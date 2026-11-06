"""
Validator Engine (Module 2).
Compares an existing, prepared Pick & Place Excel file against the current Canias ERP BOM.
Catches:
- Changed Canias material codes (Ü.A Eleman Num.)
- Changed technical specs (Value, MPN, Package, Voltage, Tolerance) with semantic differentiation
- Missing RefDes (added to BOM later but missing in prepared Excel)
- Extra/Obsolete RefDes (removed from BOM or DNP but still in prepared Excel)
- Non-SMD items accidentally placed in Excel (TH, solder paste, dummy)
- Formatting anomalies (commas in coordinates, Layer != 'T', duplicates)
- Generates auto-repaired final data keeping existing coordinates but refreshing ERP specs.
"""
import re
from decimal import Decimal
from typing import List, Dict, Tuple, Optional, Any
import openpyxl
from .models import (
    BomItem, MatchedRow, ValidationDifference, ValidationSummary
)
from .spec_parser import are_specs_semantically_equivalent, parse_specs
from .matcher_pipeline import natural_sort_key


def read_prepared_excel(file_path: str) -> List[Dict[str, Any]]:
    """
    Reads an engineer-prepared Pick & Place Excel workbook.
    Accepts sheets named 'Top', 'Bottom', or all active sheets.
    Handles headerless or headed tables.
    Returns a list of parsed component dicts.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    parsed_rows: List[Dict[str, Any]] = []

    target_sheets = []
    for sname in wb.sheetnames:
        low = sname.lower()
        if "top" in low:
            target_sheets.append((sname, "Top"))
        elif "bot" in low:
            target_sheets.append((sname, "Bottom"))

    if not target_sheets:
        target_sheets = [(wb.active.title, "Top")]

    for sname, side in target_sheets:
        sheet = wb[sname]
        start_row = 1
        
        # Check if row 1 is a header
        first_row_vals = [str(sheet.cell(1, c).value or "").strip().lower() for c in range(1, min(sheet.max_column + 1, 11))]
        if any(h in first_row_vals for h in ("designator", "refdes", "yerleşim yeri", "ref", "center-x", "katman", "layer")):
            start_row = 2

        for r in range(start_row, sheet.max_row + 1):
            refdes = str(sheet.cell(r, 1).value or "").strip()
            x = str(sheet.cell(r, 2).value or "").strip()
            y = str(sheet.cell(r, 3).value or "").strip()
            rot = str(sheet.cell(r, 4).value or "").strip()
            layer = str(sheet.cell(r, 5).value or "").strip()
            pkg = str(sheet.cell(r, 6).value or "").strip()
            spec = str(sheet.cell(r, 7).value or "").strip()
            code = str(sheet.cell(r, 8).value or "").strip()
            barcode = str(sheet.cell(r, 9).value or "").strip()
            desc = str(sheet.cell(r, 10).value or "").strip()

            if not refdes and not code and not spec:
                continue

            parsed_rows.append({
                "sheet": sname,
                "side": side,
                "row_idx": r,
                "refdes": refdes,
                "x": x,
                "y": y,
                "rotation": rot,
                "layer": layer,
                "pkg": pkg,
                "spec": spec,
                "code": code,
                "barcode": barcode,
                "desc": desc
            })

    return parsed_rows


def validate_prepared_excel_against_bom(
    prepared_excel_path: str,
    bom_items: List[BomItem],
    refdes_to_codes: Dict[str, List[str]],
    item_by_code: Dict[str, BomItem]
) -> Tuple[ValidationSummary, List[MatchedRow], List[MatchedRow]]:
    """
    Validates prepared Excel against current Canias BOM and prepares auto-repaired Top/Bottom rows.
    """
    excel_rows = read_prepared_excel(prepared_excel_path)
    summary = ValidationSummary(total_checked_refdes=len(excel_rows))

    # Precompute RefDes occurrences for O(N) duplicate detection
    ref_occurrences: Dict[str, List[Tuple[str, int]]] = {}
    for er in excel_rows:
        ru = er["refdes"].strip().upper()
        if ru:
            ref_occurrences.setdefault(ru, []).append((er["sheet"], er["row_idx"]))

    seen_excel_refs = set()
    repaired_top: List[MatchedRow] = []
    repaired_bottom: List[MatchedRow] = []

    for er in excel_rows:
        ref = er["refdes"].strip()
        ref_upper = ref.upper()
        is_duplicate_appearance = ref_upper in seen_excel_refs
        seen_excel_refs.add(ref_upper)

        sheet_name = er["sheet"]
        side = er["side"]
        row_num = er["row_idx"]

        # Check for duplicates in Excel (only flag on 2nd+ occurrence, with exact row references)
        if is_duplicate_appearance and len(ref_occurrences.get(ref_upper, [])) > 1:
            locs = ", ".join(f"{sh}:{rw}" for sh, rw in ref_occurrences[ref_upper])
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="UYARI",
                category="FORMAT",
                detail=f"Excel dosyasında mükerrer (birden fazla kez) yer alıyor ({locs}).",
                excel_value=ref,
                bom_value="Tekil olmalı",
                excel_row=row_num
            ))
            summary.warning_count += 1

        # Check format anomalies
        if "," in er["x"] or "," in er["y"]:
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="UYARI",
                category="FORMAT",
                detail="Koordinat alanında nokta '.' yerine virgül ',' kullanılmış.",
                excel_value=f"X={er['x']}, Y={er['y']}",
                bom_value="Noktalı ondalık format",
                excel_row=row_num
            ))
            summary.warning_count += 1

        if er["layer"] != "T":
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="UYARI",
                category="FORMAT",
                detail=f"Katman (Sütun E) değeri '{er['layer']}' yazılmış (Daima sabit 'T' olmalıdır).",
                excel_value=er["layer"],
                bom_value="T",
                excel_row=row_num
            ))
            summary.warning_count += 1

        # Lookup in Canias BOM
        canias_codes = refdes_to_codes.get(ref_upper)
        if not canias_codes:
            # Check without prefix
            m = re.match(r"^[A-Za-z0-9]+_(.+)$", ref_upper)
            if m:
                stripped_ref = m.group(1).upper()
                canias_codes = refdes_to_codes.get(stripped_ref)
                if canias_codes:
                    seen_excel_refs.add(stripped_ref)

        # If not found in BOM at all
        if not canias_codes:
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="FAZLA_REFDES",
                detail="Güncel Canias BOM'da bu yerleşim yeri bulunamadı (BOM'dan silinmiş veya DNP yapılmış olabilir).",
                excel_value=f"Stok: {er['code']} | Tanım: {er['spec']}",
                bom_value="BOM'da Yok",
                excel_row=row_num
            ))
            summary.error_count += 1
            summary.extra_in_excel_count += 1
            continue

        canias_code = canias_codes[0]
        bom_item = item_by_code.get(canias_code)
        if not bom_item:
            continue

        # Check if non-SMD
        if not bom_item.is_smd:
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="SMD_DISI",
                detail=f"Güncel BOM'da SMD dışı olarak işaretli ({bom_item.skip_reason}), ancak Excel'e eklenmiş!",
                excel_value=er["code"],
                bom_value=f"{bom_item.skip_reason} ({bom_item.canias_code})",
                excel_row=row_num
            ))
            summary.error_count += 1
            continue

        spec = bom_item.parsed_spec
        if not spec:
            continue

        # Check 1: Canias Stock Code change
        code_mismatch = False
        if er["code"].strip().upper() != bom_item.canias_code.strip().upper():
            code_mismatch = True
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="CANIAS_KODU",
                detail="Canias Stok Kodu (Ü.A Eleman Num.) güncel BOM ile uyuşmuyor (Malzeme kodu sonradan revize edilmiş).",
                excel_value=er["code"],
                bom_value=bom_item.canias_code,
                excel_row=row_num,
                suggested_fix={"canias_code": bom_item.canias_code}
            ))
            summary.error_count += 1

        # Check 2: Technical Spec (Column G) change
        spec_mismatch = False
        is_equiv, equiv_msg = are_specs_semantically_equivalent(er["spec"], spec.combined_spec)
        if not is_equiv:
            spec_mismatch = True
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="DEGER_PARAMETRE",
                detail=f"Birleşik Tanım (Sütun G) uyuşmuyor: {equiv_msg}",
                excel_value=er["spec"],
                bom_value=spec.combined_spec,
                excel_row=row_num,
                suggested_fix={"spec": spec.combined_spec}
            ))
            summary.error_count += 1
        elif "Semantik Denk" in equiv_msg:
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="UYARI",
                category="DEGER_PARAMETRE",
                detail=f"Yazım biçimi farkı ({equiv_msg}), elektriksel olarak denk.",
                excel_value=er["spec"],
                bom_value=spec.combined_spec,
                excel_row=row_num
            ))
            summary.warning_count += 1

        # Check 3: Package (Column F) change
        pkg_mismatch = False
        if er["pkg"].strip().upper() != spec.combined_package.strip().upper():
            pkg_mismatch = True
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="KILIF",
                detail="Birleşik Kılıf (Sütun F) güncel BOM kılıfı ile uyuşmuyor.",
                excel_value=er["pkg"],
                bom_value=spec.combined_package,
                excel_row=row_num,
                suggested_fix={"pkg": spec.combined_package}
            ))
            summary.error_count += 1

        # Check 4: Barcode (Column I) check (Canias Kodu Tekrarı)
        expected_barcode = bom_item.canias_code
        if er["barcode"].strip() and er["barcode"].strip().upper() != expected_barcode.upper():
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="UYARI",
                category="BARKOD",
                detail="Barkod (Sütun I) güncel Canias kodu tekrarı ile uyuşmuyor.",
                excel_value=er["barcode"],
                bom_value=expected_barcode,
                excel_row=row_num,
                suggested_fix={"barcode": expected_barcode}
            ))
            summary.warning_count += 1

        # Check 5: Description (Column J) check (if present in prepared Excel)
        if er["desc"].strip() and bom_item.description.strip():
            norm_excel_desc = re.sub(r"\s+", " ", er["desc"].strip().upper())
            norm_bom_desc = re.sub(r"\s+", " ", bom_item.description.strip().upper())
            if norm_excel_desc != norm_bom_desc:
                summary.differences.append(ValidationDifference(
                    refdes=ref,
                    side=side,
                    status="UYARI",
                    category="ACIKLAMA",
                    detail="BOM Açıklama metni (Sütun J) güncel Canias BOM ile farklılık gösteriyor.",
                    excel_value=er["desc"],
                    bom_value=bom_item.description,
                    excel_row=row_num,
                    suggested_fix={"description": bom_item.description}
                ))
                summary.warning_count += 1

        # If core columns (H, G, F) match
        if not code_mismatch and not spec_mismatch and not pkg_mismatch:
            summary.match_count += 1

        # Build auto-repaired row using existing coordinates
        repaired_row = MatchedRow(
            designator=ref,
            center_x=er["x"].replace(",", "."),
            center_y=er["y"].replace(",", "."),
            rotation=er["rotation"].replace(",", "."),
            layer="T",
            combined_package=spec.combined_package,
            combined_spec=spec.combined_spec,
            canias_code=bom_item.canias_code,
            barcode=expected_barcode,
            description=bom_item.description,
            side=side
        )

        if side == "Bottom":
            repaired_bottom.append(repaired_row)
        else:
            repaired_top.append(repaired_row)

    # Check for RefDes that exist in BOM but are missing in prepared Excel
    all_smd_bom_refs = {}
    for item in bom_items:
        if item.is_smd:
            for r in item.designator_list:
                all_smd_bom_refs[r.upper()] = (r, item)

    for r_upper, (orig_r, item) in all_smd_bom_refs.items():
        if r_upper not in seen_excel_refs:
            summary.differences.append(ValidationDifference(
                refdes=orig_r,
                side="Genel",
                status="HATA",
                category="EKSIK_REFDES",
                detail=f"Güncel Canias BOM'da SMD olarak mevcut, ancak hazırlanan Excel'de UNUTULMUŞ / EKSİK!",
                excel_value="Excel'de Yok",
                bom_value=f"{item.canias_code} ({item.parsed_spec.combined_spec if item.parsed_spec else ''})",
                excel_row=None
            ))
            summary.error_count += 1
            summary.missing_in_excel_count += 1

    # Calculate accuracy percentage
    total_validations = summary.total_checked_refdes + summary.missing_in_excel_count
    if total_validations > 0:
        summary.accuracy_percent = round((summary.match_count / total_validations) * 100, 2)
    else:
        summary.accuracy_percent = 100.0

    return summary, repaired_top, repaired_bottom

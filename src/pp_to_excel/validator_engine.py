"""
Validator Engine (Module 2).
Compares an existing, prepared Pick & Place Excel or .AOI file against the current Canias ERP BOM.
Focuses strictly on:
- Changed Canias material codes (Ü.A Eleman Num.)
- Missing RefDes in prepared file (SMD items present in BOM but forgotten)
- Extra RefDes in prepared file (items removed from BOM or non-SMD)
- Zero-tolerance coordinate precision (.000 3-decimal formatting and Turkish Excel thousands-separator repair)
- Generates auto-repaired final data keeping exact coordinates but refreshing ERP specs.
"""
import os
import re
import csv
from decimal import Decimal
from typing import List, Dict, Tuple, Optional, Any, Set
import openpyxl
from .models import (
    BomItem, MatchedRow, ValidationDifference, ValidationSummary
)
from .spec_parser import are_specs_semantically_equivalent, parse_specs
from .matcher_pipeline import natural_sort_key
from .pnp_parser import format_coord_3dec, repair_turkish_excel_coord, clean_angle


def read_prepared_file(file_path: str) -> List[Dict[str, Any]]:
    """
    Reads an engineer-prepared Pick & Place file (.xlsx, .AOI, .txt, .csv).
    Handles both headed and headerless formats.
    Automatically repairs Turkish Excel thousands-separator corruption (24776 -> 24.776).
    """
    parsed_rows: List[Dict[str, Any]] = []
    fp_low = file_path.lower()
    base_name = os.path.basename(file_path).lower()

    if fp_low.endswith(".xlsx"):
        wb = openpyxl.load_workbook(file_path, data_only=True)
        target_sheets = []
        for sname in wb.sheetnames:
            low = sname.lower()
            if "top" in low:
                target_sheets.append((sname, "Top"))
            elif "bot" in low:
                target_sheets.append((sname, "Bottom"))

        if not target_sheets:
            target_sheets = [(wb.active.title, "Bottom" if "bot" in base_name else "Top")]

        for sname, side in target_sheets:
            sheet = wb[sname]
            start_row = 1
            
            # Check if row 1 is a header
            first_row_vals = [str(sheet.cell(1, c).value or "").strip().lower() for c in range(1, min(sheet.max_column + 1, 11))]
            if any(h in first_row_vals for h in ("designator", "refdes", "yerleşim yeri", "ref", "center-x", "katman", "layer")):
                start_row = 2

            for r in range(start_row, sheet.max_row + 1):
                refdes = str(sheet.cell(r, 1).value or "").strip()
                raw_x = sheet.cell(r, 2).value
                raw_y = sheet.cell(r, 3).value
                raw_rot = sheet.cell(r, 4).value
                layer = str(sheet.cell(r, 5).value or "").strip()
                pkg = str(sheet.cell(r, 6).value or "").strip()
                spec = str(sheet.cell(r, 7).value or "").strip()
                code = str(sheet.cell(r, 8).value or "").strip()
                barcode = str(sheet.cell(r, 9).value or "").strip()
                desc = str(sheet.cell(r, 10).value or "").strip()

                if not refdes and not code and not spec:
                    continue

                parsed_rows.append({
                    "source_file": file_path,
                    "sheet": sname,
                    "side": side,
                    "row_idx": r,
                    "refdes": refdes,
                    "x": repair_turkish_excel_coord(raw_x),
                    "y": repair_turkish_excel_coord(raw_y),
                    "rotation": clean_angle(raw_rot),
                    "layer": layer,
                    "pkg": pkg,
                    "spec": spec,
                    "code": code,
                    "barcode": barcode,
                    "desc": desc
                })

        return parsed_rows

    # Plain text / AOI / CSV reader
    encodings = ("utf-8-sig", "utf-8", "cp1254", "iso-8859-9", "windows-1252", "latin-1")
    lines = []
    for enc in encodings:
        try:
            with open(file_path, "r", encoding=enc) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue

    if not lines:
        return []

    default_side = "Bottom" if (
        "bot" in base_name
        or "bottom" in base_name
        or re.search(r"[_\-]b(?:[_\.\-]|$)|\.b\.", base_name)
    ) else "Top"
    start_line = 0

    # Check if line 0 is a header
    first_line_low = lines[0].lower()
    if any(h in first_line_low for h in ("designator", "refdes", "yerleşim", "center-x", "layer")):
        start_line = 1

    for idx in range(start_line, len(lines)):
        line_str = lines[idx].strip()
        if not line_str or line_str.startswith("="):
            continue

        if "\t" in line_str:
            tokens = [t.strip().strip('"') for t in line_str.split("\t")]
        elif "," in line_str:
            try:
                tokens = [t.strip().strip('"') for t in next(csv.reader([line_str]))]
            except Exception:
                tokens = [t.strip().strip('"') for t in line_str.split(",")]
        else:
            tokens = [t.strip().strip('"') for t in line_str.split()]

        if not tokens:
            continue

        refdes = tokens[0] if len(tokens) > 0 else ""
        raw_x = tokens[1] if len(tokens) > 1 else "0"
        raw_y = tokens[2] if len(tokens) > 2 else "0"
        raw_rot = tokens[3] if len(tokens) > 3 else "0"
        layer = tokens[4] if len(tokens) > 4 else "T"
        pkg = tokens[5] if len(tokens) > 5 else ""
        spec = tokens[6] if len(tokens) > 6 else ""
        code = tokens[7] if len(tokens) > 7 else ""
        barcode = tokens[8] if len(tokens) > 8 else ""
        desc = tokens[9] if len(tokens) > 9 else ""

        if not refdes:
            continue

        # In AOI / SMT outputs, Column E is often forced to 'T' for both sides.
        # If the file itself is a designated Bottom file (_B_ or _BOT), preserve side as Bottom!
        if default_side == "Bottom":
            side = "Bottom"
        elif "bot" in layer.lower() or layer.upper() == "B":
            side = "Bottom"
        else:
            side = "Top"

        parsed_rows.append({
            "source_file": file_path,
            "sheet": default_side,
            "side": side,
            "row_idx": idx + 1,
            "refdes": refdes,
            "x": repair_turkish_excel_coord(raw_x),
            "y": repair_turkish_excel_coord(raw_y),
            "rotation": clean_angle(raw_rot),
            "layer": layer,
            "pkg": pkg,
            "spec": spec,
            "code": code,
            "barcode": barcode,
            "desc": desc
        })

    return parsed_rows


def read_prepared_files(file_paths: List[str]) -> List[Dict[str, Any]]:
    """Reads one or more prepared Excel / AOI files and aggregates component rows."""
    all_rows = []
    for fp in file_paths:
        rows = read_prepared_file(fp)
        all_rows.extend(rows)
    return all_rows


def validate_prepared_excel_against_bom(
    prepared_paths: List[str] | str,
    bom_items: List[BomItem],
    refdes_to_codes: Dict[str, List[str]],
    item_by_code: Dict[str, BomItem]
) -> Tuple[ValidationSummary, List[MatchedRow], List[MatchedRow]]:
    """
    Validates prepared Excel/.AOI against current Canias BOM.
    Strictly compares:
    1. RefDes existence & Canias Code (Ü.A Eleman Num.) match.
    2. Missing SMD RefDes in prepared file.
    3. Extra RefDes (or non-SMD items) in prepared file.
    No false warnings for package, spec, layer, commas, or descriptions.
    """
    if isinstance(prepared_paths, str):
        prepared_paths = [prepared_paths]

    excel_rows = read_prepared_files(prepared_paths)
    summary = ValidationSummary(total_checked_refdes=len(excel_rows))

    seen_excel_refs: Set[str] = set()
    tested_prefixes: Set[str] = set()
    repaired_top: List[MatchedRow] = []
    repaired_bottom: List[MatchedRow] = []

    for er in excel_rows:
        ref = er["refdes"].strip()
        ref_upper = ref.upper()
        sheet_name = er["sheet"]
        side = er["side"]
        row_num = er["row_idx"]

        seen_excel_refs.add(ref_upper)

        # Detect prefix of this ref
        if "_" in ref_upper:
            pfx = ref_upper.split("_", 1)[0]
            tested_prefixes.add(pfx)
        else:
            tested_prefixes.add("")

        # Lookup in Canias BOM
        canias_codes = refdes_to_codes.get(ref_upper)
        matched_key = ref_upper
        if not canias_codes:
            # Check without prefix
            m = re.match(r"^[A-Za-z0-9]+_(.+)$", ref_upper)
            if m:
                stripped_ref = m.group(1).upper()
                canias_codes = refdes_to_codes.get(stripped_ref)
                if canias_codes:
                    matched_key = stripped_ref
                    seen_excel_refs.add(stripped_ref)

        # 1. If not found in BOM at all
        if not canias_codes:
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="FAZLA_REFDES",
                detail="Güncel Canias BOM'da bu yerleşim yeri bulunamadı (BOM'dan silinmiş veya DNP yapılmış olabilir).",
                excel_value=f"Stok: {er['code']}",
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

        # 2. Check if non-SMD
        if not bom_item.is_smd:
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="SMD_DISI",
                detail=f"Güncel BOM'da SMD dışı olarak kayıtlı ({bom_item.skip_reason}), ancak dosyaya eklenmiş!",
                excel_value=er["code"],
                bom_value=f"{bom_item.skip_reason} ({bom_item.canias_code})",
                excel_row=row_num
            ))
            summary.error_count += 1
            continue

        spec = bom_item.parsed_spec
        if not spec:
            continue

        # 3. Check Canias Stock Code (Ü.A Eleman Num.)
        code_in_file = er["code"].strip().upper()
        expected_code = bom_item.canias_code.strip().upper()

        if code_in_file and code_in_file != expected_code:
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
        else:
            summary.match_count += 1

        # Build auto-repaired row using existing coordinates (guaranteed .000) and refreshed Canias specs
        repaired_row = MatchedRow(
            designator=ref,
            center_x=er["x"],
            center_y=er["y"],
            rotation=er["rotation"],
            layer="T",
            combined_package=spec.combined_package,
            combined_spec=spec.combined_spec,
            canias_code=bom_item.canias_code,
            barcode=bom_item.canias_code,
            description=bom_item.description,
            side=side
        )

        if side == "Bottom":
            repaired_bottom.append(repaired_row)
        else:
            repaired_top.append(repaired_row)

    # 4. Check for RefDes that exist in BOM but are missing in prepared file
    # Smart scope: Only check missing items for the board prefixes that were actually tested!
    all_smd_bom_refs: Dict[str, Tuple[str, BomItem]] = {}
    for item in bom_items:
        if item.is_smd:
            for r in item.designator_list:
                r_up = r.upper()
                # Check prefix
                pfx = r_up.split("_", 1)[0] if "_" in r_up else ""
                if pfx in tested_prefixes or (not pfx and "" in tested_prefixes):
                    all_smd_bom_refs[r_up] = (r, item)

    for r_upper, (orig_r, item) in all_smd_bom_refs.items():
        base_r = r_upper.split("_", 1)[1] if "_" in r_upper else r_upper
        if r_upper not in seen_excel_refs and base_r not in seen_excel_refs:
            summary.differences.append(ValidationDifference(
                refdes=orig_r,
                side="Genel",
                status="HATA",
                category="EKSIK_REFDES",
                detail="Güncel Canias BOM'da SMD olarak mevcut, ancak hazırlanan dosyada EKSİK / UNUTULMUŞ!",
                excel_value="Dosyada Yok",
                bom_value=f"{item.canias_code} ({item.parsed_spec.combined_spec if item.parsed_spec else ''})",
                excel_row=None
            ))
            summary.error_count += 1
            summary.missing_in_excel_count += 1

    # Sort repaired rows naturally
    repaired_top.sort(key=lambda r: natural_sort_key(r.designator))
    repaired_bottom.sort(key=lambda r: natural_sort_key(r.designator))

    # Calculate accuracy percentage
    total_validations = summary.total_checked_refdes + summary.missing_in_excel_count
    if total_validations > 0:
        summary.accuracy_percent = round((summary.match_count / total_validations) * 100, 2)
    else:
        summary.accuracy_percent = 100.0

    return summary, repaired_top, repaired_bottom

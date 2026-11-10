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
from .bom_parser import lookup_refdes_in_bom
from .spec_parser import are_specs_semantically_equivalent, parse_specs
from .matcher_pipeline import natural_sort_key
from .pnp_parser import (
    format_coord_3dec,
    repair_turkish_excel_coord,
    clean_angle,
    is_valid_refdes_row,
    is_valid_coordinate,
    find_column_indices,
    normalize_layer,
    _cell_str,
)


def _is_bottom_name(name: str) -> bool:
    """Checks if a sheet name or filename indicates Bottom side."""
    low = name.lower()
    return (
        "bot" in low
        or "bottom" in low
        or bool(re.search(r"(?:^|[_\-\s])b(?:[_\.\-\s]|$)|\.b\.", low))
    )


def _is_top_name(name: str) -> bool:
    """Checks if a sheet name or filename indicates Top side."""
    low = name.lower()
    return (
        "top" in low
        or bool(re.search(r"(?:^|[_\-\s])t(?:[_\.\-\s]|$)|\.t\.", low))
    )


def _unpack_prepared_columns(col_vals: List[str]) -> Tuple[str, str, str, str, str, str]:
    """
    Auto-detects 7-column format (Layer 'T' omitted: Pkg, Spec, Code, Barcode, Desc)
    vs 8-10 column format (Layer, Pkg, Spec, Code, Barcode, Desc).
    Returns (layer, pkg, spec, code, barcode, desc).
    """
    padded = (col_vals + ["", "", "", "", "", ""])[:6]
    c4_up = padded[0].strip().upper()
    is_layer_col = (
        c4_up in ("T", "B", "TOP", "BOT", "BOTTOM", "TOPLAYER", "BOTTOMLAYER")
        or (c4_up == "" and padded[3].strip() != "")
    )
    if is_layer_col:
        return (
            padded[0].strip() or "T",
            padded[1].strip(),
            padded[2].strip(),
            padded[3].strip(),
            padded[4].strip(),
            padded[5].strip(),
        )
    else:
        return (
            "T",
            padded[0].strip(),
            padded[1].strip(),
            padded[2].strip(),
            padded[3].strip(),
            padded[4].strip(),
        )


def _extract_workbook_ref_sides(wb: openpyxl.Workbook) -> Dict[str, str]:
    """
    Extracts a mapping of RefDes -> 'Top' / 'Bottom' from any raw CAD sheet
    (e.g. 'Pick&Place', 'Sheet1') or scratch sheet in the same workbook,
    so generic prepared sheets ('Sayfa1', 'Sayfa2', 'Sayfa4') with forced Layer='T'
    still get accurate Top/Bottom side classification.
    """
    ref_side_map: Dict[str, str] = {}
    for sname in wb.sheetnames:
        ws = wb[sname]
        # 1. Check headed CAD P&P sheets
        header_row = None
        col_map = {}
        for r in range(1, min(ws.max_row + 1, 25)):
            row_vals = [_cell_str(ws.cell(r, c).value) for c in range(1, min(ws.max_column + 1, 15))]
            cm, _ = find_column_indices(row_vals)
            if "designator" in cm and "layer" in cm and "x" in cm:
                header_row = r
                col_map = cm
                break
        if header_row is not None:
            d_col = col_map["designator"] + 1
            l_col = col_map["layer"] + 1
            for r in range(header_row + 1, ws.max_row + 1):
                d_val = _cell_str(ws.cell(r, d_col).value).upper()
                l_val = _cell_str(ws.cell(r, l_col).value)
                if d_val and l_val:
                    ref_side_map[d_val] = normalize_layer(l_val)
            continue

        # 2. Check headerless scratch sheets where Col 1 is RefDes and Col 2 is 'Top'/'Bottom'
        for r in range(1, ws.max_row + 1):
            c1 = _cell_str(ws.cell(r, 1).value).upper()
            c2 = _cell_str(ws.cell(r, 2).value).upper()
            if c1 and c2 in ("TOP", "BOTTOM", "TOPLAYER", "BOTTOMLAYER"):
                ref_side_map[c1] = normalize_layer(c2)
    return ref_side_map


def read_prepared_file(file_path: str) -> List[Dict[str, Any]]:
    """
    Reads an engineer-prepared Pick & Place file (.xlsx, .AOI, .txt, .csv).
    Handles both headed and headerless formats, 7-column and 8-10 column layouts,
    and multi-sheet Excel workbooks ('Sayfa1', 'Sayfa2', 'Sayfa4', 'TOP', 'BOT').
    Automatically repairs Turkish Excel thousands-separator corruption (24776 -> 24.776).
    """
    parsed_rows: List[Dict[str, Any]] = []
    fp_low = file_path.lower()
    base_name = os.path.basename(file_path).lower()

    if fp_low.endswith(".xlsx"):
        wb = openpyxl.load_workbook(file_path, data_only=True)
        workbook_ref_sides = _extract_workbook_ref_sides(wb)
        file_default_side = "Bottom" if _is_bottom_name(base_name) else "Top"

        # Scan all sheets in the workbook to find prepared production tables
        candidate_sheets: List[str] = []
        for sname in wb.sheetnames:
            sheet = wb[sname]
            start_row = 1
            first_row_vals = [_cell_str(sheet.cell(1, c).value).lower() for c in range(1, min(sheet.max_column + 1, 11))]
            if any(h in first_row_vals for h in ("designator", "refdes", "yerleşim yeri", "ref", "center-x", "katman", "layer")):
                start_row = 2

            valid_prep_rows = 0
            for r in range(start_row, min(sheet.max_row + 1, start_row + 40)):
                refdes = _cell_str(sheet.cell(r, 1).value)
                raw_x = _cell_str(sheet.cell(r, 2).value)
                raw_y = _cell_str(sheet.cell(r, 3).value)
                raw_rot = _cell_str(sheet.cell(r, 4).value)
                extra_cols = [_cell_str(sheet.cell(r, c).value) for c in range(5, 11)]
                _, pkg, spec, code, _, _ = _unpack_prepared_columns(extra_cols)
                if (
                    is_valid_refdes_row(refdes, raw_x, raw_y)
                    and (not raw_rot or is_valid_coordinate(raw_rot))
                    and (code or spec or pkg)
                ):
                    valid_prep_rows += 1
            if valid_prep_rows > 0:
                candidate_sheets.append(sname)

        if not candidate_sheets:
            for sname in wb.sheetnames:
                low = sname.lower()
                if "top" in low or "bot" in low:
                    candidate_sheets.append(sname)
        if not candidate_sheets:
            candidate_sheets = [wb.active.title]

        for sname in candidate_sheets:
            sheet = wb[sname]
            start_row = 1
            first_row_vals = [_cell_str(sheet.cell(1, c).value).lower() for c in range(1, min(sheet.max_column + 1, 11))]
            if any(h in first_row_vals for h in ("designator", "refdes", "yerleşim yeri", "ref", "center-x", "katman", "layer")):
                start_row = 2

            sheet_explicit_side = None
            if _is_bottom_name(sname):
                sheet_explicit_side = "Bottom"
            elif _is_top_name(sname):
                sheet_explicit_side = "Top"

            for r in range(start_row, sheet.max_row + 1):
                refdes = _cell_str(sheet.cell(r, 1).value)
                raw_x = sheet.cell(r, 2).value
                raw_y = sheet.cell(r, 3).value
                raw_rot = sheet.cell(r, 4).value
                extra_cols = [_cell_str(sheet.cell(r, c).value) for c in range(5, 11)]
                layer, pkg, spec, code, barcode, desc = _unpack_prepared_columns(extra_cols)

                if not refdes and not code and not spec:
                    continue
                if not is_valid_refdes_row(refdes, _cell_str(raw_x), _cell_str(raw_y)):
                    continue

                ref_u = refdes.upper()
                base_u = ref_u.split("_", 1)[1] if "_" in ref_u else ref_u

                if sheet_explicit_side:
                    side = sheet_explicit_side
                elif ref_u in workbook_ref_sides:
                    side = workbook_ref_sides[ref_u]
                elif base_u in workbook_ref_sides:
                    side = workbook_ref_sides[base_u]
                elif "bot" in layer.lower() or layer.upper() == "B":
                    side = "Bottom"
                else:
                    side = file_default_side

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

    default_side = "Bottom" if _is_bottom_name(base_name) else "Top"
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
        extra_cols = tokens[4:10]
        layer, pkg, spec, code, barcode, desc = _unpack_prepared_columns(extra_cols)

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
        side = er["side"]
        row_num = er["row_idx"]

        seen_excel_refs.add(ref_upper)

        # Detect prefix of this ref
        if "_" in ref_upper:
            pfx = ref_upper.split("_", 1)[0]
            tested_prefixes.add(pfx)
            if pfx.endswith(("L", "R")) and len(pfx) >= 3:
                tested_prefixes.add(pfx[:-1])
        else:
            tested_prefixes.add("")

        # Lookup in Canias BOM with L/R and multi-board equivalence
        canias_codes, matched_keys = lookup_refdes_in_bom(ref_upper, refdes_to_codes, item_by_code)
        seen_excel_refs.update(matched_keys)

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

        code_in_file = er["code"].strip().upper()
        matched_code = canias_codes[0]
        for c_cand in canias_codes:
            if c_cand.strip().upper() == code_in_file:
                matched_code = c_cand
                break

        bom_item = item_by_code.get(matched_code) or item_by_code.get(canias_codes[0])
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
        expected_code = bom_item.canias_code.strip().upper()

        if not code_in_file or code_in_file != expected_code:
            summary.differences.append(ValidationDifference(
                refdes=ref,
                side=side,
                status="HATA",
                category="CANIAS_KODU",
                detail="Canias Stok Kodu (Ü.A Eleman Num.) güncel BOM ile uyuşmuyor (Malzeme kodu sonradan revize edilmiş veya eksik).",
                excel_value=er["code"] or "(Boş)",
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

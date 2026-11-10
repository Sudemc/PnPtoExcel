"""
Canias ERP BOM Parser.
Handles:
- Loading .xlsx and .csv Canias ERP BOM files.
- Dynamic header detection for Kalem, Canias Code, MPN, Quantity, Unit, Description, Yerleşim Yeri.
- Multi-row continuation for same material code (e.g., Kalem 65, Kalem 66).
- Strict SMD filtering (skipping TH, Krem Lehim, Tel Lehim, PCB, Etiket, Dummy).
- Splitting, cleaning, and normalizing RefDes tokens.
- Board prefix detection (_extract_erp_prefixes) for multi-board panels.
- Parameter enrichment via spec_parser.
"""
import os
import re
import csv
from typing import List, Dict, Optional, Tuple, Set, Any
import openpyxl
from .models import BomItem, ParsedSpec
from .spec_parser import parse_specs


# RefDes token regex: R1, C5, U10, D1, TR1_C11, FM_B1, etc.
REFDES_PATTERN = re.compile(r"^[A-Za-z0-9]+(?:_[A-Za-z0-9]+)?$")
PREFIXED_REFDES_PATTERN = re.compile(r"^([A-Za-z0-9]+)_([A-Za-z]{1,4}\d+[A-Za-z0-9]*)$")


def split_references(raw_text: Optional[str]) -> List[str]:
    """
    Splits raw RefDes text into clean, normalized individual designators.
    Handles commas, semicolons, whitespace, spaces after prefix (e.g., 'TR1_ C1' -> 'TR1_C1'),
    and potential ranges (e.g. 'R1..R5' or 'R1-R5').
    """
    if not raw_text:
        return []
    
    text = str(raw_text).strip()
    if not text or text.upper() in (
        "PCB", "DUMMY", "-", "YOK", "NONE",
        "KREM LEHIM", "KREM LEHİM", "TEL LEHIM", "TEL LEHİM", "ETIKET", "ETİKET"
    ):
        return []

    # Fix space or dot after prefix: 'TR1_ C1' -> 'TR1_C1', 'TR1.C1' -> 'TR1_C1'
    text = re.sub(r"\b([A-Za-z0-9]+)[_.]\s+([A-Za-z0-9]+)\b", r"\1_\2", text)
    text = re.sub(r"\b([A-Za-z0-9]+)\.([A-Za-z0-9]+)\b", r"\1_\2", text)

    # Split by comma, semicolon, space, or newline
    raw_tokens = re.split(r"[,;\s\n\r]+", text)
    result = []

    for tok in raw_tokens:
        tok = tok.strip()
        if not tok:
            continue
        
        # Check range: R1..R5 or R1-R5 (ensure not a hyphenated part name)
        range_match = re.match(r"^([A-Za-z]{1,4})(\d+)\s*(?:\.\.|-)\s*(?:[A-Za-z]{1,4})?(\d+)$", tok)
        if range_match:
            prefix_letters, start_num, end_num = range_match.group(1), int(range_match.group(2)), int(range_match.group(3))
            if 0 < (end_num - start_num) <= 200:
                for n in range(start_num, end_num + 1):
                    result.append(f"{prefix_letters}{n}")
                continue

        # Normal token
        result.append(tok)

    return result


def is_smd_item(canias_code: str, mpn: str, description: str, unit: str, raw_refs: str) -> Tuple[bool, Optional[str]]:
    """
    Checks if a BOM item is a surface-mount device (SMD) that should be placed by SMT.
    Returns (is_smd, skip_reason).
    """
    code_up = (canias_code or "").strip().upper()
    mpn_up = (mpn or "").strip().upper()
    desc_up = (description or "").strip().upper()
    unit_up = (unit or "").strip().upper()
    refs_up = (raw_refs or "").strip().upper()

    # 1. Empty or non-component designator
    if not refs_up or refs_up == "PCB":
        return False, "Yerleşim Yeri Boş veya PCB"

    # 2. Consumable unit (GR = Krem Lehim / Tel Lehim vb.)
    if unit_up == "GR":
        return False, "Sarf Malzeme (Birim: GR - Krem/Tel Lehim)"

    # 3. Explicit Through-Hole keywords
    if desc_up.startswith("TH ") or desc_up.startswith("TH/") or desc_up.startswith("TH-") or " TH " in desc_up:
        return False, "TH (Through-Hole Malzeme)"
    if "THROUGH HOLE" in desc_up or "DIP " in desc_up:
        return False, "TH / DIP Malzeme"

    # 4. Consumables / Non-placement items in code or description
    for kw in ("KREM LEHIM", "KREM LEHİM", "TEL LEHIM", "TEL LEHİM", "LEHIM", "LEHİM", "FLUX"):
        if kw in desc_up or kw in code_up:
            return False, f"Sarf Malzeme ({kw})"

    for kw in ("ETIKET", "ETİKET", "BARKOD", "LABEL", "KOLI", "KOLİ", "KUTU", "VIDA", "VİDA", "SOMUN"):
        if kw in desc_up or kw in code_up:
            return False, f"Mekanik / Paketleme ({kw})"

    if code_up.startswith("PCB") or " PCBA " in f" {desc_up} ":
        return False, "Çıplak PCB Kartı"

    # 5. Dummy / None-PN items
    if "DUMMY" in code_up or "DUMMY" in mpn_up or "NONE/PN" in mpn_up or "NONE/PN" in desc_up:
        return False, "DUMMY / Dizgisiz Kalem"

    return True, None


def _read_bom_grid(file_path: str) -> List[List[str]]:
    """Reads an Excel (.xlsx) or CSV (.csv) BOM file into a 2D list of strings (0-indexed)."""
    if file_path.lower().endswith(".csv"):
        encodings = ("utf-8-sig", "utf-8", "cp1254", "iso-8859-9", "windows-1252", "latin-1")
        for enc in encodings:
            try:
                with open(file_path, "r", encoding=enc, newline="") as f:
                    sample = f.read(4096)
                    f.seek(0)
                    delim = ";" if sample.count(";") > sample.count(",") else ","
                    if "\t" in sample and sample.count("\t") > max(sample.count(";"), sample.count(",")):
                        delim = "\t"
                    reader = csv.reader(f, delimiter=delim)
                    return [[str(c or "").strip() for c in row] for row in reader]
            except UnicodeDecodeError:
                continue
        raise ValueError(f"CSV BOM dosyası okunamadı: {file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    grid: List[List[str]] = []
    for r in range(1, sheet.max_row + 1):
        row_vals = [str(sheet.cell(r, c).value or "").strip() for c in range(1, sheet.max_column + 1)]
        grid.append(row_vals)
    return grid


def load_canias_bom(file_path: str) -> Tuple[List[BomItem], Dict[str, List[str]], Dict[str, BomItem]]:
    """
    Loads and parses Canias ERP BOM from Excel (.xlsx) or CSV (.csv).
    Returns:
        bom_items: list of all parsed BomItem objects (including skipped ones)
        refdes_to_item: mapping of RefDes -> list of Canias stock codes (SMD prioritized, non-SMD included)
        item_by_code: mapping of Canias stock code -> BomItem
    """
    grid = _read_bom_grid(file_path)
    if not grid:
        return [], {}, {}

    max_col = max((len(r) for r in grid), default=0)

    # Detect header row and column mapping (0-based index)
    header_row_idx = None
    col_map: Dict[str, int] = {}

    for r_idx in range(min(len(grid), 25)):
        row_vals = [c.lower() for c in grid[r_idx]]
        has_code = any("eleman" in v or "malzeme" in v or "stok" in v for v in row_vals)
        has_ref = any("yerle" in v or "ref" in v for v in row_vals)
        if has_code or has_ref or any("kalem" in v for v in row_vals):
            header_row_idx = r_idx
            for c_idx, val in enumerate(row_vals):
                if not val:
                    continue
                if "kalem" in val and "tip" not in val:
                    col_map["item_no"] = c_idx
                elif "eleman" in val or "malzeme" in val or "stok" in val or "part" in val:
                    col_map["canias_code"] = c_idx
                elif "retici" in val or "mpn" in val:
                    col_map["mpn"] = c_idx
                elif "miktar" in val and "br" not in val and "birim" not in val:
                    col_map["quantity"] = c_idx
                elif "br" in val or "birim" in val or "unit" in val:
                    col_map["unit"] = c_idx
                elif "klama" in val or "desc" in val or "tan" in val:
                    col_map["description"] = c_idx
                elif "yerle" in val or "ref" in val or "designator" in val:
                    col_map["refs"] = c_idx
                elif "alternatif 1" in val or "alt 1" in val:
                    col_map["alt1"] = c_idx
            break

    if header_row_idx is None:
        header_row_idx = 0
        col_map = {
            "item_no": 0,
            "canias_code": 2,
            "mpn": 3,
            "quantity": 4,
            "unit": 5,
            "description": 6,
            "alt1": 7,
            "refs": 8 if max_col <= 10 else max(0, max_col - 1)
        }

    def _get_cell(row_list: List[str], idx: int) -> str:
        return row_list[idx].strip() if 0 <= idx < len(row_list) else ""

    raw_items: List[BomItem] = []

    for r_idx in range(header_row_idx + 1, len(grid)):
        row_list = grid[r_idx]
        item_no = _get_cell(row_list, col_map.get("item_no", 0))
        canias_code = _get_cell(row_list, col_map.get("canias_code", 2))
        mpn = _get_cell(row_list, col_map.get("mpn", 3))
        quantity = _get_cell(row_list, col_map.get("quantity", 4))
        unit = _get_cell(row_list, col_map.get("unit", 5))
        description = _get_cell(row_list, col_map.get("description", 6))
        alt1 = _get_cell(row_list, col_map.get("alt1", 7))
        raw_refs = _get_cell(row_list, col_map.get("refs", max(0, len(row_list) - 1)))

        if not canias_code and not raw_refs and not description:
            continue

        is_smd, skip_reason = is_smd_item(canias_code, mpn, description, unit, raw_refs)
        # Always split component references even for TH/Dummy items so they can be identified as non-SMD
        refs_list = split_references(raw_refs) if unit.upper() != "GR" else []

        parsed_spec = None
        if is_smd:
            parsed_spec = parse_specs(canias_code, mpn, description)

        item = BomItem(
            item_no=item_no,
            canias_code=canias_code,
            mpn=mpn,
            quantity=quantity,
            quantity_unit=unit,
            description=description,
            alt1=alt1,
            raw_designators=raw_refs,
            designator_list=refs_list,
            is_smd=is_smd,
            skip_reason=skip_reason,
            parsed_spec=parsed_spec
        )
        raw_items.append(item)

    # Combine multi-line items with the same canias_code (e.g. Kalem 65, Kalem 66)
    merged_items: List[BomItem] = []
    seen_code_map: Dict[Tuple[str, bool], BomItem] = {}

    for item in raw_items:
        key = (item.canias_code, item.is_smd)
        if item.canias_code and key in seen_code_map:
            parent = seen_code_map[key]
            for d in item.designator_list:
                if d not in parent.designator_list:
                    parent.designator_list.append(d)
            if item.raw_designators:
                parent.raw_designators = (
                    f"{parent.raw_designators}, {item.raw_designators}"
                    if parent.raw_designators else item.raw_designators
                )
        else:
            if item.canias_code:
                seen_code_map[key] = item
            merged_items.append(item)

    # Build lookup dictionaries (prioritize SMD items first, then non-SMD items so TH/Dummy RefDes are known)
    refdes_to_codes: Dict[str, List[str]] = {}
    item_by_code: Dict[str, BomItem] = {}

    # Pass 1: Non-SMD items first into item_by_code, then SMD items overwrite if same code
    for item in merged_items:
        if item.canias_code and not item.is_smd:
            item_by_code[item.canias_code] = item
    for item in merged_items:
        if item.canias_code and item.is_smd:
            item_by_code[item.canias_code] = item

    # Pass 2: Populate refdes_to_codes with SMD items first, then non-SMD items
    for item in merged_items:
        if item.is_smd and item.canias_code:
            for ref in item.designator_list:
                ref_clean = ref.upper()
                if ref_clean not in refdes_to_codes:
                    refdes_to_codes[ref_clean] = []
                if item.canias_code not in refdes_to_codes[ref_clean]:
                    refdes_to_codes[ref_clean].append(item.canias_code)

    for item in merged_items:
        if not item.is_smd and item.canias_code:
            for ref in item.designator_list:
                ref_clean = ref.upper()
                if ref_clean not in refdes_to_codes:
                    refdes_to_codes[ref_clean] = []
                if item.canias_code not in refdes_to_codes[ref_clean]:
                    refdes_to_codes[ref_clean].append(item.canias_code)

    return merged_items, refdes_to_codes, item_by_code


def extract_erp_prefixes(bom_items: List[BomItem]) -> Dict[str, Set[str]]:
    """
    Extracts all multi-board prefixes and their base designator sets from BOM.
    e.g. ['FM_C5', 'ST1_R31', 'F2_LED1', 'TR1_C11'] ->
    {'FM': {'C5'}, 'ST1': {'R31'}, 'F2': {'LED1'}, 'TR1': {'C11'}}
    """
    refs_by_prefix: Dict[str, Set[str]] = {}
    pfx_counts: Dict[str, int] = {}
    total_refs = 0

    for item in bom_items:
        for ref in item.designator_list:
            total_refs += 1
            if "_" in ref:
                parts = ref.split("_", 1)
                pfx = parts[0].strip().upper()
                base = parts[1].strip().upper()
                if 2 <= len(pfx) <= 10 and re.match(r"^[A-Za-z]+\d+", base):
                    refs_by_prefix.setdefault(pfx, set()).add(base)
                    pfx_counts[pfx] = pfx_counts.get(pfx, 0) + 1

    if not total_refs:
        return {}

    # Valid board prefixes appear in multiple components
    valid_prefixes = {
        pfx for pfx, bases in refs_by_prefix.items()
        if len(bases) >= 2 or pfx_counts.get(pfx, 0) >= 2
    }

    # If prefixed designators make up less than 10% of total designators, it's a single-board BOM
    sum_prefixed = sum(len(refs_by_prefix[p]) for p in valid_prefixes)
    if (sum_prefixed / max(1, total_refs)) < 0.10:
        return {}

    return {pfx: refs_by_prefix[pfx] for pfx in valid_prefixes}


def _clean_filename_stem(filename: str) -> str:
    """Extracts filename stem, stripping directory paths and temporary upload prefixes (e.g. cust_0_, prep_1_, 0_)."""
    base = os.path.basename(filename or "")
    base = re.sub(r"^(?:cust_\d+_|prep_\d+_|\d+_)", "", base, flags=re.IGNORECASE)
    return re.sub(r"\.[^.]+$", "", base)


def are_top_bot_pair(filenames: List[str]) -> bool:
    """Checks if a list of 2 files are simply TOP and BOT of the same board."""
    if len(filenames) != 2:
        return False
    norm_stems = []
    for fn in filenames:
        stem = _clean_filename_stem(fn).upper()
        # Remove top/bot keywords
        clean_stem = re.sub(r"\b(?:TOP|BOT|BOTTOM|T|B)\b", "", stem)
        clean_stem = re.sub(r"[_\-\s]+", "", clean_stem)
        norm_stems.append(clean_stem)
    return norm_stems[0] == norm_stems[1] and len(norm_stems[0]) >= 2


def resolve_pnp_file_prefix(
    filename: str,
    pnp_refs: Set[str],
    erp_prefixes: Set[str],
    erp_refs_by_prefix: Dict[str, Set[str]],
    all_filenames: Optional[List[str]] = None
) -> str:
    """
    Intelligently determines the board prefix for a P&P file.
    - If BOM has NO multi-board prefixes:
      - If single file or TOP/BOT pair of the same board: returns "" (NO PREFIX).
      - If multiple different board files: returns clean 2-4 letter acronym from filename.
    - If BOM HAS multi-board prefixes:
      - Uses compound word splitting (deka algorithm), L/R side detection, acronym matching, and RefDes overlap scoring.
    """
    all_fns = all_filenames or [filename]
    stem = _clean_filename_stem(filename)
    stem_upper = stem.upper()

    # 1. If ERP BOM has multi-board prefixes
    if erp_prefixes:
        tokens = [t.strip() for t in re.split(r"[_ \-]+", stem) if t.strip()]
        tokens_upper = [t.upper() for t in tokens]

        # Direct exact match of any token against erp_prefixes (e.g. 'ST1L', 'FM')
        for tok_u in tokens_upper:
            if tok_u in erp_prefixes:
                return tok_u

        # Detect Left / Right indicator in filename if prefixes have L/R variants
        has_l_in_pfx = any(p.endswith("L") for p in erp_prefixes)
        has_r_in_pfx = any(p.endswith("R") for p in erp_prefixes)
        file_side = None
        if has_l_in_pfx and any(t in ("LEFT", "LH", "SOL", "L") for t in tokens_upper):
            file_side = "L"
        elif has_r_in_pfx and any(t in ("RIGHT", "RH", "SAG", "SAĞ", "R") for t in tokens_upper):
            file_side = "R"

        # Step 1: Token & Acronym matching from filename tokens (deka algorithm)
        # e.g. StopTail1 -> ST1 (+ L -> ST1L), FogMain -> FM, TurnRev1 -> TR1
        for tok in tokens:
            tok_clean = tok.strip()
            if not tok_clean:
                continue
            spaced = re.sub(r"(?i)(FOG|STOP|TAIL|TURN|REV|MAIN|LAMP|UPPER|LOWER|REAR|FRONT|LEFT|RIGHT)", r" \1 ", tok_clean)
            sub_tokens = [t.strip().upper() for t in re.split(r"[\s_\-]+", spaced) if t.strip()]
            if sub_tokens:
                acronym = "".join(
                    t if t.isdigit() else t[0]
                    for t in sub_tokens
                    if t not in ("LEFT", "RIGHT", "LH", "RH", "LOWER", "UPPER", "FRONT", "REAR", "LAMP", "HW", "SW", "PCB", "P&P")
                ).upper()
                if not acronym:
                    continue

                # If file_side is detected and acronym + file_side exists in prefixes (e.g. ST1 + L -> ST1L)
                if file_side:
                    side_cand = acronym + file_side
                    if side_cand in erp_prefixes:
                        return side_cand

                if acronym in erp_prefixes:
                    return acronym

                # If filename is un-sided (e.g. StopTail1.xlsx) and BOM has ST1L / ST1R
                if f"{acronym}L" in erp_prefixes:
                    return f"{acronym}L"
                if f"{acronym}R" in erp_prefixes:
                    return f"{acronym}R"

                for pfx in erp_prefixes:
                    if acronym.startswith(pfx) or pfx == acronym:
                        return pfx

        # Step 2: Check longest prefix match in filename
        for pfx in sorted(erp_prefixes, key=len, reverse=True):
            if re.search(rf"(?:^|[_\-\s]){re.escape(pfx)}(?:$|[_\-\s])", stem_upper):
                return pfx

        # Step 3: RefDes set overlap fallback
        if pnp_refs:
            clean_bases = set()
            for r in pnp_refs:
                parts = r.upper().split("_", 1)
                clean_bases.add(parts[-1])

            best_pfx = None
            best_overlap = 0
            for pfx in sorted(erp_prefixes):
                bases = erp_refs_by_prefix.get(pfx, set())
                overlap = len(clean_bases & bases)
                if overlap > best_overlap or (
                    overlap == best_overlap and overlap > 0 and file_side and pfx.endswith(file_side)
                ):
                    best_overlap = overlap
                    best_pfx = pfx

            if best_pfx and best_overlap >= 2:
                return best_pfx

        return ""

    # 2. If ERP BOM has NO prefixes
    # If single file or TOP/BOT pair of same board -> NO PREFIX!
    if len(all_fns) <= 1 or are_top_bot_pair(all_fns):
        return ""

    # If multiple DIFFERENT boards are loaded for a prefixless BOM, derive short acronym to prevent RefDes collision
    cleaned = re.sub(r"(?i)\b(?:PICK|PLACE|P&P|PCB|ASSY|BOARD|SAMPLE|LOWER|UPPER|HW\d+|SW\d+|V\d+)\b", "", stem).strip()
    cleaned = re.sub(r"\b\d{4}[-_]\d{2}[-_]\d{2}\b", "", cleaned).strip()
    spaced = re.sub(r"(?i)(FOG|STOP|TAIL|TURN|REV|MAIN|LAMP|UPPER|LOWER|REAR|FRONT|LEFT|RIGHT)", r" \1 ", cleaned)
    tokens = [t.strip().upper() for t in re.split(r"[\s_\-]+", spaced) if t.strip()]
    if tokens:
        acro = "".join(t if t.isdigit() else t[0] for t in tokens if t not in ("LEFT", "RIGHT", "LH", "RH"))
        if 2 <= len(acro) <= 8:
            return acro
    return ""


def lookup_refdes_in_bom(
    ref_upper: str,
    refdes_to_codes: Dict[str, List[str]],
    item_by_code: Optional[Dict[str, BomItem]] = None
) -> Tuple[Optional[List[str]], Set[str]]:
    """
    Looks up a component designator in Canias BOM with full Left/Right (L/R) and multi-board equivalence:
    1. Exact match (e.g. 'ST1L_C1' -> 'ST1L_C1')
    2. Prefix-stripped base match for single-board BOMs ('ST1_C1' -> 'C1')
    3. L/R-stripped prefix match ('ST1L_C1' or 'ST1R_C1' -> 'ST1_C1' when BOM has un-sided 'ST1_C1')
    4. L/R-expanded prefix match ('ST1_C1' -> 'ST1L_C1' & 'ST1R_C1' when BOM has L/R prefixes)
    5. Symmetrical L/R twin fallback ('T2R_LED1' <-> 'T2L_LED1' when BOM has L/R typo like 'T2L_LED1, T2R_LED2')
    Returns (canias_codes, matched_bom_keys_set).
    """
    ref_u = ref_upper.strip().upper()
    if not ref_u:
        return None, set()

    # 1. Exact match
    codes = refdes_to_codes.get(ref_u)
    if codes:
        matched_keys = {ref_u}
        # If this is <PFX>L_<BASE> or <PFX>R_<BASE>, check if un-sided <PFX>_<BASE> also exists
        if "_" in ref_u:
            pfx, base = ref_u.split("_", 1)
            if pfx.endswith(("L", "R")) and len(pfx) >= 3:
                matched_keys.add(f"{pfx[:-1]}_{base}")
        return codes, matched_keys

    if "_" in ref_u:
        pfx, base = ref_u.split("_", 1)

        # 2. Prefix-stripped base match (for single-board BOMs without prefixes)
        codes = refdes_to_codes.get(base)
        if codes:
            return codes, {ref_u, base}

        # 3. L/R-stripped prefix match: <PFX>L_<BASE> or <PFX>R_<BASE> -> <PFX>_<BASE>
        if pfx.endswith(("L", "R")) and len(pfx) >= 3:
            base_pfx = pfx[:-1].rstrip("_")
            cand_unsided = f"{base_pfx}_{base}"
            codes = refdes_to_codes.get(cand_unsided)
            if codes:
                return codes, {ref_u, cand_unsided}

        # 4. L/R-expanded prefix match: <PFX>_<BASE> -> <PFX>L_<BASE> & <PFX>R_<BASE>
        cand_l = f"{pfx}L_{base}"
        cand_r = f"{pfx}R_{base}"
        codes_l = refdes_to_codes.get(cand_l)
        codes_r = refdes_to_codes.get(cand_r)
        if codes_l or codes_r:
            return (codes_l or codes_r), {ref_u, cand_l, cand_r}

        # 5. Symmetrical L/R twin fallback: <PFX>R_<BASE> <-> <PFX>L_<BASE>
        # Handles cases where the same part is on Left & Right boards (e.g. T2L_LED1 & T2R_LED1)
        # and ERP BOM wrote T2L_LED1, T2R_LED2 (typo on one side's index) or only listed one side.
        if pfx.endswith(("L", "R")) and len(pfx) >= 3:
            twin_side = "L" if pfx.endswith("R") else "R"
            twin_pfx = pfx[:-1] + twin_side
            twin_ref = f"{twin_pfx}_{base}"
            twin_codes = refdes_to_codes.get(twin_ref)
            if twin_codes:
                matched_keys = {ref_u, twin_ref}
                if item_by_code and twin_codes[0] in item_by_code:
                    b_item = item_by_code[twin_codes[0]]
                    m_cls = re.match(r"^([A-Z]+)\d+$", base)
                    if m_cls:
                        comp_cls = m_cls.group(1)
                        for d in b_item.designator_list:
                            d_up = d.upper()
                            if d_up.startswith(f"{pfx}_{comp_cls}"):
                                matched_keys.add(d_up)
                return twin_codes, matched_keys
    else:
        # Unprefixed ref_u (e.g. 'B1'), check if BOM only has single L/R pair like '<PFX>L_B1', '<PFX>R_B1'
        matching_lr = [k for k in refdes_to_codes if k.endswith(f"_{ref_u}")]
        if len(matching_lr) == 2:
            p0 = matching_lr[0].split("_", 1)[0]
            p1 = matching_lr[1].split("_", 1)[0]
            if len(p0) == len(p1) and p0[:-1] == p1[:-1] and {p0[-1], p1[-1]} == {"L", "R"}:
                return refdes_to_codes[matching_lr[0]], {ref_u, matching_lr[0], matching_lr[1]}

    return None, set()


def suggest_prefix_for_filename(filename: str, detected_prefixes: Optional[Dict[str, Any]] = None, all_filenames: Optional[List[str]] = None) -> str:
    """Compatibility wrapper for suggest_prefix_for_filename."""
    known_pfxs = set(detected_prefixes.keys()) if isinstance(detected_prefixes, dict) else set()
    refs_by_pfx = {k: set(v) for k, v in detected_prefixes.items()} if isinstance(detected_prefixes, dict) else {}
    return resolve_pnp_file_prefix(filename, set(), known_pfxs, refs_by_pfx, all_filenames=all_filenames)

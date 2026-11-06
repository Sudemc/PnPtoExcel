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
import re
import csv
from typing import List, Dict, Optional, Tuple, Set
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


def extract_erp_prefixes(bom_items: List[BomItem]) -> Dict[str, List[str]]:
    """
    Extracts all multi-board prefixes from BOM designators.
    e.g. ['FM_C5', 'ST1_R31', 'F2_LED1', 'TR1_C11'] -> {'FM': ['FM_C5'], 'ST1': ['ST1_R31'], 'F2': ['F2_LED1'], 'TR1': ['TR1_C11']}
    Ignores isolated underscored component names (like 'R_TH1' on a single-board BOM).
    """
    all_counts: Dict[str, List[str]] = {}
    total_refs = 0
    prefixed_refs = 0

    for item in bom_items:
        for ref in item.designator_list:
            total_refs += 1
            m = PREFIXED_REFDES_PATTERN.match(ref)
            if m:
                prefixed_refs += 1
                pfx = m.group(1).upper()
                all_counts.setdefault(pfx, []).append(ref)

    # A BOM is multi-board if prefixed designators make up at least 15% of total designators,
    # and each valid board prefix has at least 3 components.
    if total_refs == 0 or (prefixed_refs / total_refs) < 0.15:
        return {}

    return {
        pfx: refs[:5]
        for pfx, refs in all_counts.items()
        if len(refs) >= 3
    }


def suggest_prefix_for_filename(filename: str, detected_prefixes: Optional[Dict[str, List[str]]] = None) -> str:
    """
    Suggests a board prefix (e.g. 'ST1', 'FM', 'TR1', 'F2') from a Pick & Place filename.
    Uses CamelCase acronym extraction (e.g., 'StopTail1' -> 'ST1', 'FogMain' -> 'FM', 'TurnRev2' -> 'TR2')
    and matches against detected BOM prefixes if available.
    """
    base = re.sub(r"\.[^.]+$", "", filename)
    tokens = [t for t in re.split(r"[\s_\-]+", base) if t]
    known_pfxs = {k.upper() for k in (detected_prefixes or {}).keys()}

    # 1. Generate CamelCase / uppercase+digit acronyms for each token in filename
    # e.g. 'StopTail1' -> 'ST1', 'FogMain' -> 'FM', 'Fog2' -> 'F2', 'TurnRev1' -> 'TR1'
    candidate_acronyms: List[str] = []
    for tok in tokens:
        if any(c.islower() for c in tok) and any(c.isupper() for c in tok):
            acr = "".join(re.findall(r"[A-Z]|\d+", tok))
            if len(acr) >= 2:
                candidate_acronyms.append(acr.upper())
        tok_up = tok.upper()
        if known_pfxs and tok_up in known_pfxs:
            return tok_up

    if known_pfxs:
        for acr in candidate_acronyms:
            if acr in known_pfxs:
                return acr
        # Fallback: word-boundary / multi-char prefix match (longest prefix first)
        fn_up = base.upper()
        for pfx in sorted(known_pfxs, key=len, reverse=True):
            if len(pfx) >= 2 and re.search(rf"(?:^|[_\-\s]){re.escape(pfx)}(?:$|[_\-\s])", fn_up):
                return pfx
        return ""

    # Fallback when BOM is not yet loaded: return CamelCase acronym excluding common words
    ignore_acronyms = {"HW01", "HW02", "HW03", "SW01", "V1", "V2", "V8"}
    for acr in candidate_acronyms:
        if acr not in ignore_acronyms and not acr.startswith("HW"):
            return acr

    return ""

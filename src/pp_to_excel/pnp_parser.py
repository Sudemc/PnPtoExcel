"""
Pick & Place (Centroid / XYR) Parser.
Supports:
- Altium Designer (.txt, .csv)
- Cadence Allegro (.xlsx)
- Feka Multi-Board P&P (.xlsx)
- Generic CSV / Tab / Fixed-width CAD files
- Robust encoding fallback (utf-8-sig, utf-8, cp1254, iso-8859-9, latin-1)
- Dynamic header detection & heuristic sniffing
- Zero rounding mil -> mm Decimal conversion (all fractional digits preserved)
- Allegro rotation scale correction (270000 -> 270)
- Board prefix injection (<PREFIX>_<DESIGNATOR>)
"""
import re
import csv
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import List, Dict, Optional, Tuple, Any
import openpyxl
from .models import PnPRow


# Regex to match valid RefDes (including standard R1, C5, U10, ST1_R31 and non-numbered JDBG, JSPI, RSH, R_TH, IGND)
VALID_REFDES_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9_\-]{0,19}$")
PREFIXED_REFDES_PATTERN = re.compile(r"^[A-Za-z0-9]+_[A-Za-z][A-Za-z0-9_\-]*$")

RESERVED_REFDES_WORDS = {
    "DESIGNATOR", "REFDES", "REF", "PART", "COMPONENT", "PROJECT", "REVISION",
    "CREATION", "DATE", "TIME", "VARIANT", "UNITS", "NO", "LAYER", "FOOTPRINT",
    "COMMENT", "DESCRIPTION", "TOTAL", "COUNT", "PAGE", "TITLE", "COMPANY",
    "AUTHOR", "FILE", "SIDE", "KATMAN", "YERLESIM"
}

MIL_TO_MM_RATIO = Decimal("0.0254")
TOKEN_REGEX = re.compile(r'"[^"]*"|\S+')


def format_coord_3dec(val_str: Any) -> str:
    """
    Formats a coordinate string to strictly 3 decimal places with dot separator (.000).
    - 4+ decimals -> quantize to 3 decimals (e.g. '20.5792' -> '20.579', '64.1722618' -> '64.172')
    - 1-2 decimals -> pad with trailing zeroes (e.g. '29.95' -> '29.950', '6.1' -> '6.100')
    - integers -> pad with .000 (e.g. '5' -> '5.000', '0' -> '0.000')
    - handles comma or dot, strips mm/mil units.
    """
    if val_str is None:
        return "0.000"
    s = str(val_str).strip().replace(",", ".")
    s = re.sub(r"(?:mm|mil|thou)$", "", s, flags=re.IGNORECASE).strip()
    s = re.sub(r"[^\d.\-+]", "", s)
    if not s or s in ("-", "+", "."):
        return "0.000"
    try:
        dec = Decimal(s)
        dec_3 = dec.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
        return f"{dec_3:f}"
    except (InvalidOperation, ValueError):
        return s


def repair_turkish_excel_coord(raw_val: Any) -> str:
    """
    Repairs coordinates corrupted by Turkish Excel thousands-separator interpretation.
    In tr-TR Excel, '24.776' is converted to integer 24776.
    If value is an integer or string with no dot and abs(val) >= 1000, divides by 1000
    and formats to 3 decimal places ('24.776').
    """
    if raw_val is None:
        return "0.000"
    s = str(raw_val).strip()
    if "." in s or "," in s:
        return format_coord_3dec(s)
    try:
        val_int = int(s)
        if abs(val_int) >= 1000:
            dec = Decimal(val_int) / Decimal(1000)
            return format_coord_3dec(str(dec))
        return format_coord_3dec(s)
    except (ValueError, InvalidOperation):
        return format_coord_3dec(s)


def is_valid_coordinate(val_str: str) -> bool:
    """Checks if a string represents a valid numeric coordinate (with optional mm/mil suffix)."""
    if not val_str:
        return False
    s = str(val_str).strip().lower()
    s = re.sub(r"(?:mm|mil|thou)$", "", s).strip()
    s = s.replace(",", ".")
    return bool(re.match(r"^[+\-]?(?:\d+\.?\d*|\.\d+)$", s))


def is_valid_refdes_row(raw_designator: str, raw_x: str, raw_y: str) -> bool:
    """
    Validates that a row has a genuine component designator AND valid numeric X, Y coordinates.
    Accepts both numbered designators (R1, C10, ST1_R31) and symbolic ones (JDBG, JSPI, RSH, R_TH, IGND).
    """
    if not raw_designator:
        return False
    d_clean = raw_designator.strip()
    if d_clean.upper() in RESERVED_REFDES_WORDS:
        return False
    if not VALID_REFDES_PATTERN.match(d_clean) and not PREFIXED_REFDES_PATTERN.match(d_clean):
        return False
    return is_valid_coordinate(raw_x) and is_valid_coordinate(raw_y)


def _infer_unit_from_coordinates(coord_pairs: List[Tuple[str, str]]) -> str:
    """
    Heuristic unit detection (Adım 2.1): if coordinates frequently exceed 800,
    they are in mil (thou) rather than mm.
    """
    if not coord_pairs:
        return "mm"
    vals = []
    for rx, ry in coord_pairs:
        for raw in (rx, ry):
            s = re.sub(r"[^\d.\-+]", "", str(raw).strip().replace(",", "."))
            try:
                vals.append(abs(float(s)))
            except ValueError:
                pass
    if not vals:
        return "mm"
    high_count = sum(1 for v in vals if v > 800.0)
    if high_count >= max(1, int(len(vals) * 0.25)):
        return "mil"
    return "mm"


def normalize_decimal_str(val_str: str) -> str:
    """Formats coordinate string to strictly 3 decimal places (.000)."""
    return format_coord_3dec(val_str)


def convert_mil_to_mm(mil_val_str: str) -> str:
    """
    Converts mil (thou) to mm and formats to strictly 3 decimal places (.000).
    e.g. 2526.467 mil -> 64.172 mm
    """
    s = mil_val_str.strip().replace(",", ".")
    s = re.sub(r"[^\d.\-+]", "", s)
    try:
        dec_mil = Decimal(s)
        dec_mm = dec_mil * MIL_TO_MM_RATIO
        return format_coord_3dec(str(dec_mm))
    except InvalidOperation:
        return format_coord_3dec(mil_val_str)


def clean_angle(angle_val_str: Any) -> str:
    """
    Cleans rotation angle.
    Handles Cadence Allegro 1000x scale factor (e.g. 270000 -> 270, 180000 -> 180).
    Exact integers remain clean integer ('0', '90', '180', '270').
    Fractional angles are formatted to 3 decimal places (.000, e.g. 111.34 -> 111.340).
    """
    if angle_val_str is None:
        return "0"
    s = str(angle_val_str).strip().replace(",", ".").replace("°", "")
    s = re.sub(r"[^\d.\-+]", "", s)
    if not s or s in ("-", "+", "."):
        return "0"
    try:
        val = float(s)
        # Cadence Allegro 1000x angle fix
        if abs(val) > 360 and val % 1000 == 0:
            val = val / 1000.0
        
        # Check integer angle
        if val.is_integer():
            norm_deg = int(val) % 360
            return str(norm_deg)

        # Fractional angle: normalize to 0..360 and format to 3 decimals
        deg = val % 360.0
        return format_coord_3dec(str(deg))
    except (ValueError, InvalidOperation):
        return s


def normalize_layer(layer_raw: Any) -> str:
    """
    Normalizes layer string into 'Top' or 'Bottom'.
    Supports Altium ('TopLayer', 'BottomLayer'), Cadence SYM_MIRROR ('NO' -> Top, 'YES' -> Bottom),
    and standard 'T' / 'B'.
    """
    s = str(layer_raw or "").strip().upper()
    if s in ("TOPLAYER", "TOP", "T", "SURFACE", "NO"):
        return "Top"
    if s in ("BOTTOMLAYER", "BOTTOM", "BOT", "B", "YES"):
        return "Bottom"
    if "TOP" in s:
        return "Top"
    if "BOT" in s:
        return "Bottom"
    return "Top"


def find_column_indices(header_tokens: List[str]) -> Tuple[Dict[str, int], Optional[str]]:
    """
    Detects column indices for RefDes, X, Y, Rotation, Layer from header tokens.
    Also detects unit ('mil' or 'mm') if present in X/Y headers.
    """
    col_map = {}
    detected_unit = None

    for idx, tok in enumerate(header_tokens):
        t = tok.strip().lower()
        t_clean = re.sub(r"[\s\-_()]+", "", t)

        # RefDes / Designator
        if any(k in t_clean for k in ("designator", "refdes", "ref", "component", "part")) and "footprint" not in t_clean and "val" not in t_clean:
            if "designator" not in col_map:
                col_map["designator"] = idx

        # Center-X
        elif any(k in t_clean for k in ("centerx", "posx", "midx", "refx", "symcenterx")) or t_clean == "x":
            if "x" not in col_map:
                col_map["x"] = idx
                if "mil" in t:
                    detected_unit = "mil"
                elif "mm" in t:
                    detected_unit = "mm"

        # Center-Y
        elif any(k in t_clean for k in ("centery", "posy", "midy", "refy", "symcentery")) or t_clean == "y":
            if "y" not in col_map:
                col_map["y"] = idx
                if "mil" in t and not detected_unit:
                    detected_unit = "mil"
                elif "mm" in t and not detected_unit:
                    detected_unit = "mm"

        # Rotation (Avoid matching 'ac' inside 'package'!)
        elif any(k in t_clean for k in ("rotation", "rotate", "angle", "rot", "symrotate", "orientation", "aci", "açı")):
            if "rotation" not in col_map:
                col_map["rotation"] = idx

        # Layer / Side / SYM_MIRROR
        elif any(k in t_clean for k in ("layer", "side", "katman", "surface", "symmirror")):
            if "layer" not in col_map:
                col_map["layer"] = idx

    return col_map, detected_unit


def parse_pnp_xlsx(file_path: str, prefix: str = "", unit_override: Optional[str] = None) -> List[PnPRow]:
    """
    Parses an Excel (.xlsx) Pick & Place file.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    
    target_sheet = None
    preferred_names = ["pick&place", "pick place", "centroid", "placement", "pnp", "p&p"]
    for sname in wb.sheetnames:
        if any(p in sname.lower() for p in preferred_names):
            target_sheet = wb[sname]
            break

    if target_sheet is None:
        for sname in wb.sheetnames:
            s = wb[sname]
            for r in range(1, min(s.max_row + 1, 20)):
                row_vals = [str(s.cell(r, c).value or "").lower() for c in range(1, s.max_column + 1)]
                if any("designator" in v or "refdes" in v for v in row_vals):
                    target_sheet = s
                    break
            if target_sheet:
                break

    if target_sheet is None:
        target_sheet = wb.active

    header_row_idx = None
    col_map = {}
    detected_unit = None

    for r in range(1, min(target_sheet.max_row + 1, 30)):
        row_vals = [str(target_sheet.cell(r, c).value or "").strip() for c in range(1, target_sheet.max_column + 1)]
        cm, u = find_column_indices(row_vals)
        if "designator" in cm and "x" in cm and "y" in cm:
            header_row_idx = r
            col_map = cm
            detected_unit = u
            break

    if not header_row_idx or "designator" not in col_map or "x" not in col_map or "y" not in col_map:
        raise ValueError(f"Pick & Place dosyasında gerekli sütunlar (Designator, X, Y) tespit edilemedi: {file_path}")

    raw_candidates = []
    for r in range(header_row_idx + 1, target_sheet.max_row + 1):
        raw_designator = str(target_sheet.cell(r, col_map["designator"] + 1).value or "").strip()
        raw_x = str(target_sheet.cell(r, col_map["x"] + 1).value or "").strip()
        raw_y = str(target_sheet.cell(r, col_map["y"] + 1).value or "").strip()

        rot_idx = col_map.get("rotation")
        raw_rot = str(target_sheet.cell(r, rot_idx + 1).value or "").strip() if rot_idx is not None else "0"

        layer_idx = col_map.get("layer")
        raw_layer = str(target_sheet.cell(r, layer_idx + 1).value or "").strip() if layer_idx is not None else "Top"

        if not is_valid_refdes_row(raw_designator, raw_x, raw_y):
            continue

        raw_candidates.append((r, raw_designator, raw_x, raw_y, raw_rot, raw_layer))

    effective_unit = unit_override or detected_unit
    if not effective_unit:
        effective_unit = _infer_unit_from_coordinates([(rx, ry) for _, _, rx, ry, _, _ in raw_candidates])

    rows: List[PnPRow] = []
    for r, raw_designator, raw_x, raw_y, raw_rot, raw_layer in raw_candidates:
        if effective_unit == "mil" and "mm" not in raw_x.lower():
            cx = convert_mil_to_mm(raw_x)
            cy = convert_mil_to_mm(raw_y)
        else:
            cx = normalize_decimal_str(raw_x)
            cy = normalize_decimal_str(raw_y)

        clean_rot = clean_angle(raw_rot)
        norm_layer = normalize_layer(raw_layer)

        final_designator = raw_designator
        if prefix:
            clean_pfx = prefix.strip().rstrip("_")
            if not raw_designator.startswith(f"{clean_pfx}_"):
                final_designator = f"{clean_pfx}_{raw_designator}"

        pnp_row = PnPRow(
            designator=final_designator,
            raw_designator=raw_designator,
            prefix=prefix,
            center_x=cx,
            center_y=cy,
            rotation=clean_rot,
            layer=norm_layer,
            source_file=file_path,
            original_line=r
        )
        rows.append(pnp_row)

    return rows


def parse_pnp_text(file_path: str, prefix: str = "", unit_override: Optional[str] = None) -> List[PnPRow]:
    """
    Parses a plain text, TSV, or CSV (.txt, .csv) Pick & Place file.
    Uses encoding fallback and handles multiline quoted strings.
    """
    encodings = ("utf-8-sig", "utf-8", "cp1254", "iso-8859-9", "windows-1252", "latin-1")
    lines: List[str] = []

    for enc in encodings:
        try:
            with open(file_path, "r", encoding=enc) as f:
                lines = f.readlines()
            break
        except UnicodeDecodeError:
            continue

    if not lines:
        raise ValueError(f"Dosya okunamadı veya desteklenmeyen karakter kodlaması: {file_path}")

    # Check top metadata for units
    detected_unit = None
    for line in lines[:30]:
        line_low = line.lower()
        if "units used: mil" in line_low or "units: mil" in line_low:
            detected_unit = "mil"
            break
        elif "units used: mm" in line_low or "units: mm" in line_low:
            detected_unit = "mm"
            break

    # Reconstruct lines balancing quotes
    merged_lines = []
    buffer = ""
    for line in lines:
        stripped = line.rstrip("\r\n")
        if not stripped.strip():
            continue
        if not buffer:
            buffer = stripped
        else:
            buffer += " " + stripped.strip()

        if buffer.count('"') % 2 == 0:
            merged_lines.append(buffer)
            buffer = ""
    if buffer:
        merged_lines.append(buffer)

    header_idx = None
    col_map = {}
    header_unit = None
    split_mode = "whitespace"  # 'csv', 'tab', or 'whitespace'

    for idx, line in enumerate(merged_lines[:40]):
        # 1. Check tab-separated format
        if "\t" in line and line.count("\t") >= 3:
            tokens = [t.strip().strip('"') for t in line.split("\t")]
            cm, u = find_column_indices(tokens)
            if "designator" in cm and "x" in cm and "y" in cm:
                header_idx = idx
                col_map = cm
                header_unit = u
                split_mode = "tab"
                break

        # 2. Check comma-separated format
        if "," in line and line.count(",") >= 3:
            tokens = [t.strip().strip('"') for t in next(csv.reader([line]))]
            cm, u = find_column_indices(tokens)
            if "designator" in cm and "x" in cm and "y" in cm:
                header_idx = idx
                col_map = cm
                header_unit = u
                split_mode = "csv"
                break

        # 3. Try whitespace/quoted tokenization
        tokens = [t.strip().strip('"') for t in TOKEN_REGEX.findall(line)]
        cm, u = find_column_indices(tokens)
        if "designator" in cm and "x" in cm and "y" in cm:
            header_idx = idx
            col_map = cm
            header_unit = u
            split_mode = "whitespace"
            break

    if header_idx is None or "designator" not in col_map or "x" not in col_map or "y" not in col_map:
        # Fallback: Headerless .AOI or CAD file where data begins directly at line 0
        for sample_line in merged_lines[:5]:
            s_line = sample_line.strip()
            if not s_line or s_line.startswith("="):
                continue
            
            # Check tab-delimited headerless .AOI (RefDes \t X \t Y \t Rot \t Layer ...)
            if "\t" in s_line:
                sample_toks = [t.strip().strip('"') for t in s_line.split("\t")]
                if len(sample_toks) >= 3 and is_valid_refdes_row(sample_toks[0], sample_toks[1], sample_toks[2]):
                    header_idx = -1
                    col_map = {"designator": 0, "x": 1, "y": 2}
                    if len(sample_toks) > 3 and is_valid_coordinate(sample_toks[3]):
                        col_map["rotation"] = 3
                    if len(sample_toks) > 4:
                        col_map["layer"] = 4
                    split_mode = "tab"
                    break

            # Check CSV headerless
            if "," in s_line:
                try:
                    sample_toks = [t.strip().strip('"') for t in next(csv.reader([s_line]))]
                    if len(sample_toks) >= 3 and is_valid_refdes_row(sample_toks[0], sample_toks[1], sample_toks[2]):
                        header_idx = -1
                        col_map = {"designator": 0, "x": 1, "y": 2}
                        if len(sample_toks) > 3 and is_valid_coordinate(sample_toks[3]):
                            col_map["rotation"] = 3
                        if len(sample_toks) > 4:
                            col_map["layer"] = 4
                        split_mode = "csv"
                        break
                except Exception:
                    pass

            # Check whitespace headerless
            sample_toks = [t.strip().strip('"') for t in TOKEN_REGEX.findall(s_line)]
            if len(sample_toks) >= 3 and is_valid_refdes_row(sample_toks[0], sample_toks[1], sample_toks[2]):
                header_idx = -1
                col_map = {"designator": 0, "x": 1, "y": 2}
                if len(sample_toks) > 3 and is_valid_coordinate(sample_toks[3]):
                    col_map["rotation"] = 3
                if len(sample_toks) > 4:
                    col_map["layer"] = 4
                split_mode = "whitespace"
                break

    if header_idx is None or "designator" not in col_map or "x" not in col_map or "y" not in col_map:
        raise ValueError(f"Pick & Place dosyasında gerekli sütunlar (Designator, X, Y) bulunamadı: {file_path}")

    raw_candidates = []
    for r_idx in range(header_idx + 1, len(merged_lines)):
        line = merged_lines[r_idx].strip()
        if not line or line.startswith("="):
            continue

        if split_mode == "tab":
            tokens = [t.strip().strip('"') for t in merged_lines[r_idx].split("\t")]
        elif split_mode == "csv":
            try:
                tokens = [t.strip().strip('"') for t in next(csv.reader([line]))]
            except Exception:
                tokens = [t.strip().strip('"') for t in TOKEN_REGEX.findall(line)]
        else:
            tokens = [t.strip().strip('"') for t in TOKEN_REGEX.findall(line)]

        if len(tokens) <= max(col_map["designator"], col_map["x"], col_map["y"]):
            continue

        raw_designator = tokens[col_map["designator"]].strip()
        raw_x = tokens[col_map["x"]].strip()
        raw_y = tokens[col_map["y"]].strip()

        rot_idx = col_map.get("rotation")
        raw_rot = tokens[rot_idx].strip() if rot_idx is not None and rot_idx < len(tokens) else "0"

        layer_idx = col_map.get("layer")
        raw_layer = tokens[layer_idx].strip() if layer_idx is not None and layer_idx < len(tokens) else "Top"

        if not is_valid_refdes_row(raw_designator, raw_x, raw_y):
            continue

        raw_candidates.append((r_idx + 1, raw_designator, raw_x, raw_y, raw_rot, raw_layer))

    effective_unit = unit_override or detected_unit or header_unit
    if not effective_unit:
        effective_unit = _infer_unit_from_coordinates([(rx, ry) for _, _, rx, ry, _, _ in raw_candidates])

    rows: List[PnPRow] = []
    for line_no, raw_designator, raw_x, raw_y, raw_rot, raw_layer in raw_candidates:
        if effective_unit == "mil" and "mm" not in raw_x.lower():
            cx = convert_mil_to_mm(raw_x)
            cy = convert_mil_to_mm(raw_y)
        else:
            cx = normalize_decimal_str(raw_x)
            cy = normalize_decimal_str(raw_y)

        clean_rot = clean_angle(raw_rot)
        norm_layer = normalize_layer(raw_layer)

        final_designator = raw_designator
        if prefix:
            clean_pfx = prefix.strip().rstrip("_")
            if not raw_designator.startswith(f"{clean_pfx}_"):
                final_designator = f"{clean_pfx}_{raw_designator}"

        pnp_row = PnPRow(
            designator=final_designator,
            raw_designator=raw_designator,
            prefix=prefix,
            center_x=cx,
            center_y=cy,
            rotation=clean_rot,
            layer=norm_layer,
            source_file=file_path,
            original_line=line_no
        )
        rows.append(pnp_row)

    return rows


def load_pnp_file(file_path: str, prefix: str = "", unit_override: Optional[str] = None) -> List[PnPRow]:
    """
    Entry point to load any Pick & Place file (.xlsx, .txt, .csv).
    """
    fp_low = file_path.lower()
    if fp_low.endswith(".xlsx"):
        return parse_pnp_xlsx(file_path, prefix=prefix, unit_override=unit_override)
    else:
        return parse_pnp_text(file_path, prefix=prefix, unit_override=unit_override)

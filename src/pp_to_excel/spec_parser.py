"""
Technical parameter parser for Canias ERP BOM descriptions and manufacturer part numbers.
Strictly implements:
- Part Name (MPN) extraction for active components (FET, TRA, DIO, IC, MCU, etc.)
- Value extraction for passive components (C, R, L, Ferrite Bead, NTC, Crystal, etc.)
- Voltage rule: Only capacitors get voltage, and ONLY when voltage != '50V'
- Tolerance rule: Passives/actives get tolerance ONLY when tolerance != '%5'
- Semantic electrical equivalence (100nF == 0.1uF, 4k7 == 4.7k, 49R9 == 49.9R)
"""
import re
from decimal import Decimal, InvalidOperation
from typing import Optional, Tuple
from .models import ParsedSpec


# Standard imperial chip sizes
CHIP_SIZES = {"0201", "0402", "0603", "0805", "1206", "1210", "1812", "2010", "2220", "2512"}

# Common IC/Semiconductor/SMD packages (ordered from more specific to general)
SEMICON_PACKAGES = [
    r"16-TSSOP-EP", r"\d+-TSSOP(?:-EP)?", r"TSSOP-\d+(?:-EP)?",
    r"\d+-SSOP", r"SSOP-\d+",
    r"\d+-MSOP", r"MSOP-\d+",
    r"6-TSOP", r"\d+-TSOP", r"TSOP-\d+", r"TSOP",
    r"TSOT-?\d+",
    r"8-PowerSOIC", r"\d+-PowerSOIC",
    r"SOIC-\d+", r"SO-\d+", r"SOP-\d+",
    r"SOT-?23(?:-\d+[A-Z]*)?", r"SOT-?223(?:-\d+)?", r"SOT-?323", r"SOT-?363",
    r"SOT-?89", r"SOT-?563", r"SOT-?583",
    r"SOD-?123(?:FL)?", r"SOD-?323(?:FL)?", r"SOD-?523", r"SOD-?923", r"SOD-?80",
    r"PowerDI-\d+",
    r"PPAK\d+X\d+",
    r"SMA(?:-FL)?", r"SMB", r"SMC", r"DO-214(?:AA|AB|AC)?",
    r"D2PAK", r"DPAK", r"TO-252", r"TO-263", r"TO-220",
    r"\d+-PowerLFLGA", r"\d+-PowerLFQFN", r"\d+-PowerLQFN", r"\d+-BLGA",
    r"\d+-WSON", r"WSON-\d+", r"\d+-USON(?:\s*\d+X\d+)?", r"USON-\d+",
    r"LQFP-\d+", r"TQFP-\d+", r"\d+-QFP", r"QFP-\d+",
    r"QFN-\d+", r"DFN-\d+", r"BGA-\d+",
    r"PLCC-\d+",
    r"4x0402", r"4x0603", r"2x0402", r"2x0603",
    r"MELF\s*\(0207\)", r"MELF", r"0207",
    r"3225-4P", r"4-SMD", r"4-SOJ",
    r"Tantal\s+[A-E]", r"SMD\s+CASE\s+[A-E]", r"CASE\s+[A-E]"
]

# Words to ignore when detecting product part numbers in descriptions
EXCLUDE_MPN_WORDS = {
    "SMD", "TH", "PACKAGE", "CHANNEL", "P-CHANNEL", "N-CHANNEL", "DUAL", "SINGLE",
    "GENERAL", "AUTOMOTIVE", "THICK", "THIN", "FILM", "RESISTOR", "CAPACITOR", "INDUCTOR",
    "DIODE", "TRANSISTOR", "ZENER", "SCHOTTKY", "TVS", "SHUNT", "MOSFET", "NPN", "PNP",
    "NPN/PNP", "BIPOLAR", "BJT", "STANDARD", "ULTRA", "FAST", "SWITCHING",
    "LED", "AMBER", "GREEN", "RED", "BLUE", "WHITE", "YELLOW", "OPTOCOUPLER", "OPTOISOLATOR",
    "REGULATOR", "LINEAR", "BUCK", "VOLTAGE", "REFERENCE", "COMPARATOR", "ADJUSTABLE",
    "DRIVER", "TRANSCEIVER", "TRANSCIEVER", "AMPLIFIER", "AMPLIFIERS", "BUFFER",
    "NON-INVERTING", "ISOLATORS", "DIGITAL", "MEMORY", "FLASH", "EEPROM", "MCU",
    "CRYSTAL", "OSCILLATOR", "FUSE", "FUSES", "CHIP", "FERRITE", "BEAD", "CONNECTOR",
    "RECEPTACLES", "MODULE", "POWER", "MANAGEMENT", "SWITCH", "CONVERTER", "DC/DC",
    "BATTERY", "CHARGERS", "PROTECTION", "CAN", "CANBUS", "RS485", "USB", "SERIAL", "UART",
    "MOISTURE", "RESISTANT", "ANTI-SULFUR", "METAL", "ELEMENT", "SOFT", "TERMINATION",
    "MURATA", "YAGEO", "VISHAY", "ONSEMI", "DIODES", "ST", "STM", "STMICROELECTRONICS",
    "TI", "TEXAS", "INSTRUMENTS", "NXP", "NEXPERIA", "INFINEON", "MICROCHIP", "TDK",
    "TAIYO", "YUDEN", "KEMET", "AVX", "KYOCERA", "ROHM", "BOURNS", "LITTELFUSE",
    "PANASONIC", "SAMSUNG", "WURTH", "OSRAM", "LITE-ON", "ABRACON", "EPSON", "PULSE",
    "DEGSON", "CONNFLY", "KOSHIN", "JSCJ", "PANJIT", "ROYAL", "OHM", "FENGHUA",
    "HOTTECH", "WICH", "QUECTEL", "JYJE", "GCT", "TE", "CONNECTIVITY", "KEYSTONE",
    "MIKRODEV", "M-PLUSTEC", "FEKA", "DTSIS", "DTSİS"
}


def clean_part_number(s: Optional[str]) -> str:
    """
    Extracts the clean MPN from customer wrapper codes and strips reel/bin suffixes.
    Examples:
        'HM0012 (NTGS5120PT1G)' -> 'NTGS5120PT1G'
        'M-PLUSTEC(2N7002K-T1-GE3)' -> '2N7002K-T1-GE3'
        'H010101011333 (CH334F)' -> 'CH334F'
        '4E32000013(NCV317BD2TR4G)' -> 'NCV317BD2TR4G'
        'H010202005308 (SM15B-GHS-TB(LF)(SN))' -> 'SM15B-GHS-TB(LF)(SN)'
        'HM1221 (BZX384-C33,115)' -> 'BZX384-C33'
        'LA A67F BİN: AA-BB' -> 'LA A67F'
    """
    if not s:
        return ""
    text = str(s).strip()

    # Check outer parenthetical wrapper: CUSTOMER_PREFIX (MPN_WITH_OPTIONAL_NESTED_PARENS)
    m = re.match(r"^[^()]+\((.+)\)\s*$", text)
    if m:
        candidate = m.group(1).strip()
        if len(candidate) >= 2 and candidate.upper() not in {"DUMMY", "NONE/PN", "NA"}:
            text = candidate

    # Strip BIN / BİN notes from LED MPNs (e.g. 'LA G6SP.02 BIN: 8E-5F' -> 'LA G6SP.02')
    text = re.sub(r"\s+B[Iİ]N\s*:.*$", "", text, flags=re.IGNORECASE).strip()

    # Strip Nexperia/NXP reel suffix like ',115', ',125', ',215', ',135', ',315'
    text = re.sub(r",\d{3}\b", "", text).strip()

    return text


def parse_capacitance(text: str) -> Optional[str]:
    """Extracts capacitance value: 10nF, 4.7uF, 100pF, 1uF, 0.047uF, etc."""
    # Also support 4u7, 2n2 style if present
    m_split = re.search(r"\b(\d+)([pnuµ])(\d+)\b", text, re.IGNORECASE)
    if m_split:
        u = m_split.group(2).lower().replace("µ", "u")
        return f"{m_split.group(1)}.{m_split.group(3)}{u}F"

    m = re.search(r"\b(\d+(?:\.\d+)?)\s*([pnuµmk]?F)\b", text, re.IGNORECASE)
    if m:
        val, unit = m.group(1), m.group(2)
        unit_low = unit.replace("µ", "u").lower()
        if unit_low == "uf":
            unit = "uF"
        elif unit_low == "nf":
            unit = "nF"
        elif unit_low == "pf":
            unit = "pF"
        elif unit_low == "mf":
            unit = "mF"
        elif unit_low == "f":
            unit = "F"
        return f"{val}{unit}"
    return None


def parse_resistance(text: str) -> Optional[str]:
    """
    Extracts resistance value: 49R9, 25K5, 10k, 182K, 0R, 0R001, 220R, 1M, 4.7k, 0.05R, etc.
    Normalizes uppercase 'K' -> 'k' and 'r' -> 'R' for consistent formatting.
    """
    # Pattern 1: split notation like 49R9, 25K5, 2R2, 0R001, 4k7, 1K91, 33K2, 1M2
    m1 = re.search(r"\b(\d+)([RrKkMm])(\d+)\b", text)
    if m1:
        whole, letter, frac = m1.group(1), m1.group(2), m1.group(3)
        if letter in ("R", "r"):
            letter = "R"
        elif letter in ("K", "k"):
            letter = "k"
        elif letter in ("M", "m"):
            letter = "M"
        return f"{whole}{letter}{frac}"

    # Pattern 2: standard notation like 10k, 182K, 220R, 0R, 1M, 4.7k, 0.05R, 600R, 80 Ohms
    m2 = re.search(r"\b(\d+(?:\.\d+)?)\s*([RrKkMm]|ohms?|OHMS?|kOhms?|mOhms?)\b", text)
    if m2:
        val, unit = m2.group(1), m2.group(2)
        u_low = unit.lower()
        if u_low in ("ohm", "ohms", "r"):
            unit = "R"
        elif u_low in ("k", "kohm", "kohms"):
            unit = "k"
        elif u_low == "m" and unit == "M":
            unit = "M"
        elif u_low in ("mohm", "mohms"):
            unit = "mR"
        else:
            return None
        return f"{val}{unit}"
    return None


def parse_inductance(text: str) -> Optional[str]:
    """Extracts inductance: 68nH, 10uH, 2.2uH, 8.2uH, etc."""
    m = re.search(r"\b(\d+(?:\.\d+)?)\s*([nuµm]?H)\b", text, re.IGNORECASE)
    if m:
        val, unit = m.group(1), m.group(2)
        u_low = unit.replace("µ", "u").lower()
        if u_low == "uh":
            unit = "uH"
        elif u_low == "nh":
            unit = "nH"
        elif u_low == "mh":
            unit = "mH"
        elif u_low == "h":
            unit = "H"
        return f"{val}{unit}"
    return None


def parse_frequency(text: str) -> Optional[str]:
    """Extracts frequency: 12MHz, 16MHz, 50MHz, 32.768kHz, etc."""
    m = re.search(r"\b(\d+(?:\.\d+)?)\s*([kM]?Hz)\b", text, re.IGNORECASE)
    if m:
        val, unit = m.group(1), m.group(2)
        if unit.lower() == "mhz":
            unit = "MHz"
        elif unit.lower() == "khz":
            unit = "kHz"
        elif unit.lower() == "hz":
            unit = "Hz"
        return f"{val}{unit}"
    return None


def parse_voltage(text: str) -> Optional[str]:
    """Extracts voltage: 50V, 10V, 16V, 25V, 6.3V, 60V, 100V, etc."""
    # Ignore temperature/range tokens like 1.24V~18V or 1.6V-5.5V if standalone voltage exists
    m = re.search(r"(?<![~.\-\d])\b(\d+(?:\.\d+)?)\s*V(?:DC)?\b(?![~\-])", text, re.IGNORECASE)
    if not m:
        m = re.search(r"\b(\d+(?:\.\d+)?)\s*V(?:DC)?\b", text, re.IGNORECASE)
    if m:
        val = m.group(1)
        if val.endswith(".0"):
            val = val[:-2]
        return f"{val}V"
    return None


def parse_tolerance(text: str) -> Optional[str]:
    """
    Extracts percentage tolerance: %5, %1, %0.1, %10, %20, etc.
    Carefully avoids matching '4 %' inside '60R4 %1' and avoids matching '±10ppm' or '±0.1pF'.
    """
    # 1. Turkish prefix style: %1, %5, %0.1, %10, ±%1
    m_prefix = re.search(r"(?:±\s*)?%\s*(\d+(?:\.\d+)?)\b", text)
    if m_prefix:
        val = m_prefix.group(1)
        if val.endswith(".0"):
            val = val[:-2]
        return f"%{val}"

    # 2. Suffix style: 1%, 5%, 0.1%, ±1% (must be preceded by start, space, comma, semicolon, or ±)
    m_suffix = re.search(r"(?:^|[\s,;±])(\d+(?:\.\d+)?)\s*%(?!\d)", text)
    if m_suffix:
        val = m_suffix.group(1)
        if val.endswith(".0"):
            val = val[:-2]
        return f"%{val}"

    return None


def _standardize_package_name(pkg: str) -> str:
    """Standardizes common package variations into canonical DEKA/AOI names."""
    raw = pkg.strip()
    up = raw.upper()

    if "MELF" in up and "0207" in up:
        return "0207"
    if up in ("SOT23", "SOT-23") or up.startswith("SOT23-") or up.startswith("SOT-23-"):
        return "SOT-23"
    if up in ("SOT223", "SOT-223") or up.startswith("SOT-223-") or up.startswith("SOT223-"):
        return "SOT-223"
    if up in ("SOT323", "SOT-323"):
        return "SOT-323"
    if up in ("SOT363", "SOT-363"):
        return "SOT-363"
    if up in ("SOT89", "SOT-89"):
        return "SOT-89"
    if up in ("SOT563", "SOT-563"):
        return "SOT-563"
    if up in ("SOT583", "SOT-583"):
        return "SOT-583"
    if up.startswith("SOD123") or up.startswith("SOD-123"):
        return "SOD-123"
    if up.startswith("SOD323") or up.startswith("SOD-323"):
        return "SOD-323"
    if up.startswith("SOD523") or up.startswith("SOD-523"):
        return "SOD-523"
    if up.startswith("SOD80") or up.startswith("SOD-80"):
        return "SOD-80"
    if up.startswith("SMA"):
        return "SMA"
    if up == "SMB":
        return "SMB"
    if up == "SMC":
        return "SMC"
    if up in ("6-TSOP", "TSOP-6", "TSOP6"):
        return "6-TSOP"
    if up.startswith("TANTAL "):
        letter = up.split()[-1]
        return f"Tantal-{letter}"
    return raw


def extract_package(text: str) -> str:
    """Extracts package string from Canias BOM description."""
    if not text:
        return ""

    # 1. Check semiconductor/IC/special packages first (before chip sizes so '4x0402' matches before '0402')
    for pat in SEMICON_PACKAGES:
        m = re.search(r"(?<![A-Za-z0-9])" + pat + r"(?![A-Za-z0-9])", text, re.IGNORECASE)
        if m:
            return _standardize_package_name(m.group(0))

    # 2. Check standard chip sizes (0201, 0402, 0603, 0805, 1206, 1210, 1812, 2010, 2220, 2512)
    for chip in sorted(CHIP_SIZES, key=len, reverse=True):
        if re.search(r"\b" + chip + r"\b", text):
            return chip

    # 3. Fallback: inspect first comma-separated segment if it starts with 'SMD '
    first_seg = text.split(",")[0].strip()
    m_first = re.match(r"^SMD\s+([A-Za-z0-9\-_]+?)(?:\s+Package)?\s*$", first_seg, re.IGNORECASE)
    if m_first:
        cand = m_first.group(1).strip()
        if cand.upper() not in {"NON-STANDARD", "MODULE", "CHOKE", "WIRE", "RECEPTACLES", "SIM", "ALUMINIUM"}:
            return _standardize_package_name(cand)

    return ""


def extract_active_mpn(canias_code: str, mpn_raw: str, description: str, package: str) -> str:
    """
    Extracts the clean active product name / MPN for semiconductors, ICs, diodes, LEDs, etc.
    Examples:
      - Kalem 110: FETSH122M7-MIK | HM0012 (NTGS5120PT1G) | SMD 6-TSOP, P Channel, NTGS5120PT1G... -> 'NTGS5120PT1G'
      - Kalem 10:  ICRSSAC00A0L-DI | ZTL432AQFTA | SMD SOT23, Voltage Reference, Adjustable, %1... -> 'ZTL432AQFTA'
      - Kalem 20:  TRASSA2A700-MIK | HM0002 (BC807) | SMD SOT23, PNP, BC807-40, 45V... -> 'BC807-40'
    """
    clean_mpn = clean_part_number(mpn_raw)

    # Scan comma-separated tokens in description for an explicit MPN token
    tokens = [t.strip() for t in description.replace(";", ",").split(",") if t.strip()]
    desc_candidates = []

    pkg_norm_alnum = re.sub(r"[^A-Z0-9]", "", package.upper()) if package else ""

    for idx, token in enumerate(tokens):
        # Remove brand parenthetical e.g. "Mikrodev (Onsemi)"
        tok_clean = re.sub(r"\(.*?\)", "", token).strip()
        if not tok_clean:
            continue
        tok_upper = tok_clean.upper()

        # Skip the first segment if it starts with SMD / TH (that's the package segment)
        if idx == 0 and (tok_upper.startswith("SMD ") or tok_upper.startswith("TH ")):
            continue
        if tok_upper.startswith("SMD ") or tok_upper.startswith("TH "):
            continue

        # Skip if token is the package name (ignoring hyphens/spaces)
        tok_alnum = re.sub(r"[^A-Z0-9]", "", tok_upper)
        if pkg_norm_alnum and (pkg_norm_alnum == tok_alnum or tok_alnum == f"SMD{pkg_norm_alnum}"):
            continue

        # Skip BIN codes like 'Bin:AA-BB' or 'Bin:8E-5F'
        if tok_upper.startswith("BIN:") or tok_upper.startswith("BİN:"):
            continue

        # Skip electrical ratings (voltage, current, power, frequency, memory size, bits, channels, temp)
        if re.search(
            r"^\d+(?:\.\d+)?\s*(?:V|VDC|VWM|VC|A|mA|uA|W|mW|kW|°C|ohm|ohms|kOhm|mOhm|k|M|Hz|kHz|MHz|F|pF|nF|uF|H|nH|uH|mH|Bit|Bits|KB|MB|Mb|Mbit|CH|mcd|nm|h|mm|pin|LN|S)$",
            tok_clean,
            re.IGNORECASE
        ):
            continue
        if re.search(r"^\d+(?:-\d+)?\s*(?:Bit|CH)\b", tok_clean, re.IGNORECASE):
            continue
        if re.search(r"[+-]?\d+\s*°C|[~]", tok_clean):
            continue
        if re.search(r"^\d+(?:\.\d+)?\s*V\s*-\s*\d+", tok_clean, re.IGNORECASE):
            continue
        if "%" in tok_clean or "±" in tok_clean:
            continue
        if "AEC-Q" in tok_upper:
            continue

        # Strip ',115' reel suffix if present inside description token
        tok_clean = re.sub(r",\d{3}\b", "", tok_clean).strip()
        tok_upper = tok_clean.upper()

        # Check if all words in token are generic excluded words
        tok_words = [w for w in re.split(r"[\s\-_/]+", tok_upper) if w]
        if not tok_words:
            continue
        if all(w in EXCLUDE_MPN_WORDS for w in tok_words):
            continue
        # If any word is an excluded descriptive word (like 'Channel', 'Regulator', 'Output', 'Converter', 'Isolators')
        if any(w in EXCLUDE_MPN_WORDS for w in tok_words if len(w) > 2) and " " in tok_clean:
            continue

        # Must contain at least one digit and one letter, no spaces (or valid part format)
        if any(c.isdigit() for c in tok_clean) and any(c.isalpha() for c in tok_clean) and len(tok_clean) >= 3:
            desc_candidates.append(tok_clean)

    # Cross-check between clean_mpn and desc_candidates:
    if clean_mpn and clean_mpn.upper() not in EXCLUDE_MPN_WORDS:
        clean_up = clean_mpn.upper()
        for dc in desc_candidates:
            dc_up = dc.upper()
            # If desc candidate is a more specific variant of clean_mpn (e.g. BC807 -> BC807-40)
            if dc_up.startswith(clean_up) and len(dc_up) >= len(clean_up):
                return dc
            # Or if clean_mpn matches dc
            if clean_up.startswith(dc_up):
                return clean_mpn
        # Prefer clean_mpn from Üretici Kodu when it is valid
        if any(c.isdigit() for c in clean_mpn) or len(clean_mpn) >= 4:
            return clean_mpn

    if desc_candidates:
        return desc_candidates[0]

    return clean_mpn or canias_code


def parse_specs(canias_code: str, mpn_raw: str, description: str) -> ParsedSpec:
    """
    Main parser converting Canias BOM information into ParsedSpec with exact Column F and Column G rules.
    """
    c_code = (canias_code or "").strip().upper()
    mpn_str = (mpn_raw or "").strip()
    desc = (description or "").strip()
    desc_up = desc.upper()

    # 1. Determine component group and special flags (NTC, Jumper)
    is_ntc = bool(c_code.startswith("TER") or re.search(r"\bNTC\b", desc_up))
    is_jumper = bool(re.search(r"\bJUMPER\b", desc_up))

    comp_type = "OTHER"
    if c_code.startswith("CAP") or desc_up.startswith("SMD CAP") or "CAPACITOR" in desc_up or "TANTAL" in desc_up or "ALUMINIUM ELEC" in desc_up:
        comp_type = "CAPACITOR"
    elif c_code.startswith("RES") or c_code.startswith("TER") or is_ntc or "RESISTOR" in desc_up or ", NTC," in f" {desc_up}" or ", PTC," in f" {desc_up}":
        comp_type = "RESISTOR"
    elif c_code.startswith("IND") or c_code.startswith("TRF") or "INDUCTOR" in desc_up or "FERRITE" in desc_up or "F.BEAD" in desc_up or "BEAD" in desc_up or "CHOKE" in desc_up:
        comp_type = "INDUCTOR"
    elif any(c_code.startswith(p) for p in ("FET", "DIO", "TRA", "IC", "MCU", "AMP", "OPT", "VREG", "LED")):
        comp_type = "SEMICONDUCTOR"
    elif any(kw in desc_up for kw in ("MOSFET", "TRANSISTOR", "DIODE", "ZENER", "SCHOTTKY", "TVS", "IC ", "MCU", "DRIVER", "OPTOISOLATOR")):
        comp_type = "SEMICONDUCTOR"
    elif c_code.startswith("CRYS") or c_code.startswith("XTAL") or "CRYSTAL" in desc_up:
        comp_type = "CRYSTAL"
    elif c_code.startswith("FUS") or "FUSE" in desc_up:
        comp_type = "FUSE"
    elif c_code.startswith("CON") or "CONNECTOR" in desc_up or "KONNEKT" in desc_up or "HEADER" in desc_up or "RECEPTACLE" in desc_up or "SIM CARD" in desc_up:
        comp_type = "CONNECTOR"

    # 2. Extract Package
    package = extract_package(desc)

    # 3. Extract Voltage and Tolerance
    voltage = parse_voltage(desc)
    tolerance = parse_tolerance(desc)

    # 4. Form Column F (combined_package)
    # NTC kuralı: Kılıf başında R yerine NTC yazılır (örn. NTC0603)
    if comp_type == "CAPACITOR":
        if package in CHIP_SIZES:
            combined_package = f"C{package}"
        elif package:
            combined_package = package
        else:
            combined_package = "SMD"
    elif comp_type == "RESISTOR":
        if is_ntc:
            if package in CHIP_SIZES or package in ("0207", "4x0402", "4x0603"):
                combined_package = f"NTC{package}"
            elif package:
                combined_package = f"NTC_{package}"
            else:
                combined_package = "NTC"
        else:
            if package in CHIP_SIZES or package in ("0207", "4x0402", "4x0603"):
                combined_package = f"R{package}"
            elif package:
                combined_package = package
            else:
                combined_package = "SMD"
    elif comp_type == "INDUCTOR":
        if package in CHIP_SIZES:
            combined_package = f"L{package}"
        elif package:
            combined_package = package
        else:
            combined_package = "SMD"
    else:
        combined_package = package if package else "SMD"

    # 5. Determine Primary Value / Part Name & Form Column G (combined_spec)
    # Kural: % işareti kaldırılır, sadece sayısı yazılır (örn. %1 -> 1). %5 dışındakiler yazılır.
    # Jumper kuralı: Açıklamada Jumper varsa sonuna _Jumper eklenir.
    # NTC kuralı: NTC ise tolerans/sayıdan sonra _NTC eklenir.
    primary_val = ""
    clean_mpn = clean_part_number(mpn_str)

    def _format_tol(tol: Optional[str]) -> Optional[str]:
        if not tol or tol == "%5":
            return None
        return tol.replace("%", "").strip()

    tol_clean = _format_tol(tolerance)

    if comp_type == "CAPACITOR":
        c_val = parse_capacitance(desc)
        primary_val = c_val if c_val else clean_mpn

        # Capacitors: ONLY add voltage if voltage != 50V. NEVER add tolerance!
        spec_parts = [primary_val, package if package else "SMD"]
        if voltage and voltage.upper() != "50V":
            spec_parts.append(voltage)
        if is_jumper:
            spec_parts.append("Jumper")
        combined_spec = "_".join(p for p in spec_parts if p)

    elif comp_type in ("RESISTOR", "INDUCTOR", "CRYSTAL", "FUSE"):
        if comp_type == "RESISTOR":
            r_val = parse_resistance(desc)
            primary_val = r_val if r_val else clean_mpn
        elif comp_type == "INDUCTOR":
            # Try inductance (uH/nH) first; for Ferrite Beads (F.Bead: 220R / 600R), fallback to resistance!
            l_val = parse_inductance(desc)
            if not l_val:
                # Only accept resistance if it's not DC resistance 'mOhm'
                r_cand = parse_resistance(desc)
                if r_cand and not r_cand.endswith("mR"):
                    l_val = r_cand
            primary_val = l_val if l_val else clean_mpn
        elif comp_type == "CRYSTAL":
            freq_val = parse_frequency(desc)
            primary_val = freq_val if freq_val else clean_mpn
        else:
            primary_val = clean_mpn

        # Passives: NEVER add voltage! Add tolerance (without %) ONLY if tolerance != %5
        spec_parts = [primary_val, package if package else "SMD"]
        if tol_clean:
            spec_parts.append(tol_clean)
        if is_ntc:
            spec_parts.append("NTC")
        if is_jumper:
            spec_parts.append("Jumper")
        combined_spec = "_".join(p for p in spec_parts if p)

    elif comp_type == "SEMICONDUCTOR":
        # Active components: Extract Product Name / MPN (e.g. NTGS5120PT1G, ZTL432AQFTA, BC807-40)
        mpn_name = extract_active_mpn(c_code, mpn_str, desc, package)
        primary_val = mpn_name if mpn_name else clean_mpn

        # Actives: NEVER add voltage! Add tolerance ONLY if tolerance != %5
        spec_parts = [primary_val, package if package else "SMD"]
        if tol_clean:
            spec_parts.append(tol_clean)
        if is_jumper:
            spec_parts.append("Jumper")
        combined_spec = "_".join(p for p in spec_parts if p)

    else:
        primary_val = clean_mpn if clean_mpn else c_code
        spec_parts = [primary_val, package if package else "SMD"]
        if tol_clean:
            spec_parts.append(tol_clean)
        if is_jumper:
            spec_parts.append("Jumper")
        combined_spec = "_".join(p for p in spec_parts if p)

    return ParsedSpec(
        comp_type=comp_type,
        package=package,
        combined_package=combined_package,
        primary_val=primary_val,
        voltage=voltage,
        tolerance=tolerance,
        combined_spec=combined_spec,
        raw_description=desc,
        clean_mpn=clean_mpn
    )


def normalize_value_to_numeric(val_str: str) -> Optional[Tuple[Decimal, str]]:
    """
    Converts electrical values (e.g., '100nF', '0.1uF', '10000pF', '10nF', '4k7', '4.7k', '49R9')
    into numeric value (Decimal) + base unit ('F', 'R', 'H', 'HZ').
    Used for semantic equivalence checks.
    """
    if not val_str:
        return None
    s = val_str.strip().replace("µ", "u")

    # Split notation: 4k7, 25K5, 2R2, 0R001, 1M2, 4u7 -> 4.7k, 25.5k, 2.2R, 0.001R, 1.2M, 4.7uF
    m_split = re.match(r"^(\d+)([RrKkMm])(\d+)$", s)
    if m_split:
        unit_char = m_split.group(2)
        if unit_char in ("R", "r"):
            s = f"{m_split.group(1)}.{m_split.group(3)}R"
        elif unit_char in ("K", "k"):
            s = f"{m_split.group(1)}.{m_split.group(3)}kR"
        elif unit_char in ("M", "m"):
            s = f"{m_split.group(1)}.{m_split.group(3)}MR"

    m = re.match(r"^(\d+(?:\.\d+)?)\s*([pnumkKM]?)([FRH]|Hz|ohm)?$", s, re.IGNORECASE)
    if not m:
        return None

    num_part, prefix, base_unit = m.group(1), m.group(2), (m.group(3) or "").upper()
    if base_unit == "OHM":
        base_unit = "R"
    elif not base_unit:
        if prefix in ("r", "R"):
            base_unit = "R"
            prefix = ""
        elif prefix in ("k", "K", "M"):
            base_unit = "R"
        else:
            return None

    try:
        dec = Decimal(num_part)
    except InvalidOperation:
        return None

    # Normalize prefix case ('p', 'n', 'u', 'm' are lowercase; 'k'/'K' is 1e3; 'M' is 1e6 for R/Hz)
    if base_unit == "F":
        p_key = prefix.lower()
    elif base_unit in ("R", "HZ"):
        p_key = "M" if prefix == "M" else prefix.lower()
    else:
        p_key = prefix.lower()

    multiplier_map = {
        "p": Decimal("1e-12"),
        "n": Decimal("1e-9"),
        "u": Decimal("1e-6"),
        "m": Decimal("1e-3"),
        "":  Decimal("1"),
        "k": Decimal("1e3"),
        "M": Decimal("1e6"),
    }
    mult = multiplier_map.get(p_key, Decimal("1"))
    return (dec * mult, base_unit)


def are_specs_semantically_equivalent(spec_a: str, spec_b: str) -> Tuple[bool, str]:
    """
    Checks if two spec strings (Column G) are semantically identical.
    Returns (is_equivalent, note).
    e.g. ('0.1uF_0603', '100nF_0603') -> (True, 'Semantik Denk: 0.1uF == 100nF')
    """
    a_clean = (spec_a or "").strip()
    b_clean = (spec_b or "").strip()
    if a_clean.upper() == b_clean.upper():
        return (True, "Birebir Uyuşuyor")

    # Normalize away % sign for tolerance comparison
    a_norm_tol = re.sub(r"%(\d)", r"\1", a_clean)
    b_norm_tol = re.sub(r"%(\d)", r"\1", b_clean)
    if a_norm_tol.upper() == b_norm_tol.upper():
        return (True, f"Semantik Denk: {a_clean} == {b_clean} (% yazım farkı)")

    parts_a = a_clean.split("_")
    parts_b = b_clean.split("_")
    if len(parts_a) != len(parts_b):
        return (False, f"Format/Bileşen Sayısı Farklı: '{a_clean}' vs '{b_clean}'")

    if [p.upper() for p in parts_a[1:]] != [p.upper() for p in parts_b[1:]]:
        return (False, f"Kılıf/Voltaj/Tolerans Farklı: '{a_clean}' vs '{b_clean}'")

    num_a = normalize_value_to_numeric(parts_a[0])
    num_b = normalize_value_to_numeric(parts_b[0])
    if num_a and num_b and num_a[1] == num_b[1] and num_a[0] == num_b[0]:
        return (True, f"Semantik Denk: {parts_a[0]} == {parts_b[0]}")

    return (False, f"Değer/Ürün Farklı: '{parts_a[0]}' vs '{parts_b[0]}'")

"""
Data models for SMT Pick & Place to Excel/AOI automation and Canias ERP BOM validation.
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from decimal import Decimal


@dataclass
class PnPRow:
    """Represents a single component placement row from a CAD Pick & Place file."""
    designator: str          # Final RefDes (with prefix if applied, e.g., 'ST1_R31')
    raw_designator: str      # Original RefDes from file (e.g., 'R31')
    prefix: str              # Applied prefix (e.g., 'ST1' or '')
    center_x: str            # Exact coordinate string in mm with dot separator (no rounding)
    center_y: str            # Exact coordinate string in mm with dot separator (no rounding)
    rotation: str            # Angle string (0-360 with Allegro fix if needed)
    layer: str               # CAD layer string ('Top' / 'Bottom' / 'T' / 'B')
    source_file: str         # File name this row originated from
    original_line: int       # 1-based line or row index in the source file


@dataclass
class ParsedSpec:
    """Parsed technical parameters from Canias BOM item."""
    comp_type: str           # CAPACITOR, RESISTOR, INDUCTOR, SEMICONDUCTOR, CONNECTOR, OTHER
    package: str             # Normalized package string (e.g., '0603', '6-TSOP', 'SOT-23')
    combined_package: str    # Column F string (e.g., 'C0603', 'R0603', '6-TSOP')
    primary_val: str         # Electrical value (e.g. '10nF', '49R9') OR Part/MPN (e.g. 'NTGS5120PT1G')
    voltage: Optional[str]   # e.g., '50V', '10V', '60V' or None
    tolerance: Optional[str] # e.g., '%5', '%1', '%10' or None
    combined_spec: str       # Column G string (e.g., '10nF_0603', '4.7uF_0603_10V', '49R9_0603_%1', 'NTGS5120PT1G_6-TSOP')
    raw_description: str     # Original description text from BOM
    clean_mpn: Optional[str] = None # Clean manufacturer part number


@dataclass
class BomItem:
    """Represents an ERP BOM line from Canias."""
    item_no: str             # Kalem (e.g. '10', '110')
    canias_code: str         # Ü.A Eleman Num. (e.g. 'FETSH122M7-MIK')
    mpn: str                 # Üretici Kodu (e.g. 'HM0012 (NTGS5120PT1G)')
    quantity: str            # Miktar
    quantity_unit: str       # Miktar Br. ('AD', 'GR', etc.)
    description: str         # Açıklama
    alt1: str                # Alternatif 1
    raw_designators: str     # Raw RefDes text from BOM (e.g. 'FM_C5, FM_C12')
    designator_list: List[str] = field(default_factory=list) # Split normalized RefDes tokens
    is_smd: bool = True
    skip_reason: Optional[str] = None # 'TH', 'Krem Lehim', 'Dummy', 'Empty RefDes', etc.
    parsed_spec: Optional[ParsedSpec] = None


@dataclass
class MatchedRow:
    """10-column final output row for Top / Bottom Excel sheets and .AOI files."""
    designator: str          # Col A
    center_x: str            # Col B
    center_y: str            # Col C
    rotation: str            # Col D
    layer: str               # Col E (Always fixed to 'T')
    combined_package: str    # Col F
    combined_spec: str       # Col G
    canias_code: str         # Col H (Ü.A Eleman Num.)
    barcode: str             # Col I (Canias Kodu Tekrarı - Ü.A Eleman Num.)
    description: str         # Col J (BOM Açıklama)
    side: str                # 'Top' or 'Bottom'
    source_file: str = ""
    is_matched: bool = True
    unmatched_reason: Optional[str] = None

    def to_row_values(self, include_barcode: bool = True, include_description: bool = True) -> List[str]:
        """Return values for headerless Excel row or AOI line (8 to 10 columns)."""
        vals = [
            self.designator,
            self.center_x,
            self.center_y,
            self.rotation,
            self.layer,
            self.combined_package,
            self.combined_spec,
            self.canias_code,
        ]
        if include_barcode:
            vals.append(self.barcode)
        if include_description:
            vals.append(self.description)
        return vals


@dataclass
class ValidationDifference:
    """Represents a single discrepancy between prepared Final Excel and current Canias BOM."""
    refdes: str
    side: str                # 'Top', 'Bottom', or 'Genel'
    status: str              # 'UYGUN', 'HATA', 'UYARI', 'BILGI'
    category: str            # 'CANIAS_KODU', 'DEGER_PARAMETRE', 'KILIF', 'ACIKLAMA', 'BARKOD', 'EKSIK_REFDES', 'FAZLA_REFDES', 'SMD_DISI', 'FORMAT'
    detail: str              # Human readable explanation
    excel_value: str         # Current value in the prepared Excel
    bom_value: str           # Expected value from current Canias BOM
    excel_row: Optional[int] = None
    suggested_fix: Optional[Dict[str, Any]] = None


@dataclass
class ValidationSummary:
    """Summary of validation report between prepared Final Excel and current Canias BOM."""
    total_checked_refdes: int = 0
    match_count: int = 0
    error_count: int = 0
    warning_count: int = 0
    missing_in_excel_count: int = 0
    extra_in_excel_count: int = 0
    accuracy_percent: float = 0.0
    differences: List[ValidationDifference] = field(default_factory=list)

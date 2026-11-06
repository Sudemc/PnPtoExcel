"""
Matcher Pipeline.
Merges Pick & Place components with Canias ERP BOM records:
- Matches by RefDes (case-insensitive, respecting prefix).
- Strictly filters out non-SMD items (TH, Krem Lehim, Dummy, DNP, Fiducials, Test Points).
- Enforces Column E ('Layer') = 'T' for BOTH Top and Bottom sheets.
- Constructs the 10 standard output columns (A through J).
- Partitions matched components into Top and Bottom sets.
- Tracks statistics (matched, unmatched, skipped non-SMD, missing in P&P).
"""
import re
from typing import List, Dict, Tuple, Optional, Any
from dataclasses import dataclass, field
from .models import PnPRow, BomItem, MatchedRow


def natural_sort_key(s: str):
    """Sorts strings with embedded numbers naturally (e.g. R1, R2, R10)."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r"(\d+)", s)]


@dataclass
class PipelineReport:
    """Statistics and details of the matching run."""
    total_pnp_rows: int = 0
    top_placed_count: int = 0
    bottom_placed_count: int = 0
    unmatched_pnp_count: int = 0
    skipped_non_smd_count: int = 0
    bom_missing_in_pnp_count: int = 0
    unmatched_pnp_items: List[PnPRow] = field(default_factory=list)
    skipped_non_smd_items: List[Tuple[PnPRow, str]] = field(default_factory=list)
    bom_missing_in_pnp_items: List[str] = field(default_factory=list)


def run_matching_pipeline(
    pnp_rows: List[PnPRow],
    bom_items: List[BomItem],
    refdes_to_codes: Dict[str, List[str]],
    item_by_code: Dict[str, BomItem]
) -> Tuple[List[MatchedRow], List[MatchedRow], PipelineReport]:
    """
    Executes the matching pipeline between P&P rows and Canias BOM.
    Returns:
        top_rows: matched rows on Top layer (Col E = 'T')
        bottom_rows: matched rows on Bottom layer (Col E = 'T')
        report: PipelineReport with comprehensive statistics
    """
    top_rows: List[MatchedRow] = []
    bottom_rows: List[MatchedRow] = []
    report = PipelineReport(total_pnp_rows=len(pnp_rows))

    matched_bom_refdes = set()

    for pnp in pnp_rows:
        ref_clean = pnp.designator.strip().upper()
        
        # Check if RefDes exists in Canias BOM
        matched_key = ref_clean
        canias_codes = refdes_to_codes.get(ref_clean)
        if not canias_codes:
            # Try without prefix if prefixed
            m = re.match(r"^[A-Za-z0-9]+_(.+)$", ref_clean)
            if m:
                stripped_key = m.group(1).upper()
                canias_codes = refdes_to_codes.get(stripped_key)
                if canias_codes:
                    matched_key = stripped_key

        if not canias_codes:
            report.unmatched_pnp_count += 1
            report.unmatched_pnp_items.append(pnp)
            continue

        canias_code = canias_codes[0]
        bom_item = item_by_code.get(canias_code)
        if not bom_item:
            report.unmatched_pnp_count += 1
            report.unmatched_pnp_items.append(pnp)
            continue

        # Check if SMD
        if not bom_item.is_smd:
            report.skipped_non_smd_count += 1
            report.skipped_non_smd_items.append((pnp, bom_item.skip_reason or "SMD Değil"))
            continue

        spec = bom_item.parsed_spec
        if not spec:
            continue

        matched_bom_refdes.add(ref_clean)
        matched_bom_refdes.add(matched_key)

        # Build 10-column matched row
        # Sütun E daima sabit olarak "T" yazılır!
        # Sütun I (Barkod): Canias Stok Kodu tekrarı (Ü.A Eleman Num.)
        matched = MatchedRow(
            designator=pnp.designator,
            center_x=pnp.center_x,
            center_y=pnp.center_y,
            rotation=pnp.rotation,
            layer="T",  # Always 'T' for both Top and Bottom sheets
            combined_package=spec.combined_package,
            combined_spec=spec.combined_spec,
            canias_code=bom_item.canias_code,
            barcode=bom_item.canias_code,
            description=bom_item.description,
            side=pnp.layer,
            source_file=pnp.source_file,
            is_matched=True
        )

        if pnp.layer == "Bottom":
            bottom_rows.append(matched)
        else:
            top_rows.append(matched)

    # Sort rows naturally by RefDes
    top_rows.sort(key=lambda r: natural_sort_key(r.designator))
    bottom_rows.sort(key=lambda r: natural_sort_key(r.designator))

    report.top_placed_count = len(top_rows)
    report.bottom_placed_count = len(bottom_rows)

    # Check which SMD BOM designators were not present in P&P
    all_smd_bom_refs = set()
    for item in bom_items:
        if item.is_smd:
            for r in item.designator_list:
                all_smd_bom_refs.add(r.upper())

    missing_in_pnp = sorted(all_smd_bom_refs - matched_bom_refdes, key=natural_sort_key)
    report.bom_missing_in_pnp_count = len(missing_in_pnp)
    report.bom_missing_in_pnp_items = missing_in_pnp

    return top_rows, bottom_rows, report

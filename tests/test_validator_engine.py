"""
Unit tests for validator_engine.py (Module 2).
"""
import os
import openpyxl
from pp_to_excel.models import BomItem, ParsedSpec, MatchedRow
from pp_to_excel.validator_engine import validate_prepared_excel_against_bom
from pp_to_excel.exporter import export_validation_report_excel, export_final_excel


def test_validator_detects_changes_and_repairs():
    # 1. Create a simulated prepared Excel with deliberate discrepancies:
    # Row 1 (R1): Stock code outdated (RES_OLD instead of RES_NEW) -> HATA (CANIAS_KODU)
    # Row 2 (C1): Code matches CAP_01 -> UYGUN
    # Row 3 (C2): Code matches CAP_02 -> UYGUN
    # Row 4 (J1): TH connector mistakenly put in Excel -> HATA (SMD_DISI)
    # Row 5 (R99): Deleted from BOM -> HATA (FAZLA_REFDES)
    
    test_excel = "test_prepared_pnp.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Top"

    # Add 5 rows to Top
    rows_data = [
        # Col A: RefDes, B: X, C: Y, D: Rot, E: Layer, F: Pkg, G: Spec, H: Code, I: Barcode, J: Desc
        ["R1", "12,50", "25.00", "90", "Top", "R0603", "10k_0603", "RES_OLD", "RES_OLD-000001", "Old Resistor"],
        ["C1", "15.00", "30.00", "0", "T", "C0603", "100nF_0603", "CAP_01", "CAP_01-000001", "100nF Cap"],
        ["C2", "20.00", "40.00", "180", "T", "C0603", "4.7uF_0603_10V", "CAP_02", "CAP_02-000001", "4.7uF Cap"],
        ["J1", "50.00", "60.00", "90", "T", "CONN", "CONN_HEADER", "TH_CONN", "TH_CONN-000001", "TH Connector"],
        ["R99", "80.00", "90.00", "0", "T", "R0603", "1k_0603", "RES_OBS", "RES_OBS-000001", "Obsolete"],
    ]
    for r_idx, r_data in enumerate(rows_data, start=1):
        for c_idx, val in enumerate(r_data, start=1):
            ws.cell(r_idx, c_idx, val)

    wb.save(test_excel)

    # 2. Build current Canias BOM items
    r1_spec = ParsedSpec("RESISTOR", "0603", "R0603", "10k", None, "%5", "10k_0603", "SMD 10k")
    c1_spec = ParsedSpec("CAPACITOR", "0603", "C0603", "100nF", "50V", "%10", "100nF_0603", "SMD 100nF 50V")
    c2_spec = ParsedSpec("CAPACITOR", "0603", "C0603", "4.7uF", "10V", "%10", "4.7uF_0603_10V", "SMD 4.7uF 10V")
    c3_spec = ParsedSpec("CAPACITOR", "0603", "C0603", "22nF", "50V", "%10", "22nF_0603", "SMD 22nF 50V")

    bom_items = [
        BomItem("10", "RES_NEW", "MPN_R1", "1", "AD", "SMD 10k", "", "R1", ["R1"], True, None, r1_spec),
        BomItem("20", "CAP_01", "MPN_C1", "1", "AD", "SMD 100nF 50V", "", "C1", ["C1"], True, None, c1_spec),
        BomItem("30", "CAP_02", "MPN_C2", "1", "AD", "SMD 4.7uF 10V", "", "C2", ["C2"], True, None, c2_spec),
        BomItem("40", "CAP_03", "MPN_C3", "1", "AD", "SMD 22nF 50V", "", "C3", ["C3"], True, None, c3_spec), # Missing in Excel!
        BomItem("50", "TH_CONN", "MPN_J1", "1", "AD", "TH Pin Header", "", "J1", ["J1"], False, "TH (Through-Hole Malzeme)", None),
    ]

    refdes_to_codes = {
        "R1": ["RES_NEW"],
        "C1": ["CAP_01"],
        "C2": ["CAP_02"],
        "C3": ["CAP_03"],
        "J1": ["TH_CONN"],
    }
    item_by_code = {item.canias_code: item for item in bom_items}

    # 3. Run validation
    summary, repaired_top, repaired_bottom = validate_prepared_excel_against_bom(
        test_excel, bom_items, refdes_to_codes, item_by_code
    )

    # 4. Check findings
    categories = [d.category for d in summary.differences]
    assert "CANIAS_KODU" in categories     # R1 had RES_OLD instead of RES_NEW
    assert "SMD_DISI" in categories        # J1 was TH
    assert "FAZLA_REFDES" in categories    # R99 not in BOM
    assert "EKSIK_REFDES" in categories    # C3 in BOM but missing in Excel

    assert summary.error_count > 0
    assert summary.missing_in_excel_count == 1  # C3

    # Check auto-repair: R1 updated to RES_NEW with 3-decimal formatting
    repaired_r1 = next((r for r in repaired_top if r.designator == "R1"), None)
    assert repaired_r1 is not None
    assert repaired_r1.canias_code == "RES_NEW"
    assert repaired_r1.barcode == "RES_NEW"
    assert repaired_r1.center_x == "12.500"
    assert repaired_r1.center_y == "25.000"

    if os.path.exists(test_excel):
        os.remove(test_excel)

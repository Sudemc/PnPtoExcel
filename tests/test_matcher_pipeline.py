"""
Unit tests for matcher_pipeline.py and exporter.py
"""
import os
import openpyxl
from pp_to_excel.models import PnPRow, BomItem, ParsedSpec
from pp_to_excel.matcher_pipeline import run_matching_pipeline
from pp_to_excel.exporter import export_final_excel, export_aoi_file


def test_matching_and_layer_fixed_t():
    """Verifies that Col E is fixed to 'T' for both Top and Bottom rows."""
    pnp_rows = [
        PnPRow(designator="R1", raw_designator="R1", prefix="", center_x="10.5", center_y="20.2", rotation="90", layer="Top", source_file="top.txt", original_line=1),
        PnPRow(designator="C1", raw_designator="C1", prefix="", center_x="30.125", center_y="40.875", rotation="180", layer="Bottom", source_file="bot.txt", original_line=2),
    ]

    r1_spec = ParsedSpec(comp_type="RESISTOR", package="0603", combined_package="R0603", primary_val="10k", voltage=None, tolerance="%5", combined_spec="10k_0603", raw_description="SMD 10k")
    c1_spec = ParsedSpec(comp_type="CAPACITOR", package="0603", combined_package="C0603", primary_val="10nF", voltage="50V", tolerance="%10", combined_spec="10nF_0603", raw_description="SMD 10nF 50V")

    bom_items = [
        BomItem(item_no="10", canias_code="RES01", mpn="MPN_R1", quantity="1", quantity_unit="AD", description="SMD 10k", alt1="", raw_designators="R1", designator_list=["R1"], is_smd=True, parsed_spec=r1_spec),
        BomItem(item_no="20", canias_code="CAP01", mpn="MPN_C1", quantity="1", quantity_unit="AD", description="SMD 10nF 50V", alt1="", raw_designators="C1", designator_list=["C1"], is_smd=True, parsed_spec=c1_spec),
    ]

    refdes_to_codes = {"R1": ["RES01"], "C1": ["CAP01"]}
    item_by_code = {"RES01": bom_items[0], "CAP01": bom_items[1]}

    top_rows, bot_rows, report = run_matching_pipeline(pnp_rows, bom_items, refdes_to_codes, item_by_code)

    assert len(top_rows) == 1
    assert len(bot_rows) == 1
    assert top_rows[0].designator == "R1"
    assert top_rows[0].layer == "T"  # Col E is 'T'
    assert bot_rows[0].designator == "C1"
    assert bot_rows[0].layer == "T"  # Col E is also 'T'!

    # Test export to Excel
    out_xlsx = "test_out_final.xlsx"
    export_final_excel(top_rows, bot_rows, out_xlsx)
    assert os.path.exists(out_xlsx)

    wb = openpyxl.load_workbook(out_xlsx, data_only=True)
    assert "Top" in wb.sheetnames
    assert "Bottom" in wb.sheetnames
    ws_top = wb["Top"]
    # Row 1 is data, NOT header!
    assert ws_top.cell(1, 1).value == "R1"
    assert ws_top.cell(1, 5).value == "T"
    assert ws_top.cell(1, 6).value == "R0603"
    assert ws_top.cell(1, 7).value == "10k_0603"

    ws_bot = wb["Bottom"]
    assert ws_bot.cell(1, 1).value == "C1"
    assert ws_bot.cell(1, 5).value == "T"
    assert ws_bot.cell(1, 6).value == "C0603"
    assert ws_bot.cell(1, 7).value == "10nF_0603"

    # Test export to AOI
    out_aoi = "test_out_top.AOI"
    export_aoi_file(top_rows, out_aoi)
    assert os.path.exists(out_aoi)
    with open(out_aoi, "r", encoding="cp1254") as f:
        content = f.read()
    assert "R1\t10.5\t20.2\t90\tT\tR0603\t10k_0603\tRES01\tRES01\tSMD 10k" in content

    # Clean up test output files
    if os.path.exists(out_xlsx):
        os.remove(out_xlsx)
    if os.path.exists(out_aoi):
        os.remove(out_aoi)

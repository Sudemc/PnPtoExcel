"""
Full end-to-end integration tests using real workspace sample files.
"""
import os
import openpyxl
import pytest
from pp_to_excel.bom_parser import load_canias_bom, extract_erp_prefixes
from pp_to_excel.pnp_parser import load_pnp_file
from pp_to_excel.matcher_pipeline import run_matching_pipeline
from pp_to_excel.exporter import export_final_excel, export_aoi_file


def test_feka_full_pipeline():
    bom_file = "canıasbom feka.xlsx"
    pnp_file = r"2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places\2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places\2026-10-14 - Lower_StopTail1_Left_HW02_P&P.xlsx"

    if not os.path.exists(bom_file) or not os.path.exists(pnp_file):
        pytest.skip("Feka sample files not found")

    bom_items, ref_map, item_by_code = load_canias_bom(bom_file)
    prefixes = extract_erp_prefixes(bom_items)
    assert "ST1" in prefixes

    pnp_rows = load_pnp_file(pnp_file, prefix="ST1")
    assert len(pnp_rows) > 0

    top_rows, bottom_rows, report = run_matching_pipeline(
        pnp_rows, bom_items, ref_map, item_by_code
    )

    # In StopTail1 Left, components are on the Bottom side
    assert len(bottom_rows) > 0
    # Every row must have Layer == 'T'
    for r in bottom_rows:
        assert r.layer == "T"
        assert r.designator.startswith("ST1_")
        assert r.canias_code != ""
        assert r.combined_package != ""
        assert r.combined_spec != ""

    # Check export
    out_xlsx = "test_feka_final.xlsx"
    out_aoi = "test_feka_bot.AOI"
    export_final_excel(top_rows, bottom_rows, out_xlsx)
    export_aoi_file(bottom_rows, out_aoi)

    assert os.path.exists(out_xlsx)
    assert os.path.exists(out_aoi)

    wb = openpyxl.load_workbook(out_xlsx, data_only=True)
    ws_bot = wb["Bottom"]
    # Check row 1 is data, no header!
    val_a = ws_bot.cell(1, 1).value
    val_e = ws_bot.cell(1, 5).value
    assert val_a.startswith("ST1_")
    assert val_e == "T"

    # Cleanup
    for f in (out_xlsx, out_aoi):
        if os.path.exists(f):
            os.remove(f)


def test_mikrodev_hm1345_full_pipeline():
    bom_file = "Yeni Microsoft Excel Çalışma Sayfası (3).xlsx"
    pnp_file = "Pick Place for HM1345(V8) TOP.txt"

    if not os.path.exists(bom_file) or not os.path.exists(pnp_file):
        pytest.skip("Mikrodev sample files not found")

    bom_items, ref_map, item_by_code = load_canias_bom(bom_file)
    pnp_rows = load_pnp_file(pnp_file)
    assert len(pnp_rows) > 0

    top_rows, bottom_rows, report = run_matching_pipeline(
        pnp_rows, bom_items, ref_map, item_by_code
    )

    assert len(top_rows) > 0
    # TH components (CON1, CON4, D4, C74, C76, BAT1) must be accurately counted in skipped_non_smd_count
    assert report.skipped_non_smd_count > 0
    # Look for D46
    d46 = next((r for r in top_rows if r.designator == "D46"), None)
    assert d46 is not None
    assert d46.layer == "T"
    assert d46.canias_code == "DIOSZV0S1400-MIK"
    assert d46.barcode == "DIOSZV0S1400-MIK"
    assert "1N4148WS" in d46.combined_spec

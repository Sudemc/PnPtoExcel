"""
Unit tests for pnp_parser.py using real workspace sample files.
"""
import os
import pytest
from pp_to_excel.pnp_parser import load_pnp_file, convert_mil_to_mm, clean_angle, normalize_decimal_str


def test_convert_mil_to_mm():
    # 2526.467 mil -> 64.1722618 mm (exact, no rounding)
    res = convert_mil_to_mm("2526.467")
    assert res == "64.1722618"

    # 1000 mil -> 25.4 mm
    res2 = convert_mil_to_mm("1000")
    assert res2 == "25.4"


def test_clean_angle():
    # Cadence Allegro 1000x fix
    assert clean_angle("270000") == "270"
    assert clean_angle("180000") == "180"
    assert clean_angle("90000") == "90"
    assert clean_angle("0.000") == "0"
    # Decimal angles preserved
    assert clean_angle("111.34") == "111.34"
    assert clean_angle("291.34") == "291.34"


def test_load_altium_mil_text_file():
    fp = "Pick Place for Energy_Meter_Board(Variant of EnergyMeter).txt"
    if not os.path.exists(fp):
        pytest.skip(f"{fp} not found")

    rows = load_pnp_file(fp)
    assert len(rows) > 0
    # Check R31
    r31 = next((r for r in rows if r.raw_designator == "R31"), None)
    assert r31 is not None
    assert r31.rotation == "90"
    assert r31.layer == "Bottom"
    # Check exact mil to mm conversion: 2526.467 mil -> 64.1722618 mm
    assert r31.center_x == "64.1722618"


def test_load_altium_mm_text_file():
    fp = "Pick Place for HM1345(V8) TOP.txt"
    if not os.path.exists(fp):
        pytest.skip(f"{fp} not found")

    rows = load_pnp_file(fp)
    assert len(rows) > 0
    d46 = next((r for r in rows if r.raw_designator == "D46"), None)
    assert d46 is not None
    assert d46.rotation == "180"
    assert d46.layer == "Top"
    assert d46.center_x == "20.5792"
    assert d46.center_y == "37.4301"


def test_load_cadence_allegro_xlsx():
    fp = "VAL8_V2_UC2_V8.1/PP_TOP_VAL8_V2_UC2_V8_1--20260515.xlsx"
    if not os.path.exists(fp):
        pytest.skip(f"{fp} not found")

    rows = load_pnp_file(fp)
    assert len(rows) > 0
    ant3 = next((r for r in rows if r.raw_designator == "ANT3"), None)
    assert ant3 is not None
    # 270000 -> 270
    assert ant3.rotation == "270"
    assert ant3.layer == "Top"
    assert ant3.center_x == "-2.97"
    assert ant3.center_y == "48.34"


def test_load_feka_pnp_with_prefix():
    fp = r"2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places\2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places\2026-10-14 - Lower_StopTail1_Left_HW02_P&P.xlsx"
    if not os.path.exists(fp):
        pytest.skip(f"{fp} not found")

    rows = load_pnp_file(fp, prefix="ST1")
    assert len(rows) > 0
    c1 = next((r for r in rows if r.raw_designator == "C1"), None)
    assert c1 is not None
    assert c1.designator == "ST1_C1"
    assert c1.rotation == "270"
    assert c1.layer == "Bottom"
    assert c1.center_x == "22.098"
    assert c1.center_y == "5.858"

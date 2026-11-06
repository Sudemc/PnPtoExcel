"""
Unit tests for bom_parser.py
"""
import os
import pytest
from pp_to_excel.bom_parser import load_canias_bom, extract_erp_prefixes, split_references, is_smd_item


def test_split_references():
    refs = split_references("FM_C5, FM_C12, TR1_C11")
    assert refs == ["FM_C5", "FM_C12", "TR1_C11"]

    # Space after prefix
    refs2 = split_references("TR1_ C1, TR1_ C2")
    assert refs2 == ["TR1_C1", "TR1_C2"]

    # Range
    refs3 = split_references("R1..R5")
    assert refs3 == ["R1", "R2", "R3", "R4", "R5"]


def test_is_smd_item():
    # SMD item
    is_smd, reason = is_smd_item(
        canias_code="CAPSS41003C7KA0000GB-MT",
        mpn="GCM188R71H103KA37D",
        description="SMD 0603 Package, 10nF, 50V, %10, X7R, Murata",
        unit="AD",
        raw_refs="C1, C2"
    )
    assert is_smd is True
    assert reason is None

    # TH item
    is_smd, reason = is_smd_item(
        canias_code="CONTACA61211040-MIK",
        mpn="2EDGK-3.81-12P",
        description="TH PCB Klemens, Erkek, 3.81mm, 12 pin, 90°(Yatık), Mikrodev (Degson)",
        unit="AD",
        raw_refs="J1"
    )
    assert is_smd is False
    assert "TH" in reason

    # Solder paste (GR)
    is_smd, reason = is_smd_item(
        canias_code="KRMLEHIM01",
        mpn="SAC305",
        description="Krem Lehim Kursunsuz",
        unit="GR",
        raw_refs=""
    )
    assert is_smd is False

    # Dummy
    is_smd, reason = is_smd_item(
        canias_code="15-LED-NONE/PN(DUMMY)",
        mpn="DUMMY",
        description="DUMMY COMPONENT",
        unit="AD",
        raw_refs="LED1"
    )
    assert is_smd is False
    assert "DUMMY" in reason


def test_load_real_feka_bom():
    file_path = "canıasbom feka.xlsx"
    if not os.path.exists(file_path):
        pytest.skip("Feka BOM file not present")
    
    items, ref_map, item_by_code = load_canias_bom(file_path)
    assert len(items) > 0
    assert "FM_C5" in ref_map
    assert "ST1_C3" in ref_map

    # Check prefix extraction
    prefixes = extract_erp_prefixes(items)
    assert "FM" in prefixes
    assert "ST1" in prefixes
    assert "F2" in prefixes

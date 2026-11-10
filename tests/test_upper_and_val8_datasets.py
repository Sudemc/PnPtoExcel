import pytest
import os
from pp_to_excel.bom_parser import load_canias_bom, extract_erp_prefixes, resolve_pnp_file_prefix
from pp_to_excel.pnp_parser import load_pnp_file
from pp_to_excel.matcher_pipeline import run_matching_pipeline
from pp_to_excel.validator_engine import validate_prepared_excel_against_bom
from pp_to_excel.pnp_verifier import verify_pnp_against_prepared

VAL8_DIR = r"c:\Users\sudem\Desktop\P&PtoExcel\son gelenlerden\VAL8_V2_UC2_V8.1(1.kart)"
UPPER_DIR = r"c:\Users\sudem\Desktop\P&PtoExcel\UPPER\UPPER"

@pytest.mark.skipif(not os.path.exists(VAL8_DIR), reason="VAL8 dataset not present on local machine")
def test_val8_module1_and_module2():
    canias_path = os.path.join(VAL8_DIR, r"2.Pick&Place+Bom\CANIASBOMYEN.xlsx")
    pnp_top = os.path.join(VAL8_DIR, r"2.Pick&Place+Bom\PP_TOP_VAL8_V2_UC2_V8_1--20260515.xlsx")
    pnp_bot = os.path.join(VAL8_DIR, r"2.Pick&Place+Bom\PP_BOT_VAL8_V2_UC2_V8_1--20260515.xlsx")
    aoi_t = os.path.join(VAL8_DIR, r"DTS00006VAL8V2UC00_T-DTSIS VAL8 V2 UC2 REV.00\DTS00006VAL8V2UC00_T.AOI")
    aoi_b = os.path.join(VAL8_DIR, r"DTS00006VAL8V2UC00_B-DTSIS VAL8 V2 UC2 REV.00\DTS00006VAL8V2UC00_B.AOI")

    # Module 1
    bom_items, ref_map, item_by_code = load_canias_bom(canias_path)
    top_rows = load_pnp_file(pnp_top)
    bot_rows = load_pnp_file(pnp_bot)
    matched_top, matched_bot, report = run_matching_pipeline(top_rows + bot_rows, bom_items, ref_map, item_by_code)

    assert report.top_placed_count == 91
    assert report.bottom_placed_count == 3
    assert report.bom_missing_in_pnp_count == 0

    # Module 2
    summary, _, _ = validate_prepared_excel_against_bom([aoi_t, aoi_b], bom_items, ref_map, item_by_code)
    assert summary.total_checked_refdes == 94
    assert summary.match_count == 94
    assert summary.error_count == 0


@pytest.mark.skipif(not os.path.exists(UPPER_DIR), reason="UPPER dataset not present on local machine")
def test_upper_lr_prefix_resolution_and_matching():
    upper_canias = os.path.join(UPPER_DIR, r"2.Pick&Place+Bom\CANIASBOMUP.xlsx")
    upper_pnp_dir = os.path.join(UPPER_DIR, r"2.Pick&Place+Bom\2026-10-14 - MAN_Upper_Lamp_B_Sample_Pick&Places\2026-10-14 - MAN_Upper_Lamp_B_Sample_Pick&Places")

    bom_items, ref_map, item_by_code = load_canias_bom(upper_canias)
    refs_by_pfx = extract_erp_prefixes(bom_items)
    erp_prefixes = set(refs_by_pfx.keys())

    # Check prefix resolution
    assert resolve_pnp_file_prefix("Upper_StopTail1_Left_HW03_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "ST1L"
    assert resolve_pnp_file_prefix("Upper_StopTail1_Right_HW03_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "ST1R"
    assert resolve_pnp_file_prefix("Upper_StopTail2_Left_HW02_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "ST2L"
    assert resolve_pnp_file_prefix("Upper_StopTail2_Right_HW02_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "ST2R"
    assert resolve_pnp_file_prefix("Upper_Turn1_Left_HW03_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "T1L"
    assert resolve_pnp_file_prefix("Upper_Turn1_Right_HW03_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "T1R"
    assert resolve_pnp_file_prefix("Upper_Turn2_Left_HW03_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "T2L"
    assert resolve_pnp_file_prefix("Upper_Turn2_Right_HW03_P&P.xlsx", set(), erp_prefixes, refs_by_pfx) == "T2R"

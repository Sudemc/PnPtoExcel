import json
import os
import pytest
from fastapi.testclient import TestClient

from pp_to_excel.bom_parser import load_canias_bom, extract_erp_prefixes, resolve_pnp_file_prefix
from pp_to_excel.pnp_parser import load_pnp_file
from pp_to_excel.matcher_pipeline import run_matching_pipeline
from pp_to_excel.exporter import export_final_excel
from pp_to_excel.validator_engine import validate_prepared_excel_against_bom
from pp_to_excel.pnp_verifier import verify_pnp_against_prepared
from pp_to_excel.web import app

VAL8_DIR = r"c:\Users\sudem\Desktop\P&PtoExcel\son gelenlerden\VAL8_V2_UC2_V8.1(1.kart)"
UPPER_DIR = r"c:\Users\sudem\Desktop\P&PtoExcel\UPPER\UPPER"
LOWER_BOM = r"c:\Users\sudem\Desktop\P&PtoExcel\canıasbom feka.xlsx"
LOWER_PNP_DIR = r"c:\Users\sudem\Desktop\P&PtoExcel\2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places\2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places"


@pytest.mark.skipif(not os.path.exists(VAL8_DIR), reason="VAL8 dataset not present on local machine")
def test_val8_module1_module2_and_module3(tmp_path):
    canias_path = os.path.join(VAL8_DIR, r"2.Pick&Place+Bom\CANIASBOMYEN.xlsx")
    pnp_top = os.path.join(VAL8_DIR, r"2.Pick&Place+Bom\PP_TOP_VAL8_V2_UC2_V8_1--20260515.xlsx")
    pnp_bot = os.path.join(VAL8_DIR, r"2.Pick&Place+Bom\PP_BOT_VAL8_V2_UC2_V8_1--20260515.xlsx")
    aoi_t = os.path.join(VAL8_DIR, r"DTS00006VAL8V2UC00_T-DTSIS VAL8 V2 UC2 REV.00\DTS00006VAL8V2UC00_T.AOI")
    aoi_b = os.path.join(VAL8_DIR, r"DTS00006VAL8V2UC00_B-DTSIS VAL8 V2 UC2 REV.00\DTS00006VAL8V2UC00_B.AOI")
    aoi_t_xlsx = os.path.join(VAL8_DIR, r"DTS00006VAL8V2UC00_T-DTSIS VAL8 V2 UC2 REV.00\DTS00006VAL8V2UC00_T.AOI.xlsx")
    aoi_b_xlsx = os.path.join(VAL8_DIR, r"DTS00006VAL8V2UC00_B-DTSIS VAL8 V2 UC2 REV.00\DTS00006VAL8V2UC00_B.AOI.xlsx")

    # 1. Module 1
    bom_items, ref_map, item_by_code = load_canias_bom(canias_path)
    top_rows = load_pnp_file(pnp_top)
    bot_rows = load_pnp_file(pnp_bot)
    matched_top, matched_bot, report = run_matching_pipeline(top_rows + bot_rows, bom_items, ref_map, item_by_code)

    assert report.top_placed_count == 91
    assert report.bottom_placed_count == 3
    assert report.bom_missing_in_pnp_count == 0
    assert report.unmatched_pnp_count == 26  # 26 DNP items in customer P&P

    # 2. Module 2 on 7-column .AOI files
    summary, _, _ = validate_prepared_excel_against_bom([aoi_t, aoi_b], bom_items, ref_map, item_by_code)
    assert summary.total_checked_refdes == 94
    assert summary.match_count == 94
    assert summary.error_count == 0

    # 3. Module 2 on 7-column headerless .AOI.xlsx and hybrid PP_TOP/PP_BOT (Sayfa1)
    summary_xlsx, _, _ = validate_prepared_excel_against_bom([aoi_t_xlsx, aoi_b], bom_items, ref_map, item_by_code)
    assert summary_xlsx.total_checked_refdes == 94
    assert summary_xlsx.match_count == 94
    assert summary_xlsx.error_count == 0

    summary_hybrid, _, _ = validate_prepared_excel_against_bom([pnp_top, pnp_bot], bom_items, ref_map, item_by_code)
    assert summary_hybrid.total_checked_refdes == 94
    assert summary_hybrid.match_count == 94
    assert summary_hybrid.error_count == 0

    # 4. Module 2 mutation test on 7-column .AOI: ensure wrong or missing Canias code in col[6] IS caught!
    mutated_aoi = tmp_path / "mutated_T.AOI"
    with open(aoi_t, "r", encoding="utf-8") as f:
        lines = f.readlines()
    parts0 = lines[0].rstrip("\r\n").split("\t")
    assert len(parts0) == 7
    parts0[6] = "WRONG_CODE_999"
    lines[0] = "\t".join(parts0) + "\n"
    mutated_aoi.write_text("".join(lines), encoding="utf-8")

    summary_mut, _, _ = validate_prepared_excel_against_bom([str(mutated_aoi), aoi_b], bom_items, ref_map, item_by_code)
    assert summary_mut.total_checked_refdes == 94
    assert summary_mut.error_count == 1
    assert summary_mut.match_count == 93
    assert summary_mut.differences[0].category == "CANIAS_KODU"

    # 5. Module 3 (Customer P&P vs Prepared .AOI) WITH BOM
    v_sum = verify_pnp_against_prepared([pnp_top, pnp_bot], [aoi_t, aoi_b], bom_path=canias_path)
    assert v_sum.missing_in_prepared_count == 0
    assert v_sum.extra_in_prepared_count == 0
    assert v_sum.skipped_non_smd_count == 28  # 26 DNP + 2 TH = 28 legitimately excluded items
    assert v_sum.coord_error_count == 7  # Down from 188 thanks to Top board Gerber panel offset detection
    info_cats = [d.category for d in v_sum.differences if d.status == "BILGI"]
    assert "PANEL_OFFSET" in info_cats
    assert info_cats.count("SMD_DISI") == 28


@pytest.mark.skipif(not os.path.exists(UPPER_DIR), reason="UPPER dataset not present on local machine")
def test_upper_module1_module2_module3_and_web_api():
    upper_canias = os.path.join(UPPER_DIR, r"2.Pick&Place+Bom\CANIASBOMUP.xlsx")
    upper_pnp_dir = os.path.join(
        UPPER_DIR,
        r"2.Pick&Place+Bom\2026-10-14 - MAN_Upper_Lamp_B_Sample_Pick&Places\2026-10-14 - MAN_Upper_Lamp_B_Sample_Pick&Places",
    )

    bom_items, ref_map, item_by_code = load_canias_bom(upper_canias)
    refs_by_pfx = extract_erp_prefixes(bom_items)
    erp_prefixes = set(refs_by_pfx.keys())

    # 1. Check L/R prefix resolution
    expected_prefixes = {
        "2026-10-14 - Upper_StopTail1_Left_HW03_P&P.xlsx": "ST1L",
        "2026-10-14 - Upper_StopTail1_Right_HW03_P&P.xlsx": "ST1R",
        "2026-10-14 - Upper_StopTail2_Left_HW02_P&P.xlsx": "ST2L",
        "2026-10-14 - Upper_StopTail2_Right_HW02_P&P.xlsx": "ST2R",
        "2026-10-14 - Upper_Turn1_Left_HW03_P&P.xlsx": "T1L",
        "2026-10-14 - Upper_Turn1_Right_HW03_P&P.xlsx": "T1R",
        "2026-10-14 - Upper_Turn2_Left_HW03_P&P.xlsx": "T2L",
        "2026-10-14 - Upper_Turn2_Right_HW03_P&P.xlsx": "T2R",
    }
    all_pnp_rows = []
    pnp_paths = []
    for fname, exp_pfx in expected_prefixes.items():
        fpath = os.path.join(upper_pnp_dir, fname)
        pnp_paths.append(fpath)
        pfx = resolve_pnp_file_prefix(fname, set(), erp_prefixes, refs_by_pfx)
        assert pfx == exp_pfx
        all_pnp_rows.extend(load_pnp_file(fpath, prefix=pfx))

    # 2. Module 1 on all 8 boards (including T2R_LED1 L/R twin fallback)
    matched_top, matched_bot, report = run_matching_pipeline(all_pnp_rows, bom_items, ref_map, item_by_code)
    assert report.top_placed_count == 135
    assert report.bottom_placed_count == 201
    assert report.unmatched_pnp_count == 0
    # 22 items in CANIASBOMUP.xlsx are marked 'NM' (Not Mounted) in customer EBOM
    assert report.bom_missing_in_pnp_count == 22

    # 3. Module 1 single-board upload: ensure bom_missing_in_pnp is scoped to ST1L only (5 NM items on ST1L, not 22+ across all 8 boards)
    st1l_rows = [r for r in all_pnp_rows if r.designator.startswith("ST1L_")]
    _, _, report_single = run_matching_pipeline(st1l_rows, bom_items, ref_map, item_by_code)
    assert report_single.top_placed_count == 42
    assert report_single.bottom_placed_count == 56
    assert report_single.bom_missing_in_pnp_count == 5

    # 4. Module 2 on all 12 .AOI files in 2.Pick&Place+Bom
    aoi_files = []
    for root, _, files in os.walk(os.path.join(UPPER_DIR, r"2.Pick&Place+Bom")):
        for f in sorted(files):
            if f.upper().endswith(".AOI"):
                aoi_files.append(os.path.join(root, f))
    assert len(aoi_files) == 12
    summary_aoi, _, _ = validate_prepared_excel_against_bom(aoi_files, bom_items, ref_map, item_by_code)
    assert summary_aoi.total_checked_refdes == 336
    assert summary_aoi.match_count == 336
    assert summary_aoi.extra_in_excel_count == 0
    assert summary_aoi.missing_in_excel_count == 22  # Same 22 NM items in CANIASBOMUP.xlsx

    # 5. Module 2 on multi-sheet hybrid .xlsx files (Sayfa2 + Sayfa4 in ST1L, BOT + TOP in ST1R)
    st1l_xlsx = os.path.join(upper_pnp_dir, "2026-10-14 - Upper_StopTail1_Left_HW03_P&P.xlsx")
    st1r_xlsx = os.path.join(upper_pnp_dir, "2026-10-14 - Upper_StopTail1_Right_HW03_P&P.xlsx")
    summary_st1, _, _ = validate_prepared_excel_against_bom([st1l_xlsx, st1r_xlsx], bom_items, ref_map, item_by_code)
    assert summary_st1.total_checked_refdes == 196  # 98 ST1L + 98 ST1R
    assert summary_st1.match_count == 196
    assert summary_st1.extra_in_excel_count == 0

    # 6. Module 3 (Customer P&P vs Prepared .AOI) WITH and WITHOUT BOM
    v_sum_with_bom = verify_pnp_against_prepared(pnp_paths, aoi_files, bom_path=upper_canias)
    v_sum_no_bom = verify_pnp_against_prepared(pnp_paths, aoi_files, bom_path=None)
    assert v_sum_with_bom.missing_in_prepared_count == 0
    assert v_sum_with_bom.extra_in_prepared_count == 0
    assert v_sum_no_bom.missing_in_prepared_count == 0
    assert v_sum_no_bom.extra_in_prepared_count == 0
    assert v_sum_with_bom.match_count == v_sum_no_bom.match_count == 148

    # 7. Web API /api/generate test with empty string prefixes in prefixes_json
    client = TestClient(app)
    empty_prefixes = {fname: "" for fname in expected_prefixes}
    files_payload = []
    open_handles = []
    try:
        fh_bom = open(upper_canias, "rb")
        open_handles.append(fh_bom)
        files_payload.append(("bom_file", ("CANIASBOMUP.xlsx", fh_bom, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")))
        for fname in expected_prefixes:
            fh = open(os.path.join(upper_pnp_dir, fname), "rb")
            open_handles.append(fh)
            files_payload.append(("pnp_files", (fname, fh, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")))

        resp = client.post("/api/generate", files=files_payload, data={"prefixes_json": json.dumps(empty_prefixes)})
        assert resp.status_code == 200
        data = resp.json()
        assert data["report"]["top_count"] == 135
        assert data["report"]["bottom_count"] == 201
        assert data["report"]["unmatched_count"] == 0
    finally:
        for fh in open_handles:
            fh.close()


@pytest.mark.skipif(not os.path.exists(LOWER_BOM), reason="Lower Lamp dataset not present on local machine")
def test_lower_lamp_left_right_duplicate_refdes(tmp_path):
    bom_items, ref_map, item_by_code = load_canias_bom(LOWER_BOM)
    refs_by_pfx = extract_erp_prefixes(bom_items)
    erp_prefixes = set(refs_by_pfx.keys())

    pnp_files = sorted(
        os.path.join(LOWER_PNP_DIR, f)
        for f in os.listdir(LOWER_PNP_DIR)
        if f.endswith(".xlsx") and not f.startswith("~$")
    )
    assert len(pnp_files) == 16

    all_pnp_rows = []
    for fpath in pnp_files:
        raw_rows = load_pnp_file(fpath, prefix="")
        pnp_refs = {r.raw_designator for r in raw_rows}
        pfx = resolve_pnp_file_prefix(fpath, pnp_refs, erp_prefixes, refs_by_pfx)
        all_pnp_rows.extend(load_pnp_file(fpath, prefix=pfx))

    matched_top, matched_bot, report = run_matching_pipeline(all_pnp_rows, bom_items, ref_map, item_by_code)
    assert report.top_placed_count == 406
    assert report.bottom_placed_count == 578

    out_excel = str(tmp_path / "lower_lamp_out.xlsx")
    export_final_excel(matched_top, matched_bot, out_excel)

    # Module 2 on generated output with duplicate L/R RefDes
    m2_sum, _, _ = validate_prepared_excel_against_bom([out_excel], bom_items, ref_map, item_by_code)
    assert m2_sum.total_checked_refdes == 984
    assert m2_sum.match_count == 984
    assert m2_sum.extra_in_excel_count == 0

    # Module 3 on 16 Left & Right P&P files vs generated output
    m3_sum = verify_pnp_against_prepared(pnp_files, [out_excel], bom_path=LOWER_BOM)
    assert m3_sum.match_count == 984
    assert m3_sum.coord_error_count == 0
    assert m3_sum.angle_error_count == 0
    assert m3_sum.missing_in_prepared_count == 0
    assert m3_sum.extra_in_prepared_count == 0
    assert m3_sum.skipped_non_smd_count == 14

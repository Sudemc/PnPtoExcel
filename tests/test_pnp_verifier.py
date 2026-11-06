"""
Unit tests for pnp_verifier.py (Module 3).
"""
import os
import pytest
from pp_to_excel.pnp_verifier import verify_pnp_against_prepared, export_pnp_verification_report


def test_pnp_verification_catches_discrepancies():
    # Test with real files in workspace if present
    cust_file = r"2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places\2026-10-14 - MAN_Lower_Lamp_B_Sample_Pick&Places\2026-10-14 - Lower_FogMain_Left_HW02_P&P.xlsx"
    prep_file = "lastExcelBut.AOI"
    bom_file = "canıasbom feka.xlsx"

    if not os.path.exists(cust_file) or not os.path.exists(prep_file):
        pytest.skip("Test sample files not found")

    summary = verify_pnp_against_prepared(
        customer_pnp_paths=[cust_file],
        prepared_paths=[prep_file],
        bom_path=bom_file if os.path.exists(bom_file) else None
    )

    assert summary.total_customer_components == 116
    assert summary.total_prepared_components == 2
    assert summary.missing_in_prepared_count > 0

    # lastExcelBut.AOI has C1 at (24.776, 4.731) while customer P&P has (126.238, 27.139)
    # Verifier should catch X_KOORDINAT and Y_KOORDINAT mismatch for C1
    c1_diffs = [d for d in summary.differences if d.refdes == "FM_C1" or d.refdes == "C1"]
    categories = [d.category for d in c1_diffs]
    assert "X_KOORDINAT" in categories
    assert "Y_KOORDINAT" in categories

    # Export report test
    out_rep = "test_pnp_verification_report.xlsx"
    export_pnp_verification_report(summary, out_rep)
    assert os.path.exists(out_rep)
    if os.path.exists(out_rep):
        os.remove(out_rep)

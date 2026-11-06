"""
Unit tests for spec_parser.py
"""
import pytest
from pp_to_excel.spec_parser import (
    parse_specs,
    clean_part_number,
    are_specs_semantically_equivalent,
    normalize_value_to_numeric
)


def test_active_semiconductor_screenshot_rule():
    """
    Directly tests the example from user prompt & screenshot:
    Kalem 110: FETSH122M7-MIK | HM0012 (NTGS5120PT1G) | SMD 6-TSOP, P Channel, NTGS5120PT1G, 60V, 1.8A, Mikrodev (Onsemi)
    Should produce Col F: 6-TSOP, Col G: NTGS5120PT1G_6-TSOP
    """
    spec = parse_specs(
        canias_code="FETSH122M7-MIK",
        mpn_raw="HM0012 (NTGS5120PT1G)",
        description="SMD 6-TSOP, P Channel, NTGS5120PT1G, 60V, 1.8A, Mikrodev (Onsemi)"
    )
    assert spec.comp_type == "SEMICONDUCTOR"
    assert spec.package == "6-TSOP"
    assert spec.combined_package == "6-TSOP"
    assert spec.primary_val == "NTGS5120PT1G"
    assert spec.combined_spec == "NTGS5120PT1G_6-TSOP"


def test_resistor_with_tolerance_not_5_percent():
    """
    Kalem 120: RESSS24990BF08GN-MIK | HM0013 (RC0603FR-0749R9L) | SMD 0603 Package, 49R9, %1, 1/10W, Thick Film, General, Mikrodev (Yageo)
    Rule: % is stripped, only number is written -> 49R9_0603_1
    """
    spec = parse_specs(
        canias_code="RESSS24990BF08GN-MIK",
        mpn_raw="HM0013 (RC0603FR-0749R9L)",
        description="SMD 0603 Package, 49R9, %1, 1/10W, Thick Film, General, Mikrodev (Yageo)"
    )
    assert spec.comp_type == "RESISTOR"
    assert spec.package == "0603"
    assert spec.combined_package == "R0603"
    assert spec.primary_val == "49R9"
    assert spec.tolerance == "%1"
    assert spec.combined_spec == "49R9_0603_1"


def test_resistor_with_5_percent_tolerance_omitted():
    """
    Resistors with %5 should NOT have tolerance in Column G.
    """
    spec = parse_specs(
        canias_code="RESSS11002FC08AU-YG",
        mpn_raw="RC0402JR-0710KL",
        description="SMD 0402 Package, 10k, %5, 1/16W, Thick Film, General, Yageo"
    )
    assert spec.combined_package == "R0402"
    assert spec.combined_spec == "10k_0402"


def test_jumper_rule():
    """
    Resistors or jumpers with 'Jumper' in description should append _Jumper to Column G.
    """
    spec = parse_specs(
        canias_code="RESSS200000FC08AU-YG",
        mpn_raw="RC0603JR-070RL",
        description="SMD 0603 Package, 0R, Jumper, 1/10W, General, Yageo"
    )
    assert spec.combined_package == "R0603"
    assert spec.combined_spec == "0R_0603_Jumper"


def test_ntc_rule():
    """
    NTC components:
    - Column F: R prefix replaced with NTC -> NTC0603
    - Column G: _NTC appended after tolerance number -> 10k_0603_1_NTC
    """
    spec = parse_specs(
        canias_code="TERSS21002FB08B-M-P",
        mpn_raw="M-PLUSTEC(NCP18XH103F03RB)",
        description="SMD 0603 Package, 10k, %1, 0.1mW, NTC, M-Plustec(Murata)"
    )
    assert spec.combined_package == "NTC0603"
    assert spec.combined_spec == "10k_0603_1_NTC"


def test_capacitor_50v_omitted():
    """
    Capacitors with 50V should omit voltage in Column G.
    """
    spec = parse_specs(
        canias_code="CAPSS41003C7KA0000GB-MT",
        mpn_raw="GRM188R71H103KA01D",
        description="SMD 0603 Package, 10nF, 50V, %10, -55°C~+125°C, X7R, Automotive, Murata"
    )
    assert spec.comp_type == "CAPACITOR"
    assert spec.combined_package == "C0603"
    assert spec.voltage == "50V"
    assert spec.combined_spec == "10nF_0603"


def test_capacitor_non_50v_included():
    """
    Capacitors with voltage != 50V must include voltage, but NEVER tolerance.
    """
    spec = parse_specs(
        canias_code="CAPSS44704C7KA0000GB-MT",
        mpn_raw="GRM188R61A475KE15D",
        description="SMD 0603 Package, 4.7uF, 10V, %10, X5R, Murata"
    )
    assert spec.combined_package == "C0603"
    assert spec.combined_spec == "4.7uF_0603_10V"


def test_clean_part_number():
    assert clean_part_number("HM0012 (NTGS5120PT1G)") == "NTGS5120PT1G"
    assert clean_part_number("M-PLUSTEC(2N7002K-T1-GE3)") == "2N7002K-T1-GE3"
    assert clean_part_number("4E32000013(NCV317BD2TR4G)") == "NCV317BD2TR4G"
    assert clean_part_number("RC0603FR-0749R9L") == "RC0603FR-0749R9L"


def test_semantic_equivalence():
    equiv, msg = are_specs_semantically_equivalent("0.1uF_0603", "100nF_0603")
    assert equiv is True

    equiv, msg = are_specs_semantically_equivalent("4k7_0603_1", "4.7k_0603_1")
    assert equiv is True

    equiv, msg = are_specs_semantically_equivalent("49R9_0603_1", "49R9_0603_%1")
    assert equiv is True

    equiv, msg = are_specs_semantically_equivalent("10nF_0603", "100nF_0603")
    assert equiv is False


def test_edge_cases_from_real_boms():
    # 1. Uppercase K in resistor (25K5 -> 25k5, 182K -> 182k)
    s1 = parse_specs(
        canias_code="RESSS22552BF08GN-MIK",
        mpn_raw="HM0035 (RC0603FR-0725K5L)",
        description="SMD 0603 Package, 25K5 %1, 1/10W, Thick Film, General, Mikrodev (Yageo)"
    )
    assert s1.combined_package == "R0603"
    assert s1.combined_spec == "25k5_0603_1"

    # 2. Resistor value ending with digit before space + %1 ('60R4 %1' must be 1, NOT 4)
    s2 = parse_specs(
        canias_code="RESSS26040BF08GN-MIK",
        mpn_raw="HM0038 (RC0603FR-0760R4L)",
        description="SMD 0603 Package, 60R4 %1, 1/10W, Thick Film, General, Mikrodev (Yageo)"
    )
    assert s2.tolerance == "%1"
    assert s2.combined_spec == "60R4_0603_1"

    # 3. Crystal with ±10ppm must NOT extract %10
    s3 = parse_specs(
        canias_code="CRYSS01202HC02B-MIK",
        mpn_raw="HM0083 (Q22FA1280050900)",
        description="SMD 3225-4P, 12MHz, 12pF, ±10ppm, 60Ohm, -40°C~85°C, Mikrodev (Epson)"
    )
    assert s3.tolerance is None
    assert s3.combined_spec == "12MHz_3225-4P"

    # 4. Ferrite Bead (IND* with 220R -> 220R_0805)
    s4 = parse_specs(
        canias_code="INDSS22200FC08AU-TD",
        mpn_raw="MPZ2012S221AT000",
        description="SMD 0805 Package, F.Bead: 220R, 2A, 100MHz, 50mR, TDK"
    )
    assert s4.combined_package == "L0805"
    assert s4.combined_spec == "220R_0805"

    # 5. Nested parentheses in MPN & LED BIN removal
    assert clean_part_number("H010202005308 (SM15B-GHS-TB(LF)(SN))") == "SM15B-GHS-TB(LF)(SN)"
    assert clean_part_number("LA A67F BİN: AA-BB") == "LA A67F"
    assert clean_part_number("M-PLUSTEC(BC807-40,115)") == "BC807-40"

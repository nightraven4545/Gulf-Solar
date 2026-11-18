"""The workbook must be what build_model writes, and its recalculated numbers must match lcoe.py."""

import csv
import re
from pathlib import Path

import pytest
from openpyxl import load_workbook

from nrep import build_model
from nrep.lcoe import implied_capex, irr, lcoe, project_cash_flows, project_finance, real_from_flat

ROOT = Path(__file__).resolve().parents[1]
MODEL = ROOT / "model" / "NREP_Tariff_LCOE_Model.xlsx"
A = {r["id"]: float(r["value"]) for r in csv.DictReader((ROOT / "data" / "assumptions.csv").read_text().splitlines())}
PVGIS = {p: float(r["yield_kwh_kwp_yr"]) for r in csv.DictReader((ROOT / "data" / "yield.csv").read_text().splitlines())
         for p in r["projects"].split(";")}
CAL = A["SAKAKA_GWH"] * 1e6 / (A["SAKAKA_MW"] * 1e3) / PVGIS["SAKAKA"]
KD = A["BASE_RATE"] + A["MARGIN"]
WACC = A["GEAR_WACC"] * KD * (1 - A["TAX_KSA"]) + (1 - A["GEAR_WACC"]) * A["EQ_HURDLE"]
KSA = {"opex": A["OPEX"], "years": 25, "degr": A["DEGR"], "opex_esc": A["OPEX_ESC"], "tax": A["TAX_KSA"],
       "dep_years": 25}
NAJRAN_T, NAJRAN_E = 0.0109682, PVGIS["NAJRAN"] * CAL
CAPEX_2025 = 510.0  # mean of the five Saudi 2025 disclosures in capex_evidence.csv


@pytest.fixture(scope="module")
def values():
    wb = load_workbook(MODEL, data_only=True)
    if wb["Cover"]["C6"].value is None:
        pytest.skip("model has no cached values; run `make model`")
    return wb


def norm(f):
    return re.sub(r"[\s']", "", f)


def test_committed_model_matches_build(tmp_path):
    fresh = load_workbook(build_model.build(tmp_path / "m.xlsx"))
    shipped = load_workbook(MODEL)
    for ws in fresh.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    assert norm(shipped[ws.title][c.coordinate].value) == norm(c.value), (ws.title, c.coordinate)


def test_no_hardcoded_numbers_in_calc_tabs():
    allowed = {"0", "1", "2", "3", "4", "8", "10", "100", "1000", "8760", "1000000", "2024", "2025",
               "0.000001", "0.001"}
    wb = load_workbook(MODEL)
    for name in ("Engine", "Engine_CF", "PF_Model", "Bridge", "Oman", "Capex", "Checks"):
        for row in wb[name].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    body = re.sub(r"\$?[A-Z]{1,3}\$?\d+|[A-Za-z_]+\d*", "", c.value)
                    for lit in re.findall(r"\d+\.?\d*", body):
                        assert lit in allowed, (name, c.coordinate, c.value)


def test_all_checks_pass(values):
    assert values["Checks"]["C18"].value is True


def test_najran_outputs_match_reference(values):
    cover = values["Cover"]
    assert cover["C6"].value == pytest.approx(1.09682)
    assert cover["C7"].value == pytest.approx(100 * real_from_flat(NAJRAN_T, WACC, A["US_CPI"], 25, A["DEGR"]))
    assert cover["C8"].value == pytest.approx(WACC)
    assert cover["C9"].value == pytest.approx(implied_capex(NAJRAN_T, NAJRAN_E, rate=WACC, **KSA))
    assert cover["C10"].value == pytest.approx(CAPEX_2025)
    flows = project_cash_flows(NAJRAN_T, CAPEX_2025, NAJRAN_E, **KSA)
    assert cover["C11"].value == pytest.approx(irr(flows), abs=1e-7)


def test_pf_equity_irr_matches_reference(values):
    pf = project_finance(NAJRAN_T, CAPEX_2025, NAJRAN_E, debt_rate=KD, tenor=int(A["TENOR"]),
                         dscr=A["DSCR"], gear_max=A["GEAR_MAX"], **KSA)
    assert values["Cover"]["C12"].value == pytest.approx(pf["equity_irr"], abs=1e-7)


def test_oman_commercial_tariff_matches_reference(values):
    kd = KD + A["DS_OMN"] - A["DS_KSA"]
    ke = A["EQ_HURDLE"] + A["CRP_OMN"] - A["CRP_KSA"]
    w = A["GEAR_WACC"] * kd * (1 - A["TAX_OMN"]) + (1 - A["GEAR_WACC"]) * ke
    n = int(A["OMN_PPA"])
    ref = lcoe(A["OMN_CAPEX"], PVGIS["IBRI3"] * CAL, A["OPEX"], w, n, A["DEGR"], A["OPEX_ESC"], A["TAX_OMN"], n)
    assert values["Cover"]["C14"].value == pytest.approx(100 * ref)

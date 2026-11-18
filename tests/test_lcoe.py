import pytest

from nrep.lcoe import (
    implied_capex,
    irr,
    lcoe,
    project_cash_flows,
    project_finance,
    pv_factor,
    real_from_flat,
    shapley,
)

CASE = {"capex": 520.0, "energy": 3200.0, "opex": 8.0, "years": 25, "degr": 0.004, "opex_esc": 0.02,
        "tax": 0.10, "dep_years": 25}


def test_pv_factor_matches_brute_force():
    for r, n, g in [(0.06, 25, 0.0), (0.06, 25, -0.004), (0.05, 30, 0.05), (0.08, 10, 0.02)]:
        brute = sum((1 + g) ** (t - 1) / (1 + r) ** t for t in range(1, n + 1))
        assert pv_factor(r, n, g) == pytest.approx(brute, rel=1e-12)


def test_irr_known_case():
    assert irr([-100, 110]) == pytest.approx(0.10, abs=1e-9)


def test_lcoe_earns_exactly_the_rate():
    rate = 0.06
    t = lcoe(rate=rate, **CASE)
    assert irr(project_cash_flows(t, **CASE)) == pytest.approx(rate, abs=1e-9)


def test_implied_capex_inverts_lcoe():
    rate = 0.055
    t = lcoe(rate=rate, **CASE)
    kw = {k: v for k, v in CASE.items() if k != "capex"}
    assert implied_capex(t, rate=rate, **kw) == pytest.approx(CASE["capex"], rel=1e-12)


def test_flat_tariff_is_cheaper_in_real_terms():
    assert real_from_flat(0.011, 0.06, 0.023, 25) < 0.011
    assert real_from_flat(0.011, 0.06, 0.0, 25) == pytest.approx(0.011)


def test_shapley_parts_sum_to_total():
    f = lambda x: x["a"] * x["b"] / (x["c"] + 1)
    start, end = {"a": 1.0, "b": 2.0, "c": 0.5}, {"a": 3.0, "b": 0.5, "c": 2.0}
    assert sum(shapley(f, start, end).values()) == pytest.approx(f(end) - f(start))


def test_sculpted_debt_respects_covenants():
    pf = project_finance(0.0125, debt_rate=0.0525, tenor=22, dscr=1.2, gear_max=0.8, **CASE)
    assert pf["gearing"] <= 0.8 + 1e-12
    assert pf["min_dscr"] >= 1.2 - 1e-9
    assert pf["end_balance"] == pytest.approx(0.0, abs=1e-6)


def test_irena_inputs_imply_a_wacc_inside_irenas_stated_range():
    # IRENA 2025: USD 667/kWdc, 16.5% CF, USD 12.2/kWdc-yr O&M, USD 44/MWh; it states real WACCs of ~3% to >10%.
    energy = 0.165 * 8760
    lo, hi = 0.0, 0.2
    for _ in range(100):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if lcoe(667, energy, 12.2, mid, 25) < 0.044 else (lo, mid)
    assert 0.03 < lo < 0.10

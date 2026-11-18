"""Reference math for the workbook. Every formula here is mirrored on a sheet; tests hold the two together.

Units: capex USD/kW, energy kWh/kW-yr, opex USD/kW-yr, tariffs USD/kWh. Rates are decimals.
"""

from itertools import combinations
from math import factorial


def pv_factor(rate, years, growth=0.0):
    """PV of a stream that starts at 1 in year 1 and grows at `growth`: sum (1+g)^(t-1) / (1+r)^t."""
    q = (1 + growth) / (1 + rate)
    if abs(1 - q) < 1e-12:
        return years / (1 + rate)
    return q * (1 - q**years) / ((1 - q) * (1 + growth))


def lcoe(capex, energy, opex, rate, years, degr=0.0, opex_esc=0.0, tax=0.0, dep_years=None):
    """Flat tariff that returns `rate` after tax. With tax=0 this is the textbook LCOE."""
    dep_years = dep_years or years
    shield = tax * pv_factor(rate, dep_years) / dep_years
    pv_energy = energy * pv_factor(rate, years, -degr)
    pv_opex = opex * pv_factor(rate, years, opex_esc)
    return (capex * (1 - shield) + (1 - tax) * pv_opex) / ((1 - tax) * pv_energy)


def implied_capex(tariff, energy, opex, rate, years, degr=0.0, opex_esc=0.0, tax=0.0, dep_years=None):
    """Capex at which `tariff` returns exactly `rate`; the inverse of lcoe()."""
    dep_years = dep_years or years
    shield = tax * pv_factor(rate, dep_years) / dep_years
    pv_energy = energy * pv_factor(rate, years, -degr)
    pv_opex = opex * pv_factor(rate, years, opex_esc)
    return (1 - tax) * (tariff * pv_energy - pv_opex) / (1 - shield)


def project_cash_flows(tariff, capex, energy, opex, years, degr=0.0, opex_esc=0.0, tax=0.0, dep_years=None):
    """Unlevered after-tax cash flows, year 0 to `years`."""
    dep_years = dep_years or years
    flows = [-capex]
    for t in range(1, years + 1):
        ebitda = tariff * energy * (1 - degr) ** (t - 1) - opex * (1 + opex_esc) ** (t - 1)
        dep = capex / dep_years if t <= dep_years else 0.0
        flows.append(ebitda - tax * (ebitda - dep))
    return flows


def irr(flows, lo=-0.99, hi=1.0):
    """Bisection; assumes one sign change, which holds for build-then-operate flows."""
    npv = lambda r: sum(cf / (1 + r) ** t for t, cf in enumerate(flows))
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if npv(lo) * npv(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def project_finance(tariff, capex, energy, opex, years, degr, opex_esc, tax, dep_years,
                    debt_rate, tenor, dscr, gear_max):
    """Single-asset model with debt sculpted to a flat DSCR and capped by gearing.

    Sculpting uses CFADS after unlevered tax, so the schedule has no circular reference;
    the interest tax shield then flows to equity.
    """
    yrs = range(1, years + 1)
    ebitda = [tariff * energy * (1 - degr) ** (t - 1) - opex * (1 + opex_esc) ** (t - 1) for t in yrs]
    dep = [capex / dep_years if t <= dep_years else 0.0 for t in yrs]
    cfads = [e - tax * max(0.0, e - d) for e, d in zip(ebitda, dep)]
    target = [c / dscr if t <= tenor else 0.0 for t, c in zip(yrs, cfads)]
    capacity = sum(ds / (1 + debt_rate) ** t for t, ds in zip(yrs, target))
    debt = min(capacity, gear_max * capex)
    service = [ds * debt / capacity for ds in target] if capacity > 0 else [0.0] * years

    bal, equity, min_dscr = debt, [-(capex - debt)], float("inf")
    for e, d, ds, cf in zip(ebitda, dep, service, cfads):
        interest = bal * debt_rate
        bal -= ds - interest
        # shortcut: no loss carry-forward; profits are positive from year 1 at any tariff tested.
        cash_tax = tax * max(0.0, e - d - interest)
        equity.append(e - ds - cash_tax)
        if ds > 0:
            min_dscr = min(min_dscr, cf / ds)
    return {"debt": debt, "gearing": debt / capex, "equity_irr": irr(equity),
            "min_dscr": min_dscr, "end_balance": bal}


def wacc(gearing, debt_rate, equity_rate, tax):
    return gearing * debt_rate * (1 - tax) + (1 - gearing) * equity_rate


def real_from_flat(tariff, nominal_rate, inflation, years, degr=0.0):
    """Constant-real-price equivalent of a flat nominal tariff, in year-0 dollars."""
    real_rate = (1 + nominal_rate) / (1 + inflation) - 1
    flat_pv = pv_factor(nominal_rate, years, -degr)
    real_pv = pv_factor(real_rate, years, -degr)
    return tariff * flat_pv / real_pv


def shapley(f, start, end):
    """Order-independent split of f(end) - f(start) across the keys of `start`."""
    keys = list(start)
    n = len(keys)
    out = {}
    for k in keys:
        others = [x for x in keys if x != k]
        total = 0.0
        for size in range(n):
            w = factorial(size) * factorial(n - size - 1) / factorial(n)
            for subset in combinations(others, size):
                base = {x: (end[x] if x in subset else start[x]) for x in keys}
                total += w * (f({**base, k: end[k]}) - f(base))
        out[k] = total
    return out

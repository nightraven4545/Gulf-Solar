"""Export the recalculated workbook to JSON for the web case page, and copy the downloads next to it."""

import csv
import json
import shutil
from pathlib import Path

from nrep.build_deck import DECK, MODEL, ONEPAGER, Model

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
WEB = ROOT / "web" / "public"


def rows(name):
    return list(csv.DictReader((DATA / name).read_text().splitlines()))


def year(d):
    d = str(d)
    return int(d[:4]) + (int(d[5:7]) - 1) / 12 if len(d) > 4 else int(d) + 0.5


def r(x, n=6):
    return round(x, n) if isinstance(x, float) else x


def case(m):
    eng = m.table("Engine", 4)
    tr = m.table("Tracker", 4)
    najran = next(e for e in eng if e["ID"] == "NAJRAN")
    vcd = {k: v for k, v, _ in m.rows("VC_STACK")}
    wf = dict(m.rows("WATERFALL"))
    ob = {k: v for k, v, _ in m.rows("OMN_BRIDGE")}
    cases = [m[f"OMN_CASE{i}"] for i in range(1, 5)]
    oman = m.wb["Oman"]
    scen = [[c.value for c in row] for row in oman["A41:Q44"]]
    twh31 = m["OMN_SOLAR_TWH_2031"]
    omn_pv = sorted((t for t in tr if t["Country"] == "Oman" and t["Tech"] in ("PV", "PV+BESS")), key=lambda t: str(t["Date"]))
    ppa = [int(t["PPA yrs"]) for t in omn_pv if t["PPA yrs"]]

    base = {"opex": m["OPEX"], "degr": m["DEGR"], "esc": m["OPEX_ESC"]}
    presets = [
        {"key": "najran", "label": "Najran at a commercial hurdle", **base, "capex": m["CAPEX"],
         "energy": najran["Energy kWh/kWac"], "tax": m["TAX_KSA"], "years": int(najran["PPA yrs"]),
         "dep": int(m["DEP_YEARS"]), "wacc": m["WACC_KSA"], "expected": najran["Required tariff at base capex (USc)"]},
        {"key": "najran_bid", "label": "Najran at the return its bid implies", **base, "capex": m["CX_DISCLOSED_2025"],
         "energy": najran["Energy kWh/kWac"], "tax": m["TAX_KSA"], "years": int(najran["PPA yrs"]),
         "dep": int(m["DEP_YEARS"]), "wacc": m["IRR_NAJRAN"], "expected": m["T_NAJRAN"]},
    ]
    for key, label, s in (("oman_today", "Oman as today", scen[0]), ("oman_moves", "Oman with moves 1 to 3", scen[2])):
        presets.append({"key": key, "label": label, "capex": s[1], "energy": s[2], "opex": s[3], "degr": s[4],
                        "esc": s[5], "tax": s[9], "years": int(s[10]), "dep": int(s[11]), "wacc": s[12],
                        "expected": s[16]})

    moves = [("Credit support for the buyer", "Debt and equity priced nearer Saudi terms", cases[0] - cases[1],
              "SPPC: sovereign-backed single buyer", "Buyer backed by a Baa3 sovereign; Saudi Arabia is Aa3"),
             ("GW-scale, state-prepared sites", "Capex per kW; land and grid outside the bid", ob["Capex and scale"],
              "Najran 1.4 GW; PIF sites of 2-3 GW",
              f"{len(omn_pv)} projects of {min(t['MWac'] for t in omn_pv):.0f} MW each since {str(omn_pv[0]['Date'])[:4]}"),
             ("25-year PPAs (or longer)", "Spreads capex over more years", ob["Contract tenor"],
              "SPPC 25 years; three PIF projects at 35", f"PPAs of {min(ppa)} to {max(ppa)} years"),
             ("Widen the bidder pool", "Lets strategic sponsors price below commercial returns", cases[2] - cases[3],
              "Masdar, ACWA, EDF, KEPCO, SPIC, Jinko Power", ", ".join(t["Lead sponsor"] for t in omn_pv))]

    r7 = [x for x in m.rows("R7_TABLE") if x[0]]
    region = {t["ID"]: t["Region"] for t in tr}
    country = {t["ID"]: t["Country"] for t in tr}
    r6 = [e["Tariff USc/kWh"] for e in eng if e["Round"] == "R6"]
    imp = lambda p: {k.lower(): m[f"IMP_{p}_{k}"] for k in ("GW", "TWH", "KBD", "SHARE", "VALUE", "VALUE_HI", "COST", "RATIO", "CO2")}

    return {
        "asOf": "2026-11-19",
        "headline": {
            "tariff": m["T_NAJRAN"], "real": m["REAL_NAJRAN"], "irr": m["IRR_NAJRAN"], "kd": m["KD_KSA"],
            "wacc": m["WACC_KSA"], "capexDisclosed": m["CX_DISCLOSED_2025"], "capexImplied": m["IMPL_NAJRAN"],
            "irena": m["IRENA_LCOE"] / 10, "irenaKsa": m["IRENA_LCOE_KSA"] / 10, "contractedGw": m["KSA_PV_AWARDED_GW"],
            "omanCommercial": cases[0], "omanMoves": cases[2], "gas": m["OMN_AVOIDED_GAS"], "lng": m["OMN_AVOIDED_LNG"],
            "preFin": m["VC_PRE_FIN"], "left": m["VC_LEFT"], "required": m["VC_TOTAL"],
        },
        "history": [{"id": e["ID"], "project": e["Project"], "country": e["Country"], "track": e["Track"],
                     "round": e["Round"], "date": str(e["Date"]), "t": r(year(e["Date"]), 4), "mw": e["MWac"],
                     "tariff": e["Tariff USc/kWh"]} for e in eng],
        "impact": {"operating": imp("OP"), "contracted": imp("CT"), "brent": m["BRENT"], "brentHi": m["BRENT_HI"],
                   "burn": m["CRUDE_BURN_2025"]},
        "valueChain": {
            "parts": [{"name": n, "value": sum(vcd[k] for k in ks)} for n, ks in (
                ("Modules", ["Modules"]), ("Single-axis trackers", ["Single-axis trackers"]),
                ("Inverters and MV", ["Inverters and MV"]),
                ("Balance of system, civils, freight", ["Electrical BoS and civils", "Freight and duties"]),
                ("EPC margin, development, financing fees",
                 ["Installation and EPC margin", "Development and owner's costs", "Financing fees and IDC"]),
                ("O&M over 25 years", ["O&M (25 years)"]))],
            "ret": vcd["Return on capital at the hurdle WACC"], "tax": vcd["Tax and Zakat"],
            "who": [["Hardware", "Chinese makers: LONGi supplied 1.29 GW of modules for L&T's Saudi plant"],
                    ["EPC and fees", "L&T (Sudair), PowerChina, CPECC (Ibri III); sponsors' development fees"],
                    ["O&M", "Plant operator, often a sponsor affiliate"],
                    ["Return on capital", ("Banks (SNB, Riyad Bank, HSBC, KEXIM, Mizuho) and sponsors: ACWA with "
                                           "PIF's Badeel and Aramco's SAPCO; Masdar")],
                    ["Tax and Zakat", "The state (ZATCA)"]],
        },
        "whtbt": {"waccs": m.rows("WHTBT")[0][1:], "lives": [[int(x[0])] + x[1:] for x in m.rows("WHTBT")[1:]],
                  "hurdle": m["WACC_KSA"], "mw": najran["MWac"]},
        "bridge": [{"name": k, "value": v} for k, v in wf.items()],
        "round7": {"rates": {"low": m["R7_RATE_LO"], "mid": m["R7_RATE_MID"], "high": m["R7_RATE_HI"]},
                   "midAvg": m["R7_MID_AVG"], "r6": [min(r6), max(r6)],
                   "sites": [{"id": x[0], "project": x[1], "region": region[x[0]], "mw": x[2], "energy": x[3],
                              "low": x[4], "mid": x[5], "high": x[6]} for x in r7]},
        "oman": {"cases": [{"label": s[0], "value": s[16]} for s in scen], "gasPrice": m["OMN_GAS"], "heatRate": m["OMN_HR"],
                 "breakeven": [m["OMN_BE_GAS1"], m["OMN_BE_GAS3"]], "twh2031": twh31, "per01": m["OMN_IMP_PER01_2031"],
                 "gasValue": m["OMN_IMP_GAS_2031"], "lngValue": m["OMN_IMP_LNG_2031"], "movesValue": m["OMN_IMP_MOVES_2031"],
                 "moves": [{"move": a, "what": b, "cents": c, "usdm": c * twh31 * 10, "saudi": d, "oman": e}
                           for a, b, c, d, e in moves]},
        "presets": presets,
        "sites": sorted(({"id": y["ID"], "project": y["Project"].split(" (")[0], "energy": y["AC energy kWh/kWac-yr"]}
                         for y in m.table("Yield", 4) if country.get(y["ID"]) in ("KSA", "Oman")), key=lambda x: x["project"]),
        "sources": {s["id"]: {k: s[k] for k in ("publisher", "title", "date", "url")} for s in rows("sources.csv")},
    }


def site_map(m):
    tariffs = {e["ID"]: e["Tariff USc/kWh"] for e in m.table("Engine", 4)}
    y_by = {e["ID"]: e["PVGIS yield kWh/kWp-yr"] for e in m.table("Yield", 4)}
    feats = json.loads((DATA / "raw" / "borders_mena.geojson").read_text())["features"]
    rnd = lambda c: [rnd(x) for x in c] if isinstance(c[0], list) else [round(c[0], 2), round(c[1], 2)]
    for f in feats:
        f["geometry"]["coordinates"] = rnd(f["geometry"]["coordinates"])
    grid = rows("yield_grid.csv")
    pts = [(float(g["lon"]), float(g["lat"]), float(g["yield_kwh_kwp_yr"])) for g in grid]
    # Display lattice: every half-degree point near the samples takes its nearest sample's value, so the
    # clipped map has no holes at coasts or at the one cell PVGIS left empty. Stats use the samples only.
    lattice = []
    for i in range(53):
        for j in range(36):
            x, y = 34 + i * 0.5, 15.5 + j * 0.5
            d, z = min(((x - px) ** 2 + (y - py) ** 2, pz) for px, py, pz in pts)
            if d <= 1.6**2:
                lattice.append({"lon": x, "lat": y, "y": z})
    projects = [{"id": t["id"], "project": t["project"], "country": t["country"], "track": t["track"],
                 "status": t["status"], "round": t["round"], "mw": float(t["mw_ac"]), "lat": float(t["lat"]),
                 "lon": float(t["lon"]), "sponsor": t["lead_sponsor"], "tariff": r(tariffs.get(t["id"]), 5),
                 "yield": y_by.get(t["id"])}
                for t in rows("tracker.csv")
                if t["lat"] and t["tech"] in ("PV", "PV+BESS") and t["country"] in ("KSA", "Oman")]
    sau = sorted(float(g["yield_kwh_kwp_yr"]) for g in grid if g["country"] == "SAU")
    omn = sorted(float(g["yield_kwh_kwp_yr"]) for g in grid if g["country"] == "OMN")
    beats = lambda v, cells: sum(x < v for x in cells) / len(cells)
    stats = {"sauMin": sau[0], "sauMax": sau[-1], "omnMax": omn[-1], "najran": y_by["NAJRAN"], "samtah": y_by["SAMTAH"],
             "ibri": y_by["IBRI3"], "najranBeats": beats(y_by["NAJRAN"], sau), "samtahBeats": beats(y_by["SAMTAH"], sau),
             "ibriBeatsSau": beats(y_by["IBRI3"], sau), "ibriBeatsOmn": beats(y_by["IBRI3"], omn), "cells": len(grid)}
    return {"stats": stats, "borders": {"type": "FeatureCollection", "features": feats},
            "grid": lattice,
            "projects": projects}


def export(out=WEB):
    m = Model()
    (out / "data").mkdir(parents=True, exist_ok=True)
    (out / "data" / "case.json").write_text(json.dumps(case(m), separators=(",", ":"), default=r))
    sm = site_map(m)
    (out / "data" / "map.json").write_text(json.dumps(sm, separators=(",", ":")))
    c = json.loads((out / "data" / "case.json").read_text())
    c["mapStats"] = sm["stats"]
    (out / "data" / "case.json").write_text(json.dumps(c, separators=(",", ":")))
    (out / "downloads").mkdir(exist_ok=True)
    for src in (MODEL, DECK.with_suffix(".pdf"), ONEPAGER.with_suffix(".pdf")):
        shutil.copy(src, out / "downloads" / src.name)
    return out


if __name__ == "__main__":
    print(export())

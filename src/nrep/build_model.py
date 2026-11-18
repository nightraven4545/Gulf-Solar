"""Build model/NREP_Tariff_LCOE_Model.xlsx from the CSVs in data/. Every number on a calc tab is a live formula."""

import csv
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "model" / "NREP_Tariff_LCOE_Model.xlsx"
YEARS = 35  # longest PPA in the tracker

ARIAL = "Arial"
F_BASE = Font(name=ARIAL, size=10)
F_INPUT = Font(name=ARIAL, size=10, color="0000FF")
F_LINK = Font(name=ARIAL, size=10, color="008000")
F_BOLD = Font(name=ARIAL, size=10, bold=True)
F_HEAD = Font(name=ARIAL, size=10, bold=True, color="1F2A36")
F_TITLE = Font(name=ARIAL, size=14, bold=True, color="1F2A36")
F_NOTE = Font(name=ARIAL, size=9, italic=True, color="595959")
FILL_HEAD = PatternFill("solid", fgColor="DCE3EA")
FILL_KEY = PatternFill("solid", fgColor="FFF2CC")
FILL_OUT = PatternFill("solid", fgColor="E2EFDA")
THIN = Border(bottom=Side(style="thin", color="A6A6A6"))

PCT, PCT2 = "0.0%", "0.00%"
USD0, NUM0, NUM1, NUM3 = "#,##0;(#,##0);-", "#,##0", "#,##0.0", "0.000"
CENTS = '0.000 "¢"'


def load(name):
    return list(csv.DictReader((ROOT / "data" / name).read_text().splitlines()))


def num(x):
    return float(x) if x not in ("", None) else None


class Book:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

    def sheet(self, name, title, note=None, widths=None):
        ws = self.wb.create_sheet(name)
        ws.sheet_view.showGridLines = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws["A1"] = title
        ws["A1"].font = F_TITLE
        if note:
            ws["A2"] = note
            ws["A2"].font = F_NOTE
        for i, w in enumerate(widths or [], start=1):
            ws.column_dimensions[col(i)].width = w
        return ws

    def name(self, name, ref):
        self.wb.defined_names[name] = DefinedName(name, attr_text=ref)


def put(ws, ref, value, font=F_BASE, fmt=None, fill=None, bold=False):
    c = ws[ref]
    c.value = value
    c.font = Font(name=ARIAL, size=10, bold=True, color=font.color) if bold else font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def header(ws, row, labels, start=1):
    for i, label in enumerate(labels, start=start):
        c = ws.cell(row=row, column=i, value=label)
        c.font, c.fill, c.border = F_HEAD, FILL_HEAD, THIN
        c.alignment = Alignment(wrap_text=True, vertical="bottom")


def pvf(r, n, g):
    """Excel text for sum_{t=1..n} (1+g)^(t-1)/(1+r)^t, closed form."""
    q = f"((1+{g})/(1+{r}))"
    return f"({q}*(1-{q}^{n})/((1-{q})*(1+{g})))"


def annuity(r, n):
    return f"((1-(1+{r})^-{n})/{r})"


# ---------------------------------------------------------------- scenario rows
SCEN_COLS = ["Case", "Capex USD/kWac", "Energy kWh/kWac", "O&M USD/kWac-yr", "Degradation",
             "O&M escalation", "Cost of debt", "Cost of equity", "Gearing", "Tax rate", "PPA years",
             "Depreciation years", "WACC (after tax)", "PV factor: energy", "PV factor: O&M",
             "Shield factor", "Required tariff (USc/kWh)"]


def scenario_table(ws, top, rows, first_col=1):
    """One closed-form tariff per row. `rows` = [(label, [11 input formulas])]; returns {label: tariff cell}."""
    header(ws, top, SCEN_COLS, start=first_col)
    out = {}
    for i, (label, inputs) in enumerate(rows):
        r = top + 1 + i
        c = [col(first_col + k) for k in range(len(SCEN_COLS))]
        ws[f"{c[0]}{r}"] = label
        ws[f"{c[0]}{r}"].font = F_BASE
        fmts = [USD0, NUM0, NUM1, PCT2, PCT, PCT2, PCT2, PCT, PCT, NUM0, NUM0]
        for k, (f, fmt) in enumerate(zip(inputs, fmts), start=1):
            put(ws, f"{c[k]}{r}", f, F_LINK, fmt)
        cap, en, om, dg, esc, kd, ke, gr, tx, n, dep = (f"{c[k]}{r}" for k in range(1, 12))
        put(ws, f"{c[12]}{r}", f"={gr}*{kd}*(1-{tx})+(1-{gr})*{ke}", fmt=PCT2)
        w = f"{c[12]}{r}"
        put(ws, f"{c[13]}{r}", f"={pvf(w, n, '-' + dg)}", fmt=NUM3)
        put(ws, f"{c[14]}{r}", f"={pvf(w, n, esc)}", fmt=NUM3)
        put(ws, f"{c[15]}{r}", f"={tx}*{annuity(w, dep)}/{dep}", fmt=NUM3)
        pe, po, sh = f"{c[13]}{r}", f"{c[14]}{r}", f"{c[15]}{r}"
        put(ws, f"{c[16]}{r}", f"=100*({cap}*(1-{sh})+(1-{tx})*{om}*{po})/((1-{tx})*{en}*{pe})",
            fmt=CENTS, fill=FILL_OUT)
        out[label] = f"'{ws.title}'!${c[16]}${r}"
    return out


BASE = ["=CAPEX", "", "=OPEX", "=DEGR", "=OPEX_ESC", "=KD_KSA", "=EQ_HURDLE", "=GEAR_WACC", "=TAX_KSA",
        "=PPA_YEARS", "=DEP_YEARS"]


# ---------------------------------------------------------------- sheets
def build(out=OUT):
    bk = Book()
    wb = bk.wb
    assumptions = load("assumptions.csv")
    tracker = load("tracker.csv")
    sources = load("sources.csv")
    capex_ev = load("capex_evidence.csv")
    yields = {}
    for r in load("yield.csv"):
        for p in r["projects"].split(";"):
            yields[p] = r

    cover = bk.sheet("Cover", "Saudi NREP solar tariffs: what drives a ~1.1 ¢/kWh bid, and can Oman match it?",
                     widths=[3, 54, 14, 90])

    # Inputs -----------------------------------------------------------
    ws = bk.sheet("Inputs", "Inputs and assumptions",
                  "Blue = input. Low/High feed the sensitivity tab. Type: Reported (from a source), Derived (calculated from sources), Assumed.",
                  widths=[16, 11, 42, 11, 11, 11, 20, 10, 12, 80])
    header(ws, 4, ["ID", "Section", "Parameter", "Value", "Low", "High", "Unit", "Type", "Source", "Rationale"])
    fmt_for = {"per yr": PCT2, "share": PCT2, "share of capex": PCT, "AC output per kWdc": PCT}
    r = 5
    for a in assumptions:
        fmt = fmt_for.get(a["unit"], "#,##0.000" if float(a["value"]) < 10 else NUM0)
        put(ws, f"A{r}", a["id"])
        put(ws, f"B{r}", a["section"])
        put(ws, f"C{r}", a["parameter"])
        for c, k in (("D", "value"), ("E", "low"), ("F", "high")):
            put(ws, f"{c}{r}", float(a[k]), F_INPUT, fmt, FILL_KEY if c == "D" and a["type"] == "Assumed" else None)
        for c, k in (("G", "unit"), ("H", "type"), ("I", "source_id"), ("J", "rationale")):
            put(ws, f"{c}{r}", a[k])
        bk.name(a["id"], f"Inputs!$D${r}")
        bk.name(a["id"] + "_LO", f"Inputs!$E${r}")
        bk.name(a["id"] + "_HI", f"Inputs!$F${r}")
        r += 1

    r += 1
    put(ws, f"A{r}", "Derived", bold=True)
    r += 1
    header(ws, r, ["ID", "", "Parameter", "Value", "", "", "Unit", "", "", "Formula logic"])
    derived = [
        ("YIELD_CAL", "AC energy per kWac per kWh/kWp of PVGIS tracked yield",
         "=SAKAKA_GWH*1000000/(SAKAKA_MW*1000)/INDEX(Yield!$F:$F,MATCH(\"SAKAKA\",Yield!$A:$A,0))",
         "#,##0.000", "ratio", "Sakaka 2022 measured output / (MWac x PVGIS yield at Sakaka)"),
        ("KD_KSA", "Cost of senior debt (KSA)", "=BASE_RATE+MARGIN", PCT2, "per yr", "Swap + margin"),
        ("WACC_KSA", "Hurdle WACC, nominal after tax (KSA)",
         "=GEAR_WACC*KD_KSA*(1-TAX_KSA)+(1-GEAR_WACC)*EQ_HURDLE", PCT2, "per yr", "Gearing x Kd x (1-t) + equity share x Ke"),
        ("REAL_KSA", "Hurdle WACC in real terms", "=(1+WACC_KSA)/(1+US_CPI)-1", PCT2, "per yr", "Fisher"),
        ("KD_OMN", "Cost of senior debt (Oman)", "=KD_KSA+DS_OMN-DS_KSA", PCT2, "per yr", "KSA debt cost + sovereign default spread gap"),
        ("KE_OMN", "Cost of equity (Oman)", "=EQ_HURDLE+CRP_OMN-CRP_KSA", PCT2, "per yr", "KSA hurdle + country risk premium gap"),
    ]
    for name, label, f, fmt, unit, logic in derived:
        r += 1
        put(ws, f"A{r}", name)
        put(ws, f"C{r}", label)
        put(ws, f"D{r}", f, fmt=fmt, fill=FILL_OUT)
        put(ws, f"G{r}", unit)
        put(ws, f"J{r}", logic, F_NOTE)
        bk.name(name, f"Inputs!$D${r}")
    ws.freeze_panes = "D5"

    # Tracker ----------------------------------------------------------
    ws = bk.sheet("Tracker", "Tariff tracker: Saudi utility-scale PV awards and GCC comparators",
                  "Reported values in blue; column X restates every tariff in US cents/kWh at the SAR peg. In analysis = Y means a tariff with a usable site.",
                  widths=[12, 28, 7, 10, 10, 6, 8, 8, 9, 9, 9, 11, 9, 40, 11, 8, 8, 9, 6, 11, 6, 14, 50, 11])
    cols = ["id", "project", "country", "offtaker", "track", "round", "tech", "mw_ac", "tariff_usc",
            "tariff_halala", "tariff_sar", "tariff_basis", "date", "consortium", "region", "lat", "lon",
            "loc_quality", "ppa_years", "status", "in_analysis", "source_ids", "note"]
    header(ws, 4, ["ID", "Project", "Country", "Offtaker", "Track", "Round", "Tech", "MWac", "USc/kWh",
                   "Halala/kWh", "SAR/kWh", "Tariff basis", "Date", "Consortium", "Region", "Lat", "Lon",
                   "Location", "PPA yrs", "Status", "In analysis", "Sources", "Note", "Tariff USc/kWh"])
    numeric = {"mw_ac", "tariff_usc", "tariff_halala", "tariff_sar", "lat", "lon", "ppa_years"}
    for i, t in enumerate(tracker):
        r = 5 + i
        for j, k in enumerate(cols, start=1):
            v = num(t[k]) if k in numeric else t[k]
            c = ws.cell(row=r, column=j, value=v)
            c.font = F_INPUT if k in numeric or k == "date" else F_BASE
            if k == "mw_ac":
                c.number_format = NUM0
        put(ws, f"X{r}", f'=IF(I{r}<>"",I{r},IF(J{r}<>"",J{r}/FX_SAR,IF(K{r}<>"",K{r}*100/FX_SAR,"")))',
            fmt=CENTS)
    last_tr = 4 + len(tracker)
    ws.freeze_panes = "C5"

    # Yield ------------------------------------------------------------
    ws = bk.sheet("Yield", "Site resource: PVGIS 5.3 (SARAH3) horizontal single-axis tracking, calibrated to Sakaka's 2022 output",
                  "Column F is the PVGIS annual yield with zero system loss (temperature, angle and spectral losses only). Column H applies the Sakaka calibration.",
                  widths=[12, 28, 8, 8, 10, 14, 14, 14, 10])
    header(ws, 4, ["ID", "Project", "Lat", "Lon", "Location", "PVGIS yield kWh/kWp-yr", "POA irradiation kWh/m2-yr",
                   "AC energy kWh/kWac-yr", "AC capacity factor"])
    y_rows = [t for t in tracker if t["id"] in yields]
    for i, t in enumerate(y_rows):
        r = 5 + i
        yv = yields[t["id"]]
        put(ws, f"A{r}", t["id"])
        put(ws, f"B{r}", t["project"])
        put(ws, f"C{r}", float(t["lat"]), F_INPUT, "0.00")
        put(ws, f"D{r}", float(t["lon"]), F_INPUT, "0.00")
        put(ws, f"E{r}", t["loc_quality"])
        put(ws, f"F{r}", float(yv["yield_kwh_kwp_yr"]), F_INPUT, NUM0)
        put(ws, f"G{r}", float(yv["poa_kwh_m2_yr"]), F_INPUT, NUM0)
        put(ws, f"H{r}", f"=F{r}*YIELD_CAL", fmt=NUM0)
        put(ws, f"I{r}", f"=H{r}/8760", fmt=PCT)
    ws.freeze_panes = "C5"

    # Capex ------------------------------------------------------------
    ws = bk.sheet("Capex", "Capex triangulation (USD per kWac)",
                  "Four independent views of what a 2025-26 Saudi GW-scale PV plant costs.",
                  widths=[14, 32, 8, 8, 11, 9, 8, 9, 9, 11, 12])
    header(ws, 4, ["ID", "Project", "Country", "MWac", "Total cost", "Currency", "PPA yrs", "Vintage",
                   "Source", "USD m", "USD/kWac"])
    for i, e in enumerate(capex_ev):
        r = 5 + i
        put(ws, f"A{r}", e["id"])
        put(ws, f"B{r}", e["project"])
        put(ws, f"C{r}", e["country"])
        put(ws, f"D{r}", float(e["mw_ac"]), F_INPUT, NUM0)
        put(ws, f"E{r}", float(e["total_cost"]), F_INPUT, NUM0)
        put(ws, f"F{r}", e["currency"])
        put(ws, f"G{r}", num(e["ppa_years"]), F_INPUT)
        put(ws, f"H{r}", int(e["vintage"]), F_INPUT, "0")
        put(ws, f"I{r}", e["source_id"])
        put(ws, f"J{r}", f'=IF(F{r}="SAR m",E{r}/FX_SAR,E{r})', fmt=NUM0)
        put(ws, f"K{r}", f"=J{r}*1000/D{r}", fmt=NUM0)
    ev_last = 4 + len(capex_ev)
    rng = lambda c: f"Capex!${c}$5:${c}${ev_last}"

    r = ev_last + 2
    put(ws, f"A{r}", "Saudi disclosed cost by vintage (simple average)", bold=True)
    header(ws, r + 1, ["Vintage", "USD/kWac", "Projects"])
    vin_top = r + 2
    for i, yr in enumerate(range(2021, 2026)):
        rr = vin_top + i
        put(ws, f"A{rr}", yr, F_INPUT, "0")
        put(ws, f"B{rr}", f'=AVERAGEIFS({rng("K")},{rng("H")},A{rr},{rng("C")},"KSA")', fmt=NUM0)
        put(ws, f"C{rr}", f'=COUNTIFS({rng("H")},A{rr},{rng("C")},"KSA")', fmt="0")
    vin_bot = vin_top + 4
    bk.name("CX_VIN", f"Capex!$A${vin_top}:$A${vin_bot}")
    bk.name("CX_AVG", f"Capex!$B${vin_top}:$B${vin_bot}")

    r = vin_bot + 2
    put(ws, f"A{r}", "Bottom-up build at 2026 prices", bold=True)
    header(ws, r + 1, ["Item", "", "", "", "Base", "Low", "High"])
    items = [("Modules", "MODULE_PRICE*1000*DCAC"), ("Freight and duties", "CB_FREIGHT"),
             ("Single-axis trackers", "CB_TRACKER*DCAC"), ("Inverters and MV", "CB_INVERTER"),
             ("Electrical BoS and civils", "CB_BOS"), ("Installation and EPC margin", "CB_EPC"),
             ("Development and owner's costs", "CB_SOFT"), ("Financing fees and IDC", "CB_FIN")]
    top = r + 2
    for i, (label, f) in enumerate(items):
        rr = top + i
        put(ws, f"A{rr}", label)
        put(ws, f"E{rr}", "=" + f, fmt=NUM0)
        lo = f.replace("MODULE_PRICE", "MODULE_PRICE_LO").replace("CB_TRACKER", "CB_TRACKER_LO").replace("*DCAC", "*DCAC_LO")
        hi = f.replace("MODULE_PRICE", "MODULE_PRICE_HI").replace("CB_TRACKER", "CB_TRACKER_HI").replace("*DCAC", "*DCAC_HI")
        if f.startswith("CB_") and "*" not in f:
            lo, hi = f + "_LO", f + "_HI"
        put(ws, f"F{rr}", "=" + lo, fmt=NUM0)
        put(ws, f"G{rr}", "=" + hi, fmt=NUM0)
    tot = top + len(items)
    put(ws, f"A{tot}", "Total", bold=True)
    for c in "EFG":
        put(ws, f"{c}{tot}", f"=SUM({c}{top}:{c}{tot - 1})", fmt=NUM0, fill=FILL_OUT, bold=True)

    r = tot + 2
    put(ws, f"A{r}", "Football field (USD/kWac)", bold=True)
    header(ws, r + 1, ["Method", "", "", "", "Low", "High", "Point"])
    ff = r + 2
    k25 = f'{rng("H")},2025,{rng("C")},"KSA"'
    k24 = f'{rng("H")},2024,{rng("C")},"KSA"'
    football = [
        ("Disclosed, 2025 PPAs", f"=_xlfn.MINIFS({rng('K')},{k25})", f"=_xlfn.MAXIFS({rng('K')},{k25})",
         f"=AVERAGEIFS({rng('K')},{k25})"),
        ("Disclosed, 2024 financial closes", f"=_xlfn.MINIFS({rng('K')},{k24})",
         f"=_xlfn.MAXIFS({rng('K')},{k24})", f"=AVERAGEIFS({rng('K')},{k24})"),
        ("Bottom-up at 2026 prices", f"=F{tot}", f"=G{tot}", f"=E{tot}"),
        ("IRENA Saudi 2025 (commissioned)", "=IRENA_TIC_KSA*DCAC_LO", "=IRENA_TIC_KSA*DCAC_HI",
         "=IRENA_TIC_KSA*DCAC"),
        ("Implied by 2025 tariffs at hurdle WACC", "=_xlfn.MINIFS(Engine!$T:$T,Engine!$G:$G,2025)",
         "=_xlfn.MAXIFS(Engine!$T:$T,Engine!$G:$G,2025)", "=AVERAGEIFS(Engine!$T:$T,Engine!$G:$G,2025)"),
    ]
    for i, (label, lo, hi, pt) in enumerate(football):
        rr = ff + i
        put(ws, f"A{rr}", label)
        put(ws, f"E{rr}", lo, fmt=NUM0)
        put(ws, f"F{rr}", hi, fmt=NUM0)
        put(ws, f"G{rr}", pt, fmt=NUM0, fill=FILL_OUT)
    bk.name("FF_TABLE", f"Capex!$A${ff}:$G${ff + len(football) - 1}")
    bk.name("CX_DISCLOSED_2025", f"Capex!$G${ff}")
    bk.name("CX_BOTTOMUP", f"Capex!$G${ff + 2}")
    bk.name("CX_IMPLIED_2025", f"Capex!$G${ff + 4}")

    # Engine -----------------------------------------------------------
    ws = bk.sheet("Engine", "LCOE engine: what each awarded tariff implies",
                  "One row per tariff. Closed-form after-tax LCOE at the hurdle WACC. Implied capex = the capex at which the tariff earns exactly the hurdle; IRR = return at the disclosed cost of that project, or its vintage average where none is published.",
                  widths=[12, 24, 7, 9, 6, 9, 8, 7, 7, 10, 10, 8, 8, 8, 9, 8, 9, 9, 8, 11, 11, 11, 10, 10, 8, 9, 10, 11])
    hdr = ["ID", "Project", "Country", "Track", "Round", "Date", "Vintage", "MWac", "PPA yrs",
           "Tariff USc/kWh", "Energy kWh/kWac", "AC CF", "Hurdle WACC", "q energy", "PV factor energy",
           "q O&M", "PV factor O&M", "Dep. annuity", "Tax shield", "Implied capex USD/kWac",
           "Required tariff at base capex (USc)", "Disclosed capex (own, else vintage avg)", "Implied vs disclosed",
           "Project IRR at disclosed capex", "q real", "PV factor real", "Real-equivalent tariff (USc)",
           "Tariff at reference yield (USc)"]
    header(ws, 4, hdr)
    eng = [t for t in tracker if t["in_analysis"] == "Y" and t["id"] in yields]
    e_top, e_bot = 5, 4 + len(eng)
    tr = lambda c: f"Tracker!${c}$5:${c}${last_tr}"
    for i, t in enumerate(eng):
        r = e_top + i
        m = f"MATCH($A{r},{tr('A')},0)"
        put(ws, f"A{r}", t["id"])
        put(ws, f"B{r}", t["project"])
        for c, src in (("C", "C"), ("D", "E"), ("E", "F"), ("F", "M"), ("H", "H"), ("I", "S")):
            put(ws, f"{c}{r}", f"=INDEX({tr(src)},{m})", F_LINK, NUM0 if c in "HI" else None)
        put(ws, f"G{r}", f"=VALUE(LEFT(F{r},4))", fmt="0")
        put(ws, f"J{r}", f"=INDEX({tr('X')},{m})", F_LINK, CENTS)
        put(ws, f"K{r}", f"=INDEX(Yield!$H:$H,MATCH($A{r},Yield!$A:$A,0))", F_LINK, NUM0)
        put(ws, f"L{r}", f"=K{r}/8760", fmt=PCT)
        put(ws, f"M{r}", "=WACC_KSA", F_LINK, PCT2)
        put(ws, f"N{r}", f"=(1-DEGR)/(1+M{r})", fmt="0.0000")
        put(ws, f"O{r}", f"=N{r}*(1-N{r}^I{r})/((1-N{r})*(1-DEGR))", fmt=NUM3)
        put(ws, f"P{r}", f"=(1+OPEX_ESC)/(1+M{r})", fmt="0.0000")
        put(ws, f"Q{r}", f"=P{r}*(1-P{r}^I{r})/((1-P{r})*(1+OPEX_ESC))", fmt=NUM3)
        put(ws, f"R{r}", f"=(1-(1+M{r})^-DEP_YEARS)/M{r}", fmt=NUM3)
        put(ws, f"S{r}", f"=TAX_KSA*R{r}/DEP_YEARS", fmt="0.0000")
        put(ws, f"T{r}", f"=(1-TAX_KSA)*(J{r}/100*K{r}*O{r}-OPEX*Q{r})/(1-S{r})", fmt=NUM0, fill=FILL_OUT)
        put(ws, f"U{r}", f"=100*(CAPEX*(1-S{r})+(1-TAX_KSA)*OPEX*Q{r})/((1-TAX_KSA)*K{r}*O{r})", fmt=CENTS)
        put(ws, f"V{r}", f'=IFERROR(INDEX({rng("K")},MATCH(A{r},{rng("A")},0)),IFERROR(INDEX(CX_AVG,MATCH(G{r},CX_VIN,0)),""))',
            F_LINK, NUM0)
        put(ws, f"W{r}", f'=IF(V{r}="","",T{r}/V{r}-1)', fmt=PCT)
        put(ws, f"X{r}", f'=IF(V{r}="","",Engine_CF!$AP{r})', F_LINK, PCT, FILL_OUT)
        put(ws, f"Y{r}", f"=(1-DEGR)*(1+US_CPI)/(1+M{r})", fmt="0.0000")
        put(ws, f"Z{r}", f"=Y{r}*(1-Y{r}^I{r})/((1-Y{r})*(1-DEGR))", fmt=NUM3)
        put(ws, f"AA{r}", f"=J{r}*O{r}/Z{r}", fmt=CENTS)
        put(ws, f"AB{r}", f"=J{r}*K{r}/E_REF", fmt=CENTS)
        bk.name(f"E_{t['id']}", f"Engine!$K${r}")
        bk.name(f"T_{t['id']}", f"Engine!$J${r}")
        bk.name(f"IMPL_{t['id']}", f"Engine!$T${r}")
        bk.name(f"IRR_{t['id']}", f"Engine!$X${r}")
        bk.name(f"REAL_{t['id']}", f"Engine!$AA${r}")
    ref = e_bot + 2
    put(ws, f"A{ref}", "Reference yield (Saudi rows, mean)")
    put(ws, f"K{ref}", f'=AVERAGEIFS(K{e_top}:K{e_bot},C{e_top}:C{e_bot},"KSA")', fmt=NUM0, fill=FILL_OUT)
    bk.name("E_REF", f"Engine!$K${ref}")
    ws.freeze_panes = "C5"

    # Engine_CF --------------------------------------------------------
    ws = bk.sheet("Engine_CF", "Unlevered after-tax cash flows at the disclosed cost of each vintage (USD/kWac)",
                  "Feeds Engine column X. Year 0 = capex; years 1..PPA = EBITDA less tax on (EBITDA - depreciation).",
                  widths=[12, 10, 10, 10, 8] + [7] * (YEARS + 1) + [9])
    header(ws, 4, ["ID", "Tariff USD/kWh", "Energy kWh/kWac", "Capex USD/kWac", "PPA yrs"]
           + [str(t) for t in range(YEARS + 1)] + ["IRR"])
    for t in range(YEARS + 1):
        put(ws, f"{col(6 + t)}3", t, fmt="0")
    for i, t in enumerate(eng):
        r = e_top + i
        put(ws, f"A{r}", t["id"])
        put(ws, f"B{r}", f"=Engine!J{r}/100", F_LINK, "0.00000")
        put(ws, f"C{r}", f"=Engine!K{r}", F_LINK, NUM0)
        put(ws, f"D{r}", f"=Engine!V{r}", F_LINK, NUM0)
        put(ws, f"E{r}", f"=Engine!I{r}", F_LINK, NUM0)
        put(ws, f"F{r}", f'=IF($D{r}="","",-$D{r})', fmt=NUM1)
        for k in range(1, YEARS + 1):
            c = col(6 + k)
            put(ws, f"{c}{r}",
                f'=IF($D{r}="","",IF({c}$3<=$E{r},($B{r}*$C{r}*(1-DEGR)^({c}$3-1)-OPEX*(1+OPEX_ESC)^({c}$3-1))*(1-TAX_KSA)'
                f'+TAX_KSA*IF({c}$3<=DEP_YEARS,$D{r}/DEP_YEARS,0),0))', fmt=NUM1)
        put(ws, f"AP{r}", f'=IF($D{r}="","",IRR(F{r}:{col(6 + YEARS)}{r}))', fmt=PCT2, fill=FILL_OUT)
    ws.freeze_panes = "F5"

    # PF_Model ---------------------------------------------------------
    ws = bk.sheet("PF_Model", "Single-asset project finance model (per kWac)",
                  "Pick a project in C4. Debt is sculpted to the minimum DSCR on CFADS after unlevered tax, then capped by gearing. No circular references.",
                  widths=[3, 34, 12] + [8] * YEARS)
    put(ws, "B4", "Project", bold=True)
    put(ws, "C4", "NAJRAN", F_INPUT, fill=FILL_KEY)
    dv = DataValidation(type="list", formula1=f"=Engine!$A${e_top}:$A${e_bot}", allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C4")
    pf_in = [
        ("Tariff (USD/kWh)", "=INDEX(Engine!$J:$J,MATCH($C$4,Engine!$A:$A,0))/100", "0.00000"),
        ("AC energy year 1 (kWh/kWac)", "=INDEX(Engine!$K:$K,MATCH($C$4,Engine!$A:$A,0))", NUM0),
        ("PPA years", "=INDEX(Engine!$I:$I,MATCH($C$4,Engine!$A:$A,0))", NUM0),
        ("Capex (USD/kWac): disclosed for vintage, else base",
         '=IF(INDEX(Engine!$V:$V,MATCH($C$4,Engine!$A:$A,0))="",CAPEX,INDEX(Engine!$V:$V,MATCH($C$4,Engine!$A:$A,0)))', NUM0),
        ("Cost of debt", "=KD_KSA", PCT2), ("Debt tenor (yrs)", "=TENOR", NUM0), ("Minimum DSCR", "=DSCR", "0.00x"),
        ("Maximum gearing", "=GEAR_MAX", PCT), ("Tax rate", "=TAX_KSA", PCT), ("Depreciation years", "=DEP_YEARS", NUM0),
    ]
    for i, (label, f, fmt) in enumerate(pf_in):
        put(ws, f"B{5 + i}", label)
        put(ws, f"C{5 + i}", f, F_LINK, fmt)
    T, E, N, CAP, KDc, TEN, DS, GM, TX, DEP = (f"$C${5 + i}" for i in range(10))
    Y = 17
    put(ws, f"B{Y}", "Year", bold=True)
    put(ws, f"C{Y}", 0, fmt="0", bold=True)
    for k in range(1, YEARS + 1):
        put(ws, f"{col(3 + k)}{Y}", k, fmt="0", bold=True)
    rows = [
        (18, "Operating flag", "=IF({c}$17<={N},1,0)", "0"),
        (19, "Energy (kWh/kWac)", "={E}*(1-DEGR)^({c}$17-1)*{c}18", NUM0),
        (20, "Revenue", "={c}19*{T}", NUM1),
        (21, "O&M", "=OPEX*(1+OPEX_ESC)^({c}$17-1)*{c}18", NUM1),
        (22, "EBITDA", "={c}20-{c}21", NUM1),
        (23, "Tax depreciation", "=IF({c}$17<={DEP},{CAP}/{DEP},0)", NUM1),
        (24, "Unlevered tax", "={TX}*MAX(0,{c}22-{c}23)", NUM1),
        (25, "CFADS", "={c}22-{c}24", NUM1),
        (26, "Target debt service (CFADS / DSCR)", "=IF({c}$17<={TEN},{c}25/{DS},0)", NUM1),
        (27, "Discount factor at cost of debt", "=1/(1+{KD_KSA})^{c}$17", "0.000"),
        (33, "Debt service", "={c}26*$C$32", NUM1),
        (34, "Opening balance", None, NUM1),
        (35, "Interest", "={c}34*{KD_KSA}", NUM1),
        (36, "Principal", "={c}33-{c}35", NUM1),
        (37, "Closing balance", "={c}34-{c}36", NUM1),
        (38, "Taxable income", "={c}22-{c}23-{c}35", NUM1),
        (39, "Cash tax", "={TX}*MAX(0,{c}38)", NUM1),
        (40, "Equity cash flow", "={c}22-{c}33-{c}39", NUM1),
        (41, "DSCR", '=IF({c}33>0,{c}25/{c}33,"")', "0.00x"),
        (42, "Project cash flow (unlevered, after tax)", "={c}25", NUM1),
    ]
    for rr, label, f, fmt in rows:
        put(ws, f"B{rr}", label)
        for k in range(1, YEARS + 1):
            c = col(3 + k)
            if rr == 34:
                fx = "=$C$30" if k == 1 else f"={col(2 + k)}37"
            else:
                fx = f.format(c=c, N=N, E=E, T=T, DEP=DEP, CAP=CAP, TX=TX, TEN=TEN, DS=DS, KD_KSA=KDc)
            put(ws, f"{c}{rr}", fx, fmt=fmt)
    last = col(3 + YEARS)
    summary = [
        (29, "Debt capacity at target service", f"=SUMPRODUCT(D26:{last}26,D27:{last}27)", NUM1),
        (30, "Debt raised", f"=MIN(C29,{GM}*{CAP})", NUM1),
        (31, "Gearing", f"=C30/{CAP}", PCT),
        (32, "Debt service scale", "=IF(C29>0,C30/C29,0)", "0.000"),
    ]
    for rr, label, f, fmt in summary:
        put(ws, f"B{rr}", label, bold=True)
        put(ws, f"C{rr}", f, fmt=fmt)
    put(ws, "C40", f"=-({CAP}-C30)", fmt=NUM1)
    put(ws, "C42", f"=-{CAP}", fmt=NUM1)
    outs = [
        (44, "Equity IRR", f"=IFERROR(IRR(C40:{last}40),\"n/a\")", PCT2, "PF_EQ_IRR"),
        (45, "Project IRR (unlevered, after tax)", f"=IRR(C42:{last}42)", PCT2, "PF_PROJ_IRR"),
        (46, "Minimum DSCR", f"=MIN(D41:{last}41)", "0.00x", "PF_MIN_DSCR"),
        (47, "Debt balance at end of tenor", f"=INDEX(D37:{last}37,{TEN})", NUM3, "PF_END_BAL"),
        (48, "Gearing", "=C31", PCT, "PF_GEARING"),
    ]
    for rr, label, f, fmt, name in outs:
        put(ws, f"B{rr}", label, bold=True)
        put(ws, f"C{rr}", f, fmt=fmt, fill=FILL_OUT, bold=True)
        bk.name(name, f"PF_Model!$C${rr}")
    ws.freeze_panes = "D18"

    # Sensitivity ------------------------------------------------------
    ws = bk.sheet("Sensitivity", "Najran: required tariff at the hurdle WACC, one driver at a time",
                  "Each row recomputes the closed-form tariff with one input at its Low or High value. Base row equals Engine column U for Najran.",
                  widths=[30] + [10] * 16)
    drivers = [("CAPEX", 0), ("YIELD_CAL", 1), ("OPEX", 2), ("DEGR", 3), ("BASE_RATE", 5), ("MARGIN", 5),
               ("EQ_HURDLE", 6), ("TAX_KSA", 8), ("PPA_YEARS", 9)]
    base_n = list(BASE)
    base_n[1] = "=E_NAJRAN"
    srows = [("Base", base_n)]
    for d, idx in drivers:
        for side in ("LO", "HI"):
            v = list(base_n)
            if d == "YIELD_CAL":
                v[1] = f"=E_NAJRAN/YIELD_CAL*YIELD_CAL_RANGE_{side}"
            elif d == "BASE_RATE":
                v[5] = f"=BASE_RATE_{side}+MARGIN"
            elif d == "MARGIN":
                v[5] = f"=BASE_RATE+MARGIN_{side}"
            else:
                v[idx] = f"={d}_{side}"
            srows.append((f"{d} {side.lower()}", v))
    sens = scenario_table(ws, 4, srows)
    bk.name("SENS_BASE", sens["Base"])

    r = 4 + len(srows) + 3
    put(ws, f"A{r}", "Tornado summary (USc/kWh)", bold=True)
    header(ws, r + 1, ["Driver", "Low input", "High input", "Swing"])
    for i, (d, _) in enumerate(drivers):
        rr = r + 2 + i
        put(ws, f"A{rr}", d)
        put(ws, f"B{rr}", f"={sens[d + ' lo']}", F_LINK, CENTS)
        put(ws, f"C{rr}", f"={sens[d + ' hi']}", F_LINK, CENTS)
        put(ws, f"D{rr}", f"=ABS(C{rr}-B{rr})", fmt=CENTS)
    bk.name("TORNADO", f"Sensitivity!$A${r + 2}:$D${r + 1 + len(drivers)}")

    r = r + 2 + len(drivers) + 2
    put(ws, f"A{r}", "Required tariff (USc/kWh) at Najran's yield: capex (rows) x after-tax WACC (columns)", bold=True)
    hr = r + 1
    waccs = [0.03 + 0.005 * k for k in range(13)]
    capexes = [350 + 25 * k for k in range(13)]
    for k, w in enumerate(waccs):
        put(ws, f"{col(2 + k)}{hr}", w, F_INPUT, PCT)
    for j, cx in enumerate(capexes):
        rr = hr + 1 + j
        put(ws, f"A{rr}", cx, F_INPUT, NUM0)
        for k in range(len(waccs)):
            c = col(2 + k)
            w = f"{c}${hr}"
            pe, po = pvf(w, "PPA_YEARS", "-DEGR"), pvf(w, "PPA_YEARS", "OPEX_ESC")
            sh = f"(TAX_KSA*{annuity(w, 'DEP_YEARS')}/DEP_YEARS)"
            put(ws, f"{c}{rr}", f"=100*($A{rr}*(1-{sh})+(1-TAX_KSA)*OPEX*{po})/((1-TAX_KSA)*E_NAJRAN*{pe})", fmt="0.00")
    bk.name("HEATMAP", f"Sensitivity!$A${hr}:${col(1 + len(waccs))}${hr + len(capexes)}")

    # Bridge -----------------------------------------------------------
    ws = bk.sheet("Bridge", "From IRENA's global average to Najran: Shapley attribution (USD/MWh, real, per kWdc)",
                  "Both ends on IRENA's basis: per kWdc, real terms, no tax, no degradation. Each driver's bar is its average marginal effect over every order of change.",
                  widths=[34, 14, 14, 10, 10, 10, 10, 12, 10, 12, 12, 12, 12])
    header(ws, 4, ["Endpoint input", "A: IRENA global 2025", "B: Najran (implied)"])
    br = {
        "capex": (5, "Capex (USD/kWdc)", "=IRENA_TIC", "=CAPEX/DCAC", NUM0),
        "energy": (6, "Energy (kWh/kWdc-yr)", "=IRENA_CF*8760", "=E_NAJRAN/DCAC", NUM0),
        "opex": (7, "O&M (USD/kWdc-yr)", "=IRENA_OM", "=OPEX/DCAC", NUM1),
        "life": (8, "Life (yrs)", "=IRENA_LIFE", "=PPA_YEARS", NUM0),
    }
    for rr, label, a, b, fmt in br.values():
        put(ws, f"A{rr}", label)
        put(ws, f"B{rr}", a, F_LINK, fmt)
        put(ws, f"C{rr}", b, F_LINK, fmt)
    put(ws, "A9", "Real-equivalent tariff (USD/MWh, no degradation)")
    put(ws, "B9", "=IRENA_LCOE", F_LINK, NUM1)
    put(ws, "C9", f"=T_NAJRAN*10*{annuity('WACC_KSA', 'PPA_YEARS')}/{annuity('REAL_KSA', 'PPA_YEARS')}", fmt=NUM1)
    put(ws, "A10", "Implied real return (solves to the tariff)")
    for c in "BC":
        put(ws, f"{c}10", f"=RATE({c}8,{c}9/1000*{c}6-{c}7,-{c}5)", fmt=PCT2, fill=FILL_OUT)
    bk.name("BR_RATE_A", "Bridge!$B$10")
    bk.name("BR_RATE_B", "Bridge!$C$10")

    put(ws, "A12", "Shapley table: 1 = driver at Najran value", bold=True)
    header(ws, 13, ["Mask", "Capex", "Energy", "O&M", "Return+life", "Drivers moved", "Weight",
                    "LCOE USD/MWh", "Capex effect", "Energy effect", "O&M effect", "Return effect"])
    fac = [("capex", 5), ("energy", 6), ("opex", 7)]
    for m in range(16):
        rr = 14 + m
        put(ws, f"A{rr}", m, fmt="0")
        for b in range(4):
            put(ws, f"{col(2 + b)}{rr}", (m >> b) & 1, F_INPUT, "0")
        put(ws, f"F{rr}", f"=SUM(B{rr}:E{rr})", fmt="0")
        put(ws, f"G{rr}", f"=IF(F{rr}<4,FACT(F{rr})*FACT(3-F{rr})/FACT(4),0)", fmt="0.000")
        cap, en, om = (f"IF({col(2 + b)}{rr}=1,$C${row},$B${row})" for b, (_, row) in enumerate(fac))
        rate = f"IF(E{rr}=1,$C$10,$B$10)"
        life = f"IF(E{rr}=1,$C$8,$B$8)"
        put(ws, f"H{rr}", f"=1000*({cap}*{rate}/(1-(1+{rate})^-{life})+{om})/{en}", fmt=NUM1)
        for b in range(4):
            c = col(9 + b)
            flag = col(2 + b)
            put(ws, f"{c}{rr}", f"=IF({flag}{rr}=0,G{rr}*(INDEX($H$14:$H$29,A{rr}+{2 ** b}+1)-H{rr}),0)", fmt=NUM1)
    put(ws, "A31", "Waterfall (USD/MWh)", bold=True)
    header(ws, 32, ["Step", "Value"])
    steps = [("IRENA global average 2025", "=H14"), ("Capex", "=SUM(I14:I29)"), ("Energy yield", "=SUM(J14:J29)"),
             ("O&M", "=SUM(K14:K29)"), ("Return accepted (and life)", "=SUM(L14:L29)"),
             ("Najran, real equivalent", "=H29"), ("Flat nominal vs real basis", "=T_NAJRAN*10-H29"),
             ("Najran awarded tariff", "=T_NAJRAN*10")]
    for i, (label, f) in enumerate(steps):
        put(ws, f"A{33 + i}", label)
        put(ws, f"B{33 + i}", f, fmt=NUM1, fill=FILL_OUT if i in (0, 5, 7) else None)
    bk.name("WATERFALL", "Bridge!$A$33:$B$40")

    # Oman -------------------------------------------------------------
    ws = bk.sheet("Oman", "Replicability: could Nama PWP (Oman) reach a Saudi-style tariff?",
                  "Same closed-form engine. Oman's financing = KSA terms plus the sovereign spread and country risk premium gap (Damodaran, Jan 2026).",
                  widths=[34] + [11] * 16)
    ksa = list(BASE)
    ksa[1] = "=E_NAJRAN"
    omn = ["=OMN_CAPEX", "=INDEX(Yield!$H:$H,MATCH(\"IBRI3\",Yield!$A:$A,0))", "=OPEX", "=DEGR", "=OPEX_ESC",
           "=KD_OMN", "=KE_OMN", "=GEAR_WACC", "=TAX_OMN", "=OMN_PPA", "=OMN_PPA"]
    markets = scenario_table(ws, 4, [("KSA (Najran site, hurdle returns)", ksa), ("Oman (Ibri site, hurdle returns)", omn)])
    bk.name("OMN_COMMERCIAL", markets["Oman (Ibri site, hurdle returns)"])
    bk.name("KSA_COMMERCIAL", markets["KSA (Najran site, hurdle returns)"])

    groups = {"Resource": [1], "Capex and scale": [0], "Financing and tax": [5, 6, 8],
              "Contract tenor": [9, 10]}
    names = list(groups)
    shap_rows = []
    for m in range(16):
        v = [ksa[k] for k in range(11)]
        for b, g in enumerate(names):
            if (m >> b) & 1:
                for k in groups[g]:
                    v[k] = omn[k]
        shap_rows.append((f"mask {m}", v))
    put(ws, "A9", "Shapley: KSA -> Oman (each row switches the flagged driver groups to Oman)", bold=True)
    shap = scenario_table(ws, 10, shap_rows)
    header(ws, 10, ["Resource", "Capex and scale", "Financing and tax", "Contract tenor", "Moved", "Weight",
                    "Resource effect", "Capex effect", "Financing effect", "Tenor effect"], start=19)
    put(ws, "R10", "Mask", F_HEAD, fill=FILL_HEAD)
    for m in range(16):
        rr = 11 + m
        put(ws, f"R{rr}", m, F_INPUT, "0")
        for b in range(4):
            put(ws, f"{col(19 + b)}{rr}", (m >> b) & 1, F_INPUT, "0")
        put(ws, f"W{rr}", f"=SUM(S{rr}:V{rr})", fmt="0")
        put(ws, f"X{rr}", f"=IF(W{rr}<4,FACT(W{rr})*FACT(3-W{rr})/FACT(4),0)", fmt="0.000")
        for b in range(4):
            put(ws, f"{col(25 + b)}{rr}",
                f"=IF({col(19 + b)}{rr}=0,X{rr}*(INDEX($Q$11:$Q$26,$R{rr}+{2 ** b}+1)-Q{rr}),0)", fmt="0.000")
    put(ws, "A29", "KSA -> Oman bridge (USc/kWh)", bold=True)
    header(ws, 30, ["Step", "Value", "Portable?"])
    obr = [("KSA at hurdle returns", f"={shap['mask 0']}", ""),
           ("Resource", "=SUM(Y11:Y26)", "Inherited"),
           ("Capex and scale", "=SUM(Z11:Z26)", "Designable: tender size and site preparation"),
           ("Financing and tax", "=SUM(AA11:AA26)", "Partly designable: offtaker credit support"),
           ("Contract tenor", "=SUM(AB11:AB26)", "Designable: PPA length"),
           ("Oman at hurdle returns", f"={shap['mask 15']}", "")]
    for i, (label, f, note) in enumerate(obr):
        put(ws, f"A{31 + i}", label)
        put(ws, f"B{31 + i}", f, fmt=CENTS, fill=FILL_OUT if i in (0, 5) else None)
        put(ws, f"C{31 + i}", note, F_NOTE)
    bk.name("OMN_BRIDGE", "Oman!$A$31:$C$36")

    put(ws, "A39", "Oman tariff cases (USc/kWh)", bold=True)
    credit = list(omn)
    credit[5], credit[6] = "=KD_KSA", "=EQ_HURDLE"
    scaled = list(credit)
    scaled[0], scaled[9], scaled[10] = "=CAPEX", "=PPA_YEARS", "=PPA_YEARS"
    record = list(scaled)
    record[6], record[7] = "=IRR_NAJRAN", "=0"
    cases = scenario_table(ws, 40, [
        ("1. As today: 500 MW, 20-yr PPA, Oman risk", omn),
        ("2. Plus Saudi-grade financing (credit support)", credit),
        ("3. Plus GW-scale tender and 25-yr PPA", scaled),
        ("4. Plus a Najran-style sponsor return", record),
    ])
    for i, k in enumerate(cases):
        bk.name(f"OMN_CASE{i + 1}", cases[k])

    put(ws, "A47", "Avoided cost of gas-fired power in Oman (USc/kWh)", bold=True)
    put(ws, "A48", "Fuel at the regulated power-sector gas price")
    put(ws, "B48", "=OMN_GAS*OMN_HR/10", fmt=CENTS, fill=FILL_OUT)
    put(ws, "A49", "Fuel at LNG export netback (opportunity cost)")
    put(ws, "B49", "=(BRENT*LNG_SLOPE-LNG_NETBACK_COST)*OMN_HR/10", fmt=CENTS, fill=FILL_OUT)
    bk.name("OMN_AVOIDED_GAS", "Oman!$B$48")
    bk.name("OMN_AVOIDED_LNG", "Oman!$B$49")

    # Checks -----------------------------------------------------------
    ws = bk.sheet("Checks", "Model checks", "All must be TRUE. The Cover shows the combined result.",
                  widths=[60, 14, 10] + [7] * 36)
    put(ws, "A5", "Parity: cash flows at Najran's implied capex return exactly the hurdle WACC")
    put(ws, "D4", 0, fmt="0")
    for k in range(1, YEARS + 1):
        put(ws, f"{col(4 + k)}4", k, fmt="0")
    put(ws, "D5", "=-IMPL_NAJRAN", fmt=NUM1)
    for k in range(1, YEARS + 1):
        c = col(4 + k)
        put(ws, f"{c}5", f"=IF({c}$4<=PPA_YEARS,(T_NAJRAN/100*E_NAJRAN*(1-DEGR)^({c}$4-1)-OPEX*(1+OPEX_ESC)^({c}$4-1))*(1-TAX_KSA)"
                         f"+TAX_KSA*IF({c}$4<=DEP_YEARS,IMPL_NAJRAN/DEP_YEARS,0),0)", fmt=NUM1)
    put(ws, "B5", f"=IRR(D5:{col(4 + YEARS)}5)", fmt=PCT2)
    checks = [
        (5, "=ABS(B5-WACC_KSA)<0.000001"),
        (6, "=ABS(SENS_BASE-INDEX(Engine!$U:$U,MATCH(\"NAJRAN\",Engine!$A:$A,0)))<0.000001"),
        (7, "=ABS(Bridge!H14-IRENA_LCOE)<0.000001"),
        (8, "=ABS(Bridge!H29-Bridge!C9)<0.000001"),
        (9, "=ABS(SUM(Bridge!I14:L29)-(Bridge!H29-Bridge!H14))<0.000001"),
        (10, "=ABS(SUM(Oman!B32:B35)-(Oman!B36-Oman!B31))<0.000001"),
        (11, "=ABS(Oman!B31-KSA_COMMERCIAL)<0.000001"),
        (12, "=ABS(PF_END_BAL)<0.001"),
        (13, "=PF_MIN_DSCR>=DSCR-0.000001"),
        (14, "=PF_GEARING<=GEAR_MAX+0.000001"),
        (15, f'=COUNTIFS(Tracker!$U$5:$U${last_tr},"Y",Tracker!$X$5:$X${last_tr},"")=0'),
        (16, f"=SUMPRODUCT(--ISERROR(Engine!A{e_top}:AB{e_bot}))+SUMPRODUCT(--ISERROR(Engine_CF!A{e_top}:AP{e_bot}))=0"),
    ]
    labels = {6: "Sensitivity base equals Engine required tariff for Najran",
              7: "Bridge start reproduces IRENA's USD 44/MWh", 8: "Bridge end reproduces Najran's real-equivalent tariff",
              9: "Shapley effects sum to the IRENA-to-Najran gap", 10: "Oman bridge effects sum to the KSA-to-Oman gap",
              11: "Oman bridge starts at the KSA hurdle tariff", 12: "PF debt fully repaid by end of tenor",
              13: "PF minimum DSCR respects the covenant", 14: "PF gearing within the cap",
              15: "Every analysed tracker row has a tariff", 16: "No errors on Engine and Engine_CF"}
    for rr, f in checks:
        if rr != 5:
            put(ws, f"A{rr}", labels[rr])
        put(ws, f"C{rr}", f, fill=FILL_OUT)
    put(ws, "A18", "All checks pass", bold=True)
    put(ws, "C18", "=AND(C5:C16)", fill=FILL_OUT, bold=True)
    bk.name("ALL_CHECKS", "Checks!$C$18")

    # Sources and Log --------------------------------------------------
    ws = bk.sheet("Sources", "Source register", "Accessed dates are when the figure was checked.",
                  widths=[6, 26, 60, 11, 70, 11, 70])
    header(ws, 4, ["ID", "Publisher", "Title", "Date", "URL", "Accessed", "Used for"])
    for i, s in enumerate(sources):
        for j, k in enumerate(["id", "publisher", "title", "date", "url", "accessed", "used_for"], start=1):
            ws.cell(row=5 + i, column=j, value=s[k]).font = F_BASE
    ws = bk.sheet("Log", "Method notes and version log", widths=[110])
    notes = [
        f"Built {date.today().isoformat()} from data/*.csv by src/nrep/build_model.py.",
        "Tariffs are flat nominal USD for the PPA term; the engine treats them as energy-only revenue with no merchant tail.",
        "Yield: PVGIS tracked yield x a single calibration factor from Sakaka's measured 2022 output. One factor for all sites.",
        "Implied capex uses the KSA hurdle WACC for every row, including UAE and Qatar comparators.",
        "Project IRR uses the disclosed cost of that project where published, else the Saudi average for its vintage; rows outside 2021-2025 show blank.",
        "Debt is sculpted on CFADS after unlevered tax to avoid a circular reference; the interest tax shield goes to equity.",
        "No loss carry-forward in the PF model; taxable profit is positive in all base runs.",
        "Bridge uses IRENA's basis (per kWdc, real, pre-tax, no degradation) at both ends; DC/AC ratio converts Saudi inputs.",
        "Oman financing adds the Baa3-Aa3 default spread gap to debt and the country risk premium gap to equity.",
    ]
    for i, n in enumerate(notes):
        put(ws, f"A{3 + i}", n)

    # Cover ------------------------------------------------------------
    ws = cover
    put(ws, "B3", "Case question", bold=True)
    put(ws, "D3", "A Gulf utility (Oman's Nama PWP) wants to know whether Saudi Arabia's record solar tariffs are "
                  "replicable in its market. What drives a ~1.1 USc/kWh bid, and how sensitive is it to capital cost, "
                  "capacity factor and financing?")
    ws["D3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[3].height = 54
    put(ws, "B5", "Headline outputs", bold=True)
    key = [
        ("Najran awarded tariff (USc/kWh)", "=T_NAJRAN", CENTS),
        ("Najran real-equivalent tariff (USc/kWh)", "=REAL_NAJRAN", CENTS),
        ("Hurdle WACC, KSA (nominal, after tax)", "=WACC_KSA", PCT2),
        ("Capex Najran's tariff supports at the hurdle (USD/kWac)", "=IMPL_NAJRAN", USD0),
        ("Disclosed Saudi capex, 2025 PPAs (USD/kWac)", "=CX_DISCLOSED_2025", USD0),
        ("Najran project IRR at the 2025 disclosed cost", "=IRR_NAJRAN", PCT2),
        ("Equity IRR, project selected on PF_Model", "=PF_EQ_IRR", PCT2),
        ("IRENA global average LCOE 2025 (USD/MWh)", "=IRENA_LCOE", NUM1),
        ("Oman tariff at Oman hurdle returns (USc/kWh)", "=OMN_COMMERCIAL", CENTS),
        ("Oman avoided fuel cost, regulated gas (USc/kWh)", "=OMN_AVOIDED_GAS", CENTS),
        ("Oman avoided fuel cost, LNG netback (USc/kWh)", "=OMN_AVOIDED_LNG", CENTS),
        ("All model checks pass", "=ALL_CHECKS", None),
    ]
    for i, (label, f, fmt) in enumerate(key):
        put(ws, f"B{6 + i}", label)
        put(ws, f"C{6 + i}", f, F_LINK, fmt, FILL_OUT)
    r = 6 + len(key) + 1
    put(ws, f"B{r}", "Colour code", bold=True)
    put(ws, f"B{r + 1}", "Blue text: input or reported data", F_INPUT)
    put(ws, f"B{r + 2}", "Green text: link from another tab", F_LINK)
    put(ws, f"B{r + 3}", "Black text: calculation")
    put(ws, f"B{r + 4}", "Yellow fill: assumption without a direct source", fill=FILL_KEY)
    put(ws, f"B{r + 5}", "Green fill: key output", fill=FILL_OUT)
    r += 7
    put(ws, f"B{r}", "Tabs", bold=True)
    guide = [("Inputs", "Every assumption with low/high range, type and source"),
             ("Tracker", "Awarded tariffs, Saudi and GCC"), ("Yield", "PVGIS site yields and Sakaka calibration"),
             ("Capex", "Disclosed costs, bottom-up build, football field"),
             ("Engine", "What each tariff implies: capex at hurdle, IRR at disclosed cost, real equivalent"),
             ("Engine_CF", "Cash flows behind the Engine IRRs"), ("PF_Model", "Levered single-asset model"),
             ("Sensitivity", "Tornado and capex x WACC grid for Najran"),
             ("Bridge", "IRENA global average to Najran, Shapley attribution"),
             ("Oman", "Replicability bridge, Oman tariff cases, avoided cost"), ("Checks", "Integrity checks"),
             ("Sources", "Source register"), ("Log", "Method notes")]
    for i, (tab, desc) in enumerate(guide):
        c = ws.cell(row=r + 1 + i, column=2, value=tab)
        c.hyperlink = f"#'{tab}'!A1"
        c.font = Font(name=ARIAL, size=10, color="1F4E79", underline="single")
        put(ws, f"D{r + 1 + i}", desc)

    for s in wb.worksheets:
        s.sheet_properties.tabColor = "1F2A36" if s.title in ("Cover", "Checks") else None
    wb.calculation.fullCalcOnLoad = True
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    print(build())

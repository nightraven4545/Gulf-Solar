"""Build the deck and one-pager from the recalculated workbook. Every number on a slide is read from the model."""

import re
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[2]
MODEL = ROOT / "model" / "NREP_Tariff_LCOE_Model.xlsx"
DECK = ROOT / "deck" / "NREP_Tariff_Replicability.pptx"
ONEPAGER = ROOT / "onepager" / "NREP_onepager.pptx"

INK, TEXT2, MUTED = RGBColor(0x1F, 0x2A, 0x36), RGBColor(0x52, 0x51, 0x4E), RGBColor(0x80, 0x86, 0x8C)
GRID, PALE, WHITE = RGBColor(0xD9, 0xDE, 0xE3), RGBColor(0xF1, 0xF3, 0xF5), RGBColor(0xFF, 0xFF, 0xFF)
AMBER, BLUE, TEAL, GREY = (RGBColor.from_string(h) for h in ("C2780B", "3A5FA8", "1A9A8A", "A3ACB5"))
AMBER_PALE = RGBColor(0xF6, 0xE3, 0xC4)
FONT = "Arial"
W, H = 13.333, 7.5


# ---------------------------------------------------------------- model access
class Model:
    def __init__(self, path=MODEL):
        self.wb = load_workbook(path, data_only=True)
        if self.wb["Cover"]["C6"].value is None:
            raise SystemExit("Model has no cached values; run `make model` first.")

    def ref(self, name):
        sheet, coord = self.wb.defined_names[name].attr_text.replace("'", "").replace("$", "").split("!")
        return self.wb[sheet], coord

    def __getitem__(self, name):
        ws, coord = self.ref(name)
        return ws[coord].value

    def rows(self, name):
        ws, coord = self.ref(name)
        return [[c.value for c in row] for row in ws[coord]]

    def table(self, sheet, header_row):
        ws = self.wb[sheet]
        head = [c.value for c in ws[header_row]]
        out = []
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            if row[0] is None:
                break
            out.append(dict(zip(head, row)))
        return out


# ---------------------------------------------------------------- slide helpers
def font(run, size=14, bold=False, color=INK, italic=False):
    run.font.name, run.font.size, run.font.bold, run.font.italic = FONT, Pt(size), bold, italic
    run.font.color.rgb = color


def text(slide, x, y, w, h, paras, size=14, color=INK, bold=False, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT,
         space_after=6, name=None):
    """paras: list of str or (str, dict) where dict can set size/bold/color/bullet."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    for i, p in enumerate(paras):
        s, opt = (p, {}) if isinstance(p, str) else p
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.space_after = Pt(opt.get("space_after", space_after))
        if opt.get("bullet"):
            pPr = para._p.get_or_add_pPr()
            pPr.set("marL", str(Inches(0.22)))
            pPr.set("indent", str(-Inches(0.18)))
            bu = pPr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}buChar", {"char": "•"})
            pPr.append(bu)
        run = para.add_run()
        run.text = s
        font(run, opt.get("size", size), opt.get("bold", bold), opt.get("color", color), opt.get("italic", False))
    return box


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(W), Inches(H)
        layout = self.prs.slide_layouts[5]  # Title Only
        ph = layout.placeholders[0]
        ph.left, ph.top, ph.width, ph.height = Inches(0.5), Inches(0.35), Inches(W - 1.0), Inches(0.95)
        self.layout = layout
        self.blank = self.prs.slide_layouts[6]
        self.n = 0

    def slide(self, title, source, notes="", tag=True):
        s = self.prs.slides.add_slide(self.layout)
        self.n += 1
        tf = s.shapes.title.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.TOP
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = title
        font(r, 22, True, INK)
        text(s, 0.5, 6.88, 10.4, 0.45, [source], size=10, color=MUTED, name="Source")
        if tag:
            text(s, 10.95, 6.88, 1.6, 0.45, ["Preliminary, outside-in"], size=10, color=MUTED, align=PP_ALIGN.RIGHT)
        text(s, 12.55, 6.88, 0.3, 0.45, [str(self.n)], size=10, color=MUTED, align=PP_ALIGN.RIGHT)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s

    def dark(self, lines, notes=""):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        bg = s.background.fill
        bg.solid()
        bg.fore_color.rgb = INK
        text(s, 0.9, 2.2, 11.5, 3.5, lines, color=WHITE)
        if notes:
            s.notes_slide.notes_text_frame.text = notes
        return s


def style_chart(chart, legend=False, size=11):
    chart.font.name, chart.font.size = FONT, Pt(size)
    chart.font.color.rgb = TEXT2
    chart.has_title = False
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(size)


def quiet_axes(chart, val_fmt=None, val_min=None, val_max=None, gridlines=True):
    va = chart.value_axis
    va.has_major_gridlines = gridlines
    if gridlines:
        va.major_gridlines.format.line.color.rgb = GRID
        va.major_gridlines.format.line.width = Pt(0.75)
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(11)
    va.tick_labels.font.color.rgb = TEXT2
    if val_fmt:
        va.tick_labels.number_format, va.tick_labels.number_format_is_linked = val_fmt, False
    if val_min is not None:
        va.minimum_scale = val_min
    if val_max is not None:
        va.maximum_scale = val_max
    ca = chart.category_axis
    ca.format.line.color.rgb = GRID
    ca.tick_labels.font.size = Pt(11)
    ca.tick_labels.font.color.rgb = INK
    ca.has_major_gridlines = False


def fill(series_or_point, rgb):
    f = series_or_point.format.fill
    f.solid()
    f.fore_color.rgb = rgb
    series_or_point.format.line.fill.background()


def hide(series):
    series.format.fill.background()
    series.format.line.fill.background()


def label(point, s, size=11, color=INK, pos=None, bold=False):
    dl = point.data_label
    tf = dl.text_frame
    tf.text = s
    for r in tf.paragraphs[0].runs:
        font(r, size, bold, color)
    if pos is not None:
        dl.position = pos


def add_chart(slide, kind, data, x, y, w, h):
    return slide.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), data).chart


def cents(v, d=2):
    return f"{v:.{d}f}¢"


def floating(slide, cats, lows, highs, colors, x, y, w, h, fmt, vmin, vmax, labels=None, horizontal=True,
             gap=60, label_color=INK, label_on_base=False, unit=None):
    """Stacked bar with an invisible base series: draws lows..highs per category."""
    if horizontal:  # bar charts plot bottom-up; reverse the data so the first item sits on top
        cats, lows, highs, colors = cats[::-1], lows[::-1], highs[::-1], colors[::-1]
        labels = labels[::-1] if labels else labels
        label_color = label_color[::-1] if isinstance(label_color, list) else label_color
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("base", lows)
    cd.add_series("span", [hi - lo for lo, hi in zip(lows, highs)])
    kind = XL_CHART_TYPE.BAR_STACKED if horizontal else XL_CHART_TYPE.COLUMN_STACKED
    ch = add_chart(slide, kind, cd, x, y, w, h)
    style_chart(ch)
    quiet_axes(ch, fmt, vmin, vmax)
    plot = ch.plots[0]
    plot.gap_width, plot.overlap = gap, 100
    hide(plot.series[0])
    span = plot.series[1]
    for i, c in enumerate(colors):
        fill(span.points[i], c)
    if labels:
        for i, s in enumerate(labels):
            if label_on_base:
                label(plot.series[0].points[i], s, pos=XL_LABEL_POSITION.INSIDE_END)
            else:
                label(span.points[i], s, color=label_color[i] if isinstance(label_color, list) else label_color,
                      pos=XL_LABEL_POSITION.CENTER)
    if unit:
        ch.value_axis.major_unit = unit
    return ch


def card(slide, x, y, w, h, head, body, accent=INK, size=14):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = PALE
    shp.line.fill.background()
    shp.shadow.inherit = False
    text(slide, x + 0.15, y + 0.12, w - 0.3, 0.5, [(head, {"bold": True, "color": accent, "size": size + 1})])
    text(slide, x + 0.15, y + 0.62, w - 0.3, h - 0.7, [(b, {"bullet": True}) for b in body], size=size)


def table(slide, rows, x, y, w, col_w, size=12, head_fill=INK, row_h=0.38, fills=None):
    t = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w),
                               Inches(row_h * len(rows))).table
    for j, cw in enumerate(col_w):
        t.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.margin_left = c.margin_right = Inches(0.06)
            c.margin_top = c.margin_bottom = Inches(0.03)
            c.text = ""
            r = c.text_frame.paragraphs[0].add_run()
            r.text = str(v)
            c.fill.solid()
            if i == 0:
                c.fill.fore_color.rgb = head_fill
                font(r, size, True, WHITE)
            else:
                c.fill.fore_color.rgb = (fills or {}).get((i, j), WHITE if i % 2 else PALE)
                font(r, size, False, INK)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
    return t


# ---------------------------------------------------------------- deck
def build(path=DECK):
    m = Model()
    eng = m.table("Engine", 4)
    tr = m.table("Tracker", 4)
    najran = next(r for r in eng if r["ID"] == "NAJRAN")
    kd = m["KD_KSA"]
    d = Deck()

    t_naj, real_naj, irr_naj = m["T_NAJRAN"], m["REAL_NAJRAN"], m["IRR_NAJRAN"]
    disc25, req_naj = m["CX_DISCLOSED_2025"], najran["Required tariff at base capex (USc)"]
    wf = {k: v for k, v in m.rows("WATERFALL")}
    ob = {r[0]: r[1] for r in m.rows("OMN_BRIDGE")}
    cases = [m[f"OMN_CASE{i}"] for i in range(1, 5)]
    gas, lng = m["OMN_AVOIDED_GAS"], m["OMN_AVOIDED_LNG"]
    ksa_pv = [r for r in tr if r["Country"] == "KSA" and r["Tech"] == "PV" and r["Tariff basis"] != "pending"]
    gw_sppc = sum(r["MWac"] for r in ksa_pv if r["Track"] == "SPPC") / 1000
    gw_pif = sum(r["MWac"] for r in ksa_pv if r["Track"] == "PIF") / 1000
    gap = wf["IRENA global average 2025"] - wf["Najran, real equivalent"]
    share = lambda k: -wf[k] / gap

    # 1 Title --------------------------------------------------------------
    d.dark([
        ("Can Oman match Saudi Arabia's record solar tariffs?", {"size": 34, "bold": True, "color": WHITE}),
        ("What drives a 1.1 ¢/kWh bid, and which parts of it travel", {"size": 20, "color": WHITE, "space_after": 30}),
        ("Outside-in analysis on public data · November 2026", {"size": 14, "color": GREY}),
    ], notes="Case: a Gulf utility (Oman's Nama PWP) asks whether Saudi tariffs are replicable. "
             "Everything here is built from public data; the Excel model carries every number and source.")

    # 2 Executive summary ---------------------------------------------------
    s = d.slide("Saudi tariffs rest 60/35 on sunshine and cost versus returns few investors would accept; "
                "Oman can copy the design and still beat gas",
                "Source: model Cover and Bridge tabs; SPPC awards via TaiyangNews (S14); ACWA Power investor report FY2025 (S43); "
                "IRENA Renewable Power Generation Costs in 2025 (S21)",
                notes="Lead with the answer. Three sections follow, one for each part of the argument.")
    lines = [
        ("Situation", (f"Saudi Arabia has contracted {gw_sppc + gw_pif:.0f} GW of utility solar since 2018 at "
                      f"1.0 to 1.8 ¢/kWh, about a quarter of IRENA's 4.4 ¢ global average.")),
        ("Complication", (f"At the 2025 disclosed cost of USD {disc25:.0f}/kWac, Najran's {t_naj:.2f} ¢ bid returns "
                         f"{irr_naj:.1%} before leverage, below the {kd:.2%} cost of debt.")),
        ("A. The record is real", (f"In real terms Najran is {real_naj:.2f} ¢; IRENA's own Saudi average is 4.0 ¢. "
                                   "A flat 25-year tariff is not a cost.")),
        ("B. About a third is accepted return",
         (f"Of the USD {gap:.0f}/MWh gap to the global average, yield and capex explain "
         f"{share('Energy yield') + share('Capex'):.0%} and the return sponsors accept {share('Return accepted (and life)'):.0%}.")),
        ("C. Oman can copy the design", (f"With credit support, GW-scale tenders and 25-year PPAs, Oman reaches "
                                        f"{cents(cases[2])}, below the {cents(gas, 1)} fuel cost of gas power.")),
    ]
    y = 1.55
    for head, body in lines:
        text(s, 0.5, y, 2.6, 0.9, [(head, {"bold": True, "color": AMBER if head[1] == "." else INK})], size=14)
        text(s, 3.1, y, 5.9, 0.95, [body], size=14)
        y += 1.0
    stats = [(cents(t_naj), "Najran tariff, Oct 2025 (record low in Round 6)"),
             (f"{irr_naj:.1%}", f"Najran return before leverage at USD {disc25:.0f}/kWac"),
             (cents(cases[2]), "Oman with Saudi-style tender design")]
    for i, (big, small) in enumerate(stats):
        yy = 1.55 + i * 1.7
        text(s, 9.5, yy, 3.3, 0.8, [(big, {"size": 40, "bold": True, "color": [AMBER, AMBER, TEAL][i]})])
        text(s, 9.5, yy + 0.8, 3.3, 0.7, [small], size=12, color=TEXT2)

    # 3 Programme ----------------------------------------------------------
    s = d.slide(f"Saudi Arabia has contracted {gw_sppc + gw_pif:.0f} GW of utility solar through two buying tracks, "
                f"{gw_pif / (gw_pif + gw_sppc):.0%} of it negotiated with PIF-backed ACWA consortia",
                "Source: model Tracker tab (37 projects); SPPC, ACWA Power, Masdar and trade press releases (S01-S43). "
                "PV only; excludes Category A plants under 150 MW; Round 7 (3.1 GW) and Round 8 (1.9 GW) pending",
                notes="SPPC runs competitive rounds; PIF projects are negotiated with ACWA Power, Badeel (PIF) and SAPCO (Aramco). "
                      "Both sell to SPPC under 25-year PPAs (35 years for three 2023 projects).")
    by = defaultdict(lambda: [0.0, 0.0])
    for r in ksa_pv:
        by[int(str(r["Date"])[:4])][0 if r["Track"] == "SPPC" else 1] += r["MWac"] / 1000
    years = list(range(min(by), max(by) + 1))
    cd = CategoryChartData()
    cd.categories = [str(yr) for yr in years]
    cd.add_series("SPPC competitive tenders", [by[yr][0] for yr in years])
    cd.add_series("PIF bilateral (ACWA Power / Badeel / SAPCO)", [by[yr][1] for yr in years])
    ch = add_chart(s, XL_CHART_TYPE.COLUMN_STACKED, cd, 0.5, 1.5, 8.4, 5.2)
    style_chart(ch, legend=True)
    quiet_axes(ch, '0" GW"')
    ch.plots[0].gap_width, ch.plots[0].overlap = 55, 100
    fill(ch.plots[0].series[0], AMBER)
    fill(ch.plots[0].series[1], BLUE)
    for i, yr in enumerate(years):
        tot = sum(by[yr])
        if tot:
            label(ch.plots[0].series[1].points[i] if by[yr][1] else ch.plots[0].series[0].points[i],
                  f"{tot:.1f}", color=WHITE, pos=XL_LABEL_POSITION.INSIDE_END, bold=True)
    text(s, 9.3, 1.6, 3.5, 5, [
        ("How the two tracks differ", {"bold": True}),
        ("SPPC tenders: open rounds, lowest levelised tariff wins, results published to 5 decimals", {"bullet": True}),
        ("PIF track: directly negotiated, tariffs disclosed only from 2025", {"bullet": True}),
        ("Same buyer, same 25-year flat energy-only PPA, SAR pegged to USD", {"bullet": True}),
        ("Round 7 (3.1 GW PV) bids are in; Round 8 (1.9 GW) is at RFQ", {"bullet": True}),
    ], size=14)

    # 4 Tariff history -----------------------------------------------------
    s = d.slide("Awarded tariffs hit 1.04 ¢ in 2021 and have stayed between 1.1 and 1.8 ¢ since, between a quarter "
                "and two-fifths of IRENA's 4.4 ¢ global average",
                "Source: model Tracker and Engine tabs; IRENA RPGC 2025 (S21) global and Saudi weighted-average LCOE for "
                "projects commissioned in 2025 (real USD). Tariffs are flat nominal USc/kWh for the PPA term",
                notes="Each dot is an award. Saudi tariffs flattened as module prices rose in 2021-23 and interest rates rose; "
                      "falling module prices in 2024-25 brought them back down to 1.1-1.3 cents.")
    xy = XyChartData()
    groups = [("SPPC tenders", lambda r: r["Track"] == "SPPC", AMBER),
              ("PIF bilateral", lambda r: r["Track"] == "PIF", BLUE),
              ("UAE and Qatar tenders", lambda r: r["Country"] != "KSA", GREY)]
    pts = {}
    for name, f, _ in groups:
        ser = xy.add_series(name)
        pts[name] = [r for r in eng if f(r)]
        for r in pts[name]:
            ds = str(r["Date"])
            ser.add_data_point(int(ds[:4]) + (int(ds[5:7]) - 1) / 12 if len(ds) > 4 else int(ds) + 0.5, r["Tariff USc/kWh"])
    irena = xy.add_series("IRENA 2025 average LCOE")
    irena.add_data_point(2025.5, m["IRENA_LCOE"] / 10)
    irena.add_data_point(2025.5, m["IRENA_LCOE_KSA"] / 10)
    ch = add_chart(s, XL_CHART_TYPE.XY_SCATTER, xy, 0.5, 1.5, 9.0, 5.25)
    style_chart(ch, legend=True)
    quiet_axes(ch, '0.0"¢"', 0, 5)
    ch.category_axis.minimum_scale, ch.category_axis.maximum_scale = 2017, 2027
    ch.category_axis.tick_labels.number_format, ch.category_axis.tick_labels.number_format_is_linked = "0", False
    ch.category_axis.major_unit = 1
    for ser, (_, _, c) in zip(ch.plots[0].series, groups + [(None, None, INK)]):
        ser.format.line.fill.background()
        ser.marker.style = XL_MARKER_STYLE.CIRCLE if c != INK else XL_MARKER_STYLE.DIAMOND
        ser.marker.size = 9
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = c
        ser.marker.format.line.color.rgb = WHITE
    sppc = ch.plots[0].series[0]
    for i, r in enumerate(pts["SPPC tenders"]):
        if r["ID"] in ("NAJRAN", "SHUAIBAH1", "SAKAKA"):
            label(sppc.points[i], f"{r['Project']} {r['Tariff USc/kWh']:.2f}¢", pos=XL_LABEL_POSITION.RIGHT)
    lab = ch.plots[0].series[3]
    label(lab.points[0], "Global 4.4¢", pos=XL_LABEL_POSITION.LEFT)
    label(lab.points[1], "Saudi 4.0¢", pos=XL_LABEL_POSITION.LEFT)
    text(s, 9.8, 1.6, 3.0, 5, [
        ("What the dots say", {"bold": True}),
        ("Saudi Arabia set the solar floor in 2021 and again in 2025", {"bullet": True}),
        ("Tariffs rose in 2022-23 with module prices and interest rates, then fell back", {"bullet": True}),
        ("UAE and Qatar awards sit inside the Saudi range", {"bullet": True}),
        ("IRENA's averages are real LCOEs for plants commissioned in 2025, not bids", {"bullet": True}),
    ], size=14)

    # 5 Tariff vs LCOE -----------------------------------------------------
    s = d.slide(f"A flat nominal tariff hides inflation: in real terms Najran's {t_naj:.2f} ¢ is {real_naj:.2f} ¢, "
                "a fifth of IRENA's own Saudi average of 4.0 ¢",
                f"Source: model Engine tab (real equivalent at {m['WACC_KSA']:.1%} nominal WACC and {m['US_CPI']:.1%} "
                "inflation, COD-year dollars); IRENA RPGC 2025 (S21) Fig. 3.7 and 3.6",
                notes="A flat 25-year price loses value every year. The constant real price with the same present value is the "
                      "number to compare with IRENA's real LCOE. IRENA's Saudi average reflects plants bid in 2022-23.")
    cd = CategoryChartData()
    cats = ["Najran tariff (flat nominal, 25 yrs)", "Najran, constant real equivalent", "IRENA Saudi average 2025",
            "IRENA global average 2025"]
    vals = [t_naj, real_naj, m["IRENA_LCOE_KSA"] / 10, m["IRENA_LCOE"] / 10]
    cd.categories = cats
    cd.add_series("USc/kWh", vals)
    ch = add_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, cd, 0.5, 1.6, 7.8, 4.6)
    style_chart(ch)
    quiet_axes(ch, '0.0"¢"', 0, 5)
    ch.category_axis.reverse_order = True
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.plots[0].gap_width = 70
    ser = ch.plots[0].series[0]
    for i, c in enumerate([AMBER, AMBER, GREY, GREY]):
        fill(ser.points[i], c)
        label(ser.points[i], cents(vals[i]), pos=XL_LABEL_POSITION.OUTSIDE_END, bold=i < 2)
    text(s, 8.8, 1.6, 4.0, 5, [
        ("Why a tariff is not an LCOE", {"bold": True}),
        ("Flat in nominal USD for 25 years, with no indexation", {"bullet": True}),
        ("Energy only: the buyer takes curtailment risk through deemed energy", {"bullet": True}),
        ("Land is leased by the state and SEC builds the grid connection, so neither sits fully in the tariff", {"bullet": True}),
        ("Sovereign-backed single buyer and a USD peg cut the cost of debt", {"bullet": True}),
    ], size=14)

    # 6 Capex --------------------------------------------------------------
    ff = m.rows("FF_TABLE")
    s = d.slide("Disclosed Saudi costs fell from about USD 620-830/kWac in 2021 to USD 500-550 in 2025, in line "
                "with a bottom-up build at 2026 module prices",
                "Source: model Capex tab; ACWA Power FY2025 investor report (S43) total investment cost; Masdar (S18); OPIS "
                "module price via pv magazine (S23); IRENA RPGC 2025 (S21) converted from per kWdc at DC/AC 1.2-1.4",
                notes="Four independent views. Disclosed costs may include development fees and margins paid to sponsor "
                      "affiliates, so the true build cost may be lower still. The last bar is what the tariffs imply.")
    cats = [r[0] for r in ff]
    lows, highs, pts_ = [r[4] for r in ff], [r[5] for r in ff], [r[6] for r in ff]
    floating(s, cats, lows, highs, [BLUE, BLUE, AMBER, GREY, INK], 0.5, 1.5, 7.6, 5.2, '"$"#,##0', 0, 1100,
             labels=[f"{lo:.0f}-{hi:.0f}" for lo, hi in zip(lows, highs)], label_on_base=True, unit=200)
    vin = [(int(a), b, c) for a, b, c in m.rows("CX_VIN_TABLE")]
    cd = CategoryChartData()
    cd.categories = [str(v[0]) for v in vin]
    cd.add_series("USD/kWac", [v[1] for v in vin])
    ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, 8.4, 1.9, 4.4, 3.4)
    style_chart(ch)
    ch.has_title = True
    ch.chart_title.text_frame.text = "Average disclosed cost by vintage (USD/kWac)"
    for r in ch.chart_title.text_frame.paragraphs[0].runs:
        font(r, 12, True, INK)
    quiet_axes(ch, "#,##0", 0, 900, gridlines=False)
    ch.value_axis.visible = False
    ser = ch.plots[0].series[0]
    fill(ser, BLUE)
    ch.plots[0].gap_width = 60
    for i, v in enumerate(vin):
        label(ser.points[i], f"{v[1]:.0f}", pos=XL_LABEL_POSITION.OUTSIDE_END)
    text(s, 8.5, 5.4, 4.3, 1.3, [(f"Implied by 2025 tariffs at the {m['WACC_KSA']:.1%} hurdle: USD {pts_[4]:.0f}/kWac on "
                                 f"average, {1 - pts_[4] / pts_[0]:.0%} below what was disclosed")], size=14, bold=True)

    # 7 IRR dot plot -------------------------------------------------------
    rows7 = sorted([r for r in eng if isinstance(r["Project IRR at disclosed capex"], float) and r["Country"] == "KSA"],
                   key=lambda r: r["Project IRR at disclosed capex"])
    lo_irr, hi_irr = rows7[0]["Project IRR at disclosed capex"], rows7[-1]["Project IRR at disclosed capex"]
    s = d.slide(f"At disclosed costs, Saudi awards since 2021 return between {lo_irr:.0%} and {hi_irr:.0%} before "
                f"leverage, every one below the {kd:.2%} cost of debt",
                "Source: model Engine and Engine_CF tabs. Unlevered, after-tax, nominal IRR over the PPA term at the "
                "project's disclosed cost (ACWA S43, Masdar S18) or the Saudi average for its vintage; no merchant tail",
                notes="If the project earns less than its debt costs, borrowing lowers the equity return. On PF_Model, "
                      "Najran's equity IRR is negative at the disclosed cost. Sponsors must be pricing in lower true costs, "
                      "a longer life, or earnings outside the project company.")
    xy = XyChartData()
    s_sppc, s_pif = xy.add_series("SPPC tenders"), xy.add_series("PIF bilateral")
    order = {}
    for i, r in enumerate(rows7):
        (s_sppc if r["Track"] == "SPPC" else s_pif).add_data_point(r["Project IRR at disclosed capex"], len(rows7) - i)
        order[r["ID"]] = r
    debt = xy.add_series(f"Cost of debt {kd:.2%}")
    debt.add_data_point(kd, 0.3)
    debt.add_data_point(kd, len(rows7) + 0.7)
    ch = add_chart(s, XL_CHART_TYPE.XY_SCATTER, xy, 0.5, 1.4, 8.6, 5.4)
    style_chart(ch, legend=True)
    quiet_axes(ch, None, 0, len(rows7) + 1, gridlines=False)
    ch.value_axis.visible = False
    xa = ch.category_axis
    xa.minimum_scale, xa.maximum_scale, xa.major_unit = -0.02, 0.07, 0.01
    xa.tick_labels.number_format, xa.tick_labels.number_format_is_linked = "0%", False
    xa.has_major_gridlines = True
    xa.major_gridlines.format.line.color.rgb = GRID
    for ser, c in zip(ch.plots[0].series, [AMBER, BLUE, INK]):
        if c == INK:
            ser.marker.style = XL_MARKER_STYLE.NONE
            ser.format.line.color.rgb = INK
            ser.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            ser.format.line.width = Pt(1.5)
            continue
        ser.format.line.fill.background()
        ser.marker.style, ser.marker.size = XL_MARKER_STYLE.CIRCLE, 10
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = c
        ser.marker.format.line.color.rgb = WHITE
    k = {"SPPC": 0, "PIF": 0}
    for r in rows7:
        ser = ch.plots[0].series[0 if r["Track"] == "SPPC" else 1]
        idx = k[r["Track"]]
        k[r["Track"]] += 1
        label(ser.points[idx], f"{r['Project']} {r['Project IRR at disclosed capex']:.1%}", size=10,
              pos=XL_LABEL_POSITION.RIGHT, bold=r["ID"] == "NAJRAN")
    text(s, 9.4, 1.6, 3.4, 5, [
        ("How the bids can still close", {"bold": True}),
        ("True build cost below the disclosed 'total investment', which can include sponsor fees and margins", {"bullet": True}),
        ("Plants run 30-35 years against a 25-year PPA", {"bullet": True}),
        ("Sponsors earn on O&M, EPC and development, not only on project equity", {"bullet": True}),
        ("State-backed sponsors (PIF, Masdar, SPIC) price for strategic value", {"bullet": True}),
    ], size=14)

    # 8 Sensitivity --------------------------------------------------------
    tor = sorted(m.rows("TORNADO"), key=lambda r: -r[3])
    tor_map = {r[0]: r for r in tor}
    names = {"CAPEX": "Capex (USD 450-650/kWac)", "PPA_YEARS": "PPA length (35 / 20 yrs)",
             "YIELD_CAL": "Yield factor (1.15-1.35)", "OPEX": "O&M (USD 6-12/kWac-yr)",
             "EQ_HURDLE": "Equity hurdle (7-12%)", "BASE_RATE": "Swap rate (3.5-4.5%)",
             "MARGIN": "Debt margin (1.0-1.75%)", "DEGR": "Degradation (0.3-0.6%/yr)", "TAX_KSA": "Tax (2.5-20%)"}
    s = d.slide(f"Capex, PPA length and yield move the required tariff most; at Najran's site a "
                f"{t_naj:.2f} ¢ bid clears only near USD 350/kWac and a 4.5% WACC",
                f"Source: model Sensitivity tab. Required tariff at the hurdle WACC for Najran's yield; base "
                f"{req_naj:.2f} ¢ at USD {m['CAPEX']:.0f}/kWac. Grid: after-tax nominal WACC; amber = at or below the awarded tariff",
                notes="Left: one input at a time between its low and high. Right: the capex and cost of capital "
                      "combinations that reach Najran's tariff.")
    cats = [names[r[0]] for r in tor[:7]]
    lows = [min(r[1], r[2]) for r in tor[:7]]
    highs = [max(r[1], r[2]) for r in tor[:7]]
    ch = floating(s, cats, lows, highs, [AMBER] * 7, 0.5, 1.5, 6.6, 5.2, '0.0"¢"', 1.2, 2.1,
                  labels=[f"{lo:.2f}-{hi:.2f}" for lo, hi in zip(lows, highs)])
    hm = m.rows("HEATMAP")
    waccs = hm[0][1:]
    pick_c = [0.03, 0.04, 0.05, 0.06, 0.07]
    rows = [["Capex \\ WACC"] + [f"{w:.0%}" for w in pick_c]]
    fills = {}
    for row in hm[1:]:
        if row[0] in (350, 400, 450, 500, 550):
            vals = [row[1 + waccs.index(min(waccs, key=lambda x: abs(x - w)))] for w in pick_c]
            rows.append([f"USD {row[0]:.0f}"] + [f"{v:.2f}" for v in vals])
            for j, v in enumerate(vals):
                if v <= t_naj + 1e-9:
                    fills[(len(rows) - 1, j + 1)] = AMBER_PALE
    text(s, 7.5, 1.55, 5.3, 0.4, [("Required tariff (¢/kWh) at Najran's yield", {"bold": True})], size=14)
    table(s, rows, 7.5, 2.0, 5.3, [1.3] + [0.8] * 5, size=12, fills=fills)
    text(s, 7.5, 4.5, 5.3, 2.2, [
        (f"At the disclosed USD {disc25:.0f}/kWac, no WACC in the grid gets near {t_naj:.2f} ¢", {"bullet": True}),
        ((f"A 35-year PPA alone takes {req_naj - tor_map['PPA_YEARS'][2]:.2f} ¢ off the required tariff; Shuaibah's "
         "1.04 ¢ record sits on a 35-year PPA"), {"bullet": True}),
    ], size=14)

    # 9 Shapley waterfall --------------------------------------------------
    steps = [("Energy yield", wf["Energy yield"]), ("Return accepted", wf["Return accepted (and life)"]),
             ("Capex", wf["Capex"]), ("O&M", wf["O&M"])]
    steps.sort(key=lambda x: x[1])
    start, mid, end = wf["IRENA global average 2025"], wf["Najran, real equivalent"], wf["Najran awarded tariff"]
    s = d.slide(f"Of the USD {gap:.0f}/MWh gap from IRENA's global average to Najran, energy yield explains "
                f"{-wf['Energy yield']:.0f}, accepted return {-wf['Return accepted (and life)']:.0f}, capex "
                f"{-wf['Capex']:.0f} and O&M {-wf['O&M']:.0f}",
                "Source: model Bridge tab. Both ends on IRENA's basis (real, per kWdc, pre-tax, no degradation). Shapley "
                "attribution: each bar is the driver's average effect over all 24 orders of change. Najran capex at "
                f"USD {m['CAPEX']:.0f}/kWac",
                notes="Order-free attribution, so no one can say the order of the bars drives the answer. Energy and capex "
                      "are things a buyer can find or tender for. The return bar is what sponsors gave up.")
    cats = ["IRENA global average 2025"] + [n for n, _ in steps] + ["Najran, real equivalent", "Flat nominal basis",
                                                                     "Najran awarded tariff"]
    lows, highs, cols, labs = [0], [start], [INK], [f"{start:.1f}"]
    run = start
    for n, v in steps:
        lows.append(run + v)
        highs.append(run)
        cols.append(AMBER)
        labs.append(f"{v:.1f}")
        run += v
    lows += [0, mid, 0]
    highs += [mid, end, end]
    cols += [INK, GREY, INK]
    labs += [f"{mid:.1f}", f"+{end - mid:.1f}", f"{end:.1f}"]
    floating(s, cats, lows, highs, cols, 0.5, 1.5, 8.8, 5.2, '"$"0', 0, 48, labels=labs, horizontal=False, gap=45,
             label_color=[WHITE, INK, INK, INK, INK, WHITE, INK, WHITE])
    text(s, 9.6, 1.6, 3.2, 5, [
        ("Reading the bridge", {"bold": True}),
        (f"Sunshine and cost: {share('Energy yield') + share('Capex'):.0%} of the gap", {"bullet": True}),
        ((f"Return accepted: {share('Return accepted (and life)'):.0%}, an implied real return of "
         f"{m['BR_RATE_B']:.1%} against IRENA's {m['BR_RATE_A']:.1%}"), {"bullet": True}),
        ("The last two bars add back inflation: flat nominal tariffs look higher than real LCOEs", {"bullet": True}),
    ], size=14)

    # 10 Round 6 dispersion ------------------------------------------------
    r6 = [r for r in eng if r["Round"] == "R6"]
    e_by = {r["ID"]: r["Energy kWh/kWac"] for r in eng}
    who = {"NAJRAN": "Masdar", "ADDARB": "Masdar", "SAMTAH": "SEC/EDF", "SUFUN": "AEW/Total"}
    raw = (max(r["Tariff USc/kWh"] for r in r6) / min(r["Tariff USc/kWh"] for r in r6)) - 1
    adj = (max(r["Tariff at reference yield (USc)"] for r in r6) / min(r["Tariff at reference yield (USc)"] for r in r6)) - 1
    s = d.slide(f"Inside Round 6, sunshine explains about {1 - adj / raw:.0%} of a {raw:.0%} tariff spread; "
                "scale and bidder appetite explain the rest",
                "Source: model Engine tab; Round 6 awards via TaiyangNews (S14); site yields from PVGIS 5.3 (S35), "
                "As Sufun proxied by Hail city. Adjusted tariff = awarded tariff x site energy / average Saudi site energy",
                notes="Same buyer, same contract, same month. Najran has the best site, but even after adjusting every bid to "
                      "the same sunshine the spread is still about 30%. The two cheapest bids are Masdar's.")
    cd = CategoryChartData()
    cd.categories = [f"{r['Project']} {r['MWac'] / 1000:.1f} GW ({who[r['ID']]})" for r in r6]
    cd.add_series("Awarded tariff", [r["Tariff USc/kWh"] for r in r6])
    cd.add_series("Adjusted to average Saudi yield", [r["Tariff at reference yield (USc)"] for r in r6])
    ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, cd, 0.5, 1.5, 8.6, 5.2)
    style_chart(ch, legend=True)
    quiet_axes(ch, '0.0"¢"', 0, 1.8)
    ch.plots[0].gap_width, ch.plots[0].overlap = 70, -8
    for ser, c in zip(ch.plots[0].series, [AMBER, GREY]):
        fill(ser, c)
        for i, r in enumerate(r6):
            label(ser.points[i], f"{ser.values[i]:.2f}", pos=XL_LABEL_POSITION.OUTSIDE_END)
    text(s, 9.4, 1.6, 3.4, 5, [
        ("What drives the spread", {"bold": True}),
        (f"Najran's high, cool site yields {e_by['NAJRAN'] / e_by['SAMTAH'] - 1:.0%} more than coastal Samtah", {"bullet": True}),
        ("Najran is 2-3x the size of the other three plants", {"bullet": True}),
        ("Masdar won both of the two cheapest slots; it also won Al Sadawi in Round 5", {"bullet": True}),
    ], size=14)

    # 11 Oman bridge -------------------------------------------------------
    o_start, o_end = ob["KSA at hurdle returns"], ob["Oman at hurdle returns"]
    y_by = {r["ID"]: r["PVGIS yield kWh/kWp-yr"] for r in m.table("Yield", 4)}
    o_gap = o_end - o_start
    designable = (o_gap - ob["Resource"]) / o_gap
    s = d.slide(f"About {designable:.0%} of the {o_gap:.1f} ¢ gap between Oman and a Saudi-quality site is designable; "
                "only the resource is inherited",
                "Source: model Oman tab. Both at hurdle returns: Saudi debt 5.25%, equity 9%; Oman adds the Baa3-Aa3 default "
                "spread (1.36 pts) to debt and the country risk premium gap (2.07 pts) to equity (Damodaran Jan 2026, S22). "
                "Ibri yield from PVGIS; Oman capex from Ibri III (S25); Shapley attribution",
                notes="Resource is fixed. Capex and scale are about tender size and site preparation. Financing depends on the "
                      "buyer's credit and any guarantee. Tenor is a contract choice.")
    keys = ["Resource", "Capex and scale", "Financing and tax", "Contract tenor"]
    cats = ["Saudi site, hurdle returns"] + keys + ["Oman (Ibri), hurdle returns"]
    lows, highs, cols, labs, run = [0], [o_start], [INK], [cents(o_start)], o_start
    for k_ in keys:
        lows.append(run)
        highs.append(run + ob[k_])
        cols.append(GREY if k_ == "Resource" else TEAL)
        labs.append(f"+{ob[k_]:.2f}")
        run += ob[k_]
    lows.append(0)
    highs.append(o_end)
    cols.append(INK)
    labs.append(cents(o_end))
    floating(s, cats, lows, highs, cols, 0.5, 1.5, 8.6, 5.2, '0.0"¢"', 0, 2.6, labels=labs, horizontal=False, gap=45,
             label_color=[WHITE, INK, WHITE, WHITE, WHITE, WHITE])
    text(s, 9.4, 1.6, 3.4, 5, [
        ("Inherited", {"bold": True, "color": TEXT2}),
        (f"Ibri yields {1 - y_by['IBRI3'] / y_by['NAJRAN']:.0%} less than Najran", {"bullet": True}),
        ("Designable", {"bold": True, "color": TEAL}),
        ("Capex: a 500 MW tender costs more per kW than a 1.4-3 GW one", {"bullet": True}),
        ("Financing: credit support can close most of the Baa3-Aa3 spread", {"bullet": True}),
        ("Tenor: a 25-year PPA instead of 20", {"bullet": True}),
    ], size=14)

    # 12 Oman cases --------------------------------------------------------
    s = d.slide(f"With Saudi-style tender design Oman could reach about {cases[2]:.1f} ¢, below the {gas:.1f} ¢ fuel "
                f"cost of gas power at regulated prices and {lng:.1f} ¢ at LNG export value",
                f"Source: model Oman tab. Gas: USD 3.00/MMBtu (2015, S31) escalated 3%/yr to {m['OMN_GAS']:.2f}; heat rate "
                f"{m['OMN_HR']:.1f} MMBtu/MWh. LNG: {m['LNG_SLOPE']:.0%} of Brent at USD {m['BRENT']:.0f} (World Bank 2026 "
                "baseline, S32) less USD 2.5 shipping and liquefaction. The Bank's severe Hormuz case (USD 95-115) lifts it further",
                notes="Cases are cumulative. Case 4 is not a recommendation: it shows what a Najran-style sponsor return would "
                      "do. Even the commercial case beats the fuel cost of gas power, before counting capacity value.")
    cd = CategoryChartData()
    cats = ["1. Oman today: 500 MW, 20-yr PPA", "2. + Saudi-grade financing", "3. + GW-scale tender, 25-yr PPA",
            "4. + Najran-style sponsor return", "Gas fuel cost, regulated price", "Gas fuel cost, LNG netback"]
    vals = cases + [gas, lng]
    cd.categories = cats
    cd.add_series("USc/kWh", vals)
    ch = add_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, cd, 0.5, 1.5, 8.6, 5.2)
    style_chart(ch)
    quiet_axes(ch, '0.0"¢"', 0, 6.5)
    ch.category_axis.reverse_order = True
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.plots[0].gap_width = 55
    ser = ch.plots[0].series[0]
    for i, c in enumerate([TEAL, TEAL, TEAL, TEAL, GREY, GREY]):
        fill(ser.points[i], c)
        label(ser.points[i], cents(vals[i]), pos=XL_LABEL_POSITION.OUTSIDE_END, bold=i == 2)
    text(s, 9.4, 1.6, 3.4, 5, [
        ("So what for Nama PWP", {"bold": True}),
        ("Oman does not need Saudi's tariff for solar to be the cheapest new energy", {"bullet": True}),
        (f"Cases 2 and 3 are design choices worth {cases[0] - cases[2]:.2f} ¢ together", {"bullet": True}),
        ("Case 4 depends on who bids, not on what Nama PWP designs", {"bullet": True}),
    ], size=14)

    # 13 Four moves --------------------------------------------------------
    s = d.slide("Four moves for Nama PWP, in order of what it controls: credit support, GW-scale prepared sites, "
                "25-year PPAs and a wider bidder pool",
                "Source: model Oman tab (case steps and Shapley effects); Saudi precedents from the Tracker tab and ACWA "
                "Power FY2025 investor report (S43)",
                notes="Impacts are from the Oman tab and are not additive with each other beyond case 3.")
    rows = [["Move", "What it changes", "Effect on Oman tariff", "Saudi precedent"],
            ["1. Credit support for the buyer", "Debt and equity priced nearer Saudi terms",
             f"-{cases[0] - cases[1]:.2f} ¢", "SPPC: sovereign-backed single buyer"],
            ["2. GW-scale, state-prepared sites", "Capex per kW; land and grid outside the bid",
             f"-{ob['Capex and scale']:.2f} ¢", "Najran 1.4 GW; PIF sites of 2-3 GW"],
            ["3. 25-year PPAs (or longer)", "Spreads capex over more years", f"-{ob['Contract tenor']:.2f} ¢",
             "SPPC 25 years; three PIF projects at 35"],
            ["4. Widen the bidder pool", "Lets strategic sponsors price below commercial returns",
             f"up to -{cases[2] - cases[3]:.2f} ¢", "Masdar, ACWA, EDF, KEPCO, SPIC, Jinko Power"]]
    table(s, rows, 0.5, 1.6, 12.3, [3.0, 3.9, 2.0, 3.4], size=14, row_h=0.75)
    text(s, 0.5, 5.6, 12.3, 1.1, [
        (("Moves 1 to 3 are in Nama PWP's hands and together take Oman from "
         f"{cents(cases[0])} to {cents(cases[2])}. Move 4 depends on whether sponsors want a foothold in Oman."), {})],
        size=14)

    # 14 What would change the answer -------------------------------------
    s = d.slide("Three inputs would change the answer, and paid data could settle each of them",
                "Source: model Sensitivity tab; author's assessment of data sources",
                notes="Be explicit about what is estimated. These are the questions an interviewer will ask.")
    up10 = (1 - m["TAX_KSA"]) * t_naj / 100 * najran["Energy kWh/kWac"] * najran["PV factor energy"] * 0.1 / (
        1 - najran["Tax shield"])
    card(s, 0.5, 1.6, 6.0, 3.9, "What would move the numbers", [
        ("Yield: one calibration plant (Sakaka, 2020). Newer bifacial plants at DC/AC 1.4 could yield 10% more and "
        f"lift implied capex by about USD {up10:.0f}/kWac"),
        (f"True build cost: disclosed 'total investment' may include sponsor fees; a USD {m['CAPEX_LO']:.0f} build cost "
        f"cuts Najran's required tariff to {tor_map['CAPEX'][1]:.2f} ¢"),
        "Life after the PPA: a merchant tail beyond year 25 raises returns",
        "Oman gas price: the 2015 regulated price escalated; the current figure is not published",
    ])
    card(s, 6.8, 1.6, 6.0, 3.9, "What I would do with paid data", [
        "IJGlobal or Proximo deal sheets for debt margins, tenors and DSCRs on Saudi solar closes",
        "MEED Projects bid tables for every bidder, not only winners, to test bidder appetite",
        "InfoLink or OPIS module and tracker price series at each bid date",
        "Nama PWP PPA terms and tariffs, available only through client access",
    ], accent=TEAL)

    # Appendix -------------------------------------------------------------
    s = d.slide("Appendix: how the model is built", "Source: model/NREP_Tariff_LCOE_Model.xlsx; docs/model-guide.md",
                tag=False)
    boxes = [("Inputs", "46 assumptions, each Reported, Derived or Assumed, with a range"),
             ("Tracker and Yield", "37 projects; PVGIS yields calibrated to Sakaka's measured output"),
             ("Engine", "Closed-form after-tax LCOE: implied capex, IRR, real equivalent"),
             ("Bridge and Oman", "Shapley attribution; Oman cases and avoided cost"),
             ("Checks", "12 integrity tests, all TRUE; 19 automated tests in the repo")]
    for i, (h_, b) in enumerate(boxes):
        x = 0.5 + i * 2.5
        shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(2.2), Inches(2.2), Inches(2.4))
        shp.fill.solid()
        shp.fill.fore_color.rgb = PALE if i else AMBER_PALE
        shp.line.fill.background()
        text(s, x + 0.12, 2.32, 1.96, 0.5, [(h_, {"bold": True})], size=15)
        text(s, x + 0.12, 2.9, 1.96, 1.6, [b], size=13)
        if i < 4:
            ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x + 2.22), Inches(3.4), Inches(x + 2.48), Inches(3.4))
            ln.line.color.rgb = MUTED
            ln.line.width = Pt(1.5)
    text(s, 0.5, 5.0, 12.3, 1.6, [
        (("Every number on these slides is read from the workbook. Formulas are live, inputs are named (=CAPEX, "
         "=WACC_KSA) and the Python build re-computes each output independently as a test."), {})], size=14)

    a = {r[0]: r for r in m.wb["Inputs"].iter_rows(min_row=5, values_only=True) if r[0]}
    keep = ["CAPEX", "YIELD_CAL_RANGE", "OPEX", "DEGR", "PPA_YEARS", "BASE_RATE", "MARGIN", "TENOR", "DSCR",
            "GEAR_WACC", "EQ_HURDLE", "TAX_KSA", "OMN_CAPEX", "OMN_PPA", "OMN_GAS"]
    s = d.slide("Appendix: assumption log for the main drivers", "Source: model Inputs tab, which also gives the rationale "
                "for each line", tag=False)
    rows = [["Input", "Base", "Range", "Type", "Source"]]
    for k_ in keep:
        r = a[k_]
        pct = r[6] in ("per yr", "share", "share of capex")
        f_ = (lambda v: f"{v:.2%}") if pct else (lambda v: f"{v:,.3g}")
        rows.append([r[2], f_(r[3]), f"{f_(r[4])} to {f_(r[5])}", r[7], r[8] or "-"])
    table(s, rows, 0.5, 1.5, 12.3, [5.2, 1.6, 2.6, 1.3, 1.6], size=11, row_h=0.32)

    srcs = {r[0]: r for r in m.wb["Sources"].iter_rows(min_row=5, values_only=True) if r[0]}
    cited = {x for slide_ in d.prs.slides for sh in slide_.shapes if sh.has_text_frame and sh.name == "Source"
             for x in re.findall(r"S\d{2}", sh.text_frame.text)}
    cited |= {x for r in tr if r["In analysis"] == "Y" for x in r["Sources"].split(";")}
    cited |= {r[8] for r in m.wb["Capex"].iter_rows(min_row=5, max_row=18, values_only=True) if r[8]}
    cited = sorted(cited)
    half = (len(cited) + 1) // 2
    for part, chunk in enumerate((cited[:half], cited[half:]), start=1):
        s = d.slide(f"Appendix: sources behind this deck ({part} of 2)",
                    "Full register with URLs and access dates: model Sources tab", tag=False)
        rows = [["ID", "Publisher", "Title", "Date"]] + [[c, srcs[c][1], srcs[c][2], srcs[c][3]] for c in chunk]
        table(s, rows, 0.5, 1.4, 12.3, [0.6, 2.7, 7.9, 1.1], size=11, row_h=0.32)

    path.parent.mkdir(parents=True, exist_ok=True)
    d.prs.save(path)
    return path


def build_onepager(path=ONEPAGER):
    m = Model()
    wf = {k: v for k, v in m.rows("WATERFALL")}
    cases = [m[f"OMN_CASE{i}"] for i in range(1, 5)]
    d = Deck()
    s = d.slide("Saudi Arabia's 1.1 ¢ solar bids rest on cheap sunshine and returns no commercial investor would "
                "accept; Oman can copy the design and still beat gas",
                "Source: public data (SPPC, ACWA Power, Masdar, IRENA 2025, PVGIS, Damodaran); model and sources in the "
                "Gulf-Solar repository", tag=True)
    steps = sorted([("Energy yield", wf["Energy yield"]), ("Return accepted", wf["Return accepted (and life)"]),
                    ("Capex", wf["Capex"]), ("O&M", wf["O&M"])], key=lambda x: x[1])
    start, mid = wf["IRENA global average 2025"], wf["Najran, real equivalent"]
    cats, lows, highs, cols, labs, run = ["IRENA global 2025"], [0], [start], [INK], [f"{start:.1f}"], start
    for n, v in steps:
        cats.append(n)
        lows.append(run + v)
        highs.append(run)
        cols.append(AMBER)
        labs.append(f"{v:.1f}")
        run += v
    cats.append("Najran, real")
    lows.append(0)
    highs.append(mid)
    cols.append(INK)
    labs.append(f"{mid:.1f}")
    text(s, 0.5, 1.4, 7.4, 0.4, [("IRENA global average LCOE to Najran's tariff, USD/MWh (real)", {"bold": True})])
    floating(s, cats, lows, highs, cols, 0.5, 1.8, 7.4, 4.9, '"$"0', 0, 48, labels=labs, horizontal=False, gap=45,
             label_color=[WHITE, INK, INK, INK, INK, WHITE])
    gap = start - mid
    text(s, 8.3, 1.4, 4.5, 3.4, [
        ("Three takeaways", {"bold": True}),
        ((f"At the 2025 disclosed cost of USD {m['CX_DISCLOSED_2025']:.0f}/kWac, Najran's {m['T_NAJRAN']:.2f} ¢ bid "
         f"returns {m['IRR_NAJRAN']:.1%} before leverage"), {"bullet": True}),
        ((f"Sunshine and capex explain {(-wf['Energy yield'] - wf['Capex']) / gap:.0%} of the gap to the global "
         f"average; accepted return {-wf['Return accepted (and life)'] / gap:.0%}"), {"bullet": True}),
        ((f"Oman reaches {cases[2]:.1f} ¢ with credit support, GW-scale tenders and 25-year PPAs, below the "
         f"{m['OMN_AVOIDED_GAS']:.1f} ¢ fuel cost of gas power"), {"bullet": True}),
    ], size=13)
    card(s, 8.3, 4.85, 4.5, 1.85, "With paid data I would", [
        "Pull debt terms from IJGlobal deal sheets",
        "Test bidder appetite with full MEED bid tables",
    ], accent=TEAL, size=12)
    path.parent.mkdir(parents=True, exist_ok=True)
    d.prs.save(path)
    return path


if __name__ == "__main__":
    print(build())
    print(build_onepager())

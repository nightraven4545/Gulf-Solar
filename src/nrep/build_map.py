"""Resource map: PVGIS tracked yield across Saudi Arabia and Oman, with awarded projects on top."""

import csv
import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.colors import LinearSegmentedColormap  # noqa: E402
from matplotlib.patches import PathPatch  # noqa: E402
from matplotlib.path import Path as MPath  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
OUT = ROOT / "docs" / "map.png"
INK, TEXT2, GRID = "#1F2A36", "#52514E", "#D9DEE3"
COLORS = {"SPPC": "#C2780B", "PIF": "#3A5FA8", "IPP tender": "#1A9A8A"}
LABELS = {"NAJRAN": "Najran", "SHUAIBAH1": "Shuaibah 1", "SAKAKA": "Sakaka", "SADAWI": "Al Sadawi",
          "BISHA": "Bisha", "SAMTAH": "Samtah", "IBRI3": "Ibri III (Oman)"}


def rows(name):
    return list(csv.DictReader((DATA / name).read_text().splitlines()))


def ring_paths(geom):
    polys = [geom["coordinates"]] if geom["type"] == "Polygon" else geom["coordinates"]
    return [MPath(p[0]) for p in polys]


def render(out=OUT, model=ROOT / "model" / "NREP_Tariff_LCOE_Model.xlsx"):
    plt.rcParams["font.family"] = "Liberation Sans"
    borders = json.loads((DATA / "raw" / "borders_mena.geojson").read_text())["features"]
    grid = rows("yield_grid.csv")
    tariffs = {r[0]: r[9] for r in load_workbook(model, data_only=True)["Engine"].iter_rows(min_row=5, values_only=True)
               if r[0]}

    fig, ax = plt.subplots(figsize=(7.6, 5.4), dpi=220)
    ax.set_xlim(34, 60.5)
    ax.set_ylim(15.5, 33)
    ax.set_aspect(1 / 0.93)  # rough cos(latitude) correction around 22N
    ax.axis("off")
    for f in borders:
        for p in ring_paths(f["geometry"]):
            ax.add_patch(PathPatch(p, facecolor="#F4F5F7", edgecolor="#C9CED4", lw=0.4, zorder=1))

    cmap = LinearSegmentedColormap.from_list("sun", ["#FFF4DC", "#F2C46E", "#C2780B"])
    xs = [float(g["lon"]) for g in grid]
    ys = [float(g["lat"]) for g in grid]
    zs = [float(g["yield_kwh_kwp_yr"]) for g in grid]
    levels = [2100 + 100 * k for k in range(8)]
    fill = ax.tricontourf(xs, ys, zs, levels=levels, cmap=cmap, zorder=2, extend="both")
    clip = [p for f in borders if f["properties"]["iso"] in ("SAU", "OMN") for p in ring_paths(f["geometry"])]
    verts, codes = [], []
    for p in clip:
        verts += list(p.vertices)
        codes += [MPath.MOVETO] + [MPath.LINETO] * (len(p.vertices) - 2) + [MPath.CLOSEPOLY]
    clip_patch = PathPatch(MPath(verts, codes), facecolor="none", edgecolor=INK, lw=0.6, zorder=3)
    ax.add_patch(clip_patch)
    fill.set_clip_path(clip_patch)

    for t in rows("tracker.csv"):
        if not t["lat"] or t["tech"] not in ("PV", "PV+BESS") or t["country"] not in ("KSA", "Oman"):
            continue
        x, y, mw = float(t["lon"]), float(t["lat"]), float(t["mw_ac"])
        ax.scatter(x, y, s=mw / 9, color=COLORS[t["track"]], edgecolor="white", lw=0.8, alpha=0.9, zorder=5)
        if t["id"] in LABELS:
            tag = f"{LABELS[t['id']]}"
            if isinstance(tariffs.get(t["id"]), float):
                tag += f" {tariffs[t['id']]:.2f}¢"
            ax.annotate(tag, (x, y), xytext=(6, 4), textcoords="offset points", fontsize=7.5, color=INK, zorder=6)

    for text_, x, y in (("Red Sea", 37.4, 21.0), ("Arabian Gulf", 51.2, 27.6), ("Strait of Hormuz", 55.2, 27.0)):
        ax.text(x, y, text_, fontsize=7, color=TEXT2, style="italic", zorder=6)
    cbar = fig.colorbar(fill, ax=ax, orientation="horizontal", fraction=0.04, pad=0.02, shrink=0.55)
    cbar.set_label("Tracked PV yield, kWh/kWp per year (PVGIS 5.3)", fontsize=8, color=TEXT2)
    cbar.ax.tick_params(labelsize=7, colors=TEXT2)
    cbar.outline.set_visible(False)
    for label, c in (("SPPC tender", COLORS["SPPC"]), ("PIF bilateral", COLORS["PIF"]), ("Oman (Nama PWP)", COLORS["IPP tender"])):
        ax.scatter([], [], s=60, color=c, edgecolor="white", label=label)
    for mw in (500, 2000):
        ax.scatter([], [], s=mw / 9, facecolor="none", edgecolor=TEXT2, label=f"{mw:,} MW")
    ax.legend(loc="lower left", fontsize=7.5, frameon=False, labelspacing=1.1, borderpad=0.2)
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out


if __name__ == "__main__":
    print(render())

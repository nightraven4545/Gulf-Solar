"""Pull public inputs: PVGIS site yields and the KAPSARC project list. Raw responses are kept in data/raw."""

import csv
import json
import urllib.request
from datetime import date
from pathlib import Path

from matplotlib.path import Path as Polygon

DATA = Path(__file__).resolve().parents[2] / "data"
PVGIS = "https://re.jrc.ec.europa.eu/api/v5_3/PVcalc"
BORDERS = "https://raw.githubusercontent.com/nvkelso/natural-earth-vector/master/geojson/ne_50m_admin_0_countries.geojson"
REGION = {"SAU", "OMN", "ARE", "QAT", "KWT", "BHR", "YEM", "IRQ", "JOR", "EGY", "IRN", "ISR", "PSX", "SDN", "ERI", "DJI",
          "ETH", "SYR", "LBN"}
KAPSARC = (
    "https://datasource.kapsarc.org/api/explore/v2.1/catalog/datasets/"
    "saudi-arabia-planned-and-installed-renewables-by-project/records?limit=100"
)


def get_json(url):
    with urllib.request.urlopen(url, timeout=120) as r:
        return json.load(r)


def pvgis_tracked(lat, lon):
    # Horizontal N-S single-axis tracker; loss=0 so system losses stay explicit in the model.
    q = f"?lat={lat}&lon={lon}&peakpower=1&loss=0&inclined_axis=1&inclinedaxisangle=0&outputformat=json"
    return get_json(PVGIS + q)


def fetch_yields():
    out_dir = DATA / "raw" / "pvgis"
    out_dir.mkdir(parents=True, exist_ok=True)
    sites = {}
    for r in csv.DictReader((DATA / "tracker.csv").read_text().splitlines()):
        if r["lat"] and r["lon"]:
            sites.setdefault((r["lat"], r["lon"]), []).append(r["id"])

    rows = []
    for (lat, lon), ids in sorted(sites.items()):
        raw = out_dir / f"{lat}_{lon}.json"
        if not raw.exists():
            raw.write_text(json.dumps(pvgis_tracked(lat, lon)))
        d = json.loads(raw.read_text())
        t = d["outputs"]["totals"]["inclined_axis"]
        rows.append({
            "lat": lat, "lon": lon, "projects": ";".join(ids),
            "elevation_m": d["inputs"]["location"]["elevation"],
            "radiation_db": d["inputs"]["meteo_data"]["radiation_db"],
            "poa_kwh_m2_yr": t["H(i)_y"],
            "yield_kwh_kwp_yr": t["E_y"],
            "pvgis_loss_pct": t["l_total"],
        })

    with open(DATA / "yield.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader()
        w.writerows(rows)
    return rows


def fetch_borders():
    """Natural Earth 1:50m boundaries (public domain), trimmed to the region."""
    path = DATA / "raw" / "borders_mena.geojson"
    if not path.exists():
        d = get_json(BORDERS)
        feats = [{"type": "Feature", "properties": {"iso": f["properties"]["ADM0_A3"], "name": f["properties"]["NAME"]},
                  "geometry": f["geometry"]} for f in d["features"] if f["properties"]["ADM0_A3"] in REGION]
        path.write_text(json.dumps({"type": "FeatureCollection", "features": feats}))
    return json.loads(path.read_text())


def polygons(geometry):
    rings = [geometry["coordinates"]] if geometry["type"] == "Polygon" else geometry["coordinates"]
    return [Polygon(poly[0]) for poly in rings]


def fetch_grid(step=1.0):
    """PVGIS tracked yield on a regular grid over Saudi Arabia and Oman, for the resource map."""
    shapes = {f["properties"]["iso"]: polygons(f["geometry"]) for f in fetch_borders()["features"]
              if f["properties"]["iso"] in ("SAU", "OMN")}
    out_dir = DATA / "raw" / "pvgis_grid"
    out_dir.mkdir(parents=True, exist_ok=True)
    rows = []
    lat = 16.5
    while lat < 32.6:
        lon = 34.5
        while lon < 60.0:
            iso = next((k for k, ps in shapes.items() if any(p.contains_point((lon, lat)) for p in ps)), None)
            if iso:
                raw = out_dir / f"{lat:.1f}_{lon:.1f}.json"
                if not raw.exists():
                    raw.write_text(json.dumps(pvgis_tracked(lat, lon)))
                t = json.loads(raw.read_text())["outputs"]["totals"]["inclined_axis"]
                if t["E_y"] is not None:  # PVGIS returns nulls for a few cells with missing satellite data
                    rows.append({"lat": lat, "lon": lon, "country": iso, "yield_kwh_kwp_yr": t["E_y"]})
            lon += step
        lat += step
    with open(DATA / "yield_grid.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader()
        w.writerows(rows)
    return rows


def fetch_kapsarc():
    path = DATA / "raw" / f"kapsarc_renewables_by_project_{date.today()}.json"
    path.write_text(json.dumps(get_json(KAPSARC)))
    return path


if __name__ == "__main__":
    for r in fetch_yields():
        print(r["projects"], r["yield_kwh_kwp_yr"])
    print(len(fetch_grid()), "grid points")
    print(fetch_kapsarc())

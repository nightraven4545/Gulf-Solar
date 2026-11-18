"""Pull public inputs: PVGIS site yields and the KAPSARC project list. Raw responses are kept in data/raw."""

import csv
import json
import urllib.request
from datetime import date
from pathlib import Path

DATA = Path(__file__).resolve().parents[2] / "data"
PVGIS = "https://re.jrc.ec.europa.eu/api/v5_3/PVcalc"
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


def fetch_kapsarc():
    path = DATA / "raw" / f"kapsarc_renewables_by_project_{date.today()}.json"
    path.write_text(json.dumps(get_json(KAPSARC)))
    return path


if __name__ == "__main__":
    for r in fetch_yields():
        print(r["projects"], r["yield_kwh_kwp_yr"])

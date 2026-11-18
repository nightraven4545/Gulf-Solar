import csv
from pathlib import Path

DATA = Path(__file__).resolve().parents[1] / "data"


def load(name):
    return list(csv.DictReader((DATA / name).read_text().splitlines()))


def test_every_reference_resolves():
    ids = {r["id"] for r in load("sources.csv")}
    refs = [s for r in load("tracker.csv") for s in r["source_ids"].split(";")]
    refs += [s for r in load("assumptions.csv") for s in r["source_id"].split(";") if s]
    refs += [r["source_id"] for r in load("capex_evidence.csv")]
    refs += [r["source_id"] for name in ("milestones.csv", "rates.csv") for r in load(name)]
    refs += [s for r in load("value_chain.csv") for s in r["source_id"].split(";")]
    assert set(refs) <= ids


def test_reported_and_derived_inputs_are_sourced():
    assert all(r["source_id"] for r in load("assumptions.csv") if r["type"] != "Assumed")


def test_halala_and_cent_tariffs_agree_at_the_peg():
    for r in load("tracker.csv"):
        if r["tariff_usc"] and r["tariff_halala"]:
            assert abs(float(r["tariff_halala"]) / 3.75 - float(r["tariff_usc"])) < 1e-4, r["id"]


def test_analysed_rows_have_a_tariff():
    for r in load("tracker.csv"):
        if r["in_analysis"] == "Y":
            assert r["tariff_usc"] or r["tariff_halala"] or r["tariff_sar"], r["id"]


def test_ranges_bracket_base_values():
    for r in load("assumptions.csv"):
        assert float(r["low"]) <= float(r["value"]) <= float(r["high"]), r["id"]

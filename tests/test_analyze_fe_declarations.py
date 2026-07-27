"""Tests for scripts/analyze_fe_declarations.py — the SIAM BS-VI fuel-efficiency
analysis, which is the repo's real Python compute (the cost model itself is a pile
of Excel formula strings built by build_model.py; see test_build_model.py).

Covers the pure helpers (desc, base_model_key), the E20 mileage-drop constants,
and an end-to-end main() run driven by a synthetic in-memory declarations CSV with
all output paths redirected to a tmp dir — verifying the by-fuel FE aggregation,
the same-nameplate petrol->CNG distance uplift, and the E20 penalty arithmetic.
"""
import csv

import pytest

import analyze_fe_declarations as afe


# --------------------------- desc() : descriptive stats ---------------------------

def test_desc_even_length_list():
    d = afe.desc([10, 20, 30, 40])
    assert d["n"] == 4
    assert d["mean"] == 25.0
    assert d["median"] == 25.0          # (20+30)/2
    assert d["min"] == 10
    assert d["max"] == 40


def test_desc_odd_length_list_median_is_middle():
    d = afe.desc([1, 2, 3, 4, 100])
    assert d["n"] == 5
    assert d["median"] == 3             # middle element, robust to the outlier
    assert d["mean"] == 22.0            # (1+2+3+4+100)/5


def test_desc_rounds_to_two_places():
    d = afe.desc([1.0, 2.0])            # mean 1.5 exact; check rounding path
    assert d["mean"] == 1.5
    assert afe.desc([1, 1, 2])["mean"] == 1.33   # round(4/3, 2)


# --------------------------- base_model_key() : nameplate normaliser --------------

def test_base_model_key_strips_trim_and_fuel_tokens():
    assert afe.base_model_key("Maruti Swift VXI CNG") == "maruti swift"


def test_base_model_key_drops_engine_numbers_and_everything_after():
    assert afe.base_model_key("KWID 0.8L MT") == "kwid"


def test_base_model_key_collapses_petrol_and_cng_variants_together():
    # This is the whole point: a petrol variant and its factory-CNG twin must
    # map to the same key so the uplift pairing in main() can find them.
    assert afe.base_model_key("Alpha VXI") == afe.base_model_key("Alpha CNG") == "alpha"


# --------------------------- E20 mileage-drop constants ---------------------------

def test_e20_drop_band_is_ordered_low_mid_high():
    assert afe.E20_LOW < afe.E20_MID < afe.E20_HIGH
    assert (afe.E20_LOW, afe.E20_MID, afe.E20_HIGH) == (0.02, 0.04, 0.06)


def test_cited_calibration_constants_present():
    assert afe.CITED["hatchback petrol"] == 20.3
    assert afe.CITED["CNG car (km/kg)"] == 27.4


# --------------------------- main() : end-to-end (synthetic CSV, tmp outputs) -----

_HEADER = ["sno", "manufacturer", "model_variant", "kerb_weight_kg", "engine_cc",
           "emission_stage", "fuel_type", "declared_fe", "declared_co2_gkm"]

# Two petrol, one diesel, one CNG. "Alpha" exists in both petrol and CNG form so a
# petrol->CNG pair is produced. Petrol FE = {20, 16} -> mean 18.0.
_ROWS = [
    [1, "M", "Alpha VXI", 950, 1000, "BS-VI", "Gasoline", 20.0, 110.0],
    [2, "M", "Beta LXI", 1300, 1500, "BS-VI", "Gasoline", 16.0, 130.0],
    [3, "M", "Beta LXI", 1300, 1500, "BS-VI", "Diesel", 22.0, 120.0],
    [4, "M", "Alpha CNG", 970, 1000, "BS-VI", "CNG", 26.0, 90.0],
]


@pytest.fixture
def run_main(tmp_path, monkeypatch):
    """Write a synthetic declarations CSV, point all module paths at tmp, run main()."""
    src = tmp_path / "siam.csv"
    with src.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(_HEADER)
        w.writerows(_ROWS)

    out_fuel = tmp_path / "fe_by_fuel.csv"
    out_pairs = tmp_path / "fe_pairs.csv"
    out_md = tmp_path / "analysis.md"
    # main()'s closing print() does OUT_*.relative_to(ROOT); keep outputs under ROOT.
    monkeypatch.setattr(afe, "ROOT", tmp_path)
    monkeypatch.setattr(afe, "SRC", src)
    monkeypatch.setattr(afe, "OUT_FUEL", out_fuel)
    monkeypatch.setattr(afe, "OUT_PAIRS", out_pairs)
    monkeypatch.setattr(afe, "OUT_MD", out_md)

    afe.main()
    return out_fuel, out_pairs, out_md


def _read(path):
    with path.open() as f:
        return list(csv.DictReader(f))


def test_main_fuel_summary_aggregates_petrol_correctly(run_main):
    out_fuel, _, _ = run_main
    by_fuel = {r["fuel_type"]: r for r in _read(out_fuel)}
    petrol = by_fuel["Gasoline"]
    assert int(petrol["n"]) == 2
    assert float(petrol["fe_mean"]) == 18.0        # (20 + 16) / 2
    assert float(petrol["fe_min"]) == 16.0
    assert float(petrol["fe_max"]) == 20.0
    assert float(by_fuel["CNG"]["fe_mean"]) == 26.0


def test_main_petrol_to_cng_uplift(run_main):
    _, out_pairs, _ = run_main
    pairs = _read(out_pairs)
    alpha = next(p for p in pairs if p["model_key"] == "alpha")
    assert float(alpha["petrol_kmpl"]) == 20.0
    assert float(alpha["cng_km_per_kg"]) == 26.0
    assert float(alpha["distance_uplift_pct"]) == 30.0   # (26/20 - 1) * 100


def test_main_e20_penalty_uses_mean_petrol_fe(run_main):
    # petrol mean FE = 18.0; E20 @ 4% central drop -> 18.0 * 0.96 = 17.28 kmpl.
    _, _, out_md = run_main
    text = out_md.read_text()
    assert "17.28" in text
    assert "18.00 kmpl" in text          # reported mean petrol FE

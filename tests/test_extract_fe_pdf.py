"""Unit tests for the pure row parser in scripts/extract_fe_pdf.py (no PDF needed)."""
from __future__ import annotations

import pytest

extract_fe_pdf = pytest.importorskip("extract_fe_pdf")  # needs pdfplumber


def _line(sno, name, kerb, cc, fuel, fe, co2):
    return f"{sno}{name} {kerb} {cc} Bharat Stage VI {fuel} {fe} {co2}"


def test_parses_petrol_row():
    line = _line(1, "Maruti Suzuki India Limited Alto 800 LXI", 760, 796, "Gasoline", 22.05, 104)
    row = extract_fe_pdf.parse_line(line)
    assert row == {
        "sno": 1, "manufacturer": "Maruti Suzuki India Limited", "model_variant": "Alto 800 LXI",
        "kerb_weight_kg": 760, "engine_cc": 796, "emission_stage": "BS-VI",
        "fuel_type": "Gasoline", "declared_fe": 22.05, "declared_co2_gkm": 104.0,
    }


def test_thousands_separator_variants_are_collapsed():
    for kerb in ("1,000", "1 ,000", "1, 000"):
        row = extract_fe_pdf.parse_line(_line(7, "Kia Motors Seltos", kerb, "1,497", "Diesel", 18.0, 120))
        assert row["kerb_weight_kg"] == 1000 and row["engine_cc"] == 1497


def test_electric_has_no_engine_cc():
    row = extract_fe_pdf.parse_line(_line(9, "TATA Motors Limited Nexon EV", 1400, "-", "Electric", 6.0, 0))
    assert row["engine_cc"] is None and row["fuel_type"] == "Electric"


def test_gasoline_hybrid_beats_gasoline():
    row = extract_fe_pdf.parse_line(_line(3, "Toyota Kirloskar Motor Camry", 1600, 2487, "Gasoline Hybrid", 19.1, 98))
    assert row["fuel_type"] == "Gasoline Hybrid"


@pytest.mark.parametrize("line", ["S.no. Manufacturer Model", "", "random text 123"])
def test_non_data_lines_return_none(line):
    assert extract_fe_pdf.parse_line(line) is None


def test_split_oem_fallback_uses_two_tokens():
    assert extract_fe_pdf.split_oem("Foo Bar Baz Qux") == ("Foo Bar", "Baz Qux")


def test_renualt_typo_is_canonicalised():
    assert extract_fe_pdf.split_oem("Renualt India Kwid")[0] == "Renault India"


def test_committed_csv_rows_round_trip_through_parser():
    """Rebuild a PDF-style line from every committed CSV row and check the parser recovers it."""
    import csv
    from pathlib import Path

    src = Path(__file__).resolve().parent.parent / "data" / "siam_bsvi_fe_2020_4w.csv"
    rows = list(csv.DictReader(src.open()))
    assert len(rows) == 303
    for r in rows:
        line = _line(r["sno"], f"{r['manufacturer']} {r['model_variant']}", r["kerb_weight_kg"],
                     r["engine_cc"] or "-", r["fuel_type"], r["declared_fe"], r["declared_co2_gkm"])
        got = extract_fe_pdf.parse_line(line)
        assert got["model_variant"] == r["model_variant"]
        assert got["fuel_type"] == r["fuel_type"]
        assert got["declared_fe"] == float(r["declared_fe"])

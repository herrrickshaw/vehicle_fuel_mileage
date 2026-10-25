"""extract_fe_pdf.py: tail-anchored row regex and OEM split, on hand-written PDF text lines."""
import csv

import pytest

import extract_fe_pdf as m


class FakePage:
    def __init__(self, text):
        self.text = text

    def extract_text(self):
        return self.text


class FakePdf:
    def __init__(self, pages):
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


@pytest.fixture
def extract(tmp_path, monkeypatch):
    def _run(*page_texts):
        pdf = tmp_path / "in.pdf"
        pdf.write_bytes(b"%PDF")
        out = tmp_path / "out.csv"
        monkeypatch.setattr(m, "PDF", pdf)
        monkeypatch.setattr(m, "OUT", out)
        monkeypatch.setattr(m, "ROOT", tmp_path)
        monkeypatch.setattr(m.pdfplumber, "open", lambda p: FakePdf([FakePage(t) for t in page_texts]))
        m.main()
        with open(out, newline="") as fh:
            return list(csv.DictReader(fh))
    return _run


def test_rows_parsed_with_separators_dash_cc_and_hybrid(extract, capsys):
    rows = extract(
        "S.no. Manufacturer Model Kerb CC Stage Fuel FE CO2\n"
        "1Maruti Suzuki India Limited Alto 800 LXI 755 796 Bharat Stage VI Gasoline 22.05 104.8\n"
        "2TATA Motors Limited Nexon EV 1 ,400 - Bharat Stage VI Electric 15.6 0\n",
        "S.no. Manufacturer Model Kerb CC Stage Fuel FE CO2\n"
        "2TATA Motors Limited Nexon EV 1 ,400 - Bharat Stage VI Electric 15.6 0\n"     # repeated row
        "3Toyota Kirloskar Motor Camry 1,665 2,487 Bharat Stage VI Gasoline Hybrid 23.27 99.6\n"
        "4Renualt India Kwid 1.0 RXT 769 999 Bharat Stage VI Gasoline 22.3 104\n",
    )
    assert [r["sno"] for r in rows] == ["1", "2", "3", "4"]
    alto, nexon, camry, kwid = rows
    assert (alto["manufacturer"], alto["model_variant"], alto["kerb_weight_kg"], alto["engine_cc"]) == \
        ("Maruti Suzuki India Limited", "Alto 800 LXI", "755", "796")
    assert (nexon["kerb_weight_kg"], nexon["engine_cc"], nexon["fuel_type"]) == ("1400", "", "Electric")
    assert (camry["kerb_weight_kg"], camry["engine_cc"], camry["fuel_type"], camry["declared_fe"]) == \
        ("1665", "2487", "Gasoline Hybrid", "23.27")
    assert kwid["manufacturer"] == "Renault India"  # source typo canonicalised
    assert "no gaps" in capsys.readouterr().err


def test_gap_warning_and_unknown_oem_fallback(extract, capsys):
    rows = extract("1Acme Motors X1 900 1199 Bharat Stage VI CNG 30.1 95\n"
                   "3Acme Motors X2 950 1199 Bharat Stage VI Diesel 24.0 110\n")
    assert rows[0]["manufacturer"] == "Acme Motors" and rows[0]["model_variant"] == "X1"
    assert "missing S.no rows: [2]" in capsys.readouterr().err


def test_split_oem_direct():
    assert m.split_oem("BMW India 330i M Sport") == ("BMW India", "330i M Sport")

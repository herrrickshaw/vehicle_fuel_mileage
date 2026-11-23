"""Kerb-weight band edges, CNG pairing edge cases and the E20 table in analyze_fe_declarations.main(),
plus scripts/recalc.py (LibreOffice driver) with `soffice` and the subprocess faked."""
import csv
import json
import os
import subprocess

import openpyxl
import pytest

import analyze_fe_declarations as afe
import recalc

_HEADER = ["sno", "manufacturer", "model_variant", "kerb_weight_kg", "engine_cc",
           "emission_stage", "fuel_type", "declared_fe", "declared_co2_gkm"]


@pytest.fixture
def run_rows(tmp_path, monkeypatch):
    def _run(rows):
        src = tmp_path / "siam.csv"
        with src.open("w", newline="") as f:
            w = csv.writer(f)
            w.writerow(_HEADER)
            w.writerows(rows)
        outs = {k: tmp_path / f"{k}.csv" for k in ("fuel", "pairs")}
        md = tmp_path / "a.md"
        monkeypatch.setattr(afe, "ROOT", tmp_path)
        monkeypatch.setattr(afe, "SRC", src)
        monkeypatch.setattr(afe, "OUT_FUEL", outs["fuel"])
        monkeypatch.setattr(afe, "OUT_PAIRS", outs["pairs"])
        monkeypatch.setattr(afe, "OUT_MD", md)
        afe.main()
        return outs, md.read_text()
    return _run


def row(i, model, kg, fuel, fe, cc=1000, co2=100.0):
    return [i, "M", model, kg, cc, "BS-VI", fuel, fe, co2]


def test_load_coerces_types_and_blank_engine_cc(tmp_path, monkeypatch):
    src = tmp_path / "s.csv"
    with src.open("w", newline="") as f:
        csv.writer(f).writerows([_HEADER, row(1, "A", 900, "Gasoline", 20, cc=""), row(2, "B", 901, "CNG", 26.5)])
    monkeypatch.setattr(afe, "SRC", src)
    a, b = afe.load()
    assert a["engine_cc"] is None and b["engine_cc"] == 1000
    assert (a["kerb_weight_kg"], a["declared_fe"], a["declared_co2_gkm"]) == (900, 20.0, 100.0)


def test_band_edges_are_lower_inclusive_upper_exclusive(run_rows):
    # one petrol model at every edge; FE doubles as an id so each band mean is unambiguous
    kgs = [(0, 1), (899, 2), (900, 3), (1049, 4), (1050, 5), (1199, 6), (1200, 7), (1499, 8), (1500, 9), (9999, 10), (10_000, 11), (50_000, 12)]
    rows = [row(i + 1, f"Model{i} VXI", kg, "Gasoline", fe) for i, (kg, fe) in enumerate(kgs)]
    rows.append(row(99, "Zed CNG", 950, "CNG", 26))
    rows.append(row(98, "Zed VXI", 950, "Gasoline", 20.0))
    _, md = run_rows(rows)
    table = {l.split("|")[1].strip(): l for l in md.splitlines() if l.startswith("| ") and "kg)" in l}
    micro = next(v for k, v in table.items() if k.startswith("Micro"))
    hatch = next(v for k, v in table.items() if k.startswith("Hatchback"))
    prem = next(v for k, v in table.items() if k.startswith("Premium"))
    sedan = next(v for k, v in table.items() if k.startswith("Sedan"))
    suv = next(v for k, v in table.items() if k.startswith("SUV"))
    assert "| 1.5 (n=2) |" in micro                  # 0 and 899 (900 is NOT micro: upper bound exclusive)
    assert "| 9.0 (n=3) |" in hatch                   # 900, 950 (Zed VXI, FE 20) and 1049: (3 + 4 + 20) / 3
    assert "| 5.5 (n=2) |" in prem                    # 1050 (lower bound inclusive) and 1199
    assert "| 7.5 (n=2) |" in sedan                   # 1200 and 1499
    assert "| 10.5 (n=4) |" in suv                    # 1500, 9999 and the >=10,000 kg fall-through to the last band
    assert "(0-900 kg)" in micro and "(1500-+ kg)" in suv


def test_diesel_missing_in_a_band_renders_a_dash(run_rows):
    _, md = run_rows([row(1, "Alpha VXI", 950, "Gasoline", 20.0), row(2, "Alpha CNG", 960, "CNG", 26.0)])
    hatch = next(l for l in md.splitlines() if l.startswith("| Hatchback (900"))
    assert hatch.endswith("| 20.0 (n=1) | — |")


def test_models_with_no_base_key_are_not_paired(run_rows):
    outs, _ = run_rows([row(1, "VXI", 950, "Gasoline", 20.0), row(2, "CNG", 960, "CNG", 26.0),
                        row(3, "Alpha VXI", 950, "Gasoline", 20.0), row(4, "Alpha CNG", 960, "CNG", 25.0)])
    pairs = list(csv.DictReader(open(outs["pairs"])))
    assert [p["model_key"] for p in pairs] == ["alpha"]


def test_pairs_sorted_by_uplift_and_mean_uplift_reported(run_rows):
    outs, md = run_rows([row(1, "Aa VXI", 950, "Gasoline", 20.0), row(2, "Aa CNG", 960, "CNG", 22.0),
                         row(3, "Bb VXI", 950, "Gasoline", 10.0), row(4, "Bb CNG", 960, "CNG", 15.0)])
    pairs = list(csv.DictReader(open(outs["pairs"])))
    assert [p["model_key"] for p in pairs] == ["bb", "aa"]                    # +50% before +10%
    assert [float(p["distance_uplift_pct"]) for p in pairs] == [50.0, 10.0]
    assert "declared **distance per unit rises 30.0% on CNG**" in md


def test_e20_table_rows(run_rows):
    _, md = run_rows([row(1, "Aa VXI", 950, "Gasoline", 20.0), row(2, "Aa CNG", 960, "CNG", 22.0)])
    assert "| 2% (SIAM low) | 19.6 | −0.40 |" in md
    assert "| 4% (central) | 19.2 | −0.80 |" in md
    assert "| 6% (ARAI high) | 18.8 | −1.20 |" in md
    assert "| Hatchback petrol (kmpl) | 20.3 | 20.0 | consistent |" in md


def test_fuel_summary_orders_by_model_count(run_rows):
    outs, _ = run_rows([row(1, "Aa VXI", 950, "Gasoline", 20.0), row(2, "Aa CNG", 960, "CNG", 22.0),
                        row(3, "Bb CNG", 960, "CNG", 24.0)])
    assert [r["fuel_type"] for r in csv.DictReader(open(outs["fuel"]))] == ["CNG", "Gasoline"]


# ------------------------------------------------------------------ recalc.py
def test_find_soffice_prefers_soffice_then_libreoffice(monkeypatch):
    monkeypatch.setattr(recalc.shutil, "which", lambda n: {"soffice": "/x/soffice", "libreoffice": "/x/lo"}.get(n))
    assert recalc.find_soffice() == "/x/soffice"
    monkeypatch.setattr(recalc.shutil, "which", lambda n: {"libreoffice": "/x/lo"}.get(n))
    assert recalc.find_soffice() == "/x/lo"
    monkeypatch.setattr(recalc.shutil, "which", lambda n: None)
    assert recalc.find_soffice() is None


def test_recalc_without_libreoffice_or_file(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(recalc, "find_soffice", lambda: None)
    assert recalc.recalc(str(tmp_path / "m.xlsx")) == 2
    assert "not found on PATH" in capsys.readouterr().out
    monkeypatch.setattr(recalc, "find_soffice", lambda: "/x/soffice")
    assert recalc.recalc(str(tmp_path / "missing.xlsx")) == 2
    out = json.loads(capsys.readouterr().out)
    assert out["error"].startswith("file not found")


def make_xlsx(path, cells):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for coord, v in cells.items():
        ws[coord] = v
    wb.save(path)


def fake_run_factory(produce, record=None):
    def run(cmd, env=None, capture_output=None, text=None, timeout=None):
        if record is not None:
            record.update(cmd=cmd, home=env["HOME"], timeout=timeout)
        outdir = cmd[cmd.index("--outdir") + 1]
        if produce:
            produce(outdir, cmd[-1])
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="x" * 600 + "TAIL")
    return run


def test_recalc_success_replaces_file_and_reports_no_errors(monkeypatch, tmp_path, capsys):
    target = tmp_path / "model.xlsx"
    make_xlsx(target, {"A1": 1})

    def produce(outdir, src):
        make_xlsx(os.path.join(outdir, "model.xlsx"), {"A1": 1, "A2": "fine"})
    rec = {}
    monkeypatch.setattr(recalc, "find_soffice", lambda: "/x/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", fake_run_factory(produce, rec))
    assert recalc.recalc(str(target)) == 0
    out = json.loads(capsys.readouterr().out)
    assert out == {"status": "success", "total_errors": 0, "error_summary": {}}
    assert openpyxl.load_workbook(target)["S"]["A2"].value == "fine"           # file replaced in place by the recalculated copy
    assert rec["cmd"][:5] == ["/x/soffice", "--headless", "--calc", "--convert-to", "xlsx"] and rec["timeout"] == 180
    assert rec["home"] != os.environ.get("HOME")                                # isolated LibreOffice profile


def test_recalc_reports_every_error_token(monkeypatch, tmp_path, capsys):
    target = tmp_path / "model.xlsx"
    make_xlsx(target, {"A1": 1})
    cells = {"A1": "#DIV/0!", "A2": "#DIV/0!", "B1": "#REF!", "C3": "#N/A", "D1": "#NAME?", "E1": "#NULL!",
             "F1": "#NUM!", "G1": "#VALUE!", "H1": "#ERROR!", "I1": "not #REF! exactly", "J1": 5}
    monkeypatch.setattr(recalc, "find_soffice", lambda: "/x/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", fake_run_factory(lambda d, s: make_xlsx(os.path.join(d, "model.xlsx"), cells)))
    assert recalc.recalc(str(target)) == 0                                       # errors in the sheet do not fail the run
    out = json.loads(capsys.readouterr().out)
    assert out["status"] == "errors_found" and out["total_errors"] == 9
    assert out["error_summary"]["#DIV/0!"] == ["S!A1", "S!A2"] and out["error_summary"]["#REF!"] == ["S!B1"]
    assert "not #REF! exactly" not in str(out)                                   # substring is not an error cell


def test_recalc_caps_each_error_list_at_fifty(monkeypatch, tmp_path, capsys):
    target = tmp_path / "model.xlsx"
    make_xlsx(target, {"A1": 1})
    cells = {f"A{i}": "#N/A" for i in range(1, 61)}
    monkeypatch.setattr(recalc, "find_soffice", lambda: "/x/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", fake_run_factory(lambda d, s: make_xlsx(os.path.join(d, "model.xlsx"), cells)))
    recalc.recalc(str(target))
    out = json.loads(capsys.readouterr().out)
    assert out["total_errors"] == 60 and len(out["error_summary"]["#N/A"]) == 50


def test_recalc_no_output_file_reports_stderr_tail(monkeypatch, tmp_path, capsys):
    target = tmp_path / "model.xlsx"
    make_xlsx(target, {"A1": 1})
    before = target.read_bytes()
    monkeypatch.setattr(recalc, "find_soffice", lambda: "/x/soffice")
    monkeypatch.setattr(recalc.subprocess, "run", fake_run_factory(None))
    assert recalc.recalc(str(target)) == 2
    out = json.loads(capsys.readouterr().out)
    assert out["error"] == "LibreOffice did not produce output" and len(out["stderr"]) == 500 and out["stderr"].endswith("TAIL")
    assert target.read_bytes() == before                                        # original untouched


# ------------------------------------------------------------------ extract_fe_pdf.main guard
def test_extract_main_exits_when_source_pdf_is_missing(tmp_path, monkeypatch):
    import extract_fe_pdf as m
    monkeypatch.setattr(m, "PDF", tmp_path / "nope.pdf")
    with pytest.raises(SystemExit) as e:
        m.main()
    assert "Source PDF not found" in str(e.value)

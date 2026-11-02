"""Tests for scripts/recalc.py helpers that need neither LibreOffice nor the network."""
import openpyxl
import recalc


def test_scan_errors_finds_error_tokens(tmp_path):
    p = tmp_path / "m.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "#DIV/0!"
    ws["B2"] = "ok"
    ws["C3"] = "#REF!"
    wb.save(p)
    assert recalc.scan_errors(str(p)) == {"#DIV/0!": ["S!A1"], "#REF!": ["S!C3"]}


def test_recalc_missing_file_returns_2(tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(recalc, "find_soffice", lambda: "/bin/true")
    assert recalc.recalc(str(tmp_path / "nope.xlsx")) == 2
    assert "file not found" in capsys.readouterr().out


def test_recalc_without_soffice_returns_2(monkeypatch, capsys):
    monkeypatch.setattr(recalc, "find_soffice", lambda: None)
    assert recalc.recalc("x.xlsx") == 2
    assert "not found" in capsys.readouterr().out

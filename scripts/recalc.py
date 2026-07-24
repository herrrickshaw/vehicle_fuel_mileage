#!/usr/bin/env python3
"""Recalculate an openpyxl-written .xlsx and report formula errors.

openpyxl writes formulas as strings with no cached values, so any tool reading
cached values (pandas, Excel previewers) sees blanks until the file is recalculated.
This script loads the workbook in headless LibreOffice — which computes every formula
that lacks a cached value on open — rewrites the file in place, then scans for error
strings and prints a JSON summary.

Requires LibreOffice (`soffice`) on PATH. This is optional: opening the workbook in
Excel / LibreOffice / Google Sheets recalculates it automatically.

Usage:  python scripts/recalc.py outputs/model.xlsx
Exit:   0 if recalculated (even with formula errors), non-zero if it could not run.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile

ERROR_TOKENS = ("#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#ERROR!")


def find_soffice():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None


def recalc(path):
    soffice = find_soffice()
    if not soffice:
        print(json.dumps({"error": "LibreOffice (soffice) not found on PATH"}))
        return 2
    path = os.path.abspath(path)
    if not os.path.exists(path):
        print(json.dumps({"error": f"file not found: {path}"}))
        return 2

    with tempfile.TemporaryDirectory() as tmp:
        # Isolate the LibreOffice user profile so concurrent/headless runs don't clash.
        env = dict(os.environ)
        env["HOME"] = tmp
        proc = subprocess.run(
            [soffice, "--headless", "--calc", "--convert-to", "xlsx",
             "--outdir", tmp, path],
            env=env, capture_output=True, text=True, timeout=180,
        )
        produced = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
        if not os.path.exists(produced):
            print(json.dumps({"error": "LibreOffice did not produce output",
                              "stderr": proc.stderr[-500:]}))
            return 2
        shutil.move(produced, path)

    # Scan the recalculated file for error strings.
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    errors = {}
    total = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v in ERROR_TOKENS:
                    total += 1
                    errors.setdefault(v, []).append(f"{ws.title}!{cell.coordinate}")
    print(json.dumps({
        "status": "errors_found" if total else "success",
        "total_errors": total,
        "error_summary": {k: v[:50] for k, v in errors.items()},
    }, indent=2))
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python scripts/recalc.py <file.xlsx>")
        sys.exit(2)
    sys.exit(recalc(sys.argv[1]))

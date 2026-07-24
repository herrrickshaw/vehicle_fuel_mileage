#!/usr/bin/env python3
"""Extract the SIAM BS-VI Fuel Efficiency Declaration (4W passenger vehicles,
as on 1 April 2020) from the source PDF into a clean CSV.

Source: data/raw/SIAM_BSVI_FE_Declaration_4W_2019-20.pdf
Output: data/siam_bsvi_fe_2020_4w.csv

Each data row in the PDF ends with a regular numeric tail:
    <kerb_kg> <engine_cc|-> "Bharat Stage VI" <FuelType> <declared_fe> <declared_co2>
so we anchor a regex on that tail and treat everything before it as
"manufacturer + model", then split off the manufacturer using a known-OEM
prefix list. Manufacturer/model split is best-effort; the fuel-type analysis
does not depend on it.
"""
import csv
import re
import sys
from pathlib import Path

import pdfplumber

ROOT = Path(__file__).resolve().parent.parent
PDF = ROOT / "data" / "raw" / "SIAM_BSVI_FE_Declaration_4W_2019-20.pdf"
OUT = ROOT / "data" / "siam_bsvi_fe_2020_4w.csv"

# Fuel labels as they appear (longest first so "Gasoline Hybrid" wins over "Gasoline").
FUELS = ["Gasoline Hybrid", "Gasoline", "Diesel", "CNG", "Electric"]
FUEL_ALT = "|".join(re.escape(f) for f in FUELS)

# Known OEM name prefixes (as spelled in the doc, incl. the "Renualt" typo).
OEMS = [
    "Maruti Suzuki India Limited", "Hyundai Motor India", "TATA Motors Limited",
    "Toyota Kirloskar Motor", "Honda Cars India Limited", "Mahindra & Mahindra",
    "Skoda Auto Volkswagen India Pvt. Ltd.", "MG Motor India Pvt. Ltd.",
    "Mercedes-Benz India", "BMW India", "Kia Motors", "Nissan Motor India",
    "Renualt India", "Renault India", "Ford India", "FCA India",
]

# Anchored on the numeric tail. Kerb & cc may carry thousands separators; cc may be "-".
# Thousands separators are stripped from the line first (see main), so kerb &
# cc are plain 3-4 digit integers here — which also stops a model name ending
# in a digit from being mis-split into the kerb column.
ROW_RE = re.compile(
    r"^(?P<sno>\d{1,3})(?P<name>[A-Za-z].+?)\s+"
    r"(?P<kerb>\d{3,4})\s+(?P<cc>\d{3,4}|-)\s+"
    r"Bharat Stage VI\s+(?P<fuel>" + FUEL_ALT + r")\s+"
    r"(?P<fe>\d+(?:\.\d+)?)\s+(?P<co2>\d+(?:\.\d+)?)\s*$"
)


def num(s):
    return s.replace(",", "").strip()


def split_oem(name):
    for oem in OEMS:
        if name.startswith(oem):
            model = name[len(oem):].strip()
            canonical = "Renault India" if oem.startswith("Renual") else oem
            return canonical, model
    # Fallback: first two tokens as manufacturer.
    toks = name.split()
    return " ".join(toks[:2]), " ".join(toks[2:])


def main():
    if not PDF.exists():
        sys.exit(f"Source PDF not found: {PDF}")

    rows = []
    seen_sno = set()
    with pdfplumber.open(PDF) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                # Strip thousands separators so kerb/cc parse cleanly. The PDF
                # renders them irregularly ("1 ,000", "1,000", "1, 000"), so
                # collapse digit + optional-space + comma + optional-space + 3
                # digits into one integer. FE/CO2 use decimal points, not commas.
                clean = re.sub(r"(\d)\s*,\s*(\d{3})", r"\1\2", line.strip())
                m = ROW_RE.match(clean)
                if not m:
                    continue
                sno = int(m["sno"])
                # The header repeats "S.no." etc. per page; ROW_RE won't match those.
                if sno in seen_sno:
                    continue
                seen_sno.add(sno)
                oem, model = split_oem(m["name"].strip())
                rows.append({
                    "sno": sno,
                    "manufacturer": oem,
                    "model_variant": model,
                    "kerb_weight_kg": int(num(m["kerb"])),
                    "engine_cc": (None if m["cc"] == "-" else int(num(m["cc"]))),
                    "emission_stage": "BS-VI",
                    "fuel_type": m["fuel"],
                    "declared_fe": float(m["fe"]),   # kmpl (petrol/diesel), km/kg (CNG), km/kWh-eq (electric)
                    "declared_co2_gkm": float(m["co2"]),
                })

    rows.sort(key=lambda r: r["sno"])
    with OUT.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    # Sanity report to stderr.
    from collections import Counter
    by_fuel = Counter(r["fuel_type"] for r in rows)
    print(f"Extracted {len(rows)} rows -> {OUT.relative_to(ROOT)}", file=sys.stderr)
    print("By fuel:", dict(by_fuel), file=sys.stderr)
    gaps = [n for n in range(1, max(seen_sno) + 1) if n not in seen_sno]
    if gaps:
        print(f"WARNING: missing S.no rows: {gaps}", file=sys.stderr)
    else:
        print(f"Contiguous S.no 1..{max(seen_sno)} — no gaps.", file=sys.stderr)


if __name__ == "__main__":
    main()

# Which Gives More Bang for Your Buck? — India Vehicle Fuel-Cost Analysis

A reproducible cost-of-ownership and fiscal analysis comparing **ethanol-blended petrol (E20), diesel, CNG and electric** vehicles in India (July 2026 basis), and the tax revenue they generate for the exchequer.

Everything is generated from two scripts: a Python model builder (Excel) and a Node report builder (Word). No manual spreadsheet editing.

## What's inside

| Output | Built by | Description |
|---|---|---|
| `outputs/EV_CNG_Diesel_Ethanol_Vehicle_Cost_Model.xlsx` | `scripts/build_model.py` | 11-sheet, fully formula-driven cost model with 21 automated consistency checks |
| `outputs/Bang_For_Your_Buck_Vehicle_Study.docx` | `scripts/build_report.js` | 14-section research paper (abstract + tables) built from the same data |
| `data/siam_bsvi_fe_2020_4w.csv` + `docs/fe_declaration_analysis.md` | `scripts/extract_fe_pdf.py` → `scripts/analyze_fe_declarations.py` | 303 real ARAI-certified 4W fuel-efficiency declarations extracted from the SIAM PDF, with fuel-type / weight-band / petrol-vs-CNG analysis calibrating the model's mileage assumptions |

### Vehicle-efficiency evidence base (SIAM declarations)

`scripts/extract_fe_pdf.py` parses the SIAM BS-VI Fuel-Efficiency Declaration PDF (`data/raw/`, passenger vehicles as on 1 Apr 2020) into a clean 303-row CSV, and `scripts/analyze_fe_declarations.py` derives:

- **Declared FE by fuel type** — petrol 16.7 kmpl mean, diesel 17.9, CNG 27.4 km/kg, hybrid 21.9, over 303 models.
- **Diesel's structural lead** — ~30-34% more kmpl than petrol *within every kerb-weight band* (not just overall).
- **Same-nameplate petrol→CNG uplift** — +40% mean distance per unit across the 11 models sold in both forms (WagonR +54%, Alto +43%), the physical basis for CNG's sub-2-year payback.
- **Calibration cross-check** — recomputed averages match the constants cited in `data/sources.md` (CNG 27.4 km/kg exactly), and quantify the E20 mileage haircut on the mean petrol vehicle.

### The Excel model (sheets)

1. **README / Assumptions** — pick a city, edit levers (E20 drop %, usage km, charging mix); the whole model reprices.
2. **Cities** — petrol/diesel/CNG/electricity + state VAT for 6 metros, with cost-per-km.
3. **CostPerKm** — running cost per km by vehicle class × powertrain (petrol includes the E20 penalty).
4. **EthanolImpact** — what E10/E20 blending costs a petrol owner.
5. **TCO_Scenarios** — fuel + maintenance + premium and payback across Low/Average/High usage.
6. **TaxStructure** — why petrol/diesel are taxed by centre *and* state, while ethanol carries only 5% GST.
7. **FleetSales** — FADA FY26 sales mix, sales-weighted running cost, national ethanol-penalty aggregate, SIAM-vs-FADA reconciliation.
8. **OMC_PriceBoard** — 14-city OMC-published fuel prices and the national spread.
9. **ExchequerRevenue** — fuel excise+VAT + vehicle GST, additional earnings from growth, EV/CNG/ethanol erosion.
10. **Checks** — 21 automated data-consistency checks (all must read PASS).

### Companion study: The Volume Dividend (2026-07-28)

`docs/Energy_Blend_Volume_Dividend.docx/.pdf` + `docs/Volume_Dividend_Slides.pptx` — an
energy-density and fiscal follow-up built in the companion repo
[`india-omc-fuel-fleet-model`](https://github.com/herrrickshaw/india-omc-fuel-fleet-model),
with vehicle mileage anchored on **this repo's SIAM FE declarations** (petrol 16.67 kmpl,
CNG 27.40 km/kg, +40.3% same-nameplate petrol→CNG uplift):

- **The volume effect**: E20's −4% mileage pulls +2.16 bn extra litres through pumps for the same
  km — ₹22,703 cr/yr of extra consumer spend at an unchanged pump price, collected via per-litre
  levies (excise ₹4,303 cr, VAT ₹4,216 cr, dealer commission ₹886 cr, OMC margin ₹757 cr).
- **Parity pricing**: honest cost-per-km needs E20 at ₹100.80/L (E30 ₹97.65). The dual-tax exemption
  this repo's TaxStructure sheet documents (ethanol pays only 5% GST) already funds that discount
  1.3–1.5× — it is currently retained in the price build-up, not rebated.
- **The E27 grand bargain**: parity price ₹98.44 + a 5% state SGST on ethanol (₹4,787 cr/yr for
  states) still fits inside the embedded headroom with ₹1.27/L to spare.
- **CBG contrast**: sold per kg at CNG-parity energy — renewable MJ at ₹1.16 vs ethanol's ₹2.94
  (2.5× cheaper), zero mileage loss, zero hidden levy.
- **Supply check (CareEdge May-2026, DFPD register, NCDC coop scheme)**: the higher blends *cure*
  the distillery overcapacity — E20 runs FY27 capacity at 59%, E27 at 76% (top of the consolidation
  band) with zero new construction; the coop-mill scheme is 96.5% working capital (~9.7 cr L/yr of
  actual ethanol capacity); feedstock (marginal maize), not steel, is the binding constraint.
- **RON95 dividend**: ethanol's blending RON ~112 means E20 on an unchanged 91-RON blendstock is a
  free RON95 fuel; compression-ratio headroom recovers +2.2% efficiency → net E20 mileage penalty
  ≈ −1.8% on a RON95-calibrated engine — the Brazil E27 playbook, arriving with fleet turnover.

## Headline findings (Delhi basis, July 2026)

- **Running cost/km:** EV cheapest (₹1.41 hatchback, ₹0.35 scooter) < CNG (₹2.97) < diesel (₹4.33–5.01 cars) < E20 petrol (₹6.65–8.18). Ethanol penalty ≈ ₹4,000/yr per car.
- **Payback:** CNG < 2 yrs for average drivers; EV ~2.4–3.5 yrs at high mileage; diesel only at high mileage.
- **City spread:** petrol ₹101.5 (Chandigarh) → ₹115.7 (Hyderabad), ~14% — almost entirely state VAT.
- **Market check:** FADA FY26 shows CNG at 22% of PV sales and 8.5% overall EV penetration — buyers follow the value math.
- **Exchequer:** ~₹4.28 lakh cr/yr recurring fuel tax + ~₹1.66 lakh cr one-time vehicle GST; growth adds ~₹44,500 cr/yr, but EV/CNG/ethanol erode the base (~₹32,000 cr/yr foregone on ethanol alone).

## Reproduce

```bash
# Python deps
pip install -r requirements.txt

# Node deps
npm install

# Build the Excel model (then recalculate formulas with LibreOffice)
python scripts/build_model.py
python scripts/recalc.py outputs/EV_CNG_Diesel_Ethanol_Vehicle_Cost_Model.xlsx   # requires libreoffice/soffice

# Build the Word report
node scripts/build_report.js
```

Or use the Makefile: `make all`.

> **Note:** `recalc.py` uses LibreOffice (`soffice`) to compute the formula values openpyxl writes as strings. Without it, formula cells read back as empty. Install LibreOffice to run the checks locally.

## Data sources

See [`data/sources.md`](data/sources.md) for full provenance. Headline sources: ARAI/SIAM (fuel efficiency & E20 drop), PPAC/OMC (fuel prices & consumption), FADA (retail sales), SIAM (wholesale), ClearTax/PIB (taxation), GST 2.0 rates.

## Caveats

Representative estimates for comparison, not price quotes. State VAT rates are indicative effective rates; several states add fixed cesses. Fleet-level and exchequer aggregates use editable assumptions documented in the model. Verify current local prices before any purchase decision.

## License

MIT — see [LICENSE](LICENSE).

---
*Generated with [Claude](https://claude.com/claude-code). Figures are analytical estimates, not financial advice.*

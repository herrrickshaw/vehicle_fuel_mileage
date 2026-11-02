# India Vehicle Fuel-Cost Analysis (E20 petrol vs diesel vs CNG vs EV)

A reproducible cost-of-ownership and fiscal analysis comparing ethanol-blended petrol (E20), diesel, CNG and
electric vehicles in India (July 2026 basis), plus the tax revenue each generates for the exchequer. It is aimed
at analysts and policy readers and produces an 11-sheet Excel model, a Word research paper, and a
fuel-efficiency evidence base extracted from the SIAM BS-VI declarations.

> Status: maintained. Figures are analytical estimates from public sources (see Data sources), not price
> quotes or financial advice. New to AI-assisted use? See [HOW_TO_USE_WITH_AI.md](HOW_TO_USE_WITH_AI.md).

## What it does
- Builds a fully formula-driven Excel cost model (cost/km, TCO and payback, tax structure, fleet mix, exchequer
  revenue) with 21 automated consistency checks (`scripts/build_model.py`).
- Recalculates the workbook headlessly with LibreOffice and reports formula errors as JSON (`scripts/recalc.py`).
- Builds a 14-section Word paper from the same numbers (`scripts/build_report.js`).
- Parses the SIAM BS-VI Fuel-Efficiency Declaration into a 303-row CSV of ARAI-certified 4W declarations
  (`scripts/extract_fe_pdf.py`).
- Analyses the declarations by fuel type, weight band and petrol-vs-CNG nameplate pair to calibrate the model's
  mileage assumptions (`scripts/analyze_fe_declarations.py`).

## Quick start
```bash
git clone https://github.com/herrrickshaw/vehicle_fuel_mileage.git
cd vehicle_fuel_mileage
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt && npm install
python scripts/build_model.py                                                    # writes the .xlsx
python scripts/recalc.py outputs/EV_CNG_Diesel_Ethanol_Vehicle_Cost_Model.xlsx   # needs LibreOffice (soffice)
node scripts/build_report.js                                                     # writes the .docx
```
`make all` runs the same steps plus the FE-declaration extraction (see Limitations).

## How it works
| Stage | Module | Output |
|---|---|---|
| FE declarations to CSV | `scripts/extract_fe_pdf.py` (`parse_line`) | `data/siam_bsvi_fe_2020_4w.csv` |
| FE analysis | `scripts/analyze_fe_declarations.py` (stdlib only) | `data/fe_by_fuel_summary.csv`, `data/fe_petrol_vs_cng_pairs.csv`, `docs/fe_declaration_analysis.md` |
| Cost model | `scripts/build_model.py` (openpyxl) | `outputs/EV_CNG_Diesel_Ethanol_Vehicle_Cost_Model.xlsx` |
| Recalculation | `scripts/recalc.py` (LibreOffice headless) | cached formula values, JSON error report |
| Report | `scripts/build_report.js` (`docx`) | `outputs/Bang_For_Your_Buck_Vehicle_Study.docx` |

Design decisions: every model value is an Excel **formula** referencing assumption cells (never a pasted result),
so changing a lever (city, E20 drop %, usage km, charging mix) reprices the whole workbook; cell addresses are
tracked in Python dicts so cross-sheet formulas survive row moves. Numbers in the Word report are transcribed
from the recalculated model and must be kept in sync by hand. See [CLAUDE.md](CLAUDE.md) for the working rules.

### The Excel model (sheets)
1. **README / Assumptions** - pick a city, edit levers; the whole model reprices.
2. **Cities** - petrol/diesel/CNG/electricity + state VAT for 6 metros, with cost-per-km.
3. **CostPerKm** - running cost per km by vehicle class and powertrain (petrol includes the E20 penalty).
4. **EthanolImpact** - what E10/E20 blending costs a petrol owner.
5. **TCO_Scenarios** - fuel + maintenance + premium and payback across Low/Average/High usage.
6. **TaxStructure** - why petrol/diesel are taxed by centre and state, while ethanol carries only 5% GST.
7. **FleetSales** - FADA FY26 sales mix, sales-weighted running cost, national ethanol-penalty aggregate, SIAM-vs-FADA reconciliation.
8. **OMC_PriceBoard** - 14-city OMC-published fuel prices and the national spread.
9. **ExchequerRevenue** - fuel excise+VAT + vehicle GST, additional earnings from growth, EV/CNG/ethanol erosion.
10. **Checks** - 21 automated data-consistency checks (all must read PASS).

### Vehicle-efficiency evidence base (SIAM declarations)
- Declared FE by fuel type: petrol 16.7 kmpl mean, diesel 17.9, CNG 27.4 km/kg, hybrid 21.9, over 303 models.
- Diesel's structural lead: about 30-34% more kmpl than petrol within every kerb-weight band.
- Same-nameplate petrol to CNG uplift: +40% mean distance per unit across the 11 models sold in both forms
  (WagonR +54%, Alto +43%), the physical basis for CNG's sub-2-year payback.
- Calibration cross-check: recomputed averages match the constants cited in `data/sources.md` (CNG 27.4 km/kg exactly).

## Headline findings (Delhi basis, July 2026)

- **Running cost/km:** EV cheapest (₹1.41 hatchback, ₹0.35 scooter) < CNG (₹2.97) < diesel (₹4.33–5.01 cars) < E20 petrol (₹6.65–8.18). Ethanol penalty ≈ ₹4,000/yr per car.
- **Payback:** CNG < 2 yrs for average drivers; EV ~2.4–3.5 yrs at high mileage; diesel only at high mileage.
- **City spread:** petrol ₹101.5 (Chandigarh) → ₹115.7 (Hyderabad), ~14% — almost entirely state VAT.
- **Market check:** FADA FY26 shows CNG at 22% of PV sales and 8.5% overall EV penetration — buyers follow the value math.
- **Exchequer:** ~₹4.28 lakh cr/yr recurring fuel tax + ~₹1.66 lakh cr one-time vehicle GST; growth adds ~₹44,500 cr/yr, but EV/CNG/ethanol erode the base (~₹32,000 cr/yr foregone on ethanol alone).

## Companion study: The Volume Dividend (2026-07-28)

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
- **Annex A — CBG incentive stack**: 19 instruments by lever from the 70-circular GOBARdhan
  register (SATAT/CBO offtake, capital schemes, DPI pipeline, FCO Schedule VIII fertiliser leg +
  MDA, priority-sector lending, GST/excise/carbon, IS 16087, state policies) — every leg of the
  CBG business de-risked on-budget, the structural opposite of ethanol's consumer-funded dividend.
- **CBG investment pitch** (`docs/CBG_Investment_Pitch.pptx/.pdf`): 9-slide deck on a 12-TPD
  reference-plant model — mandated CBO demand, 8-lever policy moat, three revenue legs, project
  IRR 20.6% / equity IRR 32% / DSCR 1.85×, and a ₹150 cr 10-plant platform ask anchored on
  cooperative-sugar-mill press-mud.
- **Multi-feed distillery conversion pitch** (`docs/Distillery_MultiFeed_Pitch.pptx/.pdf`):
  ₹50 cr converts a 100-KLPD molasses distillery to 330-day dual-feed — slab arbitrage (maize
  ₹71.86 vs C-heavy ₹57.97), project IRR 38.4% / equity 81% on DFPD-subvented debt, with the
  damaged-grain-slab stress (17%) as the honest floor; 5-mill ₹250 cr CSM programme ask.
- **DME in LPG — the fourth quadrant** (paper §10 + slide 13, LERC basis): DME's −37%/kg dilution
  on a *subsidised* per-kg fuel inverts the volume dividend — DME20 costs households +₹312/yr AND
  adds ₹1,201 cr/yr of PMUY subsidy; methanol-route DME fails the ₹32.7/kg energy-parity ceiling,
  so guardrails are parity cylinder pricing (₹761 vs ₹803) + domestic-carbon DME only.
- **DME blending pitch** (`docs/DME_Blending_Pitch.pptx/.pdf`): the guardrails made investable —
  100-TPD dehydration plant selling at energy parity into unsubsidised commercial LPG; IRR 26% /
  equity 42% on domestic methanol, with the kg-parity "temptation margin" (₹34.7/kg) shown and
  refused; gated on a ≤₹17/kg take-or-pay methanol contract.
- **Bio-isobutanol pitch** (`docs/Isobutanol_Blending_Pitch.pptx/.pdf`): the molecule that halves
  the mileage penalty — IB20 dilutes 3.5% vs E20's 6.9%, and yields +25% more jet fuel per tonne
  via ATJ. Blocked by having no price slab and no BIS standard (at ethanol's per-litre slab the
  margin is −₹9.58/L); with energy-basis pricing, blending-only IRR 13.8% → 33.9% at 30% SAF.

## Repository layout
| Path | Purpose |
|---|---|
| `scripts/` | Model builder, recalc, report builder, FE extraction and analysis |
| `tests/` | pytest suite (cost-per-km formulas, FE analysis, row parser, recalc helpers) |
| `data/` | Tracked FE declaration CSVs and `sources.md` provenance; `data/raw/` is not tracked |
| `outputs/` | Generated Excel model and Word report (regenerate, never hand-edit) |
| `docs/` | Generated FE analysis plus companion paper and pitch decks |

## Data sources
Full provenance in [`data/sources.md`](data/sources.md). Headline sources: ARAI/SIAM (fuel efficiency and E20
drop), PPAC/OMC (fuel prices and consumption), FADA (retail sales), SIAM (wholesale), ClearTax/PIB (taxation),
GST 2.0 rates. City prices were compiled from third-party aggregators (Goodreturns, HDFCSky); treat as snapshots.

## Configuration
No environment variables or secrets are needed. `requirements-ocr.txt` is an opt-in OCR fallback client
(active only if `UNLIMITED_OCR_URL` is set). Model levers live in the workbook's Assumptions sheet.

## Testing
```bash
pip install -r requirements-dev.txt ruff
ruff check .
pytest -q
```

## Limitations and known issues
- State VAT rates are indicative effective rates; several states add fixed cesses. Fleet-level and exchequer
  aggregates use editable assumptions documented in the model. Verify current local prices before any purchase.
- The source SIAM PDF (`data/raw/SIAM_BSVI_FE_Declaration_4W_2019-20.pdf`) was removed from the repo, so
  `make data` / `extract_fe_pdf.py` exits with "Source PDF not found" unless you supply the PDF; the extracted
  CSV and analysis outputs are tracked and `analyze_fe_declarations.py` runs without it.
- `recalc.py` needs LibreOffice; without it, formula cells read back as empty until opened in a spreadsheet app.
- `build_model.py` executes at import time and writes to a relative `outputs/` path; run it from the repo root.

## Related repositories
- [`india-omc-fuel-fleet-model`](https://github.com/herrrickshaw/india-omc-fuel-fleet-model) - companion
  energy-density and fiscal study that anchors mileage on this repo's SIAM declarations.

## Licence
MIT - see [LICENSE](LICENSE).

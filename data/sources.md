# Data sources & provenance

All figures are July 2026 basis unless noted. Prices vary by city and day; treat as a snapshot.

## Fuel efficiency & ethanol impact
- **ARAI E20 study (Jul 2026)** — 2–6% fuel-efficiency drop on E20; no engine failure in compatible vehicles. (Business Today, The Tribune, ANI)
- **SIAM** — 2–4% controlled-environment drop; rejected "50% drop" claims. (Deccan Herald)
- **Autocar India** — real-world E20 tests, up to ~12% on older non-compliant vehicles.
- **SIAM BS-VI Fuel Efficiency Declarations** (4W 2019-20, 2W 2019, ARAI-certified) — used to calibrate the model's real-world mileage assumptions. Certified averages: hatchback petrol 20.3 kmpl, sedan 18.4, compact SUV 14.8, CNG car 27.4 km/kg, two-wheeler 50.7 kmpl.
  - The **4W passenger-vehicle declaration (as on 1 April 2020, 303 models)** is now included as primary data: PDF at `data/raw/SIAM_BSVI_FE_Declaration_4W_2019-20.pdf`, extracted by `scripts/extract_fe_pdf.py` to `data/siam_bsvi_fe_2020_4w.csv`, analysed by `scripts/analyze_fe_declarations.py` (see [`docs/fe_declaration_analysis.md`](../docs/fe_declaration_analysis.md)). Recomputed averages confirm the cited constants (CNG 27.4 km/kg exactly; hatchback petrol ~19-21 kmpl), and quantify the same-nameplate petrol→CNG distance uplift (mean +40%) and diesel's ~30% per-band lead.

## Fuel & electricity prices
- **OMC daily boards** (IndianOil, BPCL, HPCL) via PPAC, compiled by Goodreturns / HDFCSky — city-wise petrol/diesel; CNG city prices (Goodreturns).
- Home electricity ₹8–11/kWh (indicative domestic slabs); public DC fast charging ₹18–24/kWh (2026).

## Running costs & vehicles
- **VahanBazaar, AutoPunditz, BikeDekho (2026)** — running costs, premiums, CNG bike (Bajaj Freedom 125, 102 km/kg).
- **GoMechanic** — biofuels-vs-EVs (operating cost ₹4–5/km biofuel vs ₹1.1/km EV; 30–50% CO₂ cut).

## Sales & fleet
- **FADA FY26 retail data** (fada.in) — category volumes, PV fuel mix; overall EV penetration 8.5% (JMK Research).
- **SIAM FY26 wholesale** (press release / Outlook Business) — dispatch volumes; SIAM-vs-FADA methodology (Business Standard).

## Taxation & exchequer
- **ClearTax / Business Standard** — central excise (₹13/L petrol, ₹10/L diesel), state VAT rates, ethanol 5% GST.
- **GST 2.0** (effective Sep 2025, ClearTax) — small cars/2W 18%, large cars/SUV 40%, EV 5%.
- **PPAC** — FY26 consumption (petrol 42.66 MMT, diesel 93.95 MMT).
- **PPAC/PIB via Business Standard** — petroleum-sector tax contribution (~₹7.5 lakh crore; centre/state split).

## Key modelling assumptions (editable in the model)
- E20 mileage drop: 4% central (ARAI/SIAM range 2–6%).
- EV charging mix: 80% home / 20% public.
- Diesel modelled only for sedan & compact SUV.
- National-average state VAT (₹18/L petrol, ₹12/L diesel) for the exchequer aggregate; on-road parc counts illustrative.

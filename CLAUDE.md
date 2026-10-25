# CLAUDE.md — guidance for Claude Code working in this repo

## What this project is

A reproducible analysis comparing India's vehicle powertrains (E20 petrol, diesel, CNG, EV) on
cost-per-km, total cost of ownership, payback, city price variation, market adoption, and the tax
revenue they generate. Two generator scripts produce two artifacts; **never hand-edit the outputs**
— change the scripts and regenerate.

## Architecture

- `scripts/build_model.py` — builds the Excel model with `openpyxl`. Every value is a **formula**
  referencing assumption cells (never a hardcoded computed result), so the workbook recalculates
  when inputs change. Cell references are tracked in Python dicts (`A`, `MILE`, `PRICE`, `CPK`,
  `EFF`, `EXC`, …) so cross-sheet formulas stay correct when rows move.
- `scripts/recalc.py` — runs LibreOffice headless to compute the formula values (openpyxl writes
  formulas as strings with no cached values). Reports formula errors as JSON. **Always run after
  building** and confirm `status: success, total_errors: 0`.
- `scripts/build_report.js` — builds the Word report with `docx` (npm). Numbers in the report are
  transcribed from the recalculated model; keep them in sync when the model changes.

## Golden rules (learned the hard way)

1. **Formulas, not hardcoded numbers.** `sheet['B10'] = '=SUM(...)'`, never the Python total.
2. **Recalc after every model change.** A green recalc proves formulas *evaluate*, not that ranges
   are *right* — spot-check a few values against hand-calculation.
3. **Consistency checks are load-bearing.** The `Checks` sheet has 21 checks; all must read PASS.
   When you add data, add a check (e.g. shares sum to 100%, A within X% of B). A check that
   references a hardcoded cell address breaks when rows shift — reference via the tracked dicts.
4. **Sample-first, then scale.** Validate a new mechanic on one vehicle/city, recalc, verify the
   numbers, then extend to the full matrix.
5. **Unit discipline.** Money in ₹ crore / lakh crore; volumes in crore litres; 1 MMT = 100/density
   crore L; 1 lakh crore = 1e12 ₹. Unit-scale bugs are the most common error here — check magnitudes.
6. **Cite every hardcoded input** in an adjacent note cell, and keep `data/sources.md` current.

## Regenerate

```bash
pip install -r requirements.txt && npm install
python scripts/build_model.py
python scripts/recalc.py outputs/EV_CNG_Diesel_Ethanol_Vehicle_Cost_Model.xlsx
node scripts/build_report.js
```

## Good next-analysis directions

- Replace the illustrative on-road parc assumptions (FleetSales §D) with actual Vahan parc-by-fuel
  counts, and the national-average VAT (ExchequerRevenue) with state-wise VAT × state consumption.
- Add a time series (FY22–FY26) to trend consumption, sales mix and exchequer revenue.
- Add sensitivity/tornado analysis on the payback conclusions (crude price, electricity tariff,
  E20 drop, resale value).
- Add charts (matplotlib/openpyxl charts) to the model and figures to the report.
- Model resale value, financing and battery-replacement risk in the 5-year TCO.
- Wire in a live price feed (PPAC / OMC daily boards) to refresh the Cities and OMC_PriceBoard tabs.

## Conventions

- Excel colours: blue = input, yellow = key lever, green = cross-sheet link, black = formula.
- Keep sheet names without awkward spaces; cross-sheet refs to names with spaces must be quoted.
- Prefer Excel-2007-era functions (INDEX/MATCH/SUMPRODUCT/IFERROR); avoid XLOOKUP/FILTER/UNIQUE
  (LibreOffice recalc can't evaluate them from openpyxl-written files).

## Cross-session project memory

A Claude Code session rooted here has no persisted memory of its own. Decisions, hazards, and
credential locations for the broader market-research platform (of which this cost model is one
project) live under the `~` (home) session: `~/.claude/projects/-Users-umashankar/memory/`,
indexed by `MEMORY.md` and mapped by the `reference_memory_map` memory. From here, reach it with
`qmd search "<topic>"` (indexes that memory directory regardless of working directory) rather than
assuming this session has its own history. For fast orientation, a published snapshot of 128
incident-sourced rules from across that memory is browsable at
https://claude.ai/code/artifact/e29ceaeb-099e-4de7-ab43-dcd97a2d8f57 ("Ways of Working") — a
starting point, not authoritative over the live memory files.

<!-- source-fetch-toolkit:v1 (managed by ~/bin/add_source_fetch_section.py — edit there, not here) -->
## Source fetching for data collection & literature surveys

When a task needs a primary document (statute, regulator data, government release, paper, filing),
fetch the document itself — never cite a search snippet as if it had been read.

- **On the owner's Mac:** `fetch_source.py "<url | doi:… | celex:… | arxiv:…>" --out ./sources`
  (in `~/bin`; full guide `~/bin/SOURCE_FETCH_TOOLKIT.md`). Scheduled/unattended runs use only
  `bash /Users/umashankar/bin/fetch_source.sh "<target>"` (allowlisted) and then Read the `.txt` it reports.
- **Elsewhere (cloud sessions, CI):** follow the same ladder by hand — direct GET with a browser User-Agent →
  headless browser only if the page needs JavaScript → official alternates (DOI → OpenAlex
  `api.openalex.org/works/doi:<doi>` locations → repository copy; EUR-Lex → EU Publications Office
  `publications.europa.eu/resource/celex/<CELEX>`; arXiv pdf/abs) → Wayback Machine `id_` snapshot.
- **Hard rule:** a CAPTCHA, robot check, Cloudflare/Akamai/WAF challenge or login wall means stop on that
  host. Do not bypass it (no TLS impersonation, stealth browsers, identity rotation or challenge solving);
  look for the same document via another route, or report it as unavailable.
- **Record** for every figure: URL read, title/date, page/table, fetch route; mark abstract-only or
  snippet-only sources. Fetch politely (~1 request/second per host, no bulk crawling).
<!-- /source-fetch-toolkit -->

# GDP tempo-effect PoC (candidates A / B / D)

This directory contains a proof-of-concept testing whether the mechanisms that
matter for period-versus-cohort fertility indicators (tempo shift + forgotten
dispersion parameter) have analogues that improve GDP models.

Three candidates are implemented in parallel and compared on the same
39-country benchmark (OECD + China, D.R. Congo dropped for insufficient data):

| Candidate | Analogue to population paper | Mechanism |
|---|---|---|
| **A** | AFB shift (tempo) | Investment-to-output time-to-build lag μ(t) drifts |
| **B** | Same-living population | Mean labour entry/exit ages shift → effective L |
| **D** | σ (forgotten parameter) | Intangible capital as additional factor |

See [`reports/poc_findings.md`](reports/poc_findings.md) for the full writeup
including per-candidate results, the A/B/D comparison, and discussion of the
implications for Beyond-GDP indicators (IWI, CWON, HDI).

## Reproduce

1. Download Penn World Table 10.01 (`.dta`) to
   `/home/ubuntu/gdp_tempo_data/pwt1001.dta`:
   ```bash
   mkdir -p /home/ubuntu/gdp_tempo_data
   curl -sL "https://dataverse.nl/api/access/datafile/354098" \
     -o /home/ubuntu/gdp_tempo_data/pwt1001.dta
   ```
2. Pull World Bank data for candidates B and D (5-year age population bins,
   R&D-to-GDP share). Script `scripts/fetch_wb.py` may be added if not present;
   for now the API calls used in this PoC are documented in
   `reports/poc_findings.md` §8.
3. Run the PoCs:
   ```bash
   cd gdp_tempo_poc
   python scripts/run_poc.py         # Candidate A   (~1 min)
   python scripts/run_poc_B.py       # Candidate B   (~10 min)
   python scripts/run_poc_D.py       # Candidate D   (~30 s)
   python scripts/compare_ABD.py     # A/B/D comparison
   python scripts/run_champion_plots.py  # A time-series overlays
   ```

## Files

- `scripts/run_poc.py` — Candidate A: three PIM constructions × three tests.
- `scripts/run_poc_B.py` — Candidate B: age-profile labour with tempo drift.
- `scripts/run_poc_D.py` — Candidate D: Cobb-Douglas with intangible K.
- `scripts/compare_ABD.py` — cross-candidate comparison figures and summary.
- `scripts/run_champion_plots.py` — time-series overlay for top-6 A-countries.
- `data/poc_results.csv` + `_B` + `_D` — per-country results for each PoC.
- `data/compare_ABD.csv` — merged comparison table.
- `data/*_summary.json` — aggregate statistics per candidate.
- `figures/fig1–5*.png` — Candidate A plots (growth RMSE, improvements,
  K direct, μ₁ scatter, champion countries).
- `figures/figB1–3*.png` — Candidate B plots.
- `figures/figD1–3*.png` — Candidate D plots.
- `figures/figCompare1–2*.png` — A vs B vs D comparison.
- `reports/poc_findings.md` — full writeup with Beyond-GDP discussion.

## Headline results

| Candidate | Test B (growth-RMSE, pp) median gain | Test A (levels-MAPE, pp) median gain | Countries improved |
|---|---|---|---|
| A | **+0.028** (74% of countries) | ≈0 | 24 / 39 |
| B | ≈0 (18%) | ≈0 (38%) | 2 / 39 |
| D | +0.0003 (54%) | **+0.39** (97%) | 13 / 39 |

**Key finding**: A and D operate on different timescales — A tightens
year-to-year growth dynamics; D tightens long-run level trends. Together they
form a natural two-component generalisation of the population paper's
"tempo + forgotten parameter" structure. Candidate B does not help at annual
frequency because PWT's observed `emp` already encodes business-cycle signal.

# GDP × CWON Joint Identification

Empirical demonstrator for the flow-stock unification framework of
[`gdp_tempo_poc/reports/poc_findings.md` §9](../gdp_tempo_poc/reports/poc_findings.md).

Given per-country estimates of the hidden tempo parameter μ (candidate A)
and intangible share β (candidate D) from the GDP PoC, this subproject
reconciles the PIM-reconstructed capital trajectories against World Bank
CWON (Changing Wealth of Nations) series `NW.PCA.TO` / `NW.HCA.TO` /
`NW.TOW.TO`, and re-fits (μ, β) jointly under a combined production
+ wealth loss.

## Results in one line

```
trajectory RMSE median (PIM vs CWON produced):    0.049 log units (~5% deviation)
β median, production-only fit:                    0.01
β median, joint (production + wealth) fit:        0.06   ← wealth constraint raises β
```

See [`reports/cwon_findings.md`](reports/cwon_findings.md) for full results
and honest caveats (PPP vs market-exchange-rate confound in level-cross-country
comparison).

## Reproduce

```bash
cd gdp_cwon_integration
python scripts/run_cwon_integration.py
```

Inputs (already in place from `gdp_tempo_poc`):
- `gdp_tempo_poc/data/poc_results.csv` (μ̂ per country from candidate A)
- `gdp_tempo_poc/data/poc_D_results.csv` (β̂ per country from candidate D)
- `/home/ubuntu/gdp_tempo_data/pwt1001.dta`
- `/home/ubuntu/gdp_tempo_data/wb/rnd_gdp.json`
- `/home/ubuntu/gdp_tempo_data/wb/cwon/NW.{PCA,HCA,TOW,TOW.PC}.TO.json`

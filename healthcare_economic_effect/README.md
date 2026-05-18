# Healthcare Expenditure as Economic Effect

A neutral sustainability framework that reframes healthcare spending from
"cost to be contained" to "investment with measurable economic returns."

## Core Thesis

Healthcare expenditure is simultaneously a **cost** and an **economic effect**.
This project integrates five analytical lenses:

1. **Input-Output (I-O) Multipliers** -- healthcare spending as demand stimulus
   (Japan: 2.78x, OECD range: 1.7-2.9x)
2. **Health-Led Growth Hypothesis (HLGH)** -- bidirectional causality between
   health spending and GDP growth (confirmed across 38 OECD countries)
3. **Health-Capital Tempo Effect** -- deferred supply-side returns from
   health-capital stock accumulation (from `healthcare_tempo_poc`)
4. **Three-Layer Tempo Analogy** -- the Bongaarts-Feeney quantum-tempo
   decomposition ported from Population to GDP to Healthcare, showing that
   healthcare exhibits the largest tempo drift (+0.15 yr/yr vs GDP +0.04)
5. **Equipment Stock & Import Leakage** -- Japan's diagnostic imaging density
   (170.9 CT+MRI/million, 4x OECD median) as health capital asset; import
   leakage (~5% CHE) reducing effective I-O multiplier

## Key Results

- The **fiscal return ratio** tau*m / pf exceeds 1.0 in 5 of 13 countries on the
  demand side alone (France 1.18, Italy 1.13, Japan 1.09, Sweden 1.04, Finland 1.04),
  with the remaining countries recovering 76-96% of public costs.
- The **tempo model (M2)** outperforms the constant-lag model (M1) in **95% of
  39 countries** (mu_H1 = +0.15 yr/yr), confirming that flow-only evaluation
  systematically underestimates returns.
- The **three-layer analogy** (Population -> GDP -> Healthcare) reveals that
  healthcare has the largest tempo drift among all three domains, making it
  the field where tempo correction matters most.
- **Counterfactual analysis**: reducing Japan's equipment density to OECD average
  drops fiscal return to 0.98 (unsustainable); domestic manufacturing raises it
  to 1.09. High equipment density is a sustainability component, not excess.

## Country Selection

Countries are included based on published I-O multiplier evidence:
- **(a)** Direct I-O study of the healthcare sector in a peer-reviewed journal
  or official government report.
- **(b)** Backward-linkage coefficient from the EU-28 I-O framework study
  (Gutierrez-Hernandez & Abasolo-Alesson 2021).

13 countries with multiplier evidence + OECD synthesis average.

## Structure

```
healthcare_economic_effect/
  scripts/
    analyze_healthcare_economic_effect.py   # Data + figures (8 figs, bilingual)
    create_manuscript_ja.py                 # Japanese manuscript (docx + pptx)
    create_manuscript_en.py                 # English manuscript (docx + pptx)
  data/                                     # CSV + JSON outputs
  output/
    docx/     # Manuscripts (JA/EN)
    pptx/     # Editable figures (1 slide per figure, JA/EN)
    figures/  # PNG figures (8 EN + 4 JA variants)
  README.md
```

## Reproduce

```bash
cd healthcare_economic_effect
pip install python-docx python-pptx numpy pandas matplotlib
python scripts/analyze_healthcare_economic_effect.py
python scripts/create_manuscript_ja.py
python scripts/create_manuscript_en.py
```

## Connection to Companion Papers

This project integrates findings from three companion analyses:

| Paper | Repository Location | Key Contribution |
|---|---|---|
| Population tempo (Onishi 2026a) | `population_tempo_paper/` | Bongaarts-Feeney framework + sigma |
| GDP tempo (Onishi 2026b) | `gdp_tempo_paper/` | Time-to-build mu, intangible K beta |
| Healthcare PoC (Onishi 2026c) | `healthcare_tempo_poc/` | Candidate A-H: mu_H drift +0.15 yr/yr |
| **This paper** | `healthcare_economic_effect/` | I-O + tempo dual-return sustainability |

The GDP paper's Section 6.4 previews the healthcare companion. The three-layer
analogy (Figure 2) connects all four papers into a unified programme.

## References

Key sources (24 total in manuscripts):
- Yamada & Imanaka (2015) *Environ Health Prev Med* -- Japan I-O multiplier 2.78
- Dupor & Guerrero (2021) *Econ Inq* -- US Medicare multiplier 1.7
- Gutierrez-Hernandez & Abasolo-Alesson (2021) *Cost Eff Resour Alloc* -- EU-28 I-O framework
- Ertugrul et al. (2024) *Front Public Health* -- HLGH in 38 OECD countries
- Bongaarts & Feeney (1998) *Popul Dev Rev* -- Tempo effect origin
- Goldstein, Lutz & Scherbov (2003) *Popul Dev Rev* -- Forgotten parameter sigma
- Onishi (2026) Working Papers -- GDP tempo + healthcare PoC companion papers
- OECD (2023) *Health at a Glance* -- CT/MRI density data
- MHLW (2021) Pharmaceutical production statistics -- Japan trade data
- Preston (1975) *Popul Stud* -- Original Preston Curve (overfit demonstrated)

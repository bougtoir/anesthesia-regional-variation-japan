# Modern Re-examination of 52 Canonical Curves

**定説曲線の現代的再検証 — 外れ値依存性・サンプルサイズ・非線形関係の脆弱性**

Onishi T. 2026.

## Overview

This project systematically re-examines 52 "canonical" curvilinear relationships across 8 academic disciplines using modern model selection methods (F-test, AIC/BIC, LOOCV, Cook's distance sensitivity analysis).

## Key Findings

| Verdict | Count | % |
|---------|-------|---|
| Not Significant | 22 | 42% |
| Robust Nonlinear | 18 | 35% |
| Outlier-Dependent | 12 | 23% |

**~65% of textbook nonlinear relationships fail at least one modern robustness test.**

## Disciplines Covered

- **A. Economics (12)**: Phillips, Laffer, Kuznets, EKC, Beveridge, Okun, Engel, J-Curve, Rahn, Gravity, Great Gatsby, Balassa-Samuelson
- **B. Public Health (10)**: Preston, Easterlin, McKeown, LNT, Fries, Barker, Omran, Wilkinson, BMI-Mortality, Alcohol-Mortality
- **C. Demography (6)**: Demographic Transition, Bongaarts-Feeney, Lee-Carter, Coale-Trussell, Replacement Migration, Second Demographic Transition
- **D. Environmental Science (6)**: Species-Area, Hubbert Peak Oil, Keeling, HANPP, Jevons Paradox, Forest Transition
- **E. Psychology (5)**: Yerkes-Dodson, Ebbinghaus, Weber-Fechner, Dunning-Kruger, Happiness U-Curve
- **F. Physics (4)**: Hubble's Law, Kleiber's Law, Gutenberg-Richter, Moore's Law
- **G. Political Science (5)**: Lipset, Duverger, Zipf, Crime-Temperature, Putnam
- **H. Agriculture (4)**: Mitscherlich, Borlaug, Micronutrient U-shape, Body Weight Set-Point

## Methods

1. **Nested F-test**: Linear (restricted) vs quadratic (unrestricted)
2. **AIC/BIC**: Model selection across linear, quadratic, and log models
3. **LOOCV RMSE**: Out-of-sample predictive accuracy
4. **Cook's Distance**: Top 3 influential points removed; F-test repeated

## Structure

```
canonical_curves_reexamination/
├── README.md
├── manuscript_canonical_curves_en.docx    # English manuscript
├── manuscript_canonical_curves_ja.docx    # Japanese manuscript
├── figures_canonical_curves.pptx          # Editable figures
├── scripts/
│   ├── core_analysis.py                   # Statistical framework
│   ├── data_economics.py                  # Economics curves (1-12)
│   ├── data_health.py                     # Public health curves (13-22)
│   ├── data_demography.py                 # Demography curves (23-28)
│   ├── data_environment.py                # Environmental curves (29-34)
│   ├── data_psychology.py                 # Psychology curves (35-39)
│   ├── data_physics.py                    # Physics curves (40-43)
│   ├── data_political.py                  # Political science curves (44-48)
│   ├── data_agriculture.py                # Agriculture curves (49-52)
│   ├── run_all_analyses.py                # Master runner
│   ├── generate_figures.py                # Figure generation
│   ├── generate_manuscript.py             # English docx
│   ├── generate_manuscript_ja.py          # Japanese docx
│   └── generate_pptx.py                   # PPTX figures
├── results/
│   ├── summary_table.csv                  # Summary statistics
│   └── full_results.json                  # Complete results
└── figures/
    ├── fig1_verdict_distribution.png
    ├── fig2_sensitivity_analysis.png
    ├── fig3_model_comparison.png
    ├── fig4_loocv_comparison.png
    └── fig5_sample_size.png
```

## Reproduction

```bash
cd scripts
python run_all_analyses.py      # Run all 52 analyses
python generate_figures.py      # Generate figures
python generate_manuscript.py   # Generate English manuscript
python generate_manuscript_ja.py # Generate Japanese manuscript
python generate_pptx.py         # Generate editable figures
```

## Dependencies

```
numpy scipy pandas statsmodels scikit-learn matplotlib seaborn python-docx python-pptx
```

# Weibull Reliability Analysis of Clinical Treatment Dropout

Five-domain comparative study applying reliability engineering's Weibull distribution to characterise hazard patterns in clinical treatment dropout.

## Domains

| Domain | Datasets | Total N | Mean k (range) |
|--------|----------|---------|----------------|
| HIV/ART Treatment Discontinuation | 5 regions (IeDEA) | 516,000 | 0.597 (0.582–0.627) |
| Antipsychotic Discontinuation (Schizophrenia) | 6 (CATIE, EUFEST, Finland) | 63,682 | 0.857 (0.683–0.958) |
| Substance Use Disorder Treatment | 6 substance types | 75,000 | 0.639 (0.511–0.733) |
| Cardiac Rehabilitation | 5 phases/indications | 34,200 | 0.689 (0.613–0.775) |
| Clinical Trial Dropout | 5 therapeutic areas | 99,000 | 0.708 (0.679–0.742) |

**Key finding:** All 27 datasets across all 5 domains yielded k < 1, demonstrating a **universal decreasing failure rate (DFR) pattern** — dropout risk is highest in the early treatment period and decreases monotonically over time.

## Structure

```
weibull-clinical-dropout/
├── scripts/
│   ├── weibull_analysis_all.py      # Main analysis (data, Weibull fitting, figures)
│   └── generate_manuscripts.py       # Manuscript generation (EN/JA docx, EN pptx)
├── figures/                          # All generated figures (12 PNG files)
├── data/
│   └── weibull_results_summary.csv   # Complete results table
├── manuscripts/
│   ├── weibull_clinical_dropout_comprehensive_EN.docx
│   ├── weibull_clinical_dropout_comprehensive_JA.docx
│   └── weibull_clinical_dropout_figures_EN.pptx
└── README.md
```

## Requirements

```
pip install scipy numpy matplotlib seaborn python-docx python-pptx pandas
```

## Usage

```bash
# Run Weibull analysis and generate all figures
python scripts/weibull_analysis_all.py

# Generate manuscripts
python scripts/generate_manuscripts.py
```

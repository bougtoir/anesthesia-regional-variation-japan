# TATSUKI: Trust-Adjusted Transparent Scoring with Unified Knowledge Integration (襷)

A novel electoral accountability mechanism that institutionalizes retrospective evaluation by linking candidate trust coefficients to measured policy fulfillment, with empirical calibration using real-world pledge tracking data.

**Target Journal:** Electoral Studies (Elsevier)

## Overview

TATSUKI introduces a continuous accountability loop into representative democracy. The name derives from the Japanese 襷 (tasuki), the relay sash passed between runners in ekiden races — symbolizing the passing of accountability across electoral cycles.

### Core Mechanism
1. **Pledge Declaration** — Candidates declare weighted policy pledges before elections
2. **Election** — Citizens vote; effective votes are modulated by candidate trust coefficients
3. **Term in Office** — Elected officials implement policies
4. **Fulfillment Evaluation** — Independent body assesses pledge fulfillment
5. **Accountability Score** — Weighted sum of fulfillment scores: S = Σ wⱼ·fⱼ
6. **Trust Coefficient Update** — Score mapped to trust via influence function: τ = ω(S)

### Key Innovation
The trust coefficient attaches to **candidates**, not voters, preserving the one-person-one-vote principle while creating performance-based electoral incentives.

### Empirical Calibration
The model is calibrated using:
- **Polimeter** (polimeter.org): 1,050 coded promises from the Trudeau government across 3 parliamentary terms (2015–2025)
- **Thomson et al. (2017)**: Cross-national pledge fulfillment data from 20,000+ pledges across 57 elections in 12 countries

## Repository Structure

```
tasuki-electoral-model/
├── README.md                          # This file
├── DRD-CT_先行文献比較レポート.md       # Literature comparison report (Japanese)
├── src/
│   ├── create_figures.py              # Generate 7 PNG figures (matplotlib)
│   ├── create_pptx.py                 # Generate English & Japanese PPTX slides
│   ├── create_paper_en.py             # Generate English DOCX paper
│   ├── create_paper_ja.py             # Generate Japanese DOCX paper
│   └── empirical_calibration.py       # Polimeter + Thomson data calibration
└── output/
    ├── calibration_results.json       # Calibration output data
    ├── figures/                        # Generated PNG figures (7 files)
    │   ├── fig1_conceptual_overview.png
    │   ├── fig2_influence_functions.png
    │   ├── fig3_simulation_results.png
    │   ├── fig4_adversarial.png
    │   ├── fig5_sensitivity.png
    │   ├── fig6_positioning.png
    │   └── fig7_empirical_calibration.png
    ├── docx/                           # Generated DOCX papers
    │   ├── TATSUKI_Electoral_Studies_English.docx
    │   └── TATSUKI_Electoral_Studies_Japanese.docx
    └── pptx/                           # Generated PPTX presentations
        ├── TATSUKI_Figures_English.pptx  # Editable slides (English)
        └── TATSUKI_Figures_Japanese.pptx # Editable slides (Japanese)
```

## Dependencies

```
pip install matplotlib numpy scipy python-docx python-pptx Pillow seaborn pandas
```

## Regenerating Output Files

```bash
# Generate calibration data (required by figures)
python src/empirical_calibration.py

# Generate figures (required by papers and presentations)
python src/create_figures.py

# Generate papers
python src/create_paper_en.py
python src/create_paper_ja.py

# Generate presentations
python src/create_pptx.py
```

All output paths are relative to the repository root. No hardcoded absolute paths.

## Paper Contents

- **ODD Protocol** compliant agent-based model description
- **7 figures**: conceptual overview, influence functions, simulation results, adversarial robustness, sensitivity analysis, theoretical positioning, **empirical calibration & counterfactual analysis**
- **Empirical grounding**: Section 5.5 with Polimeter data calibration and counterfactual trust trajectory analysis
- **Bilingual**: Full paper in both English and Japanese
- **Editable presentations**: PPTX with editable shapes for flow diagrams (Slides 1 & 6) and embedded images for simulation/empirical plots (Slides 2-5, 7)

## License

All rights reserved. This is unpublished academic work.

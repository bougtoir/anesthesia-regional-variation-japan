# Spectral Causality — JMLR Submission

Causal Direction Estimation via Magnetic Laplacians and Hodge Decomposition.

## Files

| File | Description |
|------|-------------|
| `spectral_causality.tex` | Main LaTeX manuscript (jmlr2e format) |
| `references.bib` | BibTeX references (natbib/plainnat) |
| `jmlr2e.sty` | JMLR official style file |
| `figures/` | PNG figures (9 files) |
| `spectral_causality_figures.pptx` | Editable PowerPoint with all figures |
| `generate_figures_pptx.py` | Script to regenerate .pptx |
| `Makefile` | Build system |

## Build

```bash
# Full build (PDF + pptx)
make all

# PDF only
make spectral_causality.pdf

# Create submission archive
make archive
```

Requires: `texlive-latex-base`, `texlive-latex-recommended`, `texlive-latex-extra`, `texlive-fonts-recommended`, `texlive-science`, `latexmk`, `python3-pptx`.

## Structure

The manuscript is organized into 10 sections + appendix:

1. Introduction
2. Preliminaries: Graph Laplacians
3. The Magnetic Laplacian
4. Spectral Causality: Formulation (DPI, SCC, SCD)
5. Hodge Decomposition and Causal Flow
6. Related Work
7. Experiments (UCI Heart Disease)
8. Phase Transition in Causal Structure
9. Ensemble Causal Direction (ECD)
10. Identifiability
11. Discussion and Future Work
A. Proof Sketch for Phase-Direction Correspondence

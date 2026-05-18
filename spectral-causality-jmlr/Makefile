# Makefile for Spectral Causality JMLR paper

TEX = spectral_causality
PDF = $(TEX).pdf
BIB = references.bib

.PHONY: all clean pptx coverletter

all: $(PDF) pptx coverletter

$(PDF): $(TEX).tex $(BIB)
	pdflatex -interaction=nonstopmode $(TEX).tex
	bibtex $(TEX)
	pdflatex -interaction=nonstopmode $(TEX).tex
	pdflatex -interaction=nonstopmode $(TEX).tex

pptx: spectral_causality_figures.pptx

spectral_causality_figures.pptx: generate_figures_pptx.py figures/*.png
	python3 generate_figures_pptx.py

coverletter: cover_letter.pdf

cover_letter.pdf: cover_letter.tex jmlr_coverletter_preamble.tex
	pdflatex -interaction=nonstopmode cover_letter.tex

clean:
	rm -f *.aux *.log *.bbl *.blg *.fls *.fdb_latexmk *.out *.toc \
	      *.synctex.gz *.pdf *.pptx

archive: $(PDF) pptx coverletter
	tar czf spectral_causality_jmlr.tar.gz \
	    $(TEX).tex $(BIB) jmlr2e.sty Makefile \
	    cover_letter.tex jmlr_coverletter_preamble.tex \
	    generate_figures_pptx.py generate_en_docx.py generate_cover_letter.py \
	    figures/ \
	    $(PDF) cover_letter.pdf spectral_causality_figures.pptx
	@echo "Created spectral_causality_jmlr.tar.gz"

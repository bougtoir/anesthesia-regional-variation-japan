# Khmer inscription vocabulary analysis

Pilot analysis for the idea: “クメール碑文における大乗→上座部の語彙変遷”.

The script queries the public SEAlang Corpus of Khmer Inscriptions interface and produces:

- `outputs/term_summary.csv`
- `outputs/term_distribution.csv`
- `outputs/term_contexts.csv`
- `outputs/analysis_report.md`
- `outputs/term_timeline_by_decade.png`
- `outputs/term_group_by_century.png`

## Data source

- SEAlang Corpus of Khmer Inscriptions: http://sealang.net/ok/corpus-top.htm
- Corpus description: http://sealang.net/classic/khmer/cki-bot.htm

SEAlang describes this corpus as a demonstration version with incomplete, not fully proofread texts and rough era/region divisions. Treat these outputs as exploratory signals for follow-up philological review.

## Reproduce

```bash
python3 -m pip install pandas matplotlib
python3 khmer_inscription_analysis/scripts/analyze_khmer_inscriptions.py
```

The script intentionally uses only the standard library for HTTP/HTML parsing plus `pandas` and `matplotlib` for tabular output and plots.

## Current pilot signal

The generated report suggests:

- Mahayana-oriented markers such as `vajra`, `lokeśvara`, `bodhisattva`, and `mahābodhi` appear in the queried SEAlang distribution mostly between the late 9th and 12th centuries.
- Theravada/Pali-oriented markers queried here (`thera`, `sāsana`, `bhikkhu`) are sparse and appear later in this pilot sample.
- Śaiva/Vaiṣṇava/Brahmanical comparator vocabulary is more frequent overall, which is useful as a baseline for testing whether Buddhist vocabulary is embedded in broader Indic religious epigraphy rather than isolated as a separate corpus.

This is not yet proof of “absorption” or “survival”; it is a reproducible starting point for selecting inscriptions and lexical variants for closer reading.

# Regional Variation in Perioperative and Chronic Pain-Related Prescribing Across Japan

An integrated ecological study using the National Database of Health Insurance Claims (NDB) Open Data.

## Overview

This project investigates regional (prefecture-level) variation in pain-related prescribing across Japan's 47 prefectures, using publicly available NDB Open Data and DPC data.

### Phase 1: Acute Perioperative Pain
- Constructed a composite "analgesic prescribing index" (inpatient NSAIDs, acetaminophen, opioids per surgical case) for each prefecture
- Tested the hypothesis that Tohoku region residents are "more stoic" about pain
- **Finding**: Tohoku's index was 11.7% *above* the national mean (P=0.031), **rejecting** the stoicism hypothesis

### Phase 2: Chronic Postsurgical Pain (CPSP) Proxy
- Used outpatient neuropathic pain drug prescriptions (pregabalin, mirogabalin, duloxetine) per surgical case as a CPSP proxy
- Adjusted for confounders: diabetes, herpes zoster, fibromyalgia, depression, anxiety disorder prescriptions
- **Finding**: Tohoku's apparent excess was **no longer significant** after confounder adjustment (P=0.150); diabetes medication was the strongest confounder (r=0.87)

## Directory Structure

```
ndb-pain-regional-variation-japan/
├── README.md
├── scripts/           # Python analysis and manuscript generation scripts
│   ├── analyze_ndb.py                    # Phase 1 NDB data analysis
│   ├── cpsp_integrated_analysis.py       # Phase 2 integrated analysis
│   ├── cpsp_figures.py                   # Figure generation
│   ├── extract_confounder_data.py        # Confounder data extraction
│   ├── create_integrated_docx_en.py      # BJA manuscript (English)
│   ├── create_integrated_docx_ja.py      # BJA manuscript (Japanese)
│   ├── create_ja_journal_docx_en.py      # JA manuscript (English)
│   ├── create_ja_journal_docx_ja.py      # JA manuscript (Japanese)
│   ├── create_pptx_en.py                 # Presentation (English)
│   ├── create_pptx_ja.py                 # Presentation (Japanese)
│   └── ...
├── data/              # Source data (NDB Open Data extracts)
│   ├── admission_addons_prefecture.xlsx
│   ├── inpatient_drugs_prefecture.xlsx
│   ├── outpatient_drugs_prefecture.xlsx
│   ├── surgery_prefecture.xlsx
│   └── anesthesia_prefecture.xlsx
├── output/            # Generated figures, manuscripts, and results
│   ├── JA_manuscript_EN.docx             # Journal of Anesthesia (English)
│   ├── JA_manuscript_JA.docx             # Journal of Anesthesia (Japanese)
│   ├── BJA_integrated_manuscript_EN.docx # BJA format (English)
│   ├── BJA_integrated_manuscript_JA.docx # BJA format (Japanese)
│   ├── figures_EN.pptx                   # Presentation (English)
│   ├── figures_JA.pptx                   # Presentation (Japanese)
│   ├── fig1-6, sfig1 (PNG)              # All figures
│   ├── cpsp_integrated_results.csv       # Full dataset (47 prefectures)
│   └── cpsp_regression_summary.json      # Regression statistics
└── reports/           # Background research and feasibility reports
    ├── postoperative_pain_regional_differences_report.md
    ├── postoperative_pain_japan_regional_report.md
    └── cpsp_feasibility_report.md
```

## Manuscripts

| Target Journal | Language | File |
|---|---|---|
| Journal of Anesthesia | English | `output/JA_manuscript_EN.docx` |
| Journal of Anesthesia | Japanese | `output/JA_manuscript_JA.docx` |
| British Journal of Anaesthesia | English | `output/BJA_integrated_manuscript_EN.docx` |
| British Journal of Anaesthesia | Japanese | `output/BJA_integrated_manuscript_JA.docx` |

## Data Sources

- **NDB Open Data** (厚生労働省 NDBオープンデータ): Prefecture-level prescription volumes
- **DPC Open Data**: Surgical case counts by prefecture
- All data are publicly available aggregate statistics (no individual-level data)

## Reporting Guidelines

- STROBE checklist for observational studies
- RECORD extension for routinely collected health data

---

# English Translation

---

# Regional Variation in Perioperative and Chronic Pain-Related Prescribing Across Japan

An integrated ecological study using the National Database of Health Insurance Claims (NDB) Open Data.

## Overview

This project investigates regional (prefecture-level) variation in pain-related prescribing across Japan's 47 prefectures, using publicly available NDB Open Data and DPC data.

### Phase 1: Acute Perioperative Pain
- Constructed a composite "analgesic prescribing index" (inpatient NSAIDs, acetaminophen, opioids per surgical case) for each prefecture
- Tested the hypothesis that Tohoku region residents are "more stoic" about pain
- **Finding**: Tohoku's index was 11.7% *above* the national mean (P=0.031), **rejecting** the stoicism hypothesis

### Phase 2: Chronic Postsurgical Pain (CPSP) Proxy
- Used outpatient neuropathic pain drug prescriptions (pregabalin, mirogabalin, duloxetine) per surgical case as a CPSP proxy
- Adjusted for confounders: diabetes, herpes zoster, fibromyalgia, depression, anxiety disorder prescriptions
- **Finding**: Tohoku's apparent excess was **no longer significant** after confounder adjustment (P=0.150); diabetes medication was the strongest confounder (r=0.87)

## Directory Structure

```
ndb-pain-regional-variation-japan/
├── README.md
├── scripts/           # Python analysis and manuscript generation scripts
│   ├── analyze_ndb.py                    # Phase 1 NDB data analysis
│   ├── cpsp_integrated_analysis.py       # Phase 2 integrated analysis
│   ├── cpsp_figures.py                   # Figure generation
│   ├── extract_confounder_data.py        # Confounder data extraction
│   ├── create_integrated_docx_en.py      # BJA manuscript (English)
│   ├── create_integrated_docx_ja.py      # BJA manuscript (Japanese)
│   ├── create_ja_journal_docx_en.py      # JA manuscript (English)
│   ├── create_ja_journal_docx_ja.py      # JA manuscript (Japanese)
│   ├── create_pptx_en.py                 # Presentation (English)
│   ├── create_pptx_ja.py                 # Presentation (Japanese)
│   └── ...
├── data/              # Source data (NDB Open Data extracts)
│   ├── admission_addons_prefecture.xlsx
│   ├── inpatient_drugs_prefecture.xlsx
│   ├── outpatient_drugs_prefecture.xlsx
│   ├── surgery_prefecture.xlsx
│   └── anesthesia_prefecture.xlsx
├── output/            # Generated figures, manuscripts, and results
│   ├── JA_manuscript_EN.docx             # Journal of Anesthesia (English)
│   ├── JA_manuscript_JA.docx             # Journal of Anesthesia (Japanese)
│   ├── BJA_integrated_manuscript_EN.docx # BJA format (English)
│   ├── BJA_integrated_manuscript_JA.docx # BJA format (Japanese)
│   ├── figures_EN.pptx                   # Presentation (English)
│   ├── figures_JA.pptx                   # Presentation (Japanese)
│   ├── fig1-6, sfig1 (PNG)              # All figures
│   ├── cpsp_integrated_results.csv       # Full dataset (47 prefectures)
│   └── cpsp_regression_summary.json      # Regression statistics
└── reports/           # Background research and feasibility reports
    ├── postoperative_pain_regional_differences_report.md
    ├── postoperative_pain_japan_regional_report.md
    └── cpsp_feasibility_report.md
```

## Manuscripts

| Target Journal | Language | File |
|---|---|---|
| Journal of Anesthesia | English | `output/JA_manuscript_EN.docx` |
| Journal of Anesthesia | Japanese | `output/JA_manuscript_JA.docx` |
| British Journal of Anaesthesia | English | `output/BJA_integrated_manuscript_EN.docx` |
| British Journal of Anaesthesia | Japanese | `output/BJA_integrated_manuscript_JA.docx` |

## Data Sources

- **NDB Open Data** (Ministry of Health, Labor and Welfare NDB Open Data): Prefecture-level prescription volumes
- **DPC Open Data**: Surgical case counts by prefecture
- All data are publicly available aggregate statistics (no individual-level data)

## Reporting Guidelines

- STROBE checklist for observational studies
- RECORD extension for routinely collected health data


# Buddhist food lexicon geography prototype

This directory implements project idea 1 from the linked Devin session: a reproducible prototype for mapping food-related vocabulary in Buddhist Vinaya corpora with open data.

## What it does

`scripts/analyze_food_terms.py` downloads or reuses open data from:

- [CBETA API 2025R3](https://cbdata.dila.edu.tw/stable): Chinese Vinaya texts and translation metadata for T1421, T1425, T1428, T1435, and T1442.
- [SuttaCentral `bilara-data`](https://github.com/suttacentral/bilara-data/tree/published/root/pli/ms/vinaya): Pāli Vinaya root-text JSON files.

It counts a small seed lexicon of food terms such as 粥/`yāgu`, 乳/`khīra`, 酥/`sappi`, 蜜/`madhu`, 飯/`odana`, 肉/`maṃsa`, 魚/`maccha`, 油/`tela`, and related grain/seasoning/drink terms. Outputs include CSV tables, a heatmap, a category profile, and a Leaflet HTML map.

## Reproduce

From the repository root:

```bash
python3 buddhist_food_lexicon_map/scripts/analyze_food_terms.py
```

The script uses only Python standard-library networking plus `pandas` and `matplotlib`, both already available in this repo environment.

## Outputs

- `output/food_term_counts.csv`: term-level counts and normalized counts per 10,000 characters.
- `output/work_metadata.csv`: corpus, source URL, CBETA translation metadata, and map coordinates.
- `output/corpus_summary.csv`: total food-term mentions by corpus.
- `output/cbeta_context_samples.csv`: small Chinese context samples for key terms.
- `output/figures/food_terms_heatmap.png`: per-corpus term frequency heatmap.
- `output/figures/food_category_profile.png`: stacked category profile by Vinaya lineage.
- `output/figures/food_terms_map.html`: interactive map. Open this file in a browser.

## Current prototype results

| Corpus | School/lineage | Total food mentions | Text length | Mentions per 10k chars |
|---|---:|---:|---:|---:|
| T1421 | Mahīśāsaka | 500 | 346,291 | 14.44 |
| T1425 | Mahāsāṃghika | 1,535 | 562,940 | 27.27 |
| T1428 | Dharmaguptaka | 1,214 | 748,452 | 16.22 |
| T1435 | Sarvāstivāda | 1,533 | 822,652 | 18.63 |
| T1442 | Mūlasarvāstivāda | 697 | 471,949 | 14.77 |
| Pāli Vinaya | Theravāda | 321 | 3,880,316 | 0.83 |

Top aggregate English lemmas in this seed lexicon are cooked rice/meal, meat, oil, honey, ghee/butter, gruel/congee, and milk. Among the Chinese Vinaya translations, grain/preparation terms dominate the normalized category profile, followed by animal-food, dairy, fat/medicine, and sweetener/medicine terms.

## Interpretation cautions

- The map locations are evidence anchors from CBETA translation metadata plus a Sri Lanka centroid for the Pāli Vinaya, not exact composition sites.
- Direct cross-language frequency comparison is exploratory only. Chinese one-character lexical matching and romanized Pāli token matching have different recall/precision profiles.
- The seed lexicon is intentionally small. A publishable analysis should expand it with variant spellings, compound terms, named preparations, and manual validation of false positives.
- CBETA plain text in this environment appears to be normalized/simplified in places; counts should be treated as computational signals, not diplomatic editions.

## Suggested next steps

1. Expand the lexicon by category, including compounds such as 乳粥, 石蜜, 生酥, 五藥, `khīrapāyāsa`, and Pāli inflected forms.
2. Add sentence/paragraph extraction and manual validation labels for high-frequency terms.
3. Link results to explicit Vinaya chapters or rule contexts rather than whole works only.
4. Add epigraphic and temple GIS datasets such as SIDDHAM or Nanyang Buddhist Space when the question shifts from textual food vocabulary to material geography.

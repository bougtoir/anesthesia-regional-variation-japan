# Journal Submission Plan: Lancet Regional Health – Western Pacific

## Target Journal
- **Journal**: The Lancet Regional Health – Western Pacific
- **CiteScore**: 11.40
- **Editor-in-Chief**: Chi Chiu Leung
- **Submission system**: Editorial Manager
- **Open Access**: CC BY-NC-ND
- **APC**: Up to $5,010 USD (waiver policy available)

## Selection Rationale
1. **Precedent**: Taira et al. (2021) published an ecological study using NDB Open Data in this journal (Lancet Reg Health West Pac 2021; 12: 100170)
2. **Regional focus**: Journal covers Western Pacific region; Japan is a key nation
3. **Scope match**: Population-level health variation using routinely-collected data
4. **Impact**: High CiteScore (11.40) for a regional health journal
5. **After EJP desk rejection**: Manuscript repositioned from European to Asia-Pacific audience

## Previous Submission History
| Journal | Submitted | Decision | Notes |
|---------|-----------|----------|-------|
| European Journal of Pain (EJP) | — | Desk reject | Too ecological/epidemiological for clinical pain journal |

## Manuscript Details
- **Title**: Regional Heterogeneity in Pain-Related Prescribing Across Japan's 47 Prefectures Challenges a Stoic Monolithic Patient Stereotype: an ecological study
- **Type**: Original Article
- **Abstract**: 268 words (limit: 300)
- **Body**: ~2,500 words (limit: 4,000)
- **References**: 23 (limit: 30)
- **Display items**: 3 figures + 2 tables = 5 (limit: 5)
- **Reporting guidelines**: STROBE + RECORD

## Key Adaptations from EJP Version
1. References: Harvard (author-date) → Vancouver (numbered superscript [n])
2. Abstract: Unstructured → Structured (Background, Methods, Findings, Interpretation, Funding)
3. Added "Research in Context" panel (Lancet Group requirement)
4. Word count: Reduced from ~5,000 to ~2,500 words
5. Display items: Reduced to ≤5 combined figures + tables
6. Reporting: STROBE only → STROBE + RECORD extension
7. Audience: Reframed for Western Pacific regional relevance

## Submission Files
| File | Description |
|------|-------------|
| `LRH_manuscript_EN.docx` | English manuscript with inline figures/tables |
| `LRH_manuscript_JA.docx` | Japanese manuscript with inline figures/tables |
| `LRH_cover_letter.docx` | Cover letter to Prof. Chi Chiu Leung |
| `LRH_Table1_EN.docx` | Editable Table 1 (regional summary) |
| `LRH_Table2_EN.docx` | Editable Table 2 (regression models) |
| `LRH_figures_EN.pptx` | Editable figures (widescreen PPTX) |
| `LRH_STROBE_RECORD_checklist.docx` | STROBE + RECORD checklist |
| `LRH_Fig1_neuropathic_unadjusted.tiff` | Figure 1 (300 DPI TIFF) |
| `LRH_Fig2_confounder_correlations.tiff` | Figure 2 (300 DPI TIFF) |
| `LRH_Fig3_region_unadj_vs_adj.tiff` | Figure 3 (300 DPI TIFF) |

## Generation Scripts
All in `scripts/`:
- `create_lancet_rh_docx_en.py` — EN manuscript
- `create_lancet_rh_docx_ja.py` — JA manuscript
- `create_lancet_rh_cover_letter.py` — Cover letter
- `create_lancet_rh_tables_en.py` — Editable tables
- `create_lancet_rh_pptx_en.py` — Editable PPTX figures
- `create_lancet_rh_record_checklist.py` — STROBE/RECORD checklist
- `create_lancet_rh_tiff_figures.py` — TIFF figure conversion

## Submission Status
- [ ] Files generated
- [ ] Author review complete
- [ ] Submitted to Editorial Manager
- [ ] Desk decision received

## Notes
- SCR (Standardized Claim Ratios) cannot be computed: NDB Open Data lacks age-sex stratified data at drug-prefecture level. Noted as limitation and future direction.
- Per-capita rates used as complementary population standardisation.
- Figures use existing EN/JA PNG outputs from cpsp_figures_en.py / cpsp_figures.py.

# Regional Variation in Anaesthesia Practice in Japan

Cross-sectional analysis of national claims data across 335 secondary medical areas (SMAs).

## Overview

This project analyses regional variation in anaesthesia practice across Japan using publicly available standardised claim ratio (SCR) data. Three independent sensitivity analyses demonstrate that observed variation is structural rather than an artefact of insurance auditing.

**Key findings:**
- Coefficients of variation: GA 54.6%, epidural 83.2%, continuous epidural 64.9%
- University hospital presence explains 38.5% of GA (L008) variance (Cohen's d = 1.78)
- Insurance audit rate differences (0.2pp) explain <1% of observed variation
- Combined GA-epidural technique (L003) is 1.73x higher in university hospital areas

## Repository Structure

```
scripts/          Analysis and visualisation scripts
  analysis_v2.py              Main statistical analysis
  generate_all_maps.py        2D choropleth maps (JP)
  generate_maps_en.py         2D choropleth maps (EN)
  create_3d_v2.py             3D extruded maps (EN)
  create_3d_v2_jp.py          3D extruded maps (JP)
  screenshot_3d_v2.py         Screenshot 3D HTML maps (EN)
  screenshot_3d_v2_jp.py      Screenshot 3D HTML maps (JP)
  create_bmj_docx.py          BMJ-format manuscript (EN)
  create_bmj_docx_jp.py       BMJ-format manuscript (JP)
  create_pptx_en.py           Figure/table slides (EN)
  create_pptx_jp.py           Figure/table slides (JP)

data/             Input data files
  corrected_metrics_final.csv   SCR data with zero-padded area codes
  anesthesiologist_by_sma.csv   Anaesthesiologist count by SMA
  scr_*.csv                     Raw SCR data from Cabinet Office
  physician_table24.csv         Physician statistics
  univ_hospital_mapping_v2.json University hospital to SMA mapping

gis_data/         Geographic data
  merged_enriched.gpkg          335 SMAs with merged metrics
  northern_territories.gpkg     Northern Territories boundaries
  pref_simplified.gpkg          Prefectural boundaries

output/           Generated figures
  maps_2d_en/     2D choropleth maps (English labels)
  maps_2d_jp/     2D choropleth maps (Japanese labels)
  maps_3d_en/     3D extruded map screenshots (English)
  maps_3d_jp/     3D extruded map screenshots (Japanese)
  maps_3d_html/   Interactive 3D HTML maps

documents/        Manuscript and presentation files
  regional_anaesthesia_BMJ_EN.docx   BMJ manuscript (English)
  regional_anaesthesia_BMJ_JP.docx   BMJ manuscript (Japanese)
  regional_anaesthesia_figures_EN.pptx  Figure slides (English)
  regional_anaesthesia_figures_JP.pptx  Figure slides (Japanese)

reports/          Analysis reports and brainstorm documents
```

## Data Sources

- **SCR data**: Cabinet Office "Regional Variation Visualisation" (FY2022)
- **Physician statistics**: e-Stat (2022 Survey of Physicians, Dentists, and Pharmacists)
- **GIS boundaries**: National Land Numerical Information (A38-20, N03)

## Anaesthesia Codes Analysed

| Code | Procedure | Clinical significance |
|------|-----------|----------------------|
| L008 | Closed-circuit general anaesthesia | Primary GA indicator |
| L002 | Epidural anaesthesia | Regional technique |
| L003 | Continuous epidural infusion | GA+epidural combination indicator |
| L004 | Spinal anaesthesia | GA alternative |
| L009 | Anaesthesia management fee I | Specialist staffing proxy |
| L100 | Nerve block (inpatient) | Pain clinic activity |

## Requirements

- Python 3.12+
- geopandas, pandas, scipy, matplotlib, plotly, python-docx, python-pptx, Pillow
- playwright (for 3D map screenshots)

## Study Design

STROBE-compliant cross-sectional ecological study. BMJ submission format.

## License

[To be determined by authors]

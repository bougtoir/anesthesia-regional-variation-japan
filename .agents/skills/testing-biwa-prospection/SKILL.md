---
name: testing-biwa-prospection
description: Test the Lake Biwa archaeological prospection pipeline end-to-end. Use when verifying DEM fetch, terrain visualization, anomaly detection, or paleo-lake data changes.
---

# Testing: 琵琶湖考古学的探査パイプライン

## Prerequisites

- Python 3.12+ with numpy, scipy, matplotlib, pandas, requests, Pillow
- Internet access to GSI tile servers (public, no auth required)
- No secrets or credentials needed

## How to Run

```bash
cd /home/ubuntu/repos/wip
python -m biwa_archaeological_prospection.run --output /tmp/biwa_test_output --verbose
```

This fetches DEM tiles from GSI (国土地理院), runs terrain visualizations (RRIM, SVF, Hillshade, Slope, LRM), performs anomaly detection, and generates PNG outputs + report.md.

Expected runtime: ~60-90 seconds for 3 priority regions at zoom 14.

## What to Verify

### 1. Pipeline Completion
- Exit code = 0
- 5 output files: `paleo_biwa_overview.png`, 3 `analysis_*.png`, `report.md`
- All PNGs > 10KB (not trivial/empty)

### 2. DEM Data Validity
- **Awazu** (粟津): shape (768,768), elevation 80-270m, 100% valid pixels
- **Oyamada** (大山田): shape (768,1024), elevation 125-500m, 100% valid pixels  
- **Koka** (甲賀): shape (1024,1024), elevation 150-500m, 100% valid pixels
- No region should be all-NaN (would indicate tile fetch failure)

### 3. Terrain Visualizations
- Hillshade: values [0,1], std > 0.05
- Slope: values [0, ~50°], mean > 1°
- SVF: values [0,1], mean ~0.99
- LRM: has both positive and negative values
- RRIM: shape (H,W,3) uint8, variance in all 3 channels

### 4. Anomaly Detection
- >= 10 candidates per region (typically 50)
- All candidates within **tile-aligned** region bounds (slightly larger than input region)
- All scores > 0, areas > 200m²
- Sorted by score descending

### 5. Paleo-Lake Data
- 5 stages: 大山田湖 → 阿山湖 → 甲賀湖 → 蒲生沼沢地 → 堅田湖
- Center latitudes increase south→north (34.75 → 35.10)
- Ages non-strictly decrease (甲賀湖 and 蒲生沼沢地 both 2.6Ma — correct temporal overlap)
- 2 known sites with valid coordinates

## Known Gotchas

- **Tile-aligned bounds**: `actual_region` from `fetch_dem_region()` is slightly larger than input `GeoRegion` because XYZ tile boundaries don't align with lat/lon. Candidates may appear marginally outside input bounds — use `actual_region` for validation.
- **Japanese font rendering**: Matplotlib needs IPA Gothic font. If CJK characters show as boxes, check `fc-list :lang=ja` for IPA fonts and clear `~/.cache/matplotlib/fontlist-*.json`.
- **GSI server availability**: Tile fetches have 3 retries with exponential backoff. If GSI servers are down, the pipeline will log errors but may produce partial/NaN DEMs.
- **No recording needed**: All testing is CLI/shell-based. Do not start screen recordings.

## Verification Script (Quick)

```python
import sys; sys.path.insert(0, '.')
from biwa_archaeological_prospection.paleo_biwa import PALEO_LAKE_STAGES, get_priority_regions
assert len(PALEO_LAKE_STAGES) == 5
assert len(get_priority_regions()) == 3
print('Basic data structures OK')
```

## Devin Secrets Needed

None — all data sources (GSI DEM tiles, GSJ geology tiles) are public APIs with no authentication.

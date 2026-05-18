---
name: biwa-prospection-testing
description: Testing and running the Lake Biwa archaeological prospection pipeline (DEM + NDVI)
---

## Running the Pipeline

```bash
cd /home/ubuntu/repos/wip
python -m biwa_archaeological_prospection.run --output /tmp/biwa_output --verbose
```

Default regions: Awazu (粟津), Oyamada (大山田), Koka (甲賀)

## Dependencies

```bash
pip install numpy scipy matplotlib pandas requests Pillow pystac-client rasterio
```

## Key Modules

| Module | Purpose |
|--------|--------|
| data_fetch.py | GSI DEM tile fetching |
| terrain_viz.py | RRIM, SVF, LRM, Hillshade |
| anomaly_detection.py | Z-score, circular/linear structure detection |
| vegetation_analysis.py | Sentinel-2 NDVI fetch + vegetation anomaly detection |
| paleo_biwa.py | Paleo-lake stage definitions (5 stages) |
| pipeline.py | Main orchestrator |
| run.py | CLI entry point |

## Data Sources (all public, no auth required)

- **DEM**: GSI tiles `https://cyberjapandata.gsi.go.jp/xyz/dem/{z}/{x}/{y}.txt`
- **Sentinel-2**: Earth Search STAC API `https://earth-search.aws.element84.com/v1` → COGs on AWS S3

## Testing Approach

All tests are shell-based (no UI). Key validations:

1. **NDVI validity**: values in [-1, 1], mean > 0, valid pixels > 90%
2. **Vegetation anomalies**: both high_ndvi and low_ndvi types detected, Z-score > 2.0, area >= 500m²
3. **Combined score**: shape matches DEM grid, differs from terrain-only score
4. **5-row PNG layout**: aspect ratio > 1.2 (taller than 4-row)
5. **Report sections**: 「植生解析 (NDVI)」 present for each region
6. **Graceful degradation**: ocean-only bbox returns None, pipeline continues
7. **Regression**: 5 paleo-lake stages, south→north ordering, ages 4.0→0.4 Ma

## Japanese Font Rendering

Matplotlib uses IPAGothic for CJK characters. Set in pipeline.py:
```python
plt.rcParams["font.family"] = ["IPAGothic", "DejaVu Sans"]
```

## Notes

- Sentinel-2 scenes are auto-selected for lowest cloud cover in Apr-Oct range
- NDVI failure is non-fatal: pipeline falls back to DEM-only analysis
- Tile boundaries don't perfectly align with lat/lon — use `actual_region` for coordinate validation
- 甲賀湖 and 蒲生沼沢地 overlap at 2.6 Ma (geologically correct)

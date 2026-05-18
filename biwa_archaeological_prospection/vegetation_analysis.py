"""
植生解析モジュール — Sentinel-2 NDVI による植生異常検出

データソース:
- Sentinel-2 L2A (Cloud-Optimized GeoTIFF via Earth Search STAC API)
  - Band 4 (Red, 665nm, 10m)
  - Band 8 (NIR, 842nm, 10m)
- 認証不要: AWS上の公開COGに直接アクセス

処理フロー:
1. STAC APIで指定領域の低雲量Sentinel-2シーンを検索
2. rasterioでCOGからBand4/Band8を部分読み出し (HTTP range request)
3. NDVI = (NIR - Red) / (NIR + Red) を計算
4. 植生異常を統計的に検出 (Z-score)
5. DEM地形異常スコアとの重ね合わせ

参考:
- 琵琶湖湖畔域で自然範囲を超えた特定植物の分布 → 狩猟採集から栽培への移行痕跡
- 古琵琶湖各ステージの湖畔域における植生パターンの比較
"""

import logging
from dataclasses import dataclass, field
from typing import Optional

import numpy as np

logger = logging.getLogger(__name__)

EARTH_SEARCH_URL = "https://earth-search.aws.element84.com/v1"
SENTINEL2_COLLECTION = "sentinel-2-l2a"

# Sentinel-2 L2A の反射率スケール (10000で割って0-1にする)
S2_SCALE_FACTOR = 10000.0


@dataclass
class NDVIResult:
    """NDVI解析結果"""
    ndvi: np.ndarray
    red: np.ndarray
    nir: np.ndarray
    scene_date: str
    cloud_cover: float
    resolution_m: float
    bbox: tuple[float, float, float, float]  # (lon_min, lat_min, lon_max, lat_max)


@dataclass
class VegetationAnomaly:
    """植生異常候補"""
    lat: float
    lon: float
    ndvi_value: float
    ndvi_zscore: float
    area_m2: float
    anomaly_type: str  # "high_ndvi" or "low_ndvi"
    description: str = ""


def search_sentinel2_scenes(
    bbox: tuple[float, float, float, float],
    date_range: str = "2024-04-01/2024-10-31",
    max_cloud_cover: float = 20.0,
    max_items: int = 10,
) -> list[dict]:
    """
    指定領域のSentinel-2 L2Aシーンを検索。

    Parameters
    ----------
    bbox : (lon_min, lat_min, lon_max, lat_max) — WGS84
    date_range : 日付範囲 (ISO format)
    max_cloud_cover : 最大雲量(%)
    max_items : 最大取得数

    Returns
    -------
    list of STAC items as dicts with keys: id, datetime, cloud_cover, red_href, nir_href
    """
    try:
        from pystac_client import Client
    except ImportError:
        logger.error("pystac-client not installed. Run: pip install pystac-client")
        return []

    try:
        catalog = Client.open(EARTH_SEARCH_URL)
        search = catalog.search(
            collections=[SENTINEL2_COLLECTION],
            bbox=list(bbox),
            datetime=date_range,
            query={"eo:cloud_cover": {"lt": max_cloud_cover}},
            max_items=max_items,
        )
        items = list(search.items())
    except Exception as e:
        logger.error("STAC search failed: %s", e)
        return []

    results = []
    for item in items:
        if "red" not in item.assets or "nir" not in item.assets:
            continue
        results.append({
            "id": item.id,
            "datetime": str(item.datetime),
            "cloud_cover": item.properties.get("eo:cloud_cover", 0),
            "red_href": item.assets["red"].href,
            "nir_href": item.assets["nir"].href,
        })

    results.sort(key=lambda x: x["cloud_cover"])
    logger.info("Found %d Sentinel-2 scenes (bbox=%s, clouds<%.0f%%)", len(results), bbox, max_cloud_cover)
    return results


def fetch_ndvi_for_region(
    region,
    date_range: str = "2024-04-01/2024-10-31",
    max_cloud_cover: float = 20.0,
) -> Optional[NDVIResult]:
    """
    指定GeoRegionのNDVIを取得。

    Parameters
    ----------
    region : GeoRegion
    date_range : 検索日付範囲
    max_cloud_cover : 最大雲量(%)

    Returns
    -------
    NDVIResult or None
    """
    try:
        import rasterio
        from rasterio.warp import transform_bounds
        from rasterio.windows import from_bounds
    except ImportError:
        logger.error("rasterio not installed. Run: pip install rasterio")
        return None

    bbox = (region.lon_min, region.lat_min, region.lon_max, region.lat_max)
    scenes = search_sentinel2_scenes(bbox, date_range, max_cloud_cover)

    if not scenes:
        logger.warning("No Sentinel-2 scenes found for %s", region.name)
        return None

    scene = scenes[0]
    logger.info(
        "Using scene %s (date=%s, clouds=%.1f%%)",
        scene["id"], scene["datetime"], scene["cloud_cover"],
    )

    try:
        with rasterio.open(scene["red_href"]) as red_src:
            src_crs = red_src.crs
            src_bounds = transform_bounds("EPSG:4326", src_crs, *bbox)

            window = from_bounds(*src_bounds, transform=red_src.transform)
            red_data = red_src.read(1, window=window).astype(np.float64)
            actual_transform = rasterio.windows.transform(window, red_src.transform)
            resolution = red_src.res[0]

        with rasterio.open(scene["nir_href"]) as nir_src:
            nir_window = from_bounds(*src_bounds, transform=nir_src.transform)
            nir_data = nir_src.read(1, window=nir_window).astype(np.float64)

    except Exception as e:
        logger.error("Failed to read Sentinel-2 bands: %s", e)
        return None

    # Shape alignment
    min_h = min(red_data.shape[0], nir_data.shape[0])
    min_w = min(red_data.shape[1], nir_data.shape[1])
    red_data = red_data[:min_h, :min_w]
    nir_data = nir_data[:min_h, :min_w]

    # NDVI calculation
    red_refl = red_data / S2_SCALE_FACTOR
    nir_refl = nir_data / S2_SCALE_FACTOR

    denom = nir_refl + red_refl
    ndvi = np.where(denom > 0, (nir_refl - red_refl) / denom, np.nan)

    # Mask nodata (where both bands are 0)
    nodata_mask = (red_data == 0) & (nir_data == 0)
    ndvi[nodata_mask] = np.nan

    valid_pct = np.sum(~np.isnan(ndvi)) / ndvi.size * 100
    logger.info(
        "NDVI computed: shape=%s, range=[%.3f, %.3f], mean=%.3f, valid=%.1f%%",
        ndvi.shape, np.nanmin(ndvi), np.nanmax(ndvi), np.nanmean(ndvi), valid_pct,
    )

    return NDVIResult(
        ndvi=ndvi,
        red=red_refl,
        nir=nir_refl,
        scene_date=scene["datetime"],
        cloud_cover=scene["cloud_cover"],
        resolution_m=resolution,
        bbox=bbox,
    )


def detect_vegetation_anomalies(
    ndvi: np.ndarray,
    bbox: tuple[float, float, float, float],
    resolution_m: float = 10.0,
    zscore_threshold: float = 2.0,
    min_area_m2: float = 500.0,
    max_candidates: int = 50,
) -> list[VegetationAnomaly]:
    """
    NDVI配列から植生異常を検出。

    高NDVI異常: 周囲より異常に高い植生 → 栽培植物/人為的植栽の可能性
    低NDVI異常: 周囲より異常に低い植生 → 埋没遺構による植生抑制の可能性

    Parameters
    ----------
    ndvi : NDVI配列
    bbox : (lon_min, lat_min, lon_max, lat_max)
    resolution_m : ピクセル解像度(m)
    zscore_threshold : Z-scoreの閾値
    min_area_m2 : 最小面積フィルタ
    max_candidates : 最大候補数
    """
    from scipy.ndimage import uniform_filter, label

    valid_mask = ~np.isnan(ndvi)
    if np.sum(valid_mask) < 100:
        logger.warning("Too few valid NDVI pixels (%d)", np.sum(valid_mask))
        return []

    # Local Z-score
    ndvi_filled = np.where(valid_mask, ndvi, np.nanmean(ndvi))
    window = 51
    local_mean = uniform_filter(ndvi_filled, size=window)
    local_sq_mean = uniform_filter(ndvi_filled ** 2, size=window)
    local_std = np.sqrt(np.maximum(local_sq_mean - local_mean ** 2, 0))
    local_std = np.maximum(local_std, 1e-6)
    zscore = (ndvi_filled - local_mean) / local_std
    zscore[~valid_mask] = np.nan

    candidates = []
    pixel_area = resolution_m ** 2

    # High NDVI anomalies (potential cultivated/managed vegetation)
    high_mask = zscore > zscore_threshold
    labeled_high, n_high = label(high_mask)
    for i in range(1, n_high + 1):
        cluster = labeled_high == i
        area = np.sum(cluster) * pixel_area
        if area < min_area_m2:
            continue
        rows, cols = np.where(cluster)
        mean_row = np.mean(rows)
        mean_col = np.mean(cols)
        lat = bbox[3] - (mean_row / ndvi.shape[0]) * (bbox[3] - bbox[1])
        lon = bbox[0] + (mean_col / ndvi.shape[1]) * (bbox[2] - bbox[0])
        mean_ndvi = np.nanmean(ndvi[cluster])
        mean_z = np.nanmean(zscore[cluster])
        candidates.append(VegetationAnomaly(
            lat=lat, lon=lon,
            ndvi_value=mean_ndvi,
            ndvi_zscore=mean_z,
            area_m2=area,
            anomaly_type="high_ndvi",
            description=f"高NDVI異常 (NDVI={mean_ndvi:.3f}, z={mean_z:.1f}, {area:.0f}m²)"
                        f" — 栽培植物/人為的植栽の可能性",
        ))

    # Low NDVI anomalies (potential buried structures suppressing vegetation)
    low_mask = zscore < -zscore_threshold
    labeled_low, n_low = label(low_mask)
    for i in range(1, n_low + 1):
        cluster = labeled_low == i
        area = np.sum(cluster) * pixel_area
        if area < min_area_m2:
            continue
        rows, cols = np.where(cluster)
        mean_row = np.mean(rows)
        mean_col = np.mean(cols)
        lat = bbox[3] - (mean_row / ndvi.shape[0]) * (bbox[3] - bbox[1])
        lon = bbox[0] + (mean_col / ndvi.shape[1]) * (bbox[2] - bbox[0])
        mean_ndvi = np.nanmean(ndvi[cluster])
        mean_z = np.nanmean(zscore[cluster])
        candidates.append(VegetationAnomaly(
            lat=lat, lon=lon,
            ndvi_value=mean_ndvi,
            ndvi_zscore=mean_z,
            area_m2=area,
            anomaly_type="low_ndvi",
            description=f"低NDVI異常 (NDVI={mean_ndvi:.3f}, z={mean_z:.1f}, {area:.0f}m²)"
                        f" — 埋没遺構/土壌露出の可能性",
        ))

    candidates.sort(key=lambda c: abs(c.ndvi_zscore), reverse=True)
    candidates = candidates[:max_candidates]
    logger.info("Detected %d vegetation anomalies (high=%d, low=%d)",
                len(candidates),
                sum(1 for c in candidates if c.anomaly_type == "high_ndvi"),
                sum(1 for c in candidates if c.anomaly_type == "low_ndvi"))
    return candidates


def compute_combined_score(
    terrain_score: np.ndarray,
    ndvi_result: NDVIResult,
    terrain_region,
    weight_terrain: float = 0.6,
    weight_ndvi: float = 0.4,
) -> np.ndarray:
    """
    DEM地形異常スコアとNDVI植生異常スコアを重ね合わせ。

    NDVIは10m解像度、DEMは~8m解像度のため、
    NDVIをDEMグリッドにリサンプリングして統合する。

    Parameters
    ----------
    terrain_score : DEM由来の異常スコア (H_dem, W_dem)
    ndvi_result : NDVIResult
    terrain_region : 地形解析の実際の地理範囲 (GeoRegion)
    weight_terrain : 地形スコアの重み
    weight_ndvi : NDVIスコアの重み

    Returns
    -------
    combined : 統合スコア配列 (same shape as terrain_score)
    """
    from scipy.ndimage import uniform_filter, zoom

    ndvi = ndvi_result.ndvi
    valid_mask = ~np.isnan(ndvi)

    if np.sum(valid_mask) < 100:
        logger.warning("Insufficient NDVI data for combination, using terrain only")
        return terrain_score

    # NDVI Z-score (absolute value = anomaly magnitude)
    ndvi_filled = np.where(valid_mask, ndvi, np.nanmean(ndvi))
    local_mean = uniform_filter(ndvi_filled, size=51)
    local_sq_mean = uniform_filter(ndvi_filled ** 2, size=51)
    local_std = np.sqrt(np.maximum(local_sq_mean - local_mean ** 2, 0))
    local_std = np.maximum(local_std, 1e-6)
    ndvi_zscore = np.abs((ndvi_filled - local_mean) / local_std)
    ndvi_zscore[~valid_mask] = 0

    # Normalize NDVI anomaly to 0-1
    p2, p98 = np.percentile(ndvi_zscore[valid_mask], [2, 98])
    if p98 - p2 > 1e-6:
        ndvi_norm = np.clip((ndvi_zscore - p2) / (p98 - p2), 0, 1)
    else:
        ndvi_norm = np.zeros_like(ndvi_zscore)

    # Resample NDVI to terrain grid size
    target_h, target_w = terrain_score.shape
    zoom_h = target_h / ndvi_norm.shape[0]
    zoom_w = target_w / ndvi_norm.shape[1]
    ndvi_resampled = zoom(ndvi_norm, (zoom_h, zoom_w), order=1)

    # Crop/pad to exact shape
    ndvi_resampled = ndvi_resampled[:target_h, :target_w]
    if ndvi_resampled.shape != terrain_score.shape:
        padded = np.zeros_like(terrain_score)
        h = min(ndvi_resampled.shape[0], target_h)
        w = min(ndvi_resampled.shape[1], target_w)
        padded[:h, :w] = ndvi_resampled[:h, :w]
        ndvi_resampled = padded

    # Normalize terrain score
    t_valid = terrain_score[~np.isnan(terrain_score)]
    if len(t_valid) > 0:
        tp2, tp98 = np.percentile(t_valid, [2, 98])
        if tp98 - tp2 > 1e-6:
            terrain_norm = np.clip((terrain_score - tp2) / (tp98 - tp2), 0, 1)
        else:
            terrain_norm = np.zeros_like(terrain_score)
    else:
        terrain_norm = np.zeros_like(terrain_score)

    combined = weight_terrain * terrain_norm + weight_ndvi * ndvi_resampled
    return combined

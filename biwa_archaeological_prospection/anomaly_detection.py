"""
地形異常検出モジュール

遺跡候補地の自動検出のための統計的手法:
- Z-score閾値法: 局所的な地形パラメータの偏差を検出
- マルチパラメータスコアリング: 複数の地形特徴を統合
- 形状フィルタ: 円形/矩形の人工的構造をブースト
"""

import numpy as np
from scipy import ndimage
from dataclasses import dataclass


@dataclass
class AnomalyCandidate:
    """異常地形候補"""
    row: int
    col: int
    lat: float
    lon: float
    score: float
    area_m2: float
    description: str


def zscore_anomaly(data: np.ndarray, window_size: int = 51) -> np.ndarray:
    """
    局所Z-scoreによる異常度マップ。
    局所平均・標準偏差からの偏差を計算。
    """
    nan_mask = np.isnan(data)
    filled = np.where(nan_mask, np.nanmedian(data), data)

    local_mean = ndimage.uniform_filter(filled, size=window_size, mode="nearest")
    local_sq_mean = ndimage.uniform_filter(filled**2, size=window_size, mode="nearest")
    local_std = np.sqrt(np.clip(local_sq_mean - local_mean**2, 0, None))
    local_std = np.maximum(local_std, 1e-8)

    zscore = (filled - local_mean) / local_std
    zscore[nan_mask] = np.nan
    return zscore


def circular_structure_score(dem: np.ndarray, cell_size: float, radii_m: list[float] = None) -> np.ndarray:
    """
    円形構造物の検出スコア。
    様々な半径のリング状フィルタとの相関を計算。
    古墳や環濠の検出に有効。
    """
    if radii_m is None:
        radii_m = [10, 20, 30, 50, 75, 100]

    nan_mask = np.isnan(dem)
    filled = np.where(nan_mask, np.nanmedian(dem), dem)
    lrm = filled - ndimage.uniform_filter(filled, size=25, mode="nearest")

    best_score = np.zeros_like(dem)

    for radius_m in radii_m:
        radius_px = max(2, int(radius_m / cell_size))
        kernel_size = 2 * radius_px + 1

        y, x = np.ogrid[-radius_px:radius_px + 1, -radius_px:radius_px + 1]
        dist = np.sqrt(x**2 + y**2)

        ring_inner = radius_px * 0.7
        ring_outer = radius_px * 1.3
        ring_mask = ((dist >= ring_inner) & (dist <= ring_outer)).astype(float)
        center_mask = (dist <= radius_px * 0.5).astype(float)

        if ring_mask.sum() == 0 or center_mask.sum() == 0:
            continue

        ring_mask /= ring_mask.sum()
        center_mask /= center_mask.sum()

        ring_mean = ndimage.convolve(lrm, ring_mask, mode="nearest")
        center_mean = ndimage.convolve(lrm, center_mask, mode="nearest")

        score = np.abs(center_mean - ring_mean)
        best_score = np.maximum(best_score, score)

    best_score[nan_mask] = np.nan
    return best_score


def linear_feature_score(dem: np.ndarray, cell_size: float) -> np.ndarray:
    """
    線形構造物の検出スコア。
    道路、堤防、用水路などの直線的な地形変化を検出。
    """
    nan_mask = np.isnan(dem)
    filled = np.where(nan_mask, np.nanmedian(dem), dem)
    lrm = filled - ndimage.uniform_filter(filled, size=25, mode="nearest")

    angles = [0, 45, 90, 135]
    max_response = np.zeros_like(dem)

    for angle in angles:
        rad = np.radians(angle)
        length = max(3, int(30 / cell_size))
        kernel = np.zeros((length, length))
        center = length // 2

        for i in range(length):
            offset = i - center
            ky = center + int(round(offset * np.cos(rad)))
            kx = center + int(round(offset * np.sin(rad)))
            if 0 <= ky < length and 0 <= kx < length:
                kernel[ky, kx] = 1.0

        if kernel.sum() > 0:
            kernel /= kernel.sum()
            response = ndimage.convolve(np.abs(lrm), kernel, mode="nearest")
            max_response = np.maximum(max_response, response)

    max_response[nan_mask] = np.nan
    return max_response


def multi_parameter_score(
    slope_arr: np.ndarray,
    svf_arr: np.ndarray,
    lrm_arr: np.ndarray,
    weights: dict[str, float] = None,
) -> np.ndarray:
    """
    複数の地形パラメータを統合した異常スコア。

    Parameters
    ----------
    slope_arr : 傾斜量
    svf_arr : 天空率
    lrm_arr : 局所起伏モデル
    weights : パラメータ重み
    """
    if weights is None:
        weights = {"slope_z": 0.25, "svf_z": 0.25, "lrm_z": 0.5}

    slope_z = zscore_anomaly(slope_arr)
    svf_z = zscore_anomaly(svf_arr)
    lrm_z = zscore_anomaly(lrm_arr)

    combined = (
        weights["slope_z"] * np.abs(slope_z)
        + weights["svf_z"] * np.abs(svf_z)
        + weights["lrm_z"] * np.abs(lrm_z)
    )

    return combined


def extract_candidates(
    score_map: np.ndarray,
    cell_size: float,
    lat_range: tuple[float, float],
    lon_range: tuple[float, float],
    threshold_percentile: float = 97.5,
    min_area_m2: float = 200.0,
    max_candidates: int = 50,
) -> list[AnomalyCandidate]:
    """
    スコアマップから候補地を抽出。

    Parameters
    ----------
    score_map : 異常スコアマップ
    cell_size : セルサイズ (m)
    lat_range : (lat_min, lat_max)
    lon_range : (lon_min, lon_max)
    threshold_percentile : 閾値パーセンタイル
    min_area_m2 : 最小面積 (m²)
    max_candidates : 最大候補数
    """
    valid_scores = score_map[~np.isnan(score_map)]
    if len(valid_scores) == 0:
        return []

    threshold = np.percentile(valid_scores, threshold_percentile)
    binary = (score_map >= threshold) & ~np.isnan(score_map)

    labeled, n_features = ndimage.label(binary)
    cell_area = cell_size ** 2
    min_pixels = max(1, int(min_area_m2 / cell_area))

    candidates = []
    for i in range(1, n_features + 1):
        mask = labeled == i
        n_pixels = mask.sum()
        if n_pixels < min_pixels:
            continue

        region_scores = score_map[mask]
        mean_score = np.nanmean(region_scores)

        ys, xs = np.where(mask)
        cy, cx = ys.mean(), xs.mean()

        rows, cols = score_map.shape
        lat = lat_range[1] - (cy / rows) * (lat_range[1] - lat_range[0])
        lon = lon_range[0] + (cx / cols) * (lon_range[1] - lon_range[0])

        area_m2 = n_pixels * cell_area

        candidates.append(AnomalyCandidate(
            row=int(cy), col=int(cx),
            lat=lat, lon=lon,
            score=float(mean_score),
            area_m2=area_m2,
            description=f"Anomaly cluster ({n_pixels} px, {area_m2:.0f} m²)",
        ))

    candidates.sort(key=lambda c: c.score, reverse=True)
    return candidates[:max_candidates]

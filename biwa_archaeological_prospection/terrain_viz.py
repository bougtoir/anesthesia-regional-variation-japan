"""
地形可視化モジュール

以下の手法を実装:
- Hillshade (陰影起伏図)
- Slope (傾斜量図)
- Red Relief Image Map (RRIM / 赤色立体地図, Chiba et al. 2008)
- Sky-View Factor (SVF / 天空率)
- Local Relief Model (LRM / 局所起伏モデル)
- Multi-directional Hillshade (多方向陰影)
"""

import numpy as np
from scipy import ndimage


def compute_gradient(dem: np.ndarray, cell_size: float) -> tuple[np.ndarray, np.ndarray]:
    """DEM勾配をx, y方向に計算 (Hornメソッド)"""
    pad = np.pad(dem, 1, mode="edge")
    dz_dx = (
        (pad[:-2, 2:] + 2 * pad[1:-1, 2:] + pad[2:, 2:])
        - (pad[:-2, :-2] + 2 * pad[1:-1, :-2] + pad[2:, :-2])
    ) / (8 * cell_size)
    dz_dy = (
        (pad[2:, :-2] + 2 * pad[2:, 1:-1] + pad[2:, 2:])
        - (pad[:-2, :-2] + 2 * pad[:-2, 1:-1] + pad[:-2, 2:])
    ) / (8 * cell_size)
    return dz_dx, dz_dy


def hillshade(
    dem: np.ndarray, cell_size: float,
    azimuth: float = 315.0, altitude: float = 45.0,
) -> np.ndarray:
    """
    陰影起伏図を計算。

    Parameters
    ----------
    dem : 標高配列
    cell_size : セルサイズ (m)
    azimuth : 光源方位角 (度, 北=0, 時計回り)
    altitude : 光源仰角 (度)
    """
    dz_dx, dz_dy = compute_gradient(dem, cell_size)
    az_rad = np.radians(360 - azimuth + 90)
    alt_rad = np.radians(altitude)
    shade = (
        np.sin(alt_rad)
        + np.cos(alt_rad) * (dz_dx * np.cos(az_rad) + dz_dy * np.sin(az_rad))
    ) / np.sqrt(1 + dz_dx**2 + dz_dy**2)
    return np.clip(shade, 0, 1)


def multi_hillshade(dem: np.ndarray, cell_size: float, n_directions: int = 8) -> np.ndarray:
    """多方向陰影の合成"""
    azimuths = np.linspace(0, 360, n_directions, endpoint=False)
    shades = [hillshade(dem, cell_size, az, 45.0) for az in azimuths]
    return np.mean(shades, axis=0)


def slope(dem: np.ndarray, cell_size: float, degrees: bool = True) -> np.ndarray:
    """傾斜量を計算"""
    dz_dx, dz_dy = compute_gradient(dem, cell_size)
    slope_rad = np.arctan(np.sqrt(dz_dx**2 + dz_dy**2))
    if degrees:
        return np.degrees(slope_rad)
    return slope_rad


def _positive_openness_at_direction(
    dem: np.ndarray, cell_size: float, direction_rad: float, max_radius_cells: int,
) -> np.ndarray:
    """特定方向の正の開放度を計算"""
    rows, cols = dem.shape
    openness = np.full_like(dem, 90.0)

    dy = np.cos(direction_rad)
    dx = np.sin(direction_rad)

    for r in range(1, max_radius_cells + 1):
        oy = int(round(r * dy))
        ox = int(round(r * dx))
        if oy == 0 and ox == 0:
            continue
        shifted = np.full_like(dem, np.nan)
        src_y_start = max(0, -oy)
        src_y_end = rows - max(0, oy)
        src_x_start = max(0, -ox)
        src_x_end = cols - max(0, ox)
        dst_y_start = max(0, oy)
        dst_y_end = rows + min(0, oy)
        dst_x_start = max(0, ox)
        dst_x_end = cols + min(0, ox)
        shifted[dst_y_start:dst_y_end, dst_x_start:dst_x_end] = (
            dem[src_y_start:src_y_end, src_x_start:src_x_end]
        )
        dist = r * cell_size
        elev_angle = np.degrees(np.arctan2(shifted - dem, dist))
        valid = ~np.isnan(elev_angle)
        openness[valid] = np.minimum(openness[valid], 90.0 - elev_angle[valid])

    return openness


def sky_view_factor(
    dem: np.ndarray, cell_size: float,
    n_directions: int = 16, max_radius_m: float = 100.0,
) -> np.ndarray:
    """
    天空率 (SVF) を計算。

    Parameters
    ----------
    n_directions : 方向数
    max_radius_m : 探索半径 (m)
    """
    max_radius_cells = max(1, int(max_radius_m / cell_size))
    directions = np.linspace(0, 2 * np.pi, n_directions, endpoint=False)
    openness_sum = np.zeros_like(dem, dtype=float)

    for d in directions:
        op = _positive_openness_at_direction(dem, cell_size, d, max_radius_cells)
        openness_sum += np.sin(np.radians(op))

    svf = openness_sum / n_directions
    return np.clip(svf, 0, 1)


def local_relief_model(dem: np.ndarray, kernel_size: int = 25) -> np.ndarray:
    """
    局所起伏モデル (LRM)。
    大域的トレンドを平滑化で除去し、局所的な起伏のみを抽出。

    Parameters
    ----------
    kernel_size : 平滑化カーネルサイズ
    """
    smoothed = ndimage.uniform_filter(dem, size=kernel_size, mode="nearest")
    lrm = dem - smoothed
    return lrm


def rrim(
    dem: np.ndarray, cell_size: float,
    svf_directions: int = 16, svf_radius_m: float = 100.0,
) -> np.ndarray:
    """
    赤色立体地図 (RRIM: Red Relief Image Map)。
    Chiba T., Kaneta S. & Suzuki Y. (2008) の手法に基づく。

    傾斜量を赤の彩度に、天空率を明度にマッピング。

    Returns
    -------
    rrim_rgb : (H, W, 3) uint8 RGB配列
    """
    slp = slope(dem, cell_size, degrees=True)
    svf = sky_view_factor(dem, cell_size, svf_directions, svf_radius_m)

    slp_norm = np.clip(slp / 50.0, 0, 1)
    svf_norm = np.clip(svf, 0, 1)

    r = np.clip((1.0 - svf_norm) * 255 + slp_norm * 200, 0, 255)
    g = np.clip(svf_norm * 200 * (1.0 - slp_norm * 0.5), 0, 255)
    b = np.clip(svf_norm * 200 * (1.0 - slp_norm * 0.8), 0, 255)

    rgb = np.stack([r, g, b], axis=-1).astype(np.uint8)
    return rgb


def compute_all_visualizations(
    dem: np.ndarray, cell_size: float,
    svf_directions: int = 16, svf_radius_m: float = 100.0,
    lrm_kernel: int = 25,
) -> dict[str, np.ndarray]:
    """全可視化手法を一括計算"""
    nan_mask = np.isnan(dem)
    dem_filled = np.where(nan_mask, np.nanmedian(dem), dem)

    results = {
        "hillshade": hillshade(dem_filled, cell_size),
        "multi_hillshade": multi_hillshade(dem_filled, cell_size),
        "slope": slope(dem_filled, cell_size),
        "svf": sky_view_factor(dem_filled, cell_size, svf_directions, svf_radius_m),
        "lrm": local_relief_model(dem_filled, lrm_kernel),
        "rrim": rrim(dem_filled, cell_size, svf_directions, svf_radius_m),
    }

    for key in results:
        if results[key].ndim == 2:
            results[key] = np.where(nan_mask, np.nan, results[key])

    return results

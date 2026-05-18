"""
琵琶湖考古学的探査メインパイプライン

実行フロー:
1. 調査対象地域の定義（現琵琶湖湖底 + 古琵琶湖各ステージ湖畔域）
2. 国土地理院DEMタイルの取得・結合
3. 地形可視化（Hillshade, Slope, RRIM, SVF, LRM）
4. 異常地形検出（Z-score, 円形構造, マルチパラメータ統合）
5. 候補地抽出・ランキング
6. 結果の可視化・レポート出力
"""

import logging
import os
from dataclasses import dataclass

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.gridspec import GridSpec

plt.rcParams["font.family"] = ["IPAGothic", "DejaVu Sans"]

from .data_fetch import GeoRegion, fetch_dem_region, compute_cell_size_meters
from .terrain_viz import compute_all_visualizations
from .anomaly_detection import (
    zscore_anomaly,
    circular_structure_score,
    linear_feature_score,
    multi_parameter_score,
    extract_candidates,
    AnomalyCandidate,
)
from .paleo_biwa import (
    PALEO_LAKE_STAGES,
    KNOWN_SITES,
    get_priority_regions,
    BIWA_SOUTH_LAKEBED,
    AWAZU_LAKEBED_SITE,
)
from .vegetation_analysis import (
    fetch_ndvi_for_region,
    detect_vegetation_anomalies,
    compute_combined_score,
    NDVIResult,
    VegetationAnomaly,
)

logger = logging.getLogger(__name__)


@dataclass
class RegionResult:
    """1地域の解析結果"""
    region: GeoRegion
    actual_region: GeoRegion
    dem: np.ndarray
    visualizations: dict[str, np.ndarray]
    anomaly_score: np.ndarray
    circular_score: np.ndarray
    linear_score: np.ndarray
    candidates: list[AnomalyCandidate]
    cell_size: float
    ndvi_result: NDVIResult = None
    vegetation_anomalies: list[VegetationAnomaly] = None
    combined_terrain_ndvi_score: np.ndarray = None


def analyze_region(
    region: GeoRegion,
    zoom: int = 14,
    output_dir: str = "output",
) -> RegionResult:
    """
    1地域の完全解析パイプラインを実行。

    Parameters
    ----------
    region : 解析対象地域
    zoom : DEMタイルのズームレベル
    output_dir : 出力ディレクトリ
    """
    logger.info("=== Analyzing region: %s ===", region.name)

    dem, actual_region = fetch_dem_region(region, z=zoom)
    cell_size = compute_cell_size_meters(actual_region.center_lat, zoom)
    logger.info("Cell size: %.2f m", cell_size)

    logger.info("Computing terrain visualizations...")
    svf_radius = min(100.0, cell_size * 15)
    viz = compute_all_visualizations(
        dem, cell_size,
        svf_directions=16, svf_radius_m=svf_radius,
        lrm_kernel=max(5, int(200 / cell_size)),
    )

    logger.info("Running anomaly detection...")
    anomaly_score = multi_parameter_score(viz["slope"], viz["svf"], viz["lrm"])
    circ_score = circular_structure_score(dem, cell_size)
    lin_score = linear_feature_score(dem, cell_size)

    combined = (
        0.5 * _normalize(anomaly_score)
        + 0.3 * _normalize(circ_score)
        + 0.2 * _normalize(lin_score)
    )

    logger.info("Extracting candidates...")
    candidates = extract_candidates(
        combined, cell_size,
        lat_range=(actual_region.lat_min, actual_region.lat_max),
        lon_range=(actual_region.lon_min, actual_region.lon_max),
    )
    logger.info("Found %d candidates", len(candidates))

    # NDVI vegetation analysis
    ndvi_result = None
    veg_anomalies = None
    combined_terrain_ndvi = None
    try:
        logger.info("Fetching Sentinel-2 NDVI data...")
        ndvi_result = fetch_ndvi_for_region(region)
        if ndvi_result is not None:
            logger.info("Detecting vegetation anomalies...")
            veg_anomalies = detect_vegetation_anomalies(
                ndvi_result.ndvi, ndvi_result.bbox, ndvi_result.resolution_m,
            )
            logger.info("Computing combined terrain+NDVI score...")
            combined_terrain_ndvi = compute_combined_score(
                combined, ndvi_result, actual_region,
            )
        else:
            logger.info("No NDVI data available, using terrain analysis only")
    except Exception as e:
        logger.warning("NDVI analysis failed (non-fatal): %s", e)

    result = RegionResult(
        region=region,
        actual_region=actual_region,
        dem=dem,
        visualizations=viz,
        anomaly_score=combined,
        circular_score=circ_score,
        linear_score=lin_score,
        candidates=candidates,
        cell_size=cell_size,
        ndvi_result=ndvi_result,
        vegetation_anomalies=veg_anomalies or [],
        combined_terrain_ndvi_score=combined_terrain_ndvi,
    )

    return result


def _normalize(arr: np.ndarray) -> np.ndarray:
    """0-1正規化"""
    valid = arr[~np.isnan(arr)]
    if len(valid) == 0:
        return np.zeros_like(arr)
    vmin, vmax = np.nanpercentile(valid, [2, 98])
    if vmax - vmin < 1e-10:
        return np.zeros_like(arr)
    normalized = (arr - vmin) / (vmax - vmin)
    return np.clip(normalized, 0, 1)


def plot_region_results(result: RegionResult, output_dir: str) -> list[str]:
    """解析結果の可視化図を生成"""
    os.makedirs(output_dir, exist_ok=True)
    saved_files = []
    region_slug = result.region.name.replace(" ", "_").replace("/", "_")[:40]
    extent = [
        result.actual_region.lon_min, result.actual_region.lon_max,
        result.actual_region.lat_min, result.actual_region.lat_max,
    ]

    has_ndvi = result.ndvi_result is not None
    n_rows = 5 if has_ndvi else 4
    fig = plt.figure(figsize=(20, 6 * n_rows))
    gs = GridSpec(n_rows, 3, figure=fig, hspace=0.25, wspace=0.2)
    fig.suptitle(
        f"Archaeological Prospection: {result.region.name}\n"
        f"({result.actual_region.lat_min:.3f}–{result.actual_region.lat_max:.3f}°N, "
        f"{result.actual_region.lon_min:.3f}–{result.actual_region.lon_max:.3f}°E)",
        fontsize=14, fontweight="bold",
    )

    panels = [
        ("DEM (Elevation)", result.dem, "terrain", gs[0, 0]),
        ("Hillshade", result.visualizations["hillshade"], "gray", gs[0, 1]),
        ("Multi-directional Hillshade", result.visualizations["multi_hillshade"], "gray", gs[0, 2]),
        ("Slope (degrees)", result.visualizations["slope"], "YlOrRd", gs[1, 0]),
        ("Sky-View Factor (SVF)", result.visualizations["svf"], "gray_r", gs[1, 1]),
        ("Local Relief Model (LRM)", result.visualizations["lrm"], "RdBu_r", gs[1, 2]),
        ("Anomaly Score (Multi-param)", result.anomaly_score, "hot", gs[2, 0]),
        ("Circular Structure Score", result.circular_score, "magma", gs[2, 1]),
        ("Linear Feature Score", result.linear_score, "inferno", gs[2, 2]),
    ]

    for title, data, cmap, gs_pos in panels:
        ax = fig.add_subplot(gs_pos)
        if data.ndim == 2:
            vmin, vmax = np.nanpercentile(data[~np.isnan(data)], [2, 98]) if np.any(~np.isnan(data)) else (0, 1)
            im = ax.imshow(data, extent=extent, cmap=cmap, vmin=vmin, vmax=vmax, aspect="auto")
            plt.colorbar(im, ax=ax, shrink=0.7)
        ax.set_title(title, fontsize=10)
        ax.set_xlabel("Longitude")
        ax.set_ylabel("Latitude")

    ax_rrim = fig.add_subplot(gs[3, 0])
    ax_rrim.imshow(result.visualizations["rrim"], extent=extent, aspect="auto")
    ax_rrim.set_title("Red Relief Image Map (RRIM)", fontsize=10)
    ax_rrim.set_xlabel("Longitude")
    ax_rrim.set_ylabel("Latitude")

    ax_cand = fig.add_subplot(gs[3, 1:])
    ax_cand.imshow(
        result.visualizations["hillshade"], extent=extent,
        cmap="gray", aspect="auto", alpha=0.5,
    )
    score_masked = np.ma.masked_where(
        result.anomaly_score < np.nanpercentile(result.anomaly_score[~np.isnan(result.anomaly_score)], 95),
        result.anomaly_score,
    )
    ax_cand.imshow(score_masked, extent=extent, cmap="Reds", alpha=0.7, aspect="auto")

    for i, cand in enumerate(result.candidates[:20]):
        ax_cand.plot(cand.lon, cand.lat, "c^", markersize=8, markeredgecolor="k", linewidth=0.5)
        ax_cand.annotate(
            f"#{i+1}", (cand.lon, cand.lat),
            fontsize=6, color="cyan", fontweight="bold",
            textcoords="offset points", xytext=(4, 4),
        )

    for site in KNOWN_SITES:
        if (result.actual_region.lat_min <= site["lat"] <= result.actual_region.lat_max
                and result.actual_region.lon_min <= site["lon"] <= result.actual_region.lon_max):
            ax_cand.plot(site["lon"], site["lat"], "r*", markersize=15, markeredgecolor="k")
            ax_cand.annotate(
                site["name_ja"], (site["lon"], site["lat"]),
                fontsize=8, color="red", fontweight="bold",
                textcoords="offset points", xytext=(6, 6),
            )

    ax_cand.set_title("Candidate Anomalies (top 20) + Known Sites", fontsize=10)
    ax_cand.set_xlabel("Longitude")
    ax_cand.set_ylabel("Latitude")

    # Row 5: NDVI panels (if available)
    if has_ndvi:
        ndvi = result.ndvi_result.ndvi
        ndvi_bbox = result.ndvi_result.bbox
        ndvi_extent = [ndvi_bbox[0], ndvi_bbox[2], ndvi_bbox[1], ndvi_bbox[3]]

        # Panel 1: NDVI map
        ax_ndvi = fig.add_subplot(gs[4, 0])
        valid = ndvi[~np.isnan(ndvi)]
        if len(valid) > 0:
            vmin_n, vmax_n = np.nanpercentile(valid, [2, 98])
        else:
            vmin_n, vmax_n = -1, 1
        im_ndvi = ax_ndvi.imshow(
            ndvi, extent=ndvi_extent, cmap="RdYlGn", vmin=vmin_n, vmax=vmax_n, aspect="auto",
        )
        plt.colorbar(im_ndvi, ax=ax_ndvi, shrink=0.7)
        ax_ndvi.set_title(
            f"NDVI (Sentinel-2, {result.ndvi_result.scene_date[:10]})\n"
            f"clouds={result.ndvi_result.cloud_cover:.1f}%",
            fontsize=10,
        )
        ax_ndvi.set_xlabel("Longitude")
        ax_ndvi.set_ylabel("Latitude")

        # Panel 2: Combined terrain+NDVI score
        ax_combo = fig.add_subplot(gs[4, 1])
        if result.combined_terrain_ndvi_score is not None:
            combo = result.combined_terrain_ndvi_score
            valid_c = combo[~np.isnan(combo)]
            if len(valid_c) > 0:
                vc_min, vc_max = np.nanpercentile(valid_c, [2, 98])
            else:
                vc_min, vc_max = 0, 1
            im_combo = ax_combo.imshow(
                combo, extent=extent, cmap="hot", vmin=vc_min, vmax=vc_max, aspect="auto",
            )
            plt.colorbar(im_combo, ax=ax_combo, shrink=0.7)
            ax_combo.set_title("Combined Score\n(60% Terrain + 40% NDVI)", fontsize=10)
        else:
            ax_combo.text(0.5, 0.5, "N/A", ha="center", va="center", transform=ax_combo.transAxes)
            ax_combo.set_title("Combined Score (N/A)", fontsize=10)
        ax_combo.set_xlabel("Longitude")
        ax_combo.set_ylabel("Latitude")

        # Panel 3: Vegetation anomalies on hillshade
        ax_veg = fig.add_subplot(gs[4, 2])
        ax_veg.imshow(
            result.visualizations["hillshade"], extent=extent,
            cmap="gray", aspect="auto", alpha=0.5,
        )
        if result.vegetation_anomalies:
            for va in result.vegetation_anomalies[:30]:
                marker = "^" if va.anomaly_type == "high_ndvi" else "v"
                color = "green" if va.anomaly_type == "high_ndvi" else "orange"
                ax_veg.plot(va.lon, va.lat, marker, color=color, markersize=7, markeredgecolor="k", linewidth=0.5)
            n_high = sum(1 for v in result.vegetation_anomalies if v.anomaly_type == "high_ndvi")
            n_low = sum(1 for v in result.vegetation_anomalies if v.anomaly_type == "low_ndvi")
            ax_veg.set_title(
                f"Vegetation Anomalies ({len(result.vegetation_anomalies)} total)\n"
                f"▲ High NDVI ({n_high}) ▼ Low NDVI ({n_low})",
                fontsize=10,
            )
        else:
            ax_veg.set_title("Vegetation Anomalies (none detected)", fontsize=10)
        ax_veg.set_xlabel("Longitude")
        ax_veg.set_ylabel("Latitude")

    path = os.path.join(output_dir, f"analysis_{region_slug}.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    saved_files.append(path)
    logger.info("Saved: %s", path)

    return saved_files


def plot_paleo_biwa_overview(output_dir: str) -> str:
    """古琵琶湖変遷の概観図を生成"""
    os.makedirs(output_dir, exist_ok=True)

    fig, ax = plt.subplots(figsize=(12, 14))
    ax.set_title(
        "古琵琶湖の変遷と調査対象地域\n"
        "Paleo-Lake Biwa Migration & Survey Regions",
        fontsize=14, fontweight="bold",
    )

    colors = plt.cm.viridis(np.linspace(0.2, 0.9, len(PALEO_LAKE_STAGES)))

    for i, stage in enumerate(PALEO_LAKE_STAGES):
        r = stage.survey_region
        rect = plt.Rectangle(
            (r.lon_min, r.lat_min),
            r.lon_max - r.lon_min,
            r.lat_max - r.lat_min,
            linewidth=2, edgecolor=colors[i], facecolor=colors[i],
            alpha=0.3, label=f"{stage.name_ja} ({stage.age_ma_start}–{stage.age_ma_end} Ma)",
        )
        ax.add_patch(rect)
        ax.annotate(
            f"{stage.name_ja}\n{stage.age_ma_start}–{stage.age_ma_end} Ma",
            (stage.center_lon, stage.center_lat),
            fontsize=9, ha="center", va="center", fontweight="bold",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8),
        )

    from .paleo_biwa import CURRENT_BIWA
    biwa_rect = plt.Rectangle(
        (CURRENT_BIWA.lon_min, CURRENT_BIWA.lat_min),
        CURRENT_BIWA.lon_max - CURRENT_BIWA.lon_min,
        CURRENT_BIWA.lat_max - CURRENT_BIWA.lat_min,
        linewidth=3, edgecolor="blue", facecolor="lightblue",
        alpha=0.2, label="現琵琶湖 (Current Lake Biwa)",
    )
    ax.add_patch(biwa_rect)

    for site in KNOWN_SITES:
        ax.plot(site["lon"], site["lat"], "r*", markersize=15, markeredgecolor="k")
        ax.annotate(
            site["name_ja"], (site["lon"], site["lat"]),
            fontsize=8, color="red", fontweight="bold",
            textcoords="offset points", xytext=(8, 8),
        )

    for stage in PALEO_LAKE_STAGES[:-1]:
        next_idx = PALEO_LAKE_STAGES.index(stage) + 1
        next_stage = PALEO_LAKE_STAGES[next_idx]
        ax.annotate(
            "", xy=(next_stage.center_lon, next_stage.center_lat),
            xytext=(stage.center_lon, stage.center_lat),
            arrowprops=dict(arrowstyle="->", color="gray", lw=2),
        )

    ax.set_xlabel("Longitude (°E)", fontsize=12)
    ax.set_ylabel("Latitude (°N)", fontsize=12)
    ax.set_xlim(135.75, 136.40)
    ax.set_ylim(34.65, 35.60)
    ax.legend(loc="upper left", fontsize=8)
    ax.grid(True, alpha=0.3)
    ax.set_aspect("equal")

    path = os.path.join(output_dir, "paleo_biwa_overview.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    logger.info("Saved: %s", path)
    return path


def generate_report(results: list[RegionResult], output_dir: str) -> str:
    """Markdown形式の調査レポートを生成"""
    os.makedirs(output_dir, exist_ok=True)

    lines = [
        "# 琵琶湖考古学的探査レポート",
        "# Lake Biwa Archaeological Prospection Report",
        "",
        "## 概要 / Overview",
        "",
        "本レポートは、国土地理院の公開地形データ（DEM）を用いた考古学的探査の結果をまとめたものです。",
        "地形可視化技術（RRIM、SVF、LRM等）と統計的異常検出を組み合わせ、",
        "遺跡候補地の自動スクリーニングを行いました。",
        "",
        "### 使用データ",
        "- 国土地理院 標高タイル (DEM 5m/10m)",
        "- Sentinel-2 L2A 衛星画像 (Band4 Red + Band8 NIR → NDVI, 10m解像度)",
        "- 産総研 シームレス地質図V2 (古琵琶湖層群の分布確認用)",
        "",
        "### 使用手法",
        "| 手法 | 略称 | 目的 |",
        "|------|------|------|",
        "| 赤色立体地図 | RRIM | 微地形の総合的可視化 (千葉ほか, 2008) |",
        "| 天空率 | SVF | 凹凸の方向非依存検出 |",
        "| 傾斜量図 | Slope | 急傾斜部の強調 |",
        "| 局所起伏モデル | LRM | 大域トレンド除去後の局所起伏 |",
        "| 多方向陰影 | Multi-HS | 微地形の陰影強調 |",
        "| 円形構造スコア | Circ | 古墳・環濠等の円形パターン検出 |",
        "| 線形構造スコア | Linear | 道路・堤防等の線形パターン検出 |",
        "| Z-score異常度 | Z-score | 統計的に有意な地形偏差 |",
        "| 植生指標 | NDVI | 正規化植生指数 (Sentinel-2由来) |",
        "| 植生異常検出 | Veg-Z | NDVIのZ-scoreによる植生異常検出 |",
        "| 統合スコア | Combined | 地形異常(60%) + 植生異常(40%)の重ね合わせ |",
        "",
        "---",
        "",
        "## 古琵琶湖の変遷 / Paleo-Lake Biwa Migration",
        "",
        "琵琶湖は約400万年前に三重県伊賀地方で誕生し、北へ移動して現在の位置に至りました。",
        "",
        "| ステージ | 累層名 | 年代 (Ma) | 位置 | 備考 |",
        "|---------|--------|-----------|------|------|",
    ]

    for stage in PALEO_LAKE_STAGES:
        lines.append(
            f"| {stage.name_ja} | {stage.formation_ja} | "
            f"{stage.age_ma_start}–{stage.age_ma_end} | "
            f"{stage.center_lat:.2f}°N, {stage.center_lon:.2f}°E | "
            f"{stage.description_ja[:40]}... |"
        )

    lines.extend([
        "",
        "![古琵琶湖変遷図](paleo_biwa_overview.png)",
        "",
        "---",
        "",
        "## 地域別解析結果 / Regional Analysis Results",
        "",
    ])

    for result in results:
        region_slug = result.region.name.replace(" ", "_").replace("/", "_")[:40]
        lines.extend([
            f"### {result.region.name}",
            "",
            f"- DEM解像度: {result.cell_size:.1f} m/pixel",
            f"- DEM形状: {result.dem.shape[0]} x {result.dem.shape[1]} pixels",
            f"- 有効ピクセル率: {np.count_nonzero(~np.isnan(result.dem)) / result.dem.size * 100:.1f}%",
            f"- 検出候補数: {len(result.candidates)}",
            "",
            f"![解析結果](analysis_{region_slug}.png)",
            "",
        ])

        if result.candidates:
            lines.extend([
                "#### 候補地一覧 (上位10件)",
                "",
                "| # | 緯度 | 経度 | スコア | 面積 (m²) | 備考 |",
                "|---|------|------|--------|-----------|------|",
            ])
            for i, c in enumerate(result.candidates[:10]):
                lines.append(
                    f"| {i+1} | {c.lat:.5f} | {c.lon:.5f} | "
                    f"{c.score:.3f} | {c.area_m2:.0f} | {c.description} |"
                )
            lines.append("")

        # Vegetation analysis results
        if result.ndvi_result is not None:
            ndvi = result.ndvi_result
            lines.extend([
                "#### 植生解析 (NDVI)",
                "",
                f"- 使用シーン: {ndvi.scene_date[:10]}",
                f"- 雲量: {ndvi.cloud_cover:.1f}%",
                f"- 解像度: {ndvi.resolution_m:.0f}m",
                f"- NDVI範囲: [{np.nanmin(ndvi.ndvi):.3f}, {np.nanmax(ndvi.ndvi):.3f}]",
                f"- NDVI平均: {np.nanmean(ndvi.ndvi):.3f}",
                f"- 植生異常検出数: {len(result.vegetation_anomalies)}",
                "",
            ])
            if result.vegetation_anomalies:
                n_high = sum(1 for v in result.vegetation_anomalies if v.anomaly_type == "high_ndvi")
                n_low = sum(1 for v in result.vegetation_anomalies if v.anomaly_type == "low_ndvi")
                lines.extend([
                    f"  - 高NDVI異常 (植生が異常に活発): {n_high}件 — 栽培植物/人為的植栽の可能性",
                    f"  - 低NDVI異常 (植生が異常に希薄): {n_low}件 — 埋没遺構による植生抑制の可能性",
                    "",
                    "| # | 緒度 | 経度 | NDVI | Z-score | 面積(m²) | タイプ |",
                    "|---|------|------|------|---------|-----------|------|",
                ])
                for i, va in enumerate(result.vegetation_anomalies[:10]):
                    atype = "高NDVI" if va.anomaly_type == "high_ndvi" else "低NDVI"
                    lines.append(
                        f"| {i+1} | {va.lat:.5f} | {va.lon:.5f} | "
                        f"{va.ndvi_value:.3f} | {va.ndvi_zscore:.1f} | "
                        f"{va.area_m2:.0f} | {atype} |"
                    )
                lines.append("")

    lines.extend([
        "---",
        "",
        "## 考察 / Discussion",
        "",
        "### 手法の有効性",
        "- RRIM（赤色立体地図）は微地形の総合的把握に最も有効",
        "- SVF（天空率）は凹地形（堀、溝）の検出に優れる",
        "- LRM（局所起伏モデル）は大域的な傾斜に埋もれた微小な起伏を抽出",
        "- 円形構造スコアは古墳・環濠集落の検出に特化",
        "",
        "### 解像度の制約",
        "- 国土地理院公開DEM（5-10m解像度）では直径20m以上の構造物が検出限界",
        "- UAV-LiDAR（0.5-2m解像度）を取得すれば、本パイプラインで段階2の詳細解析が可能",
        "",
        "### 古琵琶湖湖畔域の考古学的ポテンシャル",
        "- 各ステージの湖畔域は当時の水辺の生活圏であり、遺跡存在の蓋然性が高い",
        "- 特に蒲生累層分布域は更新世の化石産地としても重要",
        "- 堅田累層付近は現琵琶湖と近く、水位変動による水没遺跡の可能性あり",
        "",
        "### 植生解析の意義",
        "- NDVI（正規化植生指数）により、地形異常だけでは検出できない遺跡痕跡を補完",
        "- 高NDVI異常: 古代の栽培植物（クリ、クルミ等の縄文時代植物）の子孫が野生化している可能性",
        "- 低NDVI異常: 地下の石構造物・埋没遺構が植生の成長を抑制している可能性",
        "- 古琵琶湖各ステージの湖畔域で、狩猟採集から栽培生活への移行の痕跡を植生パターンから探索",
        "",
        "### 次のステップ",
        "1. 候補地の現地踏査（地表面確認）",
        "2. UAV-LiDARによる高解像度DEM取得",
        "3. 地中レーダー(GPR)による埋没遺構確認",
        "4. シームレス地質図との重ね合わせによる古琵琶湖層群分布の精密化",
        "5. 季節ごとのNDVI変動解析（植生の年周期パターンと遺跡痕跡の相関）",
        "",
        "---",
        "",
        "## 参考文献 / References",
        "",
        "1. Chiba T., Kaneta S. & Suzuki Y. (2008) Red Relief Image Map: New visualisation "
        "for three dimensional data. ISPRS.",
        "2. 里口保文 (2025) 琵琶湖博物館研究調査報告 38巻.",
        "3. Štular B., Lozić E. & Eichert S. (2021) Airborne LiDAR-Derived Digital Elevation "
        "Model for Archaeology. Remote Sensing, 13(9), 1855.",
        "4. Zakšek K., Oštir K. & Kokalj Ž. (2011) Sky-View Factor as a Relief Visualization "
        "Technique. Remote Sensing, 3(2), 398-415.",
        "5. 横山卓雄・雨森清 (1991) 滋賀県湖東地域古琵琶湖層群地質図.",
        "6. Kokalj Ž. & Somrak M. (2019) Why Not a Single Image? Combining Visualizations "
        "to Facilitate Fieldwork and On-Screen Mapping. Remote Sensing, 11(7), 747.",
        "",
    ])

    path = os.path.join(output_dir, "report.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    logger.info("Report saved: %s", path)
    return path


def run_pipeline(
    output_dir: str = "output",
    zoom: int = 14,
    regions: list[GeoRegion] = None,
) -> tuple[list[RegionResult], list[str]]:
    """
    完全なパイプラインを実行。

    Parameters
    ----------
    output_dir : 出力ディレクトリ
    zoom : DEMタイルのズームレベル (14 = ~10m解像度)
    regions : 解析対象地域リスト (None = デフォルトの優先地域)
    """
    if regions is None:
        regions = get_priority_regions()

    all_results = []
    all_files = []

    overview_path = plot_paleo_biwa_overview(output_dir)
    all_files.append(overview_path)

    for region in regions:
        result = analyze_region(region, zoom=zoom, output_dir=output_dir)
        all_results.append(result)

        fig_paths = plot_region_results(result, output_dir)
        all_files.extend(fig_paths)

    report_path = generate_report(all_results, output_dir)
    all_files.append(report_path)

    logger.info("Pipeline complete. %d files generated.", len(all_files))
    return all_results, all_files

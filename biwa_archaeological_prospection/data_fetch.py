"""
国土地理院タイルデータ取得モジュール

データソース:
- 標高タイル (DEM): https://cyberjapandata.gsi.go.jp/xyz/dem/{z}/{x}/{y}.txt
- 湖水深タイル: https://cyberjapandata.gsi.go.jp/xyz/lakedepth/{z}/{x}/{y}.png
- シームレス地質図V2: https://gbank.gsj.jp/seamless/v2/api/1.2.1/tiles/{z}/{y}/{x}.png
- 地質凡例API: https://gbank.gsj.jp/seamless/v2/api/1.2.1/legend.json
"""

import io
import math
import time
import logging
from dataclasses import dataclass
from typing import Optional

import numpy as np
import pandas as pd
import requests
from PIL import Image

logger = logging.getLogger(__name__)

GSI_DEM_URL = "https://cyberjapandata.gsi.go.jp/xyz/dem/{z}/{x}/{y}.txt"
GSI_DEM5A_URL = "https://cyberjapandata.gsi.go.jp/xyz/dem5a_png/{z}/{x}/{y}.png"
GSI_LAKEDEPTH_URL = "https://cyberjapandata.gsi.go.jp/xyz/lakedepth/{z}/{x}/{y}.png"
GSJ_GEOLOGY_TILE_URL = "https://gbank.gsj.jp/seamless/v2/api/1.2.1/tiles/{z}/{y}/{x}.png"
GSJ_LEGEND_URL = "https://gbank.gsj.jp/seamless/v2/api/1.2.1/legend.json"

TILE_SIZE = 256


@dataclass
class TileBounds:
    """タイル座標系の矩形範囲"""
    z: int
    x_min: int
    x_max: int
    y_min: int
    y_max: int

    @property
    def nx(self) -> int:
        return self.x_max - self.x_min + 1

    @property
    def ny(self) -> int:
        return self.y_max - self.y_min + 1


@dataclass
class GeoRegion:
    """地理的矩形範囲 (WGS84)"""
    lat_min: float
    lat_max: float
    lon_min: float
    lon_max: float
    name: str = ""

    @property
    def center_lat(self) -> float:
        return (self.lat_min + self.lat_max) / 2

    @property
    def center_lon(self) -> float:
        return (self.lon_min + self.lon_max) / 2


def latlon_to_tile(lat: float, lon: float, z: int) -> tuple[int, int]:
    """緯度経度からタイル座標 (x, y) に変換"""
    n = 2 ** z
    x = int((lon + 180.0) / 360.0 * n)
    lat_rad = math.radians(lat)
    y = int((1.0 - math.log(math.tan(lat_rad) + 1.0 / math.cos(lat_rad)) / math.pi) / 2.0 * n)
    return x, y


def tile_to_latlon(x: int, y: int, z: int) -> tuple[float, float]:
    """タイル座標の北西角の緯度経度を返す"""
    n = 2 ** z
    lon = x / n * 360.0 - 180.0
    lat_rad = math.atan(math.sinh(math.pi * (1 - 2 * y / n)))
    lat = math.degrees(lat_rad)
    return lat, lon


def region_to_tiles(region: GeoRegion, z: int) -> TileBounds:
    """GeoRegionをタイル座標範囲に変換"""
    x_min, y_min = latlon_to_tile(region.lat_max, region.lon_min, z)
    x_max, y_max = latlon_to_tile(region.lat_min, region.lon_max, z)
    return TileBounds(z=z, x_min=x_min, x_max=x_max, y_min=y_min, y_max=y_max)


def fetch_dem_tile_txt(z: int, x: int, y: int, retries: int = 3) -> Optional[np.ndarray]:
    """標高タイル (テキスト形式) を取得し 256x256 numpy配列で返す"""
    url = GSI_DEM_URL.format(z=z, x=x, y=y)
    for attempt in range(retries):
        try:
            resp = requests.get(url, timeout=15)
            if resp.status_code == 404:
                return None
            resp.raise_for_status()
            df = pd.read_csv(io.StringIO(resp.text), header=None)
            arr = df.replace("e", np.nan).values.astype(float)
            return arr
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(1.0 * (attempt + 1))
                continue
            logger.warning("DEM tile fetch failed z=%d x=%d y=%d: %s", z, x, y, e)
            return None


def fetch_dem5a_png_tile(z: int, x: int, y: int, retries: int = 3) -> Optional[np.ndarray]:
    """DEM5A PNG標高タイルを取得し標高値に変換"""
    url = GSI_DEM5A_URL.format(z=z, x=x, y=y)
    for attempt in range(retries):
        try:
            resp = requests.get(url, timeout=15)
            if resp.status_code == 404:
                return None
            resp.raise_for_status()
            img = Image.open(io.BytesIO(resp.content)).convert("RGBA")
            arr = np.array(img)
            r, g, b = arr[:, :, 0].astype(float), arr[:, :, 1].astype(float), arr[:, :, 2].astype(float)
            a = arr[:, :, 3]
            elev = r * 256 + g + b / 256 - 32768
            elev[a == 0] = np.nan
            return elev
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(1.0 * (attempt + 1))
                continue
            logger.warning("DEM5A PNG tile fetch failed z=%d x=%d y=%d: %s", z, x, y, e)
            return None


def fetch_dem_region(region: GeoRegion, z: int = 14, use_png: bool = False) -> tuple[np.ndarray, GeoRegion]:
    """
    指定領域のDEMを取得し、結合した配列と実際の地理範囲を返す。

    Parameters
    ----------
    region : GeoRegion
    z : int
        ズームレベル (14推奨: ~10m解像度)
    use_png : bool
        True の場合 DEM5A PNG形式を使用 (より高解像度だがズーム15限定)

    Returns
    -------
    dem : np.ndarray
        結合DEM配列
    actual_region : GeoRegion
        タイル境界に基づく実際の地理範囲
    """
    bounds = region_to_tiles(region, z)
    logger.info(
        "Fetching DEM tiles: z=%d, x=[%d,%d], y=[%d,%d] (%d tiles)",
        z, bounds.x_min, bounds.x_max, bounds.y_min, bounds.y_max,
        bounds.nx * bounds.ny,
    )

    fetch_fn = fetch_dem5a_png_tile if use_png else fetch_dem_tile_txt

    rows = []
    for ty in range(bounds.y_min, bounds.y_max + 1):
        row_tiles = []
        for tx in range(bounds.x_min, bounds.x_max + 1):
            tile = fetch_fn(z, tx, ty)
            if tile is None:
                tile = np.full((TILE_SIZE, TILE_SIZE), np.nan)
            row_tiles.append(tile)
        rows.append(np.concatenate(row_tiles, axis=1))

    dem = np.concatenate(rows, axis=0)

    nw_lat, nw_lon = tile_to_latlon(bounds.x_min, bounds.y_min, z)
    se_lat, se_lon = tile_to_latlon(bounds.x_max + 1, bounds.y_max + 1, z)
    actual_region = GeoRegion(
        lat_min=se_lat, lat_max=nw_lat,
        lon_min=nw_lon, lon_max=se_lon,
        name=region.name,
    )

    valid_pct = np.count_nonzero(~np.isnan(dem)) / dem.size * 100
    logger.info("DEM shape: %s, valid pixels: %.1f%%", dem.shape, valid_pct)

    return dem, actual_region


def fetch_geology_legend(lat: float, lon: float) -> Optional[dict]:
    """指定座標の地質凡例情報を取得"""
    url = GSJ_LEGEND_URL
    params = {"lat": lat, "lon": lon}
    try:
        resp = requests.get(url, params=params, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        return data if data else None
    except Exception as e:
        logger.warning("Geology legend fetch failed: %s", e)
        return None


def fetch_geology_tile(z: int, x: int, y: int) -> Optional[np.ndarray]:
    """シームレス地質図V2タイル画像を取得"""
    url = GSJ_GEOLOGY_TILE_URL.format(z=z, x=x, y=y)
    try:
        resp = requests.get(url, timeout=15)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        img = Image.open(io.BytesIO(resp.content)).convert("RGBA")
        return np.array(img)
    except Exception as e:
        logger.warning("Geology tile fetch failed: %s", e)
        return None


def compute_cell_size_meters(lat: float, z: int) -> float:
    """タイルのセルサイズ（メートル）を計算"""
    equator_circumference = 40075016.686
    meters_per_pixel = equator_circumference * math.cos(math.radians(lat)) / (2 ** z * TILE_SIZE)
    return meters_per_pixel

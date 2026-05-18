"""
古琵琶湖変遷データモジュール

琵琶湖は約400万年前に三重県伊賀地方で誕生し、北へ移動して現在の位置に至った。
古琵琶湖層群の各累層が示す湖の変遷ステージと、対応する地理的位置を管理する。

参考文献:
- 里口保文 (2025) 琵琶湖博物館研究調査報告 38巻
- 川辺孝幸・吉川周作 (1992) Sedimentary History of the Paleo-Lake Biwa
- 横山卓雄・雨森清 (1991) 滋賀県湖東地域古琵琶湖層群地質図
- 滋賀県 琵琶湖の概要
"""

from dataclasses import dataclass, field

from .data_fetch import GeoRegion


@dataclass
class PaleoLakeStage:
    """古琵琶湖の各変遷ステージ"""
    name_ja: str
    name_en: str
    formation_ja: str
    formation_en: str
    age_ma_start: float
    age_ma_end: float
    center_lat: float
    center_lon: float
    approximate_extent_km2: float
    description_ja: str
    description_en: str
    survey_region: GeoRegion = field(default=None)

    def __post_init__(self):
        if self.survey_region is None:
            dlat = (self.approximate_extent_km2 ** 0.5) / 111.0 / 2
            dlon = dlat / 0.8
            self.survey_region = GeoRegion(
                lat_min=self.center_lat - dlat,
                lat_max=self.center_lat + dlat,
                lon_min=self.center_lon - dlon,
                lon_max=self.center_lon + dlon,
                name=self.name_en,
            )


PALEO_LAKE_STAGES: list[PaleoLakeStage] = [
    PaleoLakeStage(
        name_ja="大山田湖",
        name_en="Oyamada Lake",
        formation_ja="上野累層",
        formation_en="Ueno Formation",
        age_ma_start=4.0,
        age_ma_end=3.5,
        center_lat=34.75,
        center_lon=136.15,
        approximate_extent_km2=30,
        description_ja=(
            "古琵琶湖最初期の浅い湖。三重県大山田村付近に位置。"
            "服部川河床から上野層群が露出し、ミエゾウ・ワニの化石が産出。"
            "断層運動により形成され、流送土砂により埋積。亜熱帯気候。"
        ),
        description_en=(
            "Earliest shallow paleo-lake near Oyamada, Mie Prefecture. "
            "Ueno Formation exposed in Hattori River bed. Fossils of Mie elephant "
            "and crocodile found. Formed by faulting, filled by sediment transport."
        ),
    ),
    PaleoLakeStage(
        name_ja="阿山湖",
        name_en="Ayama Lake",
        formation_ja="阿山累層",
        formation_en="Ayama Formation",
        age_ma_start=3.0,
        age_ma_end=2.6,
        center_lat=34.82,
        center_lon=136.12,
        approximate_extent_km2=80,
        description_ja=(
            "大山田湖から北へ移動した第二の古琵琶湖。"
            "伊賀市から甲賀市にかけて広がる、古琵琶湖層群期で最大の湖。"
            "古瀬田川を通じて古奈良湖と接続した時期もあった。"
        ),
        description_en=(
            "Second paleo-lake, shifted northward from Oyamada. "
            "Extended from Iga City to Koka City, the largest lake during "
            "the Kobiwako Group period. Connected to Paleo-Nara Lake at times."
        ),
    ),
    PaleoLakeStage(
        name_ja="甲賀湖",
        name_en="Koka Lake",
        formation_ja="甲賀累層",
        formation_en="Koka Formation",
        age_ma_start=2.6,
        age_ma_end=2.0,
        center_lat=34.92,
        center_lon=136.17,
        approximate_extent_km2=60,
        description_ja=(
            "深い湖として存在。滋賀県甲賀市付近に位置。"
            "土地の隆起により消滅。"
        ),
        description_en=(
            "Deep lake near Koka City, Shiga Prefecture. "
            "Disappeared due to tectonic uplift."
        ),
    ),
    PaleoLakeStage(
        name_ja="蒲生沼沢地",
        name_en="Gamo Marshland",
        formation_ja="蒲生累層",
        formation_en="Gamo Formation",
        age_ma_start=2.6,
        age_ma_end=1.0,
        center_lat=35.05,
        center_lon=136.15,
        approximate_extent_km2=50,
        description_ja=(
            "滋賀県蒲生郡付近に形成された沼沢地。"
            "メタセコイア林が広がり、愛知川・野洲川河床に化石林が残存。"
            "ゾウの化石も産出。干上がって消滅。"
        ),
        description_en=(
            "Marshland near Gamo, Shiga Prefecture. "
            "Metasequoia forests spread; fossil forests remain in "
            "Echi and Yasu River beds. Elephant fossils found. Dried up."
        ),
    ),
    PaleoLakeStage(
        name_ja="堅田湖",
        name_en="Katata Lake",
        formation_ja="堅田累層",
        formation_en="Katata Formation",
        age_ma_start=1.0,
        age_ma_end=0.4,
        center_lat=35.10,
        center_lon=135.92,
        approximate_extent_km2=40,
        description_ja=(
            "堅田付近に形成された小さな湖。"
            "約40万年前の隆起と断層運動により、現在の琵琶湖の位置に移行。"
        ),
        description_en=(
            "Small lake near Katata. Tectonic uplift ~0.4 Ma ago "
            "led to transition to the current Lake Biwa location."
        ),
    ),
]


CURRENT_BIWA = GeoRegion(
    lat_min=34.90,
    lat_max=35.52,
    lon_min=135.85,
    lon_max=136.30,
    name="Current Lake Biwa",
)

BIWA_SOUTH_LAKEBED = GeoRegion(
    lat_min=34.95,
    lat_max=35.10,
    lon_min=135.85,
    lon_max=136.00,
    name="Lake Biwa South Basin (Lakebed)",
)

BIWA_NORTH_LAKEBED = GeoRegion(
    lat_min=35.15,
    lat_max=35.45,
    lon_min=135.95,
    lon_max=136.25,
    name="Lake Biwa North Basin (Lakebed)",
)

AWAZU_LAKEBED_SITE = GeoRegion(
    lat_min=34.977,
    lat_max=34.992,
    lon_min=135.893,
    lon_max=135.910,
    name="Awazu Lakebed Site (粟津湖底遺跡)",
)


KNOWN_SITES: list[dict] = [
    {
        "name_ja": "粟津湖底遺跡",
        "name_en": "Awazu Lakebed Site",
        "lat": 34.9853,
        "lon": 135.9003,
        "period": "縄文時代早期〜中期 (約10,000〜5,000年前)",
        "description": "日本最大の淡水貝塚。水深2-3mの湖底に位置。1952年発見。",
    },
    {
        "name_ja": "赤野井湾遺跡群",
        "name_en": "Akanoi Bay Sites",
        "lat": 35.07,
        "lon": 136.00,
        "period": "弥生〜古墳時代",
        "description": "琵琶湖南部の赤野井湾周辺の水没遺跡群。",
    },
]


def get_all_survey_regions() -> list[GeoRegion]:
    """全調査対象地域のリストを返す"""
    regions = []
    for stage in PALEO_LAKE_STAGES:
        regions.append(stage.survey_region)
    regions.extend([BIWA_SOUTH_LAKEBED, AWAZU_LAKEBED_SITE])
    return regions


def get_priority_regions() -> list[GeoRegion]:
    """優先調査地域 (デモ用にコンパクトな範囲) を返す"""
    return [
        GeoRegion(
            lat_min=34.96, lat_max=35.00,
            lon_min=135.88, lon_max=135.93,
            name="Awazu area (粟津周辺, 既知遺跡あり)",
        ),
        GeoRegion(
            lat_min=34.73, lat_max=34.77,
            lon_min=136.12, lon_max=136.18,
            name="Oyamada area (大山田, 最古の古琵琶湖)",
        ),
        GeoRegion(
            lat_min=34.90, lat_max=34.94,
            lon_min=136.14, lon_max=136.20,
            name="Koka area (甲賀, 古琵琶湖層群分布域)",
        ),
    ]

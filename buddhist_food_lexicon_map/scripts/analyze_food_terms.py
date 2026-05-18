#!/usr/bin/env python3
"""Extract food-term frequencies from open Buddhist text datasets."""

from __future__ import annotations

import csv
import io
import json
import math
import re
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
FIGURE_DIR = OUTPUT_DIR / "figures"
CACHE_DIR = DATA_DIR / "cache"
CBETA_CACHE_DIR = CACHE_DIR / "cbeta"
BILARA_CACHE_DIR = CACHE_DIR / "bilara"
BASE_CBETA = "https://cbdata.dila.edu.tw/stable"
BASE_BILARA_RAW = "https://raw.githubusercontent.com/suttacentral/bilara-data/published"
BILARA_TREE_API = "https://api.github.com/repos/suttacentral/bilara-data/git/trees/published?recursive=1"
REFERER = "https://app.devin.ai/"

CHINESE_VINAYA_WORKS = [
    {
        "work_id": "T1421",
        "title_zh": "彌沙塞部和醯五分律",
        "title_en": "Mahīśāsaka Vinaya (Five-Part Vinaya)",
        "school": "Mahīśāsaka",
        "indic_association": "northwest/central India (trad. Mahīśāsaka transmission)",
        "cbeta_note": "Chinese translation metadata is from CBETA /works; Indic association is contextual, not an API geocode.",
    },
    {
        "work_id": "T1425",
        "title_zh": "摩訶僧祇律",
        "title_en": "Mahāsāṃghika Vinaya",
        "school": "Mahāsāṃghika",
        "indic_association": "Andhra / broader early Mahāsāṃghika transmission",
        "cbeta_note": "Chinese translation metadata is from CBETA /works; Indic association is contextual, not an API geocode.",
    },
    {
        "work_id": "T1428",
        "title_zh": "四分律",
        "title_en": "Dharmaguptaka Vinaya (Four-Part Vinaya)",
        "school": "Dharmaguptaka",
        "indic_association": "Gandhāra / northwest India (Dharmaguptaka transmission)",
        "cbeta_note": "Chinese translation metadata is from CBETA /works; Indic association is contextual, not an API geocode.",
    },
    {
        "work_id": "T1435",
        "title_zh": "十誦律",
        "title_en": "Sarvāstivāda Vinaya (Ten-Recitation Vinaya)",
        "school": "Sarvāstivāda",
        "indic_association": "Kashmir / northwest India (Sarvāstivāda transmission)",
        "cbeta_note": "Chinese translation metadata is from CBETA /works; Indic association is contextual, not an API geocode.",
    },
    {
        "work_id": "T1442",
        "title_zh": "根本說一切有部毘奈耶",
        "title_en": "Mūlasarvāstivāda Vinaya",
        "school": "Mūlasarvāstivāda",
        "indic_association": "north India / Kashmir-linked Mūlasarvāstivāda transmission",
        "cbeta_note": "Chinese translation metadata is from CBETA /works; Indic association is contextual, not an API geocode.",
    },
]

CHINESE_TERMS = [
    {"term": "粥", "lemma_en": "gruel/congee", "category": "grain/preparation"},
    {"term": "乳", "lemma_en": "milk", "category": "dairy"},
    {"term": "酥", "lemma_en": "ghee/butter", "category": "dairy"},
    {"term": "蜜", "lemma_en": "honey", "category": "sweetener/medicine"},
    {"term": "飯", "lemma_en": "cooked rice/meal", "category": "grain/preparation"},
    {"term": "米", "lemma_en": "rice", "category": "grain"},
    {"term": "麥", "lemma_en": "wheat/barley", "category": "grain"},
    {"term": "肉", "lemma_en": "meat", "category": "animal food"},
    {"term": "魚", "lemma_en": "fish", "category": "animal food"},
    {"term": "餅", "lemma_en": "cake/bread", "category": "grain/preparation"},
    {"term": "油", "lemma_en": "oil", "category": "fat/medicine"},
    {"term": "鹽", "lemma_en": "salt", "category": "seasoning"},
    {"term": "薑", "lemma_en": "ginger", "category": "seasoning/medicine"},
    {"term": "糖", "lemma_en": "sugar", "category": "sweetener/medicine"},
    {"term": "漿", "lemma_en": "drink/syrup", "category": "drink"},
]

PALI_TERMS = [
    {"term": "yāgu", "lemma_en": "gruel/congee", "category": "grain/preparation"},
    {"term": "khīra", "lemma_en": "milk", "category": "dairy"},
    {"term": "sappi", "lemma_en": "ghee/butter", "category": "dairy"},
    {"term": "madhu", "lemma_en": "honey", "category": "sweetener/medicine"},
    {"term": "odana", "lemma_en": "cooked rice/meal", "category": "grain/preparation"},
    {"term": "taṇḍula", "lemma_en": "rice", "category": "grain"},
    {"term": "godhūma", "lemma_en": "wheat/barley", "category": "grain"},
    {"term": "maṃsa", "lemma_en": "meat", "category": "animal food"},
    {"term": "maccha", "lemma_en": "fish", "category": "animal food"},
    {"term": "pūva", "lemma_en": "cake/bread", "category": "grain/preparation"},
    {"term": "tela", "lemma_en": "oil", "category": "fat/medicine"},
    {"term": "loṇa", "lemma_en": "salt", "category": "seasoning"},
    {"term": "siṅgivera", "lemma_en": "ginger", "category": "seasoning/medicine"},
    {"term": "phāṇita", "lemma_en": "sugar", "category": "sweetener/medicine"},
    {"term": "pāna", "lemma_en": "drink/syrup", "category": "drink"},
]

FOOD_COLOR = {
    "animal food": "#d62728",
    "dairy": "#1f77b4",
    "drink": "#17becf",
    "fat/medicine": "#9467bd",
    "grain": "#8c564b",
    "grain/preparation": "#ff7f0e",
    "seasoning": "#7f7f7f",
    "seasoning/medicine": "#bcbd22",
    "sweetener/medicine": "#2ca02c",
}

ROMAN_WORD_RE = re.compile(r"[A-Za-zāīūṅñṭḍṇḷṃĀĪŪṄÑṬḌṆḶṂ]+")


@dataclass(frozen=True)
class WorkMetadata:
    work_id: str
    tradition: str
    title: str
    school: str
    source_dataset: str
    source_url: str
    source_text_path: str
    formation_from: int | None
    formation_to: int | None
    translation_dynasty: str | None
    translation_place: str | None
    latitude: float | None
    longitude: float | None
    region_note: str
    text_length_chars: int


def request_bytes(url: str, *, timeout: int = 60) -> bytes:
    request = urllib.request.Request(
        url,
        headers={
            "Referer": REFERER,
            "User-Agent": "Devin Buddhist food lexicon map/0.1",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def cached_bytes(path: Path, url: str, *, timeout: int = 60) -> bytes:
    if path.exists():
        return path.read_bytes()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = request_bytes(url, timeout=timeout)
    path.write_bytes(data)
    return data


def cached_json(path: Path, url: str, *, timeout: int = 60) -> dict:
    return json.loads(cached_bytes(path, url, timeout=timeout).decode("utf-8"))


def clean_cbeta_text(text: str) -> str:
    lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped.startswith("No. "):
            continue
        lines.append(stripped)
    return "\n".join(lines)


def download_cbeta_work(work_id: str) -> tuple[str, list[str]]:
    url = f"{BASE_CBETA}/download/text/{work_id}.txt.zip"
    zip_path = CBETA_CACHE_DIR / f"{work_id}.txt.zip"
    data = cached_bytes(zip_path, url)
    text_parts = []
    names = []
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for name in sorted(archive.namelist()):
            if not name.endswith(".txt") or name.endswith("-toc.txt"):
                continue
            names.append(name)
            text_parts.append(clean_cbeta_text(archive.read(name).decode("utf-8", errors="ignore")))
    return "\n".join(text_parts), names


def get_cbeta_metadata(work: dict) -> WorkMetadata:
    work_id = work["work_id"]
    encoded = urllib.parse.quote(work_id)
    url = f"{BASE_CBETA}/works?work={encoded}"
    meta = cached_json(CBETA_CACHE_DIR / f"{work_id}.works.json", url)
    result = meta["results"][0]
    places = result.get("places") or []
    primary_place = places[0] if places else {}
    text, names = download_cbeta_work(work_id)
    return WorkMetadata(
        work_id=work_id,
        tradition="Chinese Vinaya translation",
        title=f"{work['title_zh']} / {work['title_en']}",
        school=work["school"],
        source_dataset="CBETA API 2025R3",
        source_url=f"{BASE_CBETA}/download/text/{work_id}.txt.zip",
        source_text_path=";".join(names),
        formation_from=result.get("time_from"),
        formation_to=result.get("time_to"),
        translation_dynasty=result.get("time_dynasty"),
        translation_place=primary_place.get("name"),
        latitude=primary_place.get("latitude"),
        longitude=primary_place.get("longitude"),
        region_note=f"{work['indic_association']}. {work['cbeta_note']}",
        text_length_chars=len(text),
    )


def count_chinese_terms(text: str, metadata: WorkMetadata) -> list[dict]:
    rows = []
    for item in CHINESE_TERMS:
        count = text.count(item["term"])
        rows.append(
            {
                **metadata.__dict__,
                "script_language": "Classical Chinese",
                "term": item["term"],
                "lemma_en": item["lemma_en"],
                "category": item["category"],
                "count": count,
                "per_10k_chars": (count / metadata.text_length_chars) * 10000 if metadata.text_length_chars else 0,
                "sample_segment_id": "",
                "sample_text": "",
            }
        )
    return rows


def list_bilara_root_paths() -> list[str]:
    tree = cached_json(BILARA_CACHE_DIR / "published_tree.json", BILARA_TREE_API, timeout=90)
    paths = [
        item["path"]
        for item in tree["tree"]
        if item["path"].startswith("root/pli/ms/vinaya/") and item["path"].endswith(".json")
    ]
    return sorted(paths)


def bilara_raw_json(path: str) -> dict[str, str]:
    cache_name = path.replace("/", "__")
    url = f"{BASE_BILARA_RAW}/{path}"
    return cached_json(BILARA_CACHE_DIR / cache_name, url)


def count_pali_terms(paths: Iterable[str]) -> tuple[list[dict], WorkMetadata]:
    counts = {item["term"]: 0 for item in PALI_TERMS}
    samples = {item["term"]: ("", "") for item in PALI_TERMS}
    total_chars = 0
    file_count = 0

    for path in paths:
        file_count += 1
        data = bilara_raw_json(path)
        for segment_id, segment_text in data.items():
            text = segment_text.casefold()
            total_chars += len(segment_text)
            tokens = ROMAN_WORD_RE.findall(text)
            token_counts = {}
            for token in tokens:
                token_counts[token] = token_counts.get(token, 0) + 1
            for term in counts:
                count = token_counts.get(term.casefold(), 0)
                if count:
                    counts[term] += count
                    if not samples[term][0]:
                        samples[term] = (segment_id, segment_text)

    metadata = WorkMetadata(
        work_id="pli-tv",
        tradition="Pāli Vinaya root text",
        title="Theravāda Vinayapiṭaka (SuttaCentral Bilara root text)",
        school="Theravāda",
        source_dataset="SuttaCentral bilara-data published branch",
        source_url="https://github.com/suttacentral/bilara-data/tree/published/root/pli/ms/vinaya",
        source_text_path=f"{file_count} root/pli/ms/vinaya JSON files",
        formation_from=None,
        formation_to=None,
        translation_dynasty=None,
        translation_place="Sri Lanka (textual redaction/preservation convention)",
        latitude=7.8731,
        longitude=80.7718,
        region_note="Pāli Vinaya root text; coordinates are country centroid for Sri Lanka and should not be read as a precise composition site.",
        text_length_chars=total_chars,
    )
    rows = []
    for item in PALI_TERMS:
        term = item["term"]
        sample_id, sample_text = samples[term]
        rows.append(
            {
                **metadata.__dict__,
                "script_language": "Pāli (Romanized)",
                "term": term,
                "lemma_en": item["lemma_en"],
                "category": item["category"],
                "count": counts[term],
                "per_10k_chars": (counts[term] / total_chars) * 10000 if total_chars else 0,
                "sample_segment_id": sample_id,
                "sample_text": sample_text,
            }
        )
    return rows, metadata


def extract_cbeta_context(work_id: str, term: str, window: int = 32, limit: int = 3) -> list[str]:
    text, _ = download_cbeta_work(work_id)
    contexts = []
    for match in re.finditer(re.escape(term), text):
        start = max(0, match.start() - window)
        end = min(len(text), match.end() + window)
        contexts.append(text[start:end].replace("\n", ""))
        if len(contexts) >= limit:
            break
    return contexts


def save_csv(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def save_metadata(metadata: list[WorkMetadata]) -> None:
    rows = [m.__dict__ for m in metadata]
    save_csv(
        OUTPUT_DIR / "work_metadata.csv",
        rows,
        [
            "work_id",
            "tradition",
            "title",
            "school",
            "source_dataset",
            "source_url",
            "source_text_path",
            "formation_from",
            "formation_to",
            "translation_dynasty",
            "translation_place",
            "latitude",
            "longitude",
            "region_note",
            "text_length_chars",
        ],
    )


def make_heatmap(df: pd.DataFrame) -> None:
    pivot = df.pivot_table(index="title", columns="lemma_en", values="per_10k_chars", aggfunc="sum").fillna(0)
    ordered_terms = [item["lemma_en"] for item in CHINESE_TERMS]
    pivot = pivot[[term for term in ordered_terms if term in pivot.columns]]
    fig_height = max(5.0, 0.45 * len(pivot.index) + 1.6)
    fig_width = max(10.0, 0.55 * len(pivot.columns) + 4.0)
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    matrix = pivot.to_numpy()
    image = ax.imshow(matrix, aspect="auto", cmap="YlOrRd")
    ax.set_xticks(range(len(pivot.columns)))
    ax.set_xticklabels(pivot.columns, rotation=45, ha="right")
    ax.set_yticks(range(len(pivot.index)))
    ax.set_yticklabels([label.split(" / ")[-1] if " / " in label else label for label in pivot.index])
    ax.set_title("Food-term frequency by Vinaya corpus")
    ax.set_xlabel("Food lexeme (English lemma)")
    ax.set_ylabel("Corpus")
    cbar = fig.colorbar(image, ax=ax)
    cbar.set_label("mentions per 10,000 characters")
    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            value = matrix[i, j]
            if value > 0:
                ax.text(j, i, f"{value:.1f}", ha="center", va="center", fontsize=7)
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "food_terms_heatmap.png", dpi=200)
    plt.close(fig)


def make_map_html(df: pd.DataFrame) -> None:
    rows = []
    grouped = df.groupby(["work_id", "title", "school", "latitude", "longitude", "translation_place", "region_note"], dropna=False)
    for keys, group in grouped:
        work_id, title, school, latitude, longitude, place, note = keys
        if pd.isna(latitude) or pd.isna(longitude):
            continue
        top_terms = group.sort_values("count", ascending=False).head(8)
        radius = 8 + math.sqrt(float(group["count"].sum()))
        term_list = "".join(
            f"<li>{row.term} / {row.lemma_en}: {int(row.count)} ({row.per_10k_chars:.2f} per 10k chars)</li>"
            for row in top_terms.itertuples(index=False)
        )
        popup = f"""
        <strong>{work_id}: {title}</strong><br>
        School: {school}<br>
        CBETA/SuttaCentral location: {place}<br>
        <em>{note}</em>
        <ul>{term_list}</ul>
        """
        rows.append(
            f"L.circleMarker([{float(latitude):.6f}, {float(longitude):.6f}], "
            f"{{radius: {radius:.2f}, color: '#7b3294', fillColor: '#c2a5cf', fillOpacity: 0.65}})"
            f".addTo(map).bindPopup({json.dumps(popup, ensure_ascii=False)});"
        )
    html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Buddhist food-term geography prototype</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
  <style>
    body {{ margin: 0; font-family: system-ui, -apple-system, sans-serif; }}
    #map {{ height: 100vh; }}
    .note {{ position: absolute; z-index: 999; top: 12px; left: 56px; max-width: 430px; background: rgba(255,255,255,0.92); padding: 10px 12px; border-radius: 8px; box-shadow: 0 1px 6px rgba(0,0,0,0.25); font-size: 13px; }}
  </style>
</head>
<body>
  <div class="note"><strong>Prototype map.</strong> Marker locations are translation/redaction metadata from CBETA and a Sri Lanka centroid for the Pāli Vinaya. They are evidence anchors, not exact composition sites. Marker size follows total food-term mentions.</div>
  <div id="map"></div>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script>
    const map = L.map('map').setView([23, 91], 4);
    L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
      maxZoom: 18,
      attribution: '&copy; OpenStreetMap contributors'
    }}).addTo(map);
    {chr(10).join(rows)}
  </script>
</body>
</html>
"""
    (FIGURE_DIR / "food_terms_map.html").write_text(html, encoding="utf-8")


def make_category_bars(df: pd.DataFrame) -> None:
    summary = (
        df.groupby(["school", "category"], as_index=False)["per_10k_chars"]
        .sum()
        .sort_values(["school", "category"])
    )
    pivot = summary.pivot_table(index="school", columns="category", values="per_10k_chars", aggfunc="sum").fillna(0)
    columns = sorted(pivot.columns)
    pivot = pivot[columns]
    fig, ax = plt.subplots(figsize=(11, 6))
    bottom = [0.0] * len(pivot.index)
    x = range(len(pivot.index))
    for category in columns:
        values = pivot[category].to_list()
        ax.bar(x, values, bottom=bottom, label=category, color=FOOD_COLOR.get(category))
        bottom = [a + b for a, b in zip(bottom, values)]
    ax.set_xticks(list(x))
    ax.set_xticklabels(pivot.index, rotation=30, ha="right")
    ax.set_ylabel("mentions per 10,000 characters")
    ax.set_title("Food-term category profile by Vinaya lineage")
    ax.legend(loc="upper left", bbox_to_anchor=(1.0, 1.0), fontsize=8)
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "food_category_profile.png", dpi=200)
    plt.close(fig)


def main() -> int:
    for directory in [DATA_DIR, OUTPUT_DIR, FIGURE_DIR, CACHE_DIR, CBETA_CACHE_DIR, BILARA_CACHE_DIR]:
        directory.mkdir(parents=True, exist_ok=True)

    all_rows = []
    metadata = []
    context_rows = []

    for work in CHINESE_VINAYA_WORKS:
        work_id = work["work_id"]
        print(f"Fetching CBETA {work_id}...")
        meta = get_cbeta_metadata(work)
        text, _ = download_cbeta_work(work_id)
        metadata.append(meta)
        all_rows.extend(count_chinese_terms(text, meta))
        for term in ["粥", "乳", "酥", "蜜", "飯", "肉", "魚"]:
            for context in extract_cbeta_context(work_id, term, limit=2):
                context_rows.append({"work_id": work_id, "term": term, "context": context})

    print("Fetching SuttaCentral Bilara Pāli Vinaya tree...")
    pali_paths = list_bilara_root_paths()
    pali_rows, pali_meta = count_pali_terms(pali_paths)
    all_rows.extend(pali_rows)
    metadata.append(pali_meta)

    fieldnames = [
        "work_id",
        "tradition",
        "title",
        "school",
        "source_dataset",
        "source_url",
        "source_text_path",
        "formation_from",
        "formation_to",
        "translation_dynasty",
        "translation_place",
        "latitude",
        "longitude",
        "region_note",
        "text_length_chars",
        "script_language",
        "term",
        "lemma_en",
        "category",
        "count",
        "per_10k_chars",
        "sample_segment_id",
        "sample_text",
    ]
    save_csv(OUTPUT_DIR / "food_term_counts.csv", all_rows, fieldnames)
    save_metadata(metadata)
    save_csv(OUTPUT_DIR / "cbeta_context_samples.csv", context_rows, ["work_id", "term", "context"])

    df = pd.DataFrame(all_rows)
    make_heatmap(df)
    make_category_bars(df)
    make_map_html(df)

    summary = df.groupby(["work_id", "school"], as_index=False).agg(
        total_food_mentions=("count", "sum"),
        text_length_chars=("text_length_chars", "first"),
        total_per_10k_chars=("per_10k_chars", "sum"),
    )
    summary.to_csv(OUTPUT_DIR / "corpus_summary.csv", index=False)
    print(f"Wrote {OUTPUT_DIR}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

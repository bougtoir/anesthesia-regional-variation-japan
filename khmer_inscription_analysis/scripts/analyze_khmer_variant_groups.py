#!/usr/bin/env python3
"""Aggregate Khmer inscription vocabulary using spelling variants and near-synonyms."""

from __future__ import annotations

import csv
import sys
import time
from pathlib import Path
from typing import Iterable

import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import pandas as pd

from analyze_khmer_inscriptions import corpus_url, fetch, parse_distribution, plain_text


VARIANT_GROUPS: tuple[dict[str, str], ...] = (
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "’avalokiteśvara"},
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "lokeśvara"},
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "mahādivyalokeśvara"},
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "paramadivyalokeśvara"},
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "ratnalokeśvara"},
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "raṇadivyalokeśvara"},
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "samaradivyalokeśvara"},
    {"group": "大乗系語彙", "family": "観音・ロケーシュヴァラ系", "term": "sarvvalokeśvara"},
    {"group": "大乗系語彙", "family": "ローカナータ系", "term": "caturlokanātha"},
    {"group": "大乗系語彙", "family": "菩薩・菩提系", "term": "bodhisattva"},
    {"group": "大乗系語彙", "family": "菩薩・菩提系", "term": "bodhisambhāra"},
    {"group": "大乗系語彙", "family": "菩薩・菩提系", "term": "mahābodhi"},
    {"group": "大乗系語彙", "family": "金剛系", "term": "vajra"},
    {"group": "大乗系語彙", "family": "金剛系", "term": "ratnavajra"},
    {"group": "大乗系語彙", "family": "金剛系", "term": "somavajra"},
    {"group": "大乗系語彙", "family": "金剛系", "term": "tribhuvanavajra"},
    {"group": "大乗系語彙", "family": "金剛系", "term": "tribhūvanavajra"},
    {"group": "大乗系語彙", "family": "金剛系", "term": "vajrabheda"},
    {"group": "上座部・僧団系語彙", "family": "長老系", "term": "thera"},
    {"group": "上座部・僧団系語彙", "family": "長老系", "term": "mahāthera"},
    {"group": "上座部・僧団系語彙", "family": "長老系", "term": "therānuthera"},
    {"group": "上座部・僧団系語彙", "family": "僧団・比丘系", "term": "bhikṣu"},
    {"group": "上座部・僧団系語彙", "family": "僧団・比丘系", "term": "bhikṣusaṅgha"},
    {"group": "上座部・僧団系語彙", "family": "僧団・比丘系", "term": "bhiksusaṅgha"},
    {"group": "上座部・僧団系語彙", "family": "僧団・比丘系", "term": "saṅgha"},
    {"group": "上座部・僧団系語彙", "family": "僧団・比丘系", "term": "saṅgharāja"},
    {"group": "上座部・僧団系語彙", "family": "仏教制度・伝承系", "term": "sāsana"},
    {"group": "上座部・僧団系語彙", "family": "仏教制度・伝承系", "term": "buddhasāsa"},
    {"group": "上座部・僧団系語彙", "family": "仏教制度・伝承系", "term": "saṅgāyanā"},
    {"group": "上座部・僧団系語彙", "family": "仏教制度・伝承系", "term": "muggaliputtatissa"},
    {"group": "上座部・僧団系語彙", "family": "仏教制度・伝承系", "term": "dharmmāśoka"},
    {"group": "仏教一般語彙", "family": "仏・塔", "term": "buddha"},
    {"group": "仏教一般語彙", "family": "仏・塔", "term": "stūpa"},
)


def century(year: int) -> int:
    return ((year - 1) // 100) + 1


def write_csv(path: Path, fieldnames: Iterable[str], rows: Iterable[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(fieldnames))
        writer.writeheader()
        writer.writerows(rows)


def build_variant_outputs(output_dir: Path, pause: float) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    distribution_rows: list[dict[str, object]] = []
    term_rows: list[dict[str, object]] = []
    error_rows: list[dict[str, object]] = []
    for spec in VARIANT_GROUPS:
        term = spec["term"]
        print(f"Querying variant {term}...", file=sys.stderr)
        try:
            text = plain_text(fetch(corpus_url(term, search=" distribution ")))
            distribution = parse_distribution(text)
            total = sum(distribution.values())
            term_rows.append({**spec, "total_count": total, "years": ";".join(str(y) for y in sorted(distribution))})
            for year, count in sorted(distribution.items()):
                distribution_rows.append(
                    {
                        **spec,
                        "year": year,
                        "century": century(year),
                        "count": count,
                    }
                )
        except Exception as error:  # noqa: BLE001
            error_rows.append({"term": term, "error": repr(error)})
        time.sleep(pause)

    distribution = pd.DataFrame(
        distribution_rows,
        columns=["group", "family", "term", "year", "century", "count"],
    )
    terms = pd.DataFrame(
        term_rows,
        columns=["group", "family", "term", "total_count", "years"],
    )
    summary = (
        distribution.groupby("group", as_index=False)["count"].sum().sort_values("count", ascending=False)
        if not distribution.empty
        else pd.DataFrame(columns=["group", "count"])
    )
    output_dir.mkdir(parents=True, exist_ok=True)
    distribution.to_csv(output_dir / "variant_group_distribution.csv", index=False)
    terms.to_csv(output_dir / "variant_group_terms.csv", index=False)
    summary.to_csv(output_dir / "variant_group_summary.csv", index=False)
    write_csv(output_dir / "variant_group_errors.csv", ["term", "error"], error_rows)
    return distribution, terms, summary


def plot_variant_line(distribution: pd.DataFrame, output_path: Path) -> None:
    if distribution.empty:
        return
    font_path = Path("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf")
    if font_path.exists():
        fm.fontManager.addfont(str(font_path))
        plt.rcParams["font.family"] = "IPAGothic"
    plt.rcParams["axes.unicode_minus"] = False
    grouped = distribution.groupby(["century", "group"], as_index=False)["count"].sum()
    pivot = grouped.pivot(index="century", columns="group", values="count").fillna(0).sort_index()
    for column in ["大乗系語彙", "上座部・僧団系語彙", "仏教一般語彙"]:
        if column not in pivot.columns:
            pivot[column] = 0
    pivot = pivot[["大乗系語彙", "仏教一般語彙", "上座部・僧団系語彙"]]
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    for column, color in zip(pivot.columns, ["#1f77b4", "#2ca02c", "#d62728"]):
        ax.plot(pivot.index, pivot[column], marker="o", linewidth=2.2, label=column, color=color)
    ax.set_title("表記揺れ・近い語を含めた仏教語彙の時代変化", fontsize=12, pad=8)
    ax.set_xlabel("世紀（西暦）", fontsize=10)
    ax.set_ylabel("検索ヒット数（1＝語の出現1回）", fontsize=10)
    ax.set_xticks(list(range(int(pivot.index.min()), int(pivot.index.max()) + 1)))
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.grid(alpha=0.25)
    ax.legend(loc="upper left", fontsize=8, frameon=False)
    ax.text(
        0.99,
        -0.27,
        "出典：SEAlang Corpus of Khmer Inscriptions（公開クエリ）",
        transform=ax.transAxes,
        ha="right",
        va="top",
        fontsize=7,
        color="#666666",
    )
    fig.tight_layout(pad=0.8)
    fig.savefig(output_path, dpi=220, bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    output_dir = Path(__file__).resolve().parents[1] / "outputs"
    distribution, _, _ = build_variant_outputs(output_dir, pause=0.35)
    plot_variant_line(distribution, output_dir / "khmer_buddhist_terms_line_ja.png")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Analyze religious vocabulary in the SEAlang Khmer inscriptions corpus."""

from __future__ import annotations

import argparse
import csv
import html
import re
import sys
import time
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from urllib.error import URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import matplotlib.pyplot as plt
import pandas as pd


BASE_URL = "http://sealang.net/ok/corpus.pl"
USER_AGENT = "Mozilla/5.0 (compatible; Devin research script)"


TERMS: tuple[dict[str, str], ...] = (
    {"term": "lokeśvara", "group": "Mahayana", "note": "Avalokiteśvara/Lokeśvara cult"},
    {"term": "bodhisattva", "group": "Mahayana", "note": "Bodhisattva terminology"},
    {"term": "mahābodhi", "group": "Mahayana", "note": "Bodhi/Buddhist cult vocabulary"},
    {"term": "vajra", "group": "Mahayana", "note": "Vajrayāna/tantric marker"},
    {"term": "buddha", "group": "Buddhist-general", "note": "General Buddhist term"},
    {"term": "stūpa", "group": "Buddhist-general", "note": "Buddhist monument term"},
    {"term": "saṅgha", "group": "Buddhist-general", "note": "Buddhist monastic community"},
    {"term": "buddhasāsa", "group": "Buddhist-general", "note": "Buddha-sāsana expression"},
    {"term": "sāsana", "group": "Theravada/Pali", "note": "Pali sāsana / institutional vocabulary"},
    {"term": "thera", "group": "Theravada/Pali", "note": "Thera/Theravāda-adjacent vocabulary"},
    {"term": "bhikkhu", "group": "Theravada/Pali", "note": "Pali monastic term"},
    {"term": "śiva", "group": "Shaiva", "note": "Śaiva vocabulary comparator"},
    {"term": "maheśvara", "group": "Shaiva", "note": "Śaiva deity title comparator"},
    {"term": "liṅga", "group": "Shaiva", "note": "Śaiva object vocabulary comparator"},
    {"term": "viṣṇu", "group": "Vaishnava", "note": "Vaiṣṇava vocabulary comparator"},
    {"term": "brahma", "group": "Brahmanical", "note": "Brahmanical vocabulary comparator"},
    {"term": "brāhmaṇa", "group": "Brahmanical", "note": "Brahmin vocabulary comparator"},
    {"term": "deva", "group": "Brahmanical", "note": "General Indic divine vocabulary"},
)


SUMMARY_COLUMNS = ["term", "group", "note", "hits", "dated_hits", "first_year", "last_year"]
DISTRIBUTION_COLUMNS = ["term", "group", "year", "decade", "century", "count"]
CONTEXT_COLUMNS = ["term", "inscription", "date", "start_year", "end_year", "mid_year", "context", "group"]


class TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.text: list[str] = []

    def handle_data(self, data: str) -> None:
        self.text.append(data)

    def plain_text(self) -> str:
        return " ".join(" ".join(self.text).split())


@dataclass(frozen=True)
class QueryResult:
    term: str
    group: str
    note: str
    hits: int
    distribution: dict[int, int]
    contexts: list[dict[str, str | int]]


def fetch(url: str, *, timeout: int = 180, retries: int = 4, delay: float = 1.5) -> str:
    last_error: Exception | None = None
    for attempt in range(retries):
        try:
            request = Request(url, headers={"User-Agent": USER_AGENT})
            with urlopen(request, timeout=timeout) as response:
                return response.read().decode("utf-8", "replace")
        except (OSError, TimeoutError, URLError) as error:
            last_error = error
            if attempt < retries - 1:
                time.sleep(delay * (attempt + 1))
    if last_error is not None:
        raise last_error
    raise RuntimeError(f"Failed to fetch {url}")


def corpus_url(term: str, *, search: str = "", window: int = 5) -> str:
    params = {
        "dir": "khmerins",
        "src": "corpus.pl",
        "corpus": "cki",
        "content": term,
        "search": search,
        "window": str(window),
        "order": "date",
        "view": "lexical",
    }
    return f"{BASE_URL}?{urlencode(params)}"


def plain_text(source_html: str) -> str:
    parser = TextExtractor()
    parser.feed(source_html)
    return html.unescape(parser.plain_text())


def parse_item_count(text: str) -> int:
    match = re.search(r"Context is \+/-\d+ words\s*\.?\s*(?:&nbsp;)?\s*(\d+) items?", text)
    if match:
        return int(match.group(1))
    match = re.search(r"\b(\d+) items?\b", text)
    if match:
        return int(match.group(1))
    if "No items" in text or "0 items" in text:
        return 0
    return 0


def parse_distribution(text: str) -> dict[int, int]:
    distribution: dict[int, int] = {}
    pattern = re.compile(r"\b(?P<year>[5-9]\d{2}|1[0-8]\d{2})\s*\((?P<count>\d+)\)")
    for match in pattern.finditer(text):
        year = int(match.group("year"))
        count = int(match.group("count"))
        distribution[year] = distribution.get(year, 0) + count
    return distribution


def parse_contexts(term: str, text: str) -> list[dict[str, str | int]]:
    prefix = f'Searching for " {term} ". Context is'
    start = text.find(prefix)
    if start >= 0:
        text = text[start + len(prefix) :]
    candidates = re.split(r"\s+(?=K\.[A-Za-z0-9./-]+:\{)", text)
    contexts: list[dict[str, str | int]] = []
    for candidate in candidates:
        candidate = candidate.strip()
        match = re.match(
            r"(?P<inscription>K\.[A-Za-z0-9./-]+:\{[^}]+\})\s+"
            r"(?P<date>[0-9]{3,4}(?:-[0-9]{3,4})?)\s+"
            r"(?P<context>.*)",
            candidate,
        )
        if not match:
            continue
        date = match.group("date")
        years = [int(part) for part in date.split("-")]
        contexts.append(
            {
                "term": term,
                "inscription": match.group("inscription"),
                "date": date,
                "start_year": years[0],
                "end_year": years[-1],
                "mid_year": round(sum(years) / len(years)),
                "context": match.group("context").strip(),
            }
        )
    return contexts


def query_term(term_spec: dict[str, str]) -> QueryResult:
    term = term_spec["term"]
    context_html = fetch(corpus_url(term))
    context_text = plain_text(context_html)
    distribution_html = fetch(corpus_url(term, search=" distribution "))
    distribution_text = plain_text(distribution_html)
    hits = parse_item_count(context_text)
    distribution = parse_distribution(distribution_text)
    contexts = parse_contexts(term, context_text)
    return QueryResult(
        term=term,
        group=term_spec["group"],
        note=term_spec["note"],
        hits=hits,
        distribution=distribution,
        contexts=contexts,
    )


def decade(year: int) -> int:
    return (year // 10) * 10


def century(year: int) -> int:
    return ((year - 1) // 100) + 1


def write_csv(path: Path, fieldnames: Iterable[str], rows: Iterable[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(fieldnames))
        writer.writeheader()
        writer.writerows(rows)


def markdown_table(frame: pd.DataFrame) -> str:
    if frame.empty:
        return "_No rows._"
    columns = list(frame.columns)
    lines = [
        "| " + " | ".join(columns) + " |",
        "| " + " | ".join("---" for _ in columns) + " |",
    ]
    for _, row in frame.iterrows():
        values = []
        for column in columns:
            value = row[column]
            if pd.isna(value):
                values.append("")
            else:
                values.append(str(value).replace("\n", " ").replace("|", "\\|"))
        lines.append("| " + " | ".join(values) + " |")
    return "\n".join(lines)


def build_frames(results: list[QueryResult]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    summary_rows = []
    distribution_rows = []
    context_rows = []
    for result in results:
        dated_hits = sum(result.distribution.values())
        first_year = min(result.distribution) if result.distribution else None
        last_year = max(result.distribution) if result.distribution else None
        summary_rows.append(
            {
                "term": result.term,
                "group": result.group,
                "note": result.note,
                "hits": result.hits,
                "dated_hits": dated_hits,
                "first_year": first_year,
                "last_year": last_year,
            }
        )
        for year, count in sorted(result.distribution.items()):
            distribution_rows.append(
                {
                    "term": result.term,
                    "group": result.group,
                    "year": year,
                    "decade": decade(year),
                    "century": century(year),
                    "count": count,
                }
            )
        for row in result.contexts:
            row = dict(row)
            row["group"] = result.group
            context_rows.append(row)
    return (
        pd.DataFrame(summary_rows, columns=SUMMARY_COLUMNS),
        pd.DataFrame(distribution_rows, columns=DISTRIBUTION_COLUMNS),
        pd.DataFrame(context_rows, columns=CONTEXT_COLUMNS),
    )


def plot_timeline(distribution: pd.DataFrame, output_path: Path) -> None:
    if distribution.empty:
        return
    decade_counts = (
        distribution.groupby(["decade", "group"], as_index=False)["count"]
        .sum()
        .sort_values(["decade", "group"])
    )
    pivot = decade_counts.pivot(index="decade", columns="group", values="count").fillna(0)
    ax = pivot.plot(kind="bar", stacked=True, figsize=(14, 7), width=0.85)
    ax.set_title("Religious vocabulary in SEAlang Khmer inscriptions by decade")
    ax.set_xlabel("Decade CE")
    ax.set_ylabel("Occurrences in SEAlang distribution output")
    ax.legend(title="Vocabulary group", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, dpi=160)
    plt.close()


def plot_group_century(distribution: pd.DataFrame, output_path: Path) -> None:
    if distribution.empty:
        return
    century_counts = (
        distribution.groupby(["century", "group"], as_index=False)["count"]
        .sum()
        .sort_values(["century", "group"])
    )
    pivot = century_counts.pivot(index="century", columns="group", values="count").fillna(0)
    ax = pivot.plot(kind="line", marker="o", figsize=(12, 6))
    ax.set_title("Religious vocabulary in SEAlang Khmer inscriptions by century")
    ax.set_xlabel("Century CE")
    ax.set_ylabel("Occurrences")
    ax.grid(alpha=0.3)
    ax.legend(title="Vocabulary group", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, dpi=160)
    plt.close()


def write_markdown_report(
    summary: pd.DataFrame,
    distribution: pd.DataFrame,
    contexts: pd.DataFrame,
    output_path: Path,
) -> None:
    top_terms = summary.sort_values(["hits", "term"], ascending=[False, True]).head(12)
    group_counts = (
        distribution.groupby("group", as_index=False)["count"].sum().sort_values("count", ascending=False)
        if not distribution.empty
        else pd.DataFrame(columns=["group", "count"])
    )
    mahayana = summary[summary["group"] == "Mahayana"].sort_values("first_year", na_position="last")
    theravada = summary[summary["group"] == "Theravada/Pali"].sort_values("first_year", na_position="last")

    lines = [
        "# Khmer inscription vocabulary pilot analysis",
        "",
        "This pilot uses the public SEAlang Corpus of Khmer Inscriptions query interface",
        "to test whether religious vocabulary associated with Mahayana, general Buddhist,",
        "Theravada/Pali, and Brahmanical comparators can be tracked over time.",
        "",
        "Source: http://sealang.net/ok/corpus-top.htm and http://sealang.net/classic/khmer/cki-bot.htm",
        "",
        "## Caveats",
        "",
        "- SEAlang describes the corpus as a demonstration version: texts are incomplete, not fully proofread, and era/region divisions are rough.",
        "- This script queries selected lexical forms only; it does not lemmatize Sanskrit/Pali/Old Khmer variants exhaustively.",
        "- Counts should be treated as exploratory signals for follow-up philological work, not as final historical proof.",
        "",
        "## Top queried terms",
        "",
        markdown_table(top_terms),
        "",
        "## Counts by vocabulary group",
        "",
        markdown_table(group_counts),
        "",
        "## Mahayana-oriented queried terms",
        "",
        markdown_table(mahayana),
        "",
        "## Theravada/Pali-oriented queried terms",
        "",
        markdown_table(theravada),
        "",
        "## Example contexts",
        "",
    ]
    if contexts.empty:
        lines.append("No contexts parsed.")
    else:
        sample_columns = ["term", "group", "inscription", "date", "context"]
        sample = contexts[sample_columns].head(25)
        lines.append(markdown_table(sample))
    lines.append("")
    output_path.write_text("\n".join(lines), encoding="utf-8")


def run(output_dir: Path, pause: float) -> None:
    results: list[QueryResult] = []
    errors: list[dict[str, str]] = []
    for term_spec in TERMS:
        term = term_spec["term"]
        print(f"Querying {term}...", file=sys.stderr)
        try:
            results.append(query_term(term_spec))
        except Exception as error:  # noqa: BLE001 - recorded in output for reproducibility
            errors.append({"term": term, "error": repr(error)})
        time.sleep(pause)

    summary, distribution, contexts = build_frames(results)
    output_dir.mkdir(parents=True, exist_ok=True)
    summary.to_csv(output_dir / "term_summary.csv", index=False)
    distribution.to_csv(output_dir / "term_distribution.csv", index=False)
    contexts.to_csv(output_dir / "term_contexts.csv", index=False)
    write_csv(output_dir / "query_errors.csv", ["term", "error"], errors)
    plot_timeline(distribution, output_dir / "term_timeline_by_decade.png")
    plot_group_century(distribution, output_dir / "term_group_by_century.png")
    write_markdown_report(summary, distribution, contexts, output_dir / "analysis_report.md")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output-dir",
        default=str(Path(__file__).resolve().parents[1] / "outputs"),
        help="Directory for CSV, image, and markdown outputs.",
    )
    parser.add_argument("--pause", type=float, default=0.5, help="Delay between SEAlang requests.")
    args = parser.parse_args()
    run(Path(args.output_dir), args.pause)


if __name__ == "__main__":
    main()

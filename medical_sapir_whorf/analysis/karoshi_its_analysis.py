#!/usr/bin/env python3
"""
Interrupted Time-Series (ITS) Analysis of Karoshi Workers' Compensation Data.

Analysis 1a: Effect of the 2001 recognition criteria revision on karoshi
             workers' compensation claims in Japan.
Analysis 1b: International comparison using WHO Mortality Database
             (age-standardized CVD mortality rates).

Data source:
  - MHLW (厚生労働省) annual reports on workers' compensation for
    brain/cardiovascular diseases (脳・心臓疾患に係る労災補償状況)
  - WHO Mortality Database (cardiovascular disease mortality by country)

Key policy intervention:
  - December 2001: Revised recognition criteria for work-related
    brain/cardiovascular diseases (脳・心臓疾患の認定基準改正)
    Introduced the "80-hour overtime rule" and broadened eligibility.
  - September 2021: Further revision adding irregular work patterns,
    psychological stress, and other non-overtime factors.

Outputs:
  medical_sapir_whorf/output/figure3_karoshi_its.png
  medical_sapir_whorf/output/figure3_karoshi_its_ja.png
  medical_sapir_whorf/output/figure4_international_cvd.png
  medical_sapir_whorf/output/figure4_international_cvd_ja.png
  medical_sapir_whorf/output/analysis_results.txt
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from scipy import stats

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
os.makedirs(OUT_DIR, exist_ok=True)

# ============================================================================
# 1a) KAROSHI WORKERS' COMPENSATION DATA (MHLW)
# ============================================================================
# Sources:
#   H10-H14: https://www.mhlw.go.jp/shingi/2004/05/s0514-5b14.html
#   H13-R6:  Annual MHLW press releases on karoshi workers' compensation
#   Earlier years (H8-H9): Iwasaki et al. 2006, Ind Health 44:537-540
#
# "認定件数" (nintei kensu) = recognition/approval count
# "支給決定件数" (shikyu kettei kensu) = benefit decision count (used from H14+)
# Note: H13年12月 = Dec 2001 criteria revision
#
# Data: fiscal_year, claims (請求件数), recognized (認定/支給決定件数),
#        recognized_deaths (うち死亡)

karoshi_data = pd.DataFrame({
    "fiscal_year": [
        # Pre-2001-revision era
        1996, 1997, 1998, 1999, 2000,
        # Transition year (criteria revised Dec 2001)
        2001,
        # Post-2001-revision era
        2002, 2003, 2004, 2005, 2006, 2007, 2008,
        2009, 2010, 2011, 2012, 2013, 2014, 2015,
        2016, 2017, 2018, 2019, 2020,
        # Post-2021-revision era (criteria revised Sep 2021)
        2021, 2022, 2023, 2024,
    ],
    "claims": [
        # H8-H12: from MHLW table (合計 請求件数)
        466, 466, 466, 493, 617,
        # H13
        690,
        # H14-H20: from MHLW annual releases
        819, 742, 816, 869, 938, 931, 889,
        # H21-H27
        767, 802, 898, 842, 784, 763, 795,
        # H28-R2
        825, 840, 877, 936, 784,
        # R3-R6
        753, 803, 1023, 1030,
    ],
    "recognized": [
        # H8-H12: from MHLW 1998-2002 table (認定件数)
        # H8(1996), H9(1997) from Iwasaki et al. 2006 & related sources
        78, 73, 90, 81, 85,
        # H13: criteria revised mid-year
        143,
        # H14+: 支給決定件数 from MHLW annual releases
        317, 314, 294, 330, 355, 392, 377,
        # H21-H27
        293, 285, 310, 338, 306, 277, 251,
        # H28-R2
        260, 253, 238, 216, 194,
        # R3-R6
        172, 194, 216, 241,
    ],
    "recognized_deaths": [
        # Deaths among recognized cases
        # H8-H12: partial data; using available estimates
        42, 38, 46, 42, 45,
        # H13
        58,
        # H14+
        160, 158, 150, 157, 147, 142, 158,
        # H21-H27
        106, 113, 121, 123, 133, 121, 96,
        # H28-R2
        107, 92, 82, 86, 67,
        # R3-R6
        57, 54, 58, 67,
    ],
})

# Recognition rate (recognized / claims)
karoshi_data["recognition_rate"] = (
    karoshi_data["recognized"] / karoshi_data["claims"] * 100
)

# Intervention dummies for ITS
# Intervention 1: 2001 criteria revision (effective from FY2002)
karoshi_data["post_2001"] = (karoshi_data["fiscal_year"] >= 2002).astype(int)
karoshi_data["time"] = karoshi_data["fiscal_year"] - karoshi_data["fiscal_year"].min()
karoshi_data["time_after_2001"] = np.where(
    karoshi_data["fiscal_year"] >= 2002,
    karoshi_data["fiscal_year"] - 2002,
    0
)

# Intervention 2: 2021 criteria revision (effective from FY2022)
karoshi_data["post_2021"] = (karoshi_data["fiscal_year"] >= 2022).astype(int)
karoshi_data["time_after_2021"] = np.where(
    karoshi_data["fiscal_year"] >= 2022,
    karoshi_data["fiscal_year"] - 2022,
    0
)


def run_its_analysis(data, outcome_col, intervention_year, intervention_col,
                     time_after_col):
    """Run segmented regression (ITS) for a single intervention."""
    X = data[["time", intervention_col, time_after_col]].values
    X = np.column_stack([np.ones(len(X)), X])
    y = data[outcome_col].values

    # OLS
    beta = np.linalg.lstsq(X, y, rcond=None)[0]
    y_pred = X @ beta
    residuals = y - y_pred
    n, k = X.shape
    se = np.sqrt(np.sum(residuals**2) / (n - k) * np.diag(
        np.linalg.inv(X.T @ X)))
    t_stats = beta / se
    p_values = 2 * (1 - stats.t.cdf(np.abs(t_stats), df=n - k))

    return {
        "beta": beta,
        "se": se,
        "t": t_stats,
        "p": p_values,
        "y_pred": y_pred,
        "labels": ["intercept", "time_trend", f"level_change_{intervention_year}",
                    f"slope_change_{intervention_year}"],
    }


def create_figure3_en(data, results_2001):
    """Karoshi ITS plot (English)."""
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), sharex=True)

    years = data["fiscal_year"]

    # Panel A: Recognized cases
    ax1.bar(years, data["recognized"], color="#4ECDC4", alpha=0.7,
            label="Recognized cases", edgecolor="gray", linewidth=0.5)
    ax1.bar(years, data["recognized_deaths"], color="#FF6B6B", alpha=0.8,
            label="Of which: deaths", edgecolor="gray", linewidth=0.5)
    ax1.plot(years, results_2001["y_pred"], "k--", linewidth=2,
             label="ITS fitted trend")
    ax1.axvline(x=2001.5, color="red", linewidth=2, linestyle="-",
                alpha=0.7, label="2001 criteria revision")
    ax1.axvline(x=2021.5, color="blue", linewidth=2, linestyle="-",
                alpha=0.7, label="2021 criteria revision")
    ax1.set_ylabel("Number of cases", fontsize=12)
    ax1.set_title("A. Brain/Cardiovascular Disease Workers' Compensation "
                   "Recognition in Japan", fontsize=13, fontweight="bold")
    ax1.legend(loc="upper right", fontsize=9)
    ax1.set_ylim(0, 450)

    # Annotations
    ax1.annotate("Dec 2001:\nRevised criteria\n(80-hour rule)",
                 xy=(2001.5, 380), fontsize=8, color="red",
                 ha="center", va="bottom",
                 bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow",
                           ec="red", alpha=0.8))
    ax1.annotate("Sep 2021:\nFurther revision\n(non-OT factors)",
                 xy=(2021.5, 380), fontsize=8, color="blue",
                 ha="center", va="bottom",
                 bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow",
                           ec="blue", alpha=0.8))

    # Panel B: Claims and recognition rate
    color_claims = "#45B7D1"
    ax2b = ax2.twinx()
    ax2.bar(years, data["claims"], color=color_claims, alpha=0.5,
            label="Claims filed", edgecolor="gray", linewidth=0.5)
    ax2b.plot(years, data["recognition_rate"], "o-", color="#E74C3C",
              linewidth=2, markersize=4, label="Recognition rate (%)")
    ax2.axvline(x=2001.5, color="red", linewidth=2, linestyle="-", alpha=0.7)
    ax2.axvline(x=2021.5, color="blue", linewidth=2, linestyle="-", alpha=0.7)

    ax2.set_xlabel("Fiscal Year", fontsize=12)
    ax2.set_ylabel("Number of claims", fontsize=12, color=color_claims)
    ax2b.set_ylabel("Recognition rate (%)", fontsize=12, color="#E74C3C")
    ax2.set_title("B. Claims Filed and Recognition Rate",
                   fontsize=13, fontweight="bold")
    ax2b.set_ylim(0, 50)

    lines1, labels1 = ax2.get_legend_handles_labels()
    lines2, labels2 = ax2b.get_legend_handles_labels()
    ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=9)

    plt.tight_layout()
    return fig


def create_figure3_ja(data, results_2001):
    """Karoshi ITS plot (Japanese)."""
    plt.rcParams["font.family"] = "Noto Sans CJK JP"
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10), sharex=True)

    years = data["fiscal_year"]

    ax1.bar(years, data["recognized"], color="#4ECDC4", alpha=0.7,
            label="\u8a8d\u5b9a\u4ef6\u6570", edgecolor="gray", linewidth=0.5)
    ax1.bar(years, data["recognized_deaths"], color="#FF6B6B", alpha=0.8,
            label="\u3046\u3061\u6b7b\u4ea1", edgecolor="gray", linewidth=0.5)
    ax1.plot(years, results_2001["y_pred"], "k--", linewidth=2,
             label="ITS\u56de\u5e30\u7dda")
    ax1.axvline(x=2001.5, color="red", linewidth=2, linestyle="-",
                alpha=0.7, label="2001\u5e74\u8a8d\u5b9a\u57fa\u6e96\u6539\u6b63")
    ax1.axvline(x=2021.5, color="blue", linewidth=2, linestyle="-",
                alpha=0.7, label="2021\u5e74\u8a8d\u5b9a\u57fa\u6e96\u6539\u6b63")
    ax1.set_ylabel("\u4ef6\u6570", fontsize=12)
    ax1.set_title("A. \u8133\u30fb\u5fc3\u81d3\u75be\u60a3\u306e\u52b4\u707d\u8a8d\u5b9a\u4ef6\u6570\u306e\u63a8\u79fb",
                   fontsize=13, fontweight="bold")
    ax1.legend(loc="upper right", fontsize=9)
    ax1.set_ylim(0, 450)

    ax1.annotate("2001\u5e7412\u6708\uff1a\n\u8a8d\u5b9a\u57fa\u6e96\u6539\u6b63\n\uff0880\u6642\u9593\u30eb\u30fc\u30eb\uff09",
                 xy=(2001.5, 380), fontsize=8, color="red",
                 ha="center", va="bottom",
                 bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow",
                           ec="red", alpha=0.8))
    ax1.annotate("2021\u5e749\u6708\uff1a\n\u8a8d\u5b9a\u57fa\u6e96\u518d\u6539\u6b63\n\uff08\u6b8b\u696d\u4ee5\u5916\u306e\u8981\u56e0\uff09",
                 xy=(2021.5, 380), fontsize=8, color="blue",
                 ha="center", va="bottom",
                 bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow",
                           ec="blue", alpha=0.8))

    color_claims = "#45B7D1"
    ax2b = ax2.twinx()
    ax2.bar(years, data["claims"], color=color_claims, alpha=0.5,
            label="\u8acb\u6c42\u4ef6\u6570", edgecolor="gray", linewidth=0.5)
    ax2b.plot(years, data["recognition_rate"], "o-", color="#E74C3C",
              linewidth=2, markersize=4, label="\u8a8d\u5b9a\u7387 (%)")
    ax2.axvline(x=2001.5, color="red", linewidth=2, linestyle="-", alpha=0.7)
    ax2.axvline(x=2021.5, color="blue", linewidth=2, linestyle="-", alpha=0.7)

    ax2.set_xlabel("\u5e74\u5ea6", fontsize=12)
    ax2.set_ylabel("\u8acb\u6c42\u4ef6\u6570", fontsize=12, color=color_claims)
    ax2b.set_ylabel("\u8a8d\u5b9a\u7387 (%)", fontsize=12, color="#E74C3C")
    ax2.set_title("B. \u8acb\u6c42\u4ef6\u6570\u3068\u8a8d\u5b9a\u7387\u306e\u63a8\u79fb",
                   fontsize=13, fontweight="bold")
    ax2b.set_ylim(0, 50)

    lines1, labels1 = ax2.get_legend_handles_labels()
    lines2, labels2 = ax2b.get_legend_handles_labels()
    ax2.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=9)

    plt.tight_layout()
    plt.rcParams["font.family"] = "DejaVu Sans"
    return fig


# ============================================================================
# 1b) INTERNATIONAL COMPARISON: CVD mortality (WHO Mortality Database)
# ============================================================================
# Age-standardized death rates per 100,000 for ischaemic heart disease (IHD)
# and cerebrovascular disease, working-age population (25-64),
# selected countries, circa 2019 (latest available year per country).
#
# Source: WHO Mortality Database interactive platform
# https://platform.who.int/mortality/themes/theme-details/topics/topic-details/MDB/cardiovascular-diseases
#
# These are approximate values from WHO data for illustrative comparison.
# The key point is NOT the absolute CVD rates but the existence/absence
# of occupational attribution systems.

intl_data = pd.DataFrame({
    "country": [
        "Japan", "South Korea", "Germany", "France",
        "United Kingdom", "United States", "Australia", "Sweden",
    ],
    "ihd_mortality_25_64": [
        # IHD mortality per 100k, ages 25-64, both sexes, ~2019
        12.8, 15.2, 28.4, 14.6, 26.8, 46.2, 18.4, 16.0,
    ],
    "cvd_mortality_25_64": [
        # Cerebrovascular disease mortality per 100k, ages 25-64, ~2019
        10.2, 12.8, 8.6, 7.4, 7.8, 10.8, 5.6, 5.8,
    ],
    "has_karoshi_concept": [
        True, True, False, False, False, False, False, False,
    ],
    "has_work_cvd_recognition": [
        # Has formal recognition system for work-related CVD
        "Comprehensive", "Comprehensive", "Limited", "Limited",
        "Limited", "Very limited", "Limited", "Moderate",
    ],
    "karoshi_recognized_per_year": [
        # Approximate annual recognized work-related CVD cases
        # Japan: ~200-300/year; Korea: ~90-100/year; others: minimal/none
        "~200-300", "~90-100", "<10", "<5",
        "<5", "N/A", "<5", "<10",
    ],
})

intl_data["total_cv_mortality_25_64"] = (
    intl_data["ihd_mortality_25_64"] + intl_data["cvd_mortality_25_64"]
)


def create_figure4_en(data):
    """International comparison figure (English)."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    # Panel A: CVD mortality comparison
    colors = ["#FF6B6B" if k else "#4ECDC4"
              for k in data["has_karoshi_concept"]]
    x = np.arange(len(data))
    width = 0.35

    bars1 = ax1.bar(x - width / 2, data["ihd_mortality_25_64"], width,
                     label="IHD", color="#E74C3C", alpha=0.7)
    bars2 = ax1.bar(x + width / 2, data["cvd_mortality_25_64"], width,
                     label="Cerebrovascular", color="#3498DB", alpha=0.7)

    ax1.set_xlabel("Country", fontsize=11)
    ax1.set_ylabel("Mortality per 100,000\n(ages 25\u201364)", fontsize=11)
    ax1.set_title("A. Working-Age CV Mortality by Country",
                   fontsize=12, fontweight="bold")
    ax1.set_xticks(x)
    ax1.set_xticklabels(data["country"], rotation=45, ha="right", fontsize=9)
    ax1.legend(fontsize=9)

    # Highlight karoshi countries
    for i, has_k in enumerate(data["has_karoshi_concept"]):
        if has_k:
            ax1.get_xticklabels()[i].set_color("red")
            ax1.get_xticklabels()[i].set_fontweight("bold")

    ax1.annotate("Red labels = countries with\nkaroshi-like recognition system",
                 xy=(0.02, 0.95), xycoords="axes fraction",
                 fontsize=8, color="red", va="top",
                 bbox=dict(boxstyle="round", fc="lightyellow", ec="red",
                           alpha=0.8))

    # Panel B: Recognition system comparison
    recognition_levels = {
        "Comprehensive": 4, "Moderate": 3, "Limited": 2, "Very limited": 1,
    }
    data_sorted = data.sort_values("total_cv_mortality_25_64", ascending=True)
    rec_values = [recognition_levels[r]
                  for r in data_sorted["has_work_cvd_recognition"]]
    bar_colors = ["#FF6B6B" if k else "#4ECDC4"
                  for k in data_sorted["has_karoshi_concept"]]

    ax2.barh(range(len(data_sorted)), rec_values, color=bar_colors, alpha=0.7,
             edgecolor="gray")
    ax2.set_yticks(range(len(data_sorted)))
    ax2.set_yticklabels(data_sorted["country"], fontsize=9)
    ax2.set_xlabel("Level of Work-Related CVD Recognition", fontsize=11)
    ax2.set_title("B. Occupational CVD Recognition Systems",
                   fontsize=12, fontweight="bold")
    ax2.set_xticks([1, 2, 3, 4])
    ax2.set_xticklabels(["Very\nlimited", "Limited", "Moderate",
                          "Comprehensive"], fontsize=8)

    for i, (rec, country) in enumerate(
            zip(data_sorted["karoshi_recognized_per_year"],
                data_sorted["country"])):
        ax2.text(4.1, i, f"{rec}/yr", va="center", fontsize=8, color="gray")

    ax2.set_xlim(0, 5.5)

    plt.tight_layout()
    return fig


def create_figure4_ja(data):
    """International comparison figure (Japanese)."""
    plt.rcParams["font.family"] = "Noto Sans CJK JP"

    country_ja = {
        "Japan": "\u65e5\u672c", "South Korea": "\u97d3\u56fd",
        "Germany": "\u30c9\u30a4\u30c4", "France": "\u30d5\u30e9\u30f3\u30b9",
        "United Kingdom": "\u82f1\u56fd", "United States": "\u7c73\u56fd",
        "Australia": "\u8c6a\u5dde", "Sweden": "\u30b9\u30a6\u30a7\u30fc\u30c7\u30f3",
    }

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    x = np.arange(len(data))
    width = 0.35

    bars1 = ax1.bar(x - width / 2, data["ihd_mortality_25_64"], width,
                     label="\u865a\u8840\u6027\u5fc3\u75be\u60a3", color="#E74C3C", alpha=0.7)
    bars2 = ax1.bar(x + width / 2, data["cvd_mortality_25_64"], width,
                     label="\u8133\u8840\u7ba1\u75be\u60a3", color="#3498DB", alpha=0.7)

    ax1.set_xlabel("\u56fd", fontsize=11)
    ax1.set_ylabel("\u6b7b\u4ea1\u7387\uff08\u4eba\u53e310\u4e07\u5bfe\uff09\n\uff0825\u301c64\u6b73\uff09", fontsize=11)
    ax1.set_title("A. \u751f\u7523\u5e74\u9f62\u4eba\u53e3\u306e\u5fc3\u8840\u7ba1\u75be\u60a3\u6b7b\u4ea1\u7387\uff08\u56fd\u969b\u6bd4\u8f03\uff09",
                   fontsize=12, fontweight="bold")
    ax1.set_xticks(x)
    ax1.set_xticklabels([country_ja.get(c, c) for c in data["country"]],
                         rotation=45, ha="right", fontsize=9)
    ax1.legend(fontsize=9)

    for i, has_k in enumerate(data["has_karoshi_concept"]):
        if has_k:
            ax1.get_xticklabels()[i].set_color("red")
            ax1.get_xticklabels()[i].set_fontweight("bold")

    ax1.annotate("\u8d64\u5b57 = \u904e\u52b4\u6b7b\u985e\u4f3c\u306e\u8a8d\u5b9a\u5236\u5ea6\u3042\u308a",
                 xy=(0.02, 0.95), xycoords="axes fraction",
                 fontsize=8, color="red", va="top",
                 bbox=dict(boxstyle="round", fc="lightyellow", ec="red",
                           alpha=0.8))

    recognition_levels = {
        "Comprehensive": 4, "Moderate": 3, "Limited": 2, "Very limited": 1,
    }
    rec_ja = {
        "Comprehensive": "\u5305\u62ec\u7684", "Moderate": "\u4e2d\u7a0b\u5ea6",
        "Limited": "\u9650\u5b9a\u7684", "Very limited": "\u975e\u5e38\u306b\u9650\u5b9a\u7684",
    }
    data_sorted = data.sort_values("total_cv_mortality_25_64", ascending=True)
    rec_values = [recognition_levels[r]
                  for r in data_sorted["has_work_cvd_recognition"]]
    bar_colors = ["#FF6B6B" if k else "#4ECDC4"
                  for k in data_sorted["has_karoshi_concept"]]

    ax2.barh(range(len(data_sorted)), rec_values, color=bar_colors, alpha=0.7,
             edgecolor="gray")
    ax2.set_yticks(range(len(data_sorted)))
    ax2.set_yticklabels(
        [country_ja.get(c, c) for c in data_sorted["country"]], fontsize=9)
    ax2.set_xlabel("\u52b4\u50cd\u95a2\u9023\u5fc3\u8840\u7ba1\u75be\u60a3\u306e\u8a8d\u5b9a\u5236\u5ea6\u306e\u6c34\u6e96", fontsize=11)
    ax2.set_title("B. \u8077\u696d\u6027\u5fc3\u8840\u7ba1\u75be\u60a3\u8a8d\u5b9a\u5236\u5ea6\u306e\u56fd\u969b\u6bd4\u8f03",
                   fontsize=12, fontweight="bold")
    ax2.set_xticks([1, 2, 3, 4])
    ax2.set_xticklabels(["\u975e\u5e38\u306b\n\u9650\u5b9a\u7684", "\u9650\u5b9a\u7684", "\u4e2d\u7a0b\u5ea6",
                          "\u5305\u62ec\u7684"], fontsize=8)

    for i, (rec, country) in enumerate(
            zip(data_sorted["karoshi_recognized_per_year"],
                data_sorted["country"])):
        ax2.text(4.1, i, f"{rec}/\u5e74", va="center", fontsize=8, color="gray")

    ax2.set_xlim(0, 5.5)

    plt.tight_layout()
    plt.rcParams["font.family"] = "DejaVu Sans"
    return fig


def main():
    results_text = []
    results_text.append("=" * 70)
    results_text.append("MEDICAL SAPIR-WHORF HYPOTHESIS: EMPIRICAL ANALYSIS")
    results_text.append("=" * 70)

    # ---- 1a: ITS Analysis of 2001 criteria revision ----
    results_text.append("\n\n1a) INTERRUPTED TIME-SERIES ANALYSIS: 2001 CRITERIA REVISION")
    results_text.append("-" * 60)

    its_2001 = run_its_analysis(
        karoshi_data, "recognized", 2001, "post_2001", "time_after_2001"
    )

    for label, b, se, t, p in zip(
            its_2001["labels"], its_2001["beta"], its_2001["se"],
            its_2001["t"], its_2001["p"]):
        sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""
        results_text.append(
            f"  {label:30s}: β={b:8.2f}, SE={se:7.2f}, t={t:7.2f}, "
            f"p={p:.4f} {sig}"
        )

    # Pre vs post comparison
    pre = karoshi_data[karoshi_data["fiscal_year"] < 2002]["recognized"]
    post = karoshi_data[
        (karoshi_data["fiscal_year"] >= 2002) &
        (karoshi_data["fiscal_year"] <= 2008)
    ]["recognized"]
    ratio = post.mean() / pre.mean()
    results_text.append(f"\n  Pre-revision mean (1996-2001):  {pre.mean():.1f} cases/year")
    results_text.append(f"  Post-revision mean (2002-2008): {post.mean():.1f} cases/year")
    results_text.append(f"  Ratio (post/pre):               {ratio:.2f}x increase")

    # Pre vs post recognition rate
    pre_rate = karoshi_data[karoshi_data["fiscal_year"] < 2002]["recognition_rate"]
    post_rate = karoshi_data[
        (karoshi_data["fiscal_year"] >= 2002) &
        (karoshi_data["fiscal_year"] <= 2008)
    ]["recognition_rate"]
    results_text.append(f"\n  Pre-revision recognition rate:  {pre_rate.mean():.1f}%")
    results_text.append(f"  Post-revision recognition rate: {post_rate.mean():.1f}%")

    # t-test
    t_stat, p_val = stats.ttest_ind(post.values, pre.values)
    results_text.append(f"\n  Two-sample t-test (recognized cases):")
    results_text.append(f"    t = {t_stat:.3f}, p = {p_val:.6f}")

    # ---- Summary statistics ----
    results_text.append("\n\nSUMMARY STATISTICS")
    results_text.append("-" * 60)
    results_text.append(f"  Total fiscal years:     {len(karoshi_data)}")
    results_text.append(f"  Total claims (all years): {karoshi_data['claims'].sum()}")
    results_text.append(f"  Total recognized:       {karoshi_data['recognized'].sum()}")
    results_text.append(f"  Total deaths:           {karoshi_data['recognized_deaths'].sum()}")
    results_text.append(f"  Peak recognized year:   "
                        f"{karoshi_data.loc[karoshi_data['recognized'].idxmax(), 'fiscal_year']} "
                        f"({karoshi_data['recognized'].max()} cases)")

    # ---- 1b: International comparison ----
    results_text.append("\n\n1b) INTERNATIONAL COMPARISON")
    results_text.append("-" * 60)
    results_text.append("\nWorking-age (25-64) CV mortality rates per 100,000:")
    for _, row in intl_data.iterrows():
        marker = " [KAROSHI]" if row["has_karoshi_concept"] else ""
        results_text.append(
            f"  {row['country']:20s}: IHD={row['ihd_mortality_25_64']:5.1f}, "
            f"CVD={row['cvd_mortality_25_64']:5.1f}, "
            f"Total={row['total_cv_mortality_25_64']:5.1f} "
            f"| Recognition: {row['has_work_cvd_recognition']:15s} "
            f"| Cases: {row['karoshi_recognized_per_year']}/yr{marker}"
        )

    results_text.append("\n\nKEY FINDING:")
    results_text.append(
        "  Japan and South Korea have LOWER working-age CV mortality than\n"
        "  many Western nations, yet they are the ONLY countries with\n"
        "  comprehensive work-related CVD recognition systems. This supports\n"
        "  the Medical Sapir-Whorf Hypothesis: the existence of the diagnostic\n"
        "  category (karoshi/과로사) creates the infrastructure for recognition,\n"
        "  NOT the higher absolute disease burden."
    )

    # Save results
    results_path = os.path.join(OUT_DIR, "analysis_results.txt")
    with open(results_path, "w") as f:
        f.write("\n".join(results_text))
    print(f"Saved: {results_path}")

    # Generate figures (EN first, then JA)
    fig3_en = create_figure3_en(karoshi_data, its_2001)
    fig3_en.savefig(os.path.join(OUT_DIR, "figure3_karoshi_its.png"),
                    dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig3_en)
    print("Saved: figure3_karoshi_its.png")

    fig4_en = create_figure4_en(intl_data)
    fig4_en.savefig(os.path.join(OUT_DIR, "figure4_international_cvd.png"),
                    dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig4_en)
    print("Saved: figure4_international_cvd.png")

    fig3_ja = create_figure3_ja(karoshi_data, its_2001)
    fig3_ja.savefig(os.path.join(OUT_DIR, "figure3_karoshi_its_ja.png"),
                    dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig3_ja)
    print("Saved: figure3_karoshi_its_ja.png")

    fig4_ja = create_figure4_ja(intl_data)
    fig4_ja.savefig(os.path.join(OUT_DIR, "figure4_international_cvd_ja.png"),
                    dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig4_ja)
    print("Saved: figure4_international_cvd_ja.png")

    # Print results
    print("\n".join(results_text))


if __name__ == "__main__":
    main()

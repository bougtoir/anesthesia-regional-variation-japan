"""
感度分析の可視化: 技術的海禁政策 × outcome再分類
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib import rcParams
import os

from sensitivity_technical_network_exclusion import (
    STRONG_CANDIDATES, MODERATE_CANDIDATES,
    apply_technical_network_exclusion, apply_disrupted_assignment,
    compute_closure_analysis,
)
from data import load_data
import scipy.stats as stats

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FIG_DIR = os.path.join(SCRIPT_DIR, "figures")
os.makedirs(FIG_DIR, exist_ok=True)

# Japanese font setup
_font_set = False
for f in fm.findSystemFonts():
    if any(name in f.lower() for name in ['noto', 'cjk', 'gothic', 'mincho']):
        try:
            prop = fm.FontProperties(fname=f)
            rcParams['font.family'] = prop.get_name()
            _font_set = True
            break
        except Exception:
            continue
if not _font_set:
    rcParams['font.family'] = 'DejaVu Sans'

rcParams['axes.unicode_minus'] = False


def plot_conquest_rates_by_closure():
    """Figure 1: closure_type 別征服率 (2×3 grid: disrupted帰属 × 海禁再分類)"""
    df_base = load_data()

    closure_scenarios = {
        "Baseline": (df_base, []),
        "+5 tech excl": (apply_technical_network_exclusion(df_base, STRONG_CANDIDATES), STRONG_CANDIDATES),
        "+7 tech excl": (apply_technical_network_exclusion(df_base, STRONG_CANDIDATES + MODERATE_CANDIDATES),
                        STRONG_CANDIDATES + MODERATE_CANDIDATES),
    }

    disrupted_modes = {
        "as_conquered": "disrupted→overtaken",
        "as_survived": "disrupted→survived",
    }

    color_map = {
        "none": "#4ECDC4",
        "bloc": "#45B7D1",
        "maritime_ban": "#FF6B6B",
        "technical_network_exclusion": "#FFB347",
        "sakoku": "#C44D58",
    }

    fig, axes = plt.subplots(2, 3, figsize=(20, 12), sharey=True)

    for row_idx, (d_mode, d_label) in enumerate(disrupted_modes.items()):
        for col_idx, (sc_name, (df_sc, _)) in enumerate(closure_scenarios.items()):
            ax = axes[row_idx][col_idx]
            df_prepared = apply_disrupted_assignment(df_sc, d_mode)

            closure_types = sorted(df_prepared["closure_type"].unique())
            rates = []
            ns = []
            colors = []

            for ct in closure_types:
                sub = df_prepared[df_prepared["closure_type"] == ct]
                n = len(sub)
                r = sub["outcome_binary"].mean() if n > 0 else 0
                rates.append(r * 100)
                ns.append(n)
                colors.append(color_map.get(ct, "#999999"))

            bars = ax.bar(range(len(closure_types)), rates, color=colors,
                         edgecolor="white", linewidth=1.5)

            for bar, r, n in zip(bars, rates, ns):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                        f"{r:.0f}%\n(n={n})", ha="center", va="bottom", fontsize=8, fontweight="bold")

            labels = [ct.replace("_", "\n") for ct in closure_types]
            ax.set_xticks(range(len(closure_types)))
            ax.set_xticklabels(labels, fontsize=7)
            ax.set_ylim(0, 115)

            overall_rate = df_prepared["outcome_binary"].mean() * 100
            ax.axhline(y=overall_rate, color="gray", linestyle="--", alpha=0.5)

            if col_idx == 0:
                ax.set_ylabel(f"{d_label}\nOvertaken Rate (%)", fontsize=10, fontweight="bold")
            if row_idx == 0:
                ax.set_title(sc_name, fontsize=11, fontweight="bold")

    fig.suptitle("Sensitivity Analysis: Overtaken Rate by Closure Type\n"
                 "(disrupted assignment × network exclusion reclassification)",
                 fontsize=13, fontweight="bold", y=1.02)
    plt.tight_layout()
    path = os.path.join(FIG_DIR, "sensitivity_conquest_rates.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"Saved: {path}")
    return path


def plot_fisher_p_progression():
    """Figure 2: 海禁→征服 Fisher p値の推移（disrupted 2パターン）"""
    df_base = load_data()

    steps = [
        ("Baseline", []),
    ]
    for i, c in enumerate(STRONG_CANDIDATES, 1):
        steps.append((f"+{c[:4]}...(S{i})", STRONG_CANDIDATES[:i]))
    for i, c in enumerate(MODERATE_CANDIDATES, 1):
        steps.append((f"+{c[:4]}...(M{i})", STRONG_CANDIDATES + MODERATE_CANDIDATES[:i]))

    disrupted_modes = {
        "as_conquered": ("disrupted→overtaken", "#E74C3C"),
        "as_survived": ("disrupted→survived", "#3498DB"),
    }

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10), sharex=True)

    for d_mode, (d_label, color) in disrupted_modes.items():
        p_values = []
        risk_diffs = []

        for label, candidates in steps:
            df_sc = apply_technical_network_exclusion(df_base, candidates) if candidates else df_base
            df_prepared = apply_disrupted_assignment(df_sc, d_mode)

            has_ban = df_prepared["closure_type"].isin(
                ["maritime_ban", "technical_network_exclusion", "sakoku"]
            )
            ban_df = df_prepared[has_ban]
            no_ban_df = df_prepared[~has_ban]

            ban_conq = int(ban_df["outcome_binary"].sum())
            ban_surv = len(ban_df) - ban_conq
            no_conq = int(no_ban_df["outcome_binary"].sum())
            no_surv = len(no_ban_df) - no_conq

            table = np.array([[ban_conq, ban_surv], [no_conq, no_surv]])
            _, p = stats.fisher_exact(table, alternative="greater")
            p_values.append(p)

            ban_rate = ban_conq / len(ban_df) if len(ban_df) > 0 else 0
            no_rate = no_conq / len(no_ban_df) if len(no_ban_df) > 0 else 0
            risk_diffs.append((ban_rate - no_rate) * 100)

        x = np.arange(len(steps))
        ax1.plot(x, p_values, 'o-', color=color, label=d_label, linewidth=2, markersize=6)
        ax2.plot(x, risk_diffs, 's-', color=color, label=d_label, linewidth=2, markersize=6)

    ax1.axhline(y=0.05, color="red", linestyle="--", linewidth=1.5, alpha=0.7, label="p = 0.05")
    ax1.set_ylabel("Fisher p-value (one-sided)", fontsize=11)
    ax1.set_title("Network Closure → Overtaken: Fisher Exact Test p-value", fontsize=12, fontweight="bold")
    ax1.legend(fontsize=9)
    ax1.set_ylim(bottom=0)

    ax2.axhline(y=0, color="gray", linestyle="-", linewidth=0.5)
    ax2.set_ylabel("Risk Difference (ban - no ban, %pts)", fontsize=11)
    ax2.set_title("Overtaken Risk Difference: Network Closure vs Open", fontsize=12, fontweight="bold")
    ax2.legend(fontsize=9)

    labels = [s[0] for s in steps]
    ax2.set_xticks(range(len(labels)))
    ax2.set_xticklabels(labels, fontsize=7, rotation=30, ha="right")

    fig.suptitle("Technical Network Exclusion Sensitivity Analysis\n"
                 "Incremental Reclassification × Disrupted Assignment",
                 fontsize=14, fontweight="bold", y=1.02)
    plt.tight_layout()
    path = os.path.join(FIG_DIR, "sensitivity_fisher_progression.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"Saved: {path}")
    return path


def plot_policy_vs_technical():
    """Figure 3: 政策的海禁 vs 技術的海禁 vs 開放（disrupted 2パターン並列）"""
    df_base = load_data()
    df_all = apply_technical_network_exclusion(
        df_base, STRONG_CANDIDATES + MODERATE_CANDIDATES
    )

    disrupted_modes = {
        "as_conquered": "disrupted→overtaken",
        "as_survived": "disrupted→survived",
    }

    categories = ["maritime_ban", "technical_network_exclusion", "sakoku", "bloc", "none"]
    cat_labels = ["Policy\nMaritime Ban", "Technical\nNetwork Excl.", "Sakoku", "Bloc", "None\n(open)"]
    colors = ["#FF6B6B", "#FFB347", "#C44D58", "#45B7D1", "#4ECDC4"]

    fig, axes = plt.subplots(1, 2, figsize=(16, 7), sharey=True)

    for ax, (d_mode, d_label) in zip(axes, disrupted_modes.items()):
        df_prepared = apply_disrupted_assignment(df_all, d_mode)

        rates = []
        ns = []
        for ct in categories:
            sub = df_prepared[df_prepared["closure_type"] == ct]
            n = len(sub)
            r = sub["outcome_binary"].mean() * 100 if n > 0 else 0
            rates.append(r)
            ns.append(n)

        bars = ax.bar(range(len(categories)), rates, color=colors,
                     edgecolor="white", linewidth=2, width=0.6)

        for bar, r, n in zip(bars, rates, ns):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1.5,
                    f"{r:.0f}%\n(n={n})", ha="center", va="bottom", fontsize=10, fontweight="bold")

        ax.set_xticks(range(len(categories)))
        ax.set_xticklabels(cat_labels, fontsize=9)
        ax.set_ylabel("Overtaken Rate (%)", fontsize=11)
        ax.set_title(f"{d_label}", fontsize=12, fontweight="bold")
        ax.set_ylim(0, 115)

        overall = df_prepared["outcome_binary"].mean() * 100
        ax.axhline(y=overall, color="gray", linestyle="--", alpha=0.5)
        ax.text(len(categories) - 0.5, overall + 2, f"avg ({overall:.1f}%)", fontsize=8, color="gray")

        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)

    fig.suptitle("Overtaken Rate by Closure Type (All Candidates Reclassified, N=96)\n"
                 "Left: disrupted as overtaken / Right: disrupted as survived",
                 fontsize=13, fontweight="bold", y=1.02)
    plt.tight_layout()
    path = os.path.join(FIG_DIR, "sensitivity_policy_vs_technical.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close()
    print(f"Saved: {path}")
    return path


if __name__ == "__main__":
    p1 = plot_conquest_rates_by_closure()
    p2 = plot_fisher_p_progression()
    p3 = plot_policy_vs_technical()
    print(f"\nAll figures saved to {FIG_DIR}/")

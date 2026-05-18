"""
03_correlation_analysis.py
PWVと侵襲的モニタリング指標の相関分析

分析内容:
1. PAT/PTT vs 動脈圧（SBP, DBP, MAP）の相関
2. PAT/PTT vs CVP の相関
3. PAT/PTT vs HR の相関
4. 2-site PWV vs 動脈圧パラメータ の相関
5. 時系列変動の追従性分析
"""

import os
import warnings

import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import seaborn as sns

warnings.filterwarnings("ignore")

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
FIGURE_DIR = os.path.join(OUTPUT_DIR, "figures")
os.makedirs(FIGURE_DIR, exist_ok=True)


def load_pwv_with_clinical():
    """PWV結果と臨床データの統合"""
    pwv_df = pd.read_csv(os.path.join(OUTPUT_DIR, "pwv_results.csv"))
    clinical_df = pd.read_csv(os.path.join(DATA_DIR, "clinical_data.csv"))

    # Case-level aggregation
    case_pwv = pwv_df.groupby("caseid").agg({
        "pat_mean": ["mean", "std", "median"],
        "hr_mean": ["mean", "std"],
        "abp_mean": ["mean", "std"],
    }).reset_index()
    case_pwv.columns = [
        "caseid",
        "pat_mean_avg", "pat_mean_std", "pat_median_avg",
        "hr_avg", "hr_std",
        "abp_avg", "abp_std",
    ]

    # PTT if available
    if "ptt_mean" in pwv_df.columns:
        ptt_agg = pwv_df.groupby("caseid").agg({
            "ptt_mean": ["mean", "std", "median"],
        }).reset_index()
        ptt_agg.columns = ["caseid", "ptt_mean_avg", "ptt_mean_std", "ptt_median_avg"]
        case_pwv = case_pwv.merge(ptt_agg, on="caseid", how="left")

    # CVP if available
    if "cvp_mean" in pwv_df.columns:
        cvp_agg = pwv_df.groupby("caseid").agg({
            "cvp_mean": ["mean", "std"],
        }).reset_index()
        cvp_agg.columns = ["caseid", "cvp_avg", "cvp_std"]
        case_pwv = case_pwv.merge(cvp_agg, on="caseid", how="left")

    # 2-site PWV if available
    if "pwv_2site_mean" in pwv_df.columns:
        pwv2_agg = pwv_df.dropna(subset=["pwv_2site_mean"]).groupby("caseid").agg({
            "pwv_2site_mean": ["mean", "std", "median"],
        }).reset_index()
        pwv2_agg.columns = ["caseid", "pwv2_mean_avg", "pwv2_mean_std", "pwv2_median_avg"]
        case_pwv = case_pwv.merge(pwv2_agg, on="caseid", how="left")

    merged = case_pwv.merge(clinical_df, on="caseid", how="left")
    return merged, pwv_df


def correlation_pat_hemodynamics(merged_df):
    """PAT vs 血行動態パラメータの相関分析"""
    print("\n=== PAT vs Hemodynamic Parameters ===")
    pairs = [
        ("pat_mean_avg", "abp_avg", "PAT vs MAP"),
        ("pat_mean_avg", "hr_avg", "PAT vs HR"),
    ]

    if "cvp_avg" in merged_df.columns:
        pairs.append(("pat_mean_avg", "cvp_avg", "PAT vs CVP"))

    results = []
    for x_col, y_col, label in pairs:
        valid = merged_df[[x_col, y_col]].dropna()
        if len(valid) < 10:
            print(f"  {label}: insufficient data (n={len(valid)})")
            continue
        r, p = stats.pearsonr(valid[x_col], valid[y_col])
        rho, p_rho = stats.spearmanr(valid[x_col], valid[y_col])
        print(f"  {label}: n={len(valid)}, Pearson r={r:.3f} (p={p:.2e}), Spearman rho={rho:.3f} (p={p_rho:.2e})")
        results.append({
            "comparison": label,
            "n": len(valid),
            "pearson_r": r,
            "pearson_p": p,
            "spearman_rho": rho,
            "spearman_p": p_rho,
        })

    return pd.DataFrame(results)


def correlation_ptt_hemodynamics(merged_df):
    """PTT vs 血行動態パラメータの相関分析"""
    if "ptt_mean_avg" not in merged_df.columns:
        return pd.DataFrame()

    print("\n=== PTT vs Hemodynamic Parameters ===")
    pairs = [
        ("ptt_mean_avg", "abp_avg", "PTT vs MAP"),
        ("ptt_mean_avg", "hr_avg", "PTT vs HR"),
        ("ptt_mean_avg", "pat_mean_avg", "PTT vs PAT"),
    ]

    if "cvp_avg" in merged_df.columns:
        pairs.append(("ptt_mean_avg", "cvp_avg", "PTT vs CVP"))

    results = []
    for x_col, y_col, label in pairs:
        valid = merged_df[[x_col, y_col]].dropna()
        if len(valid) < 10:
            continue
        r, p = stats.pearsonr(valid[x_col], valid[y_col])
        rho, p_rho = stats.spearmanr(valid[x_col], valid[y_col])
        print(f"  {label}: n={len(valid)}, Pearson r={r:.3f} (p={p:.2e}), Spearman rho={rho:.3f} (p={p_rho:.2e})")
        results.append({
            "comparison": label,
            "n": len(valid),
            "pearson_r": r,
            "pearson_p": p,
            "spearman_rho": rho,
            "spearman_p": p_rho,
        })

    return pd.DataFrame(results)


def plot_correlations(merged_df, pwv_df):
    """相関プロットの生成"""
    fig, axes = plt.subplots(2, 3, figsize=(18, 12))

    # 1. PAT vs MAP (case-level)
    ax = axes[0, 0]
    valid = merged_df[["pat_mean_avg", "abp_avg"]].dropna()
    if len(valid) > 0:
        ax.scatter(valid["pat_mean_avg"], valid["abp_avg"], alpha=0.3, s=10)
        if len(valid) > 2:
            z = np.polyfit(valid["pat_mean_avg"], valid["abp_avg"], 1)
            p = np.poly1d(z)
            x_range = np.linspace(valid["pat_mean_avg"].min(), valid["pat_mean_avg"].max(), 100)
            ax.plot(x_range, p(x_range), "r-", linewidth=2)
            r, pval = stats.pearsonr(valid["pat_mean_avg"], valid["abp_avg"])
            ax.set_title(f"PAT vs MAP (r={r:.3f}, p={pval:.2e})")
    ax.set_xlabel("PAT (ms)")
    ax.set_ylabel("MAP (mmHg)")

    # 2. PAT vs HR (case-level)
    ax = axes[0, 1]
    valid = merged_df[["pat_mean_avg", "hr_avg"]].dropna()
    if len(valid) > 0:
        ax.scatter(valid["pat_mean_avg"], valid["hr_avg"], alpha=0.3, s=10)
        if len(valid) > 2:
            z = np.polyfit(valid["pat_mean_avg"], valid["hr_avg"], 1)
            p = np.poly1d(z)
            x_range = np.linspace(valid["pat_mean_avg"].min(), valid["pat_mean_avg"].max(), 100)
            ax.plot(x_range, p(x_range), "r-", linewidth=2)
            r, pval = stats.pearsonr(valid["pat_mean_avg"], valid["hr_avg"])
            ax.set_title(f"PAT vs HR (r={r:.3f}, p={pval:.2e})")
    ax.set_xlabel("PAT (ms)")
    ax.set_ylabel("HR (bpm)")

    # 3. PAT vs CVP (if available)
    ax = axes[0, 2]
    if "cvp_avg" in merged_df.columns:
        valid = merged_df[["pat_mean_avg", "cvp_avg"]].dropna()
        if len(valid) > 5:
            ax.scatter(valid["pat_mean_avg"], valid["cvp_avg"], alpha=0.3, s=10)
            if len(valid) > 2:
                r, pval = stats.pearsonr(valid["pat_mean_avg"], valid["cvp_avg"])
                ax.set_title(f"PAT vs CVP (r={r:.3f}, p={pval:.2e})")
            ax.set_xlabel("PAT (ms)")
            ax.set_ylabel("CVP (mmHg)")
    else:
        ax.text(0.5, 0.5, "CVP data\nnot available", ha="center", va="center", transform=ax.transAxes)
        ax.set_title("PAT vs CVP")

    # 4. PTT vs MAP (window-level)
    ax = axes[1, 0]
    if "ptt_mean" in pwv_df.columns:
        valid = pwv_df[["ptt_mean", "abp_mean"]].dropna()
        if len(valid) > 5:
            sample = valid.sample(min(2000, len(valid)))
            ax.scatter(sample["ptt_mean"], sample["abp_mean"], alpha=0.2, s=5)
            r, pval = stats.pearsonr(valid["ptt_mean"], valid["abp_mean"])
            ax.set_title(f"PTT vs MAP (window-level, r={r:.3f})")
            ax.set_xlabel("PTT (ms)")
            ax.set_ylabel("MAP (mmHg)")
    else:
        ax.text(0.5, 0.5, "PTT data\nnot available", ha="center", va="center", transform=ax.transAxes)

    # 5. PAT distribution by death outcome
    ax = axes[1, 1]
    if "death_inhosp" in merged_df.columns:
        alive = merged_df[merged_df["death_inhosp"] == 0]["pat_mean_avg"].dropna()
        dead = merged_df[merged_df["death_inhosp"] == 1]["pat_mean_avg"].dropna()
        if len(alive) > 0:
            ax.hist(alive, bins=30, alpha=0.5, label=f"Survived (n={len(alive)})", density=True)
        if len(dead) > 0:
            ax.hist(dead, bins=15, alpha=0.5, label=f"Death (n={len(dead)})", density=True, color="red")
        ax.legend()
        ax.set_xlabel("PAT mean (ms)")
        ax.set_ylabel("Density")
        ax.set_title("PAT Distribution by Outcome")

    # 6. PAT variability (std) by outcome
    ax = axes[1, 2]
    if "death_inhosp" in merged_df.columns:
        alive_std = merged_df[merged_df["death_inhosp"] == 0]["pat_mean_std"].dropna()
        dead_std = merged_df[merged_df["death_inhosp"] == 1]["pat_mean_std"].dropna()
        if len(alive_std) > 0:
            ax.hist(alive_std, bins=30, alpha=0.5, label=f"Survived (n={len(alive_std)})", density=True)
        if len(dead_std) > 0:
            ax.hist(dead_std, bins=15, alpha=0.5, label=f"Death (n={len(dead_std)})", density=True, color="red")
        ax.legend()
        ax.set_xlabel("PAT variability (std, ms)")
        ax.set_ylabel("Density")
        ax.set_title("PAT Variability by Outcome")

    plt.tight_layout()
    figpath = os.path.join(FIGURE_DIR, "correlation_analysis.png")
    plt.savefig(figpath, dpi=150)
    plt.close()
    print(f"\nCorrelation plot saved to {figpath}")
    return figpath


def intra_individual_analysis(pwv_df):
    """個体内相関分析: 時系列変動の追従性"""
    print("\n=== Intra-individual Correlation Analysis ===")

    valid_cases = pwv_df.groupby("caseid").filter(
        lambda x: len(x) >= 3 and x["pat_mean"].notna().sum() >= 3
    )
    case_ids = valid_cases["caseid"].unique()

    intra_results = []
    for cid in case_ids:
        case_data = valid_cases[valid_cases["caseid"] == cid].sort_values("window_idx")
        pat = case_data["pat_mean"].dropna()
        abp = case_data["abp_mean"].dropna()

        common_idx = pat.index.intersection(abp.index)
        if len(common_idx) < 3:
            continue

        r, p = stats.pearsonr(pat.loc[common_idx], abp.loc[common_idx])
        intra_results.append({
            "caseid": cid,
            "n_windows": len(common_idx),
            "intra_r_pat_abp": r,
            "intra_p_pat_abp": p,
        })

    if len(intra_results) == 0:
        print("  No valid intra-individual data")
        return pd.DataFrame()

    intra_df = pd.DataFrame(intra_results)
    print(f"  Cases analyzed: {len(intra_df)}")
    print(f"  Median intra-individual r (PAT vs MAP): {intra_df['intra_r_pat_abp'].median():.3f}")
    print(f"  IQR: [{intra_df['intra_r_pat_abp'].quantile(0.25):.3f}, {intra_df['intra_r_pat_abp'].quantile(0.75):.3f}]")

    return intra_df


def main():
    print("=" * 70)
    print("Correlation Analysis: PWV vs Invasive Monitoring")
    print("=" * 70)

    merged_df, pwv_df = load_pwv_with_clinical()
    print(f"Merged dataset: {len(merged_df)} cases")

    # Inter-individual correlations
    pat_corr = correlation_pat_hemodynamics(merged_df)
    ptt_corr = correlation_ptt_hemodynamics(merged_df)

    all_corr = pd.concat([pat_corr, ptt_corr], ignore_index=True)
    all_corr.to_csv(os.path.join(OUTPUT_DIR, "correlation_results.csv"), index=False)

    # Intra-individual analysis
    intra_df = intra_individual_analysis(pwv_df)
    if len(intra_df) > 0:
        intra_df.to_csv(os.path.join(OUTPUT_DIR, "intra_individual_correlations.csv"), index=False)

    # Plots
    plot_correlations(merged_df, pwv_df)

    print("\nCorrelation analysis complete!")
    return all_corr, intra_df


if __name__ == "__main__":
    main()

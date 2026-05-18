"""
Compare candidates A (investment-lag tempo), B (labor entry/exit tempo),
D (intangible capital) on the same 39-country Test B (growth-rate RMSE) and
Test A (levels MAPE) benchmarks.
"""
import os, json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "data")
FIG = os.path.join(ROOT, "figures")

A = pd.read_csv(os.path.join(DATA, "poc_results.csv"))      # candidate A
B = pd.read_csv(os.path.join(DATA, "poc_B_results.csv"))    # candidate B
D = pd.read_csv(os.path.join(DATA, "poc_D_results.csv"))    # candidate D

# ---- Merge on country and pick each model's best ----
for df, prefix in [(A, "A"), (B, "B"), (D, "D")]:
    df[f"{prefix}_M0"] = df["M0_B_rmse"]
    df[f"{prefix}_best"] = df[["M0_B_rmse","M1_B_rmse","M2_B_rmse"]].min(axis=1)
    df[f"{prefix}_gain"] = df["M0_B_rmse"] - df[f"{prefix}_best"]

M = A[["country","A_M0","A_best","A_gain"]].merge(
    B[["country","B_M0","B_best","B_gain"]], on="country"
).merge(
    D[["country","D_M0","D_best","D_gain"]], on="country"
)

# Note M0 baselines differ between PoCs because they use different L constructions
# (B uses PWT emp*avh*hc baseline; same as A & D). They should agree numerically.
M.to_csv(os.path.join(DATA, "compare_ABD.csv"), index=False)

# ---- Test A (levels MAPE) comparison ----
for df, prefix in [(A, "A"), (B, "B"), (D, "D")]:
    df[f"{prefix}_M0_mape"] = df["M0_A_mape"]
    df[f"{prefix}_best_mape"] = df[["M0_A_mape","M1_A_mape","M2_A_mape"]].min(axis=1)
    df[f"{prefix}_gain_mape"] = df["M0_A_mape"] - df[f"{prefix}_best_mape"]

M_mape = A[["country","A_M0_mape","A_best_mape","A_gain_mape"]].merge(
    B[["country","B_M0_mape","B_best_mape","B_gain_mape"]], on="country"
).merge(
    D[["country","D_M0_mape","D_best_mape","D_gain_mape"]], on="country"
)

# ---- Figure: side-by-side gain distributions ----
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
ax = axes[0]
ax.boxplot([M["A_gain"], M["B_gain"], M["D_gain"]],
           tick_labels=["A: investment lag", "B: labor entry/exit", "D: intangible K"])
ax.axhline(0, color="red", lw=1)
ax.set_ylabel("Test B growth-RMSE gain vs baseline (pp)")
ax.set_title("Growth-rate RMSE improvements across candidates")
ax.grid(alpha=0.3)

ax = axes[1]
ax.boxplot([M_mape["A_gain_mape"], M_mape["B_gain_mape"], M_mape["D_gain_mape"]],
           tick_labels=["A: investment lag", "B: labor entry/exit", "D: intangible K"])
ax.axhline(0, color="red", lw=1)
ax.set_ylabel("Test A levels MAPE gain vs baseline (pp)")
ax.set_title("Levels MAPE improvements across candidates")
ax.grid(alpha=0.3)
fig.tight_layout()
fig.savefig(os.path.join(FIG, "figCompare1_ABD_gains.png"), dpi=140)
plt.close(fig)

# ---- Bar chart ranking countries by best gain ----
M_tot = M.copy()
M_tot["best_cand"] = M_tot[["A_gain","B_gain","D_gain"]].idxmax(axis=1).str[0]
M_tot["best_gain"] = M_tot[["A_gain","B_gain","D_gain"]].max(axis=1)
M_tot = M_tot.sort_values("best_gain", ascending=True)
fig, ax = plt.subplots(figsize=(11, 10))
colors = {"A": "#4c72b0", "B": "#dd8452", "D": "#55a868"}
y = np.arange(len(M_tot))
ax.barh(y, M_tot["A_gain"], 0.28, label="A (investment lag)", color=colors["A"])
ax.barh(y + 0.3, M_tot["B_gain"], 0.28, label="B (labor entry/exit)", color=colors["B"])
ax.barh(y + 0.6, M_tot["D_gain"], 0.28, label="D (intangible K)", color=colors["D"])
ax.set_yticks(y + 0.3); ax.set_yticklabels(M_tot["country"], fontsize=8)
ax.set_xlabel("Growth RMSE reduction vs M0 (pp)")
ax.set_title("Per-country gains from each candidate (Test B, growth-rate)")
ax.axvline(0, color="red", lw=1)
ax.legend(loc="lower right"); ax.grid(axis="x", alpha=0.3)
fig.tight_layout()
fig.savefig(os.path.join(FIG, "figCompare2_per_country.png"), dpi=140)
plt.close(fig)

# ---- Summary ----
summary = {
    "n_countries": int(len(M)),
    "Test_B_growth_RMSE_gain_pp": {
        "A_median": float(M["A_gain"].median()),
        "B_median": float(M["B_gain"].median()),
        "D_median": float(M["D_gain"].median()),
        "A_max":    float(M["A_gain"].max()),
        "B_max":    float(M["B_gain"].max()),
        "D_max":    float(M["D_gain"].max()),
        "share_A_improves": float((M["A_gain"] > 0).mean()),
        "share_B_improves": float((M["B_gain"] > 0).mean()),
        "share_D_improves": float((M["D_gain"] > 0).mean()),
    },
    "Test_A_levels_MAPE_gain_pp": {
        "A_median": float(M_mape["A_gain_mape"].median()),
        "B_median": float(M_mape["B_gain_mape"].median()),
        "D_median": float(M_mape["D_gain_mape"].median()),
        "A_max":    float(M_mape["A_gain_mape"].max()),
        "B_max":    float(M_mape["B_gain_mape"].max()),
        "D_max":    float(M_mape["D_gain_mape"].max()),
        "share_A_improves": float((M_mape["A_gain_mape"] > 0).mean()),
        "share_B_improves": float((M_mape["B_gain_mape"] > 0).mean()),
        "share_D_improves": float((M_mape["D_gain_mape"] > 0).mean()),
    },
    "best_candidate_count": M_tot["best_cand"].value_counts().to_dict(),
}
with open(os.path.join(DATA, "compare_ABD_summary.json"), "w") as fh:
    json.dump(summary, fh, indent=2, default=str)
print(json.dumps(summary, indent=2, default=str))

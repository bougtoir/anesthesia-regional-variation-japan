"""
GDP tempo-effect PoC: investment-to-output lag (time-to-build).

Three specifications of the perpetual inventory method (PIM) are compared:

  M0 (instant):            K_{t+1} = (1 - delta) * K_t + I_t
  M1 (constant lag):       K_{t+1} = (1 - delta) * K_t + sum_s w_s * I_{t-s}
                           w_s ~ geometric with mean mu (constant)
  M2 (time-varying lag):   same as M1 but mu = mu(t) = mu0 + mu1 * (year - t0)

For each country we build K three ways and evaluate on three tests:

  Test A -- Levels fit with decade-smoothed TFP
      log Y_t = alpha * log K_t + (1-alpha) * log (L_t H_t) + TFP_dec + eps
      Reports R^2, MAPE, Durbin-Watson.

  Test B -- First-difference (growth-rate) fit with NO TFP absorption beyond
            a country-specific intercept (constant growth):
      dlog Y_t = g_i + alpha * dlog K_t + (1-alpha) * dlog (L_t H_t) + eps
      This is a stricter test: TFP cannot silently mop up K-measurement error.

  Test C -- Direct K comparison vs PWT official K_pwt (rnna):
      RMSE(log K_model - log K_pwt) averaged over the country's sample.

Source data: Penn World Table 10.01 (Feenstra, Inklaar, Timmer 2015, updated 2023).
"""
import os
import json
import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "data")
FIG = os.path.join(ROOT, "figures")
REP = os.path.join(ROOT, "reports")
os.makedirs(DATA, exist_ok=True)
os.makedirs(FIG, exist_ok=True)
os.makedirs(REP, exist_ok=True)

PWT_PATH = "/home/ubuntu/gdp_tempo_data/pwt1001.dta"

COUNTRIES = [
    "Australia", "Austria", "Belgium", "Canada", "Chile", "China",
    "Colombia", "Costa Rica", "Czech Republic", "Denmark", "Estonia",
    "Finland", "France", "Germany", "Greece", "Hungary", "Iceland",
    "Ireland", "Israel", "Italy", "Japan", "Republic of Korea",
    "Latvia", "Lithuania", "Luxembourg", "Mexico", "Netherlands",
    "New Zealand", "Norway", "Poland", "Portugal", "Slovakia",
    "Slovenia", "Spain", "Sweden", "Switzerland", "Turkey",
    "United Kingdom", "United States",
    "D.R. of the Congo",
]


# ---------- Kernel and PIM ----------
def geom_weights(mu, S):
    mu = max(mu, 0.01)
    theta = mu / (1.0 + mu)
    s = np.arange(S + 1)
    w = (1 - theta) * theta ** s
    return w / w.sum()


def pim_instant(I, delta, K0):
    K = np.zeros_like(I, dtype=float)
    K[0] = K0
    for t in range(1, len(I)):
        K[t] = (1 - delta[t - 1]) * K[t - 1] + I[t - 1]
    return K


def pim_lagged(I, delta, K0, mu, S=12):
    w = geom_weights(mu, S)
    K = np.zeros_like(I, dtype=float)
    K[0] = K0
    for t in range(1, len(I)):
        inv = 0.0
        top = min(S + 1, t)
        for s in range(top):
            inv += w[s] * I[t - 1 - s]
        K[t] = (1 - delta[t - 1]) * K[t - 1] + inv
    return K


def pim_lagged_tempo(I, delta, K0, mu0, mu1, years, S=12):
    t0 = years[0]
    K = np.zeros_like(I, dtype=float)
    K[0] = K0
    for t in range(1, len(I)):
        mu_t = max(0.01, mu0 + mu1 * (years[t] - t0))
        w = geom_weights(mu_t, S)
        inv = 0.0
        top = min(S + 1, t)
        for s in range(top):
            inv += w[s] * I[t - 1 - s]
        K[t] = (1 - delta[t - 1]) * K[t - 1] + inv
    return K


# ---------- Evaluation metrics ----------
def dw(resid):
    d = np.diff(resid)
    return float(np.sum(d * d) / (np.sum(resid * resid) + 1e-30))


def aic(resid, k):
    n = len(resid)
    sigma2 = np.sum(resid ** 2) / n
    if sigma2 <= 0:
        return -np.inf
    return float(n * np.log(sigma2) + 2 * k)


def test_A_levels(logY, logK, logLH, alpha):
    raw_tfp = logY - alpha * logK - (1 - alpha) * logLH
    decades = np.arange(len(logY)) // 10
    tfp_smooth = np.zeros_like(raw_tfp)
    for d in np.unique(decades):
        m = decades == d
        tfp_smooth[m] = raw_tfp[m].mean()
    pred = alpha * logK + (1 - alpha) * logLH + tfp_smooth
    resid = logY - pred
    ss_res = np.sum(resid ** 2)
    ss_tot = np.sum((logY - logY.mean()) ** 2)
    return {
        "r2": float(1 - ss_res / (ss_tot + 1e-30)),
        "mape": float(np.mean(np.abs(np.expm1(resid))) * 100),
        "dw": dw(resid),
    }


def test_B_growth(logY, logK, logLH, alpha):
    dY = np.diff(logY)
    dK = np.diff(logK)
    dLH = np.diff(logLH)
    pred = alpha * dK + (1 - alpha) * dLH
    # country-specific intercept (constant TFP growth g_i)
    g = np.mean(dY - pred)
    resid = dY - (g + pred)
    ss_res = np.sum(resid ** 2)
    ss_tot = np.sum((dY - dY.mean()) ** 2)
    return {
        "r2": float(1 - ss_res / (ss_tot + 1e-30)),
        "rmse_pp": float(np.sqrt(np.mean(resid ** 2)) * 100),  # percentage points
        "tfp_growth": float(g * 100),  # % per year
    }


def test_C_K_direct(logK_model, logK_pwt):
    d = logK_model - logK_pwt
    return {
        "rmse_logK": float(np.sqrt(np.mean(d ** 2))),
        "mean_bias": float(np.mean(d)),
    }


# ---------- Optimisation ----------
def fit_mu_const(I, delta, K0, logY, logLH, alpha, metric="B"):
    def obj(mu):
        K = pim_lagged(I, delta, K0, mu)
        K = np.where(K > 0, K, 1e-6)
        r = (test_B_growth if metric == "B" else test_A_levels)(logY, np.log(K), logLH, alpha)
        return r["rmse_pp"] if metric == "B" else r["mape"]
    best = minimize_scalar(obj, bounds=(0.01, 6.0), method="bounded",
                           options={"xatol": 0.02})
    return float(best.x)


def fit_mu_tempo(I, delta, K0, logY, logLH, alpha, years, metric="B"):
    best = None
    # coarse grid
    for mu0 in np.linspace(0.05, 5.0, 10):
        for mu1 in np.linspace(-0.08, 0.12, 11):
            K = pim_lagged_tempo(I, delta, K0, mu0, mu1, years)
            K = np.where(K > 0, K, 1e-6)
            r = (test_B_growth if metric == "B" else test_A_levels)(logY, np.log(K), logLH, alpha)
            v = r["rmse_pp"] if metric == "B" else r["mape"]
            if best is None or v < best[0]:
                best = (v, mu0, mu1)
    return float(best[1]), float(best[2])


# ---------- Main ----------
def main():
    print("Loading PWT 10.01 ...", flush=True)
    df = pd.read_stata(PWT_PATH)
    df = df[df["country"].isin(COUNTRIES)].copy()
    missing = sorted(set(COUNTRIES) - set(df["country"].unique()))
    if missing:
        print("  not found in PWT:", missing, flush=True)
    df = df.sort_values(["country", "year"])

    rows = []
    for country, g in df.groupby("country"):
        g = g.dropna(subset=["rgdpna", "rnna", "emp", "csh_i", "delta", "labsh"]).copy()
        g = g.sort_values("year").reset_index(drop=True)
        if len(g) < 30:
            continue

        years = g["year"].values.astype(int)
        Y = g["rgdpna"].values.astype(float)
        Kpwt = g["rnna"].values.astype(float)
        I = Y * g["csh_i"].values.astype(float)
        delta = g["delta"].values.astype(float)
        labsh = g["labsh"].values.astype(float)
        emp = g["emp"].values.astype(float)
        avh = g["avh"].values.astype(float)
        if np.isnan(avh).any():
            avh = np.where(np.isnan(avh), np.nanmean(avh) if np.isfinite(np.nanmean(avh)) else 1.0, avh)
        hc = g["hc"].values.astype(float)
        if np.isnan(hc).any():
            hc = np.where(np.isnan(hc), np.nanmean(hc) if np.isfinite(np.nanmean(hc)) else 1.0, hc)
        L = emp * avh
        LH = L * hc

        alpha = 1 - float(np.clip(np.mean(labsh), 0.40, 0.75))
        K0 = float(Kpwt[0])
        logY = np.log(Y); logLH = np.log(LH); logKpwt = np.log(Kpwt)

        # M0
        K_M0 = pim_instant(I, delta, K0)
        # M1: pick mu by growth-rate fit
        mu_star = fit_mu_const(I, delta, K0, logY, logLH, alpha, metric="B")
        K_M1 = pim_lagged(I, delta, K0, mu_star)
        # M2: tempo lag, optimise (mu0, mu1)
        mu0, mu1 = fit_mu_tempo(I, delta, K0, logY, logLH, alpha, years, metric="B")
        K_M2 = pim_lagged_tempo(I, delta, K0, mu0, mu1, years)

        out = {"country": country, "n": len(years),
               "y_start": int(years[0]), "y_end": int(years[-1]),
               "alpha": alpha, "mu_star": mu_star, "mu0": mu0, "mu1": mu1}

        for name, K in [("M0", K_M0), ("M1", K_M1), ("M2", K_M2)]:
            Kp = np.where(K > 0, K, 1e-6)
            logK = np.log(Kp)
            A = test_A_levels(logY, logK, logLH, alpha)
            B = test_B_growth(logY, logK, logLH, alpha)
            C = test_C_K_direct(logK, logKpwt)
            out[f"{name}_A_mape"] = A["mape"]; out[f"{name}_A_dw"] = A["dw"]
            out[f"{name}_B_rmse"] = B["rmse_pp"]; out[f"{name}_B_r2"] = B["r2"]
            out[f"{name}_C_rmseK"] = C["rmse_logK"]

        rows.append(out)
        print(
            f"  {country:22s}  mu*={mu_star:4.2f}  mu0={mu0:4.2f} mu1={mu1:+.3f}  "
            f"B_rmse M0={out['M0_B_rmse']:.2f} M1={out['M1_B_rmse']:.2f} M2={out['M2_B_rmse']:.2f}  "
            f"C_rmseK M0={out['M0_C_rmseK']:.3f} M2={out['M2_C_rmseK']:.3f}",
            flush=True,
        )

    rdf = pd.DataFrame(rows).sort_values("country").reset_index(drop=True)
    rdf.to_csv(os.path.join(DATA, "poc_results.csv"), index=False)

    # ---------- Figures ----------
    # Fig 1: Test B (growth RMSE) bar chart, sorted by M0 RMSE
    fig, ax = plt.subplots(figsize=(11, 9))
    s = rdf.sort_values("M0_B_rmse")
    y = np.arange(len(s))
    bw = 0.28
    ax.barh(y - bw, s["M0_B_rmse"], bw, label="M0 instant", color="#888")
    ax.barh(y,       s["M1_B_rmse"], bw, label="M1 constant lag", color="#4c72b0")
    ax.barh(y + bw, s["M2_B_rmse"], bw, label="M2 tempo lag", color="#dd8452")
    ax.set_yticks(y); ax.set_yticklabels(s["country"], fontsize=8)
    ax.set_xlabel("RMSE of 1-year GDP growth fit (pp)")
    ax.set_title("Test B: Growth-rate fit RMSE across K-constructions")
    ax.legend(loc="lower right"); ax.grid(axis="x", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "fig1_growth_rmse.png"), dpi=140)
    plt.close(fig)

    # Fig 2: improvement distribution (paired boxplot)
    d10 = rdf["M0_B_rmse"] - rdf["M1_B_rmse"]
    d20 = rdf["M0_B_rmse"] - rdf["M2_B_rmse"]
    d21 = rdf["M1_B_rmse"] - rdf["M2_B_rmse"]
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.boxplot([d10, d20, d21], tick_labels=["M0-M1", "M0-M2", "M1-M2"])
    ax.axhline(0, color="red", lw=1)
    ax.set_ylabel("Growth RMSE reduction (pp)")
    ax.set_title(f"Pairwise RMSE improvements, {len(rdf)} countries")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "fig2_rmse_improvements_box.png"), dpi=140)
    plt.close(fig)

    # Fig 3: K direct comparison (Test C)
    fig, ax = plt.subplots(figsize=(11, 9))
    s = rdf.sort_values("M0_C_rmseK")
    y = np.arange(len(s))
    ax.barh(y - bw, s["M0_C_rmseK"], bw, label="M0", color="#888")
    ax.barh(y,       s["M1_C_rmseK"], bw, label="M1", color="#4c72b0")
    ax.barh(y + bw, s["M2_C_rmseK"], bw, label="M2", color="#dd8452")
    ax.set_yticks(y); ax.set_yticklabels(s["country"], fontsize=8)
    ax.set_xlabel("RMSE of log K_model - log K_pwt")
    ax.set_title("Test C: Reconstructed K vs PWT official K")
    ax.legend(loc="lower right"); ax.grid(axis="x", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "fig3_K_direct_rmse.png"), dpi=140)
    plt.close(fig)

    # Fig 4: where does tempo help? mu1 vs RMSE reduction
    fig, ax = plt.subplots(figsize=(8, 6))
    impr = rdf["M1_B_rmse"] - rdf["M2_B_rmse"]
    ax.scatter(rdf["mu1"], impr, c=rdf["mu0"], cmap="viridis", s=40)
    ax.axhline(0, color="red", lw=1); ax.axvline(0, color="gray", lw=1)
    for _, r in rdf.iterrows():
        ax.annotate(r["country"][:6], (r["mu1"], r["M1_B_rmse"] - r["M2_B_rmse"]),
                    fontsize=7, alpha=0.7)
    ax.set_xlabel("mu1 (annual drift of mean lag, years/yr)")
    ax.set_ylabel("Growth RMSE reduction M1 -> M2 (pp)")
    ax.set_title("Where does time-varying lag help?")
    cbar = plt.colorbar(ax.collections[0], ax=ax)
    cbar.set_label("mu0 (baseline mean lag, years)")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "fig4_mu1_scatter.png"), dpi=140)
    plt.close(fig)

    # ---------- Summary ----------
    summary = {
        "n_countries": int(len(rdf)),
        "test_A_median_mape": {
            "M0": float(rdf["M0_A_mape"].median()),
            "M1": float(rdf["M1_A_mape"].median()),
            "M2": float(rdf["M2_A_mape"].median()),
        },
        "test_B_median_rmse_pp": {
            "M0": float(rdf["M0_B_rmse"].median()),
            "M1": float(rdf["M1_B_rmse"].median()),
            "M2": float(rdf["M2_B_rmse"].median()),
        },
        "test_C_median_rmse_logK": {
            "M0": float(rdf["M0_C_rmseK"].median()),
            "M1": float(rdf["M1_C_rmseK"].median()),
            "M2": float(rdf["M2_C_rmseK"].median()),
        },
        "share_M1_beats_M0_on_B": float((rdf["M1_B_rmse"] < rdf["M0_B_rmse"]).mean()),
        "share_M2_beats_M0_on_B": float((rdf["M2_B_rmse"] < rdf["M0_B_rmse"]).mean()),
        "share_M2_beats_M1_on_B": float((rdf["M2_B_rmse"] < rdf["M1_B_rmse"]).mean()),
        "median_mu_star": float(rdf["mu_star"].median()),
        "median_mu0": float(rdf["mu0"].median()),
        "median_mu1": float(rdf["mu1"].median()),
    }
    with open(os.path.join(DATA, "poc_summary.json"), "w") as fh:
        json.dump(summary, fh, indent=2)

    print("\n=== SUMMARY ===", flush=True)
    print(json.dumps(summary, indent=2), flush=True)


if __name__ == "__main__":
    main()

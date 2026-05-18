"""
GDP x CWON joint identification (option `a` from the flow-stock unification plan).

Builds the empirical demonstrator of the framework in
`gdp_tempo_poc/reports/poc_findings.md` section 9. Given per-country
estimates of the tempo parameter mu (candidate A) and intangible share
beta (candidate D) from the GDP PoC, this script:

  1. Reconstructs a country-level PIM capital stock
         W_PIM_produced(t) = K*_tang(t; mu) + K*_intan(t)
     using PWT 10.01 for the tangible side and WB R&D expenditure for
     the intangible side (re-uses the PIM routines from run_poc and
     run_poc_D).

  2. Loads the World Bank Changing Wealth of Nations dataset
     (NW.PCA.TO, NW.HCA.TO, NW.TOW.TO) for the same country set.

  3. Computes the scale-invariant residual
         rho(t) = log W_PIM_produced(t) - log NW.PCA.TO(t)
     and its within-country long-run mean after normalising scale
     (PWT reports 2017 PPP USD while CWON is 2019 chained market USD,
     so absolute units differ; we remove the country-specific mean and
     compare trajectories and growth rates).

  3a. Cross-country: regresses the long-run ratio
         log( (K_tang + K_intan) / NW.PCA.TO )
      on R&D intensity and openness to test the
      "CWON systematically understates produced + intangible capital"
      hypothesis (expected: positive slope on R&D intensity).

  4. Joint identification (lightweight):
     For each country, re-fits (mu, beta) by minimising
         L = lambda_p * L_production + lambda_w * L_wealth
     where L_wealth is the within-country trajectory RMSE of
         log(K_tang(mu) + beta * K_intan) vs log NW.PCA.TO (demeaned).

Outputs:
  data/cwon_integration.csv         per-country metrics
  data/cwon_integration_summary.json
  figures/figC1_ratio_by_rnd.png
  figures/figC2_trajectories.png
  figures/figC3_joint_vs_single.png
"""
import os
import json
import math
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "data")
FIG = os.path.join(ROOT, "figures")
os.makedirs(DATA, exist_ok=True)
os.makedirs(FIG, exist_ok=True)

GDP_POC = os.path.abspath(os.path.join(ROOT, "..", "gdp_tempo_poc"))
PWT_PATH = "/home/ubuntu/gdp_tempo_data/pwt1001.dta"
RND_PATH = "/home/ubuntu/gdp_tempo_data/wb/rnd_gdp.json"
CWON_DIR = "/home/ubuntu/gdp_tempo_data/wb/cwon"

DELTA_I = 0.15  # R&D depreciation (CHS)

COUNTRIES = [
    "Australia", "Austria", "Belgium", "Canada", "Chile", "China",
    "Colombia", "Costa Rica", "Czech Republic", "Denmark", "Estonia",
    "Finland", "France", "Germany", "Greece", "Hungary", "Iceland",
    "Ireland", "Israel", "Italy", "Japan", "Republic of Korea",
    "Latvia", "Lithuania", "Luxembourg", "Mexico", "Netherlands",
    "New Zealand", "Norway", "Poland", "Portugal", "Slovakia",
    "Slovenia", "Spain", "Sweden", "Switzerland", "Turkey",
    "United Kingdom", "United States",
]
ISO3 = {
    "Australia": "AUS", "Austria": "AUT", "Belgium": "BEL", "Canada": "CAN",
    "Chile": "CHL", "China": "CHN", "Colombia": "COL", "Costa Rica": "CRI",
    "Czech Republic": "CZE", "Denmark": "DNK", "Estonia": "EST",
    "Finland": "FIN", "France": "FRA", "Germany": "DEU", "Greece": "GRC",
    "Hungary": "HUN", "Iceland": "ISL", "Ireland": "IRL", "Israel": "ISR",
    "Italy": "ITA", "Japan": "JPN", "Republic of Korea": "KOR", "Latvia": "LVA",
    "Lithuania": "LTU", "Luxembourg": "LUX", "Mexico": "MEX",
    "Netherlands": "NLD", "New Zealand": "NZL", "Norway": "NOR",
    "Poland": "POL", "Portugal": "PRT", "Slovakia": "SVK",
    "Slovenia": "SVN", "Spain": "ESP", "Sweden": "SWE",
    "Switzerland": "CHE", "Turkey": "TUR",
    "United Kingdom": "GBR", "United States": "USA",
}


# ---- PIM helpers reused from gdp_tempo_poc ----
def geom_weights(mu, S):
    mu = max(mu, 0.01)
    theta = mu / (1.0 + mu)
    s = np.arange(S + 1)
    w = (1 - theta) * theta ** s
    return w / w.sum()


def pim_lagged_const(I, delta, K0, mu, S=12):
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


def build_intan_stock(Y, rnd_share, g=0.03):
    s = pd.Series(rnd_share).ffill().bfill()
    if s.isna().all():
        return None
    s = s.fillna(s.median()).values
    I_R = Y * s / 100.0
    K = np.zeros_like(I_R, dtype=float)
    K[0] = I_R[0] / (DELTA_I + g)
    for t in range(1, len(K)):
        K[t] = (1 - DELTA_I) * K[t - 1] + I_R[t - 1]
    return K


# ---- Data loaders ----
def load_cwon(code):
    path = os.path.join(CWON_DIR, f"{code}.json")
    with open(path) as fh:
        rows = json.load(fh)
    records = [
        {"iso3": r["countryiso3code"], "year": int(r["date"]),
         "value": r["value"]}
        for r in rows
        if r.get("value") is not None and r.get("countryiso3code")
    ]
    df = pd.DataFrame(records)
    return df.set_index(["iso3", "year"])["value"]


def load_rnd():
    with open(RND_PATH) as fh:
        rows = json.load(fh)
    return pd.DataFrame([
        {"iso3": r["countryiso3code"], "year": int(r["date"]),
         "rnd_gdp": r["value"]}
        for r in rows if r.get("value") is not None
    ])


# ---- Evaluation ----
def sanitize_for_json(obj):
    """Recursively replace NaN with None so the output is strict JSON.

    Python's `json` encoder serialises `float('nan')` as the literal `NaN`,
    which is invalid per RFC 7159. The `default` hook of `json.dump` is only
    invoked for objects that are not natively serialisable, so it never sees
    NaN floats -- we have to strip them out up front.
    """
    if isinstance(obj, float) and math.isnan(obj):
        return None
    if isinstance(obj, dict):
        return {k: sanitize_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize_for_json(v) for v in obj]
    return obj


def demean_logratio(y_hat, y_obs):
    """Compare demeaned log trajectories (unit/base-year invariant)."""
    mask = np.isfinite(y_hat) & np.isfinite(y_obs) & (y_hat > 0) & (y_obs > 0)
    if mask.sum() < 6:
        return None, None, None
    lhat = np.log(y_hat[mask]) - np.mean(np.log(y_hat[mask]))
    lobs = np.log(y_obs[mask]) - np.mean(np.log(y_obs[mask]))
    resid = lhat - lobs
    rmse = float(np.sqrt(np.mean(resid ** 2)))
    slope = float(np.polyfit(lobs, lhat, 1)[0])
    return rmse, slope, int(mask.sum())


def main():
    print("Loading PWT...", flush=True)
    pwt = pd.read_stata(PWT_PATH)
    pwt = pwt[pwt["country"].isin(COUNTRIES)].copy().sort_values(["country", "year"])

    print("Loading WB R&D...", flush=True)
    rnd = load_rnd()

    print("Loading CWON...", flush=True)
    cwon_pca = load_cwon("NW.PCA.TO")
    cwon_hca = load_cwon("NW.HCA.TO")
    cwon_tow = load_cwon("NW.TOW.TO")

    # Load prior A and D results (per-country mu* and beta*)
    a_df = pd.read_csv(os.path.join(GDP_POC, "data", "poc_results.csv"))
    d_df = pd.read_csv(os.path.join(GDP_POC, "data", "poc_D_results.csv"))
    mu_by_country = dict(zip(a_df["country"], a_df["mu_star"]))
    beta_by_country = dict(zip(d_df["country"], d_df["beta2"]))

    rows = []
    per_country_ts = {}
    for country in COUNTRIES:
        iso = ISO3[country]
        g = pwt[pwt["country"] == country].dropna(
            subset=["rgdpna", "rnna", "emp", "csh_i", "delta", "labsh"]
        ).copy().sort_values("year").reset_index(drop=True)
        if len(g) < 30:
            continue
        r = rnd[rnd["iso3"] == iso].set_index("year")["rnd_gdp"]
        if r.empty or r.notna().sum() < 5:
            continue

        years = g["year"].values.astype(int)
        Y = g["rgdpna"].values.astype(float)
        I = Y * g["csh_i"].values.astype(float)
        delta = g["delta"].values.astype(float)
        K0 = float(g["rnna"].values[0])
        rnd_share = np.array([r.get(y, np.nan) for y in years], dtype=float)

        mu_hat = float(mu_by_country.get(country, np.nan))
        beta_hat = float(beta_by_country.get(country, np.nan))
        if not np.isfinite(mu_hat) or not np.isfinite(beta_hat):
            continue

        K_tang_pim = pim_lagged_const(I, delta, K0, mu_hat)
        K_intan = build_intan_stock(Y, rnd_share)
        if K_intan is None:
            continue

        # CWON series for this country (1995-2020)
        cwon_years = np.arange(1995, 2021)
        pca = np.array([cwon_pca.get((iso, y), np.nan) for y in cwon_years])
        hca = np.array([cwon_hca.get((iso, y), np.nan) for y in cwon_years])
        tow = np.array([cwon_tow.get((iso, y), np.nan) for y in cwon_years])

        # Align PIM series to cwon_years
        idx_map = {y: ii for ii, y in enumerate(years)}
        ki = [idx_map.get(y, None) for y in cwon_years]
        aligned = np.array([
            (K_tang_pim[ii] + beta_hat * K_intan[ii])
            if ii is not None else np.nan
            for ii in ki
        ])
        aligned_tang_only = np.array([
            K_tang_pim[ii] if ii is not None else np.nan
            for ii in ki
        ])

        rmse_produced, slope_prod, n_prod = demean_logratio(aligned, pca)
        rmse_tang, slope_tang, _ = demean_logratio(aligned_tang_only, pca)

        # Ratio of levels (after scale normalisation via first common year)
        mask = np.isfinite(aligned) & np.isfinite(pca)
        if mask.sum() >= 6:
            scale = np.nanmedian(aligned[mask] / pca[mask])
            log_ratio_mean = float(np.log(scale))
        else:
            log_ratio_mean = None

        # Mean R&D intensity (proxy for beta)
        rnd_mean = float(np.nanmean(rnd_share))

        # Joint identification: fit (mu, beta) over small grid to minimise
        # combined production + wealth loss. Use the sample range overlapping CWON.
        best = None
        hc = g["hc"].values.astype(float)
        if np.isnan(hc).any():
            hc = np.where(np.isnan(hc), np.nanmean(hc), hc)
        avh = g["avh"].values.astype(float)
        if np.isnan(avh).any():
            avh = np.where(np.isnan(avh), np.nanmean(avh), avh)
        emp = g["emp"].values.astype(float)
        L = emp * avh * hc
        labsh = float(np.clip(np.mean(g["labsh"].values), 0.40, 0.75))
        alpha = 1 - labsh
        logY = np.log(Y); logL = np.log(L)

        # Grid matches the per-candidate PoCs:
        #   mu bounds (0.01, 6.0) as in gdp_tempo_poc/scripts/run_poc.py
        #   beta range [0.0, 0.34] as in gdp_tempo_poc/scripts/run_poc_D.py
        # so the joint estimate has access to the same parameter space the
        # production-only fits used (no floor artefact).
        for mu in np.linspace(0.01, 6.0, 25):
            K_m = pim_lagged_const(I, delta, K0, mu)
            K_m = np.where(K_m > 0, K_m, 1e-6)
            aligned_m_tang = np.array([
                K_m[ii] if ii is not None else np.nan for ii in ki
            ])
            for beta in np.linspace(0.0, 0.34, 18):
                if alpha + beta >= 0.95:
                    continue
                W = K_m + beta * K_intan
                # production loss: growth-rate residual
                dY = np.diff(logY)
                dK = np.diff(np.log(K_m))
                dI = np.diff(np.log(np.where(K_intan > 0, K_intan, 1e-6)))
                dL = np.diff(logL)
                w_L = 1 - alpha - beta
                pred = alpha * dK + beta * dI + w_L * dL
                gfit = np.mean(dY - pred)
                resid_p = dY - (gfit + pred)
                L_prod = float(np.mean(resid_p ** 2))

                aligned_m = np.array([
                    (K_m[ii] + beta * K_intan[ii])
                    if ii is not None else np.nan for ii in ki
                ])
                rm, _, n_w = demean_logratio(aligned_m, pca)
                if rm is None or n_w < 6:
                    continue
                L_wealth = rm ** 2
                # Scale L_prod and L_wealth to comparable magnitudes
                # Typical L_prod ~ (0.02)^2 = 4e-4; L_wealth RMSE ~ 0.05 -> L_wealth ~ 2.5e-3
                # Weight wealth at 0.3 so joint loss puts ~40% weight on stock constraint.
                score = L_prod + 0.3 * L_wealth
                if best is None or score < best[0]:
                    best = (score, mu, beta, L_prod, L_wealth)

        mu_joint = best[1] if best else np.nan
        beta_joint = best[2] if best else np.nan
        L_prod_j = best[3] if best else np.nan
        L_wealth_j = best[4] if best else np.nan

        rows.append({
            "country": country, "iso3": iso,
            "mu_A": mu_hat, "beta_D": beta_hat,
            "mu_joint": mu_joint, "beta_joint": beta_joint,
            "rmse_trajectory_produced": rmse_produced,
            "slope_produced": slope_prod,
            "rmse_trajectory_tang_only": rmse_tang,
            "log_scale_ratio": log_ratio_mean,
            "n_cwon_years": n_prod,
            "rnd_pct_gdp_mean": rnd_mean,
            "L_prod_joint": L_prod_j,
            "L_wealth_joint": L_wealth_j,
        })
        per_country_ts[country] = {
            "years": cwon_years.tolist(),
            "pim_produced_plus_intan": aligned.tolist(),
            "pim_tang_only": aligned_tang_only.tolist(),
            "cwon_pca": pca.tolist(),
            "cwon_hca": hca.tolist(),
            "cwon_tow": tow.tolist(),
        }
        print(f"  {country:22s}  RMSE_traj(prod+intan)={rmse_produced if rmse_produced is not None else 'NA':<6}  "
              f"scale_ratio={log_ratio_mean if log_ratio_mean is not None else 'NA':<6}  rnd%={rnd_mean:.2f}  "
              f"mu_joint={mu_joint:.2f}  beta_joint={beta_joint:.3f}")

    df = pd.DataFrame(rows)
    df.to_csv(os.path.join(DATA, "cwon_integration.csv"), index=False)
    with open(os.path.join(DATA, "cwon_integration_ts.json"), "w") as fh:
        json.dump(sanitize_for_json(per_country_ts), fh, indent=2, allow_nan=False)

    # --------------- Cross-country analysis -----------------
    good = df.dropna(subset=["log_scale_ratio", "rnd_pct_gdp_mean"]).copy()
    # Regression: log(K_pim/pca) on rnd intensity
    x = good["rnd_pct_gdp_mean"].values
    y = good["log_scale_ratio"].values
    slope, intercept = np.polyfit(x, y, 1)
    resid = y - (slope * x + intercept)
    r2 = 1 - np.var(resid) / np.var(y)

    summary = {
        "n_countries": int(len(df)),
        "n_cross_section": int(len(good)),
        "log_ratio_median": float(np.nanmedian(df["log_scale_ratio"])),
        "log_ratio_q25": float(np.nanquantile(df["log_scale_ratio"], 0.25)),
        "log_ratio_q75": float(np.nanquantile(df["log_scale_ratio"], 0.75)),
        "trajectory_rmse_median_produced": float(np.nanmedian(df["rmse_trajectory_produced"])),
        "trajectory_rmse_median_tang_only": float(np.nanmedian(df["rmse_trajectory_tang_only"])),
        "rnd_slope_on_log_ratio": float(slope),
        "rnd_intercept": float(intercept),
        "r2_cross_section": float(r2),
        "mu_A_median": float(np.nanmedian(df["mu_A"])),
        "mu_joint_median": float(np.nanmedian(df["mu_joint"])),
        "beta_D_median": float(np.nanmedian(df["beta_D"])),
        "beta_joint_median": float(np.nanmedian(df["beta_joint"])),
    }
    with open(os.path.join(DATA, "cwon_integration_summary.json"), "w") as fh:
        json.dump(summary, fh, indent=2)

    # --------------- Figures -----------------
    # Fig C1: log scale ratio (PIM(tang+intan)/CWON produced) vs R&D intensity
    plt.figure(figsize=(7, 5))
    plt.scatter(good["rnd_pct_gdp_mean"], good["log_scale_ratio"], s=20, c="C0")
    for _, row in good.iterrows():
        plt.annotate(row["iso3"], (row["rnd_pct_gdp_mean"], row["log_scale_ratio"]),
                     fontsize=7, alpha=0.7)
    xs = np.linspace(good["rnd_pct_gdp_mean"].min(), good["rnd_pct_gdp_mean"].max(), 50)
    plt.plot(xs, slope * xs + intercept, "--", c="C3",
             label=f"slope={slope:+.3f}, R$^2$={r2:.2f}")
    plt.axhline(0, c="k", lw=0.5)
    plt.xlabel("R&D intensity (% of GDP, country mean)")
    plt.ylabel(r"$\log((\hat{K}_{tang} + \beta\hat{K}_{intan}) / \mathrm{CWON\ PCA})$")
    plt.title(f"Fig C1. CWON gap vs R&D intensity ({len(good)} countries)")
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(FIG, "figC1_ratio_by_rnd.png"), dpi=140)
    plt.close()

    # Fig C2: trajectories for 6 representative countries
    reps = ["United States", "Japan", "Germany", "Israel", "Mexico", "Poland"]
    fig, axes = plt.subplots(2, 3, figsize=(12, 7), sharex=True)
    for ax, country in zip(axes.flat, reps):
        ts = per_country_ts.get(country)
        if ts is None:
            ax.set_visible(False)
            continue
        ys = np.array(ts["years"])
        pim = np.array(ts["pim_produced_plus_intan"], dtype=float)
        cw = np.array(ts["cwon_pca"], dtype=float)
        mask = np.isfinite(pim) & np.isfinite(cw)
        if mask.sum() < 6:
            ax.set_visible(False); continue
        # Normalise to first common year for trajectory comparison
        i0 = np.where(mask)[0][0]
        pim_n = pim / pim[i0]; cw_n = cw / cw[i0]
        ax.plot(ys, pim_n, label=r"$\hat{K}_{tang} + \beta\hat{K}_{intan}$", c="C0")
        ax.plot(ys, cw_n, label="CWON produced", c="C3", ls="--")
        ax.set_title(country, fontsize=10)
        ax.grid(alpha=0.3)
        if ax is axes[0, 0]:
            ax.legend(fontsize=8)
    fig.suptitle("Fig C2. PIM(tangible+intangible) vs CWON produced capital (indexed to first obs)")
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figC2_trajectories.png"), dpi=140)
    plt.close(fig)

    # Fig C3: joint vs individual identification
    mm = df.dropna(subset=["mu_A", "mu_joint", "beta_D", "beta_joint"]).copy()
    fig, axes = plt.subplots(1, 2, figsize=(11, 5))
    axes[0].scatter(mm["mu_A"], mm["mu_joint"], s=20)
    for _, row in mm.iterrows():
        axes[0].annotate(row["iso3"], (row["mu_A"], row["mu_joint"]),
                         fontsize=6, alpha=0.6)
    lim = max(mm["mu_A"].max(), mm["mu_joint"].max()) * 1.05
    axes[0].plot([0, lim], [0, lim], "--k", alpha=0.5)
    axes[0].set_xlabel(r"$\mu$ from Candidate A (production only)")
    axes[0].set_ylabel(r"$\mu$ from joint (prod + wealth) identification")
    axes[0].set_title("(a) tempo parameter")
    axes[0].grid(alpha=0.3)

    axes[1].scatter(mm["beta_D"], mm["beta_joint"], s=20, c="C2")
    for _, row in mm.iterrows():
        axes[1].annotate(row["iso3"], (row["beta_D"], row["beta_joint"]),
                         fontsize=6, alpha=0.6)
    lim2 = max(mm["beta_D"].max(), mm["beta_joint"].max()) * 1.05
    axes[1].plot([0, lim2], [0, lim2], "--k", alpha=0.5)
    axes[1].set_xlabel(r"$\beta$ from Candidate D (production only)")
    axes[1].set_ylabel(r"$\beta$ from joint identification")
    axes[1].set_title("(b) intangible capital share")
    axes[1].grid(alpha=0.3)

    fig.suptitle("Fig C3. Joint vs individual identification of hidden parameters")
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figC3_joint_vs_single.png"), dpi=140)
    plt.close(fig)

    print("\n=== CWON INTEGRATION SUMMARY ===")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()

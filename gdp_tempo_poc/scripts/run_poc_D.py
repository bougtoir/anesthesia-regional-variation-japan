"""
Candidate D PoC -- intangible capital as the 'forgotten parameter' in GDP.

Hypothesis: standard production functions use only tangible capital K_tang
(PWT's rnna), omitting intangible capital K_intan (R&D, software, training,
brand equity, organisational capital). This is the exact analogue of the
'forgotten sigma' in the fertility tempo paper. Corrado-Hulten-Sichel and
subsequent work show K_intan is 30-100% of K_tang in advanced economies.

Specifications:
  M0 (baseline):    log Y = alpha log K_tang + (1-alpha) log LH + A
  M1 (fixed beta):  log Y = alpha log K_tang + beta log K_intan + (1-alpha-beta) log LH + A
                    beta = 0.10 (Corrado-Hulten-Sichel 2009 approximate)
  M2 (fit beta):    same as M1 but beta estimated per country

Intangible stock is built by PIM on R&D expenditure (World Bank
GB.XPD.RSDV.GD.ZS), extrapolated to missing years by country median and
scaled to GDP level. Depreciation delta_I = 0.15 (standard in CHS).
Steady-state initial stock K_I(0) = I_R(0) / (delta_I + g).

Evaluation mirrors PoC A:
  Test A: levels fit with decade-smoothed TFP
  Test B: growth-rate fit with constant TFP growth intercept    -- main
  Test C: not applicable (no external K_intan benchmark)
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
os.makedirs(DATA, exist_ok=True); os.makedirs(FIG, exist_ok=True)

PWT_PATH = "/home/ubuntu/gdp_tempo_data/pwt1001.dta"
RND_PATH = "/home/ubuntu/gdp_tempo_data/wb/rnd_gdp.json"

COUNTRIES = [
    "Australia","Austria","Belgium","Canada","Chile","China",
    "Colombia","Costa Rica","Czech Republic","Denmark","Estonia",
    "Finland","France","Germany","Greece","Hungary","Iceland",
    "Ireland","Israel","Italy","Japan","Republic of Korea",
    "Latvia","Lithuania","Luxembourg","Mexico","Netherlands",
    "New Zealand","Norway","Poland","Portugal","Slovakia",
    "Slovenia","Spain","Sweden","Switzerland","Turkey",
    "United Kingdom","United States","D.R. of the Congo",
]
ISO3 = {
    "Australia":"AUS","Austria":"AUT","Belgium":"BEL","Canada":"CAN","Chile":"CHL",
    "China":"CHN","Colombia":"COL","Costa Rica":"CRI","Czech Republic":"CZE",
    "Denmark":"DNK","Estonia":"EST","Finland":"FIN","France":"FRA","Germany":"DEU",
    "Greece":"GRC","Hungary":"HUN","Iceland":"ISL","Ireland":"IRL","Israel":"ISR",
    "Italy":"ITA","Japan":"JPN","Republic of Korea":"KOR","Latvia":"LVA",
    "Lithuania":"LTU","Luxembourg":"LUX","Mexico":"MEX","Netherlands":"NLD",
    "New Zealand":"NZL","Norway":"NOR","Poland":"POL","Portugal":"PRT",
    "Slovakia":"SVK","Slovenia":"SVN","Spain":"ESP","Sweden":"SWE",
    "Switzerland":"CHE","Turkey":"TUR","United Kingdom":"GBR","United States":"USA",
    "D.R. of the Congo":"COD",
}

DELTA_I = 0.15


def load_rnd():
    with open(RND_PATH) as fh:
        rows = json.load(fh)
    df = pd.DataFrame([
        {"iso3": r["countryiso3code"], "year": int(r["date"]),
         "rnd_gdp": r["value"]}
        for r in rows if r.get("value") is not None
    ])
    return df


def build_intan_stock(years, Y, rnd_share, g=0.03):
    """Perpetual inventory of R&D, filling missing years by nearest-year median share."""
    # Fill NaN in rnd_share by carrying most recent value or country median
    s = pd.Series(rnd_share).ffill().bfill()
    if s.isna().all():
        return None
    s = s.fillna(s.median()).values
    I_R = Y * s / 100.0
    # Initial stock from steady state using first-year investment
    K = np.zeros_like(I_R, dtype=float)
    K[0] = I_R[0] / (DELTA_I + g)
    for t in range(1, len(K)):
        K[t] = (1 - DELTA_I) * K[t-1] + I_R[t-1]
    return K


def dw(resid):
    d = np.diff(resid)
    return float(np.sum(d*d) / (np.sum(resid*resid) + 1e-30))


def test_B_growth_two(logY, logK1, logK2, logL, alpha, beta):
    dY = np.diff(logY); dK1 = np.diff(logK1); dK2 = np.diff(logK2); dL = np.diff(logL)
    w_L = 1 - alpha - beta
    pred = alpha*dK1 + beta*dK2 + w_L*dL
    g = np.mean(dY - pred)
    resid = dY - (g + pred)
    ss_res = np.sum(resid**2); ss_tot = np.sum((dY - dY.mean())**2)
    return {"r2": float(1 - ss_res/(ss_tot+1e-30)),
            "rmse_pp": float(np.sqrt(np.mean(resid**2))*100)}


def test_B_growth_one(logY, logK, logL, alpha):
    dY = np.diff(logY); dK = np.diff(logK); dL = np.diff(logL)
    pred = alpha*dK + (1-alpha)*dL
    g = np.mean(dY - pred)
    resid = dY - (g + pred)
    ss_res = np.sum(resid**2); ss_tot = np.sum((dY - dY.mean())**2)
    return {"r2": float(1 - ss_res/(ss_tot+1e-30)),
            "rmse_pp": float(np.sqrt(np.mean(resid**2))*100)}


def test_A_levels_two(logY, logK1, logK2, logL, alpha, beta):
    w_L = 1 - alpha - beta
    raw = logY - alpha*logK1 - beta*logK2 - w_L*logL
    d = np.arange(len(logY)) // 10
    smo = np.zeros_like(raw)
    for dd in np.unique(d):
        m = d == dd
        smo[m] = raw[m].mean()
    pred = alpha*logK1 + beta*logK2 + w_L*logL + smo
    resid = logY - pred
    return {"mape": float(np.mean(np.abs(np.expm1(resid)))*100),
            "dw": dw(resid)}


def test_A_levels_one(logY, logK, logL, alpha):
    raw = logY - alpha*logK - (1-alpha)*logL
    d = np.arange(len(logY)) // 10
    smo = np.zeros_like(raw)
    for dd in np.unique(d):
        m = d == dd
        smo[m] = raw[m].mean()
    pred = alpha*logK + (1-alpha)*logL + smo
    resid = logY - pred
    return {"mape": float(np.mean(np.abs(np.expm1(resid)))*100),
            "dw": dw(resid)}


def fit_beta(logY, logK1, logK2, logL, alpha):
    best = None
    for beta in np.arange(0.0, 0.35, 0.01):
        if beta + alpha >= 0.95: continue
        r = test_B_growth_two(logY, logK1, logK2, logL, alpha, beta)
        if best is None or r["rmse_pp"] < best[0]:
            best = (r["rmse_pp"], beta)
    return best[1]


def main():
    print("Loading PWT...", flush=True)
    df = pd.read_stata(PWT_PATH)
    df = df[df["country"].isin(COUNTRIES)].copy()
    print("Loading WB R&D...", flush=True)
    rnd = load_rnd()

    rows = []
    for country in COUNTRIES:
        iso = ISO3[country]
        g = df[df["country"] == country].dropna(
            subset=["rgdpna","rnna","emp","labsh"]
        ).sort_values("year").reset_index(drop=True)
        r = rnd[rnd["iso3"] == iso].set_index("year")["rnd_gdp"] if iso else pd.Series(dtype=float)
        if len(g) < 30 or r.empty:
            print(f"  skip {country} (no R&D)"); continue
        # restrict to years after R&D first non-missing, but allow NaN forward/back fill
        years = g["year"].values.astype(int)
        rnd_share = np.array([r.get(y, np.nan) for y in years], dtype=float)
        # Skip countries with <5 R&D observations
        if np.sum(~np.isnan(rnd_share)) < 5:
            print(f"  skip {country} (<5 R&D obs)"); continue

        Y = g["rgdpna"].values.astype(float)
        K_tang = g["rnna"].values.astype(float)
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

        K_intan = build_intan_stock(years, Y, rnd_share)
        if K_intan is None:
            print(f"  skip {country}"); continue

        logY, logK, logL = np.log(Y), np.log(K_tang), np.log(L)
        logKi = np.log(np.where(K_intan > 0, K_intan, 1e-6))

        # M0: baseline (no intangible)
        A0 = test_A_levels_one(logY, logK, logL, alpha)
        B0 = test_B_growth_one(logY, logK, logL, alpha)
        # M1: fixed beta = 0.10
        beta1 = 0.10
        A1 = test_A_levels_two(logY, logK, logKi, logL, alpha, beta1)
        B1 = test_B_growth_two(logY, logK, logKi, logL, alpha, beta1)
        # M2: fit beta
        beta2 = fit_beta(logY, logK, logKi, logL, alpha)
        A2 = test_A_levels_two(logY, logK, logKi, logL, alpha, beta2)
        B2 = test_B_growth_two(logY, logK, logKi, logL, alpha, beta2)

        # intangible / tangible ratio (end of sample)
        ratio_end = float(K_intan[-1] / K_tang[-1])

        out = {"country": country, "n": len(years),
               "y_start": int(years[0]), "y_end": int(years[-1]),
               "alpha": alpha, "beta2": beta2,
               "Ki_over_K_end": ratio_end,
               "M0_A_mape": A0["mape"], "M0_A_dw": A0["dw"],
               "M0_B_rmse": B0["rmse_pp"], "M0_B_r2": B0["r2"],
               "M1_A_mape": A1["mape"], "M1_A_dw": A1["dw"],
               "M1_B_rmse": B1["rmse_pp"], "M1_B_r2": B1["r2"],
               "M2_A_mape": A2["mape"], "M2_A_dw": A2["dw"],
               "M2_B_rmse": B2["rmse_pp"], "M2_B_r2": B2["r2"]}
        rows.append(out)
        print(f"  {country:22s}  beta2={beta2:.2f}  Ki/K={ratio_end:.2f}  "
              f"B_rmse M0={B0['rmse_pp']:.2f} M1={B1['rmse_pp']:.2f} M2={B2['rmse_pp']:.2f}  "
              f"A_mape M0={A0['mape']:.2f} M2={A2['mape']:.2f}",
              flush=True)

    rdf = pd.DataFrame(rows).sort_values("country").reset_index(drop=True)
    rdf.to_csv(os.path.join(DATA, "poc_D_results.csv"), index=False)

    # ---------- Figures ----------
    fig, ax = plt.subplots(figsize=(11, 9))
    s = rdf.sort_values("M0_B_rmse")
    y = np.arange(len(s))
    bw = 0.28
    ax.barh(y - bw, s["M0_B_rmse"], bw, label="M0 K_tang only", color="#888")
    ax.barh(y,       s["M1_B_rmse"], bw, label="M1 K_tang + K_intan (beta=0.10)", color="#4c72b0")
    ax.barh(y + bw, s["M2_B_rmse"], bw, label="M2 K_tang + K_intan (beta fit)", color="#dd8452")
    ax.set_yticks(y); ax.set_yticklabels(s["country"], fontsize=8)
    ax.set_xlabel("RMSE of 1-year GDP growth fit (pp)")
    ax.set_title(f"Candidate D: growth-rate fit with intangible capital ({len(rdf)} countries)")
    ax.legend(loc="lower right"); ax.grid(axis="x", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figD1_growth_rmse.png"), dpi=140)
    plt.close(fig)

    d10 = rdf["M0_B_rmse"] - rdf["M1_B_rmse"]
    d20 = rdf["M0_B_rmse"] - rdf["M2_B_rmse"]
    d21 = rdf["M1_B_rmse"] - rdf["M2_B_rmse"]
    fig, ax = plt.subplots(figsize=(8,5))
    ax.boxplot([d10, d20, d21], tick_labels=["M0-M1","M0-M2","M1-M2"])
    ax.axhline(0, color="red", lw=1)
    ax.set_ylabel("Growth RMSE reduction (pp)")
    ax.set_title(f"Candidate D: pairwise RMSE improvements ({len(rdf)} countries)")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figD2_improvements_box.png"), dpi=140)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8,6))
    sc = ax.scatter(rdf["Ki_over_K_end"], rdf["M0_B_rmse"] - rdf["M2_B_rmse"],
                    c=rdf["beta2"], cmap="viridis", s=50, edgecolor="k")
    for _, row in rdf.iterrows():
        ax.annotate(row["country"][:6],
                    (row["Ki_over_K_end"], row["M0_B_rmse"] - row["M2_B_rmse"]),
                    fontsize=7, alpha=0.7)
    ax.set_xlabel("Intangible / Tangible capital ratio (end of sample)")
    ax.set_ylabel("Growth RMSE reduction M0 -> M2 (pp)")
    ax.set_title("Where does intangible capital matter most?")
    plt.colorbar(sc, ax=ax, label="Fitted intangible share (beta)")
    ax.axhline(0, color="red", lw=1)
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figD3_beta_scatter.png"), dpi=140)
    plt.close(fig)

    summary = {
        "n_countries": int(len(rdf)),
        "test_B_median_rmse_pp": {
            "M0": float(rdf["M0_B_rmse"].median()),
            "M1": float(rdf["M1_B_rmse"].median()),
            "M2": float(rdf["M2_B_rmse"].median()),
        },
        "test_A_median_mape": {
            "M0": float(rdf["M0_A_mape"].median()),
            "M1": float(rdf["M1_A_mape"].median()),
            "M2": float(rdf["M2_A_mape"].median()),
        },
        "share_M1_beats_M0_B": float((rdf["M1_B_rmse"] < rdf["M0_B_rmse"]).mean()),
        "share_M2_beats_M0_B": float((rdf["M2_B_rmse"] < rdf["M0_B_rmse"]).mean()),
        "median_beta_fitted": float(rdf["beta2"].median()),
        "median_Ki_over_K_end": float(rdf["Ki_over_K_end"].median()),
    }
    with open(os.path.join(DATA, "poc_D_summary.json"), "w") as fh:
        json.dump(summary, fh, indent=2)
    print("\n=== D SUMMARY ===", flush=True)
    print(json.dumps(summary, indent=2), flush=True)


if __name__ == "__main__":
    main()

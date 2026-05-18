"""
Candidate B PoC -- 'same-living working population' tempo effect in GDP.

Hypothesis: the standard labour input L in a Cobb-Douglas production function
is biased because the mean age of entry into / exit from work has been shifting
(education expansion delays entry; pension/health improvements delay exit).
Period GDP is biased in exactly the same way the period TFR is biased by AFB
shifts in the fertility tempo story.

Three specifications are compared:

  M0 (baseline):     L = emp * avh * hc   (standard PWT)
  M1 (static ages):  L = [sum_a pi(a; mu_E*, mu_X*) N(a,t)] * avh * hc
                      with country-specific constants mu_E*, mu_X*
                      pi(a) = trapezoid: 0 below mu_E, ramp 5 yrs, flat to mu_X-5,
                      decline to 0 at mu_X.
  M2 (tempo ages):   L = [sum_a pi(a; mu_E(t), mu_X(t)) N(a,t)] * avh * hc
                      mu_E(t) = mu_E0 + mu_E1 (year - t0)
                      mu_X(t) = mu_X0 + mu_X1 (year - t0)

Evaluation (same as candidate A):
  Test A (levels, decade-smoothed TFP)
  Test B (growth-rate, constant TFP growth intercept)   -- main
  Test C n/a (no external "true L" reference analogous to PWT K)

Data: PWT 10.01 (GDP, K, emp, hc, avh, labsh) + World Bank WDI 5-year age
population counts (SP.POP.{0004...6064}.{MA,FE} and SP.POP.65UP.{MA,FE}.IN).
"""
import os, json, glob
import numpy as np
import pandas as pd
from scipy.optimize import minimize
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "data")
FIG = os.path.join(ROOT, "figures")
os.makedirs(DATA, exist_ok=True); os.makedirs(FIG, exist_ok=True)

PWT_PATH = "/home/ubuntu/gdp_tempo_data/pwt1001.dta"
WB_AGE_DIR = "/home/ubuntu/gdp_tempo_data/wb/age"

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

# PWT country name -> WB ISO3 (WB age-bin data is keyed by ISO3)
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

BINS = [(0,4),(5,9),(10,14),(15,19),(20,24),(25,29),(30,34),(35,39),
        (40,44),(45,49),(50,54),(55,59),(60,64),(65,85)]  # treat 65UP as 65-85


def load_wb_age():
    """Load WB age-bin population counts into dict[(iso3, year)] -> array over BINS."""
    dfs = []
    for f in glob.glob(os.path.join(WB_AGE_DIR, "*.json")):
        base = os.path.basename(f)  # e.g. SP.POP.0004.MA.json  or SP.POP.65UP.MA.IN.json
        parts = base.replace(".json","").split(".")
        # parts = ['SP','POP','0004','MA'] or ['SP','POP','65UP','MA','IN']
        agecode = parts[2]
        sx = parts[3]
        with open(f) as fh:
            rows = json.load(fh)
        for row in rows:
            v = row.get("value")
            if v is None: continue
            dfs.append({
                "iso3": row["countryiso3code"],
                "year": int(row["date"]),
                "bin": agecode,
                "sex": sx,
                "n": v,
            })
    df = pd.DataFrame(dfs)
    df = df.groupby(["iso3","year","bin"])["n"].sum().reset_index()
    # pivot to wide
    df = df.pivot_table(index=["iso3","year"], columns="bin", values="n").reset_index()
    return df


def participation_profile(ages_lo, ages_hi, mu_E, mu_X, ramp=5.0):
    """Piecewise linear participation profile, returning integrated share per bin."""
    pi = np.zeros(len(ages_lo))
    for i,(lo,hi) in enumerate(zip(ages_lo, ages_hi)):
        # integrate profile over [lo, hi+1)
        width = hi + 1 - lo
        if hi < mu_E or lo > mu_X:
            pi[i] = 0; continue
        # sample at 0.5-year resolution
        a = np.arange(lo, hi+1, 0.5)
        f = np.zeros_like(a)
        left  = (a - mu_E) / ramp
        right = (mu_X - a) / ramp
        f = np.clip(np.minimum(left, right), 0, 1)
        pi[i] = f.mean()
    return pi


def _eff_labor(N_bins, ages_lo, ages_hi, mu_E, mu_X):
    pi = participation_profile(ages_lo, ages_hi, mu_E, mu_X)
    return N_bins @ pi  # per-year scalar (or vector if N is 2D)


def dw(resid):
    d = np.diff(resid)
    return float(np.sum(d*d) / (np.sum(resid*resid) + 1e-30))


def test_B_growth(logY, logK, logL, alpha):
    dY = np.diff(logY); dK = np.diff(logK); dL = np.diff(logL)
    pred = alpha*dK + (1-alpha)*dL
    g = np.mean(dY - pred)
    resid = dY - (g + pred)
    ss_res = np.sum(resid**2)
    ss_tot = np.sum((dY - dY.mean())**2)
    return {"r2": float(1 - ss_res/(ss_tot+1e-30)),
            "rmse_pp": float(np.sqrt(np.mean(resid**2))*100)}


def test_A_levels(logY, logK, logL, alpha):
    raw = logY - alpha*logK - (1-alpha)*logL
    d = np.arange(len(logY)) // 10
    smo = np.zeros_like(raw)
    for dd in np.unique(d):
        m = (d == dd)
        smo[m] = raw[m].mean()
    pred = alpha*logK + (1-alpha)*logL + smo
    resid = logY - pred
    return {"mape": float(np.mean(np.abs(np.expm1(resid)))*100),
            "dw": dw(resid)}


def fit_static(N_bins, years, logY, logK, avh, hc, alpha, ages_lo, ages_hi):
    best = None
    for mu_E in np.arange(14, 25, 1.0):
        for mu_X in np.arange(55, 72, 1.0):
            Lpop = _eff_labor(N_bins, ages_lo, ages_hi, mu_E, mu_X)
            if np.any(Lpop <= 0): continue
            L = Lpop * avh * hc
            r = test_B_growth(logY, logK, np.log(L), alpha)
            if best is None or r["rmse_pp"] < best[0]:
                best = (r["rmse_pp"], mu_E, mu_X)
    return best[1], best[2]


def fit_tempo(N_bins, years, logY, logK, avh, hc, alpha, ages_lo, ages_hi):
    t0 = years[0]
    best = None
    for mu_E0 in np.arange(14, 24, 2.0):
        for mu_E1 in np.arange(-0.04, 0.12, 0.04):
            for mu_X0 in np.arange(55, 70, 2.0):
                for mu_X1 in np.arange(-0.04, 0.16, 0.04):
                    Lpop = np.zeros(len(years))
                    bad = False
                    for ti, y in enumerate(years):
                        muE = mu_E0 + mu_E1 * (y - t0)
                        muX = mu_X0 + mu_X1 * (y - t0)
                        if muX <= muE + 10: bad = True; break
                        pi = participation_profile(ages_lo, ages_hi, muE, muX)
                        Lpop[ti] = N_bins[ti] @ pi
                    if bad or np.any(Lpop <= 0): continue
                    L = Lpop * avh * hc
                    r = test_B_growth(logY, logK, np.log(L), alpha)
                    if best is None or r["rmse_pp"] < best[0]:
                        best = (r["rmse_pp"], mu_E0, mu_E1, mu_X0, mu_X1)
    return best[1], best[2], best[3], best[4]


def main():
    print("Loading PWT...", flush=True)
    df_pwt = pd.read_stata(PWT_PATH)
    df_pwt = df_pwt[df_pwt["country"].isin(COUNTRIES)].copy()

    print("Loading WB age bins...", flush=True)
    wb = load_wb_age()
    # Build bin labels matching our BINS order
    bin_labels = [f"{lo:04d}" if lo < 65 else "65UP" for (lo,_) in BINS]
    # WB uses labels like '0004','0509',...,'6064','65UP'
    wb_label_map = {'0004':'0004','0509':'0509','1014':'1014','1519':'1519',
                    '2024':'2024','2529':'2529','3034':'3034','3539':'3539',
                    '4044':'4044','4549':'4549','5054':'5054','5559':'5559',
                    '6064':'6064','65UP':'65UP'}
    wb_cols = [wb_label_map[b] for b in ['0004','0509','1014','1519','2024','2529',
                                          '3034','3539','4044','4549','5054','5559',
                                          '6064','65UP']]
    for c in wb_cols:
        if c not in wb.columns:
            wb[c] = np.nan

    rows = []
    ages_lo = np.array([lo for lo,_ in BINS])
    ages_hi = np.array([hi for _,hi in BINS])

    for country in COUNTRIES:
        iso = ISO3[country]
        g = df_pwt[df_pwt["country"] == country].dropna(
            subset=["rgdpna","rnna","emp","labsh"]
        ).sort_values("year").reset_index(drop=True)
        w = wb[wb["iso3"] == iso].sort_values("year")
        if len(g) < 30 or w.empty: 
            print(f"  skip {country}"); continue
        # join on year
        w = w.set_index("year")
        years = np.array([y for y in g["year"].values if y in w.index], dtype=int)
        if len(years) < 30:
            print(f"  skip {country} (insufficient WB years)"); continue
        g = g[g["year"].isin(years)].reset_index(drop=True)

        Y = g["rgdpna"].values.astype(float)
        K = g["rnna"].values.astype(float)
        hc = g["hc"].values.astype(float)
        if np.isnan(hc).any():
            hc = np.where(np.isnan(hc), np.nanmean(hc), hc)
        avh = g["avh"].values.astype(float)
        if np.isnan(avh).any():
            avh = np.where(np.isnan(avh), np.nanmean(avh), avh)
        emp = g["emp"].values.astype(float)
        labsh = np.clip(np.mean(g["labsh"].values), 0.40, 0.75)
        alpha = 1 - labsh

        # PWT baseline L (M0)
        L0 = emp * avh * hc

        # WB age counts (bins x years)
        N_bins = np.zeros((len(years), len(BINS)))
        for ti, y in enumerate(years):
            for ci, lab in enumerate(wb_cols):
                v = w.loc[y, lab] if y in w.index else np.nan
                N_bins[ti, ci] = v if not np.isnan(v) else 0
        if np.any(N_bins.sum(axis=1) <= 0):
            print(f"  skip {country} (WB bins missing)"); continue

        # Find country-specific constant ages M1
        mu_E_s, mu_X_s = fit_static(N_bins, years, np.log(Y), np.log(K), avh, hc, alpha, ages_lo, ages_hi)
        Lpop_M1 = _eff_labor(N_bins, ages_lo, ages_hi, mu_E_s, mu_X_s)
        L1 = Lpop_M1 * avh * hc

        # Fit tempo M2
        mu_E0, mu_E1, mu_X0, mu_X1 = fit_tempo(N_bins, years, np.log(Y), np.log(K), avh, hc, alpha, ages_lo, ages_hi)
        t0 = years[0]
        Lpop_M2 = np.array([N_bins[ti] @ participation_profile(ages_lo, ages_hi,
                                                               mu_E0 + mu_E1*(y-t0),
                                                               mu_X0 + mu_X1*(y-t0))
                            for ti,y in enumerate(years)])
        L2 = Lpop_M2 * avh * hc

        out = {"country": country, "n": len(years),
               "y_start": int(years[0]), "y_end": int(years[-1]),
               "alpha": float(alpha),
               "mu_E_static": float(mu_E_s), "mu_X_static": float(mu_X_s),
               "mu_E0": float(mu_E0), "mu_E1": float(mu_E1),
               "mu_X0": float(mu_X0), "mu_X1": float(mu_X1)}

        for name, L in [("M0", L0), ("M1", L1), ("M2", L2)]:
            logL = np.log(np.where(L > 0, L, 1e-6))
            A = test_A_levels(np.log(Y), np.log(K), logL, alpha)
            B = test_B_growth(np.log(Y), np.log(K), logL, alpha)
            out[f"{name}_A_mape"] = A["mape"]; out[f"{name}_A_dw"] = A["dw"]
            out[f"{name}_B_rmse"] = B["rmse_pp"]; out[f"{name}_B_r2"] = B["r2"]
        rows.append(out)
        print(f"  {country:22s}  mu_E {mu_E0:.1f}{mu_E1:+.3f}t  mu_X {mu_X0:.1f}{mu_X1:+.3f}t  "
              f"B_rmse M0={out['M0_B_rmse']:.2f} M1={out['M1_B_rmse']:.2f} M2={out['M2_B_rmse']:.2f}",
              flush=True)

    rdf = pd.DataFrame(rows).sort_values("country").reset_index(drop=True)
    rdf.to_csv(os.path.join(DATA, "poc_B_results.csv"), index=False)

    # ---------- Figures ----------
    fig, ax = plt.subplots(figsize=(11, 9))
    s = rdf.sort_values("M0_B_rmse")
    y = np.arange(len(s))
    bw = 0.28
    ax.barh(y - bw, s["M0_B_rmse"], bw, label="M0 PWT emp*avh*hc", color="#888")
    ax.barh(y,       s["M1_B_rmse"], bw, label="M1 static age-profile", color="#4c72b0")
    ax.barh(y + bw, s["M2_B_rmse"], bw, label="M2 tempo ages", color="#dd8452")
    ax.set_yticks(y); ax.set_yticklabels(s["country"], fontsize=8)
    ax.set_xlabel("RMSE of 1-year GDP growth fit (pp)")
    ax.set_title(f"Candidate B: growth-rate fit across L-constructions ({len(rdf)} countries)")
    ax.legend(loc="lower right"); ax.grid(axis="x", alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figB1_growth_rmse.png"), dpi=140)
    plt.close(fig)

    d10 = rdf["M0_B_rmse"] - rdf["M1_B_rmse"]
    d20 = rdf["M0_B_rmse"] - rdf["M2_B_rmse"]
    d21 = rdf["M1_B_rmse"] - rdf["M2_B_rmse"]
    fig, ax = plt.subplots(figsize=(8,5))
    ax.boxplot([d10, d20, d21], tick_labels=["M0-M1","M0-M2","M1-M2"])
    ax.axhline(0, color="red", lw=1)
    ax.set_ylabel("Growth RMSE reduction (pp)")
    ax.set_title(f"Candidate B: pairwise RMSE improvements ({len(rdf)} countries)")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figB2_improvements_box.png"), dpi=140)
    plt.close(fig)

    # Figure: tempo parameters
    fig, ax = plt.subplots(figsize=(8,6))
    sc = ax.scatter(rdf["mu_E1"]*10, rdf["mu_X1"]*10,
                    c=(rdf["M1_B_rmse"] - rdf["M2_B_rmse"]), cmap="RdYlGn",
                    s=50, edgecolor='k')
    for _, r in rdf.iterrows():
        ax.annotate(r["country"][:6], (r["mu_E1"]*10, r["mu_X1"]*10), fontsize=7, alpha=0.7)
    ax.axhline(0, color="gray", lw=1); ax.axvline(0, color="gray", lw=1)
    ax.set_xlabel("Entry-age drift (years per decade)")
    ax.set_ylabel("Exit-age drift (years per decade)")
    ax.set_title("Labor tempo parameters, by country")
    plt.colorbar(sc, ax=ax, label="RMSE reduction M1->M2 (pp)")
    ax.grid(alpha=0.3)
    fig.tight_layout()
    fig.savefig(os.path.join(FIG, "figB3_tempo_params.png"), dpi=140)
    plt.close(fig)

    summary = {
        "n_countries": int(len(rdf)),
        "test_B_median_rmse_pp": {
            "M0": float(rdf["M0_B_rmse"].median()),
            "M1": float(rdf["M1_B_rmse"].median()),
            "M2": float(rdf["M2_B_rmse"].median()),
        },
        "share_M1_beats_M0_B": float((rdf["M1_B_rmse"] < rdf["M0_B_rmse"]).mean()),
        "share_M2_beats_M0_B": float((rdf["M2_B_rmse"] < rdf["M0_B_rmse"]).mean()),
        "share_M2_beats_M1_B": float((rdf["M2_B_rmse"] < rdf["M1_B_rmse"]).mean()),
        "median_mu_E0": float(rdf["mu_E0"].median()),
        "median_mu_E1_per_decade": float(rdf["mu_E1"].median() * 10),
        "median_mu_X0": float(rdf["mu_X0"].median()),
        "median_mu_X1_per_decade": float(rdf["mu_X1"].median() * 10),
    }
    with open(os.path.join(DATA, "poc_B_summary.json"), "w") as fh:
        json.dump(summary, fh, indent=2)
    print("\n=== B SUMMARY ===", flush=True)
    print(json.dumps(summary, indent=2), flush=True)


if __name__ == "__main__":
    main()

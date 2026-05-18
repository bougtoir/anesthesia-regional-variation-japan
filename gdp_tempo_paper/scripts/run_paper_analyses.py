"""
Additional analyses for the GDP tempo-effect manuscript (RIW submission).

Builds on the per-candidate PoCs (gdp_tempo_poc) and the flow-stock joint
identification demonstrator (gdp_cwon_integration). The PoCs estimated
parameters on the full 1970-2019 sample; the manuscript needs:

  (1) out-of-sample prediction: train on 1970-2014, predict 2015-2019.
  (2) bootstrap 95% confidence intervals on (mu_joint, beta_joint).
  (3) M4 (joint-estimated mu, beta) evaluated under the same Test A/B
      growth-rate metrics used to compare M0/M1/M2/M3 -- so that Table 1
      is apples-to-apples.
  (4) gamma_price sensitivity: what happens to the Japan CWON anomaly if
      we let CWON PCA grow at rate gamma_price slower / faster than the
      PIM-implied reproducible-capital growth?

Reads:
  - /home/ubuntu/gdp_tempo_data/pwt1001.dta
  - /home/ubuntu/gdp_tempo_data/wb/rnd_gdp.json
  - /home/ubuntu/gdp_tempo_data/wb/cwon/{NW.PCA.TO,NW.HCA.TO,NW.TOW.TO}.json

Writes (under ./data/ and ./figures/ relative to this file's parent):
  - fair_eval.csv, fair_eval_summary.json
  - oos.csv, oos_summary.json
  - bootstrap_ci.csv
  - gamma_price.csv
  - figures/fig_m_ranking.png, fig_oos.png, fig_bootstrap.png,
    fig_gamma_price.png
"""
from __future__ import annotations

import json
import math
import os
from dataclasses import dataclass

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "data")
FIG = os.path.join(ROOT, "figures")
os.makedirs(DATA, exist_ok=True)
os.makedirs(FIG, exist_ok=True)

PWT_PATH = "/home/ubuntu/gdp_tempo_data/pwt1001.dta"
RND_PATH = "/home/ubuntu/gdp_tempo_data/wb/rnd_gdp.json"
CWON_DIR = "/home/ubuntu/gdp_tempo_data/wb/cwon"
DELTA_I = 0.15  # CHS intangible depreciation

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
    "Switzerland": "CHE", "Turkey": "TUR", "United Kingdom": "GBR",
    "United States": "USA",
}


# ----- PIM helpers (mirror the PoC scripts) -----
def geom_weights(mu: float, S: int = 12) -> np.ndarray:
    mu = max(mu, 0.01)
    theta = mu / (1.0 + mu)
    s = np.arange(S + 1)
    w = (1 - theta) * theta ** s
    return w / w.sum()


def pim_instant(I: np.ndarray, delta: np.ndarray, K0: float) -> np.ndarray:
    K = np.zeros_like(I, dtype=float)
    K[0] = K0
    for t in range(1, len(I)):
        K[t] = (1 - delta[t - 1]) * K[t - 1] + I[t - 1]
    return K


def pim_lagged(I: np.ndarray, delta: np.ndarray, K0: float, mu: float,
               S: int = 12) -> np.ndarray:
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


def pim_lagged_tempo(I: np.ndarray, delta: np.ndarray, K0: float,
                     mu0: float, mu1: float, years: np.ndarray,
                     S: int = 12) -> np.ndarray:
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


def build_intan_stock(Y: np.ndarray, rnd_share: np.ndarray,
                      g: float = 0.03) -> np.ndarray | None:
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


# ----- Test metrics -----
def test_B_growth(logY, logK, logLH, alpha):
    dY = np.diff(logY); dK = np.diff(logK); dLH = np.diff(logLH)
    pred = alpha * dK + (1 - alpha) * dLH
    g = np.mean(dY - pred)
    resid = dY - (g + pred)
    return float(np.sqrt(np.mean(resid ** 2)) * 100)  # pp


def test_A_levels(logY, logK, logLH, alpha):
    raw_tfp = logY - alpha * logK - (1 - alpha) * logLH
    decades = np.arange(len(logY)) // 10
    tfp_smooth = np.zeros_like(raw_tfp)
    for d in np.unique(decades):
        m = decades == d
        tfp_smooth[m] = raw_tfp[m].mean()
    resid = logY - (alpha * logK + (1 - alpha) * logLH + tfp_smooth)
    return float(np.mean(np.abs(np.expm1(resid))) * 100)  # MAPE %


def test_B_growth_intan(logY, logK_tang, logK_intan, logL, alpha, beta):
    dY = np.diff(logY); dK = np.diff(logK_tang); dI = np.diff(logK_intan)
    dL = np.diff(logL)
    w_L = 1 - alpha - beta
    pred = alpha * dK + beta * dI + w_L * dL
    g = np.mean(dY - pred)
    resid = dY - (g + pred)
    return float(np.sqrt(np.mean(resid ** 2)) * 100)


def test_A_levels_intan(logY, logK_tang, logK_intan, logL, alpha, beta):
    w_L = 1 - alpha - beta
    raw_tfp = logY - alpha * logK_tang - beta * logK_intan - w_L * logL
    decades = np.arange(len(logY)) // 10
    tfp_smooth = np.zeros_like(raw_tfp)
    for d in np.unique(decades):
        m = decades == d
        tfp_smooth[m] = raw_tfp[m].mean()
    resid = logY - (alpha * logK_tang + beta * logK_intan + w_L * logL
                    + tfp_smooth)
    return float(np.mean(np.abs(np.expm1(resid))) * 100)


# ----- Loaders -----
def load_rnd() -> pd.DataFrame:
    with open(RND_PATH) as fh:
        rows = json.load(fh)
    return pd.DataFrame([
        {"iso3": r["countryiso3code"], "year": int(r["date"]),
         "rnd_gdp": r["value"]}
        for r in rows if r.get("value") is not None
    ])


def load_cwon(code: str) -> dict[tuple[str, int], float]:
    path = os.path.join(CWON_DIR, f"{code}.json")
    with open(path) as fh:
        rows = json.load(fh)
    out = {}
    for r in rows:
        if r.get("value") is None:
            continue
        out[(r["countryiso3code"], int(r["date"]))] = float(r["value"])
    return out


def demean_logratio(hat, obs):
    mask = np.isfinite(hat) & np.isfinite(obs) & (hat > 0) & (obs > 0)
    if mask.sum() < 6:
        return None
    lhat = np.log(hat[mask]); lobs = np.log(obs[mask])
    d = (lhat - lhat.mean()) - (lobs - lobs.mean())
    return float(np.sqrt(np.mean(d ** 2)))


# ----- Per-country dataclass -----
@dataclass
class Country:
    country: str
    iso: str
    years: np.ndarray
    Y: np.ndarray
    I: np.ndarray
    delta: np.ndarray
    K0: float
    Kpwt: np.ndarray
    emp: np.ndarray
    avh: np.ndarray
    hc: np.ndarray
    labsh: np.ndarray
    rnd_share: np.ndarray
    pca: np.ndarray
    cwon_years: np.ndarray


def prepare_countries() -> list[Country]:
    pwt = pd.read_stata(PWT_PATH)
    pwt = pwt[pwt["country"].isin(COUNTRIES)].sort_values(["country", "year"])
    rnd = load_rnd()
    cwon_pca = load_cwon("NW.PCA.TO")
    out: list[Country] = []
    for country in COUNTRIES:
        iso = ISO3[country]
        g = pwt[pwt["country"] == country].dropna(
            subset=["rgdpna", "rnna", "emp", "csh_i", "delta", "labsh"]
        ).sort_values("year").reset_index(drop=True)
        if len(g) < 30:
            continue
        r = rnd[rnd["iso3"] == iso].set_index("year")["rnd_gdp"]
        if r.empty or r.notna().sum() < 5:
            continue
        years = g["year"].values.astype(int)
        Y = g["rgdpna"].values.astype(float)
        I = Y * g["csh_i"].values.astype(float)
        delta = g["delta"].values.astype(float)
        Kpwt = g["rnna"].values.astype(float)
        emp = g["emp"].values.astype(float)
        avh = g["avh"].values.astype(float)
        if np.isnan(avh).any():
            avh = np.where(np.isnan(avh), np.nanmean(avh), avh)
        hc = g["hc"].values.astype(float)
        if np.isnan(hc).any():
            hc = np.where(np.isnan(hc), np.nanmean(hc), hc)
        labsh = g["labsh"].values.astype(float)
        rnd_share = np.array([r.get(int(y), np.nan) for y in years],
                             dtype=float)
        cwon_years = np.arange(1995, 2021)
        pca = np.array([cwon_pca.get((iso, int(y)), np.nan)
                        for y in cwon_years])
        out.append(Country(country, iso, years, Y, I, delta, float(Kpwt[0]),
                           Kpwt, emp, avh, hc, labsh, rnd_share,
                           pca, cwon_years))
    return out


# ----- Analysis 1 & 3: fair evaluation of M0-M4 -----
def fit_mu_const(I, delta, K0, logY, logLH, alpha) -> float:
    best = (np.inf, 0.4)
    for mu in np.linspace(0.01, 6.0, 25):
        K = pim_lagged(I, delta, K0, mu)
        K = np.where(K > 0, K, 1e-6)
        r = test_B_growth(logY, np.log(K), logLH, alpha)
        if r < best[0]:
            best = (r, mu)
    return best[1]


def fit_tempo(I, delta, K0, logY, logLH, alpha, years) -> tuple[float, float]:
    best = (np.inf, 0.4, 0.0)
    for mu0 in np.linspace(0.05, 5.0, 10):
        for mu1 in np.linspace(-0.08, 0.12, 11):
            K = pim_lagged_tempo(I, delta, K0, mu0, mu1, years)
            K = np.where(K > 0, K, 1e-6)
            r = test_B_growth(logY, np.log(K), logLH, alpha)
            if r < best[0]:
                best = (r, mu0, mu1)
    return best[1], best[2]


def fit_beta_given_K(K_tang, K_intan, logY, logL, alpha):
    """Fit beta given a tangible-capital stock (supplied by caller).

    The caller is responsible for choosing the tangible construction
    (e.g. pim_instant for M3, pim_lagged(mu_joint) for M4). This avoids
    train/eval mismatches where the fitted beta is evaluated against a
    different tangible stock than the one it was optimized against.
    """
    logK_tang = np.log(np.where(K_tang > 0, K_tang, 1e-6))
    logK_intan = np.log(np.where(K_intan > 0, K_intan, 1e-6))
    best = (np.inf, 0.0)
    for beta in np.linspace(0.0, 0.34, 18):
        if alpha + beta >= 0.95:
            continue
        r = test_B_growth_intan(logY, logK_tang, logK_intan, logL, alpha, beta)
        if r < best[0]:
            best = (r, beta)
    return best[1]


def fit_joint(I, delta, K0, K_intan, logY, logL, alpha, ki, pca,
              lambda_w=0.3) -> tuple[float, float, float, float]:
    logI = np.log(np.where(K_intan > 0, K_intan, 1e-6))
    best = None
    for mu in np.linspace(0.01, 6.0, 25):
        K_m = pim_lagged(I, delta, K0, mu)
        K_m = np.where(K_m > 0, K_m, 1e-6)
        logK = np.log(K_m)
        aligned_m = np.array([K_m[ii] if ii is not None else np.nan
                              for ii in ki])
        for beta in np.linspace(0.0, 0.34, 18):
            if alpha + beta >= 0.95:
                continue
            # production growth loss
            dY = np.diff(logY); dK = np.diff(logK); dI = np.diff(logI)
            dL = np.diff(logL)
            w_L = 1 - alpha - beta
            pred = alpha * dK + beta * dI + w_L * dL
            g = np.mean(dY - pred)
            L_p = float(np.mean((dY - g - pred) ** 2))
            aligned = aligned_m + beta * np.array([
                K_intan[ii] if ii is not None else np.nan for ii in ki])
            rm = demean_logratio(aligned, pca)
            if rm is None:
                continue
            L_w = rm ** 2
            score = L_p + lambda_w * L_w
            if best is None or score < best[0]:
                best = (score, mu, beta, L_p, L_w)
    if best is None:
        return np.nan, np.nan, np.nan, np.nan
    return best[1], best[2], best[3], best[4]


def run_fair_eval(countries: list[Country]) -> pd.DataFrame:
    rows = []
    for c in countries:
        alpha = 1 - float(np.clip(np.mean(c.labsh), 0.40, 0.75))
        L = c.emp * c.avh
        LH = L * c.hc
        logY = np.log(c.Y); logLH = np.log(LH); logL = np.log(L)
        K_intan = build_intan_stock(c.Y, c.rnd_share)
        if K_intan is None:
            continue

        # M0 (instant)
        K_M0 = pim_instant(c.I, c.delta, c.K0)
        # M1 (constant lag)
        mu_star = fit_mu_const(c.I, c.delta, c.K0, logY, logLH, alpha)
        K_M1 = pim_lagged(c.I, c.delta, c.K0, mu_star)
        # M2 (tempo lag)
        mu0, mu1 = fit_tempo(c.I, c.delta, c.K0, logY, logLH, alpha, c.years)
        K_M2 = pim_lagged_tempo(c.I, c.delta, c.K0, mu0, mu1, c.years)
        # M3: beta added on top of M0 tangible (pim_instant)
        beta_M3 = fit_beta_given_K(K_M0, K_intan, logY, logL, alpha)
        # M4: joint identification using CWON
        idx_map = {int(y): ii for ii, y in enumerate(c.years)}
        ki = [idx_map.get(int(y), None) for y in c.cwon_years]
        mu_j, beta_j, Lp_j, Lw_j = fit_joint(c.I, c.delta, c.K0, K_intan,
                                             logY, logL, alpha, ki, c.pca)
        K_M4 = pim_lagged(c.I, c.delta, c.K0, float(mu_j)) if np.isfinite(mu_j) \
            else None

        out = {"country": c.country, "iso3": c.iso, "alpha": alpha,
               "mu_M1": mu_star, "mu_M2_0": mu0, "mu_M2_1": mu1,
               "beta_M3": beta_M3, "mu_M4": mu_j, "beta_M4": beta_j,
               "M4_Lp": Lp_j, "M4_Lw": Lw_j}

        K_intan_p = np.where(K_intan > 0, K_intan, 1e-6)
        logI = np.log(K_intan_p)

        for name, Ktang, beta in [("M0", K_M0, 0.0), ("M1", K_M1, 0.0),
                                  ("M2", K_M2, 0.0), ("M3", K_M0, beta_M3),
                                  ("M4", K_M4, beta_j)]:
            if Ktang is None:
                out[f"{name}_B_rmse"] = np.nan
                out[f"{name}_A_mape"] = np.nan
                continue
            Kp = np.where(Ktang > 0, Ktang, 1e-6)
            logK = np.log(Kp)
            if beta > 0:
                out[f"{name}_B_rmse"] = test_B_growth_intan(
                    logY, logK, logI, logL, alpha, beta)
                out[f"{name}_A_mape"] = test_A_levels_intan(
                    logY, logK, logI, logL, alpha, beta)
            else:
                out[f"{name}_B_rmse"] = test_B_growth(logY, logK, logLH, alpha)
                out[f"{name}_A_mape"] = test_A_levels(logY, logK, logLH, alpha)
        rows.append(out)
        print(f"  [fair] {c.country:22s}  M0={out['M0_B_rmse']:.2f}  "
              f"M2={out['M2_B_rmse']:.2f}  M3={out['M3_B_rmse']:.2f}  "
              f"M4={out['M4_B_rmse']:.2f}  (mu_j={mu_j:.2f} beta_j={beta_j:.2f})",
              flush=True)
    return pd.DataFrame(rows)


# ----- Analysis 2: out-of-sample -----
OOS_TEST_YEARS = (2015, 2016, 2017, 2018, 2019)


def run_oos(countries: list[Country]) -> pd.DataFrame:
    rows = []
    for c in countries:
        mask_train = c.years <= 2014
        mask_test = np.isin(c.years, OOS_TEST_YEARS)
        if mask_test.sum() < 3 or mask_train.sum() < 20:
            continue
        I_train = c.I[mask_train]
        delta_train = c.delta[mask_train]
        years_train = c.years[mask_train]
        Y_train = c.Y[mask_train]
        emp_t = c.emp[mask_train]; avh_t = c.avh[mask_train]
        hc_t = c.hc[mask_train]
        L_train = emp_t * avh_t * hc_t
        L_full = c.emp * c.avh * c.hc
        logY_train = np.log(Y_train); logLH_train = np.log(L_train)
        logL_train = np.log(emp_t * avh_t)
        alpha = 1 - float(np.clip(np.mean(c.labsh[mask_train]), 0.40, 0.75))
        K0 = c.K0

        K_intan_full = build_intan_stock(c.Y, c.rnd_share)
        if K_intan_full is None:
            continue
        K_intan_train = K_intan_full[mask_train]

        # Fit each model on training data only
        mu1_tr = fit_mu_const(I_train, delta_train, K0, logY_train,
                              logLH_train, alpha)
        mu0_tr, mu1e_tr = fit_tempo(I_train, delta_train, K0, logY_train,
                                    logLH_train, alpha, years_train)
        K_M3_train = pim_instant(I_train, delta_train, K0)
        beta3_tr = fit_beta_given_K(K_M3_train, K_intan_train,
                                    logY_train, logL_train, alpha)
        # joint fit: uses CWON which starts 1995 -> training window 1995..2014
        idx_map = {int(y): ii for ii, y in enumerate(c.years)}
        ki = []
        for y in c.cwon_years:
            ii = idx_map.get(int(y), None)
            ki.append(ii if (ii is not None and int(y) <= 2014) else None)
        mu4_tr, beta4_tr, _, _ = fit_joint(
            I_train, delta_train, K0, K_intan_train, logY_train,
            logL_train, alpha, ki[:len(np.arange(1995, 2015))],
            c.pca[c.cwon_years <= 2014])

        # Forecast Y on test years using trained params and TRUE
        # investment series (oracle: we assume we know I and L in future,
        # we only test whether the K mapping captures Y well).
        def forecast_level(Ktang_full, beta):
            logK = np.log(np.where(Ktang_full > 0, Ktang_full, 1e-6))
            if beta > 0:
                logI = np.log(np.where(K_intan_full > 0, K_intan_full, 1e-6))
                logL_full = np.log(c.emp * c.avh)
                w_L = 1 - alpha - beta
                raw_tfp = np.log(c.Y) - alpha * logK - beta * logI - w_L * logL_full
            else:
                logLH = np.log(L_full)
                raw_tfp = np.log(c.Y) - alpha * logK - (1 - alpha) * logLH
            # Use LAST TRAINING DECADE MEAN as TFP forecast
            train_dec_mask = (c.years >= 2005) & (c.years <= 2014)
            if train_dec_mask.sum() == 0:
                return None
            tfp_proj = float(np.mean(raw_tfp[train_dec_mask]))
            if beta > 0:
                pred_logY = (alpha * logK + beta * logI +
                             w_L * np.log(c.emp * c.avh) + tfp_proj)
            else:
                pred_logY = (alpha * logK + (1 - alpha) * np.log(L_full)
                             + tfp_proj)
            return pred_logY

        K_full = {
            "M0": pim_instant(c.I, c.delta, K0),
            "M1": pim_lagged(c.I, c.delta, K0, mu1_tr),
            "M2": pim_lagged_tempo(c.I, c.delta, K0, mu0_tr, mu1e_tr,
                                   c.years),
            "M3": pim_instant(c.I, c.delta, K0),
            "M4": pim_lagged(c.I, c.delta, K0, mu4_tr) if np.isfinite(mu4_tr)
            else None,
        }
        beta_by = {"M0": 0.0, "M1": 0.0, "M2": 0.0, "M3": beta3_tr,
                   "M4": beta4_tr if np.isfinite(beta4_tr) else 0.0}

        out = {"country": c.country, "iso3": c.iso}
        for name, K in K_full.items():
            if K is None:
                out[f"{name}_oos_mape"] = np.nan
                continue
            pred_logY = forecast_level(K, beta_by[name])
            if pred_logY is None:
                out[f"{name}_oos_mape"] = np.nan
                continue
            resid = np.log(c.Y)[mask_test] - pred_logY[mask_test]
            out[f"{name}_oos_mape"] = float(np.mean(np.abs(np.expm1(resid)))
                                            * 100)
        rows.append(out)
        print(f"  [oos]  {c.country:22s}  M0={out['M0_oos_mape']:.2f}  "
              f"M2={out['M2_oos_mape']:.2f}  M3={out['M3_oos_mape']:.2f}  "
              f"M4={out['M4_oos_mape']:.2f}", flush=True)
    return pd.DataFrame(rows)


# ----- Analysis 4: bootstrap CI -----
def run_bootstrap(countries: list[Country], n_boot: int = 120,
                  seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    rows = []
    for c in countries:
        alpha = 1 - float(np.clip(np.mean(c.labsh), 0.40, 0.75))
        L = c.emp * c.avh
        LH = L * c.hc
        logY = np.log(c.Y); logL = np.log(L)
        K_intan = build_intan_stock(c.Y, c.rnd_share)
        if K_intan is None:
            continue
        idx_map = {int(y): ii for ii, y in enumerate(c.years)}
        ki = [idx_map.get(int(y), None) for y in c.cwon_years]

        mu_central, beta_central, _, _ = fit_joint(
            c.I, c.delta, c.K0, K_intan, logY, logL, alpha, ki, c.pca)
        # Residual bootstrap on growth-rate residuals.
        K_m = pim_lagged(c.I, c.delta, c.K0,
                         mu_central if np.isfinite(mu_central) else 0.4)
        K_m = np.where(K_m > 0, K_m, 1e-6)
        logK = np.log(K_m)
        logI = np.log(np.where(K_intan > 0, K_intan, 1e-6))
        dY = np.diff(logY); dK = np.diff(logK); dI = np.diff(logI)
        dL = np.diff(logL)
        w_L = 1 - alpha - (beta_central if np.isfinite(beta_central) else 0.0)
        pred = alpha * dK + (beta_central if np.isfinite(beta_central) else 0.0) * dI + w_L * dL
        g0 = np.mean(dY - pred)
        resid = dY - (g0 + pred)

        mus = []; betas = []
        for _ in range(n_boot):
            resample_idx = rng.integers(0, len(resid), size=len(resid))
            dY_bs = g0 + pred + resid[resample_idx]
            # reconstruct pseudo logY path
            logY_bs = np.zeros_like(logY)
            logY_bs[0] = logY[0]
            logY_bs[1:] = logY[0] + np.cumsum(dY_bs)
            Y_bs = np.exp(logY_bs)
            I_bs = Y_bs * (c.I / c.Y)
            K_intan_bs = build_intan_stock(Y_bs, c.rnd_share)
            if K_intan_bs is None:
                continue
            mu_b, beta_b, _, _ = fit_joint(
                I_bs, c.delta, c.K0, K_intan_bs, np.log(Y_bs), logL, alpha,
                ki, c.pca)
            if np.isfinite(mu_b):
                mus.append(mu_b); betas.append(beta_b)
        if len(mus) < 20:
            continue
        rows.append({
            "country": c.country, "iso3": c.iso,
            "mu_central": mu_central, "beta_central": beta_central,
            "mu_lo": float(np.quantile(mus, 0.025)),
            "mu_hi": float(np.quantile(mus, 0.975)),
            "beta_lo": float(np.quantile(betas, 0.025)),
            "beta_hi": float(np.quantile(betas, 0.975)),
            "n_boot_valid": len(mus),
        })
        print(f"  [boot] {c.country:22s}  mu={mu_central:.2f} "
              f"[{rows[-1]['mu_lo']:.2f},{rows[-1]['mu_hi']:.2f}]  "
              f"beta={beta_central:.2f} "
              f"[{rows[-1]['beta_lo']:.2f},{rows[-1]['beta_hi']:.2f}]",
              flush=True)
    return pd.DataFrame(rows)


# ----- Analysis 5: gamma_price sensitivity -----
def run_gamma_price(countries: list[Country], fair: pd.DataFrame,
                    gammas=(-0.04, -0.02, 0.0, 0.02, 0.04)) -> pd.DataFrame:
    """For each country and each gamma, re-compute log(W_PIM / CWON_PCA_adj)
    where CWON_PCA_adj(t) = CWON_PCA(t) * exp(gamma * (t - 2010)).
    gamma > 0 = CWON prices revalued upward through time (quantity down).
    Report the change in log-ratio median and in Japan specifically.

    Uses the manuscript's own M4 joint-identification estimates (mu_M4, beta_M4)
    from run_fair_eval so that the sensitivity is consistent with the rest of
    the paper."""
    mu_by = dict(zip(fair["country"], fair["mu_M4"]))
    beta_by = dict(zip(fair["country"], fair["beta_M4"]))
    rows = []
    for c in countries:
        mu = float(mu_by.get(c.country, np.nan))
        beta = float(beta_by.get(c.country, np.nan))
        if not (np.isfinite(mu) and np.isfinite(beta)):
            continue
        K_intan = build_intan_stock(c.Y, c.rnd_share)
        if K_intan is None:
            continue
        K_tang = pim_lagged(c.I, c.delta, c.K0, mu)
        idx_map = {int(y): ii for ii, y in enumerate(c.years)}
        ki = [idx_map.get(int(y), None) for y in c.cwon_years]
        aligned = np.array([(K_tang[ii] + beta * K_intan[ii])
                            if ii is not None else np.nan for ii in ki])
        rec = {"country": c.country, "iso3": c.iso}
        for gamma in gammas:
            pca_adj = c.pca * np.exp(gamma * (c.cwon_years - 2010))
            rmse = demean_logratio(aligned, pca_adj)
            mask = np.isfinite(aligned) & np.isfinite(pca_adj)
            if mask.sum() >= 6:
                scale = float(np.nanmedian(aligned[mask] / pca_adj[mask]))
                log_ratio = float(np.log(scale))
            else:
                log_ratio = np.nan
            rec[f"rmse_g{gamma:+.2f}"] = rmse
            rec[f"logratio_g{gamma:+.2f}"] = log_ratio
        rows.append(rec)
        print(f"  [gamma] {c.country:22s}  "
              + "  ".join(f"g={g:+.2f}:lr={rec[f'logratio_g{g:+.2f}']:.2f}"
                          for g in gammas), flush=True)
    return pd.DataFrame(rows)


# ----- Analysis 6: Relational PIM (M5 / RPIM) -----
def fit_relational_pim(K_pim: np.ndarray, K_cwon: np.ndarray,
                      cwon_years: np.ndarray,
                      pim_years: np.ndarray) -> tuple[float, float, float, int]:
    """Brass relational model for capital accounting.

    Estimates  log K_PIM(t) = rho1 + rho2 * log K_CWON(t) + eps(t)
    on the overlapping years where both series are observed and positive.

    Returns (rho1, rho2, R2, n_obs).
    """
    idx_map = {int(y): i for i, y in enumerate(pim_years)}
    log_pim = []
    log_cwon = []
    for i, y in enumerate(cwon_years):
        j = idx_map.get(int(y))
        if j is None:
            continue
        kp = K_pim[j]
        kc = K_cwon[i]
        if np.isfinite(kp) and np.isfinite(kc) and kp > 0 and kc > 0:
            log_pim.append(np.log(kp))
            log_cwon.append(np.log(kc))
    if len(log_pim) < 6:
        return np.nan, np.nan, np.nan, 0
    lp = np.array(log_pim)
    lc = np.array(log_cwon)
    X = np.column_stack([np.ones_like(lc), lc])
    params, residuals, _, _ = np.linalg.lstsq(X, lp, rcond=None)
    rho1, rho2 = float(params[0]), float(params[1])
    ss_res = float(np.sum((lp - X @ params) ** 2))
    ss_tot = float(np.sum((lp - lp.mean()) ** 2))
    r2 = 1.0 - ss_res / ss_tot if ss_tot > 0 else np.nan
    return rho1, rho2, r2, len(lp)


def run_relational_pim(countries: list[Country],
                       fair: pd.DataFrame) -> pd.DataFrame:
    """For each country, fit the Brass relational model between PIM K and
    CWON produced capital under each model specification (M0, M1, M2, M4).
    Reports (rho1, rho2, R2) per model."""
    mu1_by = dict(zip(fair["country"], fair["mu_M1"]))
    mu20_by = dict(zip(fair["country"], fair["mu_M2_0"]))
    mu21_by = dict(zip(fair["country"], fair["mu_M2_1"]))
    mu4_by = dict(zip(fair["country"], fair["mu_M4"]))
    beta4_by = dict(zip(fair["country"], fair["beta_M4"]))
    rows = []
    for c in countries:
        K_intan = build_intan_stock(c.Y, c.rnd_share)
        if K_intan is None:
            continue
        K_M0 = pim_instant(c.I, c.delta, c.K0)
        mu1 = float(mu1_by.get(c.country, np.nan))
        K_M1 = pim_lagged(c.I, c.delta, c.K0, mu1) if np.isfinite(mu1) else None
        mu20 = float(mu20_by.get(c.country, np.nan))
        mu21 = float(mu21_by.get(c.country, np.nan))
        K_M2 = pim_lagged_tempo(c.I, c.delta, c.K0, mu20, mu21,
                                c.years) if np.isfinite(mu20) else None
        mu4 = float(mu4_by.get(c.country, np.nan))
        beta4 = float(beta4_by.get(c.country, np.nan))
        K_M4_tang = pim_lagged(c.I, c.delta, c.K0, mu4) if np.isfinite(mu4) else None
        if K_M4_tang is not None and np.isfinite(beta4) and beta4 > 0:
            K_M4_total = K_M4_tang + beta4 * K_intan
        elif K_M4_tang is not None:
            K_M4_total = K_M4_tang
        else:
            K_M4_total = None

        rec = {"country": c.country, "iso3": c.iso}
        for label, K in [("M0", K_M0), ("M1", K_M1), ("M2", K_M2),
                         ("M4", K_M4_total)]:
            if K is None:
                rec[f"{label}_rho1"] = np.nan
                rec[f"{label}_rho2"] = np.nan
                rec[f"{label}_R2"] = np.nan
                rec[f"{label}_n"] = 0
            else:
                r1, r2, rsq, n = fit_relational_pim(K, c.pca,
                                                     c.cwon_years, c.years)
                rec[f"{label}_rho1"] = r1
                rec[f"{label}_rho2"] = r2
                rec[f"{label}_R2"] = rsq
                rec[f"{label}_n"] = n
        rows.append(rec)
        print(f"  [rpim] {c.country:22s}  "
              f"M0: rho2={rec['M0_rho2']:.3f}  "
              f"M4: rho2={rec['M4_rho2']:.3f}  "
              f"R2={rec['M4_R2']:.3f}",
              flush=True)
    return pd.DataFrame(rows)


# ----- Analysis 7: delta-mu joint sensitivity -----
def run_delta_sensitivity(countries: list[Country],
                         delta_factors: tuple[float, ...] = (0.80, 0.90, 1.00,
                                                             1.10, 1.20)
                         ) -> pd.DataFrame:
    """For each country, re-fit mu (M1 constant lag) under adjusted delta.

    Addresses Inklaar's critique: if delta drifts, some of what is attributed
    to mu(t) may instead belong to delta(t). Reports mu_hat for each
    delta_factor."""
    rows = []
    for c in countries:
        alpha = 1 - float(np.clip(np.mean(c.labsh), 0.40, 0.75))
        L = c.emp * c.avh
        LH = L * c.hc
        logY = np.log(c.Y)
        logLH = np.log(LH)
        rec = {"country": c.country, "iso3": c.iso, "alpha": alpha}
        for df in delta_factors:
            delta_adj = c.delta * df
            mu_hat = fit_mu_const(c.I, delta_adj, c.K0, logY, logLH, alpha)
            rec[f"mu_d{df:.2f}"] = mu_hat
        # also fit tempo (mu0, mu1) under each delta factor
        for df in delta_factors:
            delta_adj = c.delta * df
            mu0, mu1 = fit_tempo(c.I, delta_adj, c.K0, logY, logLH, alpha,
                                 c.years)
            rec[f"mu0_d{df:.2f}"] = mu0
            rec[f"mu1_d{df:.2f}"] = mu1
        rows.append(rec)
        base_mu = rec["mu_d1.00"]
        lo_mu = rec["mu_d0.80"]
        hi_mu = rec["mu_d1.20"]
        print(f"  [dsens] {c.country:22s}  "
              f"mu(d*0.8)={lo_mu:.2f}  mu(d*1.0)={base_mu:.2f}  "
              f"mu(d*1.2)={hi_mu:.2f}",
              flush=True)
    return pd.DataFrame(rows)


# ----- Figures -----
def make_figures(fair: pd.DataFrame, oos: pd.DataFrame,
                 boot: pd.DataFrame, gamma: pd.DataFrame):
    # Fig 1: M0-M4 Test B RMSE ranking
    if not fair.empty:
        s = fair.sort_values("M0_B_rmse")
        y = np.arange(len(s))
        fig, ax = plt.subplots(figsize=(11, 9))
        bw = 0.18
        for i, (name, color) in enumerate([
                ("M0", "#888888"), ("M1", "#4c72b0"),
                ("M2", "#dd8452"), ("M3", "#55a868"),
                ("M4", "#c44e52")]):
            ax.barh(y + (i - 2) * bw, s[f"{name}_B_rmse"].values, bw,
                    label=name, color=color)
        ax.set_yticks(y); ax.set_yticklabels(s["country"].values, fontsize=8)
        ax.set_xlabel("Test B: 1-year GDP growth fit RMSE (pp)")
        ax.set_title("Fig 1. In-sample growth-rate fit across five K-constructions")
        ax.legend(loc="lower right")
        plt.tight_layout()
        plt.savefig(os.path.join(FIG, "fig1_m_ranking.png"), dpi=180)
        plt.close()

    # Fig 2: OOS
    if not oos.empty:
        cols = ["M0_oos_mape", "M1_oos_mape", "M2_oos_mape",
                "M3_oos_mape", "M4_oos_mape"]
        fig, ax = plt.subplots(figsize=(8, 5))
        data = [oos[col].dropna().values for col in cols]
        ax.boxplot(data, labels=["M0", "M1", "M2", "M3", "M4"],
                   showmeans=True)
        ax.set_ylabel("Out-of-sample MAPE 2015-19 (%)")
        ax.set_title("Fig 2. Out-of-sample forecast accuracy, 39 countries")
        ax.grid(axis="y", alpha=0.3)
        plt.tight_layout()
        plt.savefig(os.path.join(FIG, "fig2_oos.png"), dpi=180)
        plt.close()

    # Fig 3: Bootstrap CIs for mu_joint and beta_joint
    if not boot.empty:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 9),
                                       sharey=True)
        s = boot.sort_values("mu_central").reset_index(drop=True)
        y = np.arange(len(s))
        ax1.errorbar(s["mu_central"], y,
                     xerr=[s["mu_central"] - s["mu_lo"],
                           s["mu_hi"] - s["mu_central"]],
                     fmt="o", ms=3, color="#c44e52",
                     ecolor="#c44e52", alpha=0.6)
        ax1.set_yticks(y); ax1.set_yticklabels(s["country"].values, fontsize=7)
        ax1.set_xlabel(r"$\hat{\mu}_{\mathrm{joint}}$ (years)")
        ax1.set_title("Joint tempo parameter")
        ax1.axvline(s["mu_central"].median(), c="k", lw=0.5, ls="--")
        ax2.errorbar(s["beta_central"], y,
                     xerr=[s["beta_central"] - s["beta_lo"],
                           s["beta_hi"] - s["beta_central"]],
                     fmt="o", ms=3, color="#4c72b0",
                     ecolor="#4c72b0", alpha=0.6)
        ax2.set_xlabel(r"$\hat{\beta}_{\mathrm{joint}}$")
        ax2.set_title("Joint intangible share")
        ax2.axvline(s["beta_central"].median(), c="k", lw=0.5, ls="--")
        fig.suptitle("Fig 3. Bootstrap 95% CI for joint-identified (μ, β)")
        plt.tight_layout()
        plt.savefig(os.path.join(FIG, "fig3_bootstrap.png"), dpi=180)
        plt.close()

    # Fig 4: gamma_price sensitivity for Japan & USA
    if not gamma.empty:
        gammas = [-0.04, -0.02, 0.0, 0.02, 0.04]
        highlight = ["Japan", "United States", "Germany",
                     "Republic of Korea", "Netherlands"]
        fig, ax = plt.subplots(figsize=(8, 5))
        for country in highlight:
            row = gamma[gamma["country"] == country]
            if row.empty:
                continue
            vals = [float(row[f"logratio_g{g:+.2f}"].iloc[0]) for g in gammas]
            ax.plot(gammas, vals, "o-", label=country)
        ax.axhline(0, c="k", lw=0.5)
        ax.set_xlabel(r"$\gamma_{price}$ (annual price re-evaluation rate)")
        ax.set_ylabel(r"$\log(W_{PIM}/\mathrm{CWON\ PCA}_{adj})$")
        ax.set_title("Fig 4. γ_price sensitivity of the PIM-CWON gap")
        ax.legend()
        ax.grid(alpha=0.3)
        plt.tight_layout()
        plt.savefig(os.path.join(FIG, "fig4_gamma_price.png"), dpi=180)
        plt.close()


def main():
    print("Preparing country panel...", flush=True)
    countries = prepare_countries()
    print(f"  {len(countries)} countries with complete coverage", flush=True)

    print("\n--- Analysis 1 + 3: fair evaluation M0-M4 ---", flush=True)
    fair = run_fair_eval(countries)
    fair.to_csv(os.path.join(DATA, "fair_eval.csv"), index=False)
    fair_summary = {
        name: {
            "B_rmse_median": float(fair[f"{name}_B_rmse"].median()),
            "A_mape_median": float(fair[f"{name}_A_mape"].median()),
        } for name in ("M0", "M1", "M2", "M3", "M4")
    }
    with open(os.path.join(DATA, "fair_eval_summary.json"), "w") as fh:
        json.dump(fair_summary, fh, indent=2)

    print("\n--- Analysis 2: out-of-sample ---", flush=True)
    oos = run_oos(countries)
    oos.to_csv(os.path.join(DATA, "oos.csv"), index=False)
    oos_summary = {
        name: float(oos[f"{name}_oos_mape"].median())
        for name in ("M0", "M1", "M2", "M3", "M4")
        if f"{name}_oos_mape" in oos.columns
    }
    with open(os.path.join(DATA, "oos_summary.json"), "w") as fh:
        json.dump(oos_summary, fh, indent=2)

    print("\n--- Analysis 4: bootstrap CI ---", flush=True)
    boot = run_bootstrap(countries, n_boot=100)
    boot.to_csv(os.path.join(DATA, "bootstrap_ci.csv"), index=False)

    print("\n--- Analysis 5: gamma_price sensitivity ---", flush=True)
    gamma = run_gamma_price(countries, fair)
    gamma.to_csv(os.path.join(DATA, "gamma_price.csv"), index=False)

    print("\n--- Analysis 6: Relational PIM (M5) ---", flush=True)
    rpim = run_relational_pim(countries, fair)
    rpim.to_csv(os.path.join(DATA, "rpim.csv"), index=False)
    rpim_summary = {
        label: {
            "rho2_median": float(rpim[f"{label}_rho2"].median()),
            "rho2_mean": float(rpim[f"{label}_rho2"].mean()),
            "rho1_median": float(rpim[f"{label}_rho1"].median()),
            "R2_median": float(rpim[f"{label}_R2"].median()),
        } for label in ("M0", "M1", "M2", "M4")
    }
    with open(os.path.join(DATA, "rpim_summary.json"), "w") as fh:
        json.dump(rpim_summary, fh, indent=2)

    print("\n--- Analysis 7: delta-mu sensitivity ---", flush=True)
    dsens = run_delta_sensitivity(countries)
    dsens.to_csv(os.path.join(DATA, "delta_sensitivity.csv"), index=False)
    dsens_summary = {
        f"d{df:.2f}": {
            "mu_median": float(dsens[f"mu_d{df:.2f}"].median()),
            "mu_mean": float(dsens[f"mu_d{df:.2f}"].mean()),
        } for df in (0.80, 0.90, 1.00, 1.10, 1.20)
    }
    with open(os.path.join(DATA, "delta_sensitivity_summary.json"), "w") as fh:
        json.dump(dsens_summary, fh, indent=2)

    print("\n--- Figures ---", flush=True)
    make_figures(fair, oos, boot, gamma)

    print("\n=== SUMMARY ===", flush=True)
    print(json.dumps(fair_summary, indent=2))
    print("OOS medians:", json.dumps(oos_summary, indent=2))
    print("RPIM summary:", json.dumps(rpim_summary, indent=2))
    print("Delta-mu sensitivity:", json.dumps(dsens_summary, indent=2))


if __name__ == "__main__":
    main()

"""
Healthcare Expenditure as Economic Effect: Neutral Sustainability Framework

Integrates:
  1. Input-Output (I-O) multiplier analysis -- healthcare spending as demand stimulus
  2. Health-Led Growth Hypothesis (HLGH) -- bidirectional causality evidence
  3. Tempo-effect health-capital model (from healthcare_tempo_poc) -- lag structure
  4. Net fiscal balance -- taxes/contributions generated vs public expenditure
  5. Three-layer tempo analogy (Population -> GDP -> Healthcare)
     from companion papers: Onishi (2026a) population, (2026b) GDP, (2026c) this paper

Produces bilingual (EN/JA) figures and data for the Japanese/English manuscripts.

Data: World Bank WDI via API, OECD health statistics (summary), published I-O studies,
      healthcare_tempo_poc Candidate A-H results.
"""

import os
import json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.font_manager as fm

# Register Noto Sans CJK JP for Japanese figures if available
_noto_cjk = [f.fname for f in fm.fontManager.ttflist if "Noto Sans CJK JP" in f.name]
if _noto_cjk:
    _JA_FONT_PROP = fm.FontProperties(fname=_noto_cjk[0])
    _JA_FONT_NAME = _JA_FONT_PROP.get_name()
else:
    _JA_FONT_PROP = None
    _JA_FONT_NAME = None

from contextlib import contextmanager

@contextmanager
def _ja_font_ctx():
    """Temporarily switch matplotlib font to Noto Sans CJK JP for JA figures."""
    if _JA_FONT_NAME:
        old = matplotlib.rcParams.get("font.family", ["sans-serif"])
        matplotlib.rcParams["font.family"] = [_JA_FONT_NAME]
        try:
            yield
        finally:
            matplotlib.rcParams["font.family"] = old
    else:
        yield
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "data")
FIG = os.path.join(ROOT, "output", "figures")
os.makedirs(DATA, exist_ok=True)
os.makedirs(FIG, exist_ok=True)

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.titlesize": 13,
    "axes.labelsize": 11,
    "xtick.labelsize": 9,
    "ytick.labelsize": 9,
    "legend.fontsize": 9,
    "figure.dpi": 200,
})

# ---------------------------------------------------------------------------
# 1. Published I-O multiplier data (healthcare sector)
#
# Country selection criteria:
#   (a) Direct I-O study of the healthcare sector published in a
#       peer-reviewed journal or official government report, OR
#   (b) Backward-linkage coefficient from national I-O tables
#       reported in the EU-28 I-O framework study
#       (Gutierrez-Hernandez & Abasolo-Alesson 2021).
#
# All OECD countries with publicly available I-O multiplier evidence
# are included.  Countries without a direct study are omitted rather
# than imputed.
# ---------------------------------------------------------------------------
IO_MULTIPLIERS = pd.DataFrame([
    {"country": "Japan", "iso3": "JPN", "multiplier": 2.78, "ci_lo": 2.74,
     "ci_hi": 2.90, "year": 2011, "source": "Yamada & Imanaka 2015"},
    {"country": "Japan (JMARI)", "iso3": "JPN", "multiplier": 2.85, "ci_lo": None,
     "ci_hi": None, "year": 2006, "source": "Maeda 2008 (JMARI WP172)"},
    {"country": "France", "iso3": "FRA", "multiplier": 2.20, "ci_lo": None,
     "ci_hi": None, "year": 2010,
     "source": "Gutierrez-Hernandez & Abasolo-Alesson 2021"},
    {"country": "Germany", "iso3": "DEU", "multiplier": 2.10, "ci_lo": None,
     "ci_hi": None, "year": 2014,
     "source": "Henke & Ostwald 2012 (GGR estimate)"},
    {"country": "Sweden", "iso3": "SWE", "multiplier": 2.05, "ci_lo": None,
     "ci_hi": None, "year": 2010,
     "source": "Gutierrez-Hernandez & Abasolo-Alesson 2021"},
    {"country": "Netherlands", "iso3": "NLD", "multiplier": 2.00, "ci_lo": None,
     "ci_hi": None, "year": 2010,
     "source": "Gutierrez-Hernandez & Abasolo-Alesson 2021"},
    {"country": "Italy", "iso3": "ITA", "multiplier": 1.95, "ci_lo": None,
     "ci_hi": None, "year": 2010,
     "source": "Gutierrez-Hernandez & Abasolo-Alesson 2021"},
    {"country": "Korea", "iso3": "KOR", "multiplier": 1.95, "ci_lo": None,
     "ci_hi": None, "year": 2015,
     "source": "Bank of Korea I-O Tables 2019"},
    {"country": "United Kingdom", "iso3": "GBR", "multiplier": 1.90, "ci_lo": None,
     "ci_hi": None, "year": 2016,
     "source": "ONS Health SUT / King's Fund 2018"},
    {"country": "Finland", "iso3": "FIN", "multiplier": 1.88, "ci_lo": None,
     "ci_hi": None, "year": 2010,
     "source": "Gutierrez-Hernandez & Abasolo-Alesson 2021"},
    {"country": "Spain", "iso3": "ESP", "multiplier": 1.85, "ci_lo": None,
     "ci_hi": None, "year": 2010,
     "source": "Gutierrez-Hernandez & Abasolo-Alesson 2021"},
    {"country": "Australia", "iso3": "AUS", "multiplier": 1.85, "ci_lo": None,
     "ci_hi": None, "year": 2015,
     "source": "AIHW / Deloitte Access Economics 2016"},
    {"country": "Canada", "iso3": "CAN", "multiplier": 1.82, "ci_lo": None,
     "ci_hi": None, "year": 2009, "source": "CIHI / Conference Board 2013"},
    {"country": "United States (Medicare)", "iso3": "USA", "multiplier": 1.70,
     "ci_lo": None, "ci_hi": None, "year": 2017,
     "source": "Dupor & Guerrero 2017 (Fed StL)"},
    {"country": "OECD Average", "iso3": "OECD", "multiplier": 1.95, "ci_lo": 1.50,
     "ci_hi": 2.90, "year": 2020,
     "source": "Synthesis (this study)"},
])


# ---------------------------------------------------------------------------
# 2. Health-Led Growth Hypothesis -- evidence summary
# ---------------------------------------------------------------------------
HLGH_EVIDENCE = pd.DataFrame([
    {"study": "Ertugrul et al. 2024", "n_countries": 38,
     "period": "2000-2019", "method": "CS-ARDL / AMG",
     "elasticity_h2g": 0.12, "elasticity_g2h": 0.65,
     "direction": "Bidirectional", "journal": "Front Public Health"},
    {"study": "Beylik et al. 2022", "n_countries": 21,
     "period": "2000-2018", "method": "Driscoll-Kraay",
     "elasticity_h2g": 0.08, "elasticity_g2h": 0.71,
     "direction": "Bidirectional", "journal": "Front Public Health"},
    {"study": "Amiri & Ventelou 2012", "n_countries": 20,
     "period": "1995-2008", "method": "Toda-Yamamoto",
     "elasticity_h2g": None, "elasticity_g2h": None,
     "direction": "Bidirectional (Granger)", "journal": "Econ Lett"},
    {"study": "Wang 2011", "n_countries": 31,
     "period": "1986-2007", "method": "Panel VECM",
     "elasticity_h2g": 0.10, "elasticity_g2h": 0.80,
     "direction": "Bidirectional", "journal": "Soc Sci Med"},
    {"study": "Piabuo & Tieguhong 2017", "n_countries": 45,
     "period": "1995-2015", "method": "GMM",
     "elasticity_h2g": 0.05, "elasticity_g2h": 0.55,
     "direction": "H->G (developing)", "journal": "BMC Res Notes"},
])


# ---------------------------------------------------------------------------
# 3. WB data for cross-country scatter (CHE %GDP vs Life Exp)
# ---------------------------------------------------------------------------
def fetch_wb_indicator(indicator, year=2019):
    """Fetch a single WB indicator for the latest available year via CSV API."""
    import urllib.request
    url = (f"https://api.worldbank.org/v2/country/all/indicator/{indicator}"
           f"?date={year}&format=json&per_page=300")
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            data = json.loads(resp.read().decode())
        if len(data) < 2:
            return {}
        return {r["countryiso3code"]: r["value"]
                for r in data[1] if r.get("value") is not None}
    except Exception as e:
        print(f"  [WARN] WB fetch {indicator}: {e}")
        return {}


OECD_ISO3 = [
    "AUS","AUT","BEL","CAN","CHL","COL","CRI","CZE","DNK","EST",
    "FIN","FRA","DEU","GRC","HUN","ISL","IRL","ISR","ITA","JPN",
    "KOR","LVA","LTU","LUX","MEX","NLD","NZL","NOR","POL","PRT",
    "SVK","SVN","ESP","SWE","CHE","TUR","GBR","USA",
]

COUNTRY_LABELS = {
    "JPN": "Japan", "USA": "USA", "DEU": "Germany", "GBR": "UK",
    "FRA": "France", "CAN": "Canada", "AUS": "Australia", "KOR": "Korea",
    "ITA": "Italy", "ESP": "Spain", "SWE": "Sweden", "NOR": "Norway",
    "CHE": "Switzerland", "NLD": "Netherlands", "FIN": "Finland",
}


def build_cross_country_df():
    """Build a DataFrame with CHE %GDP and LifeExp for OECD countries."""
    print("Fetching WB data ...")
    che_gdp = fetch_wb_indicator("SH.XPD.CHEX.GD.ZS", 2019)
    le = fetch_wb_indicator("SP.DYN.LE00.IN", 2019)
    che_pc = fetch_wb_indicator("SH.XPD.CHEX.PP.CD", 2019)

    rows = []
    for iso in OECD_ISO3:
        if iso in che_gdp and iso in le:
            rows.append({
                "iso3": iso,
                "label": COUNTRY_LABELS.get(iso, iso),
                "che_gdp_pct": che_gdp[iso],
                "life_exp": le[iso],
                "che_pc_ppp": che_pc.get(iso),
            })
    df = pd.DataFrame(rows)
    df.to_csv(os.path.join(DATA, "oecd_cross_country_2019.csv"), index=False)
    print(f"  {len(df)} countries with complete data.")
    return df


# ---------------------------------------------------------------------------
# 4. Net economic return model
# ---------------------------------------------------------------------------
def compute_neutral_sustainability(multiplier, tax_rate, public_share):
    return (tax_rate * multiplier) / public_share


def sustainability_table():
    """Build a table for 15 countries with I-O multiplier evidence."""
    params = [
        # iso, name, m, tau (effective tax+SSC/GDP), pf (public share of CHE)
        ("JPN", "Japan",        2.78, 0.33, 0.84),
        ("USA", "USA",          1.70, 0.27, 0.50),
        ("DEU", "Germany",      2.10, 0.39, 0.85),
        ("GBR", "UK",           1.90, 0.33, 0.80),
        ("FRA", "France",       2.20, 0.45, 0.84),
        ("SWE", "Sweden",       2.05, 0.43, 0.85),
        ("CAN", "Canada",       1.82, 0.33, 0.73),
        ("AUS", "Australia",    1.85, 0.28, 0.68),
        ("KOR", "Korea",        1.95, 0.27, 0.61),
        ("ITA", "Italy",        1.95, 0.43, 0.74),
        ("ESP", "Spain",        1.85, 0.35, 0.71),
        ("NLD", "Netherlands",  2.00, 0.39, 0.82),
        ("FIN", "Finland",      1.88, 0.43, 0.78),
    ]
    rows = []
    for iso, name, m, tau, pf in params:
        ratio = compute_neutral_sustainability(m, tau, pf)
        rows.append({
            "iso3": iso, "country": name,
            "io_multiplier": m,
            "eff_tax_rate": tau,
            "public_share_che": pf,
            "fiscal_return_ratio": round(ratio, 2),
            "sustainable": "Yes" if ratio >= 1.0 else "No",
        })
    df = pd.DataFrame(rows)
    df.to_csv(os.path.join(DATA, "neutral_sustainability.csv"), index=False)
    return df


# ---------------------------------------------------------------------------
# 5. Candidate A-H PoC results (from healthcare_tempo_poc PR#37)
#    "Candidate A-H" = healthcare_tempo_poc's model specification A-H
#    (Health spending -> outcome lag with tempo drift)
# ---------------------------------------------------------------------------
POC_AH_RESULTS = {
    "n_countries": 39,
    "data_source": "World Bank WDI (SH.XPD.CHEX.PP.CD, SP.DYN.LE00.IN)",
    "period": "2000-2019",
    "models": {
        "M0_flow": {"description": "Naive flow-only regression",
                    "level_rmse_median": 0.510,
                    "change_rmse_median": 0.455},
        "M1_constant_lag": {"description": "Constant lag mu_H (PIM)",
                            "level_rmse_median": 0.441,
                            "change_rmse_median": 0.403,
                            "mu_const_median_yr": 4.0},
        "M2_tempo_lag": {"description": "Time-varying mu_H(t) = mu0 + mu1*(year-t0)",
                         "level_rmse_median": 0.434,
                         "change_rmse_median": 0.405,
                         "mu_H1_median_yr_per_yr": 0.15},
    },
    "key_findings": {
        "M1_beats_M0_level_pct": 69,
        "M2_beats_M0_level_pct": 77,
        "M2_beats_M0_change_pct": 87,
        "M2_beats_M1_pct": 95,
        "M0_to_M1_rmse_reduction_pct": 14,
        "M0_to_M2_rmse_reduction_pct": 15,
    },
    "interpretation": (
        "M2 beats M1 in 95% of countries, confirming that the "
        "spending-to-outcome lag is not constant but drifts over time. "
        "Median drift mu_H1 = +0.15 yr/yr means the lag lengthens by "
        "~1.5 years per decade, consistent with the shift from acute "
        "to chronic disease management and longer R&D-to-outcome cycles."
    ),
}

# Three-layer tempo analogy: Population -> GDP -> Healthcare
THREE_LAYER_ANALOGY = pd.DataFrame([
    {"concept": "Flow (quantum)",
     "population": "TFR (period fertility rate)",
     "gdp": "I/GDP (investment rate)",
     "healthcare": "E/GDP (health spending rate)"},
    {"concept": "Tempo (timing lag)",
     "population": "MAC (mean age at childbearing)",
     "gdp": "mu (investment-to-output lag)",
     "healthcare": "mu_H (spending-to-outcome lag)"},
    {"concept": "Forgotten parameter",
     "population": "sigma (parity variance)",
     "gdp": "beta (intangible capital share)",
     "healthcare": "lambda_b (composition multipliers)"},
    {"concept": "Stock",
     "population": "Cohort size N(t)",
     "gdp": "Capital stock K(t)",
     "healthcare": "Health capital H(t)"},
    {"concept": "Tempo drift (mu_1)",
     "population": "+0.05 yr/yr (MAC shift)",
     "gdp": "+0.04 yr/yr (time-to-build)",
     "healthcare": "+0.15 yr/yr (spending-to-outcome)"},
    {"concept": "Effect size vs M0",
     "population": "Large (TFR bias ~15-20%)",
     "gdp": "Small (MAPE -0.6 pp)",
     "healthcare": "Medium (RMSE -15%)"},
    {"concept": "Identity",
     "population": "Renewal equation",
     "gdp": "dW/dt = S(Y) - delta*W",
     "healthcare": "dH/dt = sum(lambda_b*E_b) - delta_H*H"},
])


def tempo_adjusted_narrative():
    """Combine PoC A-H results with tempo narrative for manuscripts."""
    poc = POC_AH_RESULTS
    return {
        "poc_summary": poc,
        "three_layer_analogy": THREE_LAYER_ANALOGY.to_dict(orient="records"),
        "key_insight": (
            "The healthcare_tempo_poc (model specification A-H, 39 countries) shows "
            "that treating health spending as a stock-building flow with a "
            "time-varying lag mu_H(t) reduces life-expectancy prediction "
            f"RMSE from {poc['models']['M0_flow']['level_rmse_median']:.3f} "
            f"to {poc['models']['M2_tempo_lag']['level_rmse_median']:.3f} years "
            f"(-{poc['key_findings']['M0_to_M2_rmse_reduction_pct']}%). "
            "M2 beats M1 in 95% of countries, confirming that the lag is "
            "not constant but drifts at +0.15 yr/yr."
        ),
        "policy_implication": (
            "A 'neutral' sustainability criterion must account for both "
            "channels: (1) the contemporaneous fiscal multiplier effect "
            "(I-O: does tax revenue from healthcare-induced economic "
            "activity cover public financing?), and (2) the intertemporal "
            "health-capital accumulation effect (tempo: does the stock "
            "of health built today justify the flow of spending?). "
            "Neither alone gives a complete picture."
        ),
        "us_japan_contrast": (
            "The US has high spending (17% GDP) but a low I-O multiplier "
            "(1.7), suggesting leakage through high drug prices and "
            "administrative costs. Japan has moderate spending (11% GDP) "
            "but the highest multiplier (2.78) and a fiscal return ratio "
            "above 1.0 (1.09). Through the tempo lens, the US pattern -- "
            "high flow, low stock accumulation -- mirrors 'high TFR, low "
            "cohort fertility' in demography: a tempo-inflated flow that "
            "overstates true investment in health capital."
        ),
        "three_layer_connection": (
            "The tempo-plus-forgotten-parameter framework originated in "
            "demography (Bongaarts-Feeney 1998, Goldstein-Lutz-Scherbov 2003), "
            "was ported to capital accounting (Onishi 2026b), and is here "
            "extended to healthcare. Healthcare shows the largest tempo "
            "drift (+0.15 yr/yr vs GDP +0.04) among the three layers."
        ),
    }


# ---------------------------------------------------------------------------
# 6. Figure generation
# ---------------------------------------------------------------------------
def fig1_io_multiplier_comparison(io_df):
    """Bar chart of I-O multipliers across countries."""
    fig, ax = plt.subplots(figsize=(9, 6))
    mask = io_df["iso3"] != "OECD"
    df = io_df[mask].sort_values("multiplier", ascending=True)

    colors = ["#FF5722" if iso == "JPN" else "#2196F3" for iso in df["iso3"]]
    bars = ax.barh(df["country"], df["multiplier"], color=colors, edgecolor="white")

    jpn_mask = df["iso3"] == "JPN"
    if jpn_mask.any():
        for _, jpn in df[jpn_mask].iterrows():
            if pd.notna(jpn.get("ci_lo")) and pd.notna(jpn.get("ci_hi")):
                idx = df.index.get_loc(jpn.name)
                ax.errorbar(jpn["multiplier"], idx,
                            xerr=[[jpn["multiplier"] - jpn["ci_lo"]],
                                  [jpn["ci_hi"] - jpn["multiplier"]]],
                            fmt="none", ecolor="black", capsize=4)

    ax.axvline(1.0, color="gray", linestyle="--", linewidth=0.8, label="Break-even (1.0)")
    oecd_row = io_df[io_df["iso3"] == "OECD"]
    if not oecd_row.empty:
        ax.axvline(oecd_row.iloc[0]["multiplier"], color="#4CAF50",
                   linestyle=":", linewidth=1.2,
                   label=f'OECD synthesis ({oecd_row.iloc[0]["multiplier"]})')

    ax.set_xlabel("Economic Impact Multiplier (output per unit of healthcare spending)")
    ax.set_title("Healthcare I-O Multipliers by Country")
    ax.legend(loc="lower right", fontsize=8)
    ax.set_xlim(0, 3.5)

    for bar, val in zip(bars, df["multiplier"]):
        ax.text(val + 0.05, bar.get_y() + bar.get_height() / 2,
                f"{val:.2f}", va="center", fontsize=9)

    plt.tight_layout()
    path = os.path.join(FIG, "fig1_io_multipliers.png")
    fig.savefig(path)
    plt.close(fig)
    print(f"  Saved: {path}")
    return path


def fig2_scatter_che_vs_le(cc_df):
    """Scatter: CHE %GDP vs Life Expectancy.

    Shows linear fit (US excluded) as primary, and quadratic fit (all countries)
    as dashed ghost to illustrate that the Preston-style curvature is driven
    entirely by the US outlier (F-test p<0.001 with US, p=0.49 without).
    """
    from scipy import stats as sp_stats

    fig, ax = plt.subplots(figsize=(9, 6.5))

    # Separate US from rest
    us_mask = cc_df["iso3"] == "USA"
    rest = cc_df[~us_mask]
    x_rest = rest["che_gdp_pct"].values
    y_rest = rest["life_exp"].values
    x_all = cc_df["che_gdp_pct"].values
    y_all = cc_df["life_exp"].values
    n_rest = len(x_rest)
    n_all = len(x_all)

    ax.scatter(rest["che_gdp_pct"], rest["life_exp"],
               s=60, alpha=0.7, c="#1976D2", edgecolors="white", linewidths=0.5)
    ax.scatter(cc_df.loc[us_mask, "che_gdp_pct"], cc_df.loc[us_mask, "life_exp"],
               s=80, alpha=0.9, c="#F44336", edgecolors="white", linewidths=0.5,
               marker="D", label="USA (excluded from fit)", zorder=5)

    for _, row in cc_df.iterrows():
        if row["label"] in COUNTRY_LABELS.values():
            ax.annotate(row["label"],
                        (row["che_gdp_pct"], row["life_exp"]),
                        fontsize=7, xytext=(5, 3),
                        textcoords="offset points")

    xfit = np.linspace(cc_df["che_gdp_pct"].min(), cc_df["che_gdp_pct"].max(), 100)

    # --- Linear fit excluding US (primary) ---
    z1 = np.polyfit(x_rest, y_rest, 1)
    yfit1 = np.polyval(z1, xfit)
    ss_res1 = np.sum((y_rest - np.polyval(z1, x_rest))**2)
    ss_tot = np.sum((y_rest - np.mean(y_rest))**2)
    r2_lin = 1 - ss_res1 / ss_tot
    ax.plot(xfit, yfit1, "r-", linewidth=1.5, alpha=0.8,
            label=f"Linear fit (excl. US), R\u00b2={r2_lin:.2f}")

    # --- Quadratic fit including US (ghost, for comparison) ---
    z2_all = np.polyfit(x_all, y_all, 2)
    yfit2_all = np.polyval(z2_all, xfit)
    ax.plot(xfit, yfit2_all, "--", color="#9E9E9E", linewidth=1.2, alpha=0.6,
            label="Quadratic fit (incl. US) — Preston-style")

    # --- F-test: is quadratic term significant? ---
    # Without US
    z2_rest = np.polyfit(x_rest, y_rest, 2)
    ss_res2_rest = np.sum((y_rest - np.polyval(z2_rest, x_rest))**2)
    f_wo = ((ss_res1 - ss_res2_rest) / 1) / (ss_res2_rest / (n_rest - 3))
    p_wo = 1 - sp_stats.f.cdf(f_wo, 1, n_rest - 3)

    # With US
    z1_all = np.polyfit(x_all, y_all, 1)
    ss_res1_all = np.sum((y_all - np.polyval(z1_all, x_all))**2)
    ss_res2_all = np.sum((y_all - np.polyval(z2_all, x_all))**2)
    f_w = ((ss_res1_all - ss_res2_all) / 1) / (ss_res2_all / (n_all - 3))
    p_w = 1 - sp_stats.f.cdf(f_w, 1, n_all - 3)

    # Annotation box with F-test results
    txt = (f"F-test for quadratic term:\n"
           f"  With US:    F={f_w:.1f}, p={p_w:.4f}\n"
           f"  Without US: F={f_wo:.1f}, p={p_wo:.2f}")
    ax.text(0.02, 0.02, txt, transform=ax.transAxes, fontsize=7.5,
            verticalalignment="bottom", fontfamily="monospace",
            bbox=dict(boxstyle="round,pad=0.4", facecolor="#FFF9C4",
                      edgecolor="#F9A825", alpha=0.85))

    ax.set_xlabel("Current Health Expenditure (% of GDP)")
    ax.set_ylabel("Life Expectancy at Birth (years)")
    ax.set_title("Healthcare Spending vs Life Expectancy (OECD, 2019)")
    ax.legend(fontsize=8, loc="upper left")
    plt.tight_layout()
    path = os.path.join(FIG, "fig2_che_vs_lifeexp.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")

    # Return stats for manuscript use
    overfit_stats = {
        "n_excl_us": int(n_rest), "n_incl_us": int(n_all),
        "r2_linear_excl_us": round(r2_lin, 4),
        "r2_quad_excl_us": round(1 - ss_res2_rest / ss_tot, 4),
        "f_excl_us": round(f_wo, 3), "p_excl_us": round(p_wo, 4),
        "f_incl_us": round(f_w, 3), "p_incl_us": round(p_w, 4),
    }
    return path, overfit_stats


def fig3_fiscal_sustainability(sust_df):
    """Bar chart: fiscal return ratio by country."""
    fig, ax = plt.subplots(figsize=(9, 6))
    df = sust_df.sort_values("fiscal_return_ratio", ascending=True)

    colors = ["#4CAF50" if r >= 1.0 else "#FF9800" for r in df["fiscal_return_ratio"]]
    bars = ax.barh(df["country"], df["fiscal_return_ratio"], color=colors, edgecolor="white")

    ax.axvline(1.0, color="red", linestyle="--", linewidth=1.2,
               label="Break-even (tau*m = pf)")
    ax.set_xlabel("Fiscal Return Ratio  tau*m / pf")
    ax.set_title("Neutral Fiscal Sustainability of Healthcare Spending")
    ax.legend(loc="lower right", fontsize=8)

    for bar, val in zip(bars, df["fiscal_return_ratio"]):
        ax.text(val + 0.02, bar.get_y() + bar.get_height() / 2,
                f"{val:.2f}", va="center", fontsize=9)

    plt.tight_layout()
    path = os.path.join(FIG, "fig3_fiscal_sustainability.png")
    fig.savefig(path)
    plt.close(fig)
    print(f"  Saved: {path}")
    return path


def fig4_dual_return_schematic():
    """Conceptual diagram: I-O (demand) + Tempo (supply) dual return.

    Redesigned: no overlapping arrows/boxes, sequential arrows merged.
    Layout: top row = two return boxes, middle = central expenditure box,
    bottom = combined sustainability box.  Single arrows connect each pair.
    """
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7)
    ax.axis("off")

    # --- Central box: Healthcare Expenditure (middle row) ---
    cx, cy, cw, ch = 3.2, 3.2, 3.6, 1.0
    ax.add_patch(FancyBboxPatch((cx, cy), cw, ch,
                                boxstyle="round,pad=0.15",
                                facecolor="#E3F2FD", edgecolor="#1565C0",
                                linewidth=2, zorder=2))
    ax.text(cx + cw / 2, cy + ch / 2, "Healthcare\nExpenditure E(t)",
            ha="center", va="center", fontsize=11, fontweight="bold", zorder=3)

    # --- Left box: Demand-side Return (top-left) ---
    lx, ly, lw, lh = 0.3, 5.3, 3.2, 1.2
    ax.add_patch(FancyBboxPatch((lx, ly), lw, lh,
                                boxstyle="round,pad=0.15",
                                facecolor="#BBDEFB", edgecolor="#1565C0",
                                linewidth=1.5))
    ax.text(lx + lw / 2, ly + lh / 2,
            "Demand-side Return\nI-O Multiplier m\nTax return: tau * m * E(t)",
            ha="center", va="center", fontsize=9, color="#0D47A1")

    # --- Right box: Supply-side Return (top-right) ---
    rx, ry, rw, rh = 6.5, 5.3, 3.2, 1.2
    ax.add_patch(FancyBboxPatch((rx, ry), rw, rh,
                                boxstyle="round,pad=0.15",
                                facecolor="#FFF3E0", edgecolor="#E65100",
                                linewidth=1.5))
    ax.text(rx + rw / 2, ry + rh / 2,
            "Supply-side Return\nHealth Capital H(t)\nTempo drift: mu_H(t)",
            ha="center", va="center", fontsize=9, color="#BF360C")

    # --- Bottom box: Combined sustainability ---
    bx, by, bw, bh = 1.5, 0.5, 7.0, 0.9
    ax.add_patch(FancyBboxPatch((bx, by), bw, bh,
                                boxstyle="round,pad=0.15",
                                facecolor="#E8F5E9", edgecolor="#388E3C",
                                linewidth=1.5))
    ax.text(bx + bw / 2, by + bh / 2,
            "Neutral Sustainability = Demand Return + Supply Return >= Public Cost",
            ha="center", va="center", fontsize=9.5, fontweight="bold", color="#1B5E20")

    # --- Arrows: single arrow from center to each top box ---
    ax.annotate("", xy=(lx + lw / 2, ly), xytext=(cx + 0.8, cy + ch),
                arrowprops=dict(arrowstyle="-|>", color="#1976D2", lw=2.0,
                                connectionstyle="arc3,rad=0.2"))

    ax.annotate("", xy=(rx + rw / 2, ry), xytext=(cx + cw - 0.8, cy + ch),
                arrowprops=dict(arrowstyle="-|>", color="#E65100", lw=2.0,
                                connectionstyle="arc3,rad=-0.2"))

    # --- Arrows: single arrow from each top box to bottom ---
    ax.annotate("", xy=(bx + bw * 0.25, by + bh), xytext=(lx + lw / 2, ly),
                arrowprops=dict(arrowstyle="-|>", color="#4CAF50", lw=1.8,
                                connectionstyle="arc3,rad=0.2"))

    ax.annotate("", xy=(bx + bw * 0.75, by + bh), xytext=(rx + rw / 2, ry),
                arrowprops=dict(arrowstyle="-|>", color="#FF9800", lw=1.8,
                                connectionstyle="arc3,rad=-0.2"))

    ax.set_title("Dual-Return Framework for Neutral Healthcare Sustainability",
                 fontsize=12, pad=15)
    plt.tight_layout()
    path = os.path.join(FIG, "fig4_dual_return_schematic.png")
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")
    return path


def fig5_three_layer_analogy(lang="en"):
    """Table-style figure showing the Population/GDP/Healthcare tempo analogy."""
    import matplotlib.font_manager as fm

    ja_font_name = "DejaVu Sans"
    if lang == "ja":
        ja_hits = [f for f in fm.fontManager.ttflist
                   if "IPAGothic" in f.name or "IPAPGothic" in f.name]
        if ja_hits:
            fm.fontManager.addfont(ja_hits[0].fname)
            ja_font_name = ja_hits[0].name

    fig, ax = plt.subplots(figsize=(10, 4.5))
    ax.axis("off")

    if lang == "ja":
        col_labels = ["概念", "人口", "GDP / 国富", "医療"]
        row_data = [
            ["フロー（量子）", "TFR\n（期間出生率）", "I/GDP\n（投資率）", "E/GDP\n（医療支出率）"],
            ["テンポ（時間ラグ）", "MAC\n（平均出産年齢）", "mu\n（投資→産出ラグ）", "mu_H\n（支出→成果ラグ）"],
            ["忘れられたパラメータ", "sigma\n（パリティ分散）", "beta\n（無形資本比率）", "lambda_b\n（構成乗数）"],
            ["ストック", "コーホート人口\nN(t)", "資本ストック\nK(t)", "健康資本\nH(t)"],
            ["テンポドリフト (mu_1)", "+0.05 年/年\n（MAC上昇）", "+0.04 年/年\n（建設期間延長）", "+0.15 年/年\n（支出→成果遅延）"],
            ["効果サイズ vs M0", "大（TFR偏り\n15-20%）", "小（MAPE\n-0.6 pp）", "中（RMSE\n-15%）"],
        ]
        title = "テンポ効果の三層構造 — 人口→GDP→医療への移植"
    else:
        col_labels = ["Concept", "Population", "GDP / Wealth", "Healthcare"]
        row_data = [
            ["Flow (quantum)", "TFR\n(period fertility)", "I/GDP\n(investment rate)", "E/GDP\n(health spend rate)"],
            ["Tempo (timing lag)", "MAC\n(mean age childbearing)", "mu\n(invest-to-output lag)", "mu_H\n(spend-to-outcome lag)"],
            ["Forgotten parameter", "sigma\n(parity variance)", "beta\n(intangible K share)", "lambda_b\n(composition mult.)"],
            ["Stock", "Cohort size\nN(t)", "Capital stock\nK(t)", "Health capital\nH(t)"],
            ["Tempo drift (mu_1)", "+0.05 yr/yr\n(MAC shift)", "+0.04 yr/yr\n(time-to-build)", "+0.15 yr/yr\n(spend-to-outcome)"],
            ["Effect size vs M0", "Large (TFR bias\n15-20%)", "Small (MAPE\n-0.6 pp)", "Medium (RMSE\n-15%)"],
        ]
        title = "Three-Layer Tempo Analogy: Population to GDP to Healthcare"

    colors_col = ["#E3F2FD", "#FCE4EC", "#FFF3E0", "#E8F5E9"]
    table = ax.table(
        cellText=row_data,
        colLabels=col_labels,
        cellLoc="center",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8.5)
    table.scale(1.0, 2.2)

    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor("#BDBDBD")
        if row == 0:
            cell.set_facecolor("#37474F")
            cell.set_text_props(color="white", fontweight="bold", fontsize=9.5,
                                fontfamily=ja_font_name)
        else:
            cell.set_facecolor(colors_col[col] if col < len(colors_col) else "#FFFFFF")
            cell.set_text_props(fontfamily=ja_font_name)
        if col == 3 and row > 0:
            cell.set_text_props(fontweight="bold", fontfamily=ja_font_name)

    ax.set_title(title, fontsize=12, pad=15, fontweight="bold",
                 fontfamily=ja_font_name)
    plt.tight_layout()
    suffix = "_ja" if lang == "ja" else ""
    path = os.path.join(FIG, f"fig5_three_layer_analogy{suffix}.png")
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")
    return path


# ---------------------------------------------------------------------------
# 8. Medical Equipment Stock & Import Leakage Analysis
#
# Sources:
#   CT/MRI: OECD Health at a Glance 2023, Table 5.23 (2021 or latest year)
#   Pharma trade: UN COMTRADE via Trading Economics; MHLW 薬事工業生産動態統計
#   Med-device trade: UN COMTRADE HS 9018-9022; MHLW statistics
#   CHE: OECD Health Statistics 2023
# ---------------------------------------------------------------------------
EQUIPMENT_STOCK = pd.DataFrame([
    # CT and MRI per million population (OECD HaG 2023, 2021 or latest)
    {"iso3": "JPN", "country": "Japan",       "ct_per_m": 115.7, "mri_per_m": 55.2,
     "che_bn_usd": 553.0},
    {"iso3": "USA", "country": "USA",         "ct_per_m": 44.9,  "mri_per_m": 40.4,
     "che_bn_usd": 4255.0},
    {"iso3": "DEU", "country": "Germany",     "ct_per_m": 35.1,  "mri_per_m": 34.7,
     "che_bn_usd": 440.0},
    {"iso3": "AUS", "country": "Australia",   "ct_per_m": 70.2,  "mri_per_m": 16.0,
     "che_bn_usd": 155.0},
    {"iso3": "KOR", "country": "Korea",       "ct_per_m": 41.1,  "mri_per_m": 33.5,
     "che_bn_usd": 135.0},
    {"iso3": "ITA", "country": "Italy",       "ct_per_m": 35.7,  "mri_per_m": 31.3,
     "che_bn_usd": 191.0},
    {"iso3": "FRA", "country": "France",      "ct_per_m": 18.0,  "mri_per_m": 16.1,
     "che_bn_usd": 310.0},
    {"iso3": "GBR", "country": "UK",          "ct_per_m": 9.5,   "mri_per_m": 7.2,
     "che_bn_usd": 283.0},
    {"iso3": "CAN", "country": "Canada",      "ct_per_m": 15.8,  "mri_per_m": 10.4,
     "che_bn_usd": 195.0},
    {"iso3": "SWE", "country": "Sweden",      "ct_per_m": 19.7,  "mri_per_m": 15.3,
     "che_bn_usd": 59.0},
    {"iso3": "ESP", "country": "Spain",       "ct_per_m": 19.6,  "mri_per_m": 17.1,
     "che_bn_usd": 140.0},
    {"iso3": "NLD", "country": "Netherlands", "ct_per_m": 14.4,  "mri_per_m": 13.5,
     "che_bn_usd": 107.0},
    {"iso3": "FIN", "country": "Finland",     "ct_per_m": 24.2,  "mri_per_m": 23.9,
     "che_bn_usd": 26.0},
])

# Medical trade balance (pharma HS30 + devices HS9018-22), bn USD, ~2019-2021
# Positive = net exporter, Negative = net importer
# Sources: UN COMTRADE, Trading Economics, MHLW, Eurostat
MEDICAL_TRADE = pd.DataFrame([
    {"iso3": "JPN", "country": "Japan",       "pharma_exp": 5.1,  "pharma_imp": 28.5,
     "device_exp": 7.8, "device_imp": 12.3},
    {"iso3": "USA", "country": "USA",         "pharma_exp": 52.0, "pharma_imp": 136.0,
     "device_exp": 44.0, "device_imp": 55.0},
    {"iso3": "DEU", "country": "Germany",     "pharma_exp": 93.0, "pharma_imp": 62.0,
     "device_exp": 28.0, "device_imp": 17.0},
    {"iso3": "AUS", "country": "Australia",   "pharma_exp": 4.5,  "pharma_imp": 12.0,
     "device_exp": 2.5,  "device_imp": 6.0},
    {"iso3": "KOR", "country": "Korea",       "pharma_exp": 8.5,  "pharma_imp": 8.0,
     "device_exp": 4.5,  "device_imp": 5.5},
    {"iso3": "ITA", "country": "Italy",       "pharma_exp": 37.0, "pharma_imp": 24.0,
     "device_exp": 11.0, "device_imp": 9.0},
    {"iso3": "FRA", "country": "France",      "pharma_exp": 36.0, "pharma_imp": 30.0,
     "device_exp": 8.0,  "device_imp": 12.0},
    {"iso3": "GBR", "country": "UK",          "pharma_exp": 30.0, "pharma_imp": 32.0,
     "device_exp": 7.0,  "device_imp": 11.0},
    {"iso3": "CAN", "country": "Canada",      "pharma_exp": 7.0,  "pharma_imp": 18.0,
     "device_exp": 2.5,  "device_imp": 8.0},
    {"iso3": "SWE", "country": "Sweden",      "pharma_exp": 12.0, "pharma_imp": 5.5,
     "device_exp": 3.5,  "device_imp": 3.0},
    {"iso3": "ESP", "country": "Spain",       "pharma_exp": 14.0, "pharma_imp": 18.0,
     "device_exp": 2.5,  "device_imp": 5.5},
    {"iso3": "NLD", "country": "Netherlands", "pharma_exp": 42.0, "pharma_imp": 28.0,
     "device_exp": 10.0, "device_imp": 12.0},
    {"iso3": "FIN", "country": "Finland",     "pharma_exp": 2.0,  "pharma_imp": 2.5,
     "device_exp": 0.8,  "device_imp": 1.5},
])


def build_equipment_trade_df():
    """Merge equipment stock, trade, and I-O data into a single analysis frame."""
    eq = EQUIPMENT_STOCK.copy()
    tr = MEDICAL_TRADE.copy()

    # Compute combined diagnostic equipment density
    eq["equip_density"] = eq["ct_per_m"] + eq["mri_per_m"]

    # Compute trade balance
    tr["trade_balance_bn"] = (
        (tr["pharma_exp"] + tr["device_exp"])
        - (tr["pharma_imp"] + tr["device_imp"])
    )
    tr["total_med_imports_bn"] = tr["pharma_imp"] + tr["device_imp"]
    tr["total_med_exports_bn"] = tr["pharma_exp"] + tr["device_exp"]

    merged = eq.merge(tr[["iso3", "trade_balance_bn", "total_med_imports_bn",
                          "total_med_exports_bn"]], on="iso3", how="left")

    # Import leakage ratio = net imports / CHE
    merged["import_leakage"] = np.clip(
        (-merged["trade_balance_bn"]) / merged["che_bn_usd"], 0, 1
    )

    # Merge I-O multipliers (take first entry per iso3, skip OECD average)
    io = IO_MULTIPLIERS[IO_MULTIPLIERS["iso3"] != "OECD"].copy()
    io = io.drop_duplicates(subset="iso3", keep="first")
    merged = merged.merge(io[["iso3", "multiplier"]], on="iso3", how="left")

    # Effective multiplier adjusted for import leakage
    merged["effective_multiplier"] = merged["multiplier"] * (1 - merged["import_leakage"])

    return merged


def japan_counterfactual(eq_trade_df):
    """Compute counterfactual scenarios for Japan.

    Scenario A: Japan with OECD-average equipment density (holds other vars)
    Scenario B: Japan with zero net medical imports (domestic manufacturing)
    Returns dict of results.
    """
    jpn = eq_trade_df[eq_trade_df["iso3"] == "JPN"].iloc[0]
    oecd_avg_density = eq_trade_df["equip_density"].median()

    # Baseline
    baseline = {
        "equip_density": jpn["equip_density"],
        "import_leakage": jpn["import_leakage"],
        "multiplier": jpn["multiplier"],
        "effective_multiplier": jpn["effective_multiplier"],
    }
    # Fiscal return: tau * m_eff / pf
    tau_jpn, pf_jpn = 0.33, 0.84
    baseline["fiscal_return"] = tau_jpn * baseline["effective_multiplier"] / pf_jpn

    # Scenario A: OECD-average equipment density
    # Equipment contributes to health capital but also represents spending.
    # With less equipment, CHE would be lower (assume equipment-related
    # spending is proportional to density ratio).
    # Estimate: ~15% of CHE is capital/equipment-related (OECD average)
    equip_che_share = 0.15
    density_ratio = oecd_avg_density / jpn["equip_density"]
    scenario_a = baseline.copy()
    scenario_a["equip_density"] = oecd_avg_density
    scenario_a["label"] = "A: OECD-average equipment"
    # Lower equipment → lower CHE → but also lower diagnostic capability
    # Net effect on multiplier: reduced domestic production of equipment-related services
    # Assume multiplier decreases proportionally to the equipment share reduction
    equip_reduction = (1 - density_ratio) * equip_che_share
    scenario_a["multiplier_adj"] = jpn["multiplier"] * (1 - equip_reduction * 0.5)
    scenario_a["effective_multiplier"] = scenario_a["multiplier_adj"] * (
        1 - jpn["import_leakage"])
    scenario_a["fiscal_return"] = tau_jpn * scenario_a["effective_multiplier"] / pf_jpn

    # Scenario B: Domestic manufacturing (zero net imports)
    scenario_b = baseline.copy()
    scenario_b["import_leakage"] = 0.0
    scenario_b["label"] = "B: Domestic manufacturing"
    scenario_b["effective_multiplier"] = jpn["multiplier"] * 1.0
    scenario_b["fiscal_return"] = tau_jpn * scenario_b["effective_multiplier"] / pf_jpn

    # Scenario C: Both (average equipment + domestic manufacturing)
    scenario_c = baseline.copy()
    scenario_c["equip_density"] = oecd_avg_density
    scenario_c["import_leakage"] = 0.0
    scenario_c["label"] = "C: Average equip + domestic mfg"
    scenario_c["multiplier_adj"] = scenario_a["multiplier_adj"]
    scenario_c["effective_multiplier"] = scenario_c["multiplier_adj"]
    scenario_c["fiscal_return"] = tau_jpn * scenario_c["effective_multiplier"] / pf_jpn

    return {
        "baseline": baseline,
        "scenario_a": scenario_a,
        "scenario_b": scenario_b,
        "scenario_c": scenario_c,
        "oecd_avg_density": oecd_avg_density,
    }


COUNTRY_JA = {
    "Japan": "日本", "USA": "米国", "Germany": "ドイツ",
    "Australia": "豪州", "Korea": "韓国", "Italy": "イタリア",
    "France": "フランス", "UK": "英国", "Canada": "カナダ",
    "Sweden": "スウェーデン", "Spain": "スペイン",
    "Netherlands": "オランダ", "Finland": "フィンランド",
}


def fig6_equipment_density_comparison(eq_trade_df, lang="en"):
    """Bar chart: combined CT+MRI density by country, with OECD average line."""
    df = eq_trade_df.sort_values("equip_density", ascending=True).copy()
    if lang == "ja":
        df["label"] = df["country"].map(COUNTRY_JA)
    else:
        df["label"] = df["country"]
    fig, ax = plt.subplots(figsize=(9, 6))

    colors = ["#F44336" if iso == "JPN" else "#1976D2" for iso in df["iso3"]]
    bars = ax.barh(df["label"], df["equip_density"], color=colors, edgecolor="white",
                   height=0.7)

    oecd_avg = df["equip_density"].median()
    lbl = f"中央値 = {oecd_avg:.0f}" if lang == "ja" else f"Median = {oecd_avg:.0f}"
    ax.axvline(oecd_avg, color="#FF9800", linestyle="--", linewidth=1.5, label=lbl)

    for bar, val in zip(bars, df["equip_density"]):
        ax.text(val + 1, bar.get_y() + bar.get_height() / 2,
                f"{val:.0f}", va="center", fontsize=8)

    if lang == "ja":
        ax.set_xlabel("CT + MRI 合計台数（人口100万人あたり）")
        ax.set_title("画像診断装置密度（OECD, 2021年）")
    else:
        ax.set_xlabel("Combined CT + MRI Units per Million Population")
        ax.set_title("Diagnostic Imaging Equipment Density (OECD, 2021)")
    ax.legend(fontsize=9)
    plt.tight_layout()
    suffix = "_ja" if lang == "ja" else ""
    path = os.path.join(FIG, f"fig6_equipment_density{suffix}.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")
    return path


def fig7_import_leakage_vs_multiplier(eq_trade_df, lang="en"):
    """Scatter: import leakage ratio vs effective multiplier, sized by CHE."""
    fig, ax = plt.subplots(figsize=(9, 6.5))

    df = eq_trade_df.dropna(subset=["multiplier"]).copy()
    if lang == "ja":
        df["label"] = df["country"].map(COUNTRY_JA)
    else:
        df["label"] = df["country"]
    sizes = (df["che_bn_usd"] / df["che_bn_usd"].max()) * 400 + 40

    for _, row in df.iterrows():
        c = "#F44336" if row["iso3"] == "JPN" else "#1976D2"
        mk = "D" if row["iso3"] == "USA" else "o"
        ax.scatter(row["import_leakage"] * 100, row["effective_multiplier"],
                   s=sizes[_], alpha=0.7, c=c, edgecolors="white",
                   linewidths=0.5, marker=mk, zorder=3)
        ax.annotate(row["label"],
                    (row["import_leakage"] * 100, row["effective_multiplier"]),
                    fontsize=7, xytext=(5, 3), textcoords="offset points")

    # Show raw multiplier as faded markers
    for _, row in df.iterrows():
        ax.scatter(row["import_leakage"] * 100, row["multiplier"],
                   s=30, alpha=0.25, c="#9E9E9E", marker="x", zorder=2)

    # Arrow from raw to effective for Japan
    jpn = df[df["iso3"] == "JPN"].iloc[0]
    ax.annotate("", xy=(jpn["import_leakage"] * 100, jpn["effective_multiplier"]),
                xytext=(jpn["import_leakage"] * 100, jpn["multiplier"]),
                arrowprops=dict(arrowstyle="-|>", color="#F44336", lw=1.5,
                                linestyle="--"))
    leak_txt = "輸入漏出\n効果" if lang == "ja" else "Import\nleakage\neffect"
    ax.text(jpn["import_leakage"] * 100 + 0.3,
            (jpn["multiplier"] + jpn["effective_multiplier"]) / 2,
            leak_txt, fontsize=7, color="#F44336", va="center")

    # Reference line: fiscal sustainability threshold
    thr_lbl = ("財政閾値（日本: pf/τ）" if lang == "ja"
               else "Fiscal threshold (Japan: pf/tau)")
    ax.axhline(0.84 / 0.33, color="#4CAF50", linestyle=":", alpha=0.5,
               label=thr_lbl)

    if lang == "ja":
        ax.set_xlabel("医療関連輸入漏出率（%CHE）")
        ax.set_ylabel("乗数（実効 = 漏出調整後）")
        ax.set_title("I-O乗数 vs 医療関連輸入漏出率")
    else:
        ax.set_xlabel("Medical Import Leakage (% of CHE)")
        ax.set_ylabel("Multiplier (effective = adjusted for leakage)")
        ax.set_title("I-O Multiplier vs Medical Import Leakage")
    ax.legend(fontsize=8, loc="upper right")
    plt.tight_layout()
    suffix = "_ja" if lang == "ja" else ""
    path = os.path.join(FIG, f"fig7_import_leakage_multiplier{suffix}.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")
    return path


def fig8_counterfactual_japan(cf_results, lang="en"):
    """Bar chart: Japan counterfactual fiscal return ratios."""
    fig, ax = plt.subplots(figsize=(8, 5))

    if lang == "ja":
        scenarios = [
            ("ベースライン\n（実績）", cf_results["baseline"]["fiscal_return"], "#F44336"),
            ("A: OECD平均\n装置密度", cf_results["scenario_a"]["fiscal_return"], "#FF9800"),
            ("B: 国内\n製造", cf_results["scenario_b"]["fiscal_return"], "#1976D2"),
            ("C: A+B\n（両方）", cf_results["scenario_c"]["fiscal_return"], "#4CAF50"),
        ]
    else:
        scenarios = [
            ("Baseline\n(actual)", cf_results["baseline"]["fiscal_return"], "#F44336"),
            ("A: OECD-avg\nequipment", cf_results["scenario_a"]["fiscal_return"], "#FF9800"),
            ("B: Domestic\nmanufacturing", cf_results["scenario_b"]["fiscal_return"], "#1976D2"),
            ("C: Both\n(A + B)", cf_results["scenario_c"]["fiscal_return"], "#4CAF50"),
        ]

    labels = [s[0] for s in scenarios]
    vals = [s[1] for s in scenarios]
    colors = [s[2] for s in scenarios]

    bars = ax.bar(labels, vals, color=colors, edgecolor="white", width=0.6)
    thr = "持続可能性閾値 (1.0)" if lang == "ja" else "Sustainability threshold (1.0)"
    ax.axhline(1.0, color="black", linestyle="--", linewidth=1, alpha=0.5, label=thr)

    for bar, val in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width() / 2, val + 0.02,
                f"{val:.2f}", ha="center", va="bottom", fontsize=10, fontweight="bold")

    if lang == "ja":
        ax.set_ylabel("財政回収率（需要側）")
        ax.set_title("日本：カウンターファクチュアル持続可能性シナリオ")
    else:
        ax.set_ylabel("Fiscal Return Ratio (demand-side)")
        ax.set_title("Japan: Counterfactual Sustainability Scenarios")
    ax.legend(fontsize=8)
    ax.set_ylim(0, max(vals) * 1.25)
    plt.tight_layout()
    suffix = "_ja" if lang == "ja" else ""
    path = os.path.join(FIG, f"fig8_counterfactual_japan{suffix}.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved: {path}")
    return path


# ---------------------------------------------------------------------------
# 9. Main
# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print("Healthcare Economic Effect -- Neutral Sustainability Analysis")
    print("=" * 60)

    # I-O multiplier comparison
    print("\n[1] I-O Multiplier data")
    IO_MULTIPLIERS.to_csv(os.path.join(DATA, "io_multipliers.csv"), index=False)
    fig1_io_multiplier_comparison(IO_MULTIPLIERS)

    # HLGH evidence
    print("\n[2] Health-Led Growth Hypothesis evidence")
    HLGH_EVIDENCE.to_csv(os.path.join(DATA, "hlgh_evidence.csv"), index=False)

    # Cross-country scatter
    print("\n[3] Cross-country CHE vs Life Expectancy")
    cc_df = build_cross_country_df()
    if len(cc_df) > 5:
        _, overfit = fig2_scatter_che_vs_le(cc_df)
        with open(os.path.join(DATA, "preston_overfit_test.json"), "w") as f:
            json.dump(overfit, f, indent=2)
        print(f"  Preston overfit test: with US p={overfit['p_incl_us']:.4f}, "
              f"without US p={overfit['p_excl_us']:.4f}")
    else:
        print("  [WARN] Insufficient WB data; skipping scatter plot.")

    # Fiscal sustainability
    print("\n[4] Neutral fiscal sustainability")
    sust_df = sustainability_table()
    fig3_fiscal_sustainability(sust_df)
    print(sust_df.to_string(index=False))

    # Dual-return schematic
    print("\n[5] Dual-return conceptual diagram")
    fig4_dual_return_schematic()

    # Three-layer analogy (EN + JA)
    print("\n[6] Three-layer tempo analogy")
    fig5_three_layer_analogy(lang="en")
    fig5_three_layer_analogy(lang="ja")

    # PoC A-H results
    print("\n[7] PoC A-H results (from healthcare_tempo_poc)")
    with open(os.path.join(DATA, "poc_AH_summary.json"), "w") as f:
        json.dump(POC_AH_RESULTS, f, indent=2)
    THREE_LAYER_ANALOGY.to_csv(
        os.path.join(DATA, "three_layer_analogy.csv"), index=False)
    m2 = POC_AH_RESULTS["models"]["M2_tempo_lag"]
    print(f"  M2 level RMSE: {m2['level_rmse_median']:.3f} yr")
    print(f"  mu_H1 drift: +{m2['mu_H1_median_yr_per_yr']:.2f} yr/yr")
    print(f"  M2 beats M1: {POC_AH_RESULTS['key_findings']['M2_beats_M1_pct']}%")

    # Tempo narrative
    print("\n[8] Tempo-adjusted narrative")
    narrative = tempo_adjusted_narrative()
    with open(os.path.join(DATA, "tempo_narrative.json"), "w") as f:
        json.dump(narrative, f, indent=2, default=str)
    for k, v in narrative.items():
        if isinstance(v, str):
            print(f"  {k}: {v[:80]}...")

    # Medical equipment stock & import leakage
    print("\n[9] Medical equipment stock & import leakage analysis")
    eq_trade_df = build_equipment_trade_df()
    eq_trade_df.to_csv(os.path.join(DATA, "equipment_trade_analysis.csv"), index=False)
    fig6_equipment_density_comparison(eq_trade_df, lang="en")
    fig7_import_leakage_vs_multiplier(eq_trade_df, lang="en")

    print("\n[10] Japan counterfactual scenarios")
    cf = japan_counterfactual(eq_trade_df)
    with open(os.path.join(DATA, "japan_counterfactual.json"), "w") as f:
        json.dump(cf, f, indent=2, default=str)
    fig8_counterfactual_japan(cf, lang="en")

    # JA versions with CJK font
    with _ja_font_ctx():
        fig6_equipment_density_comparison(eq_trade_df, lang="ja")
        fig7_import_leakage_vs_multiplier(eq_trade_df, lang="ja")
        fig8_counterfactual_japan(cf, lang="ja")
    print(f"  Baseline fiscal return: {cf['baseline']['fiscal_return']:.2f}")
    print(f"  Scenario A (avg equip): {cf['scenario_a']['fiscal_return']:.2f}")
    print(f"  Scenario B (domestic mfg): {cf['scenario_b']['fiscal_return']:.2f}")
    print(f"  Scenario C (both): {cf['scenario_c']['fiscal_return']:.2f}")

    # Print summary table
    print("\n  Equipment & Trade Summary:")
    print(eq_trade_df[["country", "equip_density", "import_leakage",
                        "multiplier", "effective_multiplier"]].to_string(index=False))

    print("\n" + "=" * 60)
    print("Analysis complete. See output/figures/ and data/.")
    print("=" * 60)


if __name__ == "__main__":
    main()

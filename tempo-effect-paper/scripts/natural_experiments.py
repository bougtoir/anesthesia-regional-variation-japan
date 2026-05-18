#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Natural Experiments Analysis: Political/Border Changes as Exogenous Shocks
to the Endogenous Renewal Model

This script analyses countries where major political events caused large-scale
population redistribution equivalent to massive migration shocks:

1. Germany (1990): Reunification — synthetic East+West vs observed unified trajectory
2. Czechoslovakia (1993): Velvet Divorce — synthetic sum vs observed pre-split trajectory
3. Yugoslavia (1991–2001): Breakup — synthetic sum of successors vs observed pre-breakup
4. Baltic states (1991): USSR dissolution — Estonia, Latvia, Lithuania independence
5. Ethiopia/Eritrea (1993): Eritrean independence

For Germany (the primary case study), we model East and West Germany separately
using their distinct demographic parameters, then compare the synthetic
combination against the observed unified Germany trajectory after 1990.

Data sources:
- World Bank Open Data (SP.POP.TOTL, SP.DYN.TFRT.IN, SP.DYN.LE00.IN)
- East/West Germany demographic parameters from:
  * Statistisches Bundesamt (Federal Statistical Office of Germany)
  * Human Mortality Database (HMD)
  * Goldstein & Klüsener (2014), "Spatial analysis of the causes of the
    fertility decline in the former GDR"
  * Witte & Wagner (1995), "Declining fertility in East Germany after
    unification"
"""

import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar
from scipy.stats import norm
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
import warnings
warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SCRIPT_DIR)
FIG_DIR = os.path.join(BASE_DIR, 'figures')
DATA_DIR = '/home/ubuntu/wpp_data'
os.makedirs(FIG_DIR, exist_ok=True)

# ============================================================
# 1. MODEL FUNCTIONS (from model_fit_v2.py)
# ============================================================

def gompertz_survival(x, a_g, b_g):
    x = np.asarray(x, dtype=float)
    return np.exp(-(a_g / b_g) * np.expm1(b_g * x))

def gompertz_le(a_g, b_g, max_age=110):
    ages = np.arange(0, max_age + 1, dtype=float)
    surv = gompertz_survival(ages, a_g, b_g)
    return np.trapezoid(surv, ages) if hasattr(np, 'trapezoid') else np.trapz(surv, ages)

def calibrate_gompertz(target_le, b_g=0.085):
    def obj(log_a):
        a = np.exp(log_a)
        return (gompertz_le(a, b_g) - target_le) ** 2
    res = minimize_scalar(obj, bounds=(-14, -1), method='bounded')
    return np.exp(res.x), b_g

def make_asfr(mac, sigma, tfr, max_age=100):
    asfr = np.zeros(max_age + 1)
    fertile = np.arange(15, 50)
    pdf = norm.pdf(fertile, loc=mac, scale=sigma)
    total = pdf.sum()
    if total > 0:
        asfr[15:50] = pdf * (tfr / total)
    return asfr

def run_model_dynamic(init_pop_total, params_by_decade, start_year, end_year,
                      female_ratio=0.4886, max_age=100):
    """
    Run renewal model with parameters updated every decade.

    params_by_decade: dict of {decade_start: {'tfr': x, 'le': x, 'mac': x, 'sigma': x}}
    init_pop_total: total population at start_year
    """
    n_years = end_year - start_year
    # Build initial age distribution (approximate uniform-ish with survival decay)
    a_g, b_g = calibrate_gompertz(list(params_by_decade.values())[0]['le'])
    ages = np.arange(0, max_age + 1, dtype=float)
    surv = gompertz_survival(ages, a_g, b_g)
    pop0 = surv / surv.sum() * init_pop_total

    pop = np.zeros((n_years + 1, max_age + 1))
    pop[0] = pop0
    total_pop = np.zeros(n_years + 1)
    total_pop[0] = pop0.sum()

    sorted_decades = sorted(params_by_decade.keys())

    for t in range(1, n_years + 1):
        current_year = start_year + t
        # Find which decade's parameters to use
        decade = sorted_decades[0]
        for d in sorted_decades:
            if current_year > d:
                decade = d
        params = params_by_decade[decade]

        a_g, b_g = calibrate_gompertz(params['le'])
        surv_arr = gompertz_survival(np.arange(0, max_age + 2, dtype=float), a_g, b_g)
        asfr = make_asfr(params['mac'], params.get('sigma', 5.0), params['tfr'], max_age)

        # Survival ratios
        sr = np.zeros(max_age + 1)
        for a in range(1, max_age + 1):
            if surv_arr[a-1] > 1e-15:
                sr[a] = min(1.0, max(0.0, surv_arr[a] / surv_arr[a-1]))
        inf_surv = min(1.0, surv_arr[1] / surv_arr[0]) if surv_arr[0] > 1e-15 else 0.95

        # Births
        births = np.sum(pop[t-1, 15:50] * female_ratio * asfr[15:50])
        # Age with survival
        pop[t, 1:] = pop[t-1, :-1] * sr[1:]
        pop[t, 0] = max(0, births * inf_surv)
        total_pop[t] = pop[t].sum()

    years = np.arange(start_year, end_year + 1)
    return years, total_pop


# ============================================================
# 2. LOAD WORLD BANK DATA
# ============================================================

def load_wb_data():
    """Load World Bank demographic data."""
    pop_df = pd.read_csv(os.path.join(DATA_DIR, 'wb_total_population.csv'), index_col=[0, 1])
    tfr_df = pd.read_csv(os.path.join(DATA_DIR, 'wb_tfr.csv'), index_col=[0, 1])
    le_df = pd.read_csv(os.path.join(DATA_DIR, 'wb_life_expectancy.csv'), index_col=[0, 1])
    return pop_df, tfr_df, le_df

def get_country_series(df, code):
    """Extract time series for a country from World Bank DataFrame."""
    try:
        row = df.loc[code]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        years = []
        vals = []
        for col in row.index:
            if col.startswith('YR'):
                yr = int(col[2:])
                v = row[col]
                if pd.notna(v):
                    years.append(yr)
                    vals.append(float(v))
        return np.array(years), np.array(vals)
    except KeyError:
        return np.array([]), np.array([])

def get_value_at_year(df, code, year):
    """Get a single value for a country at a specific year."""
    col = f'YR{year}'
    try:
        row = df.loc[code]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        v = row[col]
        return float(v) if pd.notna(v) else None
    except (KeyError, IndexError):
        return None

def get_decade_params(tfr_df, le_df, code, decades, default_mac=28.0, default_sigma=5.0,
                      mac_overrides=None):
    """Build decade-parameter dict for a country."""
    params = {}
    for d in decades:
        tfr = get_value_at_year(tfr_df, code, d)
        le = get_value_at_year(le_df, code, d)
        if tfr is None or le is None:
            # Try adjacent years
            for offset in [1, -1, 2, -2, 5, -5]:
                if tfr is None:
                    tfr = get_value_at_year(tfr_df, code, d + offset)
                if le is None:
                    le = get_value_at_year(le_df, code, d + offset)
        if tfr is None:
            tfr = 2.0
        if le is None:
            le = 70.0
        mac = default_mac
        if mac_overrides and d in mac_overrides:
            mac = mac_overrides[d]
        sigma = default_sigma
        if tfr > 4.0:
            sigma = 6.5
        elif tfr > 2.5:
            sigma = 5.5
        elif tfr < 1.5:
            sigma = 4.5
        params[d] = {'tfr': tfr, 'le': le, 'mac': mac, 'sigma': sigma}
    return params


# ============================================================
# 3. EAST/WEST GERMANY HISTORICAL DATA
# ============================================================
#
# Sources:
#   Population: Statistisches Bundesamt, HMD
#   TFR: Goldstein & Klüsener 2014; Sobotka 2011; Witte & Wagner 1995
#   Life expectancy: HMD (mortality.org), Kibele 2012
#   MAC: Eurostat, Sobotka & Lutz 2011
#
# Note: World Bank "Germany" (DEU) = unified territory for all years
# (retrospective estimates for pre-1990). We need separate E/W data.

EAST_GERMANY = {
    # Population (thousands -> individuals)
    'pop': {
        1970: 17_068_000, 1980: 16_740_000, 1990: 16_028_000,
    },
    # Decade parameters
    'params': {
        1970: {'tfr': 2.13, 'le': 70.0, 'mac': 23.8, 'sigma': 4.5},
        1980: {'tfr': 1.94, 'le': 72.8, 'mac': 24.5, 'sigma': 4.5},
        1990: {'tfr': 1.52, 'le': 74.5, 'mac': 25.1, 'sigma': 4.5},
        2000: {'tfr': 1.21, 'le': 77.3, 'mac': 28.0, 'sigma': 4.5},
        2010: {'tfr': 1.43, 'le': 79.5, 'mac': 29.5, 'sigma': 4.5},
        2020: {'tfr': 1.53, 'le': 80.5, 'mac': 30.2, 'sigma': 4.5},
    },
}

WEST_GERMANY = {
    'pop': {
        1970: 60_651_000, 1980: 61_566_000, 1990: 63_254_000,
    },
    'params': {
        1970: {'tfr': 2.03, 'le': 71.0, 'mac': 26.6, 'sigma': 5.0},
        1980: {'tfr': 1.44, 'le': 73.8, 'mac': 27.2, 'sigma': 5.0},
        1990: {'tfr': 1.45, 'le': 76.0, 'mac': 28.3, 'sigma': 5.0},
        2000: {'tfr': 1.41, 'le': 78.7, 'mac': 29.5, 'sigma': 5.0},
        2010: {'tfr': 1.39, 'le': 80.5, 'mac': 30.5, 'sigma': 5.0},
        2020: {'tfr': 1.53, 'le': 81.2, 'mac': 31.2, 'sigma': 5.0},
    },
}

# ============================================================
# 4. ANALYSIS FUNCTIONS
# ============================================================

def analyse_germany_reunification(pop_df, tfr_df, le_df):
    """
    Germany reunification (1990) as natural experiment.
    
    Compare:
    A) Synthetic: model East + West separately from 1970, sum their populations
    B) Observed: actual unified Germany trajectory from World Bank
    C) Model: unified Germany modeled from 1970 with unified parameters
    
    The divergence between synthetic (A) and observed (B) quantifies the
    migration/integration shock of reunification.
    """
    print("\n" + "="*70)
    print("GERMANY REUNIFICATION ANALYSIS")
    print("="*70)

    start_year = 1970
    end_year = 2023

    # Run East Germany model
    east_years, east_pop = run_model_dynamic(
        EAST_GERMANY['pop'][1970],
        EAST_GERMANY['params'],
        start_year, end_year
    )
    print(f"East Germany model: {east_pop[0]/1e6:.1f}M → {east_pop[-1]/1e6:.1f}M")

    # Run West Germany model
    west_years, west_pop = run_model_dynamic(
        WEST_GERMANY['pop'][1970],
        WEST_GERMANY['params'],
        start_year, end_year
    )
    print(f"West Germany model: {west_pop[0]/1e6:.1f}M → {west_pop[-1]/1e6:.1f}M")

    # Synthetic combined
    synthetic_pop = east_pop + west_pop
    print(f"Synthetic combined: {synthetic_pop[0]/1e6:.1f}M → {synthetic_pop[-1]/1e6:.1f}M")

    # Actual Germany trajectory
    actual_years, actual_pop = get_country_series(pop_df, 'DEU')
    mask = (actual_years >= start_year) & (actual_years <= end_year)
    actual_years = actual_years[mask]
    actual_pop = actual_pop[mask]
    print(f"Observed Germany: {actual_pop[0]/1e6:.1f}M → {actual_pop[-1]/1e6:.1f}M")

    # Unified Germany model (single entity from 1970)
    unified_init = EAST_GERMANY['pop'][1970] + WEST_GERMANY['pop'][1970]
    # Weighted average params
    unified_params = {}
    e_frac = EAST_GERMANY['pop'][1970] / unified_init
    w_frac = 1 - e_frac
    for decade in [1970, 1980, 1990, 2000, 2010, 2020]:
        ep = EAST_GERMANY['params'][decade]
        wp = WEST_GERMANY['params'][decade]
        unified_params[decade] = {
            'tfr': ep['tfr'] * e_frac + wp['tfr'] * w_frac,
            'le': ep['le'] * e_frac + wp['le'] * w_frac,
            'mac': ep['mac'] * e_frac + wp['mac'] * w_frac,
            'sigma': ep['sigma'] * e_frac + wp['sigma'] * w_frac,
        }
    unified_years, unified_pop = run_model_dynamic(
        unified_init, unified_params, start_year, end_year
    )
    print(f"Unified model: {unified_pop[0]/1e6:.1f}M → {unified_pop[-1]/1e6:.1f}M")

    # Calculate MAPE
    common_years = np.intersect1d(east_years, actual_years)
    synth_at_common = []
    actual_at_common = []
    unified_at_common = []
    for y in common_years:
        si = y - start_year
        ai = np.where(actual_years == y)[0][0]
        synth_at_common.append(synthetic_pop[si])
        actual_at_common.append(actual_pop[ai])
        unified_at_common.append(unified_pop[si])
    synth_at_common = np.array(synth_at_common)
    actual_at_common = np.array(actual_at_common)
    unified_at_common = np.array(unified_at_common)

    mape_synth = np.mean(np.abs((synth_at_common - actual_at_common) / actual_at_common)) * 100
    mape_unified = np.mean(np.abs((unified_at_common - actual_at_common) / actual_at_common)) * 100
    print(f"\nMAPE (synthetic E+W vs observed): {mape_synth:.1f}%")
    print(f"MAPE (unified model vs observed): {mape_unified:.1f}%")

    # Reunification shock: divergence at 1990
    idx_1990 = 1990 - start_year
    synth_1990 = synthetic_pop[idx_1990]
    ai_1990 = actual_pop[np.where(actual_years == 1990)[0][0]] if 1990 in actual_years else None
    if ai_1990:
        print(f"\nAt reunification (1990):")
        print(f"  Synthetic E+W: {synth_1990/1e6:.1f}M")
        print(f"  Observed:      {ai_1990/1e6:.1f}M")
        print(f"  Difference:    {(synth_1990 - ai_1990)/1e6:.1f}M ({(synth_1990/ai_1990 - 1)*100:.1f}%)")

    # Post-reunification divergence
    for yr in [1995, 2000, 2005, 2010, 2015, 2020, 2023]:
        idx = yr - start_year
        if yr in actual_years:
            ai = actual_pop[np.where(actual_years == yr)[0][0]]
            sp = synthetic_pop[idx]
            print(f"  {yr}: Synth={sp/1e6:.1f}M, Obs={ai/1e6:.1f}M, Δ={(sp-ai)/1e6:.1f}M ({(sp/ai-1)*100:.1f}%)")

    # Return data for plotting
    return {
        'east_years': east_years, 'east_pop': east_pop,
        'west_years': west_years, 'west_pop': west_pop,
        'synthetic_years': east_years, 'synthetic_pop': synthetic_pop,
        'actual_years': actual_years, 'actual_pop': actual_pop,
        'unified_years': unified_years, 'unified_pop': unified_pop,
        'mape_synth': mape_synth, 'mape_unified': mape_unified,
    }


def analyse_country_split(pop_df, tfr_df, le_df, name, predecessor_code,
                          successor_codes, split_year, start_year=1970, end_year=2023):
    """
    Analyse a country that split: compare pre-split observed trajectory
    with synthetic sum of successor states modeled independently.
    """
    print(f"\n{'='*70}")
    print(f"{name} (split: {split_year})")
    print("="*70)

    # Get actual data for predecessor and each successor
    pred_years, pred_pop = get_country_series(pop_df, predecessor_code)
    if len(pred_years) == 0:
        # Some predecessors don't have WB data — use sum of successors
        print(f"  No data for predecessor {predecessor_code}, using sum of successors")

    results = {'name': name, 'split_year': split_year, 'successors': {}}

    # Model each successor independently
    total_synth = None
    for code, sname in successor_codes.items():
        s_years, s_pop = get_country_series(pop_df, code)
        if len(s_years) == 0:
            print(f"  No data for {sname} ({code})")
            continue

        # Get initial population at start_year or earliest available
        mask = s_years >= start_year
        if not mask.any():
            continue
        s_years = s_years[mask]
        s_pop = s_pop[mask]

        init_year = s_years[0]
        init_pop = s_pop[0]

        # Get parameters
        decades = list(range(max(init_year, 1970), end_year, 10))
        if not decades:
            continue
        params = get_decade_params(tfr_df, le_df, code, decades)

        # Run model
        try:
            m_years, m_pop = run_model_dynamic(init_pop, params, init_year, end_year)
        except Exception as e:
            print(f"  Error modeling {sname}: {e}")
            continue

        # Calculate MAPE vs actual
        common = np.intersect1d(m_years, s_years)
        if len(common) > 5:
            mv = np.array([m_pop[y - init_year] for y in common])
            av = np.array([s_pop[np.where(s_years == y)[0][0]] for y in common])
            mape = np.mean(np.abs((mv - av) / av)) * 100
        else:
            mape = np.nan

        results['successors'][code] = {
            'name': sname, 'model_years': m_years, 'model_pop': m_pop,
            'actual_years': s_years, 'actual_pop': s_pop, 'mape': mape,
            'init_year': init_year,
        }
        print(f"  {sname}: {init_pop/1e6:.1f}M → model {m_pop[-1]/1e6:.1f}M, "
              f"actual {s_pop[-1]/1e6:.1f}M, MAPE={mape:.1f}%")

        # Accumulate synthetic total (aligned to common timeline)
        # Use actual data where available, model where not
        full_years = np.arange(start_year, end_year + 1)
        full_pop = np.zeros(len(full_years))
        for i, yr in enumerate(full_years):
            if yr in s_years:
                full_pop[i] = s_pop[np.where(s_years == yr)[0][0]]
            elif yr >= init_year and (yr - init_year) < len(m_pop):
                full_pop[i] = m_pop[yr - init_year]
        if total_synth is None:
            total_synth = full_pop.copy()
        else:
            total_synth += full_pop

    # Compare synthetic total vs predecessor observed
    if total_synth is not None and len(pred_years) > 0:
        full_years = np.arange(start_year, end_year + 1)
        common = np.intersect1d(full_years, pred_years)
        pre_split = common[common < split_year]
        if len(pre_split) > 5:
            sv = np.array([total_synth[y - start_year] for y in pre_split])
            pv = np.array([pred_pop[np.where(pred_years == y)[0][0]] for y in pre_split])
            valid = pv > 0
            if valid.any():
                mape_pre = np.mean(np.abs((sv[valid] - pv[valid]) / pv[valid])) * 100
                print(f"\n  Pre-split synthetic sum vs predecessor MAPE: {mape_pre:.1f}%")
                results['mape_pre_split'] = mape_pre

    results['synthetic_years'] = np.arange(start_year, end_year + 1)
    results['synthetic_total'] = total_synth
    results['pred_years'] = pred_years
    results['pred_pop'] = pred_pop

    return results


# ============================================================
# 5. PLOTTING
# ============================================================

def plot_germany(data, save_path):
    """Plot Germany reunification analysis."""
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))

    # Panel A: East and West separately + synthetic combined
    ax = axes[0]
    ax.plot(data['west_years'], data['west_pop']/1e6, 'b-', linewidth=1.5,
            label='West Germany (model)', alpha=0.8)
    ax.plot(data['east_years'], data['east_pop']/1e6, 'r-', linewidth=1.5,
            label='East Germany (model)', alpha=0.8)
    ax.plot(data['synthetic_years'], data['synthetic_pop']/1e6, 'g--', linewidth=2,
            label='Synthetic E+W sum')
    ax.plot(data['actual_years'], data['actual_pop']/1e6, 'k-', linewidth=2,
            label='Observed Germany')
    ax.axvline(x=1990, color='gray', linestyle=':', alpha=0.7, label='Reunification (1990)')
    ax.set_xlabel('Year', fontsize=12)
    ax.set_ylabel('Population (millions)', fontsize=12)
    ax.set_title('A. Germany: Synthetic East+West vs Observed', fontsize=13, fontweight='bold')
    ax.legend(fontsize=9, loc='upper left')
    ax.set_xlim(1970, 2023)
    ax.set_ylim(bottom=0)
    ax.grid(True, alpha=0.3)

    # Panel B: Percentage deviation from observed
    ax = axes[1]
    common_years = np.intersect1d(data['synthetic_years'], data['actual_years'])
    synth_dev = []
    unified_dev = []
    for y in common_years:
        si = y - data['synthetic_years'][0]
        ai = np.where(data['actual_years'] == y)[0][0]
        ui = y - data['unified_years'][0]
        obs = data['actual_pop'][ai]
        synth_dev.append((data['synthetic_pop'][si] / obs - 1) * 100)
        unified_dev.append((data['unified_pop'][ui] / obs - 1) * 100)

    ax.plot(common_years, synth_dev, 'g-', linewidth=2,
            label=f'Synthetic E+W (MAPE={data["mape_synth"]:.1f}%)')
    ax.plot(common_years, unified_dev, 'm--', linewidth=2,
            label=f'Unified model (MAPE={data["mape_unified"]:.1f}%)')
    ax.axhline(y=0, color='k', linewidth=0.5)
    ax.axvline(x=1990, color='gray', linestyle=':', alpha=0.7)
    ax.fill_between(common_years, synth_dev, 0, alpha=0.15, color='green')
    ax.set_xlabel('Year', fontsize=12)
    ax.set_ylabel('Deviation from observed (%)', fontsize=12)
    ax.set_title('B. Model deviation from observed trajectory', fontsize=13, fontweight='bold')
    ax.legend(fontsize=9, loc='best')
    ax.set_xlim(1970, 2023)
    # Keep auto y-range but ensure 0 is included as a tick
    ymin, ymax = ax.get_ylim()
    if ymax < 0:
        ax.set_ylim(top=0)
    import matplotlib.ticker as mticker
    ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins='auto', steps=[1, 2, 5, 10]))
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"\nSaved: {save_path}")


def plot_natural_experiments_summary(germany_data, split_results, save_path):
    """Summary figure of all natural experiments."""
    n_cases = 1 + len(split_results)  # Germany + splits
    fig, axes = plt.subplots(2, 3, figsize=(18, 11))
    axes = axes.flatten()

    # Panel 0: Germany
    ax = axes[0]
    ax.plot(germany_data['actual_years'], germany_data['actual_pop']/1e6, 'k-',
            linewidth=2, label='Observed')
    ax.plot(germany_data['synthetic_years'], germany_data['synthetic_pop']/1e6,
            'g--', linewidth=2, label='Synthetic E+W')
    ax.axvline(x=1990, color='red', linestyle=':', alpha=0.7)
    ax.set_title('Germany (reunification 1990)', fontsize=11, fontweight='bold')
    ax.set_ylabel('Population (millions)', fontsize=10)
    ax.set_ylim(bottom=0)
    ax.legend(fontsize=8)
    ax.grid(True, alpha=0.3)

    # Remaining panels: split cases
    for i, (key, res) in enumerate(split_results.items()):
        if i + 1 >= len(axes):
            break
        ax = axes[i + 1]

        # Plot predecessor if available
        if len(res.get('pred_years', [])) > 0:
            mask = res['pred_years'] <= res['split_year'] + 5
            ax.plot(res['pred_years'][mask], res['pred_pop'][mask]/1e6, 'k-',
                    linewidth=2, label='Pre-split observed')

        # Plot synthetic sum
        if res.get('synthetic_total') is not None:
            synth_years = res['synthetic_years']
            synth_pop = res['synthetic_total']
            valid = synth_pop > 0
            ax.plot(synth_years[valid], synth_pop[valid]/1e6, 'g--',
                    linewidth=2, label='Sum of successors')

        # Plot individual successors
        colors = ['blue', 'red', 'orange', 'purple', 'brown', 'cyan']
        for j, (code, sdata) in enumerate(res.get('successors', {}).items()):
            c = colors[j % len(colors)]
            ax.plot(sdata['actual_years'], sdata['actual_pop']/1e6,
                    f'-', color=c, linewidth=1, alpha=0.6,
                    label=f"{sdata['name']} (obs)")

        ax.axvline(x=res['split_year'], color='red', linestyle=':', alpha=0.7)
        ax.set_title(f"{res['name']} ({res['split_year']})", fontsize=11, fontweight='bold')
        ax.set_ylabel('Population (millions)', fontsize=10)
        ax.set_ylim(bottom=0)
        ax.legend(fontsize=7, loc='best')
        ax.grid(True, alpha=0.3)

    # Hide unused axes
    for i in range(n_cases, len(axes)):
        axes[i].set_visible(False)

    for ax in axes[:n_cases]:
        ax.set_xlabel('Year', fontsize=10)

    fig.suptitle('Natural Experiments: Political/Border Changes as Exogenous Shocks\n'
                  'to the Endogenous Renewal Model',
                  fontsize=14, fontweight='bold')
    fig.tight_layout(rect=[0, 0, 1, 0.93])
    fig.savefig(save_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"\nSaved: {save_path}")


# ============================================================
# 6. MAIN EXECUTION
# ============================================================

if __name__ == '__main__':
    print("Loading World Bank data...")
    pop_df, tfr_df, le_df = load_wb_data()
    print("Data loaded.\n")

    # --- A. Germany Reunification ---
    germany_data = analyse_germany_reunification(pop_df, tfr_df, le_df)
    plot_germany(germany_data, os.path.join(FIG_DIR, 'fig_germany_reunification.png'))

    # --- B. Czechoslovakia (1993 Velvet Divorce) ---
    # Note: WB 'CZE' is Czechia only (~10M), not Czechoslovakia (~15M).
    # Using CZE as predecessor while CZE+SVK are successors creates a
    # scale mismatch (sum ~15M vs predecessor ~10M). No WB code exists
    # for Czechoslovakia (CSK), so we skip the pre-split comparison.
    czechoslovakia = analyse_country_split(
        pop_df, tfr_df, le_df,
        name="Czechoslovakia → Czechia + Slovakia",
        predecessor_code="CSK",  # no WB entry for Czechoslovakia → skips pre-split comparison
        successor_codes={'CZE': 'Czechia', 'SVK': 'Slovakia'},
        split_year=1993,
    )

    # --- C. Yugoslavia (1991-2001 breakup) ---
    yugoslavia = analyse_country_split(
        pop_df, tfr_df, le_df,
        name="Yugoslavia breakup",
        predecessor_code="YUG",  # may not exist in WB
        successor_codes={
            'HRV': 'Croatia', 'SVN': 'Slovenia', 'BIH': 'Bosnia & Herz.',
            'SRB': 'Serbia', 'MKD': 'N. Macedonia', 'MNE': 'Montenegro',
        },
        split_year=1991,
    )

    # --- D. Baltic States (USSR dissolution 1991) ---
    # Note: predecessor_code uses a non-existent code to skip the pre-split
    # comparison, because WB 'RUS' is Russia only (~130M), not the USSR.
    # Comparing ~7M (Baltic total) vs ~130M (Russia) is meaningless.
    # Individual successor MAPEs for each Baltic state are still correct.
    baltics = analyse_country_split(
        pop_df, tfr_df, le_df,
        name="Baltic states (USSR dissolution)",
        predecessor_code="SUN",  # no WB entry for USSR → skips pre-split comparison
        successor_codes={'EST': 'Estonia', 'LVA': 'Latvia', 'LTU': 'Lithuania'},
        split_year=1991,
    )

    # --- E. Ethiopia/Eritrea (1993) ---
    # Note: predecessor_code uses a non-existent code to skip the pre-split
    # comparison, because WB 'ETH' data uses current borders (excluding Eritrea)
    # for all years.  Using ETH as both predecessor and successor would create
    # an invalid self-referential comparison.  Individual successor MAPEs for
    # Ethiopia and Eritrea are still computed correctly against their own data.
    ethiopia_eritrea = analyse_country_split(
        pop_df, tfr_df, le_df,
        name="Ethiopia/Eritrea separation",
        predecessor_code="ETH_ERI_COMBINED",  # no WB entry → skips pre-split comparison
        successor_codes={'ETH': 'Ethiopia', 'ERI': 'Eritrea'},
        split_year=1993,
    )

    # --- Summary Figure ---
    split_results = {
        'czechoslovakia': czechoslovakia,
        'yugoslavia': yugoslavia,
        'baltics': baltics,
        'ethiopia_eritrea': ethiopia_eritrea,
    }
    plot_natural_experiments_summary(
        germany_data, split_results,
        os.path.join(FIG_DIR, 'fig_natural_experiments_summary.png')
    )

    # --- Print summary table ---
    print("\n" + "="*70)
    print("SUMMARY OF NATURAL EXPERIMENTS")
    print("="*70)
    print(f"\n{'Case':<40} {'MAPE (%)':<12} {'Key finding'}")
    print("-"*80)
    print(f"{'Germany (synthetic E+W vs observed)':<40} {germany_data['mape_synth']:<12.1f} "
          f"Reunification = migration shock")
    print(f"{'Germany (unified model vs observed)':<40} {germany_data['mape_unified']:<12.1f} "
          f"Baseline model fit")

    for key, res in split_results.items():
        for code, sdata in res.get('successors', {}).items():
            if not np.isnan(sdata.get('mape', np.nan)):
                print(f"{'  ' + sdata['name']:<40} {sdata['mape']:<12.1f}")

    print("\nDone!")

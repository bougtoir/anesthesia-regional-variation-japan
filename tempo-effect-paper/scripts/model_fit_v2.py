"""
Optimized model fit evaluation: Endogenous Renewal Model + Gompertz Survival
vs UN WPP 2024 for OECD + China + DRC
"""
import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar
from scipy.stats import norm
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os, warnings
warnings.filterwarnings('ignore')

print("Loading demographic indicators...")
demo = pd.read_csv('/home/ubuntu/wpp_data/WPP2024_Demographic_Indicators_Medium.csv.gz', low_memory=False)
print(f"Demo shape: {demo.shape}")

# Country IDs
country_ids = {
    "Australia": 36, "Austria": 40, "Belgium": 56, "Canada": 124, "Chile": 152,
    "China": 156, "Colombia": 170, "Costa Rica": 188, "Czechia": 203,
    "DRC": 180, "Denmark": 208, "Estonia": 233, "Finland": 246, "France": 250,
    "Germany": 276, "Greece": 300, "Hungary": 348, "Iceland": 352,
    "Ireland": 372, "Israel": 376, "Italy": 380, "Japan": 392,
    "Republic of Korea": 410, "Latvia": 428, "Lithuania": 440, "Luxembourg": 442,
    "Mexico": 484, "Netherlands": 528, "New Zealand": 554, "Norway": 578,
    "Poland": 616, "Portugal": 620, "Slovakia": 703, "Slovenia": 705,
    "Spain": 724, "Sweden": 752, "Switzerland": 756, "Türkiye": 792,
    "United Kingdom": 826, "United States": 840
}
id_to_name = {v: k for k, v in country_ids.items()}

# Filter demo to our countries only
demo = demo[demo['LocID'].isin(country_ids.values())].copy()
demo['Country'] = demo['LocID'].map(id_to_name)
print(f"Filtered demo: {demo.shape}")

# Load pop by age5 - filter immediately to reduce memory
print("Loading pop by age5 (filtered)...")
chunks = []
for chunk in pd.read_csv('/home/ubuntu/wpp_data/WPP2024_PopulationByAge5GroupSex_Medium.csv.gz',
                          low_memory=False, chunksize=50000):
    filtered = chunk[chunk['LocID'].isin(country_ids.values())]
    if len(filtered) > 0:
        chunks.append(filtered)
pop_age5 = pd.concat(chunks, ignore_index=True)
print(f"Pop age5 shape: {pop_age5.shape}")
print(f"Columns: {list(pop_age5.columns[:15])}")

# ============================================================
# MODEL FUNCTIONS
# ============================================================

def gompertz_survival(x, a_g, b_g):
    x = np.asarray(x, dtype=float)
    return np.exp(-(a_g / b_g) * np.expm1(b_g * x))

def gompertz_le(a_g, b_g, max_age=110):
    ages = np.arange(0, max_age + 1, dtype=float)
    surv = gompertz_survival(ages, a_g, b_g)
    return np.trapezoid(surv, ages)

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

def run_model(init_pop_5yr, tfr, mac, sigma, target_le,
              n_years=50, female_ratio=0.4886, max_age=100):
    a_g, b_g = calibrate_gompertz(target_le)
    surv_arr = gompertz_survival(np.arange(0, max_age + 2, dtype=float), a_g, b_g)
    asfr = make_asfr(mac, sigma, tfr, max_age)
    
    # Build initial single-year pop from 5-year groups
    pop0 = np.zeros(max_age + 1)
    for grp, val in init_pop_5yr.items():
        grp_s = str(grp)
        if '-' in grp_s:
            parts = grp_s.split('-')
            a_start, a_end = int(parts[0]), int(parts[1])
            span = a_end - a_start + 1
            for a in range(a_start, min(a_end + 1, max_age + 1)):
                pop0[a] = val / span
        elif '+' in grp_s:
            a_start = int(grp_s.replace('+', ''))
            pop0[min(a_start, max_age)] += val
    
    # Pre-compute survival ratios
    sr = np.zeros(max_age + 1)
    for a in range(1, max_age + 1):
        if surv_arr[a-1] > 1e-15:
            sr[a] = min(1.0, max(0.0, surv_arr[a] / surv_arr[a-1]))
    
    inf_surv = min(1.0, surv_arr[1] / surv_arr[0]) if surv_arr[0] > 1e-15 else 0.95
    
    pop = np.zeros((n_years + 1, max_age + 1))
    pop[0] = pop0
    total_pop = np.zeros(n_years + 1)
    total_pop[0] = pop0.sum()
    
    for t in range(1, n_years + 1):
        # Births from female pop * ASFR
        births = np.sum(pop[t-1, 15:50] * female_ratio * asfr[15:50])
        
        # Age with survival (vectorized)
        pop[t, 1:] = pop[t-1, :-1] * sr[1:]
        pop[t, 0] = max(0, births * inf_surv)
        total_pop[t] = pop[t].sum()
    
    return total_pop

# ============================================================
# RUN FOR ALL COUNTRIES
# ============================================================

BASE_YEARS = [1970, 1980, 1990, 2000]
results = []

for cname, loc_id in sorted(country_ids.items()):
    print(f"\n--- {cname} ---")
    
    for base_year in BASE_YEARS:
        # Get params
        row = demo[(demo['LocID'] == loc_id) & (demo['Time'] == base_year)]
        if len(row) == 0:
            continue
        row = row.iloc[0]
        tfr = float(row['TFR'])
        le = float(row['LEx'])
        mac = float(row['MAC'])
        total_pop_actual = float(row['TPopulation1July']) * 1000
        fem_ratio = float(row['TPopulationFemale1July']) / float(row['TPopulation1July']) if row['TPopulation1July'] > 0 else 0.5
        
        # Get initial age dist
        sub = pop_age5[(pop_age5['LocID'] == loc_id) & (pop_age5['Time'] == base_year)]
        if len(sub) == 0:
            continue
        init_pop = {}
        for _, r in sub.iterrows():
            init_pop[str(r['AgeGrp'])] = float(r['PopTotal']) * 1000
        
        # Sigma
        sigma = 5.0
        if tfr > 4.0: sigma = 6.5
        elif tfr > 2.5: sigma = 5.5
        elif tfr < 1.5: sigma = 4.5
        
        end_year = min(base_year + 50, 2023)
        n_sim = end_year - base_year
        if n_sim < 10:
            continue
        
        try:
            model_pop = run_model(init_pop, tfr, mac, sigma, le, n_sim, fem_ratio)
        except Exception as e:
            print(f"  Error {base_year}: {e}")
            continue
        
        # Actual trajectory
        actual_sub = demo[(demo['LocID'] == loc_id) & 
                          (demo['Time'] >= base_year) & 
                          (demo['Time'] <= end_year)].sort_values('Time')
        actual_years = actual_sub['Time'].values
        actual_pop = actual_sub['TPopulation1July'].values * 1000
        
        model_years = np.arange(base_year, base_year + n_sim + 1)
        common = np.intersect1d(model_years, actual_years)
        if len(common) < 5:
            continue
        
        mv, av = [], []
        for y in common:
            mi = y - base_year
            ai = np.where(actual_years == y)[0]
            if mi < len(model_pop) and len(ai) > 0:
                mv.append(model_pop[mi])
                av.append(actual_pop[ai[0]])
        mv, av = np.array(mv), np.array(av)
        
        if len(mv) < 5 or av[0] == 0:
            continue
        
        # Metrics
        mi_idx = mv / mv[0]
        ai_idx = av / av[0]
        mape_idx = np.mean(np.abs((mi_idx - ai_idx) / ai_idx)) * 100
        mape_abs = np.mean(np.abs((mv - av) / av)) * 100
        final_ratio = mv[-1] / av[-1] if av[-1] > 0 else np.nan
        ss_res = np.sum((mi_idx - ai_idx) ** 2)
        ss_tot = np.sum((ai_idx - np.mean(ai_idx)) ** 2)
        r2 = 1 - ss_res / ss_tot if ss_tot > 0 else np.nan
        
        results.append({
            'country': cname, 'loc_id': loc_id, 'base_year': base_year,
            'horizon': n_sim, 'end_year': end_year,
            'tfr': tfr, 'le': le, 'mac': mac,
            'init_pop': av[0], 'final_model': mv[-1], 'final_actual': av[-1],
            'mape_index': mape_idx, 'mape_abs': mape_abs,
            'final_ratio': final_ratio, 'r2': r2, 'sigma': sigma,
        })
        
        print(f"  {base_year}→{end_year}: MAPE={mape_idx:.1f}%, R²={r2:.3f}, ratio={final_ratio:.3f}")

# Save results
rdf = pd.DataFrame(results)
rdf.to_csv('/home/ubuntu/model_fit_results.csv', index=False)
print(f"\n{'='*60}")
print(f"Saved {len(rdf)} results")

# Summary
print(f"\n{'='*60}")
print("SUMMARY BY BASE YEAR")
print(f"{'='*60}")
for by in BASE_YEARS:
    sub = rdf[rdf['base_year'] == by]
    if len(sub) == 0: continue
    print(f"\n{by} ({len(sub)} countries):")
    print(f"  MAPE(idx): mean={sub['mape_index'].mean():.1f}% med={sub['mape_index'].median():.1f}% max={sub['mape_index'].max():.1f}%")
    print(f"  R²: mean={sub['r2'].mean():.3f} med={sub['r2'].median():.3f}")
    print(f"  Final ratio: mean={sub['final_ratio'].mean():.3f} std={sub['final_ratio'].std():.3f}")

print(f"\nOVERALL ({len(rdf)} runs):")
print(f"  MAPE(idx): mean={rdf['mape_index'].mean():.1f}% med={rdf['mape_index'].median():.1f}%")
print(f"  R²: mean={rdf['r2'].mean():.3f} med={rdf['r2'].median():.3f}")

print("\nBest 10:")
for _, r in rdf.nsmallest(10, 'mape_index').iterrows():
    print(f"  {r['country']} {int(r['base_year'])}→{int(r['end_year'])}: MAPE={r['mape_index']:.1f}% R²={r['r2']:.3f}")

print("\nWorst 10:")
for _, r in rdf.nlargest(10, 'mape_index').iterrows():
    print(f"  {r['country']} {int(r['base_year'])}→{int(r['end_year'])}: MAPE={r['mape_index']:.1f}% R²={r['r2']:.3f}")

print("\nDone!")

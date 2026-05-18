"""
Three-variant model + Visualizations.

Variants:
  1. Fixed-parameter: all params frozen at 1970
  2. Tempo-invariant: TFR, e₀, σ updated decadally; MAC frozen at 1970
  3. Tempo-responsive: all params (TFR, e₀, MAC, σ) updated decadally

Data source: UN Population Data Portal API (WPP 2024).
"""
import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar
from scipy.stats import norm
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os, warnings, json, time
warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FIG_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), 'figures')
os.makedirs(FIG_DIR, exist_ok=True)
DATA_DIR = os.path.join(os.path.dirname(SCRIPT_DIR), 'data')
os.makedirs(DATA_DIR, exist_ok=True)

country_ids = {
    "Australia":36,"Austria":40,"Belgium":56,"Canada":124,"Chile":152,
    "China":156,"Colombia":170,"Costa Rica":188,"Czechia":203,
    "DRC":180,"Denmark":208,"Estonia":233,"Finland":246,"France":250,
    "Germany":276,"Greece":300,"Hungary":348,"Iceland":352,
    "Ireland":372,"Israel":376,"Italy":380,"Japan":392,
    "Republic of Korea":410,"Latvia":428,"Lithuania":440,"Luxembourg":442,
    "Mexico":484,"Netherlands":528,"New Zealand":554,"Norway":578,
    "Poland":616,"Portugal":620,"Slovakia":703,"Slovenia":705,
    "Spain":724,"Sweden":752,"Switzerland":756,"Türkiye":792,
    "United Kingdom":826,"United States":840
}
id_to_name = {v:k for k,v in country_ids.items()}

# ============================================================
# DATA FETCHING FROM UN API
# ============================================================
def fetch_un_indicator(ind_id, loc_ids_str, extra_params=None, label=""):
    """Fetch all pages from the UN Data Portal API."""
    import requests
    base = "https://population.un.org/dataportalapi/api/v1"
    all_data = []
    page = 1
    params = {
        "locationIds": loc_ids_str,
        "startYear": 1950, "endYear": 2025,
        "pageSize": 100, "page": page,
    }
    if extra_params:
        params.update(extra_params)
    while True:
        params["page"] = page
        for attempt in range(5):
            try:
                resp = requests.get(f"{base}/data/indicators/{ind_id}",
                                    params=params, timeout=60)
                if resp.status_code == 200:
                    break
            except Exception:
                time.sleep(2)
        else:
            print(f"  Failed after 5 attempts on page {page}")
            break
        if resp.status_code != 200:
            print(f"  HTTP {resp.status_code} on page {page}")
            break
        result = resp.json()
        items = result.get("data", [])
        if not items:
            break
        all_data.extend(items)
        total_pages = result.get("pages", 1)
        if page % 10 == 0 or page == total_pages:
            print(f"  {label}: page {page}/{total_pages} ({len(all_data)} rows)", flush=True)
        if page >= total_pages:
            break
        page += 1
        time.sleep(0.15)
    return all_data

def load_or_fetch_data():
    """Load cached data or fetch from UN API."""
    demo_path = os.path.join(DATA_DIR, 'demo_filtered.csv')
    pop_path = os.path.join(DATA_DIR, 'pop_age5_filtered.csv')

    if os.path.exists(demo_path) and os.path.exists(pop_path):
        print("Loading cached data...", flush=True)
        demo = pd.read_csv(demo_path, low_memory=False)
        pop5 = pd.read_csv(pop_path, low_memory=False)
        print(f"  Demo: {demo.shape}, PopAge5: {pop5.shape}", flush=True)
        return demo, pop5

    print("Fetching data from UN Data Portal API...", flush=True)
    loc_str = ",".join(str(v) for v in country_ids.values())

    # Fetch demographic indicators
    # 49=TotalPop, 19=TFR, 61=LEx, 18=MAC
    # We fetch them separately since they have different dimensions
    print("Fetching total population (ind 49)...", flush=True)
    pop_data = fetch_un_indicator(49, loc_str, {"sexId": "3"}, "TotalPop")

    print("Fetching TFR (ind 19)...", flush=True)
    tfr_data = fetch_un_indicator(19, loc_str, {}, "TFR")

    print("Fetching life expectancy (ind 61)...", flush=True)
    le_data = fetch_un_indicator(61, loc_str, {"sexId": "3"}, "LEx")

    print("Fetching MAC (ind 18)...", flush=True)
    mac_data = fetch_un_indicator(18, loc_str, {}, "MAC")

    # Build demo dataframe
    def to_df(data, val_col):
        rows = []
        for item in data:
            rows.append({
                'LocID': item['locationId'],
                'Time': item['timeLabel'],
                val_col: item['value'],
            })
        df = pd.DataFrame(rows)
        df['Time'] = pd.to_numeric(df['Time'], errors='coerce')
        return df.dropna(subset=['Time'])

    pop_df = to_df(pop_data, 'TPopulation1July')  # in thousands
    tfr_df = to_df(tfr_data, 'TFR')
    le_df = to_df(le_data, 'LEx')
    mac_df = to_df(mac_data, 'MAC')

    # For population, the API may return values not in thousands — check
    # UN API returns population in thousands for indicator 49
    demo = pop_df.merge(tfr_df, on=['LocID','Time'], how='outer')
    demo = demo.merge(le_df, on=['LocID','Time'], how='outer')
    demo = demo.merge(mac_df, on=['LocID','Time'], how='outer')
    demo['Country'] = demo['LocID'].map(id_to_name)
    demo = demo.dropna(subset=['Country'])

    # Female population ratio — approximate as 0.5 (close enough for model)
    demo['FemRatio'] = 0.5

    # Fetch population by age (5-year groups, ind 46)
    print("Fetching population by age (ind 46)...", flush=True)
    age_data = fetch_un_indicator(46, loc_str, {"sexId": "3"}, "PopAge5")

    age_rows = []
    for item in age_data:
        age_rows.append({
            'LocID': item['locationId'],
            'Time': item['timeLabel'],
            'AgeGrp': item.get('ageLabel', ''),
            'AgeGrpStart': item.get('ageStart', 0),
            'PopTotal': item['value'],  # in thousands
        })
    pop5 = pd.DataFrame(age_rows)
    pop5['Time'] = pd.to_numeric(pop5['Time'], errors='coerce')
    pop5 = pop5.dropna(subset=['Time'])

    # Save
    demo.to_csv(demo_path, index=False)
    pop5.to_csv(pop_path, index=False)
    print(f"Saved: Demo {demo.shape}, PopAge5 {pop5.shape}", flush=True)
    return demo, pop5

demo_f, pop_age5 = load_or_fetch_data()

# ============================================================
# MODEL FUNCTIONS
# ============================================================
def gompertz_survival(x, a_g, b_g):
    x = np.asarray(x, dtype=float)
    return np.exp(-(a_g/b_g)*np.expm1(b_g*x))

def calibrate_gompertz(target_le, b_g=0.085):
    def obj(log_a):
        a = np.exp(log_a)
        ages = np.arange(0, 111, dtype=float)
        return (np.trapezoid(gompertz_survival(ages, a, b_g), ages) - target_le)**2
    res = minimize_scalar(obj, bounds=(-14,-1), method='bounded')
    return np.exp(res.x), b_g

def make_asfr(mac, sigma, tfr, max_age=100):
    asfr = np.zeros(max_age+1)
    fertile = np.arange(15,50)
    pdf = norm.pdf(fertile, loc=mac, scale=sigma)
    t = pdf.sum()
    if t > 0: asfr[15:50] = pdf*(tfr/t)
    return asfr

def get_init_pop(loc_id, year, max_age=100):
    sub = pop_age5[(pop_age5['LocID']==loc_id)&(pop_age5['Time']==year)]
    if len(sub)==0:
        # Try mid-period year (WPP uses .5 for July 1)
        sub = pop_age5[(pop_age5['LocID']==loc_id)&(pop_age5['Time'].between(year-0.6, year+0.6))]
    if len(sub)==0: return None
    pop0 = np.zeros(max_age+1)
    for _,r in sub.iterrows():
        grp = str(r['AgeGrp']).strip()
        val = float(r['PopTotal'])*1000  # WPP values in thousands
        if '-' in grp:
            parts = grp.split('-')
            try:
                a_s,a_e = int(parts[0]),int(parts[1])
            except ValueError:
                continue
            span = a_e - a_s + 1
            for a in range(a_s, min(a_e+1, max_age+1)):
                pop0[a] = val/span
        elif '+' in grp:
            try:
                a_s = int(grp.replace('+',''))
            except ValueError:
                continue
            pop0[min(a_s,max_age)] += val
    return pop0

def get_sigma(tfr):
    if tfr>4: return 6.5
    elif tfr>2.5: return 5.5
    elif tfr<1.5: return 4.5
    return 5.0

def make_params(loc_id, year):
    row = demo_f[(demo_f['LocID']==loc_id)&(demo_f['Time']==year)]
    if len(row)==0:
        for dy in range(-2,3):
            row = demo_f[(demo_f['LocID']==loc_id)&(demo_f['Time']==year+dy)]
            if len(row)>0: break
    if len(row)==0: return None
    row = row.iloc[0]
    tfr = float(row['TFR']) if pd.notna(row['TFR']) else None
    le = float(row['LEx']) if pd.notna(row['LEx']) else None
    mac = float(row['MAC']) if pd.notna(row['MAC']) else None
    if tfr is None or le is None or mac is None: return None
    if 'TPopulationFemale1July' in row.index and pd.notna(row.get('TPopulationFemale1July')) and pd.notna(row.get('TPopulation1July')) and row['TPopulation1July'] > 0:
        fem = float(row['TPopulationFemale1July']) / float(row['TPopulation1July'])
    else:
        fem = 0.5
    sigma = get_sigma(tfr)
    a_g,b_g = calibrate_gompertz(le)
    surv_arr = gompertz_survival(np.arange(0,102,dtype=float), a_g, b_g)
    asfr = make_asfr(mac, sigma, tfr)
    sr = np.zeros(101)
    for a in range(1,101):
        sr[a] = min(1.0, max(0.0, surv_arr[a]/surv_arr[a-1])) if surv_arr[a-1]>1e-15 else 0
    inf_surv = min(1.0, surv_arr[1]/surv_arr[0]) if surv_arr[0]>1e-15 else 0.95
    return {'asfr':asfr,'sr':sr,'inf_surv':inf_surv,'fem':fem,
            'tfr':tfr,'le':le,'mac':mac,'sigma':sigma}

def _step(pop_prev, p):
    """One time step of the renewal model."""
    pop_next = np.zeros(101)
    births = np.sum(pop_prev[15:50]*p['fem']*p['asfr'][15:50])
    pop_next[1:] = pop_prev[:-1]*p['sr'][1:]
    pop_next[0] = max(0, births*p['inf_surv'])
    return pop_next

# --- Three model variants ---

def run_tempo_responsive(loc_id, start=1970, end=2023):
    """All 4 params (TFR, e₀, MAC, σ) updated decadally."""
    pop0 = get_init_pop(loc_id, start)
    if pop0 is None: return None, None
    n = end-start; pop = np.zeros((n+1,101)); pop[0]=pop0
    total = np.zeros(n+1); total[0]=pop0.sum()
    cache = {}
    for t in range(1,n+1):
        py = start+((t-1)//10)*10
        if py not in cache:
            p = make_params(loc_id, py)
            if p is None: return None, None
            cache[py] = p
        pop[t] = _step(pop[t-1], cache[py])
        total[t] = pop[t].sum()
    return np.arange(start, start+n+1), total

def run_tempo_invariant(loc_id, start=1970, end=2023):
    """TFR, e₀, σ updated decadally; MAC frozen at 1970."""
    pop0 = get_init_pop(loc_id, start)
    if pop0 is None: return None, None
    p0 = make_params(loc_id, start)
    if p0 is None: return None, None
    mac_1970 = p0['mac']

    n = end-start; pop = np.zeros((n+1,101)); pop[0]=pop0
    total = np.zeros(n+1); total[0]=pop0.sum()
    cache = {}
    for t in range(1,n+1):
        py = start+((t-1)//10)*10
        if py not in cache:
            p = make_params(loc_id, py)
            if p is None: return None, None
            # Override MAC with 1970 value, rebuild ASFR
            p = dict(p)  # copy
            p['asfr'] = make_asfr(mac_1970, p['sigma'], p['tfr'])
            p['mac'] = mac_1970
            cache[py] = p
        pop[t] = _step(pop[t-1], cache[py])
        total[t] = pop[t].sum()
    return np.arange(start, start+n+1), total

def run_fixed_parameter(loc_id, start=1970, end=2023):
    """All params frozen at start year."""
    pop0 = get_init_pop(loc_id, start)
    if pop0 is None: return None, None
    p = make_params(loc_id, start)
    if p is None: return None, None
    n = end-start; pop = np.zeros((n+1,101)); pop[0]=pop0
    total = np.zeros(n+1); total[0]=pop0.sum()
    for t in range(1,n+1):
        pop[t] = _step(pop[t-1], p)
        total[t] = pop[t].sum()
    return np.arange(start, start+n+1), total

def _compute_tfr_star(loc_id, year):
    """Compute Bongaarts-Feeney tempo-adjusted TFR: TFR* = TFR / (1 - dMAC/dt).

    dMAC/dt is estimated from the MAC values surrounding the target year.
    Clipped to [-0.5, 0.5] to avoid extreme values.
    """
    row_cur = demo_f[(demo_f['LocID']==loc_id)&(demo_f['Time']==year)]
    if len(row_cur)==0: return None, None
    mac_cur = float(row_cur.iloc[0]['MAC']) if pd.notna(row_cur.iloc[0]['MAC']) else None
    tfr_cur = float(row_cur.iloc[0]['TFR']) if pd.notna(row_cur.iloc[0]['TFR']) else None
    if mac_cur is None or tfr_cur is None: return None, None

    # Estimate dMAC/dt using centred difference over ±5 years (or available span)
    best_dmac = None
    for span in [5, 4, 3, 2, 1]:
        row_prev = demo_f[(demo_f['LocID']==loc_id)&(demo_f['Time']==year-span)]
        row_next = demo_f[(demo_f['LocID']==loc_id)&(demo_f['Time']==year+span)]
        if len(row_prev)>0 and len(row_next)>0:
            mac_p = float(row_prev.iloc[0]['MAC'])
            mac_n = float(row_next.iloc[0]['MAC'])
            if pd.notna(mac_p) and pd.notna(mac_n):
                best_dmac = (mac_n - mac_p) / (2 * span)
                break
    if best_dmac is None:
        # Fallback: forward or backward difference
        for span in [1, 2, 3, 5]:
            row_next = demo_f[(demo_f['LocID']==loc_id)&(demo_f['Time']==year+span)]
            if len(row_next)>0 and pd.notna(row_next.iloc[0]['MAC']):
                best_dmac = (float(row_next.iloc[0]['MAC']) - mac_cur) / span
                break
    if best_dmac is None:
        for span in [1, 2, 3, 5]:
            row_prev = demo_f[(demo_f['LocID']==loc_id)&(demo_f['Time']==year-span)]
            if len(row_prev)>0 and pd.notna(row_prev.iloc[0]['MAC']):
                best_dmac = (mac_cur - float(row_prev.iloc[0]['MAC'])) / span
                break
    if best_dmac is None:
        return tfr_cur, 0.0  # No MAC change info → TFR* = TFR

    # Clip to prevent extreme values
    best_dmac = np.clip(best_dmac, -0.5, 0.5)

    # Bongaarts-Feeney: TFR* = TFR / (1 - dMAC/dt)
    denom = 1.0 - best_dmac
    if abs(denom) < 0.1:
        denom = 0.1 * np.sign(denom) if denom != 0 else 0.1
    tfr_star = tfr_cur / denom
    # Bound TFR* to reasonable range
    tfr_star = np.clip(tfr_star, 0.3, 10.0)
    return tfr_star, best_dmac

def run_tempo_adjusted(loc_id, start=1970, end=2023):
    """All params updated decadally using tempo-adjusted TFR* instead of period TFR.

    TFR* = TFR / (1 - dMAC/dt)  [Bongaarts & Feeney 1998]
    This removes the tempo distortion from the fertility input, revealing quantum fertility.
    MAC, e₀, σ are updated as in the tempo-responsive model.
    """
    pop0 = get_init_pop(loc_id, start)
    if pop0 is None: return None, None
    n = end-start; pop = np.zeros((n+1,101)); pop[0]=pop0
    total = np.zeros(n+1); total[0]=pop0.sum()
    cache = {}
    for t in range(1,n+1):
        py = start+((t-1)//10)*10
        if py not in cache:
            p = make_params(loc_id, py)
            if p is None: return None, None
            # Replace TFR with TFR*
            tfr_star, dmac = _compute_tfr_star(loc_id, py)
            if tfr_star is not None:
                p = dict(p)
                p['tfr_star'] = tfr_star
                p['dmac'] = dmac
                p['asfr'] = make_asfr(p['mac'], p['sigma'], tfr_star)
            cache[py] = p
        pop[t] = _step(pop[t-1], cache[py])
        total[t] = pop[t].sum()
    return np.arange(start, start+n+1), total

def calc_metrics(yrs, pops, actual_years, actual_pop):
    common = np.intersect1d(yrs, actual_years)
    if len(common)<5: return None
    mv = np.array([pops[np.where(yrs==y)[0][0]] for y in common])
    av = np.array([actual_pop[np.where(actual_years==y)[0][0]] for y in common])
    if av[0]==0: return None
    mi=mv/mv[0]; ai=av/av[0]
    mape = np.mean(np.abs((mi-ai)/ai))*100
    mape_abs = np.mean(np.abs((mv-av)/av))*100
    ratio = mv[-1]/av[-1] if av[-1]>0 else np.nan
    return {'mape_index':mape,'mape_abs':mape_abs,'final_ratio':ratio}

# ============================================================
# RUN ALL MODELS
# ============================================================
print("\nRunning models for all 40 countries...", flush=True)
START=1970; END=2023
all_r = {}
resp_res=[]; inv_res=[]; fix_res=[]; adj_res=[]

for cn, lid in sorted(country_ids.items()):
    sub = demo_f[(demo_f['LocID']==lid)&(demo_f['Time']>=START)&(demo_f['Time']<=END)].sort_values('Time')
    ay = sub['Time'].values
    ap = sub['TPopulation1July'].values*1000 if 'TPopulation1July' in sub.columns else np.zeros(len(ay))

    ry, rp = run_tempo_responsive(lid, START, END)
    iy, ip = run_tempo_invariant(lid, START, END)
    fy, fp = run_fixed_parameter(lid, START, END)
    ay2, ap2 = run_tempo_adjusted(lid, START, END)

    all_r[cn] = {'ay':ay,'ap':ap, 'ry':ry,'rp':rp, 'iy':iy,'ip':ip,
                 'fy':fy,'fp':fp, 'ay2':ay2,'ap2':ap2}

    if ry is not None:
        m = calc_metrics(ry,rp,ay,ap)
        if m: resp_res.append({**m,'country':cn})
    if iy is not None:
        m = calc_metrics(iy,ip,ay,ap)
        if m: inv_res.append({**m,'country':cn})
    if fy is not None:
        m = calc_metrics(fy,fp,ay,ap)
        if m: fix_res.append({**m,'country':cn})
    if ay2 is not None:
        m = calc_metrics(ay2,ap2,ay,ap)
        if m: adj_res.append({**m,'country':cn})
    print(f"  {cn}: done", flush=True)

resp_df = pd.DataFrame(resp_res)
inv_df = pd.DataFrame(inv_res)
fix_df = pd.DataFrame(fix_res)
adj_df = pd.DataFrame(adj_res)

print(f"\nTempo-responsive: N={len(resp_df)}, MAPE mean={resp_df['mape_index'].mean():.1f}% med={resp_df['mape_index'].median():.1f}%", flush=True)
print(f"Tempo-invariant:  N={len(inv_df)}, MAPE mean={inv_df['mape_index'].mean():.1f}% med={inv_df['mape_index'].median():.1f}%", flush=True)
print(f"Fixed-parameter:  N={len(fix_df)}, MAPE mean={fix_df['mape_index'].mean():.1f}% med={fix_df['mape_index'].median():.1f}%", flush=True)
print(f"Tempo-adjusted:   N={len(adj_df)}, MAPE mean={adj_df['mape_index'].mean():.1f}% med={adj_df['mape_index'].median():.1f}%", flush=True)

# ============================================================
# VISUALISATIONS
# ============================================================
print("\nGenerating figures...", flush=True)

COLORS = {
    'observed': 'black',
    'responsive': '#2166ac',   # blue
    'invariant': '#e08214',    # orange
    'fixed': '#b2182b',        # red
    'adjusted': '#4dac26',     # green
}

# Fig 1: 6-panel showcase
showcase = ["Japan","China","United States","Republic of Korea","Germany","DRC"]
fig, axes = plt.subplots(2,3,figsize=(18,10))
for i,cn in enumerate(showcase):
    ax=axes.flatten()[i]; r=all_r.get(cn,{})
    if 'ay' in r and len(r['ay'])>0:
        ax.plot(r['ay'],r['ap']/1e6,'k-',lw=2.5,label='Observed (UN WPP 2024)')
    if r.get('ry') is not None:
        ax.plot(r['ry'],r['rp']/1e6,'-',color=COLORS['responsive'],lw=1.8,label='Tempo-responsive')
    if r.get('iy') is not None:
        ax.plot(r['iy'],r['ip']/1e6,'--',color=COLORS['invariant'],lw=1.8,label='Tempo-invariant')
    if r.get('ay2') is not None:
        ax.plot(r['ay2'],r['ap2']/1e6,'-.',color=COLORS['adjusted'],lw=1.8,label='Tempo-adjusted (TFR*)')
    if r.get('fy') is not None:
        ax.plot(r['fy'],r['fp']/1e6,':',color=COLORS['fixed'],lw=1.5,label='Fixed-parameter (1970)')
    ax.set_title(cn,fontsize=14,fontweight='bold')
    ax.set_xlabel('Year'); ax.set_ylabel('Population (millions)')
    ax.legend(fontsize=6,loc='best'); ax.grid(True,alpha=0.3)
    ax.set_ylim(bottom=0)
fig.suptitle('Five Model Variants vs Observed Population, 1970\u20132023',fontsize=16,fontweight='bold',y=0.99)
fig.tight_layout(rect=[0, 0, 1, 0.95])
fig.savefig(os.path.join(FIG_DIR,'fig1_showcase.png'),dpi=300,bbox_inches='tight')
fig.savefig(os.path.join(FIG_DIR,'Figure_1.eps'),dpi=300,bbox_inches='tight',format='eps')
plt.close()
print("  Fig1 done", flush=True)

# Fig 2: All 40 countries
fig, axes = plt.subplots(8,5,figsize=(25,32)); af=axes.flatten()
for i,cn in enumerate(sorted(country_ids.keys())):
    if i>=40: break
    ax=af[i]; r=all_r.get(cn,{})
    if 'ay' in r and len(r['ay'])>0:
        ax.plot(r['ay'],r['ap']/1e6,'k-',lw=1.5)
    if r.get('ry') is not None:
        ax.plot(r['ry'],r['rp']/1e6,'-',color=COLORS['responsive'],lw=1)
    if r.get('iy') is not None:
        ax.plot(r['iy'],r['ip']/1e6,'--',color=COLORS['invariant'],lw=1)
    if r.get('ay2') is not None:
        ax.plot(r['ay2'],r['ap2']/1e6,'-.',color=COLORS['adjusted'],lw=1)
    if r.get('fy') is not None:
        ax.plot(r['fy'],r['fp']/1e6,':',color=COLORS['fixed'],lw=1)
    dr=resp_df[resp_df['country']==cn]
    if len(dr)>0:
        ax.text(0.95,0.95,f'MAPE={dr.iloc[0]["mape_index"]:.1f}%',
                transform=ax.transAxes,fontsize=7,va='top',ha='right',
                bbox=dict(boxstyle='round,pad=0.2',facecolor='lightyellow',alpha=0.8))
    ax.set_title(cn,fontsize=9,fontweight='bold'); ax.tick_params(labelsize=7)
    ax.grid(True,alpha=0.2); ax.set_ylim(bottom=0)
fig.suptitle('All Countries: Observed (black) Responsive (blue) Invariant (orange) Adjusted (green) Fixed (red)',
             fontsize=13,fontweight='bold',y=0.995)
fig.tight_layout(rect=[0, 0, 1, 0.98])
fig.savefig(os.path.join(FIG_DIR,'fig2_all_countries.png'),dpi=300,bbox_inches='tight')
fig.savefig(os.path.join(FIG_DIR,'Figure_2.eps'),dpi=300,bbox_inches='tight',format='eps')
plt.close()
print("  Fig2 done", flush=True)

# Fig 3: MAPE heatmap (fixed-parameter by base year)
# Run fixed-parameter for multiple base years
BASE_YEARS = [1970,1980,1990,2000]
static_results = []
for cn, lid in sorted(country_ids.items()):
    sub = demo_f[(demo_f['LocID']==lid)&(demo_f['Time']>=1970)&(demo_f['Time']<=END)].sort_values('Time')
    ay = sub['Time'].values
    ap = sub['TPopulation1July'].values*1000 if 'TPopulation1July' in sub.columns else np.zeros(len(ay))
    for by in BASE_YEARS:
        fy2, fp2 = run_fixed_parameter(lid, by, END)
        if fy2 is not None:
            m = calc_metrics(fy2,fp2,ay,ap)
            if m:
                static_results.append({**m,'country':cn,'base_year':by,'horizon':END-by})

static_csv = pd.DataFrame(static_results)
if len(static_csv) > 0:
    pivot = static_csv.pivot_table(values='mape_index',index='country',columns='base_year')
    fig,ax = plt.subplots(figsize=(10,16))
    im = ax.imshow(pivot.values,cmap='RdYlGn_r',aspect='auto',vmin=0,vmax=30)
    ax.set_xticks(range(len(pivot.columns))); ax.set_xticklabels([int(c) for c in pivot.columns])
    ax.set_yticks(range(len(pivot.index))); ax.set_yticklabels(pivot.index,fontsize=9)
    for ii in range(len(pivot.index)):
        for jj in range(len(pivot.columns)):
            v=pivot.values[ii,jj]
            if not np.isnan(v):
                ax.text(jj,ii,f'{v:.1f}',ha='center',va='center',fontsize=7,
                        color='white' if v>20 else 'black')
    plt.colorbar(im,ax=ax,label='MAPE (%)',shrink=0.8)
    ax.set_title('Fixed-Parameter Model MAPE (%) by Country \u00d7 Base Year',fontsize=14,fontweight='bold')
    ax.set_xlabel('Base Year')
    plt.tight_layout()
    plt.savefig(os.path.join(FIG_DIR,'fig3_heatmap.png'),dpi=300,bbox_inches='tight')
    plt.savefig(os.path.join(FIG_DIR,'Figure_3.eps'),dpi=300,bbox_inches='tight',format='eps')
    plt.close()
    print("  Fig3 done", flush=True)
    static_csv.to_csv(os.path.join(DATA_DIR,'model_fit_results.csv'),index=False)

# Fig 4: Four-variant comparison bar chart (with tempo-adjusted)
merged = resp_df.merge(inv_df,on='country',suffixes=('_resp','_inv'))
merged = merged.merge(fix_df,on='country')
merged.rename(columns={'mape_index':'mape_index_fix','final_ratio':'final_ratio_fix'},inplace=True)
merged = merged.merge(adj_df[['country','mape_index','final_ratio']],on='country',how='left')
merged.rename(columns={'mape_index':'mape_index_adj','final_ratio':'final_ratio_adj'},inplace=True)
fig,(ax1,ax2) = plt.subplots(1,2,figsize=(24,14))
ms = merged.sort_values('mape_index_resp')
x=np.arange(len(ms)); w=0.2
ax1.barh(x-1.5*w,ms['mape_index_fix'],w,label='Fixed-parameter (1970)',color=COLORS['fixed'],alpha=0.8)
ax1.barh(x-0.5*w,ms['mape_index_inv'],w,label='Tempo-invariant',color=COLORS['invariant'],alpha=0.8)
ax1.barh(x+0.5*w,ms['mape_index_resp'],w,label='Tempo-responsive',color=COLORS['responsive'],alpha=0.8)
ax1.barh(x+1.5*w,ms['mape_index_adj'],w,label='Tempo-adjusted (TFR*)',color=COLORS['adjusted'],alpha=0.8)
ax1.set_yticks(x); ax1.set_yticklabels(ms['country'],fontsize=8)
ax1.set_xlabel('MAPE (%)'); ax1.set_title('MAPE: Four Model Variants',fontweight='bold')
ax1.legend(fontsize=9); ax1.grid(True,alpha=0.3,axis='x')

# Scatter: final ratio comparison
ax2.scatter(ms['final_ratio_fix'],ms['final_ratio_resp'],s=60,alpha=0.7,
            c=COLORS['fixed'],edgecolors='black',lw=0.5,label='Fixed-param',marker='s')
ax2.scatter(ms['final_ratio_inv'],ms['final_ratio_resp'],s=60,alpha=0.7,
            c=COLORS['invariant'],edgecolors='black',lw=0.5,label='Tempo-invariant',marker='o')
ax2.scatter(ms['final_ratio_adj'],ms['final_ratio_resp'],s=60,alpha=0.7,
            c=COLORS['adjusted'],edgecolors='black',lw=0.5,label='Tempo-adjusted (TFR*)',marker='D')
ax2.axhline(1,color='gray',ls='--',alpha=0.5); ax2.axvline(1,color='gray',ls='--',alpha=0.5)
ax2.plot([0.4,3],[0.4,3],'k:',alpha=0.3)
for _,rr in ms.iterrows():
    ax2.annotate(rr['country'],(rr['final_ratio_inv'],rr['final_ratio_resp']),fontsize=5,alpha=0.6)
ax2.set_xlabel('Comparator: Final Ratio'); ax2.set_ylabel('Tempo-responsive: Final Ratio')
ax2.set_title('Final Population Ratio (Model / Observed, 2023)',fontweight='bold')
ax2.legend(fontsize=8); ax2.grid(True,alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(FIG_DIR,'fig4_comparison.png'),dpi=300,bbox_inches='tight')
plt.savefig(os.path.join(FIG_DIR,'Figure_4.eps'),dpi=300,bbox_inches='tight',format='eps')
plt.close()
print("  Fig4 done", flush=True)

# Fig 5: Bias analysis (base year 2000 fixed-parameter)
if len(static_csv) > 0:
    s2k = static_csv[static_csv['base_year']==2000]
    if len(s2k) > 0:
        fig,axes = plt.subplots(1,3,figsize=(18,6))
        for ax,col,color,title in [(axes[0],'tfr','steelblue','(A) Fit vs TFR'),
                                    (axes[1],'le','coral','(B) Fit vs LE'),
                                    (axes[2],'mac','forestgreen','(C) Bias vs MAC')]:
            # Merge with demo params
            for _,r in s2k.iterrows():
                cn = r['country']
                lid = country_ids.get(cn)
                if lid is None: continue
                row2 = demo_f[(demo_f['LocID']==lid)&(demo_f['Time']==2000)]
                if len(row2)==0: continue
                row2 = row2.iloc[0]
                xval = float(row2.get(col.upper() if col!='le' else 'LEx', np.nan))
                if col == 'tfr': xval = float(row2.get('TFR', np.nan))
                elif col == 'le': xval = float(row2.get('LEx', np.nan))
                elif col == 'mac': xval = float(row2.get('MAC', np.nan))
                if np.isnan(xval): continue
                ycol_val = r['mape_index'] if 'Fit' in title else r['final_ratio']
                ax.scatter(xval, ycol_val, s=40, alpha=0.7, c=color, edgecolors='black', lw=0.5)
                ax.annotate(cn[:3],(xval,ycol_val),fontsize=6,alpha=0.6)
            ax.set_xlabel(f'{col.upper()} (2000)')
            ax.set_ylabel('MAPE (%)' if 'Fit' in title else 'Final Ratio')
            ax.set_title(title,fontweight='bold'); ax.grid(True,alpha=0.3)
            if 'Bias' in title: ax.axhline(1,color='gray',ls='--',alpha=0.5)
        plt.suptitle('Model Bias Analysis (Fixed-Parameter, Base Year=2000)',fontsize=14,fontweight='bold')
        plt.tight_layout()
        plt.savefig(os.path.join(FIG_DIR,'fig5_bias.png'),dpi=300,bbox_inches='tight')
        plt.savefig(os.path.join(FIG_DIR,'Figure_5.eps'),dpi=300,bbox_inches='tight',format='eps')
        plt.close()
        print("  Fig5 done", flush=True)

# Save summaries
resp_df.to_csv(os.path.join(DATA_DIR,'model_fit_responsive_results.csv'),index=False)
inv_df.to_csv(os.path.join(DATA_DIR,'model_fit_invariant_results.csv'),index=False)
fix_df.to_csv(os.path.join(DATA_DIR,'model_fit_fixed_results.csv'),index=False)
adj_df.to_csv(os.path.join(DATA_DIR,'model_fit_adjusted_results.csv'),index=False)

summary = []
for by in BASE_YEARS:
    sub=static_csv[static_csv['base_year']==by] if len(static_csv)>0 else pd.DataFrame()
    if len(sub)==0: continue
    summary.append({'Variant':f'Fixed-param ({by})','Horizon':int(sub['horizon'].median()),'N':len(sub),
        'MAPE Mean':f"{sub['mape_index'].mean():.1f}",'MAPE Median':f"{sub['mape_index'].median():.1f}",
        'Ratio Mean':f"{sub['final_ratio'].mean():.3f}",'Ratio Std':f"{sub['final_ratio'].std():.3f}"})
summary.append({'Variant':'Tempo-invariant','Horizon':53,'N':len(inv_df),
    'MAPE Mean':f"{inv_df['mape_index'].mean():.1f}",'MAPE Median':f"{inv_df['mape_index'].median():.1f}",
    'Ratio Mean':f"{inv_df['final_ratio'].mean():.3f}",'Ratio Std':f"{inv_df['final_ratio'].std():.3f}"})
summary.append({'Variant':'Tempo-responsive','Horizon':53,'N':len(resp_df),
    'MAPE Mean':f"{resp_df['mape_index'].mean():.1f}",'MAPE Median':f"{resp_df['mape_index'].median():.1f}",
    'Ratio Mean':f"{resp_df['final_ratio'].mean():.3f}",'Ratio Std':f"{resp_df['final_ratio'].std():.3f}"})
summary.append({'Variant':'Tempo-adjusted (TFR*)','Horizon':53,'N':len(adj_df),
    'MAPE Mean':f"{adj_df['mape_index'].mean():.1f}",'MAPE Median':f"{adj_df['mape_index'].median():.1f}",
    'Ratio Mean':f"{adj_df['final_ratio'].mean():.3f}",'Ratio Std':f"{adj_df['final_ratio'].std():.3f}"})
pd.DataFrame(summary).to_csv(os.path.join(DATA_DIR,'model_fit_summary.csv'),index=False)

# Japan-specific results for manuscript
print("\n" + "="*60, flush=True)
print("JAPAN-SPECIFIC RESULTS", flush=True)
print("="*60, flush=True)
for lbl, df in [('Fixed-param', fix_df), ('Tempo-invariant', inv_df),
                ('Tempo-responsive', resp_df), ('Tempo-adjusted', adj_df)]:
    jp = df[df['country']=='Japan']
    if len(jp) > 0:
        j = jp.iloc[0]
        print(f"  {lbl:25s}  MAPE={j['mape_index']:.1f}%  ratio={j['final_ratio']:.3f}", flush=True)

print("\n" + "="*60, flush=True)
print("OVERALL SUMMARY", flush=True)
print("="*60, flush=True)
for s in summary:
    print(f"  {s['Variant']:25s}  N={s['N']:3d}  MAPE={s['MAPE Mean']:>5s}% (med {s['MAPE Median']:>5s}%)  "
          f"ratio={s['Ratio Mean']}\u00b1{s['Ratio Std']}", flush=True)

print("\nAll done!", flush=True)

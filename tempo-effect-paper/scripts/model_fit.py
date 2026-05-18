"""
Endogenous Renewal Model + Gompertz Survival
Fit evaluation against UN WPP 2024 data for OECD countries + China + DRC

Model specification (from ChatGPT conversation):
- Endogenous renewal: population regenerates via births, with age-specific fertility
- Gompertz survival: S(x) = exp(-(b/c)*(exp(c*x) - 1)), parameterized by life expectancy
- AFB (age at first birth) as anchor for fertility schedule
- σ controls spread of fertility around AFB
- Tracks "simultaneously living population" over time
"""

import numpy as np
import pandas as pd
from scipy.optimize import minimize_scalar, minimize
from scipy.integrate import quad
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# 1. LOAD UN WPP 2024 DATA
# ============================================================
print("Loading UN WPP 2024 data...")

demo = pd.read_csv('/home/ubuntu/wpp_data/WPP2024_Demographic_Indicators_Medium.csv.gz', low_memory=False)

# OECD + China + DRC location IDs
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

# Filter for our countries
df = demo[demo['LocID'].isin(country_ids.values())].copy()
df['Country'] = df['LocID'].map(id_to_name)

print(f"Loaded {len(df)} rows for {df['Country'].nunique()} countries")

# ============================================================
# 2. GOMPERTZ SURVIVAL MODEL
# ============================================================

def gompertz_survival(x, a_gomp, b_gomp):
    """
    Gompertz survival function: S(x) = exp(-(a/b)*(exp(b*x) - 1))
    a_gomp = initial mortality rate (Makeham 'a' parameter)
    b_gomp = rate of mortality increase with age
    """
    return np.exp(-(a_gomp / b_gomp) * (np.expm1(b_gomp * x)))

def gompertz_life_expectancy(a_gomp, b_gomp, max_age=120):
    """Compute life expectancy from Gompertz parameters by numerical integration"""
    ages = np.arange(0, max_age + 1, dtype=float)
    surv = gompertz_survival(ages, a_gomp, b_gomp)
    # Trapezoidal integration
    le = np.trapz(surv, ages)
    return le

def calibrate_gompertz_to_le(target_le, b_gomp=0.085):
    """
    Given a target life expectancy, find Gompertz 'a' parameter.
    b is typically around 0.085 for human populations.
    
    Note: This is a simplified Gompertz without infant/child mortality correction.
    We add a simple correction for infant mortality.
    """
    def objective(log_a):
        a = np.exp(log_a)
        le = gompertz_life_expectancy(a, b_gomp)
        return (le - target_le) ** 2
    
    result = minimize_scalar(objective, bounds=(-12, -2), method='bounded')
    return np.exp(result.x), b_gomp

# ============================================================
# 3. FERTILITY SCHEDULE (NORMAL DISTRIBUTION AROUND AFB)
# ============================================================

def fertility_schedule(ages, afb, sigma, tfr):
    """
    Generate age-specific fertility rates as a normal distribution centered on AFB.
    
    ages: array of ages (e.g., 15 to 49)
    afb: age at first birth (mean of fertility schedule)  
    sigma: standard deviation of fertility schedule
    tfr: total fertility rate (integral of ASFR)
    """
    from scipy.stats import norm
    pdf = norm.pdf(ages, loc=afb, scale=sigma)
    # Normalize so that sum * 1 (single-year ages) = TFR
    total = pdf.sum()
    if total > 0:
        asfr = pdf * (tfr / total)
    else:
        asfr = np.zeros_like(ages, dtype=float)
    return asfr

# ============================================================
# 4. ENDOGENOUS RENEWAL MODEL (DISCRETE-TIME)
# ============================================================

def run_renewal_model(init_pop_by_age, asfr_by_age, survival_func,
                      n_years=100, female_ratio=0.4886, max_age=100):
    """
    Discrete-time population renewal model.
    
    init_pop_by_age: initial population by single year of age (array of length max_age+1)
    asfr_by_age: age-specific fertility rates (births per woman per year) for ages 0..max_age
    survival_func: function that returns survival probability S(x) for age x
    female_ratio: fraction of births that are female
    n_years: number of years to simulate
    max_age: maximum age
    
    Returns: total_pop array of length n_years+1
    """
    pop = np.zeros((n_years + 1, max_age + 1))
    pop[0, :] = init_pop_by_age[:max_age + 1]
    
    total_pop = np.zeros(n_years + 1)
    total_pop[0] = pop[0].sum()
    
    for t in range(1, n_years + 1):
        # Births: sum of (female pop at age a) * ASFR(a)
        # Simplification: use total pop * female_ratio for female population
        births = 0.0
        for a in range(15, min(50, max_age + 1)):
            births += pop[t-1, a] * female_ratio * asfr_by_age[a]
        
        # Age the population: everyone moves up one year, with survival
        for a in range(max_age, 0, -1):
            # Survival from age a-1 to a
            s_prev = survival_func(a - 1)
            s_curr = survival_func(a)
            if s_prev > 0:
                survival_rate = s_curr / s_prev
            else:
                survival_rate = 0.0
            survival_rate = max(0, min(1, survival_rate))
            pop[t, a] = pop[t-1, a-1] * survival_rate
        
        # New births enter age 0
        # Apply infant survival
        s0 = survival_func(0)
        s1 = survival_func(1)
        infant_survival = s1 / s0 if s0 > 0 else 0.9
        pop[t, 0] = births * min(1, max(0, infant_survival))
        
        total_pop[t] = pop[t].sum()
    
    return total_pop, pop

# ============================================================
# 5. EXTRACT REAL PARAMETERS FOR EACH COUNTRY
# ============================================================

def extract_country_params(df, loc_id, base_year=2024):
    """Extract key demographic parameters for a country at base_year"""
    row = df[(df['LocID'] == loc_id) & (df['Time'] == base_year)]
    if len(row) == 0:
        # Try closest year
        for y in [2023, 2022, 2021, 2020]:
            row = df[(df['LocID'] == loc_id) & (df['Time'] == y)]
            if len(row) > 0:
                break
    
    if len(row) == 0:
        return None
    
    row = row.iloc[0]
    return {
        'total_pop': row['TPopulation1July'] * 1000,  # Convert to individuals
        'tfr': row['TFR'],
        'le': row['LEx'],
        'mac': row['MAC'],
        'births': row['Births'] * 1000,
        'deaths': row['Deaths'] * 1000,
        'female_pop_ratio': row['TPopulationFemale1July'] / row['TPopulation1July'] if row['TPopulation1July'] > 0 else 0.5,
    }

def get_actual_pop_trajectory(df, loc_id, start_year=1950, end_year=2100):
    """Get actual + projected total population trajectory"""
    country_df = df[(df['LocID'] == loc_id) & 
                    (df['Time'] >= start_year) & 
                    (df['Time'] <= end_year)].sort_values('Time')
    return country_df['Time'].values, country_df['TPopulation1July'].values * 1000

# ============================================================
# 6. RUN MODEL FOR EACH COUNTRY AND COMPARE
# ============================================================

print("\nExtracting parameters and running models...")

# Load population by age data for initial conditions
print("Loading population by age data...")
pop_age = pd.read_csv('/home/ubuntu/wpp_data/WPP2024_PopulationBySingleAgeSex_Medium_1950-2023.csv.gz', low_memory=False)
print(f"Pop by age shape: {pop_age.shape}")
print(f"Pop by age columns: {list(pop_age.columns)}")

# Check sample
sample = pop_age[pop_age['LocID'] == 392].head(3)
print(f"\nSample (Japan):")
print(sample[['LocID', 'Location', 'Time', 'AgeGrp', 'AgeGrpStart', 'PopTotal']].to_string())


"""
Create CurveReexamination objects from real World Bank data.

Replaces representative data with actual WDI data for:
- Preston Curve (#13)
- Kuznets Curve (#3)
- Demographic Transition (#23)
- Forest Transition (#34)
- Second Demographic Transition (#28)
- Engel Curve (#7) - partial
- Lipset Hypothesis (#44) - GDP + democracy proxy

Also updates Phillips, Okun with real US macro time series.
"""

import os
import numpy as np
import pandas as pd
from core_analysis import CurveReexamination

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')

# World Bank aggregate/regional codes (not independent country observations)
WB_AGGREGATE_CODES = {
    'WLD', 'UMC', 'TSS', 'SSA', 'SSF', 'TSA', 'SAS', 'SST', 'PRE', 'PST',
    'PSS', 'OSS', 'OED', 'INX', 'NAC', 'MIC', 'TMN', 'MNA', 'MEA', 'LMC',
    'LIC', 'LMY', 'LDC', 'TLA', 'LAC', 'LCN', 'LTE', 'IDA', 'IDX', 'IDB',
    'IBT', 'IBD', 'HIC', 'HPC', 'FCS', 'EUU', 'TEC', 'ECA', 'ECS', 'EMU',
    'TEA', 'EAP', 'EAS', 'EAR', 'CEB', 'CSS', 'ARB', 'AFW', 'AFE',
}


def _load_csv(fname):
    path = os.path.join(DATA_DIR, fname)
    if os.path.exists(path):
        df = pd.read_csv(path)
        if 'country_code' in df.columns:
            df = df[~df['country_code'].isin(WB_AGGREGATE_CODES)]
        return df
    return None


def get_preston_real():
    """Preston Curve from real World Bank data."""
    df = _load_csv('preston_curve_real.csv')
    if df is None:
        return None
    # Filter to countries with GDP > 500 and reasonable life expectancy
    df = df[(df['gdp_pc_ppp'] > 500) & (df['life_expectancy'] > 30) & (df['life_expectancy'] < 90)]
    df = df.dropna(subset=['gdp_pc_ppp', 'life_expectancy'])
    # Use log GDP for better spread
    x = np.log10(df['gdp_pc_ppp'].values)
    y = df['life_expectancy'].values
    labels = df['country'].values if 'country' in df.columns else None
    return CurveReexamination(
        name="Preston Curve",
        x=x, y=y,
        x_label="log₁₀(GDP per capita PPP, $)",
        y_label="Life expectancy (years)",
        country_labels=labels,
        category="Public Health",
        x_is_logged=True
    ), len(df)


def get_kuznets_real():
    """Kuznets Curve from real World Bank data."""
    df = _load_csv('kuznets_curve_real.csv')
    if df is None:
        return None
    df = df[(df['gdp_pc_ppp'] > 500) & (df['gini'] > 0) & (df['gini'] < 70)]
    df = df.dropna(subset=['gdp_pc_ppp', 'gini'])
    x = np.log10(df['gdp_pc_ppp'].values)
    y = df['gini'].values
    labels = df['country'].values if 'country' in df.columns else None
    return CurveReexamination(
        name="Kuznets Curve",
        x=x, y=y,
        x_label="log₁₀(GDP per capita PPP, $)",
        y_label="Gini coefficient",
        country_labels=labels,
        category="Economics",
        x_is_logged=True
    ), len(df)


def get_demographic_transition_real():
    """Demographic Transition from real WB data: GDP vs TFR."""
    df = _load_csv('demographic_transition_real.csv')
    if df is None:
        return None
    df = df[(df['gdp_pc_ppp'] > 500) & (df['tfr'] > 0) & (df['tfr'] < 10)]
    df = df.dropna(subset=['gdp_pc_ppp', 'tfr'])
    x = np.log10(df['gdp_pc_ppp'].values)
    y = df['tfr'].values
    labels = df['country'].values if 'country' in df.columns else None
    return CurveReexamination(
        name="Demographic Transition (TFR)",
        x=x, y=y,
        x_label="log₁₀(GDP per capita PPP, $)",
        y_label="Total fertility rate",
        country_labels=labels,
        category="Demography",
        x_is_logged=True
    ), len(df)


def get_forest_transition_real():
    """Forest Transition from real WB data."""
    df = _load_csv('forest_transition_real.csv')
    if df is None:
        return None
    df = df[(df['gdp_pc_ppp'] > 500) & (df['forest_pct'] > 0)]
    df = df.dropna(subset=['gdp_pc_ppp', 'forest_pct'])
    x = np.log10(df['gdp_pc_ppp'].values)
    y = df['forest_pct'].values
    labels = df['country'].values if 'country' in df.columns else None
    return CurveReexamination(
        name="Forest Transition Curve",
        x=x, y=y,
        x_label="log₁₀(GDP per capita PPP, $)",
        y_label="Forest area (% of land area)",
        country_labels=labels,
        category="Environmental Science",
        x_is_logged=True
    ), len(df)


def get_second_demographic_transition_real():
    """Second Demographic Transition: among high-HDI only."""
    df = _load_csv('demographic_transition_real.csv')
    if df is None:
        return None
    # Filter to higher-income countries only (GDP > $15,000 PPP)
    df = df[(df['gdp_pc_ppp'] > 15000) & (df['tfr'] > 0) & (df['tfr'] < 5)]
    df = df.dropna(subset=['gdp_pc_ppp', 'tfr'])
    x = np.log10(df['gdp_pc_ppp'].values)
    y = df['tfr'].values
    labels = df['country'].values if 'country' in df.columns else None
    return CurveReexamination(
        name="Second Demographic Transition",
        x=x, y=y,
        x_label="log₁₀(GDP per capita PPP, $)",
        y_label="Total fertility rate",
        country_labels=labels,
        category="Demography",
        x_is_logged=True
    ), len(df)


def _load_us_macro():
    """Load and parse US macro time series, skipping the 'Country' header row."""
    df = _load_csv('us_macro_timeseries.csv')
    if df is None:
        return None
    # First row may be 'Country' / 'United States' labels — skip it
    if df.iloc[0, 0] == 'Country' or (isinstance(df.index[0], str) and 'Country' in str(df.iloc[0])):
        df = df.iloc[1:].copy()
    # Extract year from index (YR1960 format)
    df.index = df.iloc[:, 0] if df.columns[0] == '' or df.columns[0] == 'Unnamed: 0' else df.index
    years = []
    for idx in df.index:
        s = str(idx)
        if s.startswith('YR'):
            years.append(int(s[2:]))
        else:
            years.append(None)
    df['year'] = years
    df = df[df['year'].notna()]
    for col in ['gdp_growth', 'unemployment', 'inflation']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    return df


def get_phillips_real():
    """Phillips Curve from real US macro time series."""
    df = _load_us_macro()
    if df is None:
        return None
    df = df.dropna(subset=['unemployment', 'inflation'])
    if len(df) < 10:
        return None
    x = df['unemployment'].values.astype(float)
    y = df['inflation'].values.astype(float)
    return CurveReexamination(
        name="Phillips Curve",
        x=x, y=y,
        x_label="Unemployment rate (%)",
        y_label="Inflation rate (CPI, %)",
        category="Economics"
    ), len(df)


def get_okun_real():
    """Okun's Law from real US macro time series."""
    df = _load_us_macro()
    if df is None:
        return None
    df = df.dropna(subset=['gdp_growth', 'unemployment'])
    if len(df) < 10:
        return None
    # Okun's law: GDP growth vs change in unemployment
    df = df.sort_values('year')
    unemp = df['unemployment'].values.astype(float)
    unemp_change = np.diff(unemp)
    gdp_growth = df['gdp_growth'].values[1:].astype(float)
    mask = np.isfinite(unemp_change) & np.isfinite(gdp_growth)
    x = gdp_growth[mask]
    y = unemp_change[mask]
    if len(x) < 10:
        return None
    return CurveReexamination(
        name="Okun's Law",
        x=x, y=y,
        x_label="Real GDP growth (%)",
        y_label="Change in unemployment rate (pp)",
        category="Economics"
    ), len(x)


def get_green_revolution_real():
    """Green Revolution yield curve from real WB data."""
    df = _load_csv('global_cereal_yield.csv')
    if df is None:
        return None
    # Parse the WB time series format
    # The format from wbgapi is: economy, YR1961, YR1962, ...
    if 'economy' in df.columns or df.shape[0] == 1:
        # Melt the wide format
        val_cols = [c for c in df.columns if c.startswith('YR')]
        if val_cols:
            vals = df[val_cols].iloc[0]
            years = [int(c.replace('YR', '')) for c in val_cols]
            data = pd.DataFrame({'year': years, 'yield_kg': vals})
            data = data.dropna()
            if len(data) > 10:
                x = data['year'].values
                y = data['yield_kg'].values / 1000  # kg/ha to tonnes/ha
                return CurveReexamination(
                    name="Green Revolution Yield Curve",
                    x=x, y=y,
                    x_label="Year",
                    y_label="Global cereal yield (tonnes/ha)",
                    category="Agriculture"
                ), len(data)
    return None


def get_all_real_data():
    """Return dict of curve_id -> (CurveReexamination, N) for all real data."""
    real = {}
    fetchers = {
        'preston': get_preston_real,
        'kuznets': get_kuznets_real,
        'demographic_transition': get_demographic_transition_real,
        'forest_transition': get_forest_transition_real,
        'second_demographic_transition': get_second_demographic_transition_real,
        'phillips': get_phillips_real,
        'okun': get_okun_real,
        'green_revolution': get_green_revolution_real,
    }
    for key, func in fetchers.items():
        try:
            result = func()
            if result is not None:
                real[key] = result
                print(f"  Real data loaded: {key} (N={result[1]})")
        except Exception as e:
            print(f"  Warning: {key} real data failed: {e}")
    return real


if __name__ == "__main__":
    real = get_all_real_data()
    print(f"\nLoaded {len(real)} curves with real data")

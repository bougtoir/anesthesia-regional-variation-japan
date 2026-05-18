"""
Data collection and analysis for Public Health / Epidemiology curves (B13-B22).
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_preston_curve_data():
    """#13 Preston Curve: Income vs Life expectancy (cross-country)."""
    countries = ['Norway', 'Switzerland', 'Australia', 'Japan', 'Sweden',
                 'Spain', 'Italy', 'France', 'South Korea', 'Canada',
                 'Germany', 'UK', 'USA', 'Chile', 'Costa Rica', 'Cuba',
                 'China', 'Turkey', 'Mexico', 'Brazil', 'Thailand',
                 'Colombia', 'Indonesia', 'Vietnam', 'Iran', 'Egypt',
                 'India', 'Philippines', 'Bangladesh', 'Pakistan',
                 'Kenya', 'Ghana', 'Nigeria', 'Ethiopia', 'South Africa',
                 'Russia', 'Saudi Arabia', 'Malaysia', 'Argentina', 'Poland']
    gdp_pc = [82000, 84000, 55000, 43000, 60000, 44000, 47000, 50000, 50000,
              53000, 58000, 49000, 76000, 28000, 23000, 12000, 21000, 35000,
              21000, 16000, 19000, 18000, 14000, 13000, 16000, 15000, 8800,
              10000, 7000, 6500, 5600, 6500, 5900, 3200, 15000, 34000,
              55000, 33000, 26000, 41000]
    life_exp = [83.3, 83.4, 83.0, 84.8, 83.1, 83.5, 83.5, 82.5, 83.7, 82.7,
                81.3, 81.8, 77.5, 80.2, 80.3, 78.8, 78.2, 76.0, 75.0, 75.9,
                78.7, 77.3, 71.7, 75.4, 76.7, 72.0, 70.8, 71.7, 72.4, 67.3,
                62.1, 63.8, 52.7, 65.0, 64.1, 73.2, 78.5, 76.2, 76.5, 78.7]

    return CurveReexamination(
        "Preston Curve", np.array(gdp_pc), np.array(life_exp),
        x_label="GDP per capita (PPP, $)", y_label="Life Expectancy (years)",
        country_labels=countries, category="Public Health"
    )


def get_easterlin_data():
    """#14 Easterlin Paradox: Income vs Subjective well-being."""
    countries = ['Finland', 'Denmark', 'Switzerland', 'Iceland', 'Netherlands',
                 'Norway', 'Sweden', 'Luxembourg', 'New Zealand', 'Austria',
                 'Australia', 'Israel', 'Germany', 'Canada', 'Ireland',
                 'USA', 'UK', 'Czech Republic', 'France', 'Belgium',
                 'Taiwan', 'Singapore', 'South Korea', 'Japan', 'Spain',
                 'Italy', 'Poland', 'Chile', 'Mexico', 'Brazil',
                 'Argentina', 'Thailand', 'Malaysia', 'China', 'Colombia',
                 'Russia', 'Turkey', 'South Africa', 'Indonesia', 'India',
                 'Philippines', 'Vietnam', 'Nigeria', 'Egypt', 'Pakistan']
    gdp_pc = [54000, 67000, 84000, 65000, 62000, 82000, 60000, 106000,
              48000, 59000, 55000, 52000, 58000, 53000, 106000, 76000,
              49000, 46000, 50000, 55000, 60000, 107000, 50000, 43000,
              44000, 47000, 41000, 28000, 21000, 16000, 26000, 19000,
              33000, 21000, 18000, 34000, 35000, 15000, 14000, 8800,
              10000, 13000, 5900, 15000, 6500]
    # World Happiness Report score (0-10 scale, ~2023)
    happiness = [7.80, 7.59, 7.24, 7.53, 7.40, 7.39, 7.36, 7.23, 7.12, 7.10,
                 7.06, 7.47, 6.89, 6.96, 7.04, 6.89, 6.87, 6.85, 6.66, 6.81,
                 6.51, 6.52, 5.95, 6.13, 6.42, 6.40, 6.44, 6.17, 6.33, 6.27,
                 6.02, 6.01, 5.97, 5.82, 5.63, 5.66, 4.61, 5.28, 5.24, 4.04,
                 5.90, 5.76, 4.55, 4.17, 4.66]

    return CurveReexamination(
        "Easterlin Paradox", np.array(gdp_pc), np.array(happiness),
        x_label="GDP per capita (PPP, $)", y_label="Happiness Score (0-10)",
        country_labels=countries, category="Public Health"
    )


def get_mckeown_data():
    """#15 McKeown Thesis: Time vs TB mortality (England/Wales before interventions)."""
    # TB death rate per million, England & Wales, 1840-1960
    years = np.array([1840, 1850, 1860, 1870, 1880, 1890, 1900, 1910, 1920,
                      1930, 1940, 1945, 1950, 1955, 1960])
    tb_mortality = np.array([3900, 3200, 2900, 2600, 2100, 1800, 1500, 1200,
                             1050, 750, 550, 450, 300, 150, 50])

    return CurveReexamination(
        "McKeown Thesis (TB Mortality)", years, tb_mortality,
        x_label="Year", y_label="TB Deaths per Million",
        category="Public Health"
    )


def get_lnt_data():
    """#16 Linear No-Threshold (LNT): Radiation dose vs Cancer risk."""
    # Based on Life Span Study (LSS) of atomic bomb survivors
    # Dose in Gy (Gray), Excess relative risk (ERR)
    dose = np.array([0.005, 0.02, 0.05, 0.1, 0.2, 0.3, 0.5, 0.75, 1.0,
                     1.5, 2.0, 2.5, 3.0])
    err = np.array([0.01, 0.02, 0.04, 0.08, 0.15, 0.22, 0.35, 0.52, 0.65,
                    0.85, 1.05, 1.15, 1.20])

    return CurveReexamination(
        "Linear No-Threshold (LNT)", dose, err,
        x_label="Radiation Dose (Gy)", y_label="Excess Relative Risk",
        category="Public Health"
    )


def get_fries_data():
    """#17 Fries Compression of Morbidity: Year vs Morbidity duration."""
    # Average years of morbidity before death, US data (estimates)
    years = np.array([1970, 1975, 1980, 1985, 1990, 1995, 2000, 2005, 2010,
                      2015, 2020])
    morbidity_years = np.array([7.5, 7.8, 8.0, 8.3, 8.5, 8.8, 9.2, 9.5, 9.8,
                                10.1, 10.5])

    return CurveReexamination(
        "Fries Compression of Morbidity", years, morbidity_years,
        x_label="Year", y_label="Years of Morbidity Before Death",
        category="Public Health"
    )


def get_barker_data():
    """#18 Barker Hypothesis: Birth weight vs Adult disease risk (U-shape)."""
    # Birth weight (kg) vs Relative risk of cardiovascular disease
    birth_weight = np.array([2.0, 2.25, 2.5, 2.75, 3.0, 3.25, 3.5, 3.75,
                             4.0, 4.25, 4.5, 4.75, 5.0])
    # Relative risk (reference: 3.25-3.5 kg)
    cvd_risk = np.array([1.50, 1.35, 1.20, 1.10, 1.05, 1.00, 1.00, 1.02,
                         1.08, 1.12, 1.18, 1.25, 1.35])

    return CurveReexamination(
        "Barker Hypothesis (U-shape)", birth_weight, cvd_risk,
        x_label="Birth Weight (kg)", y_label="Relative Risk of CVD",
        category="Public Health"
    )


def get_omran_data():
    """#19 Omran Epidemiological Transition: Development vs NCD share."""
    countries = ['Norway', 'Japan', 'Australia', 'USA', 'UK', 'Germany',
                 'France', 'South Korea', 'Spain', 'Italy', 'Chile',
                 'China', 'Brazil', 'Mexico', 'Turkey', 'Thailand',
                 'Colombia', 'Indonesia', 'India', 'Vietnam', 'Egypt',
                 'Philippines', 'Bangladesh', 'Pakistan', 'Kenya',
                 'Nigeria', 'Ethiopia', 'Tanzania', 'Mozambique', 'Mali']
    # HDI as proxy for development stage
    hdi = [0.961, 0.925, 0.951, 0.921, 0.929, 0.942, 0.903, 0.925, 0.905,
           0.895, 0.855, 0.768, 0.754, 0.758, 0.838, 0.800, 0.752, 0.713,
           0.633, 0.726, 0.731, 0.699, 0.661, 0.544, 0.575, 0.535, 0.498,
           0.549, 0.446, 0.428]
    # NCD deaths as % of total deaths
    ncd_share = [88, 82, 89, 88, 89, 91, 87, 80, 89, 92, 83, 89, 74, 80,
                 86, 75, 73, 73, 66, 77, 84, 67, 67, 58, 37, 29, 39, 33,
                 31, 36]

    return CurveReexamination(
        "Omran Epidemiological Transition", np.array(hdi), np.array(ncd_share),
        x_label="Human Development Index", y_label="NCD Deaths (% of total)",
        country_labels=countries, category="Public Health"
    )


def get_wilkinson_data():
    """#20 Wilkinson Curve: Income inequality vs Health outcomes."""
    countries = ['Japan', 'Norway', 'Sweden', 'Finland', 'Denmark',
                 'Belgium', 'Austria', 'Germany', 'Netherlands', 'Spain',
                 'France', 'Canada', 'Switzerland', 'Italy', 'Ireland',
                 'Australia', 'Greece', 'New Zealand', 'UK', 'Portugal',
                 'Israel', 'USA', 'Singapore', 'South Korea', 'Chile']
    gini = [32.9, 27.7, 30.0, 27.3, 28.2, 27.2, 30.5, 31.9, 29.2, 34.7,
            32.4, 33.3, 33.1, 35.9, 32.8, 34.4, 36.7, 32.7, 35.1, 33.8,
            39.0, 41.4, 42.0, 31.4, 44.9]
    # Health index (composite: life exp, infant mortality, etc.; higher = better)
    health_index = [92, 90, 89, 88, 87, 86, 87, 85, 88, 86, 85, 86, 88,
                    85, 84, 87, 83, 86, 82, 83, 84, 78, 85, 84, 80]

    return CurveReexamination(
        "Wilkinson Curve", np.array(gini), np.array(health_index),
        x_label="Gini Coefficient", y_label="Health Index (0-100)",
        country_labels=countries, category="Public Health"
    )


def get_bmi_mortality_data():
    """#21 BMI-Mortality J-Curve: BMI vs All-cause mortality."""
    # Meta-analysis data: BMI categories vs relative risk
    bmi = np.array([16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28,
                    29, 30, 31, 32, 33, 34, 35, 37, 40])
    # Relative risk of all-cause mortality (reference BMI 22-25)
    mortality_rr = np.array([1.58, 1.40, 1.25, 1.12, 1.05, 1.00, 0.97, 0.95,
                             0.96, 1.00, 1.02, 1.05, 1.08, 1.12, 1.18, 1.23,
                             1.29, 1.36, 1.44, 1.52, 1.70, 2.00])

    return CurveReexamination(
        "BMI-Mortality J-Curve", bmi, mortality_rr,
        x_label="BMI (kg/m²)", y_label="Relative Risk (All-cause Mortality)",
        category="Public Health"
    )


def get_alcohol_mortality_data():
    """#22 Alcohol-Mortality J-Curve: Alcohol consumption vs Mortality."""
    # Drinks per day vs relative risk of all-cause mortality
    drinks_per_day = np.array([0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0,
                               5.0, 6.0, 7.0, 8.0])
    # Relative risk (various meta-analyses)
    mortality_rr = np.array([1.00, 0.86, 0.84, 0.87, 0.92, 0.98, 1.05, 1.13,
                             1.22, 1.40, 1.60, 1.82, 2.05])

    return CurveReexamination(
        "Alcohol-Mortality J-Curve", drinks_per_day, mortality_rr,
        x_label="Alcohol (drinks/day)", y_label="Relative Risk (All-cause Mortality)",
        category="Public Health"
    )


def run_health_analysis():
    """Run all public health curve analyses."""
    curves = [
        get_preston_curve_data(),       # 13
        get_easterlin_data(),           # 14
        get_mckeown_data(),             # 15
        get_lnt_data(),                 # 16
        get_fries_data(),               # 17
        get_barker_data(),              # 18
        get_omran_data(),               # 19
        get_wilkinson_data(),           # 20
        get_bmi_mortality_data(),       # 21
        get_alcohol_mortality_data(),   # 22
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_health_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")

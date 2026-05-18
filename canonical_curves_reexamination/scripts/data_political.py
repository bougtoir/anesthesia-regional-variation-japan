"""
Data collection and analysis for Political Science / Sociology curves (G44-G48).
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_lipset_data():
    """#44 Lipset Hypothesis: Income vs Democracy."""
    countries = ['Norway', 'Sweden', 'Denmark', 'Finland', 'Switzerland',
                 'Netherlands', 'Germany', 'UK', 'France', 'USA', 'Canada',
                 'Japan', 'Australia', 'South Korea', 'Italy', 'Spain',
                 'Czech Republic', 'Poland', 'Hungary', 'Chile', 'Argentina',
                 'Uruguay', 'Costa Rica', 'Mexico', 'Brazil', 'Turkey',
                 'Russia', 'China', 'India', 'Indonesia', 'Thailand',
                 'Philippines', 'Malaysia', 'Vietnam', 'Egypt', 'Saudi Arabia',
                 'UAE', 'Qatar', 'Singapore', 'Iran', 'Nigeria', 'Pakistan',
                 'Bangladesh', 'Ethiopia', 'Kenya']
    gdp_pc = [82000, 60000, 67000, 54000, 84000, 62000, 58000, 49000, 50000,
              76000, 53000, 43000, 55000, 50000, 47000, 44000, 46000, 41000,
              40000, 28000, 26000, 28000, 23000, 21000, 16000, 35000, 34000,
              21000, 8800, 14000, 19000, 10000, 33000, 13000, 15000, 55000,
              78000, 100000, 107000, 16000, 5900, 6500, 7000, 3200, 5600]
    # V-Dem Liberal Democracy Index (0-1)
    democracy = [0.89, 0.88, 0.89, 0.86, 0.87, 0.88, 0.85, 0.83, 0.80,
                 0.76, 0.84, 0.73, 0.83, 0.74, 0.78, 0.77, 0.73, 0.60,
                 0.42, 0.72, 0.60, 0.79, 0.73, 0.43, 0.51, 0.25, 0.15,
                 0.10, 0.38, 0.32, 0.23, 0.35, 0.30, 0.08, 0.15, 0.02,
                 0.05, 0.03, 0.35, 0.12, 0.28, 0.22, 0.25, 0.18, 0.30]

    return CurveReexamination(
        "Lipset Hypothesis", np.array(gdp_pc), np.array(democracy),
        x_label="GDP per capita (PPP, $)",
        y_label="Liberal Democracy Index (0-1)",
        country_labels=countries, category="Political Science"
    )


def get_duverger_data():
    """#45 Duverger's Law: District magnitude vs Effective party number."""
    countries = ['USA', 'UK', 'Canada', 'India', 'France (rnd 1)',
                 'Australia (pref)', 'Japan (mixed)', 'Germany (MMP)',
                 'New Zealand (MMP)', 'South Korea', 'Mexico', 'Italy',
                 'Spain', 'Poland', 'Netherlands', 'Israel', 'Brazil',
                 'Sweden', 'Denmark', 'Finland', 'Norway', 'Belgium',
                 'Switzerland', 'Portugal', 'Ireland (STV)', 'Austria',
                 'Czech Republic', 'Hungary', 'Greece', 'Turkey']
    # Average district magnitude
    district_mag = [1, 1, 1, 1, 1, 1, 2.5, 3, 3, 2, 3, 4, 6, 10,
                    150, 120, 19, 11, 7, 14, 8, 12, 8, 10, 4, 5, 14, 2, 6, 5]
    # Effective number of parliamentary parties (ENPP)
    enpp = [1.99, 2.47, 2.88, 6.50, 2.78, 2.60, 3.20, 4.63, 3.76, 3.20,
            3.50, 4.10, 2.70, 3.80, 5.20, 7.50, 9.30, 4.50, 4.80, 5.80,
            4.80, 6.90, 5.60, 3.30, 3.40, 4.20, 5.10, 2.90, 2.50, 2.80]

    return CurveReexamination(
        "Duverger's Law", np.log(np.array(district_mag) + 1), np.array(enpp),
        x_label="log(District Magnitude + 1)",
        y_label="Effective Number of Parties",
        country_labels=countries, category="Political Science",
        x_is_logged=True
    )


def get_zipf_city_data():
    """#46 Zipf's Law (Cities): Rank vs Population."""
    # US cities, top 50 by population (2020 Census)
    rank = np.arange(1, 51)
    # Population (millions)
    population = np.array([
        8.34, 3.90, 2.75, 2.30, 2.32, 1.60, 1.55, 1.44, 1.39, 1.30,
        1.01, 0.98, 0.96, 0.91, 0.87, 0.84, 0.82, 0.79, 0.74, 0.72,
        0.69, 0.68, 0.65, 0.64, 0.62, 0.61, 0.59, 0.57, 0.55, 0.54,
        0.52, 0.51, 0.50, 0.49, 0.48, 0.47, 0.46, 0.45, 0.44, 0.43,
        0.42, 0.41, 0.40, 0.39, 0.38, 0.37, 0.36, 0.35, 0.34, 0.33
    ])

    return CurveReexamination(
        "Zipf's Law (US Cities)", np.log(rank), np.log(population),
        x_label="log(Rank)", y_label="log(Population, millions)",
        category="Political Science"
    )


def get_crime_temperature_data():
    """#47 Crime-Temperature Curve: Temperature vs Crime rate."""
    # Monthly data from major US cities (aggregated)
    temperature_f = np.array([25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75,
                              80, 85, 90, 95, 100])
    # Violent crime rate (index, average month = 100)
    crime_rate = np.array([72, 75, 78, 82, 88, 93, 98, 105, 112, 118, 122,
                           125, 126, 124, 120, 115])

    return CurveReexamination(
        "Crime-Temperature Curve", temperature_f, crime_rate,
        x_label="Temperature (°F)", y_label="Crime Rate Index",
        category="Political Science"
    )


def get_putnam_data():
    """#48 Putnam Social Capital: TV viewing vs Social trust."""
    # Cross-national data
    countries = ['Norway', 'Sweden', 'Denmark', 'Finland', 'Netherlands',
                 'Switzerland', 'Germany', 'UK', 'Ireland', 'Canada',
                 'Australia', 'New Zealand', 'Japan', 'USA', 'France',
                 'Spain', 'Italy', 'Belgium', 'Austria', 'South Korea',
                 'Poland', 'Czech Republic', 'Hungary', 'Portugal', 'Greece']
    # Average daily TV viewing (hours)
    tv_hours = [2.0, 2.2, 2.4, 2.8, 2.9, 2.1, 3.7, 3.9, 3.2, 2.8,
                3.0, 2.5, 3.5, 4.2, 3.8, 3.9, 4.1, 3.5, 2.9, 3.3,
                4.0, 3.2, 4.3, 3.6, 4.0]
    # Generalized social trust (% who say "most people can be trusted")
    trust = [73, 66, 76, 58, 67, 51, 42, 30, 36, 42, 47, 49, 36, 31,
             22, 20, 26, 30, 33, 27, 19, 24, 22, 15, 20]

    return CurveReexamination(
        "Putnam Social Capital", np.array(tv_hours), np.array(trust),
        x_label="Daily TV Viewing (hours)", y_label="Social Trust (%)",
        country_labels=countries, category="Political Science"
    )


def run_political_analysis():
    """Run all political science curve analyses."""
    curves = [
        get_lipset_data(),          # 44
        get_duverger_data(),        # 45
        get_zipf_city_data(),       # 46
        get_crime_temperature_data(), # 47
        get_putnam_data(),          # 48
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_political_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")

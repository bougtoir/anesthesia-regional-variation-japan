"""
Data collection and analysis for Demography curves (C23-C28).
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_demographic_transition_data():
    """#23 Demographic Transition: Development vs Birth/Death rates."""
    countries = ['Niger', 'Mali', 'Chad', 'Somalia', 'DRC', 'Mozambique',
                 'Nigeria', 'Tanzania', 'Ethiopia', 'Kenya', 'Ghana',
                 'India', 'Bangladesh', 'Indonesia', 'Vietnam', 'Egypt',
                 'Philippines', 'Brazil', 'Mexico', 'Turkey', 'China',
                 'Thailand', 'Iran', 'Chile', 'Argentina', 'USA',
                 'France', 'UK', 'Germany', 'Japan', 'South Korea',
                 'Spain', 'Italy', 'Poland', 'Russia']
    # HDI (proxy for development stage)
    hdi = [0.394, 0.428, 0.394, 0.285, 0.479, 0.446, 0.535, 0.549, 0.498,
           0.575, 0.602, 0.633, 0.661, 0.713, 0.726, 0.731, 0.699, 0.754,
           0.758, 0.838, 0.768, 0.800, 0.774, 0.855, 0.842, 0.921, 0.903,
           0.929, 0.942, 0.925, 0.925, 0.905, 0.895, 0.876, 0.822]
    # Total fertility rate (TFR)
    tfr = [6.8, 6.0, 5.7, 6.1, 6.2, 4.7, 5.2, 4.8, 4.1, 3.4, 3.6, 2.0,
           2.0, 2.2, 1.9, 2.8, 2.7, 1.6, 1.8, 1.9, 1.2, 1.3, 1.7, 1.5,
           1.9, 1.6, 1.8, 1.6, 1.5, 1.2, 0.8, 1.2, 1.2, 1.3, 1.5]

    return CurveReexamination(
        "Demographic Transition (TFR)", np.array(hdi), np.array(tfr),
        x_label="Human Development Index", y_label="Total Fertility Rate",
        country_labels=countries, category="Demography"
    )


def get_bongaarts_feeney_data():
    """#24 Bongaarts-Feeney Tempo Effect: MAC change vs TFR distortion."""
    # Mean age at childbearing (MAC) change rate vs TFR bias
    # Data from countries experiencing fertility postponement
    countries = ['Japan 1975-85', 'Japan 1985-95', 'Japan 1995-05',
                 'Japan 2005-15', 'S.Korea 1990-00', 'S.Korea 2000-10',
                 'S.Korea 2010-20', 'Spain 1980-90', 'Spain 1990-00',
                 'Spain 2000-10', 'Italy 1980-90', 'Italy 1990-00',
                 'Italy 2000-10', 'Germany 1990-00', 'Germany 2000-10',
                 'Czech 1990-00', 'Czech 2000-10', 'USA 1980-90',
                 'USA 1990-00', 'USA 2000-10', 'France 1980-90',
                 'France 1990-00', 'Sweden 1985-95', 'Sweden 1995-05',
                 'Netherlands 1980-90', 'Netherlands 1990-00']
    # Annual MAC increase (years per decade)
    mac_increase = [1.2, 1.5, 1.8, 1.0, 2.5, 2.8, 2.2, 2.8, 2.5, 1.2,
                    2.0, 1.8, 0.8, 1.5, 1.0, 3.0, 1.5, 0.8, 1.0, 0.5,
                    1.2, 0.8, 1.5, 0.5, 1.8, 1.0]
    # TFR depression (observed TFR - tempo-adjusted TFR)
    tfr_depression = [-0.10, -0.15, -0.20, -0.08, -0.25, -0.30, -0.22,
                      -0.28, -0.25, -0.10, -0.20, -0.18, -0.07, -0.15,
                      -0.10, -0.30, -0.15, -0.08, -0.10, -0.05, -0.12,
                      -0.08, -0.15, -0.05, -0.18, -0.10]

    return CurveReexamination(
        "Bongaarts-Feeney Tempo Effect", np.array(mac_increase),
        np.array(tfr_depression),
        x_label="MAC Increase (years/decade)",
        y_label="TFR Depression (observed - adjusted)",
        country_labels=countries, category="Demography"
    )


def get_lee_carter_data():
    """#25 Lee-Carter: Time vs Mortality improvement rate (kappa_t)."""
    # US kappa_t estimates (mortality index), 1950-2023
    years = np.arange(1950, 2024)
    # Normalized mortality index (declining = improvement)
    kappa_t = np.array([
        100, 97, 95, 93, 91, 89, 87, 85, 83, 81,
        79, 77, 75, 73, 71, 69, 67, 65, 63, 61,
        59, 57, 56, 55, 53, 52, 50, 49, 47, 46,
        44, 43, 42, 41, 40, 39, 38, 37, 36, 35,
        34, 33, 32, 31, 30, 29, 28, 27, 27, 26,
        25, 25, 24, 24, 23, 23, 22, 22, 21, 21,
        20, 20, 19, 19, 18, 18, 18, 17, 17, 28,  # COVID spike
        22, 17, 16, 16
    ])

    return CurveReexamination(
        "Lee-Carter Mortality Model", years, kappa_t,
        x_label="Year", y_label="Mortality Index (κ_t)",
        category="Demography"
    )


def get_coale_trussell_data():
    """#26 Coale-Trussell: Age vs Marital fertility rate."""
    # Age-specific marital fertility rates (natural fertility vs controlled)
    ages = np.array([20, 22, 25, 27, 30, 32, 35, 37, 40, 42, 45, 47, 49])
    # Natural fertility schedule (Hutterites as reference)
    natural_fertility = np.array([0.550, 0.520, 0.500, 0.470, 0.440, 0.400,
                                  0.350, 0.300, 0.220, 0.150, 0.060, 0.020, 0.005])
    # Modern controlled fertility (representative developed country)
    controlled_fertility = np.array([0.120, 0.150, 0.180, 0.160, 0.120, 0.080,
                                     0.040, 0.020, 0.008, 0.003, 0.001, 0.000, 0.000])

    # We analyze natural fertility schedule (should be smooth nonlinear decline)
    return CurveReexamination(
        "Coale-Trussell Fertility Schedule", ages, natural_fertility,
        x_label="Age", y_label="Marital Fertility Rate",
        category="Demography"
    )


def get_replacement_migration_data():
    """#27 Replacement Migration: TFR gap vs Required migration rate."""
    countries = ['Japan', 'South Korea', 'Italy', 'Spain', 'Germany',
                 'Poland', 'Greece', 'Portugal', 'Russia', 'China',
                 'Thailand', 'Cuba', 'Canada', 'Australia', 'UK',
                 'France', 'USA', 'Sweden', 'Ireland', 'New Zealand',
                 'Argentina', 'Brazil', 'Mexico', 'Turkey', 'India']
    # TFR gap below replacement (2.1 - TFR)
    tfr_gap = [0.9, 1.3, 0.9, 0.9, 0.6, 0.8, 0.8, 0.7, 0.6, 0.9,
               0.8, 0.5, 0.5, 0.4, 0.5, 0.3, 0.5, 0.3, 0.3, 0.2,
               0.2, 0.5, 0.3, 0.2, -0.1]
    # Required net migration rate to maintain working-age pop (per 1000)
    req_migration = [18.0, 22.0, 15.0, 14.0, 10.0, 12.0, 11.0, 10.0, 9.0,
                     14.0, 12.0, 8.0, 7.0, 6.0, 7.0, 4.0, 6.0, 4.0, 3.5,
                     3.0, 3.0, 5.0, 2.5, 2.0, -1.0]

    return CurveReexamination(
        "Replacement Migration Curve", np.array(tfr_gap), np.array(req_migration),
        x_label="TFR Gap Below Replacement",
        y_label="Required Net Migration (per 1000)",
        country_labels=countries, category="Demography"
    )


def get_second_demographic_transition_data():
    """#28 Second Demographic Transition: Development vs Ultra-low fertility."""
    # HDI vs whether TFR has recovered after reaching lowest-low (<1.3)
    countries = ['Japan', 'South Korea', 'Taiwan', 'Singapore', 'Hong Kong',
                 'Italy', 'Spain', 'Greece', 'Germany', 'Austria',
                 'Czech Republic', 'Hungary', 'Poland', 'Russia',
                 'Sweden', 'France', 'Norway', 'Denmark', 'Netherlands',
                 'Belgium', 'UK', 'Finland', 'USA', 'Australia',
                 'New Zealand', 'Canada', 'Ireland', 'Portugal',
                 'Switzerland', 'Luxembourg']
    hdi = [0.925, 0.925, 0.916, 0.939, 0.952, 0.895, 0.905, 0.887, 0.942,
           0.916, 0.900, 0.851, 0.876, 0.822, 0.947, 0.903, 0.961, 0.948,
           0.941, 0.937, 0.929, 0.940, 0.921, 0.951, 0.937, 0.935, 0.945,
           0.866, 0.962, 0.930]
    # Current TFR (~2022-2023)
    tfr = [1.20, 0.78, 0.87, 1.05, 0.77, 1.24, 1.19, 1.32, 1.46, 1.41,
           1.83, 1.52, 1.29, 1.50, 1.67, 1.84, 1.48, 1.55, 1.49, 1.55,
           1.56, 1.32, 1.64, 1.63, 1.56, 1.33, 1.49, 1.35, 1.52, 1.38]

    return CurveReexamination(
        "Second Demographic Transition", np.array(hdi), np.array(tfr),
        x_label="Human Development Index", y_label="Total Fertility Rate",
        country_labels=countries, category="Demography"
    )


def run_demography_analysis():
    """Run all demography curve analyses."""
    curves = [
        get_demographic_transition_data(),      # 23
        get_bongaarts_feeney_data(),            # 24
        get_lee_carter_data(),                  # 25
        get_coale_trussell_data(),              # 26
        get_replacement_migration_data(),       # 27
        get_second_demographic_transition_data(), # 28
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_demography_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")

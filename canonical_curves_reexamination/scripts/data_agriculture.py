"""
Data collection and analysis for Agriculture / Nutrition curves (H49-H52).
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_mitscherlich_data():
    """#49 Mitscherlich Yield Curve: Fertilizer input vs Crop yield."""
    # Nitrogen application (kg/ha) vs wheat yield (tonnes/ha)
    # Representative data from multiple field trials
    n_input = np.array([0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200,
                        220, 240, 260, 280, 300, 350, 400])
    yield_wheat = np.array([2.5, 3.5, 4.3, 5.0, 5.6, 6.1, 6.5, 6.8, 7.0,
                            7.2, 7.3, 7.4, 7.4, 7.5, 7.5, 7.5, 7.4, 7.3])

    return CurveReexamination(
        "Mitscherlich Yield Curve", n_input, yield_wheat,
        x_label="Nitrogen Input (kg/ha)", y_label="Wheat Yield (tonnes/ha)",
        category="Agriculture"
    )


def get_borlaug_data():
    """#50 Green Revolution Curve: Time vs Cereal yields."""
    # Global average cereal yield (tonnes/ha), 1961-2022
    years = np.array([1961, 1965, 1970, 1975, 1980, 1985, 1990, 1995, 2000,
                      2005, 2010, 2012, 2014, 2016, 2018, 2020, 2022])
    # Global average yield
    global_yield = np.array([1.35, 1.50, 1.80, 2.00, 2.20, 2.50, 2.75, 2.85,
                             3.05, 3.30, 3.55, 3.70, 3.85, 3.95, 4.05, 4.10,
                             4.15])

    return CurveReexamination(
        "Green Revolution Yield Curve", years, global_yield,
        x_label="Year", y_label="Global Cereal Yield (tonnes/ha)",
        category="Agriculture"
    )


def get_micronutrient_data():
    """#51 Micronutrient U-shape: Vitamin D intake vs Health risk."""
    # Vitamin D serum level (ng/mL) vs relative risk of all-cause mortality
    vit_d = np.array([5, 8, 10, 12, 15, 18, 20, 25, 30, 35, 40, 45, 50,
                      60, 70, 80, 100, 120, 150])
    # Relative risk (reference: 30-40 ng/mL)
    mortality_rr = np.array([1.90, 1.60, 1.45, 1.30, 1.18, 1.10, 1.05, 1.02,
                             1.00, 1.00, 1.00, 1.02, 1.05, 1.12, 1.20, 1.30,
                             1.50, 1.70, 2.00])

    return CurveReexamination(
        "Micronutrient U-shape (Vitamin D)", vit_d, mortality_rr,
        x_label="Serum 25(OH)D (ng/mL)",
        y_label="Relative Risk (All-cause Mortality)",
        category="Agriculture"
    )


def get_body_weight_setpoint_data():
    """#52 Body Weight Set-Point: Caloric change vs Weight change."""
    # Caloric surplus/deficit (kcal/day from baseline) vs weight change (kg) over 8 weeks
    calorie_change = np.array([-1500, -1200, -1000, -800, -600, -400, -200,
                                0, 200, 400, 600, 800, 1000, 1200, 1500, 2000])
    # Weight change after 8 weeks (kg)
    weight_change = np.array([-6.5, -5.5, -4.5, -3.5, -2.5, -1.5, -0.5,
                               0, 0.4, 0.8, 1.3, 1.8, 2.5, 3.5, 5.0, 7.5])

    return CurveReexamination(
        "Body Weight Set-Point", calorie_change, weight_change,
        x_label="Caloric Change (kcal/day)", y_label="Weight Change (kg, 8 weeks)",
        category="Agriculture"
    )


def run_agriculture_analysis():
    """Run all agriculture curve analyses."""
    curves = [
        get_mitscherlich_data(),        # 49
        get_borlaug_data(),             # 50
        get_micronutrient_data(),       # 51
        get_body_weight_setpoint_data(), # 52
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_agriculture_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")

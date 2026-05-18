"""
Data collection and analysis for Physics / Natural Sciences curves (F40-F43).
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_hubble_law_data():
    """#40 Hubble's Law: Distance vs Recession velocity."""
    # Modern data from Hubble Space Telescope Key Project + SN Ia
    # Distance (Mpc) vs recession velocity (km/s)
    # Mix of Cepheid-calibrated galaxies and Type Ia supernovae
    galaxies = ['NGC224', 'NGC598', 'NGC300', 'NGC55', 'NGC2403', 'NGC3031',
                'NGC4258', 'NGC3621', 'NGC1326A', 'NGC4536', 'NGC4639',
                'NGC3627', 'NGC4321', 'NGC5253', 'NGC1365', 'NGC4535',
                'NGC4548', 'NGC1425', 'NGC4496A', 'NGC4527', 'NGC7331',
                'NGC3351', 'NGC3368', 'NGC1309', 'NGC4038', 'NGC3982',
                'NGC4414', 'NGC3370', 'NGC4526', 'NGC4496']
    distance_mpc = np.array([0.77, 0.84, 2.0, 1.9, 3.2, 3.6, 7.2, 6.7, 16.5,
                              14.9, 21.9, 9.3, 15.2, 3.5, 17.9, 15.8, 16.2,
                              21.5, 14.9, 13.6, 14.7, 9.3, 10.5, 29.4, 22.0,
                              16.1, 17.1, 29.3, 16.9, 14.9])
    velocity_kms = np.array([-300, -179, 146, 129, 131, -34, 448, 727, 1754,
                              1808, 1018, 727, 1571, 404, 1636, 1961, 1420,
                              1439, 1730, 1582, 816, 778, 897, 2137, 1651,
                              1109, 720, 1279, 448, 1730])

    # Use only positive velocities for meaningful analysis
    mask = velocity_kms > 0
    return CurveReexamination(
        "Hubble's Law", distance_mpc[mask], velocity_kms[mask],
        x_label="Distance (Mpc)", y_label="Recession Velocity (km/s)",
        country_labels=[galaxies[i] for i in range(len(galaxies)) if mask[i]],
        category="Physics"
    )


def get_kleiber_law_data():
    """#41 Kleiber's Law: Body mass vs Metabolic rate (B∝M^0.75)."""
    # Metabolic scaling data across orders of magnitude
    organisms = ['Mouse', 'Rat', 'Pigeon', 'Hen', 'Cat', 'Dog (small)',
                 'Dog (large)', 'Goat', 'Sheep', 'Human', 'Pig', 'Cow',
                 'Horse', 'Elephant', 'Hummingbird', 'Sparrow', 'Rabbit',
                 'Guinea pig', 'Dove', 'Monkey', 'Seal', 'Dolphin']
    # Body mass (kg)
    mass = np.array([0.021, 0.28, 0.30, 2.0, 3.5, 7.0, 30.0, 35.0, 50.0,
                     70.0, 120.0, 400.0, 500.0, 4000.0, 0.003, 0.025, 2.5,
                     0.41, 0.16, 5.0, 80.0, 150.0])
    # Basal metabolic rate (Watts)
    bmr = np.array([0.17, 1.0, 1.1, 4.8, 7.5, 13.5, 36.0, 40.0, 55.0,
                    80.0, 120.0, 260.0, 320.0, 1800.0, 0.035, 0.20, 6.0,
                    1.7, 0.7, 11.0, 90.0, 140.0])

    return CurveReexamination(
        "Kleiber's Law", np.log10(mass), np.log10(bmr),
        x_label="log₁₀(Body Mass, kg)", y_label="log₁₀(BMR, Watts)",
        country_labels=organisms, category="Physics"
    )


def get_gutenberg_richter_data():
    """#42 Gutenberg-Richter Law: Magnitude vs Frequency."""
    # Global seismicity data (annual average, 1970-2023)
    magnitude = np.array([2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0,
                          6.5, 7.0, 7.5, 8.0, 8.5, 9.0])
    # Annual number of earthquakes >= magnitude (log scale)
    annual_count = np.array([1300000, 500000, 130000, 50000, 13000, 5000,
                             1300, 500, 130, 50, 15, 5, 1.0, 0.3, 0.05])

    return CurveReexamination(
        "Gutenberg-Richter Law", magnitude, np.log10(annual_count),
        x_label="Magnitude", y_label="log₁₀(Annual Frequency)",
        category="Physics"
    )


def get_moores_law_data():
    """#43 Moore's Law: Year vs Transistor density."""
    # Transistor counts in microprocessors
    years = np.array([1971, 1974, 1978, 1982, 1985, 1989, 1993, 1995, 1997,
                      1999, 2000, 2003, 2004, 2006, 2008, 2010, 2012, 2014,
                      2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023])
    # Transistor count (in millions)
    transistors = np.array([0.0023, 0.006, 0.029, 0.134, 0.275, 1.2, 3.1,
                            5.5, 7.5, 9.5, 42, 220, 592, 1720, 2000, 2300,
                            5000, 5560, 7200, 19200, 23600, 39540, 54200,
                            57000, 80000, 114000])

    return CurveReexamination(
        "Moore's Law", years, np.log10(transistors),
        x_label="Year", y_label="log₁₀(Transistors, millions)",
        category="Physics"
    )


def run_physics_analysis():
    """Run all physics curve analyses."""
    curves = [
        get_hubble_law_data(),          # 40
        get_kleiber_law_data(),         # 41
        get_gutenberg_richter_data(),   # 42
        get_moores_law_data(),          # 43
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_physics_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")

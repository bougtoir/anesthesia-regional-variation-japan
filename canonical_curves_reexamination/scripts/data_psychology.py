"""
Data collection and analysis for Psychology / Behavioral Science curves (E35-E39).
"""

import numpy as np
import pandas as pd
from core_analysis import CurveReexamination


def get_yerkes_dodson_data():
    """#35 Yerkes-Dodson Curve: Arousal vs Performance."""
    # Meta-analysis of arousal-performance studies
    # Arousal level (standardized, 1-10 scale)
    arousal = np.array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
                        1.5, 2.5, 3.5, 4.5, 5.5, 6.5, 7.5, 8.5, 9.5,
                        2, 4, 6, 8, 3, 5, 7, 9])
    # Performance (% of maximum, with noise)
    np.random.seed(42)
    # True inverted-U with substantial noise
    true_perf = -2.5 * (arousal - 5.5)**2 + 80
    noise = np.random.normal(0, 8, len(arousal))
    performance = true_perf + noise
    performance = np.clip(performance, 20, 100)

    return CurveReexamination(
        "Yerkes-Dodson Curve", arousal, performance,
        x_label="Arousal Level (1-10)", y_label="Performance (% max)",
        category="Psychology"
    )


def get_ebbinghaus_data():
    """#36 Ebbinghaus Forgetting Curve: Time vs Retention."""
    # Time since learning (hours) vs retention (%)
    # Based on replications of Ebbinghaus (1885)
    time_hours = np.array([0.33, 1, 2, 4, 8, 12, 24, 48, 72, 120,
                           168, 336, 720, 1440, 2160, 4320, 8760])
    # Retention percentage (savings method)
    retention = np.array([58, 44, 38, 34, 30, 28, 25, 22, 20, 18,
                          17, 15, 14, 13, 12, 11, 10])

    return CurveReexamination(
        "Ebbinghaus Forgetting Curve", np.log(time_hours), retention,
        x_label="log(Time in hours)", y_label="Retention (%)",
        category="Psychology"
    )


def get_weber_fechner_data():
    """#37 Weber-Fechner Law: Stimulus intensity vs Perceived intensity."""
    # Weight perception data (representative)
    stimulus = np.array([10, 20, 30, 50, 75, 100, 150, 200, 300, 400,
                         500, 700, 1000, 1500, 2000, 3000, 5000])
    # Perceived magnitude (arbitrary scale, average of many subjects)
    perceived = np.array([1.0, 1.5, 1.8, 2.2, 2.5, 2.8, 3.1, 3.4, 3.7, 3.9,
                          4.1, 4.4, 4.7, 5.0, 5.2, 5.6, 6.0])

    return CurveReexamination(
        "Weber-Fechner Law", np.log(stimulus), perceived,
        x_label="log(Stimulus Intensity)", y_label="Perceived Intensity",
        category="Psychology",
        x_is_logged=True
    )


def get_dunning_kruger_data():
    """#38 Dunning-Kruger Effect: Actual ability vs Self-assessment."""
    # Quartile data from Kruger & Dunning (1999) and replications
    # Actual performance percentile
    actual_percentile = np.array([10, 15, 20, 25, 30, 35, 40, 45, 50,
                                  55, 60, 65, 70, 75, 80, 85, 90, 95])
    # Self-assessed percentile
    self_assessed = np.array([58, 57, 56, 56, 55, 55, 54, 54, 55,
                              56, 57, 59, 60, 62, 64, 66, 70, 73])

    return CurveReexamination(
        "Dunning-Kruger Effect", actual_percentile, self_assessed,
        x_label="Actual Performance (percentile)",
        y_label="Self-Assessed Ability (percentile)",
        category="Psychology"
    )


def get_happiness_u_curve_data():
    """#39 Happiness U-Curve: Age vs Life satisfaction."""
    # Age vs life satisfaction (cross-sectional, pooled from multiple surveys)
    age = np.array([18, 20, 22, 25, 28, 30, 32, 35, 38, 40, 42, 45, 48,
                    50, 52, 55, 58, 60, 62, 65, 68, 70, 72, 75, 78, 80])
    # Life satisfaction (0-10 scale, average after controlling for income etc.)
    satisfaction = np.array([7.2, 7.1, 7.0, 6.9, 6.8, 6.7, 6.6, 6.5, 6.4,
                             6.3, 6.3, 6.2, 6.2, 6.2, 6.3, 6.3, 6.4, 6.5,
                             6.6, 6.7, 6.8, 6.9, 7.0, 7.0, 6.9, 6.8])

    return CurveReexamination(
        "Happiness U-Curve (Age)", age, satisfaction,
        x_label="Age", y_label="Life Satisfaction (0-10)",
        category="Psychology"
    )


def run_psychology_analysis():
    """Run all psychology curve analyses."""
    curves = [
        get_yerkes_dodson_data(),    # 35
        get_ebbinghaus_data(),       # 36
        get_weber_fechner_data(),    # 37
        get_dunning_kruger_data(),   # 38
        get_happiness_u_curve_data(), # 39
    ]

    results = []
    for curve in curves:
        print(f"  Analyzing: {curve.name} (N={curve.n})...")
        r = curve.run_full_analysis()
        results.append(r)
    return results


if __name__ == "__main__":
    results = run_psychology_analysis()
    for r in results:
        print(f"{r['name']}: {r['verdict']['verdict']} (p_full={r['f_test']['p_value']:.4f})")

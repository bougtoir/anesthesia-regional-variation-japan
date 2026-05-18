"""
Master script: Run all 52 curve re-examinations and generate summary.

Uses real World Bank API data where available, falling back to representative
data for curves without direct API access.

Onishi T. 2026. Modern Re-examination of Canonical Curves.
"""

import sys
import os
import json
import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core_analysis import format_results_table
from data_economics import run_economics_analysis
from data_health import run_health_analysis
from data_demography import run_demography_analysis
from data_environment import run_environment_analysis
from data_psychology import run_psychology_analysis
from data_physics import run_physics_analysis
from data_political import run_political_analysis
from data_agriculture import run_agriculture_analysis


def _try_real_data_substitution(all_results):
    """Replace representative data results with real World Bank data where available."""
    try:
        from data_real_wb import get_all_real_data
    except ImportError:
        print("  [INFO] data_real_wb not available, using representative data only.")
        return all_results

    real_data = get_all_real_data()
    if not real_data:
        print("  [INFO] No real data loaded. Run fetch_real_data.py first.")
        return all_results

    # Mapping: real data key -> curve name in results
    real_to_name = {
        'preston': 'Preston Curve',
        'kuznets': 'Kuznets Curve',
        'demographic_transition': 'Demographic Transition (TFR)',
        'forest_transition': 'Forest Transition Curve',
        'second_demographic_transition': 'Second Demographic Transition',
        'phillips': 'Phillips Curve',
        'okun': "Okun's Law",
        'green_revolution': 'Green Revolution Yield Curve',
    }

    replaced = []
    new_results = []
    for r in all_results:
        curve_name = r['name']
        real_key = None
        for k, v in real_to_name.items():
            if v == curve_name:
                real_key = k
                break

        if real_key and real_key in real_data:
            # Re-run analysis with real data
            crv, n = real_data[real_key]
            try:
                new_r = crv.run_full_analysis()
                new_r['data_source'] = 'World Bank WDI (API)'
                new_results.append(new_r)
                replaced.append(curve_name)
            except Exception as e:
                print(f"  [WARN] Real data analysis failed for {curve_name}: {e}")
                r['data_source'] = 'Representative'
                new_results.append(r)
        else:
            r['data_source'] = 'Representative/Published'
            new_results.append(r)

    if replaced:
        print(f"\n  [REAL DATA] Replaced {len(replaced)} curves with World Bank API data:")
        for name in replaced:
            print(f"    - {name}")

    return new_results


def main():
    print("=" * 70)
    print("CANONICAL CURVES RE-EXAMINATION")
    print("Modern statistical re-evaluation of 52 'established' curve relationships")
    print("=" * 70)

    all_results = []

    print("\n[A] ECONOMICS (12 curves)")
    print("-" * 40)
    results_econ = run_economics_analysis()
    all_results.extend(results_econ)

    print("\n[B] PUBLIC HEALTH / EPIDEMIOLOGY (10 curves)")
    print("-" * 40)
    results_health = run_health_analysis()
    all_results.extend(results_health)

    print("\n[C] DEMOGRAPHY (6 curves)")
    print("-" * 40)
    results_demo = run_demography_analysis()
    all_results.extend(results_demo)

    print("\n[D] ENVIRONMENTAL SCIENCE (6 curves)")
    print("-" * 40)
    results_env = run_environment_analysis()
    all_results.extend(results_env)

    print("\n[E] PSYCHOLOGY (5 curves)")
    print("-" * 40)
    results_psy = run_psychology_analysis()
    all_results.extend(results_psy)

    print("\n[F] PHYSICS (4 curves)")
    print("-" * 40)
    results_phys = run_physics_analysis()
    all_results.extend(results_phys)

    print("\n[G] POLITICAL SCIENCE (5 curves)")
    print("-" * 40)
    results_pol = run_political_analysis()
    all_results.extend(results_pol)

    print("\n[H] AGRICULTURE (4 curves)")
    print("-" * 40)
    results_agr = run_agriculture_analysis()
    all_results.extend(results_agr)

    # Substitute real data where available
    print("\n" + "=" * 70)
    print("SUBSTITUTING REAL DATA (World Bank API)")
    print("=" * 70)
    all_results = _try_real_data_substitution(all_results)

    # Summary
    print("\n" + "=" * 70)
    print("SUMMARY OF RESULTS")
    print("=" * 70)

    df = format_results_table(all_results)

    # Print verdict summary
    verdict_counts = df['Verdict'].value_counts()
    print(f"\nTotal curves analyzed: {len(df)}")
    print("\nVerdict Distribution:")
    for v, c in verdict_counts.items():
        print(f"  {v}: {c} ({100*c/len(df):.1f}%)")

    # Save results
    output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'results')
    os.makedirs(output_dir, exist_ok=True)

    # Save summary table
    df.to_csv(os.path.join(output_dir, 'summary_table.csv'), index=False)
    print(f"\nSummary table saved to: {output_dir}/summary_table.csv")

    # Save full results as JSON
    def convert_numpy(obj):
        if isinstance(obj, (np.integer,)):
            return int(obj)
        elif isinstance(obj, (np.floating,)):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, np.bool_):
            return bool(obj)
        return obj

    results_serializable = []
    for r in all_results:
        r_clean = json.loads(json.dumps(r, default=convert_numpy))
        results_serializable.append(r_clean)

    with open(os.path.join(output_dir, 'full_results.json'), 'w') as f:
        json.dump(results_serializable, f, indent=2)
    print(f"Full results saved to: {output_dir}/full_results.json")

    # Print formatted table
    print("\n" + "=" * 70)
    print("DETAILED RESULTS TABLE")
    print("=" * 70)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 200)
    pd.set_option('display.max_colwidth', 25)

    # Print by category
    for cat in df['Category'].unique():
        print(f"\n--- {cat} ---")
        cat_df = df[df['Category'] == cat][['Curve', 'N', 'p (full)', 'p (clean)',
                                             'AIC best', 'BIC best', 'Verdict']]
        print(cat_df.to_string(index=False))

    return all_results, df


if __name__ == "__main__":
    all_results, df = main()

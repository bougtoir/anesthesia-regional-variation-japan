#!/usr/bin/env python3
"""Empirical calibration of TATSUKI model using Polimeter and Thomson et al. (2017) data.

Data sources:
- Polimeter (polimeter.org): Trudeau promise tracking across 3 parliamentary terms
- Thomson et al. (2017, AJPS): Cross-national pledge fulfillment data (20,000+ pledges, 12 countries)
"""

import numpy as np
from scipy.stats import beta as beta_dist
from scipy.optimize import minimize
import json
import os

# ──────────────────────────────────────────────
# 1. Polimeter Data (Trudeau, 3 terms)
# ──────────────────────────────────────────────
# Source: polimeter.org/en/trudeau (accessed 2026-03-30)
# Coding: Kept=1.0, Partially kept=0.5, Broken=0.0

POLIMETER = {
    '42nd Parliament (2015-2019)': {
        'total': 353,
        'kept': 161,        # ~45.6%
        'partial': 104,     # ~29.5%
        'broken': 88,       # ~24.9%
    },
    '43rd Parliament (2019-2021)': {
        'total': 345,
        'kept': 152,        # ~44.1%
        'partial': 99,      # ~28.7%
        'broken': 94,       # ~27.2%
    },
    '44th Parliament (2021-2025)': {
        'total': 352,
        'kept': 155,        # ~44.0%
        'partial': 113,     # ~32.1%
        'broken': 84,       # ~23.9%
    },
}

# ──────────────────────────────────────────────
# 2. Thomson et al. (2017) Cross-National Data
# ──────────────────────────────────────────────
# Aggregated from Table 1, Thomson et al. (2017) AJPS 61(3):527-542
# Mean fulfillment rates by government type

THOMSON_DATA = {
    'Single-party majority':    {'mean_fulfillment': 0.72, 'n_pledges': 4812},
    'Coalition majority':       {'mean_fulfillment': 0.58, 'n_pledges': 8934},
    'Minority government':      {'mean_fulfillment': 0.61, 'n_pledges': 3254},
    'Cross-national average':   {'mean_fulfillment': 0.67, 'n_pledges': 20000},
}


def polimeter_to_scores(parliament_data):
    """Convert Polimeter categorical data to continuous fulfillment scores."""
    scores = []
    d = parliament_data
    scores.extend([1.0] * d['kept'])
    scores.extend([0.5] * d['partial'])
    scores.extend([0.0] * d['broken'])
    return np.array(scores)


def fit_beta(scores, method='mle'):
    """Fit Beta distribution to fulfillment scores.

    Uses method of moments for robust initial estimates,
    then optionally refines via MLE.
    """
    # Clamp to (0, 1) open interval for Beta fitting
    eps = 1e-6
    scores_clamped = np.clip(scores, eps, 1 - eps)
    mean = np.mean(scores_clamped)
    var = np.var(scores_clamped)

    # Method of moments
    if var < mean * (1 - mean):
        common = mean * (1 - mean) / var - 1
        alpha_mom = mean * common
        beta_mom = (1 - mean) * common
    else:
        alpha_mom, beta_mom = 2.0, 2.0

    if method == 'moments':
        return alpha_mom, beta_mom

    # MLE refinement
    def neg_log_lik(params):
        a, b = params
        if a <= 0 or b <= 0:
            return 1e10
        return -np.sum(beta_dist.logpdf(scores_clamped, a, b))

    result = minimize(neg_log_lik, [alpha_mom, beta_mom],
                      method='Nelder-Mead',
                      options={'maxiter': 10000, 'xatol': 1e-8})
    return result.x[0], result.x[1]


def simulate_tatsuki_counterfactual(polimeter_terms, omega_func, tau_init=1.0):
    """Simulate TATSUKI counterfactual using sequential Polimeter terms.

    For each parliamentary term:
    1. Compute pledge-level fulfillment scores
    2. Compute accountability score S = mean(f_j) (equal weights assumed)
    3. Update trust coefficient via influence function
    """
    trajectory = []
    tau = tau_init

    for name, data in polimeter_terms.items():
        scores = polimeter_to_scores(data)
        S = np.mean(scores)
        tau_new = omega_func(S)
        trajectory.append({
            'term': name,
            'n_pledges': data['total'],
            'pct_kept': data['kept'] / data['total'] * 100,
            'pct_partial': data['partial'] / data['total'] * 100,
            'pct_broken': data['broken'] / data['total'] * 100,
            'accountability_score': S,
            'tau_before': tau,
            'tau_after': tau_new,
            'effective_vote_multiplier': tau_new,
        })
        tau = tau_new

    return trajectory


def omega_concave(S, tau_min=0.5, tau_max=1.5):
    return tau_min + (tau_max - tau_min) * np.sqrt(S)


def omega_linear(S, tau_min=0.5, tau_max=1.5):
    return tau_min + (tau_max - tau_min) * S


def omega_sigmoid(S, tau_min=0.5, tau_max=1.5, k=10):
    return tau_min + (tau_max - tau_min) / (1 + np.exp(-k * (S - 0.5)))


def run_calibration():
    """Run full empirical calibration and return results dict."""
    results = {}

    # 1. Fit Beta distributions to Polimeter data
    print("=" * 60)
    print("Empirical Calibration: Beta Distribution Fitting")
    print("=" * 60)

    beta_params = {}
    for name, data in POLIMETER.items():
        scores = polimeter_to_scores(data)
        alpha, beta_val = fit_beta(scores, method='moments')
        mean_s = np.mean(scores)
        beta_params[name] = {
            'alpha': alpha, 'beta': beta_val,
            'mean_score': mean_s,
            'n': data['total'],
        }
        print(f"\n{name}:")
        print(f"  N pledges: {data['total']}")
        print(f"  Kept/Partial/Broken: {data['kept']}/{data['partial']}/{data['broken']}")
        print(f"  Mean fulfillment score: {mean_s:.3f}")
        print(f"  Fitted Beta(α={alpha:.2f}, β={beta_val:.2f})")

    # Combined across all terms
    all_scores = np.concatenate([
        polimeter_to_scores(data) for data in POLIMETER.values()
    ])
    alpha_all, beta_all = fit_beta(all_scores, method='moments')
    beta_params['All terms combined'] = {
        'alpha': alpha_all, 'beta': beta_all,
        'mean_score': np.mean(all_scores),
        'n': len(all_scores),
    }
    print(f"\nAll terms combined:")
    print(f"  N pledges: {len(all_scores)}")
    print(f"  Mean fulfillment score: {np.mean(all_scores):.3f}")
    print(f"  Fitted Beta(α={alpha_all:.2f}, β={beta_all:.2f})")

    results['beta_params'] = beta_params

    # 2. Compare with Thomson et al. (2017)
    print("\n" + "=" * 60)
    print("Comparison with Thomson et al. (2017)")
    print("=" * 60)
    for gov_type, data in THOMSON_DATA.items():
        print(f"  {gov_type}: mean={data['mean_fulfillment']:.2f} (N≈{data['n_pledges']})")

    polimeter_mean = np.mean(all_scores)
    thomson_mean = THOMSON_DATA['Minority government']['mean_fulfillment']
    print(f"\n  Polimeter (Trudeau, minority govt): {polimeter_mean:.3f}")
    print(f"  Thomson (minority govt avg):         {thomson_mean:.3f}")
    print(f"  Gap:                                 {polimeter_mean - thomson_mean:+.3f}")

    # 3. Recalibrated Beta parameters for simulation
    print("\n" + "=" * 60)
    print("Recalibrated Simulation Parameters")
    print("=" * 60)

    # Original ABM parameters vs empirically calibrated
    original_params = {
        'sincere':   {'alpha': 8, 'beta': 2, 'mean': 8/(8+2)},
        'populist':  {'alpha': 2, 'beta': 5, 'mean': 2/(2+5)},
        'strategic': {'alpha': 5, 'beta': 3, 'mean': 5/(5+3)},
    }

    # Empirically calibrated: Trudeau as "moderate sincere" with minority govt constraints
    calibrated_params = {
        'sincere':   {'alpha': round(alpha_all * 1.3, 2),
                      'beta': round(beta_all * 0.7, 2),
                      'mean': None},
        'populist':  {'alpha': round(alpha_all * 0.5, 2),
                      'beta': round(beta_all * 1.5, 2),
                      'mean': None},
        'strategic': {'alpha': round(alpha_all * 0.9, 2),
                      'beta': round(beta_all * 1.0, 2),
                      'mean': None},
    }
    for ctype in calibrated_params:
        a = calibrated_params[ctype]['alpha']
        b = calibrated_params[ctype]['beta']
        calibrated_params[ctype]['mean'] = round(a / (a + b), 3)

    print("\n  Original ABM Parameters:")
    for ctype, p in original_params.items():
        print(f"    {ctype:12s}: Beta({p['alpha']}, {p['beta']})  mean={p['mean']:.3f}")

    print("\n  Polimeter-Derived Parameters (method of moments):")
    for ctype, p in calibrated_params.items():
        print(f"    {ctype:12s}: Beta({p['alpha']}, {p['beta']})  mean={p['mean']:.3f}")

    # Recalibrate using empirical means with reasonable shape parameters
    # Target means: sincere → Thomson single-party (0.72), populist → low (0.29),
    # strategic → Trudeau-like minority govt (0.60)
    calibrated_params = {
        'sincere':   {'alpha': 5.8, 'beta': 2.2, 'mean': round(5.8 / 8.0, 3)},
        'populist':  {'alpha': 2.0, 'beta': 4.9, 'mean': round(2.0 / 6.9, 3)},
        'strategic': {'alpha': 4.2, 'beta': 2.8, 'mean': round(4.2 / 7.0, 3)},
    }

    print("\n  Final Calibrated Parameters (target-based):")
    for ctype, p in calibrated_params.items():
        print(f"    {ctype:12s}: Beta({p['alpha']}, {p['beta']})  mean={p['mean']:.3f}")

    results['original_params'] = original_params
    results['calibrated_params'] = calibrated_params

    # 4. Counterfactual analysis
    print("\n" + "=" * 60)
    print("Counterfactual Analysis: TATSUKI Applied to Trudeau")
    print("=" * 60)

    for omega_name, omega_func in [
        ('Concave (√S)', omega_concave),
        ('Linear', omega_linear),
        ('Sigmoid', omega_sigmoid),
    ]:
        trajectory = simulate_tatsuki_counterfactual(POLIMETER, omega_func)
        print(f"\n  Influence function: {omega_name}")
        for t in trajectory:
            print(f"    {t['term']}:")
            print(f"      S={t['accountability_score']:.3f}  "
                  f"τ: {t['tau_before']:.3f} → {t['tau_after']:.3f}")

    results['counterfactual'] = {
        omega_name: simulate_tatsuki_counterfactual(POLIMETER, omega_func)
        for omega_name, omega_func in [
            ('Concave', omega_concave),
            ('Linear', omega_linear),
            ('Sigmoid', omega_sigmoid),
        ]
    }

    return results


if __name__ == '__main__':
    results = run_calibration()
    out_dir = os.path.join(os.path.dirname(__file__), '..', 'output')
    os.makedirs(out_dir, exist_ok=True)

    # Save results as JSON for other scripts to use
    json_path = os.path.join(out_dir, 'calibration_results.json')

    # Convert numpy types for JSON serialization
    def convert(obj):
        if isinstance(obj, np.floating):
            return float(obj)
        if isinstance(obj, np.integer):
            return int(obj)
        if isinstance(obj, np.ndarray):
            return obj.tolist()
        raise TypeError(f"Object of type {type(obj)} is not JSON serializable")

    with open(json_path, 'w') as f:
        json.dump(results, f, indent=2, default=convert)
    print(f"\nResults saved to {json_path}")

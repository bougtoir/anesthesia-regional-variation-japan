#!/usr/bin/env python3
"""
CNIP Simulation: Perioperative Outcome Trial
=============================================

Scenario: RCT comparing a new enhanced recovery protocol (ERP) vs standard care
for 30-day postoperative complications after major abdominal surgery.

Ground-truth biological model:
  Y_obs(i) = S(i) + N_conf(i) + N_bio(i) + N_meas(i) + N_rand(i)

  S(i)     = treatment effect (ERP reduces complication risk)
  N_conf   = age, BMI, ASA-PS, diabetes, surgery duration (measured confounders)
  N_bio    = baseline CRP (inflammation), preop hemoglobin (fitness)
  N_meas   = inter-site complication grading variability
  N_rand   = irreducible stochastic noise

Biological hypotheses to test:
  H1: CNIP recovers the true treatment effect (ATE) more precisely than naive analysis
  H2: Noise model explains ~40% of outcome variance (ρ² ≈ 0.4)
  H3: CNIP residuals are independent of measured confounders
  H4: CNIP achieves equivalent statistical power with fewer patients
  H5: Phase 4 detects a planted subgroup effect (elderly patients benefit more)
"""

import os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy import stats
from sklearn.linear_model import LinearRegression, LogisticRegression
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import cross_val_predict, cross_val_score
from sklearn.metrics import r2_score
import warnings
warnings.filterwarnings('ignore')

np.random.seed(42)

# ---------------------------------------------------------------------------
# 1. DATA GENERATION — biologically plausible perioperative trial
# ---------------------------------------------------------------------------

def generate_trial_data(n=2000):
    """Generate synthetic RCT data with known ground truth.

    Returns dict with patient data, true signal, true noise components.
    """
    # Treatment assignment (1:1 randomization)
    treatment = np.random.binomial(1, 0.5, n)

    # --- Measured confounders (N_conf) ---
    age = np.random.normal(65, 12, n).clip(30, 90)
    bmi = np.random.normal(26, 5, n).clip(16, 45)
    asa_ps = np.random.choice([1, 2, 3, 4], n, p=[0.05, 0.40, 0.45, 0.10])
    diabetes = np.random.binomial(1, 0.25, n)
    surgery_duration_min = np.random.normal(180, 60, n).clip(60, 420)

    # Confounder noise contribution (continuous complication risk score)
    n_conf = (
        0.012 * (age - 65)
        + 0.006 * (bmi - 26)
        + 0.10 * (asa_ps - 2)
        + 0.12 * diabetes
        + 0.0008 * (surgery_duration_min - 180)
    )

    # --- Biological variability (N_bio) ---
    crp_baseline = np.random.lognormal(1.5, 0.8, n).clip(0.1, 150)   # mg/L
    hemoglobin = np.random.normal(13, 2, n).clip(7, 18)               # g/dL

    n_bio = (
        0.003 * (crp_baseline - 5)
        + 0.04 * (13 - hemoglobin)
    )

    # --- Measurement error (N_meas) ---
    site = np.random.choice(range(10), n)
    site_effect = np.random.normal(0, 0.06, 10)
    n_meas = site_effect[site]

    # --- Irreducible randomness (N_rand) ---
    n_rand = np.random.normal(0, 0.25, n)

    # --- True treatment effect (Signal) ---
    # Base ATE: -0.15 (ERP reduces complication risk by 15 percentage points)
    # Subgroup: elderly (age > 75) benefit more (additional -0.10)
    base_effect = -0.15
    elderly_bonus = -0.10 * (age > 75).astype(float)
    signal = treatment * (base_effect + elderly_bonus)

    # --- Total noise and observed outcome ---
    total_noise = n_conf + n_bio + n_meas + n_rand
    y_latent = 0.50 + signal + total_noise   # baseline ~50% keeps most values in (0,1)
    y_obs = y_latent  # no clipping — continuous outcome (e.g., complication severity index)

    # For binary outcome (complication yes/no)
    y_binary = np.random.binomial(1, y_obs.clip(0.01, 0.99))

    data = {
        'n': n,
        'treatment': treatment,
        'age': age,
        'bmi': bmi,
        'asa_ps': asa_ps,
        'diabetes': diabetes,
        'surgery_duration_min': surgery_duration_min,
        'crp_baseline': crp_baseline,
        'hemoglobin': hemoglobin,
        'site': site,
        'y_obs': y_obs,
        'y_binary': y_binary,
        'signal': signal,
        'n_conf': n_conf,
        'n_bio': n_bio,
        'n_meas': n_meas,
        'n_rand': n_rand,
        'total_noise': total_noise,
        'true_ate': base_effect,
        'true_elderly_bonus': -0.10,
    }
    return data


# ---------------------------------------------------------------------------
# 2. CNIP PIPELINE
# ---------------------------------------------------------------------------

def phase1_noise_forward_model(data):
    """Phase 1: Build noise forward model F(θ; X) from auxiliary data.

    Key design: model is trained ONLY on the control arm (where Y = baseline + noise,
    no treatment signal) then applied to ALL patients. This prevents signal leakage
    — the model never sees treatment effects during training.

    Uses cross-validated predictions on the control arm to get honest ρ²,
    then refits on full control arm for final predictions on the treatment arm.
    """
    X = np.column_stack([
        data['age'],
        data['bmi'],
        data['asa_ps'],
        data['diabetes'],
        data['surgery_duration_min'],
        data['crp_baseline'],
        data['hemoglobin'],
        np.eye(10)[data['site']],   # one-hot site encoding
    ])

    control_mask = data['treatment'] == 0
    X_ctrl = X[control_mask]
    y_ctrl = data['y_obs'][control_mask]
    X_treat = X[~control_mask]

    model_spec = GradientBoostingRegressor(
        n_estimators=200, max_depth=3, learning_rate=0.05,
        subsample=0.8, min_samples_leaf=20, random_state=42
    )

    # Cross-validated predictions on control arm (honest, out-of-fold)
    noise_pred_ctrl_cv = cross_val_predict(model_spec, X_ctrl, y_ctrl, cv=5)

    # Refit on full control arm, then predict for treatment arm
    model_ctrl = GradientBoostingRegressor(
        n_estimators=200, max_depth=3, learning_rate=0.05,
        subsample=0.8, min_samples_leaf=20, random_state=42
    )
    model_ctrl.fit(X_ctrl, y_ctrl)
    noise_pred_treat = model_ctrl.predict(X_treat)

    # Assemble full noise prediction vector
    noise_pred = np.empty(len(data['y_obs']))
    noise_pred[control_mask] = noise_pred_ctrl_cv    # CV predictions (no overfit)
    noise_pred[~control_mask] = noise_pred_treat     # from control-arm model

    # Feature importance (from control-arm model)
    feature_names = [
        'Age', 'BMI', 'ASA-PS', 'Diabetes', 'Surgery Duration',
        'CRP', 'Hemoglobin',
    ] + [f'Site_{i}' for i in range(10)]
    importances = model_ctrl.feature_importances_

    # Cross-validated R² on control arm
    cv_r2_ctrl = cross_val_score(model_spec, X_ctrl, y_ctrl, cv=5, scoring='r2')

    return {
        'model': model_ctrl,
        'noise_pred': noise_pred,
        'X': X,
        'feature_names': feature_names,
        'importances': importances,
        'cv_r2_ctrl': cv_r2_ctrl,
    }


def phase2_bayesian_estimation(data, phase1):
    """Phase 2: Estimate noise parameters (bootstrap CI for ρ²).

    ρ² is computed from cross-validated predictions on the control arm
    to avoid overfitting bias. This gives an honest estimate of how much
    outcome variance the noise model can explain.
    """
    noise_pred = phase1['noise_pred']
    control_mask = data['treatment'] == 0
    y_ctrl = data['y_obs'][control_mask]
    pred_ctrl = noise_pred[control_mask]   # these are CV predictions (no overfit)

    # ρ² from cross-validated predictions on control arm
    ss_total = np.sum((y_ctrl - y_ctrl.mean()) ** 2)
    ss_residual = np.sum((y_ctrl - pred_ctrl) ** 2)
    rho_squared = 1 - ss_residual / ss_total

    # Bootstrap CI for ρ²
    n_boot = 1000
    rho2_boot = []
    for _ in range(n_boot):
        idx = np.random.choice(len(y_ctrl), len(y_ctrl), replace=True)
        y_b = y_ctrl[idx]
        p_b = pred_ctrl[idx]
        ss_t = np.sum((y_b - y_b.mean()) ** 2)
        ss_r = np.sum((y_b - p_b) ** 2)
        if ss_t > 0:
            rho2_boot.append(1 - ss_r / ss_t)
    rho2_ci = np.percentile(rho2_boot, [2.5, 97.5])

    # Cross-validated R² from Phase 1 (control arm, k-fold scoring)
    cv_r2 = phase1['cv_r2_ctrl']

    return {
        'rho_squared': rho_squared,
        'rho2_ci': rho2_ci,
        'cv_r2_mean': cv_r2.mean(),
        'cv_r2_std': cv_r2.std(),
        'rho2_bootstrap': np.array(rho2_boot),
    }


def phase3_residual_generation(data, phase1):
    """Phase 3: Generate residuals (clean signal extraction)."""
    residuals = data['y_obs'] - phase1['noise_pred']
    return {
        'residuals': residuals,
    }


def phase4_discovery(data, phase3):
    """Phase 4: Hypothesis-free discovery in residuals."""
    residuals = phase3['residuals']
    treatment = data['treatment']

    # --- CNIP ATE from residuals ---
    ate_cnip = residuals[treatment == 1].mean() - residuals[treatment == 0].mean()
    se_cnip = np.sqrt(
        np.var(residuals[treatment == 1]) / np.sum(treatment == 1)
        + np.var(residuals[treatment == 0]) / np.sum(treatment == 0)
    )
    ci_cnip = (ate_cnip - 1.96 * se_cnip, ate_cnip + 1.96 * se_cnip)
    p_cnip = 2 * (1 - stats.norm.cdf(abs(ate_cnip / se_cnip)))

    # --- Naive ATE (no adjustment) ---
    ate_naive = (data['y_obs'][treatment == 1].mean()
                 - data['y_obs'][treatment == 0].mean())
    se_naive = np.sqrt(
        np.var(data['y_obs'][treatment == 1]) / np.sum(treatment == 1)
        + np.var(data['y_obs'][treatment == 0]) / np.sum(treatment == 0)
    )
    ci_naive = (ate_naive - 1.96 * se_naive, ate_naive + 1.96 * se_naive)
    p_naive = 2 * (1 - stats.norm.cdf(abs(ate_naive / se_naive)))

    # --- Subgroup analysis: elderly vs non-elderly ---
    elderly = data['age'] > 75
    subgroups = {}
    for name, mask in [('Elderly (>75)', elderly), ('Non-elderly (≤75)', ~elderly)]:
        r_t = residuals[(treatment == 1) & mask]
        r_c = residuals[(treatment == 0) & mask]
        ate_sg = r_t.mean() - r_c.mean()
        se_sg = np.sqrt(np.var(r_t) / len(r_t) + np.var(r_c) / len(r_c))
        subgroups[name] = {
            'ate': ate_sg,
            'se': se_sg,
            'ci': (ate_sg - 1.96 * se_sg, ate_sg + 1.96 * se_sg),
            'p': 2 * (1 - stats.norm.cdf(abs(ate_sg / se_sg))),
            'n': int(mask.sum()),
        }

    # Interaction test (age > 75 × treatment)
    X_int = np.column_stack([treatment, elderly, treatment * elderly])
    model_int = LinearRegression()
    model_int.fit(X_int, residuals)
    interaction_coef = model_int.coef_[2]

    return {
        'ate_cnip': ate_cnip, 'se_cnip': se_cnip, 'ci_cnip': ci_cnip, 'p_cnip': p_cnip,
        'ate_naive': ate_naive, 'se_naive': se_naive, 'ci_naive': ci_naive, 'p_naive': p_naive,
        'subgroups': subgroups,
        'interaction_coef': interaction_coef,
    }


# ---------------------------------------------------------------------------
# 3. VALIDATION AGAINST BIOLOGICAL HYPOTHESES
# ---------------------------------------------------------------------------

def validate_hypotheses(data, phase1, phase2, phase3, phase4):
    """Validate CNIP results against biological hypotheses."""
    results = {}

    # H1: CNIP recovers true ATE more precisely
    true_ate = data['true_ate']
    ate_cnip = phase4['ate_cnip']
    ate_naive = phase4['ate_naive']
    bias_cnip = abs(ate_cnip - true_ate)
    bias_naive = abs(ate_naive - true_ate)
    se_reduction = (phase4['se_naive'] - phase4['se_cnip']) / phase4['se_naive'] * 100

    results['H1'] = {
        'description': 'CNIP recovers true ATE more precisely than naive analysis',
        'true_ate': true_ate,
        'ate_cnip': ate_cnip,
        'ate_naive': ate_naive,
        'bias_cnip': bias_cnip,
        'bias_naive': bias_naive,
        'se_cnip': phase4['se_cnip'],
        'se_naive': phase4['se_naive'],
        'se_reduction_pct': se_reduction,
        'passed': bias_cnip < bias_naive and phase4['se_cnip'] < phase4['se_naive'],
    }

    # H2: Noise model explains ~40% of outcome variance
    rho2 = phase2['rho_squared']
    results['H2'] = {
        'description': 'Noise model ρ² ≈ 0.4 (explains ~40% of variance)',
        'rho_squared': rho2,
        'rho2_ci': phase2['rho2_ci'],
        'cv_r2': phase2['cv_r2_mean'],
        'passed': 0.25 <= rho2 <= 0.60,
    }

    # H3: Residuals are independent of measured confounders
    residuals = phase3['residuals']
    confounder_names = ['age', 'bmi', 'asa_ps', 'diabetes', 'surgery_duration_min',
                        'crp_baseline', 'hemoglobin']
    correlations = {}
    max_abs_r = 0.0
    for name in confounder_names:
        r, p = stats.pearsonr(residuals, data[name])
        correlations[name] = {'r': r, 'p': p}
        max_abs_r = max(max_abs_r, abs(r))

    results['H3'] = {
        'description': 'Residuals largely independent of confounders (max |r| < 0.10)',
        'correlations': correlations,
        'max_abs_r': max_abs_r,
        'passed': max_abs_r < 0.10,
    }

    # H4: Effective sample size multiplier
    n_eff_multiplier = 1 / (1 - rho2) if rho2 < 1 else float('inf')
    results['H4'] = {
        'description': 'Effective sample size multiplier > 1.3',
        'n_eff_multiplier': n_eff_multiplier,
        'actual_n': data['n'],
        'effective_n': data['n'] * n_eff_multiplier,
        'passed': n_eff_multiplier > 1.3,
    }

    # H5: Subgroup discovery — elderly benefit more
    elderly_ate = phase4['subgroups']['Elderly (>75)']['ate']
    non_elderly_ate = phase4['subgroups']['Non-elderly (≤75)']['ate']
    interaction = phase4['interaction_coef']
    results['H5'] = {
        'description': 'Elderly (>75) show larger treatment benefit (planted subgroup effect)',
        'elderly_ate': elderly_ate,
        'non_elderly_ate': non_elderly_ate,
        'interaction_coef': interaction,
        'true_elderly_bonus': data['true_elderly_bonus'],
        'passed': elderly_ate < non_elderly_ate,
    }

    return results


# ---------------------------------------------------------------------------
# 4. POWER SIMULATION (varying sample sizes)
# ---------------------------------------------------------------------------

def power_simulation(n_sims=500):
    """Compare CNIP vs naive power across sample sizes.

    CNIP uses control-arm-only model (same as main pipeline) to avoid
    signal leakage.
    """
    sample_sizes = [100, 200, 400, 600, 800, 1000]
    results = {'naive': {}, 'cnip': {}}

    for n in sample_sizes:
        sig_naive = 0
        sig_cnip = 0
        for sim in range(n_sims):
            np.random.seed(sim * 1000 + n)
            d = generate_trial_data(n)

            t = d['treatment']
            y = d['y_obs']

            # Naive test
            t_stat, p_naive = stats.ttest_ind(y[t == 1], y[t == 0])
            if p_naive < 0.05:
                sig_naive += 1

            # CNIP test: fit noise model on CONTROL arm only, apply to all
            X = np.column_stack([
                d['age'], d['bmi'], d['asa_ps'], d['diabetes'],
                d['surgery_duration_min'], d['crp_baseline'], d['hemoglobin'],
            ])
            ctrl = t == 0
            model = LinearRegression()
            model.fit(X[ctrl], y[ctrl])
            residuals = y - model.predict(X)
            t_stat_c, p_cnip = stats.ttest_ind(residuals[t == 1], residuals[t == 0])
            if p_cnip < 0.05:
                sig_cnip += 1

        results['naive'][n] = sig_naive / n_sims
        results['cnip'][n] = sig_cnip / n_sims
        print(f'  n={n:4d}: naive power={sig_naive/n_sims:.3f}, '
              f'CNIP power={sig_cnip/n_sims:.3f}')

    return results


# ---------------------------------------------------------------------------
# 5. VISUALIZATION
# ---------------------------------------------------------------------------

def create_result_figures(data, phase1, phase2, phase3, phase4, hypotheses, power_results):
    """Generate comprehensive result figures."""

    fig = plt.figure(figsize=(20, 24))
    gs = fig.add_gridspec(4, 3, hspace=0.35, wspace=0.3)

    # --- Panel A: ATE comparison ---
    ax = fig.add_subplot(gs[0, 0])
    methods = ['True ATE', 'Naive', 'CNIP']
    ates = [data['true_ate'], phase4['ate_naive'], phase4['ate_cnip']]
    ses = [0, phase4['se_naive'], phase4['se_cnip']]
    colors = ['#27AE60', '#E74C3C', '#2980B9']
    ax.barh(methods, ates, xerr=[1.96 * s for s in ses], color=colors, alpha=0.8,
            capsize=5, height=0.5)
    ax.axvline(x=data['true_ate'], color='#27AE60', linestyle='--', alpha=0.7, label='True ATE')
    ax.set_xlabel('Average Treatment Effect (ATE)')
    ax.set_title('A. Treatment Effect Recovery', fontweight='bold')
    ax.legend(fontsize=8)

    # --- Panel B: Noise model ρ² distribution ---
    ax = fig.add_subplot(gs[0, 1])
    ax.hist(phase2['rho2_bootstrap'], bins=40, color='#3498DB', alpha=0.7, edgecolor='white')
    ax.axvline(x=phase2['rho_squared'], color='red', linewidth=2, label=f'ρ²={phase2["rho_squared"]:.3f}')
    ax.axvline(x=phase2['rho2_ci'][0], color='orange', linestyle='--',
               label=f'95% CI [{phase2["rho2_ci"][0]:.3f}, {phase2["rho2_ci"][1]:.3f}]')
    ax.axvline(x=phase2['rho2_ci'][1], color='orange', linestyle='--')
    ax.set_xlabel('ρ² (Noise Model Accuracy)')
    ax.set_ylabel('Bootstrap Count')
    ax.set_title('B. Noise Model Accuracy (ρ²)', fontweight='bold')
    ax.legend(fontsize=8)

    # --- Panel C: Residual vs confounder correlations ---
    ax = fig.add_subplot(gs[0, 2])
    corr_data = hypotheses['H3']['correlations']
    names = list(corr_data.keys())
    r_values = [corr_data[n]['r'] for n in names]
    bar_colors = ['#E74C3C' if abs(r) > 0.05 else '#27AE60' for r in r_values]
    ax.barh(names, r_values, color=bar_colors, alpha=0.8)
    ax.axvline(x=0.05, color='gray', linestyle='--', alpha=0.5)
    ax.axvline(x=-0.05, color='gray', linestyle='--', alpha=0.5)
    ax.set_xlabel('Pearson r (Residual vs Confounder)')
    ax.set_title('C. Residual Independence (H3)', fontweight='bold')
    ax.set_xlim(-0.15, 0.15)

    # --- Panel D: Feature importance ---
    ax = fig.add_subplot(gs[1, 0])
    n_features = min(10, len(phase1['feature_names']))
    sorted_idx = np.argsort(phase1['importances'])[-n_features:]
    ax.barh(
        [phase1['feature_names'][i] for i in sorted_idx],
        phase1['importances'][sorted_idx],
        color='#8E44AD', alpha=0.8
    )
    ax.set_xlabel('Feature Importance (Gradient Boosting)')
    ax.set_title('D. Noise Model Feature Importance', fontweight='bold')

    # --- Panel E: Subgroup effects ---
    ax = fig.add_subplot(gs[1, 1])
    sg = phase4['subgroups']
    sg_names = list(sg.keys())
    sg_ates = [sg[n]['ate'] for n in sg_names]
    sg_ses = [sg[n]['se'] for n in sg_names]
    sg_colors = ['#E74C3C', '#3498DB']
    ax.barh(sg_names, sg_ates,
            xerr=[1.96 * s for s in sg_ses],
            color=sg_colors, alpha=0.8, capsize=5, height=0.4)
    ax.axvline(x=data['true_ate'], color='gray', linestyle='--', alpha=0.5, label='Overall True ATE')
    ax.axvline(x=data['true_ate'] + data['true_elderly_bonus'], color='#E74C3C',
               linestyle=':', alpha=0.7, label='True Elderly ATE')
    ax.set_xlabel('Subgroup ATE (CNIP)')
    ax.set_title('E. Subgroup Discovery (H5)', fontweight='bold')
    ax.legend(fontsize=8)

    # --- Panel F: Power curves ---
    ax = fig.add_subplot(gs[1, 2])
    ns = sorted(power_results['naive'].keys())
    power_naive = [power_results['naive'][n] * 100 for n in ns]
    power_cnip = [power_results['cnip'][n] * 100 for n in ns]
    ax.plot(ns, power_naive, 'o-', color='#95A5A6', linewidth=2, markersize=6, label='Naive')
    ax.plot(ns, power_cnip, 's-', color='#2980B9', linewidth=2, markersize=6, label='CNIP')
    ax.axhline(y=80, color='red', linestyle='--', alpha=0.5, linewidth=1.5)
    ax.text(max(ns) * 0.95, 82, '80% power', fontsize=9, color='red', ha='right')
    ax.set_xlabel('Sample Size (n)')
    ax.set_ylabel('Statistical Power (%)')
    ax.set_title('F. Power Comparison (H4)', fontweight='bold')
    ax.legend(fontsize=9)
    ax.set_ylim(0, 105)
    ax.grid(True, alpha=0.3)

    # --- Panel G: Observed vs predicted noise ---
    ax = fig.add_subplot(gs[2, 0])
    true_noise = data['total_noise']
    pred_noise = phase1['noise_pred'] - np.mean(phase1['noise_pred'])
    ax.scatter(true_noise, pred_noise, alpha=0.15, s=8, color='#2980B9')
    ax.plot([-1, 1], [-1, 1], 'r--', alpha=0.5, label='Perfect recovery')
    r_noise = np.corrcoef(true_noise, pred_noise)[0, 1]
    ax.set_xlabel('True Noise N(i)')
    ax.set_ylabel('Predicted Noise F(θ̂; X_i)')
    ax.set_title(f'G. Noise Recovery (r={r_noise:.3f})', fontweight='bold')
    ax.legend(fontsize=9)
    ax.set_xlim(-1, 1)
    ax.set_ylim(-1, 1)

    # --- Panel H: Residual distribution by treatment ---
    ax = fig.add_subplot(gs[2, 1])
    t = data['treatment']
    residuals = phase3['residuals']
    ax.hist(residuals[t == 0], bins=40, alpha=0.6, color='#95A5A6', label='Control', density=True)
    ax.hist(residuals[t == 1], bins=40, alpha=0.6, color='#2980B9', label='Treatment (ERP)', density=True)
    ax.axvline(x=residuals[t == 0].mean(), color='#95A5A6', linestyle='--', linewidth=2)
    ax.axvline(x=residuals[t == 1].mean(), color='#2980B9', linestyle='--', linewidth=2)
    ax.set_xlabel('CNIP Residual')
    ax.set_ylabel('Density')
    ax.set_title('H. Residual Distribution by Treatment', fontweight='bold')
    ax.legend(fontsize=9)

    # --- Panel I: SE reduction visualization ---
    ax = fig.add_subplot(gs[2, 2])
    se_data = {
        'SE (Naive)': phase4['se_naive'],
        'SE (CNIP)': phase4['se_cnip'],
    }
    bars = ax.bar(se_data.keys(), se_data.values(),
                  color=['#E74C3C', '#2980B9'], alpha=0.8, width=0.5)
    reduction = hypotheses['H1']['se_reduction_pct']
    ax.set_ylabel('Standard Error of ATE')
    ax.set_title(f'I. SE Reduction: {reduction:.1f}%', fontweight='bold')
    for bar, val in zip(bars, se_data.values()):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.001,
                f'{val:.4f}', ha='center', fontsize=10, fontweight='bold')

    # --- Panel J: Hypothesis summary table ---
    ax = fig.add_subplot(gs[3, :])
    ax.axis('off')
    table_data = []
    for h_id in ['H1', 'H2', 'H3', 'H4', 'H5']:
        h = hypotheses[h_id]
        status = 'PASS' if h['passed'] else 'FAIL'
        if h_id == 'H1':
            detail = (f'Bias: CNIP={h["bias_cnip"]:.4f} vs Naive={h["bias_naive"]:.4f}; '
                      f'SE reduction={h["se_reduction_pct"]:.1f}%')
        elif h_id == 'H2':
            detail = f'ρ²={h["rho_squared"]:.3f}, 95%CI=[{h["rho2_ci"][0]:.3f}, {h["rho2_ci"][1]:.3f}]'
        elif h_id == 'H3':
            max_r = max(abs(h['correlations'][n]['r']) for n in h['correlations'])
            detail = f'Max |r|={max_r:.4f}'
        elif h_id == 'H4':
            detail = (f'n_eff multiplier={h["n_eff_multiplier"]:.2f}× '
                      f'(n={h["actual_n"]} → eff. n={h["effective_n"]:.0f})')
        else:
            detail = (f'Elderly ATE={h["elderly_ate"]:.3f} vs '
                      f'Non-elderly ATE={h["non_elderly_ate"]:.3f}; '
                      f'Interaction={h["interaction_coef"]:.4f}')
        table_data.append([h_id, h['description'], status, detail])

    col_labels = ['ID', 'Biological Hypothesis', 'Result', 'Details']
    table = ax.table(cellText=table_data, colLabels=col_labels,
                     loc='center', cellLoc='left',
                     colWidths=[0.05, 0.35, 0.06, 0.54])
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.8)

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_facecolor('#2C3E50')
            cell.set_text_props(color='white', fontweight='bold')
        elif col == 2:
            text = cell.get_text().get_text()
            if text == 'PASS':
                cell.set_facecolor('#D5F5E3')
                cell.set_text_props(fontweight='bold', color='#1E8449')
            else:
                cell.set_facecolor('#FADBD8')
                cell.set_text_props(fontweight='bold', color='#C0392B')
        else:
            cell.set_facecolor('#F8F9FA' if row % 2 == 0 else 'white')

    ax.set_title('J. Biological Hypothesis Validation Summary',
                 fontweight='bold', fontsize=13, pad=20)

    fig.suptitle('CNIP Simulation Results: Perioperative RCT (n=2000)',
                 fontsize=16, fontweight='bold', y=0.995)

    out = 'simulation_results.png'
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f'Saved: {out}')
    return out


# ---------------------------------------------------------------------------
# 6. MAIN
# ---------------------------------------------------------------------------

def main():
    print('=' * 70)
    print('CNIP SIMULATION: Perioperative Outcome Trial')
    print('=' * 70)

    # Generate data
    print('\n[1] Generating synthetic trial data (n=2000)...')
    data = generate_trial_data(n=2000)
    print(f'  Treatment: {data["treatment"].sum()} ERP, '
          f'{(1-data["treatment"]).sum():.0f} control')
    print(f'  True ATE: {data["true_ate"]} (base) + {data["true_elderly_bonus"]} (elderly bonus)')
    print(f'  Observed complication rate: {data["y_obs"].mean():.3f}')
    print(f'  Binary complication rate: {data["y_binary"].mean():.3f}')

    # Phase 1
    print('\n[2] Phase 1: Building noise forward model...')
    p1 = phase1_noise_forward_model(data)
    top_features = sorted(
        zip(p1['feature_names'], p1['importances']),
        key=lambda x: x[1], reverse=True
    )[:5]
    print('  Top features:')
    for name, imp in top_features:
        print(f'    {name}: {imp:.4f}')

    # Phase 2
    print('\n[3] Phase 2: Bayesian noise parameter estimation...')
    p2 = phase2_bayesian_estimation(data, p1)
    print(f'  ρ² = {p2["rho_squared"]:.4f} '
          f'(95% CI: [{p2["rho2_ci"][0]:.4f}, {p2["rho2_ci"][1]:.4f}])')
    print(f'  Cross-validated R² = {p2["cv_r2_mean"]:.4f} ± {p2["cv_r2_std"]:.4f}')

    # Phase 3
    print('\n[4] Phase 3: Residual generation...')
    p3 = phase3_residual_generation(data, p1)
    print(f'  Residual mean: {p3["residuals"].mean():.6f}')
    print(f'  Residual std:  {p3["residuals"].std():.6f}')
    print(f'  Original std:  {data["y_obs"].std():.6f}')
    print(f'  Variance reduction: '
          f'{(1 - p3["residuals"].var()/data["y_obs"].var())*100:.1f}%')

    # Phase 4
    print('\n[5] Phase 4: Hypothesis-free discovery...')
    p4 = phase4_discovery(data, p3)
    print(f'  Naive ATE:  {p4["ate_naive"]:.4f} '
          f'(SE={p4["se_naive"]:.4f}, p={p4["p_naive"]:.6f})')
    print(f'  CNIP ATE:   {p4["ate_cnip"]:.4f} '
          f'(SE={p4["se_cnip"]:.4f}, p={p4["p_cnip"]:.6f})')
    print(f'  True ATE:   {data["true_ate"]}')
    print(f'\n  Subgroup analysis:')
    for name, sg in p4['subgroups'].items():
        print(f'    {name} (n={sg["n"]}): ATE={sg["ate"]:.4f} '
              f'(SE={sg["se"]:.4f}, p={sg["p"]:.6f})')
    print(f'  Interaction coefficient (elderly×treatment): '
          f'{p4["interaction_coef"]:.4f}')

    # Validate
    print('\n[6] Validating biological hypotheses...')
    hypotheses = validate_hypotheses(data, p1, p2, p3, p4)
    for h_id in ['H1', 'H2', 'H3', 'H4', 'H5']:
        h = hypotheses[h_id]
        status = 'PASS' if h['passed'] else 'FAIL'
        print(f'  {h_id}: [{status}] {h["description"]}')

    # Power simulation
    print('\n[7] Running power simulation (500 sims × 6 sample sizes)...')
    power_results = power_simulation(n_sims=500)

    # Visualization
    print('\n[8] Generating result figures...')
    fig_path = create_result_figures(data, p1, p2, p3, p4, hypotheses, power_results)

    # Summary
    print('\n' + '=' * 70)
    print('SIMULATION SUMMARY')
    print('=' * 70)
    n_pass = sum(1 for h in hypotheses.values() if h['passed'])
    n_total = len(hypotheses)
    print(f'Hypotheses passed: {n_pass}/{n_total}')
    print(f'Noise model accuracy (ρ²): {p2["rho_squared"]:.4f}')
    print(f'Effective sample size multiplier: '
          f'{1/(1-p2["rho_squared"]):.2f}×')
    print(f'SE reduction: '
          f'{hypotheses["H1"]["se_reduction_pct"]:.1f}%')
    print(f'True ATE={data["true_ate"]:.2f}, '
          f'CNIP ATE={p4["ate_cnip"]:.4f}, '
          f'Naive ATE={p4["ate_naive"]:.4f}')
    print(f'Figure saved: {fig_path}')
    print('=' * 70)

    return {
        'data': data, 'phase1': p1, 'phase2': p2, 'phase3': p3, 'phase4': p4,
        'hypotheses': hypotheses, 'power_results': power_results,
    }


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    main()

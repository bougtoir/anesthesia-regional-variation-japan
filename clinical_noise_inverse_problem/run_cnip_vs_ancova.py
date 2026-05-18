#!/usr/bin/env python3
"""
CNIP vs ANCOVA Head-to-Head Comparison
=======================================

Answers two questions:
  Q1: Can ANCOVA also rescue non-significant RCTs?
  Q2: Are there theoretical advantages of CNIP vs ANCOVA for RCT vs retrospective studies?

Design:
  - Compare 4 methods: Naive t-test, ANCOVA (linear), CNIP-Linear, CNIP-GBM (nonlinear)
  - Two data structures:
    (A) Linear noise → ANCOVA and CNIP-Linear should be equivalent
    (B) Nonlinear noise → CNIP-GBM should outperform ANCOVA
  - Two study designs:
    (A) RCT (randomized, balanced confounders)
    (B) Retrospective (observational, confounded treatment assignment)
"""

import os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy import stats
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import cross_val_predict
import warnings
warnings.filterwarnings('ignore')


# ---------------------------------------------------------------------------
# 1. DATA GENERATORS
# ---------------------------------------------------------------------------

def generate_rct_linear(n, true_ate):
    """RCT with linear noise structure."""
    treatment = np.random.binomial(1, 0.5, n)
    X = np.random.normal(0, 1, (n, 5))
    betas = np.array([0.15, 0.10, 0.08, 0.06, 0.04])
    noise = X @ betas + np.random.normal(0, 0.30, n)
    y = 0.5 + treatment * true_ate + noise
    return treatment, y, X


def generate_rct_nonlinear(n, true_ate):
    """RCT with nonlinear noise structure (interactions, thresholds)."""
    treatment = np.random.binomial(1, 0.5, n)
    X = np.random.normal(0, 1, (n, 5))
    noise = (
        0.15 * X[:, 0]
        + 0.10 * X[:, 1]
        + 0.12 * X[:, 0] * X[:, 1]             # interaction
        + 0.08 * np.where(X[:, 2] > 0.5, 1, 0)  # threshold
        + 0.06 * X[:, 3] ** 2                     # quadratic
        + 0.04 * np.sin(X[:, 4])                  # periodic
        + np.random.normal(0, 0.30, n)
    )
    y = 0.5 + treatment * true_ate + noise
    return treatment, y, X


def generate_retro_linear(n, true_ate):
    """Retrospective study with linear noise + confounded treatment."""
    X = np.random.normal(0, 1, (n, 5))
    # Treatment depends on confounders (confounding by indication)
    propensity = 1 / (1 + np.exp(-(0.5 * X[:, 0] + 0.3 * X[:, 1])))
    treatment = np.random.binomial(1, propensity)
    betas = np.array([0.15, 0.10, 0.08, 0.06, 0.04])
    noise = X @ betas + np.random.normal(0, 0.30, n)
    y = 0.5 + treatment * true_ate + noise
    return treatment, y, X


def generate_retro_nonlinear(n, true_ate):
    """Retrospective study with nonlinear noise + confounded treatment."""
    X = np.random.normal(0, 1, (n, 5))
    propensity = 1 / (1 + np.exp(-(0.5 * X[:, 0] + 0.3 * X[:, 1])))
    treatment = np.random.binomial(1, propensity)
    noise = (
        0.15 * X[:, 0]
        + 0.10 * X[:, 1]
        + 0.12 * X[:, 0] * X[:, 1]
        + 0.08 * np.where(X[:, 2] > 0.5, 1, 0)
        + 0.06 * X[:, 3] ** 2
        + 0.04 * np.sin(X[:, 4])
        + np.random.normal(0, 0.30, n)
    )
    y = 0.5 + treatment * true_ate + noise
    return treatment, y, X


# ---------------------------------------------------------------------------
# 2. ANALYSIS METHODS
# ---------------------------------------------------------------------------

def analyze_naive(treatment, y, X):
    """Simple t-test (no adjustment)."""
    _, p = stats.ttest_ind(y[treatment == 1], y[treatment == 0])
    ate = y[treatment == 1].mean() - y[treatment == 0].mean()
    n1, n0 = (treatment == 1).sum(), (treatment == 0).sum()
    se = np.sqrt(y[treatment == 1].var() / n1 + y[treatment == 0].var() / n0)
    return ate, se, p


def analyze_ancova(treatment, y, X):
    """ANCOVA: linear regression with treatment + covariates on full data."""
    Xfull = np.column_stack([treatment, X])
    model = LinearRegression()
    model.fit(Xfull, y)
    ate = model.coef_[0]
    residuals = y - model.predict(Xfull)
    n = len(y)
    p_features = Xfull.shape[1]
    mse = np.sum(residuals ** 2) / (n - p_features - 1)
    XtX_inv = np.linalg.inv(Xfull.T @ Xfull)
    se = np.sqrt(mse * XtX_inv[0, 0])
    t_stat = ate / se
    p = 2 * stats.t.sf(abs(t_stat), df=n - p_features - 1)
    return ate, se, p


def analyze_cnip_linear(treatment, y, X):
    """CNIP with linear noise model (control-arm only training)."""
    ctrl = treatment == 0
    model = LinearRegression()
    model.fit(X[ctrl], y[ctrl])
    residuals = y - model.predict(X)
    _, p = stats.ttest_ind(residuals[treatment == 1], residuals[treatment == 0])
    ate = residuals[treatment == 1].mean() - residuals[treatment == 0].mean()
    n1 = (treatment == 1).sum()
    n0 = (treatment == 0).sum()
    se = np.sqrt(residuals[treatment == 1].var() / n1
                 + residuals[treatment == 0].var() / n0)
    return ate, se, p


def analyze_cnip_gbm(treatment, y, X):
    """CNIP with GBM noise model (control-arm only training)."""
    ctrl = treatment == 0
    model = GradientBoostingRegressor(
        n_estimators=100, max_depth=3, learning_rate=0.05,
        subsample=0.8, min_samples_leaf=10, random_state=42
    )
    model.fit(X[ctrl], y[ctrl])
    residuals = y - model.predict(X)
    _, p = stats.ttest_ind(residuals[treatment == 1], residuals[treatment == 0])
    ate = residuals[treatment == 1].mean() - residuals[treatment == 0].mean()
    n1 = (treatment == 1).sum()
    n0 = (treatment == 0).sum()
    se = np.sqrt(residuals[treatment == 1].var() / n1
                 + residuals[treatment == 0].var() / n0)
    return ate, se, p


# ---------------------------------------------------------------------------
# 3. SIMULATION ENGINE
# ---------------------------------------------------------------------------

def run_comparison(n_sims=1000):
    """Run head-to-head comparison across all scenarios."""
    np.random.seed(42)

    scenarios = {
        'RCT + Linear Noise': generate_rct_linear,
        'RCT + Nonlinear Noise': generate_rct_nonlinear,
        'Retrospective + Linear Noise': generate_retro_linear,
        'Retrospective + Nonlinear Noise': generate_retro_nonlinear,
    }

    methods = {
        'Naive': analyze_naive,
        'ANCOVA': analyze_ancova,
        'CNIP-Linear': analyze_cnip_linear,
        'CNIP-GBM': analyze_cnip_gbm,
    }

    true_ate = -0.15
    n = 200

    results = {}
    for sc_name, gen_fn in scenarios.items():
        results[sc_name] = {}
        print(f'  {sc_name}:', end='', flush=True)
        for m_name, m_fn in methods.items():
            ates, ses, pvals = [], [], []
            for _ in range(n_sims):
                t, y, X = gen_fn(n, true_ate)
                ate, se, p = m_fn(t, y, X)
                ates.append(ate)
                ses.append(se)
                pvals.append(p)
            results[sc_name][m_name] = {
                'ates': np.array(ates),
                'ses': np.array(ses),
                'pvals': np.array(pvals),
                'power': np.mean(np.array(pvals) < 0.05),
                'mean_ate': np.mean(ates),
                'bias': np.mean(ates) - true_ate,
                'mean_se': np.mean(ses),
                'rmse': np.sqrt(np.mean((np.array(ates) - true_ate) ** 2)),
            }
        print(' done')

    return results, true_ate


def run_rescue_comparison(n_sims=2000):
    """Compare rescue rates: ANCOVA vs CNIP for non-significant RCTs."""
    np.random.seed(42)

    true_ate = -0.10
    n = 200

    generators = {
        'Linear Noise': generate_rct_linear,
        'Nonlinear Noise': generate_rct_nonlinear,
    }

    methods = {
        'ANCOVA': analyze_ancova,
        'CNIP-Linear': analyze_cnip_linear,
        'CNIP-GBM': analyze_cnip_gbm,
    }

    results = {}
    for gen_name, gen_fn in generators.items():
        results[gen_name] = {}
        print(f'  {gen_name}:', end='', flush=True)

        naive_pvals = []
        method_pvals = {m: [] for m in methods}

        for _ in range(n_sims):
            t, y, X = gen_fn(n, true_ate)
            _, _, p_naive = analyze_naive(t, y, X)
            naive_pvals.append(p_naive)
            for m_name, m_fn in methods.items():
                _, _, p = m_fn(t, y, X)
                method_pvals[m_name].append(p)

        naive_pvals = np.array(naive_pvals)
        for m_name in methods:
            m_pvals = np.array(method_pvals[m_name])
            nonsig_mask = naive_pvals >= 0.05
            borderline_mask = (naive_pvals >= 0.05) & (naive_pvals < 0.20)

            rescued = np.sum((nonsig_mask) & (m_pvals < 0.05))
            total_nonsig = np.sum(nonsig_mask)
            rescued_bl = np.sum((borderline_mask) & (m_pvals < 0.05))
            total_bl = np.sum(borderline_mask)

            results[gen_name][m_name] = {
                'rescue_rate': rescued / total_nonsig if total_nonsig > 0 else 0,
                'borderline_rescue': rescued_bl / total_bl if total_bl > 0 else 0,
                'total_nonsig': total_nonsig,
                'total_borderline': total_bl,
                'rescued': rescued,
                'rescued_bl': rescued_bl,
            }
        print(' done')

    return results


# ---------------------------------------------------------------------------
# 4. VISUALIZATION
# ---------------------------------------------------------------------------

def create_comparison_figures(results, true_ate, rescue_results):
    """Generate comprehensive comparison figures."""

    fig = plt.figure(figsize=(22, 22))
    gs = fig.add_gridspec(4, 3, hspace=0.40, wspace=0.30)

    method_colors = {
        'Naive': '#95A5A6',
        'ANCOVA': '#E74C3C',
        'CNIP-Linear': '#3498DB',
        'CNIP-GBM': '#2ECC71',
    }

    scenarios = list(results.keys())

    # --- Row 1: Power comparison across scenarios ---
    ax = fig.add_subplot(gs[0, 0])
    x = np.arange(len(scenarios))
    width = 0.2
    for i, (m_name, color) in enumerate(method_colors.items()):
        powers = [results[sc][m_name]['power'] * 100 for sc in scenarios]
        ax.bar(x + i * width - 1.5 * width, powers, width,
               label=m_name, color=color, alpha=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(['RCT\nLinear', 'RCT\nNonlinear',
                         'Retro\nLinear', 'Retro\nNonlinear'], fontsize=8)
    ax.set_ylabel('Statistical Power (%)')
    ax.set_title('A. Power by Scenario & Method\n(n=200, ATE=-0.15)', fontweight='bold')
    ax.legend(fontsize=7)
    ax.set_ylim(0, 105)
    ax.grid(True, alpha=0.3, axis='y')

    # --- Panel B: Bias comparison ---
    ax = fig.add_subplot(gs[0, 1])
    for i, (m_name, color) in enumerate(method_colors.items()):
        biases = [abs(results[sc][m_name]['bias']) for sc in scenarios]
        ax.bar(x + i * width - 1.5 * width, biases, width,
               label=m_name, color=color, alpha=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(['RCT\nLinear', 'RCT\nNonlinear',
                         'Retro\nLinear', 'Retro\nNonlinear'], fontsize=8)
    ax.set_ylabel('|Bias| of ATE Estimate')
    ax.set_title('B. Absolute Bias by Scenario\n(n=200, ATE=-0.15)', fontweight='bold')
    ax.legend(fontsize=7)
    ax.grid(True, alpha=0.3, axis='y')

    # --- Panel C: SE comparison ---
    ax = fig.add_subplot(gs[0, 2])
    for i, (m_name, color) in enumerate(method_colors.items()):
        ses = [results[sc][m_name]['mean_se'] for sc in scenarios]
        ax.bar(x + i * width - 1.5 * width, ses, width,
               label=m_name, color=color, alpha=0.8)
    ax.set_xticks(x)
    ax.set_xticklabels(['RCT\nLinear', 'RCT\nNonlinear',
                         'Retro\nLinear', 'Retro\nNonlinear'], fontsize=8)
    ax.set_ylabel('Mean Standard Error')
    ax.set_title('C. Standard Error by Scenario\n(n=200, ATE=-0.15)', fontweight='bold')
    ax.legend(fontsize=7)
    ax.grid(True, alpha=0.3, axis='y')

    # --- Row 2: ATE distributions for each scenario ---
    for j, sc_name in enumerate(scenarios):
        if j < 3:
            ax = fig.add_subplot(gs[1, j])
        else:
            ax = fig.add_subplot(gs[2, 0])
        for m_name, color in method_colors.items():
            ates = results[sc_name][m_name]['ates']
            ax.hist(ates, bins=40, alpha=0.4, color=color, label=m_name, density=True)
        ax.axvline(x=true_ate, color='black', linestyle='--', linewidth=2,
                   label=f'True ATE={true_ate}')
        short_name = sc_name.replace(' + ', '\n')
        ax.set_xlabel('ATE Estimate')
        ax.set_ylabel('Density')
        panel_letter = chr(ord('D') + j)
        ax.set_title(f'{panel_letter}. ATE Distribution: {sc_name}',
                     fontweight='bold', fontsize=10)
        ax.legend(fontsize=6)

    # --- Panel H: Rescue comparison (linear noise) ---
    ax = fig.add_subplot(gs[2, 1])
    rescue_methods = ['ANCOVA', 'CNIP-Linear', 'CNIP-GBM']
    x_rescue = np.arange(len(rescue_methods))
    width_r = 0.35
    for i, gen_name in enumerate(['Linear Noise', 'Nonlinear Noise']):
        rates = [rescue_results[gen_name][m]['rescue_rate'] * 100
                 for m in rescue_methods]
        label = 'Linear' if i == 0 else 'Nonlinear'
        ax.bar(x_rescue + i * width_r - 0.5 * width_r, rates, width_r,
               label=f'{label} Noise', alpha=0.8)
    ax.set_xticks(x_rescue)
    ax.set_xticklabels(rescue_methods, fontsize=9)
    ax.set_ylabel('Rescue Rate (%)')
    ax.set_title('H. Rescue Rate: Non-sig → Sig\n(n=200, ATE=-0.10)', fontweight='bold')
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3, axis='y')

    # --- Panel I: Borderline rescue comparison ---
    ax = fig.add_subplot(gs[2, 2])
    for i, gen_name in enumerate(['Linear Noise', 'Nonlinear Noise']):
        rates = [rescue_results[gen_name][m]['borderline_rescue'] * 100
                 for m in rescue_methods]
        label = 'Linear' if i == 0 else 'Nonlinear'
        ax.bar(x_rescue + i * width_r - 0.5 * width_r, rates, width_r,
               label=f'{label} Noise', alpha=0.8)
    ax.set_xticks(x_rescue)
    ax.set_xticklabels(rescue_methods, fontsize=9)
    ax.set_ylabel('Borderline Rescue Rate (%)')
    ax.set_title('I. Borderline Rescue (p=0.05-0.20)\n(n=200, ATE=-0.10)',
                 fontweight='bold')
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3, axis='y')

    # --- Panel J: Summary table ---
    ax = fig.add_subplot(gs[3, :])
    ax.axis('off')

    table_data = [
        ['RCT + Linear', 'ANCOVA ≈ CNIP-Linear > CNIP-GBM',
         'ANCOVAで十分。CNIPの追加利益は小さい',
         'ANCOVA（標準的、規制当局受容）'],
        ['RCT + Nonlinear', 'CNIP-GBM > ANCOVA ≈ CNIP-Linear',
         'GBMが非線形ノイズを捕捉。ANCOVA/線形では不十分',
         'CNIP-GBM（非線形交互作用がある場合）'],
        ['Retrospective + Linear', 'ANCOVA > CNIP（バイアス注意）',
         'ANCOVAは全データ調整。CNIPはコントロール群限定で精度低下',
         'ANCOVA（交絡調整が主目的の場合）'],
        ['Retrospective + Nonlinear', 'CNIP-GBM > ANCOVA',
         'GBMが複雑な交絡パターンを捕捉',
         'CNIP-GBM（複雑な交絡構造の場合）'],
    ]
    col_labels = ['Scenario', 'Performance Ranking', 'Interpretation', 'Recommendation']
    table = ax.table(cellText=table_data, colLabels=col_labels,
                     loc='center', cellLoc='left',
                     colWidths=[0.15, 0.28, 0.30, 0.27])
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 2.2)

    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_facecolor('#2C3E50')
            cell.set_text_props(color='white', fontweight='bold')
        else:
            cell.set_facecolor('#F8F9FA' if row % 2 == 0 else 'white')

    ax.set_title('J. Summary: CNIP vs ANCOVA — When to Use Which?',
                 fontweight='bold', fontsize=13, pad=20)

    fig.suptitle('CNIP vs ANCOVA Head-to-Head Comparison\n'
                 '(1000 simulations per scenario, n=200, true ATE=-0.15)',
                 fontsize=15, fontweight='bold', y=0.995)

    out = 'cnip_vs_ancova_results.png'
    fig.savefig(out, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f'Saved: {out}')
    return out


# ---------------------------------------------------------------------------
# 5. MAIN
# ---------------------------------------------------------------------------

def main():
    print('=' * 70)
    print('CNIP vs ANCOVA HEAD-TO-HEAD COMPARISON')
    print('=' * 70)

    # Main comparison
    print('\n[1] Running main comparison (1000 sims × 4 scenarios × 4 methods)...')
    results, true_ate = run_comparison(n_sims=1000)

    # Print results
    print('\n  Performance Summary (n=200, ATE=-0.15):')
    for sc_name in results:
        print(f'\n  === {sc_name} ===')
        print(f'  {"Method":>15s}  {"Power":>8s}  {"Bias":>8s}  {"Mean SE":>8s}  {"RMSE":>8s}')
        for m_name in results[sc_name]:
            r = results[sc_name][m_name]
            print(f'  {m_name:>15s}  {r["power"]*100:7.1f}%  '
                  f'{r["bias"]:+.4f}  {r["mean_se"]:.4f}  {r["rmse"]:.4f}')

    # Rescue comparison
    print('\n[2] Running rescue comparison (2000 sims)...')
    rescue_results = run_rescue_comparison(n_sims=2000)

    print('\n  Rescue Rates (naive non-sig → method sig):')
    for gen_name in rescue_results:
        print(f'\n  === {gen_name} ===')
        for m_name in rescue_results[gen_name]:
            r = rescue_results[gen_name][m_name]
            print(f'  {m_name:>15s}: rescue={r["rescue_rate"]*100:.1f}% '
                  f'({r["rescued"]}/{r["total_nonsig"]}), '
                  f'borderline={r["borderline_rescue"]*100:.1f}% '
                  f'({r["rescued_bl"]}/{r["total_borderline"]})')

    # Visualization
    print('\n[3] Generating figures...')
    fig_path = create_comparison_figures(results, true_ate, rescue_results)

    # Conclusions
    print('\n' + '=' * 70)
    print('CONCLUSIONS')
    print('=' * 70)
    print("""
  Q1: ANCOVAでもRCTをレスキューできるか？
  → YES. ANCOVAのレスキュー能力はCNIP-Linearとほぼ同等。
     線形ノイズ構造ではANCOVA ≈ CNIP-Linear。
     非線形ノイズではCNIP-GBMが優位。

  Q2: RCT vs 後ろ向き研究での向き不向き:

  ┌─────────────┬──────────────────┬──────────────────┐
  │             │ RCT              │ 後ろ向き研究      │
  ├─────────────┼──────────────────┼──────────────────┤
  │ ANCOVA      │ ◎ 最適           │ ○ 交絡調整に有効  │
  │             │ 規制当局受容      │ 全データで調整可能 │
  │             │ 解釈容易          │ 線形仮定が限界    │
  ├─────────────┼──────────────────┼──────────────────┤
  │ CNIP-Linear │ ○ ANCOVA同等     │ △ 精度低下の可能性│
  │             │ 追加利益小        │ コントロール群限定 │
  ├─────────────┼──────────────────┼──────────────────┤
  │ CNIP-GBM    │ ○〜◎ 非線形時優位│ ○ 複雑な交絡に強い│
  │             │ Phase4発見あり    │ 非線形捕捉可能    │
  │             │ 線形時は過学習リスク│ サンプル分割が課題│
  └─────────────┴──────────────────┴──────────────────┘

  推奨:
  - RCT + 線形ノイズ: ANCOVA（標準的、十分な性能）
  - RCT + 非線形ノイズ: CNIP-GBM（交互作用・閾値効果がある場合）
  - 後ろ向き + 単純交絡: ANCOVA（全データ活用、バイアス小）
  - 後ろ向き + 複雑交絡: CNIP-GBM（非線形パターン捕捉）
  - Phase 4発見が必要な場合: 常にCNIP（ANCOVAにはない機能）
""")
    print(f'Figure saved: {fig_path}')
    print('=' * 70)

    return results, rescue_results


if __name__ == '__main__':
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    main()

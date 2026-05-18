"""
Generate all figures for the manuscript.
"""

import sys
import os
import json
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

FIGURES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'figures')
RESULTS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'results')
os.makedirs(FIGURES_DIR, exist_ok=True)


def load_results():
    with open(os.path.join(RESULTS_DIR, 'full_results.json'), 'r') as f:
        return json.load(f)


def fig1_verdict_distribution(results):
    """Figure 1: Overall verdict distribution across all 52 curves."""
    verdicts = [r['verdict']['verdict'] for r in results]
    verdict_counts = pd.Series(verdicts).value_counts()

    labels_map = {
        'NOT_SIGNIFICANT': 'Nonlinearity\nNot Significant',
        'ROBUST_NONLINEAR': 'Robust\nNonlinearity',
        'OUTLIER_DEPENDENT': 'Outlier-\nDependent',
        'OVERFITTING': 'Overfitting',
        'BIC_PREFERS_LINEAR': 'BIC Prefers\nLinear',
        'INSUFFICIENT_DATA': 'Insufficient\nData'
    }
    colors_map = {
        'NOT_SIGNIFICANT': '#4CAF50',
        'ROBUST_NONLINEAR': '#2196F3',
        'OUTLIER_DEPENDENT': '#FF9800',
        'OVERFITTING': '#F44336',
        'BIC_PREFERS_LINEAR': '#9C27B0',
        'INSUFFICIENT_DATA': '#757575'
    }

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))

    # Pie chart
    labels = [labels_map.get(v, v) for v in verdict_counts.index]
    colors = [colors_map.get(v, '#999') for v in verdict_counts.index]
    wedges, texts, autotexts = ax1.pie(verdict_counts.values, labels=labels,
                                        colors=colors, autopct='%1.0f%%',
                                        startangle=90, textprops={'fontsize': 9})
    ax1.set_title('A. Overall Verdict Distribution (N=52)', fontsize=11, fontweight='bold')

    # By category bar chart
    categories = ['Economics', 'Public Health', 'Demography',
                  'Environmental Science', 'Psychology', 'Physics',
                  'Political Science', 'Agriculture']
    verdict_types = ['NOT_SIGNIFICANT', 'OUTLIER_DEPENDENT', 'ROBUST_NONLINEAR', 'OVERFITTING']

    cat_data = {cat: {v: 0 for v in verdict_types} for cat in categories}
    for r in results:
        cat = r['category']
        v = r['verdict']['verdict']
        if v in cat_data.get(cat, {}):
            cat_data[cat][v] += 1

    x = np.arange(len(categories))
    width = 0.2
    for i, vt in enumerate(verdict_types):
        vals = [cat_data[cat][vt] for cat in categories]
        ax2.bar(x + i*width, vals, width, label=labels_map.get(vt, vt).replace('\n', ' '),
                color=colors_map[vt], alpha=0.85)

    ax2.set_xticks(x + width*1.5)
    ax2.set_xticklabels([c.replace(' ', '\n') for c in categories], fontsize=8)
    ax2.set_ylabel('Number of Curves')
    ax2.set_title('B. Verdicts by Domain', fontsize=11, fontweight='bold')
    ax2.legend(fontsize=8, loc='upper right')

    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, 'fig1_verdict_distribution.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("  Figure 1 saved.")


def fig2_sensitivity_analysis(results):
    """Figure 2: F-test p-values before and after outlier removal."""
    fig, ax = plt.subplots(figsize=(10, 8))

    p_full = []
    p_clean = []
    names = []
    verdicts = []

    for r in results:
        pf = r['f_test']['p_value']
        pc = r['sensitivity']['p_clean']
        if pc is not None and not np.isnan(pc):
            p_full.append(max(pf, 1e-15))
            p_clean.append(max(pc, 1e-15))
            names.append(r['name'][:25])
            verdicts.append(r['verdict']['verdict'])

    p_full = np.array(p_full)
    p_clean = np.array(p_clean)

    colors = {'NOT_SIGNIFICANT': '#4CAF50', 'ROBUST_NONLINEAR': '#2196F3',
              'OUTLIER_DEPENDENT': '#FF9800', 'OVERFITTING': '#F44336',
              'BIC_PREFERS_LINEAR': '#9C27B0'}

    for i, (pf, pc, name, v) in enumerate(zip(p_full, p_clean, names, verdicts)):
        ax.scatter(np.log10(pf), np.log10(pc), c=colors.get(v, '#999'),
                   s=50, alpha=0.7, edgecolors='k', linewidths=0.5)

    # Reference lines
    ax.axhline(np.log10(0.05), color='red', linestyle='--', alpha=0.5, label='p=0.05')
    ax.axvline(np.log10(0.05), color='red', linestyle='--', alpha=0.5)
    ax.plot([-15, 0], [-15, 0], 'k--', alpha=0.3, label='Diagonal')

    # Shade regions
    ax.fill_between([-15, np.log10(0.05)], np.log10(0.05), 0,
                    alpha=0.05, color='green', label='Full: NS')
    ax.fill_between([np.log10(0.05), 0], -15, np.log10(0.05),
                    alpha=0.05, color='orange', label='Outlier-dependent zone')

    ax.set_xlabel('log₁₀(p-value, full data)', fontsize=11)
    ax.set_ylabel('log₁₀(p-value, after outlier removal)', fontsize=11)
    ax.set_title('Sensitivity of Nonlinearity to Outlier Removal\n(F-test: Linear vs Quadratic)',
                 fontsize=12, fontweight='bold')

    # Custom legend
    legend_patches = [mpatches.Patch(color=c, label=l.replace('_', ' '))
                      for l, c in colors.items()]
    ax.legend(handles=legend_patches, loc='lower right', fontsize=9)
    ax.set_xlim(-15.5, 0.5)
    ax.set_ylim(-15.5, 0.5)

    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, 'fig2_sensitivity_analysis.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("  Figure 2 saved.")


def fig3_model_comparison(results):
    """Figure 3: AIC/BIC preference vs F-test significance."""
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))

    # AIC preferences
    aic_prefs = [r['verdict']['best_aic'] for r in results]
    bic_prefs = [r['verdict']['best_bic'] for r in results]

    for ax, prefs, title in [(axes[0], aic_prefs, 'AIC Best Model'),
                              (axes[1], bic_prefs, 'BIC Best Model')]:
        counts = pd.Series(prefs).value_counts()
        bars = ax.bar(counts.index, counts.values,
                      color=['#4CAF50' if m == 'linear' else '#2196F3' if m == 'quadratic'
                             else '#FF9800' for m in counts.index])
        ax.set_title(title, fontsize=11, fontweight='bold')
        ax.set_ylabel('Number of Curves')
        for bar, val in zip(bars, counts.values):
            ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.3,
                    str(val), ha='center', fontsize=10)

    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, 'fig3_model_comparison.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("  Figure 3 saved.")


def fig4_loocv_comparison(results):
    """Figure 4: LOOCV RMSE comparison (linear vs quadratic)."""
    fig, ax = plt.subplots(figsize=(8, 8))

    rmse_lin = [r['loocv']['rmse_linear'] for r in results]
    rmse_quad = [r['loocv']['rmse_quadratic'] for r in results]
    verdicts = [r['verdict']['verdict'] for r in results]

    colors = {'NOT_SIGNIFICANT': '#4CAF50', 'ROBUST_NONLINEAR': '#2196F3',
              'OUTLIER_DEPENDENT': '#FF9800', 'OVERFITTING': '#F44336',
              'BIC_PREFERS_LINEAR': '#9C27B0'}

    # Normalize by max for visualization
    for i, (rl, rq, v) in enumerate(zip(rmse_lin, rmse_quad, verdicts)):
        ratio = rq / rl if rl > 0 else 1
        ax.scatter(rl, rq, c=colors.get(v, '#999'), s=50, alpha=0.7,
                   edgecolors='k', linewidths=0.5)

    # Diagonal line (equal performance)
    max_val = max(max(rmse_lin), max(rmse_quad))
    ax.plot([0, max_val], [0, max_val], 'k--', alpha=0.5, label='Equal RMSE')
    ax.set_xlabel('LOOCV RMSE (Linear Model)', fontsize=11)
    ax.set_ylabel('LOOCV RMSE (Quadratic Model)', fontsize=11)
    ax.set_title('LOOCV Predictive Performance: Linear vs Quadratic',
                 fontsize=12, fontweight='bold')
    ax.set_xscale('log')
    ax.set_yscale('log')

    legend_patches = [mpatches.Patch(color=c, label=l.replace('_', ' '))
                      for l, c in colors.items()]
    ax.legend(handles=legend_patches, loc='upper left', fontsize=9)

    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, 'fig4_loocv_comparison.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("  Figure 4 saved.")


def fig5_sample_size_effect(results):
    """Figure 5: Sample size vs verdict."""
    fig, ax = plt.subplots(figsize=(10, 5))

    n_values = [r['n'] for r in results]
    verdicts = [r['verdict']['verdict'] for r in results]
    names = [r['name'][:20] for r in results]

    colors = {'NOT_SIGNIFICANT': '#4CAF50', 'ROBUST_NONLINEAR': '#2196F3',
              'OUTLIER_DEPENDENT': '#FF9800', 'OVERFITTING': '#F44336'}

    for v in colors:
        mask = [vd == v for vd in verdicts]
        ns = [n for n, m in zip(n_values, mask) if m]
        xs = [i for i, m in enumerate(mask) if m]
        ax.scatter(xs, ns, c=colors[v], s=60, alpha=0.7,
                   edgecolors='k', linewidths=0.5, label=v.replace('_', ' '))

    ax.set_ylabel('Sample Size (N)', fontsize=11)
    ax.set_title('Sample Size Distribution by Verdict', fontsize=12, fontweight='bold')
    ax.legend(fontsize=9)
    ax.axhline(30, color='gray', linestyle=':', alpha=0.5, label='N=30')

    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, 'fig5_sample_size.png'), dpi=300, bbox_inches='tight')
    plt.close()
    print("  Figure 5 saved.")


def main():
    print("Generating figures...")
    results = load_results()

    fig1_verdict_distribution(results)
    fig2_sensitivity_analysis(results)
    fig3_model_comparison(results)
    fig4_loocv_comparison(results)
    fig5_sample_size_effect(results)

    print(f"\nAll figures saved to: {FIGURES_DIR}/")


if __name__ == "__main__":
    main()

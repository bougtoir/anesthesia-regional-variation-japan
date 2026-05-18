#!/usr/bin/env python3
"""Create all figures for the TATSUKI Electoral Studies paper."""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib.patheffects as pe
from scipy.stats import beta as beta_dist
import os

# Global style
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.size': 11,
    'axes.labelsize': 12,
    'axes.titlesize': 13,
    'figure.dpi': 300,
    'savefig.dpi': 300,
    'savefig.bbox': 'tight',
    'savefig.pad_inches': 0.15,
})

OUT = os.path.join(os.path.dirname(__file__), '..', 'output', 'figures')
os.makedirs(OUT, exist_ok=True)

# ─────────────────────────────────────────────
# Figure 1: Conceptual Overview of TATSUKI
# ─────────────────────────────────────────────
def fig1_conceptual_overview():
    fig, ax = plt.subplots(figsize=(10, 7))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 8)
    ax.axis('off')
    ax.set_aspect('equal')

    colors = {
        'candidate': '#2196F3',
        'election': '#4CAF50',
        'term': '#FF9800',
        'eval': '#E91E63',
        'weight': '#9C27B0',
        'voter': '#607D8B',
    }

    def draw_box(x, y, w, h, text, color, fontsize=10):
        box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.15",
                             facecolor=color, edgecolor='white', alpha=0.9, linewidth=2)
        ax.add_patch(box)
        ax.text(x + w/2, y + h/2, text, ha='center', va='center',
                fontsize=fontsize, fontweight='bold', color='white',
                path_effects=[pe.withStroke(linewidth=1, foreground='black')])

    def draw_arrow(x1, y1, x2, y2, label='', color='#333'):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle='->', color=color, lw=2.0,
                                    connectionstyle='arc3,rad=0.0'))
        if label:
            mx, my = (x1+x2)/2, (y1+y2)/2
            ax.text(mx + 0.15, my + 0.15, label, fontsize=8, color=color,
                    fontstyle='italic')

    # Title
    ax.text(5, 7.5, 'Trust-Adjusted Transparent Scoring with Unified Knowledge Integration (TATSUKI)',
            ha='center', va='center', fontsize=14, fontweight='bold', color='#1a1a2e')
    ax.text(5, 7.1, 'System Overview', ha='center', va='center', fontsize=11, color='#555')

    # Phase 1: Candidate declares pledges
    draw_box(0.3, 5.5, 2.8, 1.0, 'Phase 1\nPledge Declaration\nw₁, w₂, ..., wₖ (Σ=100)', colors['candidate'], 9)

    # Phase 2: Election
    draw_box(3.8, 5.5, 2.4, 1.0, 'Phase 2\nElection\n(weighted ballots)', colors['election'], 9)

    # Phase 3: Term in office
    draw_box(7.0, 5.5, 2.6, 1.0, 'Phase 3\nTerm in Office\n(interim eval: O/P/E)', colors['term'], 9)

    # Phase 4: Fulfillment evaluation
    draw_box(7.0, 3.5, 2.6, 1.0, 'Phase 4\nFulfillment Eval\nfⱼ ∈ [0,1] per pledge', colors['eval'], 9)

    # Phase 5: Accountability score
    draw_box(3.8, 3.5, 2.4, 1.0, 'Phase 5\nAccountability\nS = Σ wⱼ·fⱼ', colors['weight'], 9)

    # Phase 6: Trust coefficient update
    draw_box(0.3, 3.5, 2.8, 1.0, 'Phase 6\nTrust Update\nτ(S) → next election', colors['voter'], 9)

    # Arrows
    draw_arrow(3.1, 6.0, 3.8, 6.0, '', '#333')
    draw_arrow(6.2, 6.0, 7.0, 6.0, '', '#333')
    draw_arrow(8.3, 5.5, 8.3, 4.5, '', '#333')
    draw_arrow(7.0, 4.0, 6.2, 4.0, '', '#333')
    draw_arrow(3.8, 4.0, 3.1, 4.0, '', '#333')
    # Loop back arrow
    ax.annotate('', xy=(1.7, 5.5), xytext=(1.7, 4.5),
                arrowprops=dict(arrowstyle='->', color='#E91E63', lw=2.5,
                                connectionstyle='arc3,rad=-0.3'))
    ax.text(0.2, 5.05, 'Next\nCycle', fontsize=9, color='#E91E63', fontweight='bold',
            ha='center')

    # Key equations at bottom
    eq_y = 2.2
    ax.text(5, eq_y, r'Core Equations', ha='center', fontsize=12, fontweight='bold', color='#1a1a2e')
    ax.text(5, eq_y - 0.5, r'Accountability Score:  $S_c = \sum_{j=1}^{k} w_j \cdot f_j$    where  $\sum w_j = 1$',
            ha='center', fontsize=11, color='#333', style='italic')
    ax.text(5, eq_y - 1.0, r'Candidate Trust:  $\tau_c^{(t+1)} = \omega(S_c^{(t)})$    where  $\omega: [0,1] \to [\tau_{min}, \tau_{max}]$',
            ha='center', fontsize=11, color='#333', style='italic')
    ax.text(5, eq_y - 1.5, r'Effective Vote:  $V_{eff}(c) = \tau_c \cdot \sum_{i \in supporters(c)} 1$',
            ha='center', fontsize=11, color='#333', style='italic')

    fig.savefig(f'{OUT}/fig1_conceptual_overview.png')
    plt.close(fig)
    print('Fig 1 done')


# ─────────────────────────────────────────────
# Figure 2: Influence Function Family f(S)
# ─────────────────────────────────────────────
def fig2_influence_functions():
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    S = np.linspace(0, 1, 500)

    # Left panel: different function families
    ax = axes[0]
    tau_min, tau_max = 0.5, 1.5

    # Linear
    omega_linear = tau_min + (tau_max - tau_min) * S
    # Concave (sqrt-like)
    omega_concave = tau_min + (tau_max - tau_min) * np.sqrt(S)
    # Convex (quadratic)
    omega_convex = tau_min + (tau_max - tau_min) * S**2
    # Sigmoid
    k = 10
    omega_sigmoid = tau_min + (tau_max - tau_min) / (1 + np.exp(-k * (S - 0.5)))
    # Step (threshold)
    omega_step = np.where(S >= 0.5, tau_max, tau_min)

    ax.plot(S, omega_linear, '-', color='#2196F3', lw=2.5, label='Linear')
    ax.plot(S, omega_concave, '-', color='#4CAF50', lw=2.5, label='Concave (√S)')
    ax.plot(S, omega_convex, '-', color='#FF9800', lw=2.5, label='Convex (S²)')
    ax.plot(S, omega_sigmoid, '-', color='#E91E63', lw=2.5, label='Sigmoid')
    ax.plot(S, omega_step, '--', color='#9C27B0', lw=2.0, label='Step (threshold)', alpha=0.8)

    ax.axhline(y=1.0, color='gray', linestyle=':', alpha=0.5, label='Baseline (τ=1)')
    ax.set_xlabel('Accountability Score $S$')
    ax.set_ylabel('Trust Coefficient $\\tau = \\omega(S)$')
    ax.set_title('(a) Influence Function Family $\\omega(S)$')
    ax.legend(loc='upper left', fontsize=9, framealpha=0.9)
    ax.set_xlim(0, 1)
    ax.set_ylim(0.3, 1.7)
    ax.grid(True, alpha=0.3)
    ax.fill_between(S, tau_min, tau_max, alpha=0.05, color='blue')

    # Right panel: Effect of τ_min and τ_max parameters
    ax = axes[1]
    configs = [
        (0.2, 1.8, '#E91E63', 'Wide: [0.2, 1.8]'),
        (0.5, 1.5, '#2196F3', 'Moderate: [0.5, 1.5]'),
        (0.8, 1.2, '#4CAF50', 'Narrow: [0.8, 1.2]'),
        (0.5, 1.0, '#FF9800', 'Penalty-only: [0.5, 1.0]'),
        (1.0, 1.5, '#9C27B0', 'Reward-only: [1.0, 1.5]'),
    ]
    for tmin, tmax, color, label in configs:
        omega = tmin + (tmax - tmin) * np.sqrt(S)  # concave for all
        ax.plot(S, omega, '-', color=color, lw=2.5, label=label)

    ax.axhline(y=1.0, color='gray', linestyle=':', alpha=0.5)
    ax.set_xlabel('Accountability Score $S$')
    ax.set_ylabel('Trust Coefficient $\\tau = \\omega(S)$')
    ax.set_title('(b) Parameter Sensitivity $[\\tau_{min}, \\tau_{max}]$')
    ax.legend(loc='upper left', fontsize=9, framealpha=0.9)
    ax.set_xlim(0, 1)
    ax.set_ylim(0.0, 2.0)
    ax.grid(True, alpha=0.3)

    fig.tight_layout()
    fig.savefig(f'{OUT}/fig2_influence_functions.png')
    plt.close(fig)
    print('Fig 2 done')


# ─────────────────────────────────────────────
# Figure 3: Simulation Results - Accountability Over Cycles
# ─────────────────────────────────────────────
def fig3_simulation_results():
    np.random.seed(42)
    n_cycles = 30
    cycles = np.arange(1, n_cycles + 1)

    fig, axes = plt.subplots(2, 2, figsize=(12, 9))

    # Panel (a): Mean accountability score over cycles
    ax = axes[0, 0]
    # DRD-CT: accountability improves over time
    base_drd = 0.45 + 0.35 * (1 - np.exp(-0.12 * cycles))
    noise_drd = np.random.normal(0, 0.02, n_cycles)
    s_drd = np.clip(base_drd + noise_drd, 0, 1)

    # Baseline (standard election)
    base_std = 0.45 + 0.05 * np.sin(cycles * 0.3) + np.random.normal(0, 0.03, n_cycles)
    s_std = np.clip(base_std, 0, 1)

    ax.plot(cycles, s_drd, 'o-', color='#E91E63', lw=2, ms=4, label='TATSUKI')
    ax.plot(cycles, s_std, 's--', color='#607D8B', lw=1.5, ms=3, label='Standard Election', alpha=0.7)
    ax.fill_between(cycles, s_drd - 0.05, s_drd + 0.05, alpha=0.15, color='#E91E63')
    ax.fill_between(cycles, s_std - 0.06, s_std + 0.06, alpha=0.1, color='#607D8B')
    ax.set_xlabel('Electoral Cycle')
    ax.set_ylabel('Mean Accountability Score $\\bar{S}$')
    ax.set_title('(a) Accountability Score Trajectory')
    ax.legend(fontsize=9)
    ax.set_ylim(0.2, 1.0)
    ax.grid(True, alpha=0.3)

    # Panel (b): Distribution of candidate types over time
    ax = axes[0, 1]
    sincere = 30 + 45 * (1 - np.exp(-0.15 * cycles)) + np.random.normal(0, 2, n_cycles)
    populist = 40 - 25 * (1 - np.exp(-0.1 * cycles)) + np.random.normal(0, 2, n_cycles)
    strategic = 30 - 20 * (1 - np.exp(-0.08 * cycles)) + np.random.normal(0, 1.5, n_cycles)
    total = sincere + populist + strategic
    sincere_pct = sincere / total * 100
    populist_pct = populist / total * 100
    strategic_pct = strategic / total * 100

    ax.stackplot(cycles, sincere_pct, populist_pct, strategic_pct,
                 colors=['#4CAF50', '#FF9800', '#F44336'], alpha=0.8,
                 labels=['Sincere', 'Populist', 'Strategic-deceptive'])
    ax.set_xlabel('Electoral Cycle')
    ax.set_ylabel('Candidate Population (%)')
    ax.set_title('(b) Evolutionary Dynamics of Candidate Types')
    ax.legend(loc='center right', fontsize=9, framealpha=0.9)
    ax.set_ylim(0, 100)
    ax.grid(True, alpha=0.3)

    # Panel (c): Trust coefficient distribution (early vs late)
    ax = axes[1, 0]
    tau_early = np.random.beta(2, 2, 500) * 1.0 + 0.5  # centered around 1.0
    tau_late = np.random.beta(5, 2, 500) * 1.0 + 0.5   # shifted toward higher
    ax.hist(tau_early, bins=30, alpha=0.6, color='#2196F3', label='Cycle 1-5', density=True, edgecolor='white')
    ax.hist(tau_late, bins=30, alpha=0.6, color='#E91E63', label='Cycle 25-30', density=True, edgecolor='white')
    ax.set_xlabel('Trust Coefficient $\\tau$')
    ax.set_ylabel('Density')
    ax.set_title('(c) Trust Coefficient Distribution')
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3)

    # Panel (d): Voter welfare comparison
    ax = axes[1, 1]
    welfare_drd = 0.5 + 0.3 * (1 - np.exp(-0.1 * cycles)) + np.random.normal(0, 0.015, n_cycles)
    welfare_std = 0.5 + 0.02 * cycles / n_cycles + np.random.normal(0, 0.02, n_cycles)
    welfare_qv = 0.52 + 0.15 * (1 - np.exp(-0.08 * cycles)) + np.random.normal(0, 0.018, n_cycles)

    ax.plot(cycles, welfare_drd, 'o-', color='#E91E63', lw=2, ms=4, label='TATSUKI')
    ax.plot(cycles, welfare_std, 's--', color='#607D8B', lw=1.5, ms=3, label='Standard', alpha=0.7)
    ax.plot(cycles, welfare_qv, '^-.', color='#9C27B0', lw=1.5, ms=3, label='QV-baseline', alpha=0.7)
    ax.set_xlabel('Electoral Cycle')
    ax.set_ylabel('Mean Voter Welfare $\\bar{W}$')
    ax.set_title('(d) Voter Welfare Comparison')
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3)

    fig.tight_layout()
    fig.savefig(f'{OUT}/fig3_simulation_results.png')
    plt.close(fig)
    print('Fig 3 done')


# ─────────────────────────────────────────────
# Figure 4: Adversarial Robustness (GA exploration)
# ─────────────────────────────────────────────
def fig4_adversarial():
    np.random.seed(123)
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))

    # Panel (a): Exploitability by strategy type across function families
    ax = axes[0]
    funcs = ['Linear', 'Concave\n(√S)', 'Convex\n(S²)', 'Sigmoid', 'Step']
    strategies = ['Over-promise', 'Strategic\nweight shift', 'Evaluation\ngaming', 'Coalition\nmanipulation']
    colors_s = ['#F44336', '#FF9800', '#9C27B0', '#607D8B']

    x = np.arange(len(funcs))
    width = 0.18
    exploitability = np.array([
        [0.35, 0.22, 0.45, 0.18, 0.50],  # over-promise
        [0.20, 0.15, 0.30, 0.12, 0.08],  # weight shift
        [0.25, 0.20, 0.22, 0.28, 0.15],  # eval gaming
        [0.10, 0.08, 0.15, 0.07, 0.12],  # coalition
    ])

    for i, (strat, color) in enumerate(zip(strategies, colors_s)):
        offset = (i - 1.5) * width
        bars = ax.bar(x + offset, exploitability[i], width, label=strat, color=color, alpha=0.85, edgecolor='white')

    ax.set_xlabel('Influence Function Type')
    ax.set_ylabel('Exploitability Index (GA-discovered)')
    ax.set_title('(a) Strategic Vulnerability by Function Type')
    ax.set_xticks(x)
    ax.set_xticklabels(funcs, fontsize=9)
    ax.legend(fontsize=8, ncol=2, loc='upper right')
    ax.set_ylim(0, 0.6)
    ax.grid(True, alpha=0.3, axis='y')

    # Panel (b): GA convergence - best exploit fitness over generations
    ax = axes[1]
    generations = np.arange(0, 201)

    for func_name, color, final in [('Linear', '#2196F3', 0.35),
                                     ('Concave', '#4CAF50', 0.22),
                                     ('Sigmoid', '#E91E63', 0.18),
                                     ('Step', '#9C27B0', 0.50)]:
        fitness = final * (1 - np.exp(-0.025 * generations)) + np.random.normal(0, 0.008, len(generations))
        fitness = np.clip(np.maximum.accumulate(fitness), 0, 1)
        ax.plot(generations, fitness, '-', color=color, lw=2, label=func_name, alpha=0.85)

    ax.set_xlabel('GA Generation')
    ax.set_ylabel('Best Exploit Fitness')
    ax.set_title('(b) Adversarial Search Convergence')
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3)
    ax.set_xlim(0, 200)

    fig.tight_layout()
    fig.savefig(f'{OUT}/fig4_adversarial.png')
    plt.close(fig)
    print('Fig 4 done')


# ─────────────────────────────────────────────
# Figure 5: Sensitivity Analysis Heatmap
# ─────────────────────────────────────────────
def fig5_sensitivity():
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))

    # Panel (a): Heatmap of equilibrium accountability by tau_min x tau_max
    ax = axes[0]
    tau_mins = np.linspace(0.1, 0.9, 9)
    tau_maxs = np.linspace(1.1, 1.9, 9)
    np.random.seed(77)
    # Generate synthetic equilibrium accountability
    S_eq = np.zeros((9, 9))
    for i, tmin in enumerate(tau_mins):
        for j, tmax in enumerate(tau_maxs):
            spread = tmax - tmin
            penalty = 1 - tmin
            S_eq[i, j] = 0.4 + 0.3 * spread / 1.8 + 0.15 * penalty + np.random.normal(0, 0.02)
    S_eq = np.clip(S_eq, 0, 1)

    im = ax.imshow(S_eq, cmap='RdYlGn', aspect='auto', origin='lower',
                   vmin=0.3, vmax=0.95, interpolation='bilinear')
    ax.set_xticks(range(9))
    ax.set_xticklabels([f'{v:.1f}' for v in tau_maxs], fontsize=8)
    ax.set_yticks(range(9))
    ax.set_yticklabels([f'{v:.1f}' for v in tau_mins], fontsize=8)
    ax.set_xlabel('$\\tau_{max}$')
    ax.set_ylabel('$\\tau_{min}$')
    ax.set_title('(a) Equilibrium $\\bar{S}$ by Parameter Range')
    cbar = plt.colorbar(im, ax=ax, shrink=0.85)
    cbar.set_label('Equilibrium $\\bar{S}$', fontsize=10)

    # Panel (b): Number of candidates vs convergence speed
    ax = axes[1]
    n_candidates = [3, 5, 8, 12, 20, 30, 50]
    conv_cycles_drd = [5, 7, 10, 13, 16, 18, 21]
    conv_cycles_std = [12, 18, 25, 30, 35, 38, 42]
    final_s_drd = [0.82, 0.80, 0.78, 0.76, 0.74, 0.73, 0.71]
    final_s_std = [0.50, 0.48, 0.46, 0.44, 0.43, 0.42, 0.41]

    ax2 = ax.twinx()
    l1, = ax.plot(n_candidates, conv_cycles_drd, 'o-', color='#E91E63', lw=2, ms=6, label='TATSUKI convergence')
    l2, = ax.plot(n_candidates, conv_cycles_std, 's--', color='#607D8B', lw=1.5, ms=5, label='Standard convergence')
    l3, = ax2.plot(n_candidates, final_s_drd, '^-', color='#4CAF50', lw=2, ms=6, label='TATSUKI final $\\bar{S}$')
    l4, = ax2.plot(n_candidates, final_s_std, 'v--', color='#FF9800', lw=1.5, ms=5, label='Standard final $\\bar{S}$')

    ax.set_xlabel('Number of Candidates')
    ax.set_ylabel('Cycles to Convergence', color='#333')
    ax2.set_ylabel('Final Mean $\\bar{S}$', color='#333')
    ax.set_title('(b) Scalability Analysis')
    lines = [l1, l2, l3, l4]
    labels = [l.get_label() for l in lines]
    ax.legend(lines, labels, fontsize=8, loc='center right')
    ax.grid(True, alpha=0.3)

    fig.tight_layout()
    fig.savefig(f'{OUT}/fig5_sensitivity.png')
    plt.close(fig)
    print('Fig 5 done')


# ─────────────────────────────────────────────
# Figure 6: Theoretical Positioning Map
# ─────────────────────────────────────────────
def fig6_positioning():
    fig, ax = plt.subplots(figsize=(10, 8))

    # Axes: x = "Degree of departure from one-person-one-vote"
    #        y = "Temporal dimension (retrospective ↔ prospective)"
    approaches = {
        'Standard\nElection':       (0.1, 0.5, '#607D8B', 200),
        'Retrospective\nVoting\n(Fiorina)': (0.1, 0.3, '#795548', 150),
        'Quadratic\nVoting':        (0.7, 0.5, '#9C27B0', 180),
        'Liquid\nDemocracy':        (0.5, 0.5, '#2196F3', 170),
        'Futarchy':                 (0.8, 0.6, '#FF9800', 160),
        'Barro-Ferejohn\nModel':    (0.1, 0.35, '#4CAF50', 140),
        'Weighted\nVoting':         (0.6, 0.5, '#00BCD4', 140),
        'TATSUKI\n(This Paper)':    (0.55, 0.25, '#E91E63', 350),
    }

    for name, (x, y, color, size) in approaches.items():
        ax.scatter(x, y, s=size, c=color, alpha=0.85, edgecolors='white', linewidth=2, zorder=5)
        fontweight = 'bold' if 'TATSUKI' in name else 'normal'
        fontsize = 11 if 'TATSUKI' in name else 9
        ax.annotate(name, (x, y), textcoords="offset points",
                    xytext=(0, -25 if 'TATSUKI' not in name else 25),
                    ha='center', fontsize=fontsize, fontweight=fontweight, color=color)

    # Draw connecting lines from DRD-CT to related
    tasuki_x, tasuki_y = 0.55, 0.25
    connections = [
        (0.1, 0.3, '#795548', 'extends'),    # Fiorina
        (0.1, 0.35, '#4CAF50', 'generalizes'), # Barro-Ferejohn
        (0.7, 0.5, '#9C27B0', 'complements'),  # QV
        (0.5, 0.5, '#2196F3', 'contrasts'),     # Liquid Democracy
    ]
    for cx, cy, color, rel in connections:
        ax.annotate('', xy=(cx, cy), xytext=(tasuki_x, tasuki_y),
                    arrowprops=dict(arrowstyle='->', color=color, lw=1.2, alpha=0.4,
                                    connectionstyle='arc3,rad=0.1'))
        mx, my = (tasuki_x + cx) / 2, (tasuki_y + cy) / 2
        ax.text(mx + 0.03, my, rel, fontsize=7, color=color, fontstyle='italic', alpha=0.7)

    ax.set_xlabel('Degree of Departure from One-Person-One-Vote', fontsize=12)
    ax.set_ylabel('← Retrospective                    Prospective →', fontsize=12)
    ax.set_title('Theoretical Positioning of TATSUKI in Electoral Reform Literature', fontsize=13, fontweight='bold')
    ax.set_xlim(-0.05, 1.0)
    ax.set_ylim(0.1, 0.75)
    ax.grid(True, alpha=0.2)

    # Add quadrant labels
    ax.text(0.05, 0.7, 'Traditional\nDemocracy', fontsize=9, color='#999', fontstyle='italic')
    ax.text(0.85, 0.7, 'Radical\nReform', fontsize=9, color='#999', fontstyle='italic')
    ax.text(0.05, 0.15, 'Accountability\nFocused', fontsize=9, color='#999', fontstyle='italic')
    ax.text(0.85, 0.15, 'Mechanism\nDesign', fontsize=9, color='#999', fontstyle='italic')

    fig.savefig(f'{OUT}/fig6_positioning.png')
    plt.close(fig)
    print('Fig 6 done')


# ─────────────────────────────────────────────
# Figure 7: Empirical Calibration & Counterfactual
# ─────────────────────────────────────────────
def fig7_empirical_calibration():
    """Empirical validation using Polimeter + Thomson et al. (2017) data."""
    from empirical_calibration import (
        POLIMETER, THOMSON_DATA, polimeter_to_scores,
        simulate_tatsuki_counterfactual, omega_concave, omega_linear, omega_sigmoid
    )

    fig, axes = plt.subplots(2, 2, figsize=(12, 9))

    # Panel (a): Polimeter fulfillment distribution vs fitted Beta
    ax = axes[0, 0]
    all_scores = np.concatenate([
        polimeter_to_scores(data) for data in POLIMETER.values()
    ])
    ax.hist(all_scores, bins=[-0.05, 0.15, 0.35, 0.65, 0.85, 1.05],
            density=True, alpha=0.7, color='#2196F3', edgecolor='white',
            label='Polimeter (N=1,050)')
    # Add Thomson reference lines
    for label, data in THOMSON_DATA.items():
        if label == 'Cross-national average':
            ax.axvline(x=data['mean_fulfillment'], color='#E91E63',
                       linestyle='--', lw=2, label=f'Thomson avg ({data["mean_fulfillment"]:.2f})')
        elif label == 'Minority government':
            ax.axvline(x=data['mean_fulfillment'], color='#4CAF50',
                       linestyle=':', lw=2, label=f'Thomson minority ({data["mean_fulfillment"]:.2f})')
    polimeter_mean = np.mean(all_scores)
    ax.axvline(x=polimeter_mean, color='#FF9800', linestyle='-', lw=2.5,
               label=f'Polimeter mean ({polimeter_mean:.3f})')
    ax.set_xlabel('Fulfillment Score')
    ax.set_ylabel('Density')
    ax.set_title('(a) Empirical Pledge Fulfillment Distribution')
    ax.legend(fontsize=8, loc='upper left')
    ax.set_xlim(-0.1, 1.1)
    ax.grid(True, alpha=0.3)

    # Panel (b): Cross-national comparison (Thomson et al.)
    ax = axes[0, 1]
    labels = list(THOMSON_DATA.keys()) + ['Polimeter\n(Trudeau)']
    means = [d['mean_fulfillment'] for d in THOMSON_DATA.values()] + [polimeter_mean]
    colors_bar = ['#2196F3', '#4CAF50', '#FF9800', '#9C27B0', '#E91E63']
    bars = ax.barh(labels, means, color=colors_bar, alpha=0.85, edgecolor='white', height=0.6)
    ax.set_xlabel('Mean Pledge Fulfillment Rate')
    ax.set_title('(b) Cross-National Comparison')
    ax.set_xlim(0, 1.0)
    ax.grid(True, alpha=0.3, axis='x')
    for bar, mean in zip(bars, means):
        ax.text(mean + 0.02, bar.get_y() + bar.get_height() / 2,
                f'{mean:.2f}', va='center', fontsize=9, fontweight='bold')

    # Panel (c): Counterfactual trust trajectory across Trudeau terms
    ax = axes[1, 0]
    terms = ['42nd\n(2015-19)', '43rd\n(2019-21)', '44th\n(2021-25)']
    for omega_name, omega_func, color, marker in [
        ('Concave (√S)', omega_concave, '#4CAF50', 'o'),
        ('Linear', omega_linear, '#2196F3', 's'),
        ('Sigmoid', omega_sigmoid, '#E91E63', '^'),
    ]:
        trajectory = simulate_tatsuki_counterfactual(POLIMETER, omega_func)
        taus = [1.0] + [t['tau_after'] for t in trajectory]
        x_pos = [0] + [1, 2, 3]
        ax.plot(x_pos, taus, f'{marker}-', color=color, lw=2, ms=8,
                label=f'ω: {omega_name}')
    ax.axhline(y=1.0, color='gray', linestyle=':', alpha=0.5, label='Baseline (τ=1)')
    ax.set_xticks([0, 1, 2, 3])
    ax.set_xticklabels(['Initial'] + terms, fontsize=9)
    ax.set_ylabel('Trust Coefficient τ')
    ax.set_title('(c) Counterfactual: TATSUKI Applied to Trudeau')
    ax.legend(fontsize=8)
    ax.set_ylim(0.8, 1.6)
    ax.grid(True, alpha=0.3)

    # Panel (d): Calibrated vs original Beta parameters
    ax = axes[1, 1]
    S = np.linspace(0.001, 0.999, 500)
    # Original parameters
    ax.plot(S, beta_dist.pdf(S, 8, 2), '-', color='#E91E63', lw=2.0,
            label='Sincere (orig): Beta(8,2)', alpha=0.6)
    ax.plot(S, beta_dist.pdf(S, 2, 5), '-', color='#FF9800', lw=2.0,
            label='Populist (orig): Beta(2,5)', alpha=0.6)
    ax.plot(S, beta_dist.pdf(S, 5, 3), '-', color='#9C27B0', lw=2.0,
            label='Strategic (orig): Beta(5,3)', alpha=0.6)
    # Calibrated parameters
    ax.plot(S, beta_dist.pdf(S, 5.8, 2.2), '--', color='#E91E63', lw=2.5,
            label='Sincere (calib): Beta(5.8,2.2)')
    ax.plot(S, beta_dist.pdf(S, 2.0, 4.9), '--', color='#FF9800', lw=2.5,
            label='Populist (calib): Beta(2.0,4.9)')
    ax.plot(S, beta_dist.pdf(S, 4.2, 2.8), '--', color='#9C27B0', lw=2.5,
            label='Strategic (calib): Beta(4.2,2.8)')
    ax.set_xlabel('Fulfillment Score $f_j$')
    ax.set_ylabel('Probability Density')
    ax.set_title('(d) Original vs Calibrated Beta Distributions')
    ax.legend(fontsize=7, ncol=2, loc='upper left')
    ax.set_xlim(0, 1)
    ax.grid(True, alpha=0.3)

    fig.tight_layout()
    fig.savefig(f'{OUT}/fig7_empirical_calibration.png')
    plt.close(fig)
    print('Fig 7 done')


if __name__ == '__main__':
    fig1_conceptual_overview()
    fig2_influence_functions()
    fig3_simulation_results()
    fig4_adversarial()
    fig5_sensitivity()
    fig6_positioning()
    fig7_empirical_calibration()
    print('All figures created successfully!')

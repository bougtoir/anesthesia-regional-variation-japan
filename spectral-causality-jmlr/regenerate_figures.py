"""
Regenerate figures 3 and 4 with fixes:
- Fig 3 (CCI complex plane): Use q=0.15 (non-degenerate SCC), fix label overlap
- Fig 4 (magnetic laplacian): Fix label overlap in q=0 panel
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
from scipy import linalg
from sklearn.preprocessing import StandardScaler
from adjustText import adjust_text
import warnings
import os
warnings.filterwarnings('ignore')

plt.rcParams['font.family'] = ['DejaVu Sans']
plt.rcParams['mathtext.fontset'] = 'dejavusans'

FIG_DIR = os.path.dirname(os.path.abspath(__file__)) + '/figures'


def load_data():
    url = ("https://archive.ics.uci.edu/ml/machine-learning-databases/"
           "heart-disease/processed.cleveland.data")
    columns = ['age', 'sex', 'cp', 'trestbps', 'chol', 'fbs', 'restecg',
               'thalach', 'exang', 'oldpeak', 'slope', 'ca', 'thal', 'target']
    df = pd.read_csv(url, names=columns, na_values='?').dropna()
    clinical_vars = ['age', 'trestbps', 'chol', 'thalach', 'oldpeak']
    labels = ['Age', 'RestingBP', 'Cholesterol', 'MaxHR', 'STDepression']
    X_raw = df[clinical_vars].astype(float).values
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X_raw)
    return X_scaled, X_raw, labels


def compute_dpi(X_scaled, X_raw, n):
    corr_base = np.abs(np.corrcoef(X_scaled.T))
    np.fill_diagonal(corr_base, 0)
    M_reg = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                cov_ij = np.cov(X_raw[:, i], X_raw[:, j])[0, 1]
                var_i = np.var(X_raw[:, i], ddof=1)
                M_reg[i, j] = abs(cov_ij / (var_i + 1e-10))

    def median_bw(x):
        x = x.reshape(-1, 1)
        dists = np.abs(x - x.T)
        return max(np.median(dists[dists > 0]), 1e-5)

    def neg_hsic(x, y):
        m = len(x)
        sx, sy = median_bw(x), median_bw(y)
        x, y = x.reshape(-1, 1), y.reshape(-1, 1)
        K = np.exp(-(x - x.T)**2 / (2 * sx**2))
        L = np.exp(-(y - y.T)**2 / (2 * sy**2))
        H = np.eye(m) - np.ones((m, m)) / m
        return -np.trace(K @ H @ L @ H) / (m - 1)**2

    M_anm = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                beta = (np.dot(X_scaled[:, i], X_scaled[:, j]) /
                        (np.dot(X_scaled[:, i], X_scaled[:, i]) + 1e-10))
                res = X_scaled[:, j] - beta * X_scaled[:, i]
                M_anm[i, j] = neg_hsic(X_scaled[:, i], res)

    def knn_ent(x, k=5):
        x = np.sort(x)
        dists = np.array([np.sort(np.abs(x - x[idx]))[min(k, len(x)-1)]
                          for idx in range(len(x))])
        dists = dists[dists > 0]
        return np.mean(np.log(dists + 1e-10))

    M_ent = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                beta = (np.dot(X_scaled[:, i], X_scaled[:, j]) /
                        (np.dot(X_scaled[:, i], X_scaled[:, i]) + 1e-10))
                res = X_scaled[:, j] - beta * X_scaled[:, i]
                M_ent[i, j] = max(0, knn_ent(X_scaled[:, j]) - knn_ent(res))

    def normalize_asym(M):
        A = M - M.T
        mx = np.max(np.abs(A))
        return A / (mx + 1e-10)

    A_reg = normalize_asym(M_reg)
    A_anm = normalize_asym(M_anm)
    A_ent = normalize_asym(M_ent)
    A_combined = (A_reg + A_anm + A_ent) / 3.0
    DPI = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI[i, j] = corr_base[i, j] * (1 + A_combined[i, j])
    return DPI


def build_magnetic_laplacian(W, direction_matrix, q=0.25):
    n = W.shape[0]
    W_sym = (W + W.T) / 2
    H = np.zeros((n, n), dtype=complex)
    for i in range(n):
        for j in range(n):
            if i != j:
                d_ij = (np.sign(direction_matrix[i, j])
                        if np.abs(direction_matrix[i, j]) > 1e-10 else 0.0)
                H[i, j] = W_sym[i, j] * np.exp(1j * 2 * np.pi * q * d_ij)
    D = np.diag(np.real(np.sum(np.abs(H), axis=1)))
    D_inv_sqrt = np.diag(1.0 / np.sqrt(np.diag(D) + 1e-10))
    L_mag = np.eye(n) - D_inv_sqrt @ H @ D_inv_sqrt
    return L_mag


# ============================================================
# Load and compute
# ============================================================
print("Loading data...")
X_scaled, X_raw, labels = load_data()
n_vars = len(labels)

print("Computing DPI...")
DPI = compute_dpi(X_scaled, X_raw, n_vars)

clinical_influence = np.array([
    [0.0, 0.6, 0.4, 0.5, 0.3],
    [0.1, 0.0, 0.2, 0.3, 0.4],
    [0.1, 0.3, 0.0, 0.1, 0.3],
    [0.1, 0.2, 0.1, 0.0, 0.5],
    [0.0, 0.1, 0.0, 0.2, 0.0],
])
alpha = 0.6
utility = alpha * clinical_influence + (1 - alpha) * DPI
asymmetry = utility - utility.T

# ============================================================
# FIGURE 3: CCI Complex Plane (q=0.15 for non-degenerate SCC)
# ============================================================
print("\nGenerating Figure 3: CCI Complex Plane (q=0.15)...")

# Use q=0.15 so that cos(2*pi*0.15) = cos(54 deg) != 0
# At q=0.25, cos(pi/2)=0 makes SCC identically zero (degenerate)
q_cci = 0.15
L_mag_cci = build_magnetic_laplacian(utility, asymmetry, q=q_cci)
evals_cci, evecs_cci = linalg.eigh(L_mag_cci)

fig, ax = plt.subplots(figsize=(10, 9))

# Compute SCC and SCD
pair_labels = []
scc_vals, scd_vals = [], []
for i in range(n_vars):
    for j in range(i + 1, n_vars):
        scc = sum(evals_cci[k] * np.abs(evecs_cci[i, k]) *
                  np.abs(evecs_cci[j, k]) *
                  np.cos(np.angle(evecs_cci[i, k]) - np.angle(evecs_cci[j, k]))
                  for k in range(n_vars))
        scd = sum(evals_cci[k] * np.abs(evecs_cci[i, k]) *
                  np.abs(evecs_cci[j, k]) *
                  np.sin(np.angle(evecs_cci[i, k]) - np.angle(evecs_cci[j, k]))
                  for k in range(n_vars))
        pair_labels.append(f"{labels[i]} \u2192 {labels[j]}")
        scc_vals.append(scc)
        scd_vals.append(scd)

scc_vals = np.array(scc_vals)
scd_vals = np.array(scd_vals)

print(f"  SCC range: [{scc_vals.min():.4f}, {scc_vals.max():.4f}]")
print(f"  SCD range: [{scd_vals.min():.4f}, {scd_vals.max():.4f}]")

# Color by SCD sign
colors = ['#E74C3C' if scd > 0 else '#3498DB' for scd in scd_vals]
sizes = 200 + 300 * (np.abs(scc_vals) + np.abs(scd_vals)) / (
    np.max(np.abs(scc_vals)) + np.max(np.abs(scd_vals)) + 1e-10)

ax.scatter(scc_vals, scd_vals, c=colors, s=sizes, edgecolors='k',
           linewidth=1.5, zorder=5, alpha=0.85)

# Sort points by SCD (y-value) for systematic label placement
order = np.argsort(scd_vals)[::-1]  # top to bottom

# Place labels to the right side, evenly spaced vertically
label_x_base = 0.05  # fixed x position for labels (right of origin)
y_min_label = scd_vals[order[-1]] - 0.03
y_max_label = scd_vals[order[0]] + 0.03
y_positions = np.linspace(y_max_label, y_min_label, len(pair_labels))

for rank, k in enumerate(order):
    ax.annotate(pair_labels[k],
                xy=(scc_vals[k], scd_vals[k]),
                xytext=(label_x_base + 0.02, y_positions[rank]),
                fontsize=8, ha='left', va='center',
                bbox=dict(boxstyle='round,pad=0.15',
                          facecolor='white', alpha=0.85,
                          edgecolor='gray', linewidth=0.5),
                arrowprops=dict(arrowstyle='->', color='gray',
                                lw=0.8, alpha=0.6,
                                connectionstyle='arc3,rad=0.1'))

# Axes
ax.axhline(0, color='gray', lw=1, ls='--', alpha=0.5)
ax.axvline(0, color='gray', lw=1, ls='--', alpha=0.5)

# Set symmetric limits based on data
max_scc = max(abs(scc_vals.min()), abs(scc_vals.max())) * 1.4
max_scd = max(abs(scd_vals.min()), abs(scd_vals.max())) * 1.3
ax.set_xlim(-max_scc, max_scc)
ax.set_ylim(-max_scd, max_scd)

# Quadrant labels
xlim = ax.get_xlim()
ylim = ax.get_ylim()
ax.text(xlim[1] * 0.92, ylim[1] * 0.88,
        'Strong coupling\n+ Forward ($i \\to j$)',
        fontsize=9, ha='right', va='top', color='#E74C3C',
        fontweight='bold', style='italic')
ax.text(xlim[1] * 0.92, ylim[0] * 0.88,
        'Strong coupling\n+ Reverse ($j \\to i$)',
        fontsize=9, ha='right', va='bottom', color='#3498DB',
        fontweight='bold', style='italic')

ax.set_xlabel('Re[CCI] = SCC (Coupling Strength)', fontsize=12,
              fontweight='bold')
ax.set_ylabel('Im[CCI] = SCD (Causal Direction)', fontsize=12,
              fontweight='bold')
ax.set_title('Complex Causal Index in the Complex Plane\n'
             r'CCI$(i,j) = \mathrm{SCC}(i,j) + i \cdot \mathrm{SCD}(i,j)$'
             f'  ($q = {q_cci}$)',
             fontsize=13, fontweight='bold')

legend_elements = [
    plt.scatter([], [], c='#E74C3C', s=100, edgecolors='k',
                label=r'SCD > 0 ($i \to j$)'),
    plt.scatter([], [], c='#3498DB', s=100, edgecolors='k',
                label=r'SCD < 0 ($j \to i$)'),
]
ax.legend(handles=legend_elements, loc='lower left', fontsize=10,
          framealpha=0.9)

ax.grid(True, alpha=0.2)
plt.tight_layout()
plt.savefig(f'{FIG_DIR}/fig_cci_complex_plane.png', dpi=200,
            bbox_inches='tight', facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_cci_complex_plane.png")


# ============================================================
# FIGURE 4: Magnetic Laplacian Eigenvectors (fix label overlap)
# ============================================================
print("\nGenerating Figure 4: Magnetic Laplacian Eigenvectors...")

fig, axes = plt.subplots(1, 3, figsize=(16, 5.5))
q_values = [0, 0.1, 0.25]
subtitles = [
    '$q = 0$ (no directionality)\nAll points on real axis',
    '$q = 0.1$ (partial directionality)',
    '$q = 0.25$ (max directionality)\nPhase = causal direction'
]

for idx, (q_val, subtitle) in enumerate(zip(q_values, subtitles)):
    ax = axes[idx]
    L_mag = build_magnetic_laplacian(utility, asymmetry, q=q_val)
    evals, evecs = linalg.eigh(L_mag)

    # Use Fiedler eigenvector (2nd smallest eigenvalue)
    fiedler_idx = 1
    fiedler = evecs[:, fiedler_idx]

    re_parts = np.real(fiedler)
    im_parts = np.imag(fiedler)

    # Unit circle
    circle = Circle((0, 0), np.max(np.abs(fiedler)) * 1.1,
                    fill=False, color='gray', ls='--', lw=0.8, alpha=0.4)
    ax.add_patch(circle)

    # Set limits first
    lim = np.max(np.abs(fiedler)) * 1.6
    ax.set_xlim(-lim, lim)
    ax.set_ylim(-lim, lim)

    # Plot points
    ax.scatter(re_parts, im_parts, s=200, c='#2c3e50', edgecolors='k',
               linewidth=1.5, zorder=5, alpha=0.85)

    # Labels with manual offset for clarity
    for k in range(n_vars):
        offset_x = lim * 0.15
        offset_y = lim * 0.10
        # Stagger offsets to avoid overlap
        if k % 2 == 0:
            ox, oy = offset_x, offset_y
            ha = 'left'
        else:
            ox, oy = -offset_x, -offset_y
            ha = 'right'
        ax.annotate(labels[k],
                    (re_parts[k], im_parts[k]),
                    xytext=(re_parts[k] + ox, im_parts[k] + oy),
                    fontsize=9, ha=ha, va='center', fontweight='bold',
                    arrowprops=dict(arrowstyle='-', color='gray',
                                    lw=0.7, alpha=0.5))

    # Axes
    ax.axhline(0, color='gray', lw=0.5, alpha=0.4)
    ax.axvline(0, color='gray', lw=0.5, alpha=0.4)
    ax.set_aspect('equal')
    ax.set_xlabel('Real part', fontsize=10)
    ax.set_ylabel('Imaginary part', fontsize=10)
    ax.set_title(subtitle, fontsize=11, fontweight='bold')
    ax.grid(True, alpha=0.15)

fig.suptitle('Magnetic Laplacian Eigenvector (2nd) on Complex Plane\n'
             'Phase angle encodes causal direction; $q$ controls direction '
             'sensitivity',
             fontsize=13, fontweight='bold', y=1.02)

plt.tight_layout()
plt.savefig(f'{FIG_DIR}/fig2_magnetic_laplacian_q.png', dpi=200,
            bbox_inches='tight', facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig2_magnetic_laplacian_q.png")


# ============================================================
# FIGURE 2: DPI Architecture (fix arrowheads hidden behind boxes)
# ============================================================
print("\nGenerating Figure 2: DPI Architecture (arrowhead fix)...")

import matplotlib.patches as mpatches
import matplotlib.patheffects as pe

fig, ax = plt.subplots(figsize=(15, 9))
ax.set_xlim(-0.5, 15.5)
ax.set_ylim(-0.8, 8.5)
ax.set_aspect('equal')
ax.axis('off')

# Colors
c_data = '#3498DB'
c_reg = '#E74C3C'
c_anm = '#2ECC71'
c_ent = '#9B59B6'
c_dpi = '#F39C12'

def draw_box(ax, x, y, w, h, text, color, fontsize=11, alpha=0.85):
    box = mpatches.FancyBboxPatch(
        (x, y), w, h, boxstyle="round,pad=0.15",
        facecolor=color, edgecolor='k', linewidth=1.5, alpha=alpha)
    ax.add_patch(box)
    ax.text(x + w/2, y + h/2, text, ha='center', va='center',
            fontsize=fontsize, fontweight='bold', color='white',
            path_effects=[pe.withStroke(linewidth=2, foreground='black')])
    return (x, y, w, h)

def draw_arrow(ax, x1, y1, x2, y2, color='k'):
    """Draw arrow with visible arrowhead by using shrinkA/shrinkB."""
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=2.5,
                                shrinkA=8, shrinkB=12,
                                mutation_scale=20))

# Data input box
draw_box(ax, 0.5, 3.2, 2.5, 1.5,
         'Observed\nData\n$(X_1, ..., X_n)$', c_data, fontsize=11)

# Three component boxes
draw_box(ax, 4.0, 6.0, 3.0, 1.5,
         r'$\hat{A}_{\mathrm{reg}}$' + '\nRegression\nAsymmetry',
         c_reg, fontsize=10)
draw_box(ax, 4.0, 3.2, 3.0, 1.5,
         r'$\hat{A}_{\mathrm{ANM}}$' + '\nANM Residual\nIndependence',
         c_anm, fontsize=10)
draw_box(ax, 4.0, 0.5, 3.0, 1.5,
         r'$\hat{A}_{\mathrm{ent}}$' + '\nConditional\nEntropy',
         c_ent, fontsize=10)

# Normalization box
draw_box(ax, 8.0, 3.2, 2.2, 1.5,
         'Normalize\n& Average\n' + r'$\bar{A} = \frac{1}{3}\sum$',
         '#7F8C8D', fontsize=10)

# DPI output box
draw_box(ax, 11.0, 3.2, 2.5, 1.5,
         'DPI\n' + r'$|\hat{\rho}_{ij}|(1+\gamma\bar{A})$',
         c_dpi, fontsize=11)

# Correlation input box
draw_box(ax, 0.5, 6.0, 2.5, 1.2,
         r'$|\hat{\rho}_{ij}|$' + '\nCorrelation\n(symmetric)',
         '#95A5A6', fontsize=9)

# Arrows from data to components
draw_arrow(ax, 3.0, 4.4, 4.0, 6.75, c_data)
draw_arrow(ax, 3.0, 3.95, 4.0, 3.95, c_data)
draw_arrow(ax, 3.0, 3.5, 4.0, 1.25, c_data)

# Arrows from components to normalization
draw_arrow(ax, 7.0, 6.75, 8.0, 4.4, c_reg)
draw_arrow(ax, 7.0, 3.95, 8.0, 3.95, c_anm)
draw_arrow(ax, 7.0, 1.25, 8.0, 3.5, c_ent)

# Arrow from normalization to DPI
draw_arrow(ax, 10.2, 3.95, 11.0, 3.95, '#7F8C8D')

# Arrow from correlation to DPI
draw_arrow(ax, 3.0, 6.6, 11.0, 4.4, '#95A5A6')

# Properties annotations
ax.text(5.5, 7.8, r'$\mathrm{Var}(X_i) \neq \mathrm{Var}(X_j)$',
        fontsize=9, ha='center', color=c_reg, style='italic')
ax.text(5.5, 2.7, r'HSIC: $\varepsilon \perp X_i$',
        fontsize=9, ha='center', color=c_anm, style='italic')
ax.text(5.5, 0.1, r'$H(X_j) - H(X_j|X_i)$',
        fontsize=9, ha='center', color=c_ent, style='italic')

# Key insight labels
ax.text(12.25, 2.5, 'Asymmetric\n(directional)', fontsize=10,
        ha='center', color=c_dpi, fontweight='bold', style='italic')
ax.text(1.75, 7.5, 'Symmetric\n(no direction)', fontsize=10,
        ha='center', color='#95A5A6', fontweight='bold', style='italic')

plt.tight_layout()
plt.savefig(f'{FIG_DIR}/fig_dpi_architecture.png', dpi=200,
            bbox_inches='tight', facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_dpi_architecture.png")

print("\nDone! All figures regenerated.")

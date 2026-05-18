"""
Generate 6 new figures for A1 (厳密/JMLR) and A2 (応用/ECD) papers.
A1: DPI architecture, CCI complex plane, p_flip U-curve
A2: ECD pipeline flowchart, interventionability scatter, clinical feedback network
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib.patheffects as pe
import networkx as nx
from scipy import linalg
from sklearn.preprocessing import StandardScaler
import lingam
import warnings
import os
warnings.filterwarnings('ignore')

plt.rcParams['font.family'] = ['DejaVu Sans']
plt.rcParams['mathtext.fontset'] = 'dejavusans'

A1_FIG_DIR = '/home/ubuntu/repos/wip/spectral-causality-jmlr/figures'
A2_FIG_DIR = '/home/ubuntu/repos/wip/spectral-causality-a2-ecd/figures'
os.makedirs(A1_FIG_DIR, exist_ok=True)
os.makedirs(A2_FIG_DIR, exist_ok=True)

# ============================================================
# Data and computation (reuse from demo_spectral_causality.py)
# ============================================================

def load_data():
    url = "https://archive.ics.uci.edu/ml/machine-learning-databases/heart-disease/processed.cleveland.data"
    columns = ['age', 'sex', 'cp', 'trestbps', 'chol', 'fbs', 'restecg',
               'thalach', 'exang', 'oldpeak', 'slope', 'ca', 'thal', 'target']
    df = pd.read_csv(url, names=columns, na_values='?')
    df = df.dropna()
    clinical_vars = ['age', 'trestbps', 'chol', 'thalach', 'oldpeak']
    labels = ['Age', 'RestingBP', 'Cholesterol', 'MaxHR', 'STDepression']
    X_raw = df[clinical_vars].astype(float).values
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X_raw)
    return X_scaled, X_raw, labels, clinical_vars


def compute_dpi(X_scaled, X_raw):
    n = X_scaled.shape[1]
    corr_base = np.abs(np.corrcoef(X_scaled.T))
    np.fill_diagonal(corr_base, 0)

    # Regression asymmetry
    M_reg = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                cov_ij = np.cov(X_raw[:, i], X_raw[:, j])[0, 1]
                var_i = np.var(X_raw[:, i], ddof=1)
                M_reg[i, j] = abs(cov_ij / (var_i + 1e-10))

    # ANM (HSIC)
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
                beta = np.dot(X_scaled[:, i], X_scaled[:, j]) / (np.dot(X_scaled[:, i], X_scaled[:, i]) + 1e-10)
                res = X_scaled[:, j] - beta * X_scaled[:, i]
                M_anm[i, j] = neg_hsic(X_scaled[:, i], res)

    # Conditional entropy
    def knn_ent(x, k=5):
        x = np.sort(x)
        dists = np.array([np.sort(np.abs(x - x[idx]))[min(k, len(x)-1)] for idx in range(len(x))])
        dists = dists[dists > 0]
        return np.mean(np.log(dists + 1e-10))
    M_ent = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                beta = np.dot(X_scaled[:, i], X_scaled[:, j]) / (np.dot(X_scaled[:, i], X_scaled[:, i]) + 1e-10)
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
    gamma = 1.0
    DPI = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI[i, j] = corr_base[i, j] * (1 + gamma * A_combined[i, j])
    return DPI, A_reg, A_anm, A_ent


def build_magnetic_laplacian(W, direction_matrix, q=0.25):
    n = W.shape[0]
    W_sym = (W + W.T) / 2
    H = np.zeros((n, n), dtype=complex)
    for i in range(n):
        for j in range(n):
            if i != j:
                d_ij = np.sign(direction_matrix[i, j]) if np.abs(direction_matrix[i, j]) > 1e-10 else 0.0
                H[i, j] = W_sym[i, j] * np.exp(1j * 2 * np.pi * q * d_ij)
    D = np.diag(np.real(np.sum(np.abs(H), axis=1)))
    D_inv_sqrt = np.diag(1.0 / np.sqrt(np.diag(D) + 1e-10))
    L_mag = np.eye(n) - D_inv_sqrt @ H @ D_inv_sqrt
    return L_mag


def hodge_decomposition(flow_matrix):
    n = flow_matrix.shape[0]
    edges, flows = [], []
    for i in range(n):
        for j in range(i+1, n):
            edges.append((i, j))
            flows.append(flow_matrix[i, j])
    m = len(edges)
    B = np.zeros((m, n))
    for k, (i, j) in enumerate(edges):
        B[k, i] = -1
        B[k, j] = 1
    flows = np.array(flows)
    BtB = B.T @ B
    Btf = B.T @ flows
    BtB[0, :] = 0; BtB[0, 0] = 1; Btf[0] = 0
    phi = np.linalg.solve(BtB, Btf)
    gradient_flows = B @ phi
    curl_flows = flows - gradient_flows
    total_energy = np.sum(flows**2)
    gradient_energy = np.sum(gradient_flows**2)
    curl_energy = np.sum(curl_flows**2)
    return -phi, gradient_energy, curl_energy, total_energy, gradient_flows, curl_flows, edges, flows


def compute_cci_matrix(evecs, evals, n):
    """Compute Complex Causal Index matrix."""
    CCI = np.zeros((n, n), dtype=complex)
    for i in range(n):
        for j in range(n):
            if i != j:
                for k in range(n):
                    amp_i = np.abs(evecs[i, k])
                    amp_j = np.abs(evecs[j, k])
                    phase_diff = np.angle(evecs[i, k]) - np.angle(evecs[j, k])
                    CCI[i, j] += evals[k] * amp_i * amp_j * np.exp(1j * phase_diff)
    return CCI


# ============================================================
# Load data and compute
# ============================================================
print("Loading data...")
X_scaled, X_raw, labels, clinical_vars = load_data()
n_vars = len(labels)

print("Running LiNGAM...")
model = lingam.DirectLiNGAM()
model.fit(X_scaled)
B_lingam = model.adjacency_matrix_

print("Computing DPI...")
DPI, A_reg, A_anm, A_ent = compute_dpi(X_scaled, X_raw)

# Clinical knowledge matrix
clinical_influence = np.array([
    [0.0, 0.6, 0.4, 0.5, 0.3],
    [0.1, 0.0, 0.2, 0.3, 0.4],
    [0.1, 0.3, 0.0, 0.1, 0.3],
    [0.1, 0.2, 0.1, 0.0, 0.5],
    [0.0, 0.1, 0.0, 0.2, 0.0],
])

alpha = 0.6
utility_matrix = alpha * clinical_influence + (1 - alpha) * DPI
asymmetry = utility_matrix - utility_matrix.T

# Magnetic Laplacian
L_mag = build_magnetic_laplacian(utility_matrix, asymmetry, q=0.25)
eigenvalues, eigenvectors = linalg.eigh(L_mag)

# Hodge decomposition
phi, grad_e, curl_e, total_e, grad_flows, curl_flows, edges, flows = hodge_decomposition(asymmetry)

# CCI matrix
CCI = compute_cci_matrix(eigenvectors, eigenvalues, n_vars)

print(f"r_gradient = {grad_e/total_e:.3f}" if total_e > 1e-12 else "r_gradient = N/A (zero total energy)")
print(f"Causal potential: {dict(zip(labels, phi.round(3)))}")

# ============================================================
# FIGURE A1-New-1: DPI Architecture Diagram
# ============================================================
print("\nGenerating A1-New-1: DPI Architecture...")

fig, ax = plt.subplots(figsize=(14, 8))
ax.set_xlim(0, 14)
ax.set_ylim(0, 8)
ax.axis('off')

# Color scheme
c_data = '#3498DB'
c_reg = '#E74C3C'
c_anm = '#2ECC71'
c_ent = '#9B59B6'
c_dpi = '#F39C12'
c_output = '#1ABC9C'

def draw_box(ax, x, y, w, h, text, color, fontsize=10, alpha=0.85):
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.15",
                         facecolor=color, edgecolor='k', linewidth=1.5, alpha=alpha)
    ax.add_patch(box)
    ax.text(x + w/2, y + h/2, text, ha='center', va='center',
            fontsize=fontsize, fontweight='bold', color='white',
            path_effects=[pe.withStroke(linewidth=2, foreground='black')])

def draw_arrow(ax, x1, y1, x2, y2, color='k'):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=2))

# Data input
draw_box(ax, 0.5, 3.2, 2.5, 1.5, 'Observed\nData\n$(X_1, ..., X_n)$', c_data, fontsize=11)

# Three components
draw_box(ax, 4.0, 6.0, 3.0, 1.5, r'$\hat{A}_{\mathrm{reg}}$' + '\nRegression\nAsymmetry', c_reg, fontsize=10)
draw_box(ax, 4.0, 3.2, 3.0, 1.5, r'$\hat{A}_{\mathrm{ANM}}$' + '\nANM Residual\nIndependence', c_anm, fontsize=10)
draw_box(ax, 4.0, 0.5, 3.0, 1.5, r'$\hat{A}_{\mathrm{ent}}$' + '\nConditional\nEntropy', c_ent, fontsize=10)

# Arrows from data to components
draw_arrow(ax, 3.0, 4.4, 4.0, 6.75, c_data)
draw_arrow(ax, 3.0, 3.95, 4.0, 3.95, c_data)
draw_arrow(ax, 3.0, 3.5, 4.0, 1.25, c_data)

# Normalization box
draw_box(ax, 8.0, 3.2, 2.2, 1.5, 'Normalize\n& Average\n' + r'$\bar{A} = \frac{1}{3}\sum$', '#7F8C8D', fontsize=10)

# Arrows from components to normalization
draw_arrow(ax, 7.0, 6.75, 8.0, 4.4, c_reg)
draw_arrow(ax, 7.0, 3.95, 8.0, 3.95, c_anm)
draw_arrow(ax, 7.0, 1.25, 8.0, 3.5, c_ent)

# DPI output
draw_box(ax, 11.0, 3.2, 2.5, 1.5, 'DPI\n' + r'$|\hat{\rho}_{ij}|(1+\gamma\bar{A})$', c_dpi, fontsize=11)

# Arrow from normalization to DPI
draw_arrow(ax, 10.2, 3.95, 11.0, 3.95, '#7F8C8D')

# Correlation input
draw_box(ax, 0.5, 6.0, 2.5, 1.2, r'$|\hat{\rho}_{ij}|$' + '\nCorrelation\n(symmetric)', '#95A5A6', fontsize=9)
draw_arrow(ax, 3.0, 6.6, 11.0, 4.4, '#95A5A6')

# Properties annotations
ax.text(5.5, 7.8, r'$\mathrm{Var}(X_i) \neq \mathrm{Var}(X_j)$', fontsize=9,
        ha='center', color=c_reg, style='italic')
ax.text(5.5, 2.7, r'HSIC: $\varepsilon \perp X_i$', fontsize=9,
        ha='center', color=c_anm, style='italic')
ax.text(5.5, 0.1, r'$H(X_j) - H(X_j|X_i)$', fontsize=9,
        ha='center', color=c_ent, style='italic')

# Key insight
ax.text(12.25, 2.5, 'Asymmetric\n(directional)', fontsize=10, ha='center',
        color=c_dpi, fontweight='bold', style='italic')
ax.text(1.75, 7.5, 'Symmetric\n(no direction)', fontsize=10, ha='center',
        color='#95A5A6', fontweight='bold', style='italic')

plt.tight_layout()
plt.savefig(f'{A1_FIG_DIR}/fig_dpi_architecture.png', dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_dpi_architecture.png")


# ============================================================
# FIGURE A1-New-2: CCI Complex Plane Plot
# ============================================================
print("\nGenerating A1-New-2: CCI Complex Plane...")

fig, ax = plt.subplots(figsize=(12, 10))

# Compute SCC and SCD separately for clearer values
def compute_scc(evecs, evals, i, j):
    scc = 0
    for k in range(len(evals)):
        amp_i = np.abs(evecs[i, k])
        amp_j = np.abs(evecs[j, k])
        phase_diff = np.angle(evecs[i, k]) - np.angle(evecs[j, k])
        scc += evals[k] * amp_i * amp_j * np.cos(phase_diff)
    return scc

def compute_scd(evecs, evals, i, j):
    scd = 0
    for k in range(len(evals)):
        amp_i = np.abs(evecs[i, k])
        amp_j = np.abs(evecs[j, k])
        phase_diff = np.angle(evecs[i, k]) - np.angle(evecs[j, k])
        scd += evals[k] * amp_i * amp_j * np.sin(phase_diff)
    return scd

# Plot CCI for all ordered pairs (i<j only, upper triangle)
pair_labels = []
re_vals, im_vals = [], []
for i in range(n_vars):
    for j in range(i+1, n_vars):
        scc_val = compute_scc(eigenvectors, eigenvalues, i, j)
        scd_val = compute_scd(eigenvectors, eigenvalues, i, j)
        pair_labels.append(f"{labels[i]} → {labels[j]}")
        re_vals.append(scc_val)
        im_vals.append(scd_val)

re_vals = np.array(re_vals)
im_vals = np.array(im_vals)

# Normalize for better visualization
max_abs = max(np.max(np.abs(re_vals)), np.max(np.abs(im_vals)))
if max_abs < 1e-6:
    # Values too small — rescale to show relative positions
    scale = 1.0 / (max_abs + 1e-10)
    re_vals_plot = re_vals * scale
    im_vals_plot = im_vals * scale
    xlabel_suffix = ' (normalized)'
    ylabel_suffix = ' (normalized)'
else:
    re_vals_plot = re_vals
    im_vals_plot = im_vals
    xlabel_suffix = ''
    ylabel_suffix = ''

# Color by direction (SCD sign = Im part sign)
colors = ['#E74C3C' if im > 0 else '#3498DB' for im in im_vals_plot]

# Fixed size for visibility
ax.scatter(re_vals_plot, im_vals_plot, c=colors, s=400, edgecolors='k',
           linewidth=2, zorder=5, alpha=0.85)

# Label each point with offset to avoid overlap
offsets = [
    (0.05, 0.05), (-0.15, 0.05), (0.05, -0.08), (-0.15, -0.08), (0.05, 0.05),
    (-0.15, 0.05), (0.05, -0.08), (-0.15, -0.08), (0.05, 0.05), (-0.15, -0.08),
]
for k in range(len(pair_labels)):
    data_range_x = max(re_vals_plot) - min(re_vals_plot)
    data_range_y = max(im_vals_plot) - min(im_vals_plot)
    ox = data_range_x * 0.06 * (1 if k % 2 == 0 else -1.5)
    oy = data_range_y * 0.06 * (1 if k % 3 == 0 else -1)
    ha = 'left' if ox > 0 else 'right'
    ax.annotate(pair_labels[k],
                (re_vals_plot[k], im_vals_plot[k]),
                xytext=(re_vals_plot[k] + ox, im_vals_plot[k] + oy),
                fontsize=9, ha=ha, va='center', fontweight='bold',
                arrowprops=dict(arrowstyle='-', color='gray', lw=0.8, alpha=0.5),
                bbox=dict(boxstyle='round,pad=0.2', facecolor='white', alpha=0.8, edgecolor='gray'))

# Axes
ax.axhline(0, color='gray', lw=1, ls='--', alpha=0.5)
ax.axvline(0, color='gray', lw=1, ls='--', alpha=0.5)

# Add padding to limits
pad_x = (max(re_vals_plot) - min(re_vals_plot)) * 0.3 + 0.1
pad_y = (max(im_vals_plot) - min(im_vals_plot)) * 0.3 + 0.1
ax.set_xlim(min(re_vals_plot) - pad_x, max(re_vals_plot) + pad_x)
ax.set_ylim(min(im_vals_plot) - pad_y, max(im_vals_plot) + pad_y)

# Quadrant labels
xlim = ax.get_xlim()
ylim = ax.get_ylim()
ax.text(xlim[1] * 0.95, ylim[1] * 0.9, 'Strong coupling\n+ Forward ($i \\to j$)',
        fontsize=10, ha='right', va='top', color='#E74C3C', fontweight='bold', style='italic')
ax.text(xlim[1] * 0.95, ylim[0] * 0.9, 'Strong coupling\n+ Reverse ($j \\to i$)',
        fontsize=10, ha='right', va='bottom', color='#3498DB', fontweight='bold', style='italic')

ax.set_xlabel(f'Re[CCI] = SCC (Coupling Strength){xlabel_suffix}', fontsize=13, fontweight='bold')
ax.set_ylabel(f'Im[CCI] = SCD (Causal Direction){ylabel_suffix}', fontsize=13, fontweight='bold')
ax.set_title('Complex Causal Index in the Complex Plane\n'
             r'CCI$(i,j) = \mathrm{SCC}(i,j) + i \cdot \mathrm{SCD}(i,j)$',
             fontsize=14, fontweight='bold')

# Legend
legend_elements = [
    plt.scatter([], [], c='#E74C3C', s=150, edgecolors='k', label=r'SCD > 0 ($i \to j$)'),
    plt.scatter([], [], c='#3498DB', s=150, edgecolors='k', label=r'SCD < 0 ($j \to i$)'),
]
ax.legend(handles=legend_elements, loc='lower left', fontsize=11, framealpha=0.9)

ax.grid(True, alpha=0.2)
plt.tight_layout()
plt.savefig(f'{A1_FIG_DIR}/fig_cci_complex_plane.png', dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_cci_complex_plane.png")


# ============================================================
# FIGURE A1-New-3: p_flip U-shaped Curve
# ============================================================
print("\nGenerating A1-New-3: p_flip U-curve...")

n_trials = 200
p_flips = np.array([0.0, 0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5,
                     0.55, 0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1.0])

# True asymmetry from clinical knowledge
C_true = clinical_influence
A_true = C_true - C_true.T

r_grad_means = []
r_grad_stds = []
r_grad_all = []

for p in p_flips:
    r_grads = []
    for trial in range(n_trials):
        # Flip edge directions with probability p
        A_flipped = A_true.copy()
        for i in range(n_vars):
            for j in range(i+1, n_vars):
                if np.random.random() < p:
                    A_flipped[i, j] = -A_true[i, j]
                    A_flipped[j, i] = -A_true[j, i]

        # Add DPI data component (alpha=0.6)
        data_asym = DPI - DPI.T
        combined_asym = alpha * A_flipped + (1 - alpha) * data_asym

        phi_t, grad_t, curl_t, total_t, _, _, _, _ = hodge_decomposition(combined_asym)
        r_grads.append(grad_t / total_t if total_t > 1e-12 else 0)

    r_grad_means.append(np.mean(r_grads))
    r_grad_stds.append(np.std(r_grads))
    r_grad_all.append(r_grads)

r_grad_means = np.array(r_grad_means)
r_grad_stds = np.array(r_grad_stds)

# Theoretical curve: (1-2p)^2 * r_star
r_star = r_grad_means[0]  # r_gradient at p=0
p_theory = np.linspace(0, 1, 200)
r_theory = (1 - 2 * p_theory)**2 * r_star

fig, ax = plt.subplots(figsize=(12, 7))

# Error bars
ax.fill_between(p_flips, r_grad_means - r_grad_stds, r_grad_means + r_grad_stds,
                color='#3498DB', alpha=0.2, label='±1 SD (200 trials)')
ax.plot(p_flips, r_grad_means, 'o-', color='#3498DB', lw=2.5, markersize=8,
        label=r'Empirical $r_{\mathrm{gradient}}$ (mean)', zorder=5)

# Theoretical curve
ax.plot(p_theory, r_theory, '--', color='#E74C3C', lw=2,
        label=r'Theory: $(1-2p)^2 \cdot r^*$', alpha=0.8)

# Critical threshold line
ax.axhline(0.5, color='gray', ls=':', lw=1.5, alpha=0.6)
ax.text(0.02, 0.52, r'DAG threshold ($r_{\mathrm{gradient}} = 0.5$)',
        fontsize=10, color='gray', fontweight='bold')

# Mark critical p*
p_star = 0.15
ax.axvline(p_star, color='#E74C3C', ls=':', lw=1.5, alpha=0.6)
ax.annotate(r'$p^*_{\mathrm{flip}} \approx 0.15$' + '\n(85% correct required)',
            xy=(p_star, 0.5), xytext=(p_star + 0.08, 0.35),
            fontsize=11, fontweight='bold', color='#E74C3C',
            arrowprops=dict(arrowstyle='->', color='#E74C3C', lw=1.5))

# Mark minimum
min_idx = np.argmin(r_grad_means)
ax.annotate(f'Minimum at\n$p_{{flip}} = {p_flips[min_idx]:.2f}$',
            xy=(p_flips[min_idx], r_grad_means[min_idx]),
            xytext=(p_flips[min_idx] + 0.1, r_grad_means[min_idx] - 0.08),
            fontsize=10, fontweight='bold', color='#9B59B6',
            arrowprops=dict(arrowstyle='->', color='#9B59B6', lw=1.5))

# Annotations for endpoints
ax.annotate('Correct\nDAG', xy=(0, r_star), xytext=(0.05, r_star + 0.06),
            fontsize=10, fontweight='bold', color='#2ECC71')
ax.annotate('Reversed\nDAG', xy=(1, r_grad_means[-1]), xytext=(0.88, r_grad_means[-1] + 0.06),
            fontsize=10, fontweight='bold', color='#2ECC71')

ax.set_xlabel(r'Flip probability $p_{\mathrm{flip}}$ (fraction of reversed edges)', fontsize=13, fontweight='bold')
ax.set_ylabel(r'Gradient energy ratio $r_{\mathrm{gradient}}$', fontsize=13, fontweight='bold')
ax.set_title('Knowledge Quality Phase Transition\n'
             'Partial misinformation is worse than complete ignorance',
             fontsize=14, fontweight='bold')
ax.legend(fontsize=11, loc='upper center', framealpha=0.9)
ax.set_xlim(-0.02, 1.02)
ax.set_ylim(0, 1.0)
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(f'{A1_FIG_DIR}/fig_pflip_ucurve.png', dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_pflip_ucurve.png")


# ============================================================
# FIGURE A2-New-1: ECD Pipeline Flowchart
# ============================================================
print("\nGenerating A2-New-1: ECD Pipeline Flowchart...")

fig, ax = plt.subplots(figsize=(16, 9))
ax.set_xlim(0, 16)
ax.set_ylim(0, 9)
ax.axis('off')

c_step1 = '#E74C3C'
c_step2 = '#3498DB'
c_step3 = '#2ECC71'
c_step4 = '#9B59B6'
c_input = '#95A5A6'
c_output_box = '#F39C12'

def draw_rounded_box(ax, x, y, w, h, text, color, fontsize=10, text_color='white'):
    box = FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.2",
                         facecolor=color, edgecolor='k', linewidth=2, alpha=0.9)
    ax.add_patch(box)
    ax.text(x + w/2, y + h/2, text, ha='center', va='center',
            fontsize=fontsize, fontweight='bold', color=text_color,
            path_effects=[pe.withStroke(linewidth=1.5, foreground='black')] if text_color == 'white' else [])

def draw_thick_arrow(ax, x1, y1, x2, y2, color='k', lw=3):
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', color=color, lw=lw))

# Input
draw_rounded_box(ax, 0.3, 3.5, 2.2, 2.0, 'Observational\nData\n$(X_1,...,X_n)$', c_input, fontsize=11)

# Step 1: LiNGAM
draw_rounded_box(ax, 3.3, 3.5, 2.5, 2.0, 'Step 1\nLiNGAM\nBootstrap\n(1000 iter)', c_step1, fontsize=10)
ax.text(4.55, 2.9, 'DAG + $B_{ij}$\nconfidence', fontsize=8, ha='center', color=c_step1, style='italic')

# Step 2: DPI + Spectral
draw_rounded_box(ax, 6.6, 3.5, 2.5, 2.0, 'Step 2\nDPI-Augmented\nSpectral\nAnalysis', c_step2, fontsize=10)
ax.text(7.85, 2.9, '$\\alpha = 0.01$–$0.1$\nDCG edges', fontsize=8, ha='center', color=c_step2, style='italic')

# Step 3: Hodge
draw_rounded_box(ax, 9.9, 3.5, 2.5, 2.0, 'Step 3\nHodge\nDecomposition\n$\\omega = \\delta_0\\phi + ...$', c_step3, fontsize=10)
ax.text(11.15, 2.9, '$r_{\\mathrm{gradient}}$\nfeedback rates', fontsize=8, ha='center', color=c_step3, style='italic')

# Step 4: Interventionability
draw_rounded_box(ax, 13.2, 3.5, 2.5, 2.0, 'Step 4\nIntervention-\nability\nScoring', c_step4, fontsize=10)
ax.text(14.45, 2.9, '$\\phi \\to \\iota$\nprioritization', fontsize=8, ha='center', color=c_step4, style='italic')

# Arrows between steps
draw_thick_arrow(ax, 2.5, 4.5, 3.3, 4.5, c_input)
draw_thick_arrow(ax, 5.8, 4.5, 6.6, 4.5, c_step1)
draw_thick_arrow(ax, 9.1, 4.5, 9.9, 4.5, c_step2)
draw_thick_arrow(ax, 12.4, 4.5, 13.2, 4.5, c_step3)

# Output boxes at top
draw_rounded_box(ax, 3.3, 6.8, 2.5, 1.5, 'Causal DAG\n+ Edge\nConfidences', c_step1, fontsize=9, text_color='white')
draw_rounded_box(ax, 6.6, 6.8, 2.5, 1.5, 'Directed\nCyclic Graph\n+ DPI Scores', c_step2, fontsize=9, text_color='white')
draw_rounded_box(ax, 9.9, 6.8, 2.5, 1.5, 'Gradient/Curl\nDecomposition\n+ $\\phi$ Potential', c_step3, fontsize=9, text_color='white')
draw_rounded_box(ax, 13.2, 6.8, 2.5, 1.5, 'Treatment\nTarget\nRanking', c_step4, fontsize=9, text_color='white')

# Vertical arrows to outputs
for x_center in [4.55, 7.85, 11.15, 14.45]:
    draw_thick_arrow(ax, x_center, 5.5, x_center, 6.8, 'gray', lw=2)

# Bottom: Hill's criteria coverage
draw_rounded_box(ax, 3.3, 0.3, 12.4, 1.5, "Hill's Nine Criteria Coverage:  "
                 "H1 Strength ● | H2 Consistency ● | H3 Specificity ● | H4 Temporality ● | "
                 "H5 Dose-response ● | H6 Plausibility ● | H7 Coherence ● | H8 Experiment ○ | H9 Analogy ●",
                 c_output_box, fontsize=9, text_color='black')

# Bracket from all steps to Hill
ax.annotate('', xy=(9.5, 1.8), xytext=(9.5, 3.5),
            arrowprops=dict(arrowstyle='->', color=c_output_box, lw=2, ls='--'))

plt.tight_layout()
plt.savefig(f'{A2_FIG_DIR}/fig_ecd_pipeline.png', dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_ecd_pipeline.png")


# ============================================================
# FIGURE A2-New-2: Interventionability Scatter (φ vs ι)
# ============================================================
print("\nGenerating A2-New-2: Interventionability Scatter...")

# Data from manuscript
var_data = {
    'Age':          {'phi': 0.000,  'iota': 0.0, 'color': '#E74C3C', 'intervention': 'Non-modifiable'},
    'MaxHR':        {'phi': -0.204, 'iota': 0.3, 'color': '#F39C12', 'intervention': 'Exercise training\n(limited)'},
    'STDepression': {'phi': -0.324, 'iota': 0.5, 'color': '#9B59B6', 'intervention': 'PCI / CABG'},
    'Cholesterol':  {'phi': -0.168, 'iota': 0.9, 'color': '#2ECC71', 'intervention': 'Statins\n(HMG-CoA inhibitor)'},
    'RestingBP':    {'phi': -0.204, 'iota': 0.8, 'color': '#3498DB', 'intervention': 'ACE-i / ARB / CCB'},
}

fig, ax = plt.subplots(figsize=(12, 8))

# Custom label offsets to avoid overlap with y-axis and each other
label_offsets = {
    'Age':          {'dx':  0.015, 'dy': 0.04, 'ha': 'left',  'int_dy': -0.04},
    'MaxHR':        {'dx':  0.015, 'dy': 0.04, 'ha': 'left',  'int_dy': -0.06},
    'STDepression': {'dx':  0.015, 'dy': 0.04, 'ha': 'left',  'int_dy': -0.04},
    'Cholesterol':  {'dx':  0.015, 'dy': 0.04, 'ha': 'left',  'int_dy': -0.04},
    'RestingBP':    {'dx': -0.015, 'dy': 0.04, 'ha': 'right', 'int_dy': -0.04},
}

for name, d in var_data.items():
    ax.scatter(d['phi'], d['iota'], s=600, c=d['color'], edgecolors='k',
              linewidth=2, zorder=5, alpha=0.9)
    lo = label_offsets[name]
    # Name label (inside graph, to the right of point)
    ax.annotate(name, (d['phi'] + lo['dx'], d['iota'] + lo['dy']),
                fontsize=12, fontweight='bold', ha=lo['ha'], va='bottom', zorder=6)
    # Intervention label (inside graph, below name)
    ax.annotate(d['intervention'], (d['phi'] + lo['dx'], d['iota'] + lo['int_dy']),
                fontsize=9, ha=lo['ha'], va='top', color='gray', style='italic', zorder=6)

# Trend line (conceptual)
phi_range = np.linspace(-0.35, 0.02, 50)
trend = -2.2 * phi_range + 0.1
trend = np.clip(trend, 0, 1)
ax.plot(phi_range, trend, '--', color='gray', alpha=0.4, lw=1.5)

# Quadrant annotations
ax.text(-0.33, 0.95, 'High interventionability\n(treatment targets)',
        fontsize=11, color='#2ECC71', fontweight='bold', ha='left', va='top',
        bbox=dict(boxstyle='round', facecolor='white', alpha=0.8, edgecolor='#2ECC71'))
ax.text(0.01, 0.05, 'Low interventionability\n(confounders to control)',
        fontsize=11, color='#E74C3C', fontweight='bold', ha='right', va='bottom',
        bbox=dict(boxstyle='round', facecolor='white', alpha=0.8, edgecolor='#E74C3C'))

# Arrow showing the correspondence
ax.annotate('', xy=(-0.30, 0.85), xytext=(-0.02, 0.15),
            arrowprops=dict(arrowstyle='->', color='gray', lw=2, ls='--', alpha=0.4))
ax.text(-0.16, 0.55, 'Upstream → Downstream\n= Non-modifiable → Actionable',
        fontsize=9, ha='center', va='center', color='gray', style='italic', rotation=50)

ax.set_xlabel(r'Hodge Causal Potential $\phi$ (higher = more upstream)', fontsize=13, fontweight='bold')
ax.set_ylabel(r'Interventionability $\iota$ (higher = more actionable)', fontsize=13, fontweight='bold')
ax.set_title('Causal Potential vs. Clinical Interventionability\n'
             r'$\phi$ (mathematical) $\longleftrightarrow$ $\iota$ (clinical)',
             fontsize=14, fontweight='bold')
ax.set_xlim(-0.37, 0.05)
ax.set_ylim(-0.05, 1.05)
ax.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(f'{A2_FIG_DIR}/fig_interventionability.png', dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_interventionability.png")


# ============================================================
# FIGURE A2-New-3: Clinical Feedback Loop Network
# ============================================================
print("\nGenerating A2-New-3: Clinical Feedback Network...")

# Feedback data from manuscript
edge_data = [
    ('Age', 'RestingBP',    0.00, 'Age → RestBP',    'Aging → hypertension'),
    ('Age', 'Cholesterol',  0.01, 'Age → Chol',      'Aging → dyslipidemia'),
    ('Age', 'MaxHR',        0.34, 'Age → MaxHR',     'Aging-fitness decline'),
    ('Age', 'STDepression', 0.10, 'Age → STDep',     'Aging → ischemia risk'),
    ('RestingBP', 'STDepression', 0.24, 'RestBP ↔ STDep', 'Hypertension-ischemia'),
    ('MaxHR', 'STDepression', 0.73, 'MaxHR ↔ STDep',  'Exercise-ischemia\nfeedback loop'),
    ('RestingBP', 'Cholesterol', 0.05, 'RestBP → Chol', 'Metabolic syndrome'),
]

fig, ax = plt.subplots(figsize=(12, 10))

# Node positions (circular layout with Age at top)
pos = {
    'Age':          (0.5, 0.85),
    'RestingBP':    (0.15, 0.55),
    'Cholesterol':  (0.85, 0.55),
    'MaxHR':        (0.25, 0.15),
    'STDepression': (0.75, 0.15),
}

# Node properties
node_iota = {'Age': 0.0, 'RestingBP': 0.8, 'Cholesterol': 0.9, 'MaxHR': 0.3, 'STDepression': 0.5}
node_colors = {'Age': '#E74C3C', 'RestingBP': '#3498DB', 'Cholesterol': '#2ECC71',
               'MaxHR': '#F39C12', 'STDepression': '#9B59B6'}

# Draw edges
for src, dst, fb_rate, label, clinical in edge_data:
    x1, y1 = pos[src]
    x2, y2 = pos[dst]
    # Midpoint for label
    mx, my = (x1 + x2) / 2, (y1 + y2) / 2

    if fb_rate < 0.10:
        # Unidirectional (solid arrow)
        style = '->'
        color = '#2ECC71'
        lw = 2
        ls = '-'
    elif fb_rate < 0.50:
        # Weak feedback (dashed, bidirectional-ish)
        style = '<->'
        color = '#F39C12'
        lw = 2.5
        ls = '--'
    else:
        # Strong feedback (thick, double arrow)
        style = '<->'
        color = '#E74C3C'
        lw = 4
        ls = '-'

    # Shorten arrows to not overlap with node circles
    dx, dy = x2 - x1, y2 - y1
    dist = np.sqrt(dx**2 + dy**2)
    shrink = 0.06 / dist
    ax.annotate('', xy=(x2 - dx * shrink, y2 - dy * shrink),
                xytext=(x1 + dx * shrink, y1 + dy * shrink),
                arrowprops=dict(arrowstyle=style, color=color, lw=lw, linestyle=ls))

    # Edge label
    # Offset perpendicular to edge
    perp_x, perp_y = -dy / dist * 0.04, dx / dist * 0.04
    ax.text(mx + perp_x, my + perp_y, f'{int(fb_rate*100)}%',
            fontsize=10, fontweight='bold', ha='center', va='center', color=color,
            bbox=dict(boxstyle='round,pad=0.15', facecolor='white', alpha=0.9, edgecolor=color))

    # Clinical interpretation (smaller)
    ax.text(mx + perp_x * 3, my + perp_y * 3 - 0.035, clinical,
            fontsize=7, ha='center', va='top', color='gray', style='italic')

# Draw nodes
for name, (x, y) in pos.items():
    circle = plt.Circle((x, y), 0.055, facecolor=node_colors[name],
                        edgecolor='k', linewidth=2, zorder=10, alpha=0.9)
    ax.add_patch(circle)
    ax.text(x, y + 0.005, name.replace('Depression', '\nDepression'),
            ha='center', va='center', fontsize=9, fontweight='bold', color='white',
            path_effects=[pe.withStroke(linewidth=2, foreground='black')], zorder=11)
    # Interventionability below node
    ax.text(x, y - 0.075, f'$\\iota$ = {node_iota[name]}',
            ha='center', va='top', fontsize=8, color=node_colors[name], fontweight='bold')

# Legend
legend_elements = [
    plt.Line2D([0], [0], color='#2ECC71', lw=2, ls='-', label='Unidirectional (< 10% feedback)'),
    plt.Line2D([0], [0], color='#F39C12', lw=2.5, ls='--', label='Weak feedback (10–50%)'),
    plt.Line2D([0], [0], color='#E74C3C', lw=4, ls='-', label='Strong feedback (> 50%)'),
]
ax.legend(handles=legend_elements, loc='lower center', fontsize=11,
          ncol=3, framealpha=0.9, bbox_to_anchor=(0.5, -0.02))

# Title
ax.set_title('Clinical Feedback Network: Edge-Level Feedback Rates\n'
             'from Hodge Decomposition (DAG assumption: all edges unidirectional)',
             fontsize=14, fontweight='bold')
ax.set_xlim(-0.05, 1.05)
ax.set_ylim(-0.05, 1.0)
ax.axis('off')

plt.tight_layout()
plt.savefig(f'{A2_FIG_DIR}/fig_feedback_network.png', dpi=200, bbox_inches='tight',
            facecolor='white', edgecolor='none')
plt.close()
print("  Saved: fig_feedback_network.png")


print("\n" + "=" * 60)
print("All 6 figures generated successfully!")
print(f"  A1 figures: {A1_FIG_DIR}/")
print(f"  A2 figures: {A2_FIG_DIR}/")
print("=" * 60)

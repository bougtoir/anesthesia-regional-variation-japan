"""
Spectral Causality Demo — 実データ（UCI Heart Disease）でスペクトル因果性を計算・可視化
専門外向け解説文書用の図を生成する
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
import networkx as nx
from scipy import linalg
from sklearn.preprocessing import StandardScaler
from sklearn.metrics.pairwise import cosine_similarity
import lingam
import warnings
warnings.filterwarnings('ignore')

# --- Japanese font support ---
plt.rcParams['font.family'] = ['DejaVu Sans']
# We'll use English labels for figures (knowledge note: EN figures for EN docs)

OUTPUT_DIR = '/home/ubuntu/repos/wip/spectral-causality-brainstorm/figures'
import os
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ============================================================
# 1. Data Preparation — UCI Heart Disease (Cleveland)
# ============================================================
print("=" * 60)
print("Step 1: Loading UCI Heart Disease Data")
print("=" * 60)

url = "https://archive.ics.uci.edu/ml/machine-learning-databases/heart-disease/processed.cleveland.data"
columns = ['age', 'sex', 'cp', 'trestbps', 'chol', 'fbs', 'restecg',
           'thalach', 'exang', 'oldpeak', 'slope', 'ca', 'thal', 'target']
df = pd.read_csv(url, names=columns, na_values='?')
df = df.dropna()

# Select continuous clinical variables for spectral analysis
clinical_vars = ['age', 'trestbps', 'chol', 'thalach', 'oldpeak']
var_labels = {
    'age': 'Age',
    'trestbps': 'Resting BP',
    'chol': 'Cholesterol',
    'thalach': 'Max Heart Rate',
    'oldpeak': 'ST Depression'
}
var_labels_ja = {
    'age': '年齢',
    'trestbps': '安静時血圧',
    'chol': 'コレステロール',
    'thalach': '最大心拍数',
    'oldpeak': 'ST低下'
}

X = df[clinical_vars].astype(float).values
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)
X_df = pd.DataFrame(X_scaled, columns=clinical_vars)

print(f"  Samples: {len(X_df)}, Variables: {len(clinical_vars)}")
print(f"  Variables: {list(var_labels.values())}")

# ============================================================
# 2. Standard Approach: LiNGAM Causal Discovery
# ============================================================
print("\n" + "=" * 60)
print("Step 2: DirectLiNGAM — Variable-level Causal DAG")
print("=" * 60)

model = lingam.DirectLiNGAM()
model.fit(X_scaled)
B_lingam = model.adjacency_matrix_
causal_order = model.causal_order_

print("  Causal order:", [var_labels[clinical_vars[i]] for i in causal_order])
print("  Adjacency matrix (non-zero = causal effect):")
for i, row in enumerate(B_lingam):
    nonzero = [(var_labels[clinical_vars[j]], f"{row[j]:.3f}") for j in range(len(row)) if abs(row[j]) > 0.05]
    if nonzero:
        print(f"    {var_labels[clinical_vars[i]]} <-- {nonzero}")

# ============================================================
# 3. New Approach: Utility-Based Graph Construction
# ============================================================
print("\n" + "=" * 60)
print("Step 3: Utility-Based Graph — 'What question can this answer?'")
print("=" * 60)

# Simulate utility questions via clinical knowledge encoding
# In practice, LLM would generate these; here we use domain knowledge heuristics
utility_descriptions = {
    'age': ['cardiovascular risk assessment', 'metabolic age estimation', 'treatment candidacy'],
    'trestbps': ['hypertension diagnosis', 'cardiovascular risk assessment', 'end-organ damage'],
    'chol': ['dyslipidemia screening', 'cardiovascular risk assessment', 'statin candidacy'],
    'thalach': ['exercise capacity evaluation', 'cardiac output assessment', 'prognosis estimation'],
    'oldpeak': ['myocardial ischemia detection', 'coronary artery disease severity', 'prognosis estimation']
}

# Create utility embeddings (simplified: binary presence of shared utility)
n_vars = len(clinical_vars)

# Build asymmetric utility matrix: Util(i -> j) = how much does knowing i help answer j's questions?
# Encoded as: correlation of variable i with the "answer domain" of variable j
# Using data-driven proxy + clinical knowledge weight

# Clinical knowledge: directed influence weights (expert-encoded)
# Scale: 0=no influence, 1=strong influence
clinical_influence = np.array([
    # age  BP   chol  HR   ST
    [0.0, 0.6, 0.4, 0.5, 0.3],  # age -> others
    [0.1, 0.0, 0.2, 0.3, 0.4],  # BP -> others
    [0.1, 0.3, 0.0, 0.1, 0.3],  # chol -> others
    [0.1, 0.2, 0.1, 0.0, 0.5],  # HR -> others
    [0.0, 0.1, 0.0, 0.2, 0.0],  # ST -> others (ST is mostly an outcome)
])

# Data-driven component: how predictive is i of j (absolute correlation)
data_corr = np.abs(np.corrcoef(X_scaled.T))
np.fill_diagonal(data_corr, 0)

# Combine: Utility(i -> j) = alpha * clinical + (1-alpha) * data
alpha = 0.6
utility_matrix = alpha * clinical_influence + (1 - alpha) * data_corr

# Asymmetry = directionality signal
asymmetry = utility_matrix - utility_matrix.T
print("  Utility asymmetry matrix (positive = i -> j direction):")
for i in range(n_vars):
    for j in range(i+1, n_vars):
        if abs(asymmetry[i, j]) > 0.05:
            direction = "->" if asymmetry[i, j] > 0 else "<-"
            print(f"    {var_labels[clinical_vars[i]]} {direction} {var_labels[clinical_vars[j]]}  "
                  f"(asymmetry={asymmetry[i,j]:+.3f})")

# ============================================================
# 4. Magnetic Laplacian — Encoding Directionality into Complex Phase
# ============================================================
print("\n" + "=" * 60)
print("Step 4: Magnetic Laplacian — Direction as Complex Phase")
print("=" * 60)

def build_magnetic_laplacian(W, direction_matrix, q=0.25):
    """
    Build the normalized magnetic Laplacian L^(q) = I - D^{-1/2} H^(q) D^{-1/2}
    where H^(q)_{ij} = W_{ij} * exp(i * 2*pi*q * direction(i,j))
    """
    n = W.shape[0]
    # Symmetrize weights for Hermitian property
    W_sym = (W + W.T) / 2
    
    # Build Hermitian matrix with complex phases
    H = np.zeros((n, n), dtype=complex)
    for i in range(n):
        for j in range(n):
            if i != j:
                # direction sign from asymmetry (threshold to avoid noise amplification)
                d_ij = np.sign(direction_matrix[i, j]) if np.abs(direction_matrix[i, j]) > 1e-10 else 0.0
                H[i, j] = W_sym[i, j] * np.exp(1j * 2 * np.pi * q * d_ij)
    
    # Degree matrix
    D = np.diag(np.real(np.sum(np.abs(H), axis=1)))
    D_inv_sqrt = np.diag(1.0 / np.sqrt(np.diag(D) + 1e-10))
    
    # Normalized Magnetic Laplacian: L = I - D^{-1/2} H D^{-1/2}
    L_mag = np.eye(n) - D_inv_sqrt @ H @ D_inv_sqrt
    return L_mag, H

q_values = [0, 0.10, 0.25]
results = {}

for q in q_values:
    L_mag, H_mag = build_magnetic_laplacian(utility_matrix, asymmetry, q=q)
    eigenvalues, eigenvectors = linalg.eigh(L_mag)
    results[q] = {
        'L': L_mag, 'eigenvalues': eigenvalues, 'eigenvectors': eigenvectors
    }
    print(f"\n  q={q}: Eigenvalues = {np.real(eigenvalues).round(4)}")
    if q > 0:
        # Show complex phases in eigenvectors (= directional information)
        phases = np.angle(eigenvectors[:, 1])  # second eigenvector
        print(f"    Eigenvector phases (2nd): {np.degrees(phases).round(1)} degrees")
        print(f"    -> Phase ordering encodes causal flow direction")

# ============================================================
# 5. Spectral Causal Intensity (SCI) and Direction (SCD)
# ============================================================
print("\n" + "=" * 60)
print("Step 5: Computing SCI and SCD")
print("=" * 60)

q_opt = 0.25
evals = results[q_opt]['eigenvalues']
evecs = results[q_opt]['eigenvectors']

# SCI(i -> j) = sum_k f(lambda_k) * |<u_k, e_i>| * |<u_k, e_j>| * cos(phase_i - phase_j)
def compute_spectral_direction(evecs, evals, i, j):
    """
    Compute Spectral Causal Direction from i to j.
    Uses sin(phase_diff) which is antisymmetric: sin(a-b) = -sin(b-a)
    This captures directional information from the magnetic Laplacian.
    """
    scd = 0
    for k in range(len(evals)):
        amp_i = np.abs(evecs[i, k])
        amp_j = np.abs(evecs[j, k])
        phase_diff = np.angle(evecs[i, k]) - np.angle(evecs[j, k])
        weight = evals[k]  # f(lambda) = lambda
        # sin is antisymmetric: encodes direction
        # cos is symmetric: encodes coupling strength
        scd += weight * amp_i * amp_j * np.sin(phase_diff)
    return scd

def compute_sci(evecs, evals, i, j):
    """Compute Spectral Causal Intensity (symmetric coupling strength)"""
    sci = 0
    for k in range(len(evals)):
        amp_i = np.abs(evecs[i, k])
        amp_j = np.abs(evecs[j, k])
        phase_diff = np.angle(evecs[i, k]) - np.angle(evecs[j, k])
        weight = evals[k]
        sci += weight * amp_i * amp_j * np.cos(phase_diff)
    return sci

# Compute SCD for all pairs using sin (antisymmetric)
SCI_matrix = np.zeros((n_vars, n_vars))
SCD_matrix = np.zeros((n_vars, n_vars))

for i in range(n_vars):
    for j in range(n_vars):
        if i != j:
            SCI_matrix[i, j] = compute_sci(evecs, evals, i, j)
            SCD_matrix[i, j] = compute_spectral_direction(evecs, evals, i, j)

print("  Spectral Causal Direction (SCD > 0 means i -> j):")
for i in range(n_vars):
    for j in range(i+1, n_vars):
        scd = SCD_matrix[i, j]
        sci = (SCI_matrix[i, j] + SCI_matrix[j, i]) / 2
        direction = "->" if scd > 0 else "<-"
        print(f"    {var_labels[clinical_vars[i]]} {direction} {var_labels[clinical_vars[j]]}  "
              f"(SCD={scd:+.4f}, coupling={sci:.4f})")

# ============================================================
# 6. Hodge Decomposition
# ============================================================
print("\n" + "=" * 60)
print("Step 6: Hodge Decomposition — Gradient (Causal) vs Curl (Feedback)")
print("=" * 60)

def hodge_decomposition(flow_matrix):
    """
    Simplified Hodge decomposition on a complete graph.
    Decomposes edge flows into gradient (potential-driven) and curl components.
    Returns potential phi (causal depth) and curl energy.
    """
    n = flow_matrix.shape[0]
    # Gradient component: find potential phi such that flow(i,j) ≈ phi(j) - phi(i)
    # Least squares: min_phi sum_{i,j} (flow(i,j) - (phi(j) - phi(i)))^2
    
    edges = []
    flows = []
    for i in range(n):
        for j in range(i+1, n):
            edges.append((i, j))
            flows.append(flow_matrix[i, j])
    
    m = len(edges)
    B = np.zeros((m, n))  # incidence matrix
    for k, (i, j) in enumerate(edges):
        B[k, i] = -1
        B[k, j] = 1
    
    flows = np.array(flows)
    
    # Solve B^T B phi = B^T flows (normal equation)
    BtB = B.T @ B
    Btf = B.T @ flows
    # Add regularization (fix one node potential to 0)
    BtB[0, :] = 0
    BtB[0, 0] = 1
    Btf[0] = 0
    
    phi = np.linalg.solve(BtB, Btf)
    
    # Gradient component
    gradient_flows = B @ phi
    
    # Curl component = total - gradient
    curl_flows = flows - gradient_flows
    
    # Energies
    total_energy = np.sum(flows**2)
    gradient_energy = np.sum(gradient_flows**2)
    curl_energy = np.sum(curl_flows**2)
    
    return phi, gradient_energy, curl_energy, total_energy

# Use asymmetry matrix as flow
phi, grad_e, curl_e, total_e = hodge_decomposition(asymmetry)

# Negate phi so that higher = more upstream (cause side)
# In the flow convention, positive asymmetry(i,j) means i influences j,
# so the gradient phi(j) > phi(i) means j is downstream.
# We flip to make interpretation intuitive.
phi_causal = -phi

print("  Causal Potential (higher = more 'upstream'/causal):")
potential_order = np.argsort(-phi_causal)
for rank, idx in enumerate(potential_order):
    print(f"    #{rank+1}: {var_labels[clinical_vars[idx]]:20s} phi={phi_causal[idx]:+.4f}")

print(f"\n  Gradient (causal) energy: {grad_e:.4f} ({100*grad_e/total_e:.1f}%)")
print(f"  Curl (feedback) energy:   {curl_e:.4f} ({100*curl_e/total_e:.1f}%)")
print(f"  -> {100*grad_e/total_e:.1f}% of information flow is DAG-like (causal)")
print(f"  -> {100*curl_e/total_e:.1f}% is cyclic (feedback loops)")

# Use phi_causal for all subsequent analysis
phi = phi_causal

# ============================================================
# 7. Comparison: LiNGAM causal order vs Hodge potential
# ============================================================
print("\n" + "=" * 60)
print("Step 7: Comparison — LiNGAM Order vs Hodge Potential vs SCD")
print("=" * 60)

lingam_order_labels = [var_labels[clinical_vars[i]] for i in causal_order]
potential_order = np.argsort(-phi)  # re-sort with corrected phi
hodge_order_labels = [var_labels[clinical_vars[i]] for i in potential_order]

print(f"  LiNGAM causal order:  {lingam_order_labels}")
print(f"  Hodge potential order: {hodge_order_labels}")

# Agreement analysis
def rank_agreement(order1, order2, n):
    """Kendall's tau between two orderings"""
    concordant = 0
    discordant = 0
    total = 0
    for i in range(n):
        for j in range(i+1, n):
            pos1_i = list(order1).index(i)
            pos1_j = list(order1).index(j)
            pos2_i = list(order2).index(i)
            pos2_j = list(order2).index(j)
            if (pos1_i - pos1_j) * (pos2_i - pos2_j) > 0:
                concordant += 1
            elif (pos1_i - pos1_j) * (pos2_i - pos2_j) < 0:
                discordant += 1
            total += 1
    return (concordant - discordant) / total

agreement = rank_agreement(causal_order, potential_order, n_vars)
print(f"\n  Order agreement (Kendall tau): {agreement:.2f}")

# ============================================================
# 8. Generate Figures
# ============================================================
print("\n" + "=" * 60)
print("Step 8: Generating Figures")
print("=" * 60)

# Color scheme
COLORS = {
    'upstream': '#E74C3C',    # red = cause
    'midstream': '#F39C12',   # orange
    'downstream': '#3498DB',  # blue = effect
    'neutral': '#95A5A6',
    'gradient': '#2ECC71',    # green
    'curl': '#9B59B6',        # purple
    'lingam': '#E74C3C',
    'spectral': '#3498DB',
    'hodge': '#2ECC71',
}

labels_en = [var_labels[v] for v in clinical_vars]

# --- Figure 1: Conceptual Overview (3-panel) ---
fig, axes = plt.subplots(1, 3, figsize=(18, 6))

# Panel A: Traditional scatter (no structure)
ax = axes[0]
ax.set_title('(A) Traditional View:\nVariables as Isolated Points', fontsize=13, fontweight='bold')
np.random.seed(42)
x_pos = np.random.uniform(0.1, 0.9, n_vars)
y_pos = np.random.uniform(0.1, 0.9, n_vars)
for i in range(n_vars):
    ax.scatter(x_pos[i], y_pos[i], s=800, c=COLORS['neutral'], zorder=3, edgecolors='k')
    ax.annotate(labels_en[i], (x_pos[i], y_pos[i]), ha='center', va='center', fontsize=8, fontweight='bold')
ax.set_xlim(0, 1); ax.set_ylim(0, 1)
ax.set_xticks([]); ax.set_yticks([])
ax.text(0.5, 0.02, 'No structure, no causality', ha='center', fontsize=11, style='italic', color='gray')

# Panel B: LiNGAM DAG (variable-level)
ax = axes[1]
ax.set_title('(B) LiNGAM:\nVariable-Level Causal DAG', fontsize=13, fontweight='bold')
# Layout based on causal order
y_positions = {causal_order[i]: 0.9 - i * 0.18 for i in range(n_vars)}
x_fixed = {i: 0.5 + 0.15 * np.sin(i * 1.2) for i in range(n_vars)}

for i in range(n_vars):
    color = plt.cm.RdYlBu(1 - list(causal_order).index(i) / (n_vars - 1))
    ax.scatter(x_fixed[i], y_positions[i], s=800, c=[color], zorder=3, edgecolors='k')
    ax.annotate(labels_en[i], (x_fixed[i], y_positions[i]), ha='center', va='center', fontsize=8, fontweight='bold')

# Draw causal edges
for i in range(n_vars):
    for j in range(n_vars):
        if abs(B_lingam[j, i]) > 0.05:  # i -> j
            ax.annotate('', xy=(x_fixed[j], y_positions[j] + 0.03),
                        xytext=(x_fixed[i], y_positions[i] - 0.03),
                        arrowprops=dict(arrowstyle='->', color='k', lw=1.5 + abs(B_lingam[j, i]) * 2))
ax.set_xlim(0, 1); ax.set_ylim(0, 1)
ax.set_xticks([]); ax.set_yticks([])
ax.text(0.5, 0.02, 'Directed edges = "X causes Y"', ha='center', fontsize=11, style='italic', color='gray')

# Panel C: Spectral Causality (theme-level + direction)
ax = axes[2]
ax.set_title('(C) Spectral Causality:\nTheme-Level Causal Flow', fontsize=13, fontweight='bold')
# Layout by Hodge potential
y_by_phi = {i: 0.15 + 0.7 * (phi[i] - phi.min()) / (phi.max() - phi.min() + 1e-10) for i in range(n_vars)}
x_by_ev = {i: 0.5 + 0.3 * np.real(evecs[i, 1]) for i in range(n_vars)}

# Draw flow arrows (gradient component)
for i in range(n_vars):
    for j in range(i+1, n_vars):
        if abs(asymmetry[i, j]) > 0.08:
            src, dst = (i, j) if asymmetry[i, j] > 0 else (j, i)
            alpha_val = min(1.0, abs(asymmetry[i, j]) * 3)
            ax.annotate('', xy=(x_by_ev[dst], y_by_phi[dst]),
                        xytext=(x_by_ev[src], y_by_phi[src]),
                        arrowprops=dict(arrowstyle='->', color=COLORS['gradient'],
                                        lw=2, alpha=alpha_val))

for i in range(n_vars):
    norm_phi = (phi[i] - phi.min()) / (phi.max() - phi.min() + 1e-10)
    color = plt.cm.RdYlBu(1 - norm_phi)
    ax.scatter(x_by_ev[i], y_by_phi[i], s=800, c=[color], zorder=3, edgecolors='k')
    ax.annotate(labels_en[i], (x_by_ev[i], y_by_phi[i]), ha='center', va='center', fontsize=8, fontweight='bold')

ax.set_xlim(0, 1); ax.set_ylim(0, 1)
ax.set_xticks([]); ax.set_yticks([])
ax.annotate('Cause\n(upstream)', xy=(0.95, 0.9), fontsize=10, color=COLORS['upstream'], ha='right', fontweight='bold')
ax.annotate('Effect\n(downstream)', xy=(0.95, 0.15), fontsize=10, color=COLORS['downstream'], ha='right', fontweight='bold')
ax.text(0.5, 0.02, 'Vertical position = causal potential', ha='center', fontsize=11, style='italic', color='gray')

plt.tight_layout()
plt.savefig(f'{OUTPUT_DIR}/fig1_three_approaches.png', dpi=150, bbox_inches='tight')
plt.close()
print("  Saved: fig1_three_approaches.png")

# --- Figure 2: Magnetic Laplacian — Effect of q parameter ---
fig, axes = plt.subplots(1, 3, figsize=(18, 5))
for ax_idx, q in enumerate(q_values):
    ax = axes[ax_idx]
    evecs_q = results[q]['eigenvectors']
    evals_q = results[q]['eigenvalues']
    
    # Plot eigenvector components on complex plane (2nd eigenvector)
    ev2 = evecs_q[:, 1]
    for i in range(n_vars):
        ax.scatter(np.real(ev2[i]), np.imag(ev2[i]), s=300, zorder=3, edgecolors='k', c='steelblue')
        offset_x = 0.02 if np.real(ev2[i]) >= 0 else -0.02
        ax.annotate(labels_en[i], (np.real(ev2[i]) + offset_x, np.imag(ev2[i]) + 0.02),
                    fontsize=9, fontweight='bold', ha='left' if offset_x > 0 else 'right')
    
    # Draw unit circle for reference
    theta = np.linspace(0, 2 * np.pi, 100)
    max_r = max(np.abs(ev2)) * 1.1
    ax.plot(max_r * np.cos(theta), max_r * np.sin(theta), 'k--', alpha=0.2, lw=0.5)
    ax.axhline(0, color='gray', lw=0.5, alpha=0.5)
    ax.axvline(0, color='gray', lw=0.5, alpha=0.5)
    
    if q == 0:
        ax.set_title(f'q = {q} (no directionality)\nAll points on real axis', fontsize=12, fontweight='bold')
    elif q == 0.25:
        ax.set_title(f'q = {q} (max directionality)\nPhase = causal direction', fontsize=12, fontweight='bold')
    else:
        ax.set_title(f'q = {q} (partial directionality)', fontsize=12, fontweight='bold')
    
    ax.set_xlabel('Real part', fontsize=11)
    ax.set_ylabel('Imaginary part', fontsize=11)
    ax.set_aspect('equal')

plt.suptitle('Figure 2: Magnetic Laplacian Eigenvector (2nd) on Complex Plane\n'
             'Phase angle encodes causal direction; q controls direction sensitivity',
             fontsize=14, fontweight='bold', y=1.05)
plt.tight_layout()
plt.savefig(f'{OUTPUT_DIR}/fig2_magnetic_laplacian_q.png', dpi=150, bbox_inches='tight')
plt.close()
print("  Saved: fig2_magnetic_laplacian_q.png")

# --- Figure 3: Hodge Decomposition Pie Chart + Potential Bar ---
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# Panel A: Pie chart
ax = axes[0]
sizes = [grad_e, curl_e]
labels_pie = [f'Gradient (Causal)\n{100*grad_e/total_e:.1f}%',
              f'Curl (Feedback)\n{100*curl_e/total_e:.1f}%']
colors_pie = [COLORS['gradient'], COLORS['curl']]
wedges, texts = ax.pie(sizes, labels=labels_pie, colors=colors_pie,
                        startangle=90, textprops={'fontsize': 12, 'fontweight': 'bold'})
ax.set_title('(A) Hodge Decomposition\nof Information Flow', fontsize=13, fontweight='bold')

# Panel B: Causal potential bar chart
ax = axes[1]
order = np.argsort(phi)[::-1]
bars = ax.barh(range(n_vars), phi[order],
               color=[plt.cm.RdYlBu(1 - i / (n_vars - 1)) for i in range(n_vars)],
               edgecolor='k')
ax.set_yticks(range(n_vars))
ax.set_yticklabels([labels_en[i] for i in order], fontsize=12, fontweight='bold')
ax.set_xlabel('Causal Potential (phi)', fontsize=12)
ax.set_title('(B) Causal Potential\n(higher = more upstream)', fontsize=13, fontweight='bold')
ax.axvline(0, color='k', lw=0.5)

# Add annotations
ax.annotate('CAUSE\nside', xy=(ax.get_xlim()[1] * 0.8, 0.3), fontsize=11,
            color=COLORS['upstream'], fontweight='bold', ha='center')
ax.annotate('EFFECT\nside', xy=(ax.get_xlim()[0] * 0.5 if ax.get_xlim()[0] < 0 else ax.get_xlim()[1] * 0.2, n_vars - 1.3),
            fontsize=11, color=COLORS['downstream'], fontweight='bold', ha='center')

plt.tight_layout()
plt.savefig(f'{OUTPUT_DIR}/fig3_hodge_decomposition.png', dpi=150, bbox_inches='tight')
plt.close()
print("  Saved: fig3_hodge_decomposition.png")

# --- Figure 4: Three Methods Comparison ---
fig, ax = plt.subplots(figsize=(14, 7))

# For each variable pair, compare LiNGAM direction, SCD, and Hodge gradient
pairs = []
lingam_dirs = []
scd_dirs = []
hodge_dirs = []

for i in range(n_vars):
    for j in range(i+1, n_vars):
        pair_label = f"{labels_en[i]}\nvs\n{labels_en[j]}"
        pairs.append(pair_label)
        
        # LiNGAM direction: positive if i->j (B[j,i] > 0)
        lingam_d = B_lingam[j, i] - B_lingam[i, j]
        lingam_dirs.append(np.sign(lingam_d) if abs(lingam_d) > 0.01 else 0)
        
        # SCD direction (using antisymmetric sin-based formula)
        scd_dirs.append(np.sign(SCD_matrix[i, j]) if abs(SCD_matrix[i, j]) > 1e-6 else 0)
        
        # Hodge direction: phi difference
        hodge_d = phi[i] - phi[j]  # positive means i is more upstream
        hodge_dirs.append(np.sign(hodge_d) if abs(hodge_d) > 0.001 else 0)

x = np.arange(len(pairs))
width = 0.25

bars1 = ax.bar(x - width, lingam_dirs, width, label='LiNGAM', color=COLORS['lingam'], alpha=0.8, edgecolor='k')
bars2 = ax.bar(x, scd_dirs, width, label='Spectral (SCD)', color=COLORS['spectral'], alpha=0.8, edgecolor='k')
bars3 = ax.bar(x + width, hodge_dirs, width, label='Hodge Potential', color=COLORS['hodge'], alpha=0.8, edgecolor='k')

ax.set_ylabel('Direction (+1 = first->second, -1 = second->first)', fontsize=11)
ax.set_title('Figure 4: Causal Direction Comparison\nLiNGAM vs Spectral (SCD) vs Hodge Potential',
             fontsize=14, fontweight='bold')
ax.set_xticks(x)
ax.set_xticklabels(pairs, fontsize=8, ha='center')
ax.legend(fontsize=12, loc='upper right')
ax.axhline(0, color='k', lw=0.5)
ax.set_ylim(-1.5, 1.5)

# Highlight agreements
for idx in range(len(pairs)):
    if lingam_dirs[idx] == scd_dirs[idx] == hodge_dirs[idx] and lingam_dirs[idx] != 0:
        ax.axvspan(idx - 0.4, idx + 0.4, alpha=0.1, color='green')

ax.text(0.01, 0.01, 'Green shading = all three methods agree on direction',
        transform=ax.transAxes, fontsize=10, style='italic', color='green')

plt.tight_layout()
plt.savefig(f'{OUTPUT_DIR}/fig4_direction_comparison.png', dpi=150, bbox_inches='tight')
plt.close()
print("  Saved: fig4_direction_comparison.png")

# --- Figure 5: Hill Criteria Coverage Radar ---
fig, axes = plt.subplots(1, 3, figsize=(18, 6), subplot_kw=dict(polar=True))

hill_labels = ['H1\nStrength', 'H2\nConsistency', 'H3\nSpecificity',
               'H4\nTemporality', 'H5\nGradient', 'H6\nPlausibility',
               'H7\nCoherence', 'H8\nExperiment', 'H9\nAnalogy']
n_hill = len(hill_labels)
angles = np.linspace(0, 2 * np.pi, n_hill, endpoint=False).tolist()
angles += angles[:1]

# Scores (0-3 scale: 0=none, 1=weak, 2=partial, 3=strong)
lingam_scores = [3, 1, 3, 0, 1, 0, 0, 2, 0] + [3]
utility_scores = [2, 1, 1, 2, 2, 3, 3, 0, 3] + [2]
ecd_scores = [3, 3, 3, 2, 2, 2, 3, 2, 2] + [3]

methods = [
    ('DirectLiNGAM', lingam_scores, COLORS['lingam']),
    ('Spectral Causality', utility_scores, COLORS['spectral']),
    ('Ensemble (HIKONE)', ecd_scores, COLORS['hodge']),
]

for ax_idx, (name, scores, color) in enumerate(methods):
    ax = axes[ax_idx]
    ax.plot(angles, scores, 'o-', linewidth=2, color=color, markersize=8)
    ax.fill(angles, scores, alpha=0.25, color=color)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(hill_labels, fontsize=9)
    ax.set_ylim(0, 3)
    ax.set_yticks([1, 2, 3])
    ax.set_yticklabels(['weak', 'partial', 'strong'], fontsize=8)
    ax.set_title(name, fontsize=14, fontweight='bold', pad=20)

plt.suptitle("Hill's 9 Criteria Coverage — Three Approaches",
             fontsize=14, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(f'{OUTPUT_DIR}/fig5_hill_radar.png', dpi=150, bbox_inches='tight')
plt.close()
print("  Saved: fig5_hill_radar.png")

print("\n" + "=" * 60)
print("All figures generated successfully!")
print("=" * 60)

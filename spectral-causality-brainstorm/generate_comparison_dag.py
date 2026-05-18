"""
Generate side-by-side comparison:
  (A) LiNGAM DAG
  (B) Spectral Causality DCG (with domain knowledge, alpha=0.6)
  (C) Spectral Causality DCG (pure data-driven, alpha=0)

Shows that spectral causality allows cyclic structures (DCG),
and compares the effect of domain knowledge injection.

v2: Uses Directional Predictability Index (DPI) as data-driven component
    instead of |corr| (which was symmetric and produced 0 edges at α=0).
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import matplotlib.patches as mpatches
from scipy import linalg
from sklearn.preprocessing import StandardScaler
import lingam
import warnings
warnings.filterwarnings('ignore')

OUTPUT_DIR = '/home/ubuntu/repos/wip/spectral-causality-brainstorm/figures'


def load_data():
    url = "https://archive.ics.uci.edu/ml/machine-learning-databases/heart-disease/processed.cleveland.data"
    columns = ['age', 'sex', 'cp', 'trestbps', 'chol', 'fbs', 'restecg',
               'thalach', 'exang', 'oldpeak', 'slope', 'ca', 'thal', 'target']
    df = pd.read_csv(url, names=columns, na_values='?')
    df = df.dropna()

    clinical_vars = ['age', 'trestbps', 'chol', 'thalach', 'oldpeak']
    labels = ['Age', 'RestBP', 'Chol', 'MaxHR', 'STDep']

    X_raw = df[clinical_vars].astype(float).values
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X_raw)
    return X_scaled, X_raw, labels, clinical_vars


def run_lingam(X_scaled):
    model = lingam.DirectLiNGAM()
    model.fit(X_scaled)
    return model.adjacency_matrix_, model.causal_order_


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


def compute_scd_matrix(evecs, evals, n):
    SCD = np.zeros((n, n))
    SCC = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                for k in range(len(evals)):
                    amp_i = np.abs(evecs[i, k])
                    amp_j = np.abs(evecs[j, k])
                    phase_diff = np.angle(evecs[i, k]) - np.angle(evecs[j, k])
                    w = evals[k]
                    SCD[i, j] += w * amp_i * amp_j * np.sin(phase_diff)
                    SCC[i, j] += w * amp_i * amp_j * np.cos(phase_diff)
    return SCD, SCC


def hodge_decomposition(flow_matrix):
    n = flow_matrix.shape[0]
    edges = []
    flows = []
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
    BtB[0, :] = 0
    BtB[0, 0] = 1
    Btf[0] = 0
    phi = np.linalg.solve(BtB, Btf)
    gradient_flows = B @ phi
    curl_flows = flows - gradient_flows
    total_energy = np.sum(flows**2)
    gradient_energy = np.sum(gradient_flows**2)
    curl_energy = np.sum(curl_flows**2)
    return -phi, gradient_energy, curl_energy, total_energy


def compute_dpi(X_scaled, X_raw):
    """
    Compute Directional Predictability Index (DPI) — inherently asymmetric
    data-driven component replacing symmetric |corr|.
    
    DPI(i->j) = |corr(Xi,Xj)| * (1 + gamma * A_bar(i,j))
    """
    n = X_scaled.shape[1]
    corr_base = np.abs(np.corrcoef(X_scaled.T))
    np.fill_diagonal(corr_base, 0)
    
    # Component 1: Regression on raw data
    M_reg = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                cov_ij = np.cov(X_raw[:, i], X_raw[:, j])[0, 1]
                var_i = np.var(X_raw[:, i], ddof=1)
                M_reg[i, j] = abs(cov_ij / (var_i + 1e-10))
    
    # Component 2: ANM residual independence (HSIC)
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
    
    # Component 3: Conditional entropy reduction (kNN)
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
    
    # Combine: normalized asymmetry average
    def normalize_asym(M):
        A = M - M.T
        mx = np.max(np.abs(A))
        return A / (mx + 1e-10)
    
    A_combined = (normalize_asym(M_reg) + normalize_asym(M_anm) + normalize_asym(M_ent)) / 3.0
    
    gamma = 1.0
    DPI = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI[i, j] = corr_base[i, j] * (1 + gamma * A_combined[i, j])
    return DPI


def run_spectral(X_scaled, alpha, DPI=None):
    n = X_scaled.shape[1]
    clinical_influence = np.array([
        [0.0, 0.6, 0.4, 0.5, 0.3],
        [0.1, 0.0, 0.2, 0.3, 0.4],
        [0.1, 0.3, 0.0, 0.1, 0.3],
        [0.1, 0.2, 0.1, 0.0, 0.5],
        [0.0, 0.1, 0.0, 0.2, 0.0],
    ])
    if DPI is None:
        data_component = np.abs(np.corrcoef(X_scaled.T))
        np.fill_diagonal(data_component, 0)
    else:
        data_component = DPI
    utility_matrix = alpha * clinical_influence + (1 - alpha) * data_component
    asymmetry = utility_matrix - utility_matrix.T

    L_mag = build_magnetic_laplacian(utility_matrix, asymmetry, q=0.25)
    eigenvalues, eigenvectors = linalg.eigh(L_mag)
    SCD, SCC = compute_scd_matrix(eigenvectors, eigenvalues, n)
    phi, grad_e, curl_e, total_e = hodge_decomposition(asymmetry)

    grad_ratio = grad_e / total_e if total_e > 1e-12 else 0.0
    curl_ratio = curl_e / total_e if total_e > 1e-12 else 0.0
    return SCD, SCC, phi, grad_ratio, curl_ratio


def draw_graph(ax, labels, edges, title, subtitle, is_dag=True):
    """
    Draw a directed graph.
    edges: list of (src, dst, weight) tuples
    """
    n = len(labels)

    # Position nodes in a circle for DCG, or layered for DAG
    if is_dag:
        # Sort by causal depth (edges go top to bottom)
        # Compute a simple depth based on edges
        depth = {i: 0 for i in range(n)}
        changed = True
        while changed:
            changed = False
            for src, dst, w in edges:
                new_depth = depth[src] + 1
                if new_depth > depth[dst]:
                    depth[dst] = new_depth
                    changed = True
        max_depth = max(depth.values()) if depth else 1

        # Group by depth level
        levels = {}
        for node, d in depth.items():
            levels.setdefault(d, []).append(node)

        pos = {}
        for d, nodes in levels.items():
            y = (max_depth - d) * 1.4
            width = len(nodes)
            for idx, node in enumerate(nodes):
                x = (idx - (width - 1) / 2) * 1.8
                pos[node] = (x, y)
    else:
        # Circle layout for DCG
        angles = np.linspace(np.pi/2, np.pi/2 + 2*np.pi, n, endpoint=False)
        radius = 1.8
        pos = {}
        for i in range(n):
            pos[i] = (radius * np.cos(angles[i]), radius * np.sin(angles[i]))

    # Draw edges
    for src, dst, w in edges:
        color = '#1976D2' if w > 0 else '#D32F2F'
        width = max(abs(w) * 8, 1.0)
        alpha_val = min(abs(w) * 3, 0.9)

        x1, y1 = pos[src]
        x2, y2 = pos[dst]
        dx = x2 - x1
        dy = y2 - y1
        dist = np.sqrt(dx**2 + dy**2)
        if dist < 0.01:
            continue
        shrink = 0.45 / dist
        sx1 = x1 + dx * shrink
        sy1 = y1 + dy * shrink
        sx2 = x2 - dx * shrink
        sy2 = y2 - dy * shrink

        # Check for bidirectional edge — use more curve
        has_reverse = any(s == dst and d == src for s, d, _ in edges)
        rad = 0.2 if has_reverse else 0.08

        ax.annotate("",
            xy=(sx2, sy2), xytext=(sx1, sy1),
            arrowprops=dict(
                arrowstyle='-|>',
                color=color,
                lw=width,
                alpha=alpha_val,
                connectionstyle=f'arc3,rad={rad}',
                mutation_scale=16,
            ),
            zorder=1
        )

        # Edge weight label
        sign = "+" if w > 0 else "\u2212"
        mid_x = (x1 + x2) / 2
        mid_y = (y1 + y2) / 2
        perp_x = -dy / dist * 0.22
        perp_y = dx / dist * 0.22
        if has_reverse:
            perp_x *= 1.8
            perp_y *= 1.8
        ax.text(mid_x + perp_x, mid_y + perp_y, f"{sign}{abs(w):.2f}",
                fontsize=7, color=color, fontweight='bold',
                ha='center', va='center',
                bbox=dict(boxstyle='round,pad=0.12', facecolor='white',
                          edgecolor=color, alpha=0.85, lw=0.5))

    # Draw nodes
    box_w = 1.0
    box_h = 0.45
    for i in range(n):
        x, y = pos[i]
        box = FancyBboxPatch(
            (x - box_w/2, y - box_h/2), box_w, box_h,
            boxstyle="round,pad=0.08",
            facecolor='#1565C0', edgecolor='white', linewidth=2, zorder=3
        )
        ax.add_patch(box)
        ax.text(x, y, labels[i], ha='center', va='center',
               fontsize=9, fontweight='bold', color='white', zorder=4)

    ax.set_title(title, fontsize=12, fontweight='bold', pad=10)
    ax.text(0.5, -0.02, subtitle, transform=ax.transAxes,
            fontsize=8, ha='center', va='top', color='#666666', style='italic')
    ax.set_aspect('equal')
    ax.axis('off')

    # Auto-scale
    all_x = [pos[i][0] for i in range(n)]
    all_y = [pos[i][1] for i in range(n)]
    margin = 1.2
    ax.set_xlim(min(all_x) - margin, max(all_x) + margin)
    ax.set_ylim(min(all_y) - margin, max(all_y) + margin)


def main():
    X_scaled, X_raw, labels, clinical_vars = load_data()
    n = len(labels)

    # Compute DPI (asymmetric data-driven component)
    print("Computing DPI (Directional Predictability Index)...")
    DPI = compute_dpi(X_scaled, X_raw)
    print("  DPI asymmetry norm:", np.linalg.norm(DPI - DPI.T, 'fro'))

    # (A) LiNGAM
    B_lingam, causal_order = run_lingam(X_scaled)
    lingam_edges = []
    for i in range(n):
        for j in range(n):
            if abs(B_lingam[i][j]) > 0.05:
                lingam_edges.append((j, i, B_lingam[i][j]))

    # (B) Spectral with domain knowledge (alpha=0.6) — uses DPI as base
    SCD_06, SCC_06, phi_06, grad_r_06, curl_r_06 = run_spectral(X_scaled, alpha=0.6, DPI=DPI)

    # (C) Spectral pure data-driven (alpha=0) — DPI provides directional signal
    SCD_00, SCC_00, phi_00, grad_r_00, curl_r_00 = run_spectral(X_scaled, alpha=0.0, DPI=DPI)

    # Build spectral edges: use SCD for direction and magnitude
    def spectral_edges(SCD, threshold=0.05):
        edges = []
        for i in range(n):
            for j in range(i+1, n):
                scd = SCD[i, j]
                if abs(scd) > threshold:
                    if scd > 0:
                        edges.append((i, j, scd))
                    else:
                        edges.append((j, i, -scd))
        return edges

    spec_edges_06 = spectral_edges(SCD_06)
    spec_edges_00 = spectral_edges(SCD_00)

    # Print results
    print("=" * 60)
    print("(A) LiNGAM DAG")
    print("=" * 60)
    for s, d, w in lingam_edges:
        print(f"  {labels[s]} -> {labels[d]}: {w:+.3f}")

    print(f"\n{'=' * 60}")
    print("(B) Spectral DCG (alpha=0.6, with domain knowledge + DPI)")
    print(f"    Gradient: {grad_r_06:.1%}, Curl: {curl_r_06:.1%}")
    print("=" * 60)
    for s, d, w in spec_edges_06:
        print(f"  {labels[s]} -> {labels[d]}: SCD={w:+.4f}")
    print(f"  Causal potential (phi): {dict(zip(labels, [f'{p:.3f}' for p in phi_06]))}")

    print(f"\n{'=' * 60}")
    print("(C) Spectral DCG (alpha=0, pure data-driven with DPI)")
    print(f"    Gradient: {grad_r_00:.1%}, Curl: {curl_r_00:.1%}")
    print("=" * 60)
    for s, d, w in spec_edges_00:
        print(f"  {labels[s]} -> {labels[d]}: SCD={w:+.4f}")
    print(f"  Causal potential (phi): {dict(zip(labels, [f'{p:.3f}' for p in phi_00]))}")

    # ========== Generate Figure ==========
    fig, axes = plt.subplots(1, 3, figsize=(18, 7))

    draw_graph(axes[0], labels, lingam_edges,
               '(A) LiNGAM (DAG)',
               'Assumption: linear, non-Gaussian, acyclic\nNo domain knowledge required',
               is_dag=True)

    draw_graph(axes[1], labels, spec_edges_06,
               '(B) Spectral Causality (DCG)\n\u03b1 = 0.6 (with domain knowledge)',
               f'Gradient: {grad_r_06:.0%}, Curl: {curl_r_06:.0%}\nCycles allowed via Hodge decomposition',
               is_dag=False)

    draw_graph(axes[2], labels, spec_edges_00,
               '(C) Spectral Causality (DCG)\n\u03b1 = 0 (pure data-driven, DPI)',
               f'Gradient: {grad_r_00:.0%}, Curl: {curl_r_00:.0%}\nDPI: regression + ANM + entropy',
               is_dag=False)

    # Legend
    pos_patch = mpatches.Patch(color='#1976D2', label='Positive direction')
    neg_patch = mpatches.Patch(color='#D32F2F', label='Negative direction')
    fig.legend(handles=[pos_patch, neg_patch], loc='lower center',
              ncol=2, fontsize=11, framealpha=0.9)

    fig.suptitle('Causal Structure Comparison: LiNGAM vs Spectral Causality\nUCI Heart Disease Data (n = 297)',
                fontsize=15, fontweight='bold', y=1.02)

    plt.tight_layout()
    output_path = f'{OUTPUT_DIR}/fig7_lingam_vs_spectral.png'
    fig.savefig(output_path, dpi=200, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"\nSaved: {output_path}")
    return output_path


if __name__ == '__main__':
    main()

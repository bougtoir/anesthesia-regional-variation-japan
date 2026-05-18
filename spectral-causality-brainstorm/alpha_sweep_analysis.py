"""
Alpha sweep analysis: How does domain knowledge injection (α) affect DAG-ness?

Sweeps α from 0 to 1 and computes:
- r_gradient (Hodge gradient ratio = DAG-ness)
- Number of significant SCD edges
- LiNGAM direction agreement rate
- Spectral gap (Fiedler value)
- Effective rank of asymmetry matrix

Goal: Find the DAG transition point α* where r_gradient sharply increases.
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
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


def compute_dpi(X_scaled, X_raw):
    """Compute Directional Predictability Index (asymmetric data component)."""
    n = X_scaled.shape[1]
    corr_base = np.abs(np.corrcoef(X_scaled.T))
    np.fill_diagonal(corr_base, 0)
    
    # Regression on raw data
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
    A_combined = (normalize_asym(M_reg) + normalize_asym(M_anm) + normalize_asym(M_ent)) / 3.0
    gamma = 1.0
    DPI = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI[i, j] = corr_base[i, j] * (1 + gamma * A_combined[i, j])
    return DPI


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


def compute_scd_matrix(W, asymmetry, q=0.25):
    n = W.shape[0]
    L_mag = build_magnetic_laplacian(W, asymmetry, q=q)
    eigenvalues, eigenvectors = linalg.eigh(L_mag)
    SCD = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                for k in range(n):
                    amp_i = np.abs(eigenvectors[i, k])
                    amp_j = np.abs(eigenvectors[j, k])
                    phase_diff = np.angle(eigenvectors[i, k]) - np.angle(eigenvectors[j, k])
                    SCD[i, j] += eigenvalues[k] * amp_i * amp_j * np.sin(phase_diff)
    return SCD, eigenvalues


def run_sweep(X_scaled, alphas, lingam_directions, DPI=None):
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

    results = []
    for alpha in alphas:
        utility_matrix = alpha * clinical_influence + (1 - alpha) * data_component
        asymmetry = utility_matrix - utility_matrix.T

        # Hodge decomposition
        phi, grad_e, curl_e, total_e = hodge_decomposition(asymmetry)
        r_gradient = grad_e / total_e if total_e > 1e-12 else 0

        # SCD and edges
        SCD, eigenvalues = compute_scd_matrix(utility_matrix, asymmetry)
        n_edges = 0
        n_agree = 0
        n_compared = 0
        for i in range(n):
            for j in range(i+1, n):
                if abs(SCD[i, j]) > 0.05:
                    n_edges += 1
                    if (i, j) in lingam_directions:
                        n_compared += 1
                        lingam_dir = lingam_directions[(i, j)]
                        spectral_dir = 1 if SCD[i, j] > 0 else -1
                        if lingam_dir == spectral_dir:
                            n_agree += 1

        agreement = n_agree / n_compared if n_compared > 0 else float('nan')

        # Asymmetry Frobenius norm
        asym_norm = np.linalg.norm(asymmetry, 'fro')

        # Spectral gap (Fiedler value)
        fiedler = eigenvalues[1] if len(eigenvalues) > 1 else 0

        # Effective number of significant edges
        scd_values = []
        for i in range(n):
            for j in range(i+1, n):
                scd_values.append(abs(SCD[i, j]))
        scd_values = sorted(scd_values, reverse=True)

        results.append({
            'alpha': alpha,
            'r_gradient': r_gradient,
            'r_curl': 1 - r_gradient,
            'n_edges': n_edges,
            'agreement': agreement,
            'asym_norm': asym_norm,
            'fiedler': fiedler,
            'max_scd': max(scd_values) if scd_values else 0,
            'phi_range': max(phi) - min(phi),
        })

    return pd.DataFrame(results)


def main():
    X_scaled, X_raw, labels, clinical_vars = load_data()
    n = len(labels)

    # LiNGAM baseline
    model = lingam.DirectLiNGAM()
    model.fit(X_scaled)
    B = model.adjacency_matrix_

    # Compute DPI
    print("Computing DPI...")
    DPI = compute_dpi(X_scaled, X_raw)
    print(f"  DPI asymmetry norm: {np.linalg.norm(DPI - DPI.T, 'fro'):.4f}")

    # LiNGAM directions for comparison
    lingam_directions = {}
    for i in range(n):
        for j in range(n):
            if abs(B[i][j]) > 0.05:
                src, dst = j, i
                if src < dst:
                    lingam_directions[(src, dst)] = 1   # src -> dst
                else:
                    lingam_directions[(dst, src)] = -1  # dst -> src (i.e., reversed)

    # Sweep alpha from 0 to 1 (using DPI)
    alphas = np.linspace(0, 1, 101)
    df = run_sweep(X_scaled, alphas, lingam_directions, DPI=DPI)

    # Print key values
    print("=" * 80)
    print("Alpha Sweep Results")
    print("=" * 80)
    for _, row in df.iterrows():
        if any(abs(row['alpha'] - v) < 1e-6 for v in [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0]):
            print(f"  α={row['alpha']:.1f}: r_grad={row['r_gradient']:.3f}, "
                  f"edges={row['n_edges']:.0f}, agree={row['agreement']:.2f}, "
                  f"asym_norm={row['asym_norm']:.4f}, fiedler={row['fiedler']:.4f}, "
                  f"phi_range={row['phi_range']:.4f}")

    # Find transition point: maximum second derivative of r_gradient
    r_grad = df['r_gradient'].values
    d1 = np.gradient(r_grad, df['alpha'].values)
    d2 = np.gradient(d1, df['alpha'].values)
    transition_idx = np.argmax(d2[5:-5]) + 5  # skip edges
    alpha_star = df['alpha'].values[transition_idx]
    r_star = r_grad[transition_idx]
    print(f"\n  DAG transition point: α* ≈ {alpha_star:.2f} (r_gradient = {r_star:.3f})")
    print(f"  Max d²r/dα² at α = {alpha_star:.2f}")

    # Also find where r_gradient crosses 0.5 (50% DAG)
    cross_50 = None
    for i in range(len(r_grad) - 1):
        if r_grad[i] < 0.5 and r_grad[i+1] >= 0.5:
            cross_50 = df['alpha'].values[i]
            break
    if cross_50:
        print(f"  r_gradient crosses 0.50 at α ≈ {cross_50:.2f}")

    # =========== Generate Figure ===========
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))

    # (A) r_gradient vs alpha — the main plot
    ax = axes[0, 0]
    ax.plot(df['alpha'], df['r_gradient'], 'b-', lw=2.5, label='$r_{\\mathrm{gradient}}$ (DAG-ness)')
    ax.fill_between(df['alpha'], df['r_gradient'], alpha=0.15, color='blue')
    ax.axhline(y=0.5, color='gray', ls='--', lw=1, alpha=0.7, label='50% threshold')
    ax.axvline(x=alpha_star, color='red', ls='--', lw=1.5, alpha=0.7,
              label=f'Transition α* ≈ {alpha_star:.2f}')
    if cross_50:
        ax.axvline(x=cross_50, color='orange', ls=':', lw=1.5, alpha=0.7,
                  label=f'50% crossing α ≈ {cross_50:.2f}')
    ax.set_xlabel('α (domain knowledge weight)', fontsize=12)
    ax.set_ylabel('$r_{\\mathrm{gradient}}$', fontsize=12)
    ax.set_title('(A) DAG-ness vs Domain Knowledge', fontsize=13, fontweight='bold')
    ax.legend(fontsize=9, loc='lower right')
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.grid(True, alpha=0.3)

    # (B) Number of edges and agreement
    ax = axes[0, 1]
    ax2 = ax.twinx()
    l1, = ax.plot(df['alpha'], df['n_edges'], 'g-', lw=2, label='# significant edges')
    l2, = ax2.plot(df['alpha'], df['agreement'], 'r-', lw=2, label='LiNGAM agreement')
    ax.set_xlabel('α', fontsize=12)
    ax.set_ylabel('Number of edges (|SCD| > 0.05)', fontsize=11, color='green')
    ax2.set_ylabel('Direction agreement with LiNGAM', fontsize=11, color='red')
    ax.set_title('(B) Edge Count & LiNGAM Agreement', fontsize=13, fontweight='bold')
    ax.legend(handles=[l1, l2], fontsize=9, loc='center right')
    ax.set_xlim(0, 1)
    ax.grid(True, alpha=0.3)

    # (C) Asymmetry norm and Fiedler value
    ax = axes[1, 0]
    ax.plot(df['alpha'], df['asym_norm'], 'm-', lw=2, label='Asymmetry ‖A−Aᵀ‖_F')
    ax.plot(df['alpha'], df['phi_range'] * 3, 'c-', lw=2, label='φ range (×3)')
    ax.set_xlabel('α', fontsize=12)
    ax.set_ylabel('Norm / Range', fontsize=12)
    ax.set_title('(C) Asymmetry Strength & Potential Range', fontsize=13, fontweight='bold')
    ax.legend(fontsize=9)
    ax.set_xlim(0, 1)
    ax.grid(True, alpha=0.3)

    # (D) Phase diagram: r_gradient vs asymmetry norm
    ax = axes[1, 1]
    scatter = ax.scatter(df['asym_norm'], df['r_gradient'],
                        c=df['alpha'], cmap='viridis', s=30, alpha=0.8)
    cbar = plt.colorbar(scatter, ax=ax)
    cbar.set_label('α', fontsize=11)
    ax.set_xlabel('Asymmetry norm ‖A−Aᵀ‖_F', fontsize=12)
    ax.set_ylabel('$r_{\\mathrm{gradient}}$ (DAG-ness)', fontsize=12)
    ax.set_title('(D) Phase Diagram: Asymmetry → DAG-ness', fontsize=13, fontweight='bold')
    ax.axhline(y=0.5, color='gray', ls='--', lw=1, alpha=0.5)
    ax.grid(True, alpha=0.3)

    # Annotate key points
    for a_val in [0, 0.3, 0.6, 1.0]:
        row = df.iloc[(df['alpha'] - a_val).abs().idxmin()]
        ax.annotate(f'α={a_val}', (row['asym_norm'], row['r_gradient']),
                   fontsize=8, fontweight='bold',
                   xytext=(5, 5), textcoords='offset points')

    fig.suptitle('DAG Transition Analysis: Effect of Domain Knowledge (α) on Causal Structure\n'
                'UCI Heart Disease Data (n=297)',
                fontsize=14, fontweight='bold', y=1.02)

    plt.tight_layout()
    output_path = f'{OUTPUT_DIR}/fig8_alpha_sweep.png'
    fig.savefig(output_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f"\nSaved: {output_path}")

    return df, alpha_star


if __name__ == '__main__':
    df, alpha_star = main()

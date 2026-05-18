#!/usr/bin/env python3
"""
Final asymmetric data-driven component: Directional Predictability Index (DPI).

Combines three complementary asymmetric measures:
1. Regression asymmetry on raw data (variance-based direction)
2. ANM residual independence (nonlinear causal direction)  
3. Conditional entropy reduction (information-theoretic direction)

Also evaluates: direction-voting ensemble, Kendall tau of causal orders.
"""

import numpy as np
import pandas as pd
from scipy import linalg, stats
from sklearn.preprocessing import StandardScaler
import lingam
import warnings
warnings.filterwarnings('ignore')


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
    return X_raw, X_scaled, labels


# ============================================================
# Component 1: Regression asymmetry (raw data)
# ============================================================
def regression_component(X_raw):
    n = X_raw.shape[1]
    M = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                cov_ij = np.cov(X_raw[:, i], X_raw[:, j])[0, 1]
                var_i = np.var(X_raw[:, i], ddof=1)
                M[i, j] = abs(cov_ij / (var_i + 1e-10))
    return M


# ============================================================
# Component 2: ANM residual independence (HSIC)
# ============================================================
def anm_component(X):
    n = X.shape[1]
    M = np.zeros((n, n))
    
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
    
    for i in range(n):
        for j in range(n):
            if i != j:
                beta = np.dot(X[:, i], X[:, j]) / (np.dot(X[:, i], X[:, i]) + 1e-10)
                res = X[:, j] - beta * X[:, i]
                M[i, j] = neg_hsic(X[:, i], res)
    return M


# ============================================================
# Component 3: Conditional entropy reduction (kNN)
# ============================================================
def entropy_component(X, k=5):
    n = X.shape[1]
    M = np.zeros((n, n))
    
    def knn_ent(x, k=5):
        x = np.sort(x)
        dists = np.array([np.sort(np.abs(x - x[i]))[min(k, len(x)-1)] for i in range(len(x))])
        dists = dists[dists > 0]
        return np.mean(np.log(dists + 1e-10))
    
    for i in range(n):
        for j in range(n):
            if i != j:
                beta = np.dot(X[:, i], X[:, j]) / (np.dot(X[:, i], X[:, i]) + 1e-10)
                res = X[:, j] - beta * X[:, i]
                M[i, j] = max(0, knn_ent(X[:, j], k) - knn_ent(res, k))
    return M


# ============================================================
# DPI: Direction-Voting Ensemble
# ============================================================
def direction_voting_dpi(M_list, labels):
    """
    For each pair (i,j), each method votes on direction.
    Direction = sign(M(i,j) - M(j,i)) from the raw M matrix.
    DPI(i,j) = |avg_vote| * avg_strength.
    """
    n = M_list[0].shape[0]
    DPI = np.zeros((n, n))
    K = len(M_list)
    
    for i in range(n):
        for j in range(n):
            if i != j:
                votes = []
                strengths = []
                for M in M_list:
                    asym = M[i, j] - M[j, i]
                    if abs(asym) > 1e-10:
                        votes.append(np.sign(asym))
                        # Normalize strength by max asymmetry in this matrix
                        A = M - M.T
                        max_a = np.max(np.abs(A)) + 1e-10
                        strengths.append(abs(asym) / max_a)
                
                if votes:
                    avg_vote = np.mean(votes)
                    avg_strength = np.mean(strengths)
                    # Consensus-weighted: stronger signal when methods agree
                    DPI[i, j] = avg_vote * avg_strength * abs(avg_vote)
                    # Make positive: DPI(i,j) = strength of i→j signal
                    # If avg_vote > 0, i→j; if < 0, j→i
    
    # Convert to utility-like matrix (positive values = strength of association)
    # Keep asymmetry but add base correlation for magnitude
    return DPI


def rank_based_dpi(M_list, labels):
    """
    Rank-based fusion: for each method, rank all (i,j) asymmetry values.
    Then average ranks across methods. More robust than value-based averaging.
    """
    n = M_list[0].shape[0]
    
    # Collect asymmetry values from each method
    combined_asym = np.zeros((n, n))
    
    for M in M_list:
        A = M - M.T
        # Rank the upper triangle asymmetry values
        upper_vals = []
        upper_idx = []
        for i in range(n):
            for j in range(i+1, n):
                upper_vals.append(A[i, j])
                upper_idx.append((i, j))
        
        # Convert to ranks
        ranks = stats.rankdata(upper_vals) / len(upper_vals)  # normalized to [0, 1]
        
        for idx, (i, j) in enumerate(upper_idx):
            combined_asym[i, j] += ranks[idx] - 0.5  # center at 0
            combined_asym[j, i] -= ranks[idx] - 0.5
    
    combined_asym /= len(M_list)
    
    # Convert asymmetry to utility matrix
    # U(i,j) = base_strength + direction_signal
    DPI = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI[i, j] = 0.5 + combined_asym[i, j]  # [0, 1] range
    
    return DPI


# ============================================================
# Evaluation
# ============================================================
def build_mag_lap(W, direction_matrix, q=0.25):
    n = W.shape[0]
    W_sym = (W + W.T) / 2
    H = np.zeros((n, n), dtype=complex)
    for i in range(n):
        for j in range(n):
            if i != j:
                d = np.sign(direction_matrix[i, j]) if abs(direction_matrix[i, j]) > 1e-10 else 0.0
                H[i, j] = W_sym[i, j] * np.exp(1j * 2 * np.pi * q * d)
    D = np.diag(np.real(np.sum(np.abs(H), axis=1)))
    D_inv = np.diag(1.0 / np.sqrt(np.diag(D) + 1e-10))
    return np.eye(n) - D_inv @ H @ D_inv


def hodge(flow):
    n = flow.shape[0]
    edges, flows = [], []
    for i in range(n):
        for j in range(i+1, n):
            edges.append((i, j))
            flows.append(flow[i, j])
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
    gf = B @ phi
    cf = flows - gf
    return -phi, np.sum(gf**2), np.sum(cf**2), np.sum(flows**2)


def scd_matrix(W, asym, q=0.25):
    n = W.shape[0]
    L = build_mag_lap(W, asym, q)
    evals, evecs = linalg.eigh(L)
    SCD = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                for k in range(n):
                    ai, aj = np.abs(evecs[i, k]), np.abs(evecs[j, k])
                    pd = np.angle(evecs[i, k]) - np.angle(evecs[j, k])
                    SCD[i, j] += evals[k] * ai * aj * np.sin(pd)
    return SCD


def full_eval(X, M, labels, B_lingam, lingam_order, name):
    n = X.shape[1]
    asym = M - M.T
    asym_norm = np.linalg.norm(asym, 'fro')
    
    phi, ge, ce, te = hodge(asym)
    rg = ge / te if te > 1e-12 else float('nan')
    
    SCD = scd_matrix(M, asym)
    scd_vals = np.abs(SCD[np.triu_indices(n, 1)])
    threshold = max(np.percentile(scd_vals, 40), 0.001)
    
    edges = []
    for i in range(n):
        for j in range(i+1, n):
            if abs(SCD[i, j]) > threshold:
                src, dst = (i, j) if SCD[i, j] > 0 else (j, i)
                edges.append((src, dst, abs(SCD[i, j])))
    
    # LiNGAM edge directions
    lingam_dirs = {}
    for i in range(n):
        for j in range(n):
            if abs(B_lingam[i][j]) > 0.05:
                src, dst = j, i
                pair = (min(src, dst), max(src, dst))
                lingam_dirs[pair] = 1 if src == pair[0] else -1
    
    n_agree, n_compared = 0, 0
    for src, dst, _ in edges:
        pair = (min(src, dst), max(src, dst))
        if pair in lingam_dirs:
            n_compared += 1
            spec_dir = 1 if src == pair[0] else -1
            if spec_dir == lingam_dirs[pair]:
                n_agree += 1
    
    agreement = n_agree / n_compared if n_compared > 0 else float('nan')
    
    # Hodge potential order
    phi_order = np.argsort(-phi)
    our_order = [labels[i] for i in phi_order]
    
    # Kendall tau with LiNGAM order
    lingam_rank = {v: i for i, v in enumerate(lingam_order)}
    our_rank = [lingam_rank[v] for v in our_order]
    tau, p_val = stats.kendalltau(list(range(n)), our_rank)
    
    # Direct edge direction comparison using asymmetry matrix (not SCD)
    n_direct_agree = 0
    n_direct_total = 0
    for i in range(n):
        for j in range(i+1, n):
            pair = (i, j)
            if pair in lingam_dirs:
                n_direct_total += 1
                asym_dir = 1 if asym[i, j] > 0 else -1 if asym[i, j] < 0 else 0
                if abs(asym[i, j]) > 1e-10 and asym_dir == lingam_dirs[pair]:
                    n_direct_agree += 1
    direct_agreement = n_direct_agree / n_direct_total if n_direct_total > 0 else float('nan')
    
    print(f"\n{'='*65}")
    print(f"  {name}")
    print(f"{'='*65}")
    print(f"  Asymmetry norm:      {asym_norm:.6f}")
    rg_str = f"{rg:.3f}" if not np.isnan(rg) else "NaN"
    print(f"  r_gradient (DAG):    {rg_str}")
    print(f"  Edges (SCD):         {len(edges)} (threshold={threshold:.4f})")
    ag_str = f"{n_agree}/{n_compared} = {agreement:.0%}" if not np.isnan(agreement) else "N/A"
    print(f"  LiNGAM agree (SCD):  {ag_str}")
    dag_str = f"{n_direct_agree}/{n_direct_total} = {direct_agreement:.0%}" if not np.isnan(direct_agreement) else "N/A"
    print(f"  LiNGAM agree (raw):  {dag_str}")
    print(f"  Kendall tau:         {tau:.3f} (p={p_val:.3f})")
    print(f"  Hodge order:         {' > '.join(our_order)}")
    print(f"  LiNGAM order:        {' > '.join(lingam_order)}")
    print(f"  Edges:")
    for src, dst, s in sorted(edges, key=lambda x: -x[2]):
        pair = (min(src, dst), max(src, dst))
        m = ""
        if pair in lingam_dirs:
            d = 1 if src == pair[0] else -1
            m = " MATCH" if d == lingam_dirs[pair] else " MISS"
        print(f"    {labels[src]:>6} → {labels[dst]:<6}: |SCD|={s:.4f}{m}")
    
    return {
        'name': name, 'asym_norm': asym_norm, 'r_gradient': rg,
        'n_edges': len(edges), 'agreement': agreement,
        'direct_agreement': direct_agreement,
        'tau': tau, 'tau_p': p_val,
        'phi_order': our_order, 'M': M, 'asymmetry': asym,
        'SCD': SCD, 'phi': phi, 'edges': edges,
    }


def main():
    X_raw, X_scaled, labels = load_data()
    n = len(labels)
    
    model = lingam.DirectLiNGAM()
    model.fit(X_scaled)
    B = model.adjacency_matrix_
    lingam_order = [labels[i] for i in model.causal_order_]
    
    print(f"LiNGAM order: {' > '.join(lingam_order)}")
    print("LiNGAM edges:")
    for i in range(n):
        for j in range(n):
            if abs(B[i][j]) > 0.05:
                print(f"  {labels[j]} → {labels[i]}: {B[i][j]:+.3f}")
    
    # Compute components
    M_reg = regression_component(X_raw)
    M_anm = anm_component(X_scaled)
    M_ent = entropy_component(X_scaled)
    
    # Individual evaluations
    r1 = full_eval(X_scaled, M_reg, labels, B, lingam_order, "C1: Regression β (raw)")
    r3 = full_eval(X_scaled, M_anm, labels, B, lingam_order, "C3: ANM Score (HSIC)")
    r5 = full_eval(X_scaled, M_ent, labels, B, lingam_order, "C5: Conditional Entropy")
    
    # Direction-voting DPI
    M_vote = direction_voting_dpi([M_reg, M_anm, M_ent], labels)
    # Use |corr| as base strength, vote for direction
    corr_base = np.abs(np.corrcoef(X_scaled.T))
    np.fill_diagonal(corr_base, 0)
    # DPI = base_strength * (1 + direction_signal)
    DPI_vote = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI_vote[i, j] = corr_base[i, j] * (1 + M_vote[i, j])
    rv = full_eval(X_scaled, DPI_vote, labels, B, lingam_order, "DPI (Direction Voting)")
    
    # Rank-based DPI
    M_rank = rank_based_dpi([M_reg, M_anm, M_ent], labels)
    rr = full_eval(X_scaled, M_rank, labels, B, lingam_order, "DPI (Rank-based)")
    
    # ============================================================
    # NEW: Proper DPI definition for the model
    # ============================================================
    # DPI(i→j) = corr_base * (1 + γ * normalized_asymmetry)
    # where normalized_asymmetry averages rank-normalized directions
    # from regression, ANM, and entropy
    
    def normalize_asym(M):
        """Normalize asymmetry matrix to [-1, 1] by max abs value."""
        A = M - M.T
        mx = np.max(np.abs(A))
        return A / (mx + 1e-10)
    
    A_reg = normalize_asym(M_reg)
    A_anm = normalize_asym(M_anm)
    A_ent = normalize_asym(M_ent)
    A_combined = (A_reg + A_anm + A_ent) / 3.0
    
    # Final DPI: symmetric base + asymmetric direction
    gamma = 1.0  # direction strength parameter
    DPI_final = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI_final[i, j] = corr_base[i, j] * (1 + gamma * A_combined[i, j])
    
    rf = full_eval(X_scaled, DPI_final, labels, B, lingam_order, "DPI Final: corr*(1+γ*A_combined)")
    
    # Sweep γ to find optimal
    print(f"\n\n{'='*65}")
    print("  γ sweep: DPI = |corr| * (1 + γ * A_combined)")
    print(f"{'='*65}")
    print(f"  {'γ':>5} {'r_grad':>7} {'Edges':>5} {'SCD%':>5} {'Raw%':>5} {'τ':>6}")
    print(f"  {'-'*40}")
    
    best_tau = -2
    best_gamma = 0
    for g in [0.1, 0.3, 0.5, 0.7, 1.0, 1.5, 2.0, 3.0, 5.0]:
        DPI_g = np.zeros((n, n))
        for i in range(n):
            for j in range(n):
                if i != j:
                    DPI_g[i, j] = corr_base[i, j] * (1 + g * A_combined[i, j])
        
        asym_g = DPI_g - DPI_g.T
        phi_g, ge_g, ce_g, te_g = hodge(asym_g)
        rg_g = ge_g / te_g if te_g > 1e-12 else float('nan')
        
        SCD_g = scd_matrix(DPI_g, asym_g)
        scd_vals_g = np.abs(SCD_g[np.triu_indices(n, 1)])
        thr_g = max(np.percentile(scd_vals_g, 40), 0.001)
        
        lingam_dirs = {}
        for i in range(n):
            for j in range(n):
                if abs(B[i][j]) > 0.05:
                    src, dst = j, i
                    pair = (min(src, dst), max(src, dst))
                    lingam_dirs[pair] = 1 if src == pair[0] else -1
        
        n_e, na_scd, nc_scd = 0, 0, 0
        for i in range(n):
            for j in range(i+1, n):
                if abs(SCD_g[i, j]) > thr_g:
                    n_e += 1
                    pair = (i, j)
                    if pair in lingam_dirs:
                        nc_scd += 1
                        d = 1 if SCD_g[i, j] > 0 else -1
                        if d == lingam_dirs[pair]:
                            na_scd += 1
        
        na_raw, nc_raw = 0, 0
        for i in range(n):
            for j in range(i+1, n):
                pair = (i, j)
                if pair in lingam_dirs:
                    nc_raw += 1
                    if abs(asym_g[i, j]) > 1e-10:
                        d = 1 if asym_g[i, j] > 0 else -1
                        if d == lingam_dirs[pair]:
                            na_raw += 1
        
        phi_ord = np.argsort(-phi_g)
        lingam_rank = {v: i for i, v in enumerate(lingam_order)}
        our_r = [lingam_rank[labels[idx]] for idx in phi_ord]
        tau_g, _ = stats.kendalltau(list(range(n)), our_r)
        
        if tau_g > best_tau:
            best_tau = tau_g
            best_gamma = g
        
        rg_s = f"{rg_g:.3f}" if not np.isnan(rg_g) else "NaN"
        scd_s = f"{na_scd}/{nc_scd}" if nc_scd > 0 else "N/A"
        raw_s = f"{na_raw}/{nc_raw}" if nc_raw > 0 else "N/A"
        print(f"  {g:5.1f} {rg_s:>7} {n_e:5d} {scd_s:>5} {raw_s:>5} {tau_g:6.3f}")
    
    print(f"\n  Best γ: {best_gamma} (τ = {best_tau:.3f})")
    
    # ============================================================
    # Final evaluation with best γ
    # ============================================================
    DPI_best = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i != j:
                DPI_best[i, j] = corr_base[i, j] * (1 + best_gamma * A_combined[i, j])
    
    print(f"\n\n{'='*65}")
    print(f"  FINAL MODEL: DPI = |corr| * (1 + {best_gamma} * A_combined)")
    print(f"{'='*65}")
    r_best = full_eval(X_scaled, DPI_best, labels, B, lingam_order, f"FINAL DPI (γ={best_gamma})")
    
    # Also compare with domain knowledge versions
    print(f"\n\n{'='*65}")
    print("  Comparison with domain knowledge (α sweep)")
    print(f"{'='*65}")
    
    clinical_influence = np.array([
        [0.0, 0.6, 0.4, 0.5, 0.3],
        [0.1, 0.0, 0.2, 0.3, 0.4],
        [0.1, 0.3, 0.0, 0.1, 0.3],
        [0.1, 0.2, 0.1, 0.0, 0.5],
        [0.0, 0.1, 0.0, 0.2, 0.0],
    ])
    
    print(f"\n  {'α':>5} {'r_grad':>7} {'Edges':>5} {'τ':>6} {'Order'}")
    print(f"  {'-'*60}")
    
    for alpha in [0, 0.01, 0.1, 0.3, 0.5, 0.6, 0.8, 1.0]:
        U = alpha * clinical_influence + (1 - alpha) * DPI_best
        asym_a = U - U.T
        phi_a, ge_a, ce_a, te_a = hodge(asym_a)
        rg_a = ge_a / te_a if te_a > 1e-12 else float('nan')
        
        phi_ord = np.argsort(-phi_a)
        lingam_rank = {v: i for i, v in enumerate(lingam_order)}
        our_r = [lingam_rank[labels[idx]] for idx in phi_ord]
        tau_a, _ = stats.kendalltau(list(range(n)), our_r)
        
        SCD_a = scd_matrix(U, asym_a)
        scd_v = np.abs(SCD_a[np.triu_indices(n, 1)])
        thr_a = max(np.percentile(scd_v, 40), 0.001)
        n_e = sum(1 for i in range(n) for j in range(i+1, n) if abs(SCD_a[i,j]) > thr_a)
        
        order = ' > '.join([labels[i] for i in phi_ord])
        rg_s = f"{rg_a:.3f}" if not np.isnan(rg_a) else "NaN"
        print(f"  {alpha:5.2f} {rg_s:>7} {n_e:5d} {tau_a:6.3f} {order}")
    
    return r_best


if __name__ == '__main__':
    result = main()

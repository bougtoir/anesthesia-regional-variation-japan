"""
Generate a causal DAG figure from LiNGAM results on UCI Heart Disease data.
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import networkx as nx
from sklearn.preprocessing import StandardScaler
import lingam
import warnings
warnings.filterwarnings('ignore')

OUTPUT_DIR = '/home/ubuntu/repos/wip/spectral-causality-brainstorm/figures'

def generate_dag():
    # Load data
    url = "https://archive.ics.uci.edu/ml/machine-learning-databases/heart-disease/processed.cleveland.data"
    columns = ['age', 'sex', 'cp', 'trestbps', 'chol', 'fbs', 'restecg',
               'thalach', 'exang', 'oldpeak', 'slope', 'ca', 'thal', 'target']
    df = pd.read_csv(url, names=columns, na_values='?')
    df = df.dropna()

    clinical_vars = ['age', 'trestbps', 'chol', 'thalach', 'oldpeak']
    var_labels = {
        'age': 'Age',
        'trestbps': 'Resting BP',
        'chol': 'Cholesterol',
        'thalach': 'Max HR',
        'oldpeak': 'ST Depression'
    }

    X = df[clinical_vars].astype(float).values
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    # LiNGAM
    model = lingam.DirectLiNGAM()
    model.fit(X_scaled)
    B = model.adjacency_matrix_
    causal_order = model.causal_order_

    labels = [var_labels[clinical_vars[i]] for i in range(len(clinical_vars))]

    edges = []
    for i in range(len(B)):
        for j in range(len(B)):
            if abs(B[i][j]) > 0.05:
                # B[i][j] != 0 means j -> i in LiNGAM convention
                edges.append((j, i, B[i][j]))

    print("Causal order:", [labels[i] for i in causal_order])
    print("Edges:")
    for src, dst, w in edges:
        sign = "+" if w > 0 else "-"
        print(f"  {labels[src]} -> {labels[dst]}: {sign}{abs(w):.3f}")

    # Layout: position by causal order (upstream = top)
    order_rank = {node: rank for rank, node in enumerate(causal_order)}
    n = len(labels)

    # Manual positions for clean layout (x, y) — wider spread
    # causal_order: Age(0), MaxHR(3), STDep(4), RestBP(1), Chol(2)
    pos = {}
    layer_y = {0: 5.5, 1: 3.5, 2: 2.0, 3: 0.0, 4: -1.5}
    layer_x = {0: 3.0, 1: 0.5, 2: 5.5, 3: 1.5, 4: 4.0}
    for node in range(n):
        rank = order_rank[node]
        pos[node] = (layer_x[rank], layer_y[rank])

    # Create figure
    fig, ax = plt.subplots(1, 1, figsize=(12, 10))

    # Node box dimensions
    box_w = 1.6
    box_h = 0.55

    # Draw edges (first pass: arrows only)
    edge_labels = []
    for src, dst, w in edges:
        color = '#1976D2' if w > 0 else '#D32F2F'
        width = max(abs(w) * 6, 1.5)
        sign = "+" if w > 0 else "\u2212"
        label_text = f"{sign}{abs(w):.2f}"

        x1, y1 = pos[src]
        x2, y2 = pos[dst]

        # Shorten arrows so they don't overlap with box
        dx = x2 - x1
        dy = y2 - y1
        dist = np.sqrt(dx**2 + dy**2)
        shrink = 0.75 / dist if dist > 0 else 0
        sx1 = x1 + dx * shrink
        sy1 = y1 + dy * shrink
        sx2 = x2 - dx * shrink
        sy2 = y2 - dy * shrink

        ax.annotate("",
            xy=(sx2, sy2), xytext=(sx1, sy1),
            arrowprops=dict(
                arrowstyle='-|>',
                color=color,
                lw=width,
                connectionstyle='arc3,rad=0.08',
                mutation_scale=18,
            ),
            zorder=5
        )

        # Place label at 35% along edge (closer to source, away from congestion)
        t = 0.35
        mid_x = x1 + t * (x2 - x1)
        mid_y = y1 + t * (y2 - y1)
        perp_x = -dy / dist * 0.35 if dist > 0 else 0
        perp_y = dx / dist * 0.35 if dist > 0 else 0
        lbl_x = mid_x + perp_x
        lbl_y = mid_y + perp_y

        # Iteratively shift label away from overlapping nodes (max 4 times)
        for _ in range(4):
            overlap = False
            for node_id in range(n):
                nx_pos, ny_pos = pos[node_id]
                if abs(lbl_x - nx_pos) < box_w/2 + 0.2 and abs(lbl_y - ny_pos) < box_h/2 + 0.2:
                    overlap = True
                    break
            if not overlap:
                break
            # Try flipping to the other side of the edge
            perp_x = -perp_x
            perp_y = -perp_y
            lbl_x = mid_x + perp_x
            lbl_y = mid_y + perp_y

        edge_labels.append((lbl_x, lbl_y, label_text, color))

    # Draw edge labels (second pass: on top of everything)
    for lbl_x, lbl_y, label_text, color in edge_labels:
        ax.text(lbl_x, lbl_y, label_text,
                fontsize=10, color=color, fontweight='bold',
                ha='center', va='center',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                          edgecolor=color, alpha=0.95, lw=1),
                zorder=10)

    # Draw nodes as rounded rectangles
    for node in range(n):
        rank = order_rank[node]
        x, y = pos[node]
        t = rank / (n - 1)
        # Blue gradient: upstream=dark, downstream=light
        r_c = int(25 + t * 120)
        g_c = int(118 + t * 60)
        b_c = int(210 - t * 40)
        fc = f'#{r_c:02x}{g_c:02x}{b_c:02x}'

        box = FancyBboxPatch(
            (x - box_w/2, y - box_h/2), box_w, box_h,
            boxstyle="round,pad=0.1",
            facecolor=fc, edgecolor='white', linewidth=2,
            zorder=3
        )
        ax.add_patch(box)
        ax.text(x, y, labels[node], ha='center', va='center',
               fontsize=12, fontweight='bold', color='white', zorder=4)

    # Causal order sidebar
    sidebar_x = 7.5
    ax.text(sidebar_x, 6.0, 'Causal Order', fontsize=11, ha='center',
            fontweight='bold', color='#444444')
    ax.text(sidebar_x, 5.6, '(upstream \u2192 downstream)', fontsize=9,
            ha='center', style='italic', color='#888888')
    for node in causal_order:
        rank = order_rank[node]
        y_pos = 4.8 - rank * 0.8
        marker = '\u25B6' if rank == 0 else '\u25B7'
        ax.text(sidebar_x, y_pos, f'{marker}  #{rank+1}  {labels[node]}',
                fontsize=10, ha='center', va='center', color='#555555',
                fontfamily='monospace')

    # Vertical arrow for direction
    ax.annotate('', xy=(sidebar_x - 1.0, 0.8), xytext=(sidebar_x - 1.0, 5.0),
                arrowprops=dict(arrowstyle='->', color='#AAAAAA', lw=1.5))

    # Legend
    pos_patch = mpatches.Patch(color='#1976D2', label='Positive effect (+)')
    neg_patch = mpatches.Patch(color='#D32F2F', label='Negative effect (\u2212)')
    ax.legend(handles=[pos_patch, neg_patch], loc='lower left',
             fontsize=11, framealpha=0.9, edgecolor='#CCCCCC')

    ax.set_xlim(-0.5, 9.0)
    ax.set_ylim(-2.5, 6.8)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title('Estimated Causal DAG (DirectLiNGAM)\nUCI Heart Disease Data (n = 297)',
                fontsize=15, fontweight='bold', pad=18)

    plt.tight_layout()
    output_path = f'{OUTPUT_DIR}/fig6_causal_dag.png'
    fig.savefig(output_path, dpi=200, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    # Also save to A2 directory
    a2_path = '/home/ubuntu/repos/wip/spectral-causality-a2-ecd/figures/fig6_causal_dag.png'
    fig.savefig(a2_path, dpi=200, bbox_inches='tight',
                facecolor='white', edgecolor='none')
    plt.close()
    print(f"\nSaved: {output_path}")
    print(f"Saved: {a2_path}")
    return output_path


if __name__ == '__main__':
    generate_dag()

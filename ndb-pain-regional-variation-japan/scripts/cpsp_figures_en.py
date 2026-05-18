#!/usr/bin/env python3
"""
Generate figures for CPSP integrated analysis (Phase 2) - English version.
All labels, legends, titles, and annotations in English.
"""
import csv
import json
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib import rcParams
from scipy import stats
from collections import defaultdict

# Font setup
rcParams['font.family'] = 'sans-serif'
rcParams['font.sans-serif'] = ['DejaVu Sans', 'IPAexGothic', 'Noto Sans CJK JP']
rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = '/home/ubuntu/analysis/output/'

# Load data
rows = []
with open(OUTPUT_DIR + 'cpsp_integrated_results.csv', 'r', encoding='utf-8') as f:
    for r in csv.DictReader(f):
        for k in r:
            if k not in ('pref_name', 'region', 'is_tohoku', 'pref_code'):
                try:
                    r[k] = float(r[k])
                except:
                    pass
        r['pref_code'] = int(r['pref_code'])
        r['is_tohoku'] = int(float(r['is_tohoku']))
        rows.append(r)

with open(OUTPUT_DIR + 'cpsp_regression_summary.json', 'r') as f:
    reg_summary = json.load(f)

# Translation mappings
REGION_EN = {
    '北海道': 'Hokkaido', '東北': 'Tohoku', '関東': 'Kanto',
    '北陸・甲信越': 'Hokuriku-Koshinetsu', '東海': 'Tokai', '近畿': 'Kinki',
    '中国': 'Chugoku', '四国': 'Shikoku', '九州・沖縄': 'Kyushu-Okinawa',
}

PREF_EN = {
    '北海道': 'Hokkaido',
    '青森県': 'Aomori', '岩手県': 'Iwate', '宮城県': 'Miyagi',
    '秋田県': 'Akita', '山形県': 'Yamagata', '福島県': 'Fukushima',
    '茨城県': 'Ibaraki', '栃木県': 'Tochigi', '群馬県': 'Gunma',
    '埼玉県': 'Saitama', '千葉県': 'Chiba',
    '東京都': 'Tokyo', '神奈川県': 'Kanagawa',
    '新潟県': 'Niigata', '富山県': 'Toyama',
    '石川県': 'Ishikawa', '福井県': 'Fukui',
    '山梨県': 'Yamanashi', '長野県': 'Nagano',
    '岐阜県': 'Gifu', '静岡県': 'Shizuoka',
    '愛知県': 'Aichi', '三重県': 'Mie',
    '滋賀県': 'Shiga', '京都府': 'Kyoto',
    '大阪府': 'Osaka', '兵庫県': 'Hyogo',
    '奈良県': 'Nara', '和歌山県': 'Wakayama',
    '鳥取県': 'Tottori', '島根県': 'Shimane',
    '岡山県': 'Okayama', '広島県': 'Hiroshima',
    '山口県': 'Yamaguchi',
    '徳島県': 'Tokushima', '香川県': 'Kagawa',
    '愛媛県': 'Ehime', '高知県': 'Kochi',
    '福岡県': 'Fukuoka', '佐賀県': 'Saga',
    '長崎県': 'Nagasaki', '熊本県': 'Kumamoto',
    '大分県': 'Oita', '宮崎県': 'Miyazaki',
    '鹿児島県': 'Kagoshima', '沖縄県': 'Okinawa',
}

def pref_en(name): return PREF_EN.get(name, name)
def region_en(name): return REGION_EN.get(name, name)

# Color scheme
REGION_COLORS = {
    '北海道': '#1f77b4', '東北': '#d62728', '関東': '#ff7f0e',
    '北陸・甲信越': '#2ca02c', '東海': '#9467bd', '近畿': '#8c564b',
    '中国': '#e377c2', '四国': '#7f7f7f', '九州・沖縄': '#bcbd22'
}

REGION_ORDER = ['北海道','東北','関東','北陸・甲信越','東海','近畿','中国','四国','九州・沖縄']

# ============================================================
# Figure 1: Unadjusted neuropathic pain drugs per surgery by prefecture
# ============================================================
fig, ax = plt.subplots(figsize=(16, 7))

# Sort by value
sorted_rows = sorted(rows, key=lambda x: x['neuropathic_per_surgery'])
names = [pref_en(r['pref_name']) for r in sorted_rows]
vals = [r['neuropathic_per_surgery'] for r in sorted_rows]
colors = [REGION_COLORS[r['region']] for r in sorted_rows]
tohoku_mask = [r['is_tohoku'] for r in sorted_rows]

bars = ax.bar(range(len(names)), vals, color=colors, edgecolor='white', linewidth=0.5)

# Highlight Tohoku with border
for i, (bar, is_t) in enumerate(zip(bars, tohoku_mask)):
    if is_t:
        bar.set_edgecolor('#d62728')
        bar.set_linewidth(2)

ax.axhline(y=np.mean(vals), color='black', linestyle='--', linewidth=1, alpha=0.7, label='National mean')
ax.set_xticks(range(len(names)))
ax.set_xticklabels(names, rotation=90, fontsize=7)
ax.set_ylabel('Outpatient neuropathic pain drug prescriptions / surgical cases', fontsize=11)
# Title removed per Pain submission guidelines (legends in manuscript text only)

# Legend
handles = [mpatches.Patch(color=REGION_COLORS[r], label=region_en(r)) for r in REGION_ORDER]
handles.append(plt.Line2D([0],[0], color='black', linestyle='--', label='National mean'))
ax.legend(handles=handles, loc='upper left', fontsize=8, ncol=2)
ax.set_xlim(-0.5, len(names)-0.5)

plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig1_neuropathic_unadjusted_en.png', dpi=600, bbox_inches='tight')
plt.close()
print("Saved fig1_neuropathic_unadjusted_en.png")

# ============================================================
# Figure 2: Confounder correlation scatter panels
# ============================================================
fig, axes = plt.subplots(2, 2, figsize=(14, 12))

confounders = [
    ('diabetes_per_surgery', 'Diabetes drugs / surgery', axes[0,0]),
    ('herpes_per_surgery', 'Herpes zoster antivirals / surgery', axes[0,1]),
    ('antidep_per_surgery', 'Antidepressants / surgery', axes[1,0]),
    ('anxiolytic_per_surgery', 'Anxiolytics / surgery', axes[1,1]),
]

for conf_key, conf_label, ax in confounders:
    x = [r[conf_key] for r in rows]
    y = [r['neuropathic_per_surgery'] for r in rows]
    colors_pts = [REGION_COLORS[r['region']] for r in rows]
    
    for i, r in enumerate(rows):
        ax.scatter(x[i], y[i], c=REGION_COLORS[r['region']], s=40, alpha=0.8, zorder=3,
                  edgecolors='#d62728' if r['is_tohoku'] else 'white', linewidths=1.5 if r['is_tohoku'] else 0.5)
    
    # Regression line
    slope, intercept, r_val, p_val, se = stats.linregress(x, y)
    x_line = np.linspace(min(x), max(x), 100)
    ax.plot(x_line, intercept + slope * x_line, 'k--', alpha=0.5, linewidth=1)
    
    ax.set_xlabel(conf_label, fontsize=10)
    ax.set_ylabel('Neuropathic pain drugs / surgery', fontsize=10)
    ax.set_title(f'r = {r_val:.3f}, p = {p_val:.4f}', fontsize=10)
    ax.grid(True, alpha=0.3)

# Title removed per Pain submission guidelines (legends in manuscript text only)
handles = [mpatches.Patch(color=REGION_COLORS[r], label=region_en(r)) for r in REGION_ORDER]
fig.legend(handles=handles, loc='lower center', ncol=5, fontsize=8, bbox_to_anchor=(0.5, -0.02))
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig2_confounder_correlations_en.png', dpi=600, bbox_inches='tight')
plt.close()
print("Saved fig2_confounder_correlations_en.png")

# ============================================================
# Figure 3: Adjusted CPSP index by prefecture
# ============================================================
fig, ax = plt.subplots(figsize=(16, 7))

sorted_adj = sorted(rows, key=lambda x: x['adjusted_cpsp_index'])
names_adj = [pref_en(r['pref_name']) for r in sorted_adj]
vals_adj = [r['adjusted_cpsp_index'] for r in sorted_adj]
colors_adj = [REGION_COLORS[r['region']] for r in sorted_adj]
tohoku_adj = [r['is_tohoku'] for r in sorted_adj]

bars = ax.bar(range(len(names_adj)), vals_adj, color=colors_adj, edgecolor='white', linewidth=0.5)
for i, (bar, is_t) in enumerate(zip(bars, tohoku_adj)):
    if is_t:
        bar.set_edgecolor('#d62728')
        bar.set_linewidth(2)

ax.axhline(y=0, color='black', linestyle='-', linewidth=1, alpha=0.5)
ax.set_xticks(range(len(names_adj)))
ax.set_xticklabels(names_adj, rotation=90, fontsize=7)
ax.set_ylabel('Adjusted CPSP index (residual)', fontsize=11)
# Title removed per Pain submission guidelines (legends in manuscript text only)

handles = [mpatches.Patch(color=REGION_COLORS[r], label=region_en(r)) for r in REGION_ORDER]
ax.legend(handles=handles, loc='upper left', fontsize=8, ncol=2)
ax.set_xlim(-0.5, len(names_adj)-0.5)

plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig3_adjusted_cpsp_index_en.png', dpi=600, bbox_inches='tight')
plt.close()
print("Saved fig3_adjusted_cpsp_index_en.png")

# ============================================================
# Figure 4: Region comparison - unadjusted vs adjusted (paired bar)
# ============================================================
fig, axes = plt.subplots(1, 2, figsize=(16, 7))

# Left: unadjusted by region
region_raw = defaultdict(list)
region_adj = defaultdict(list)
for r in rows:
    region_raw[r['region']].append(r['neuropathic_per_surgery'])
    region_adj[r['region']].append(r['adjusted_cpsp_index'])

# Sort by unadjusted mean
region_order_sorted = sorted(REGION_ORDER, key=lambda x: np.mean(region_raw[x]))

ax = axes[0]
x_pos = range(len(region_order_sorted))
means = [np.mean(region_raw[r]) for r in region_order_sorted]
sds = [np.std(region_raw[r]) for r in region_order_sorted]
colors_reg = [REGION_COLORS[r] for r in region_order_sorted]
bars = ax.bar(x_pos, means, yerr=sds, color=colors_reg, edgecolor='white', capsize=3)
ax.axhline(y=np.mean([r['neuropathic_per_surgery'] for r in rows]), color='black', linestyle='--', alpha=0.5)
ax.set_xticks(x_pos)
ax.set_xticklabels([region_en(r) for r in region_order_sorted], rotation=45, ha='right', fontsize=9)
ax.set_ylabel('Neuropathic pain drugs / surgery', fontsize=10)
ax.set_title('(A) Unadjusted', fontsize=12, fontweight='bold')
ax.grid(axis='y', alpha=0.3)

# Highlight Tohoku
for i, (bar, rname) in enumerate(zip(bars, region_order_sorted)):
    if rname == '東北':
        bar.set_edgecolor('#d62728')
        bar.set_linewidth(2)

# Right: adjusted by region
ax = axes[1]
means_adj = [np.mean(region_adj[r]) for r in region_order_sorted]
sds_adj = [np.std(region_adj[r]) for r in region_order_sorted]
bars = ax.bar(x_pos, means_adj, yerr=sds_adj, color=colors_reg, edgecolor='white', capsize=3)
ax.axhline(y=0, color='black', linestyle='-', alpha=0.5)
ax.set_xticks(x_pos)
ax.set_xticklabels([region_en(r) for r in region_order_sorted], rotation=45, ha='right', fontsize=9)
ax.set_ylabel('Adjusted CPSP index', fontsize=10)
ax.set_title('(B) Confounder-adjusted', fontsize=12, fontweight='bold')
ax.grid(axis='y', alpha=0.3)

for i, (bar, rname) in enumerate(zip(bars, region_order_sorted)):
    if rname == '東北':
        bar.set_edgecolor('#d62728')
        bar.set_linewidth(2)

# Title removed per Pain submission guidelines (legends in manuscript text only)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig4_region_unadj_vs_adj_en.png', dpi=600, bbox_inches='tight')
plt.close()
print("Saved fig4_region_unadj_vs_adj_en.png")

# ============================================================
# Figure 5: Phase 1 (acute) vs Phase 2 (chronic) scatter
# ============================================================
fig, axes = plt.subplots(1, 2, figsize=(14, 6))

# Left: Acute vs Chronic (raw)
ax = axes[0]
x_acute = [r['acute_analgesic_per_surgery'] for r in rows]
y_chronic = [r['neuropathic_per_surgery'] for r in rows]

for i, r in enumerate(rows):
    ax.scatter(x_acute[i], y_chronic[i], c=REGION_COLORS[r['region']], s=50, alpha=0.8,
              edgecolors='#d62728' if r['is_tohoku'] else 'white',
              linewidths=2 if r['is_tohoku'] else 0.5, zorder=3)
    if r['is_tohoku'] or r['pref_name'] in ['東京都', '大阪府', '鹿児島県', '岐阜県']:
        ax.annotate(pref_en(r['pref_name']), (x_acute[i], y_chronic[i]), fontsize=6,
                   xytext=(5, 5), textcoords='offset points')

slope, intercept, r_val, p_val, _ = stats.linregress(x_acute, y_chronic)
x_line = np.linspace(min(x_acute), max(x_acute), 100)
ax.plot(x_line, intercept + slope * x_line, 'k--', alpha=0.5)
ax.set_xlabel('Phase 1: Inpatient analgesics / surgery (acute pain proxy)', fontsize=10)
ax.set_ylabel('Phase 2: Outpatient neuropathic pain drugs / surgery\n(persistent pain proxy: unadjusted)', fontsize=10)
ax.set_title(f'(A) Unadjusted  r={r_val:.3f}, p={p_val:.4f}', fontsize=11, fontweight='bold')
ax.grid(True, alpha=0.3)

# Right: Acute vs Chronic (adjusted)
ax = axes[1]
y_adj = [r['adjusted_cpsp_index'] for r in rows]

for i, r in enumerate(rows):
    ax.scatter(x_acute[i], y_adj[i], c=REGION_COLORS[r['region']], s=50, alpha=0.8,
              edgecolors='#d62728' if r['is_tohoku'] else 'white',
              linewidths=2 if r['is_tohoku'] else 0.5, zorder=3)
    if r['is_tohoku'] or r['pref_name'] in ['東京都', '大阪府', '鹿児島県', '岐阜県']:
        ax.annotate(pref_en(r['pref_name']), (x_acute[i], y_adj[i]), fontsize=6,
                   xytext=(5, 5), textcoords='offset points')

slope2, intercept2, r_val2, p_val2, _ = stats.linregress(x_acute, y_adj)
ax.plot(x_line, intercept2 + slope2 * x_line, 'k--', alpha=0.5)
ax.set_xlabel('Phase 1: Inpatient analgesics / surgery (acute pain proxy)', fontsize=10)
ax.set_ylabel('Phase 2: Adjusted CPSP index', fontsize=10)
ax.set_title(f'(B) Confounder-adjusted  r={r_val2:.3f}, p={p_val2:.4f}', fontsize=11, fontweight='bold')
ax.grid(True, alpha=0.3)
ax.axhline(y=0, color='gray', linestyle='-', alpha=0.3)

# Title removed per Pain submission guidelines (legends in manuscript text only)
handles = [mpatches.Patch(color=REGION_COLORS[r], label=region_en(r)) for r in REGION_ORDER]
fig.legend(handles=handles, loc='lower center', ncol=5, fontsize=8, bbox_to_anchor=(0.5, -0.03))
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig5_phase1_vs_phase2_en.png', dpi=600, bbox_inches='tight')
plt.close()
print("Saved fig5_phase1_vs_phase2_en.png")

# ============================================================
# Figure 6: Regression coefficient forest plot
# ============================================================
fig, ax = plt.subplots(figsize=(10, 8))

# Model 2 coefficients (from the regression output)
# Re-run quick regression to get exact values
from scipy.stats import t as t_dist

# We'll read from regression output - using stored values
# Model 2: Neuropathic ~ Diabetes + Herpes + Antidep + Anxiolytic + is_Tohoku
# Need to re-extract coefficients. Let's just plot the key models.

models = [
    ('Model 1: Unadjusted\nis_Tohoku', reg_summary['model1_unadjusted']['cohens_d'], 
     0, reg_summary['model1_unadjusted']['p_value']),
    ('Model 2: Fully adjusted\nis_Tohoku', None, None, reg_summary['model2_adjusted']['tohoku_p']),
    ('Model 3: Core neuropathic\nis_Tohoku', None, None, reg_summary['model3_core_neuropathic']['tohoku_p']),
    ('Model 5: Acute integrated\nis_Tohoku', None, None, reg_summary['model5_integrated']['tohoku_p']),
    ('Adjusted CPSP index\nis_Tohoku', reg_summary['adjusted_cpsp_test']['cohens_d'],
     0, reg_summary['adjusted_cpsp_test']['p_value']),
]

# Better: Show p-values and effect sizes as a summary table figure
fig, ax = plt.subplots(figsize=(12, 5))
ax.axis('off')

table_data = [
    ['Model', 'Dependent variable', 'Tohoku coeff. / effect size', 'P value', 'Result'],
    ['Model 1', 'Neuropathic pain drugs / surgery (unadjusted)', f"d = {reg_summary['model1_unadjusted']['cohens_d']:.3f}", 
     f"{reg_summary['model1_unadjusted']['p_value']:.4f}", '***'],
    ['Model 2', 'Neuropathic pain drugs / surgery (fully adjusted)', f"β = {reg_summary['model2_adjusted']['tohoku_coef']:.1f}",
     f"{reg_summary['model2_adjusted']['tohoku_p']:.4f}", 'ns'],
    ['Model 3', 'Core neuropathic drugs (PGB+MGB) (fully adjusted)', f"β = {reg_summary['model3_core_neuropathic']['tohoku_coef']:.1f}",
     f"{reg_summary['model3_core_neuropathic']['tohoku_p']:.4f}", 'ns'],
    ['Model 5', 'Neuropathic pain drugs (acute + confounder adj.)', f"β = {reg_summary['model5_integrated']['tohoku_coef']:.1f}",
     f"{reg_summary['model5_integrated']['tohoku_p']:.4f}", 'ns'],
    ['Adj CPSP', 'Confounder-removed residual', f"d = {reg_summary['adjusted_cpsp_test']['cohens_d']:.3f}",
     f"{reg_summary['adjusted_cpsp_test']['p_value']:.4f}", 'ns'],
]

table = ax.table(cellText=table_data[1:], colLabels=table_data[0], 
                loc='center', cellLoc='center')
table.auto_set_font_size(False)
table.set_fontsize(10)
table.scale(1.2, 1.8)

# Color significant cells
for i in range(len(table_data)-1):
    for j in range(len(table_data[0])):
        cell = table[i+1, j]
        if i == 0:  # Model 1 (significant)
            cell.set_facecolor('#ffcccc')
        else:
            cell.set_facecolor('#ccffcc')
    # Header
    for j in range(len(table_data[0])):
        table[0, j].set_facecolor('#4472C4')
        table[0, j].set_text_props(color='white', fontweight='bold')

# Title removed per Pain submission guidelines (legends in manuscript text only)

plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig6_model_comparison_table_en.png', dpi=600, bbox_inches='tight')
plt.close()
print("Saved fig6_model_comparison_table_en.png")

# ============================================================
# Supplementary Figure: Heatmap of per-surgery indices
# ============================================================
fig, ax = plt.subplots(figsize=(14, 12))

# Sort prefectures by neuropathic/surgery
sorted_idx = sorted(range(len(rows)), key=lambda i: rows[i]['neuropathic_per_surgery'])

vars_to_plot = [
    ('acute_analgesic_per_surgery', 'Phase 1: Inpatient analgesics / surgery'),
    ('neuropathic_per_surgery', 'Phase 2: Neuropathic pain drugs / surgery'),
    ('adjusted_cpsp_index', 'Adjusted CPSP index'),
    ('diabetes_per_surgery', 'Diabetes drugs / surgery'),
    ('herpes_per_surgery', 'Herpes zoster drugs / surgery'),
    ('antidep_per_surgery', 'Antidepressants / surgery'),
    ('anxiolytic_per_surgery', 'Anxiolytics / surgery'),
    ('nerve_block_per_surgery', 'Nerve blocks / surgery'),
]

# Normalize each variable to z-scores for heatmap
data_matrix = []
for var_key, var_label in vars_to_plot:
    vals = [rows[i][var_key] for i in sorted_idx]
    mean_v = np.mean(vals)
    sd_v = np.std(vals)
    if sd_v > 0:
        z_vals = [(v - mean_v) / sd_v for v in vals]
    else:
        z_vals = [0] * len(vals)
    data_matrix.append(z_vals)

data_array = np.array(data_matrix)
pref_labels = [pref_en(rows[i]['pref_name']) for i in sorted_idx]
var_labels = [v[1] for v in vars_to_plot]

im = ax.imshow(data_array, aspect='auto', cmap='RdYlBu_r', vmin=-2.5, vmax=2.5)
ax.set_xticks(range(len(pref_labels)))
ax.set_xticklabels(pref_labels, rotation=90, fontsize=7)
ax.set_yticks(range(len(var_labels)))
ax.set_yticklabels(var_labels, fontsize=9)

# Highlight Tohoku columns
for i in sorted_idx:
    if rows[i]['is_tohoku']:
        pos = [j for j, si in enumerate(sorted_idx) if si == i][0]
        ax.axvline(x=pos-0.5, color='red', linewidth=0.5, alpha=0.5)
        ax.axvline(x=pos+0.5, color='red', linewidth=0.5, alpha=0.5)

plt.colorbar(im, ax=ax, label='Z-score', shrink=0.8)
# Title removed per Pain submission guidelines (legends in manuscript text only)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'sfig1_heatmap_en.png', dpi=600, bbox_inches='tight')
plt.close()
print("Saved sfig1_heatmap_en.png")

print("\nAll English figures saved successfully.")

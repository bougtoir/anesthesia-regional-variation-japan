#!/usr/bin/env python3
"""
Phase 1 Analysis: Regional Differences in Analgesic Use per Surgery in Japan
Using NDB Open Data (10th Edition, 2023/04-2024/03)

Research Question: Do Tohoku region residents show lower analgesic use per surgery?
"""

import pandas as pd
import numpy as np
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import warnings
warnings.filterwarnings('ignore')

try:
    import japanize_matplotlib
    HAS_JP_FONT = True
except:
    HAS_JP_FONT = False

# ============================================================
# CONFIGURATION
# ============================================================
DATA_DIR = '/home/ubuntu/analysis/data/'
OUTPUT_DIR = '/home/ubuntu/analysis/output/'

import os
os.makedirs(OUTPUT_DIR, exist_ok=True)

PREFECTURE_NAMES = {
    1: '北海道', 2: '青森県', 3: '岩手県', 4: '宮城県', 5: '秋田県',
    6: '山形県', 7: '福島県', 8: '茨城県', 9: '栃木県', 10: '群馬県',
    11: '埼玉県', 12: '千葉県', 13: '東京都', 14: '神奈川県', 15: '新潟県',
    16: '富山県', 17: '石川県', 18: '福井県', 19: '山梨県', 20: '長野県',
    21: '岐阜県', 22: '静岡県', 23: '愛知県', 24: '三重県', 25: '滋賀県',
    26: '京都府', 27: '大阪府', 28: '兵庫県', 29: '奈良県', 30: '和歌山県',
    31: '鳥取県', 32: '島根県', 33: '岡山県', 34: '広島県', 35: '山口県',
    36: '徳島県', 37: '香川県', 38: '愛媛県', 39: '高知県', 40: '福岡県',
    41: '佐賀県', 42: '長崎県', 43: '熊本県', 44: '大分県', 45: '宮崎県',
    46: '鹿児島県', 47: '沖縄県'
}

REGIONS = {
    '北海道': [1],
    '東北': [2, 3, 4, 5, 6, 7],
    '関東': [8, 9, 10, 11, 12, 13, 14],
    '北陸・甲信越': [15, 16, 17, 18, 19, 20],
    '東海': [21, 22, 23, 24],
    '近畿': [25, 26, 27, 28, 29, 30],
    '中国': [31, 32, 33, 34, 35],
    '四国': [36, 37, 38, 39],
    '九州・沖縄': [40, 41, 42, 43, 44, 45, 46, 47]
}

ANALGESIC_CLASSES = {
    114: '解熱鎮痛消炎剤(NSAIDs等)',
    811: 'あへんアルカロイド系麻薬(オピオイド)',
    821: '合成麻薬(フェンタニル等)'
}

POPULATION_2023 = {
    1: 5140, 2: 1204, 3: 1181, 4: 2280, 5: 929, 6: 1040, 7: 1790,
    8: 2840, 9: 1908, 10: 1912, 11: 7337, 12: 6266, 13: 14048, 14: 9232,
    15: 2153, 16: 1017, 17: 1118, 18: 752, 19: 802, 20: 2019,
    21: 1943, 22: 3582, 23: 7495, 24: 1745, 25: 1412, 26: 2546,
    27: 8781, 28: 5402, 29: 1306, 30: 903, 31: 543, 32: 657,
    33: 1856, 34: 2737, 35: 1313, 36: 704, 37: 936, 38: 1304,
    39: 675, 40: 5104, 41: 800, 42: 1283, 43: 1718, 44: 1106,
    45: 1052, 46: 1562, 47: 1468
}

# ============================================================
# STEP 1: Load and parse surgery data
# ============================================================
print("=" * 60)
print("STEP 1: Loading surgery data")
print("=" * 60)

surgery_df = pd.read_excel(DATA_DIR + 'surgery_prefecture.xlsx', sheet_name='入院', header=None)
print(f"Raw shape: {surgery_df.shape}")

surgery_data = surgery_df.iloc[4:].copy()
surgery_data.columns = list(range(surgery_data.shape[1]))

header_row = surgery_df.iloc[3]
pref_cols = {}
for i, val in enumerate(header_row):
    if isinstance(val, str):
        for code, name in PREFECTURE_NAMES.items():
            if name == val:
                pref_cols[code] = i
                break

print(f"Found {len(pref_cols)} prefecture columns")

surgery_by_pref = {}
for pref_code, col_idx in pref_cols.items():
    col_data = surgery_data[col_idx].copy()
    col_data = col_data.replace('-', 0).replace('\u2010', 0).replace('\u2012', 0).replace('\u2013', 0).replace('\u2014', 0).replace('\u2015', 0).replace('\u2212', 0).replace('\uff0d', 0).replace('\u2500', 0)
    col_data = pd.to_numeric(col_data, errors='coerce').fillna(0)
    surgery_by_pref[pref_code] = col_data.sum()

print("\n=== Total Surgery Counts by Prefecture ===")
for code in sorted(surgery_by_pref.keys()):
    print(f"{PREFECTURE_NAMES[code]}: {surgery_by_pref[code]:,.0f}")

total_surgeries = sum(surgery_by_pref.values())
print(f"\nTotal surgeries (all prefectures): {total_surgeries:,.0f}")

# ============================================================
# STEP 2: Load and parse inpatient prescription drug data
# ============================================================
print("\n" + "=" * 60)
print("STEP 2: Loading inpatient prescription drug data")
print("=" * 60)

drug_df = pd.read_excel(DATA_DIR + 'inpatient_drugs_prefecture.xlsx',
                         sheet_name='内服薬 入院', header=None)
print(f"Raw shape: {drug_df.shape}")

drug_header = drug_df.iloc[3]
drug_pref_cols = {}
for i, val in enumerate(drug_header):
    if isinstance(val, str):
        for code, name in PREFECTURE_NAMES.items():
            if name == val:
                drug_pref_cols[code] = i
                break

print(f"Found {len(drug_pref_cols)} prefecture columns in drug data")

drug_data = drug_df.iloc[4:].copy()
drug_data.columns = list(range(drug_data.shape[1]))

analgesic_by_pref = {code: 0.0 for code in range(1, 48)}
analgesic_detail = {cls: {code: 0.0 for code in range(1, 48)} for cls in ANALGESIC_CLASSES}

# Track current drug class for rows where it's NaN (inherited)
current_class = None
for idx, row in drug_data.iterrows():
    drug_class = row[0]
    if pd.notna(drug_class):
        try:
            current_class = int(float(drug_class))
        except (ValueError, TypeError):
            current_class = None
            continue

    if current_class is not None and current_class in ANALGESIC_CLASSES:
        for pref_code, col_idx in drug_pref_cols.items():
            val = row[col_idx]
            if isinstance(val, str):
                val = 0
            try:
                val = float(val)
                if np.isnan(val):
                    val = 0
            except (ValueError, TypeError):
                val = 0
            analgesic_by_pref[pref_code] += val
            analgesic_detail[current_class][pref_code] += val

print("\n=== Analgesic Prescription Quantities by Prefecture (Inpatient) ===")
for code in sorted(analgesic_by_pref.keys()):
    print(f"{PREFECTURE_NAMES[code]}: {analgesic_by_pref[code]:,.1f}")

print("\n=== Detail by Drug Class ===")
for cls, name in ANALGESIC_CLASSES.items():
    total = sum(analgesic_detail[cls].values())
    print(f"{cls} {name}: Total = {total:,.1f}")

# ============================================================
# STEP 3: Calculate "analgesic per surgery" metric
# ============================================================
print("\n" + "=" * 60)
print("STEP 3: Calculating 'analgesic quantity per surgery' metric")
print("=" * 60)

results = []
for code in range(1, 48):
    surgeries = surgery_by_pref.get(code, 0)
    analgesics = analgesic_by_pref.get(code, 0)
    population = POPULATION_2023.get(code, 0)

    region = None
    for reg_name, prefectures in REGIONS.items():
        if code in prefectures:
            region = reg_name
            break

    is_tohoku = 1 if region == '東北' else 0

    if surgeries > 0:
        analgesic_per_surgery = analgesics / surgeries
        analgesic_per_capita = analgesics / (population * 1000) if population > 0 else 0
        surgery_per_capita = surgeries / (population * 1000) if population > 0 else 0
    else:
        analgesic_per_surgery = 0
        analgesic_per_capita = 0
        surgery_per_capita = 0

    results.append({
        'pref_code': code,
        'pref_name': PREFECTURE_NAMES[code],
        'region': region,
        'is_tohoku': is_tohoku,
        'surgery_count': surgeries,
        'analgesic_quantity': analgesics,
        'population_thousands': population,
        'analgesic_per_surgery': analgesic_per_surgery,
        'analgesic_per_capita': analgesic_per_capita,
        'surgery_per_capita': surgery_per_capita
    })

results_df = pd.DataFrame(results)
results_df.to_csv(OUTPUT_DIR + 'prefecture_results.csv', index=False, encoding='utf-8-sig')

print("\n=== Analgesic Quantity per Surgery (Top/Bottom) ===")
sorted_df = results_df.sort_values('analgesic_per_surgery')
print("\nBottom 10 (lowest analgesic use per surgery):")
for _, row in sorted_df.head(10).iterrows():
    marker = " *** TOHOKU" if row['is_tohoku'] else ""
    print(f"  {row['pref_name']}: {row['analgesic_per_surgery']:.2f}{marker}")

print("\nTop 10 (highest analgesic use per surgery):")
for _, row in sorted_df.tail(10).iterrows():
    marker = " *** TOHOKU" if row['is_tohoku'] else ""
    print(f"  {row['pref_name']}: {row['analgesic_per_surgery']:.2f}{marker}")

# ============================================================
# STEP 4: Statistical comparison - Tohoku vs Rest
# ============================================================
print("\n" + "=" * 60)
print("STEP 4: Statistical comparison - Tohoku vs Rest of Japan")
print("=" * 60)

tohoku = results_df[results_df['is_tohoku'] == 1]['analgesic_per_surgery']
rest = results_df[results_df['is_tohoku'] == 0]['analgesic_per_surgery']

print(f"\nTohoku (n={len(tohoku)}):")
print(f"  Mean: {tohoku.mean():.2f}")
print(f"  SD: {tohoku.std():.2f}")
print(f"  Median: {tohoku.median():.2f}")
print(f"  Range: {tohoku.min():.2f} - {tohoku.max():.2f}")

print(f"\nRest of Japan (n={len(rest)}):")
print(f"  Mean: {rest.mean():.2f}")
print(f"  SD: {rest.std():.2f}")
print(f"  Median: {rest.median():.2f}")
print(f"  Range: {rest.min():.2f} - {rest.max():.2f}")

t_stat, t_pval = stats.ttest_ind(tohoku, rest)
print(f"\nIndependent t-test: t={t_stat:.4f}, p={t_pval:.4f}")

u_stat, u_pval = stats.mannwhitneyu(tohoku, rest, alternative='two-sided')
print(f"Mann-Whitney U test: U={u_stat:.1f}, p={u_pval:.4f}")

_, u_pval_less = stats.mannwhitneyu(tohoku, rest, alternative='less')
print(f"Mann-Whitney U (one-sided, Tohoku < Rest): p={u_pval_less:.4f}")

pooled_std = np.sqrt(((len(tohoku)-1)*tohoku.std()**2 + (len(rest)-1)*rest.std()**2) /
                      (len(tohoku)+len(rest)-2))
cohens_d = (tohoku.mean() - rest.mean()) / pooled_std if pooled_std > 0 else 0
print(f"Cohen's d: {cohens_d:.4f}")

diff_mean = tohoku.mean() - rest.mean()
se_diff = np.sqrt(tohoku.std()**2/len(tohoku) + rest.std()**2/len(rest))
ci_95 = (diff_mean - 1.96*se_diff, diff_mean + 1.96*se_diff)
print(f"Difference (Tohoku - Rest): {diff_mean:.2f} (95% CI: {ci_95[0]:.2f} to {ci_95[1]:.2f})")

# ============================================================
# STEP 4b: Regional block comparison
# ============================================================
print("\n=== Regional Block Comparison ===")
regional_stats = results_df.groupby('region')['analgesic_per_surgery'].agg(['mean', 'std', 'median', 'count'])
regional_stats = regional_stats.sort_values('mean')
print(regional_stats.to_string())

region_groups = [group['analgesic_per_surgery'].values for _, group in results_df.groupby('region')]
h_stat, h_pval = stats.kruskal(*region_groups)
print(f"\nKruskal-Wallis test (all regions): H={h_stat:.4f}, p={h_pval:.4f}")

# ============================================================
# STEP 5: Drug class breakdown by region
# ============================================================
print("\n" + "=" * 60)
print("STEP 5: Drug class breakdown")
print("=" * 60)

for cls, name in ANALGESIC_CLASSES.items():
    print(f"\n--- {cls}: {name} ---")
    tohoku_vals = []
    rest_vals = []
    for code in range(1, 48):
        surgeries = surgery_by_pref.get(code, 0)
        if surgeries > 0:
            ratio = analgesic_detail[cls][code] / surgeries
        else:
            ratio = 0
        if code in REGIONS['東北']:
            tohoku_vals.append(ratio)
        else:
            rest_vals.append(ratio)

    tohoku_arr = np.array(tohoku_vals)
    rest_arr = np.array(rest_vals)
    print(f"  Tohoku mean: {tohoku_arr.mean():.4f} (SD: {tohoku_arr.std():.4f})")
    print(f"  Rest mean: {rest_arr.mean():.4f} (SD: {rest_arr.std():.4f})")
    if len(tohoku_arr) > 0 and len(rest_arr) > 0:
        u, p = stats.mannwhitneyu(tohoku_arr, rest_arr, alternative='two-sided')
        print(f"  Mann-Whitney U: U={u:.1f}, p={p:.4f}")

# ============================================================
# STEP 6: Visualizations
# ============================================================
print("\n" + "=" * 60)
print("STEP 6: Creating visualizations")
print("=" * 60)

# --- Figure 1: Prefecture bar chart ---
fig, ax = plt.subplots(figsize=(16, 10))
sorted_results = results_df.sort_values('analgesic_per_surgery')
colors = ['#e74c3c' if x == 1 else '#3498db' for x in sorted_results['is_tohoku']]
bars = ax.barh(range(len(sorted_results)), sorted_results['analgesic_per_surgery'], color=colors)
ax.set_yticks(range(len(sorted_results)))
ax.set_yticklabels(sorted_results['pref_name'], fontsize=7)
ax.set_xlabel('Analgesic Quantity per Surgery (Inpatient NSAIDs + Opioids)', fontsize=11)
ax.set_title('Prefecture-level Analgesic Use per Surgery\n(NDB 10th Open Data, 2023/04-2024/03)\nRed = Tohoku Region', fontsize=13)
ax.axvline(x=results_df['analgesic_per_surgery'].mean(), color='gray', linestyle='--', alpha=0.5, label='National mean')
ax.legend()
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig1_prefecture_bar.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig1_prefecture_bar.png")

# --- Figure 2: Regional block comparison ---
fig, ax = plt.subplots(figsize=(12, 6))
region_order = regional_stats.sort_values('mean').index.tolist()
region_means = [regional_stats.loc[r, 'mean'] for r in region_order]
region_stds = [regional_stats.loc[r, 'std'] for r in region_order]
colors2 = ['#e74c3c' if r == '東北' else '#3498db' for r in region_order]
bars2 = ax.bar(range(len(region_order)), region_means, yerr=region_stds,
               color=colors2, capsize=5, alpha=0.8)
ax.set_xticks(range(len(region_order)))
ax.set_xticklabels(region_order, fontsize=10)
ax.set_ylabel('Mean Analgesic Quantity per Surgery', fontsize=11)
ax.set_title('Regional Block Comparison: Analgesic Use per Surgery\n(Error bars = SD)\nRed = Tohoku', fontsize=13)
ax.axhline(y=results_df['analgesic_per_surgery'].mean(), color='gray', linestyle='--', alpha=0.5, label='National mean')
ax.legend()
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig2_regional_comparison.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig2_regional_comparison.png")

# --- Figure 3: Scatter plot ---
fig, ax = plt.subplots(figsize=(10, 8))
tohoku_mask = results_df['is_tohoku'] == 1
ax.scatter(results_df[~tohoku_mask]['surgery_count'],
           results_df[~tohoku_mask]['analgesic_per_surgery'],
           c='#3498db', alpha=0.7, s=80, label='Other regions', edgecolors='white')
ax.scatter(results_df[tohoku_mask]['surgery_count'],
           results_df[tohoku_mask]['analgesic_per_surgery'],
           c='#e74c3c', alpha=0.9, s=120, label='Tohoku', edgecolors='white', marker='D')
for _, row in results_df[tohoku_mask].iterrows():
    ax.annotate(row['pref_name'], (row['surgery_count'], row['analgesic_per_surgery']),
                textcoords="offset points", xytext=(5, 5), fontsize=8, color='#e74c3c')
ax.set_xlabel('Total Surgery Count (Inpatient)', fontsize=11)
ax.set_ylabel('Analgesic Quantity per Surgery', fontsize=11)
ax.set_title('Surgery Volume vs Analgesic Use per Surgery\nTohoku highlighted in red', fontsize=13)
ax.legend()
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig3_scatter.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig3_scatter.png")

# --- Figure 4: Box plot by region ---
fig, ax = plt.subplots(figsize=(12, 6))
region_data = []
region_labels = []
for reg in region_order:
    prefs = REGIONS[reg]
    vals = results_df[results_df['pref_code'].isin(prefs)]['analgesic_per_surgery'].values
    region_data.append(vals)
    region_labels.append(reg)
bp = ax.boxplot(region_data, labels=region_labels, patch_artist=True)
for i, patch in enumerate(bp['boxes']):
    if region_labels[i] == '東北':
        patch.set_facecolor('#e74c3c')
        patch.set_alpha(0.7)
    else:
        patch.set_facecolor('#3498db')
        patch.set_alpha(0.5)
ax.set_ylabel('Analgesic Quantity per Surgery', fontsize=11)
ax.set_title('Distribution of Analgesic Use per Surgery by Region\n(Box plot, Red = Tohoku)', fontsize=13)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig4_boxplot.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig4_boxplot.png")

# --- Figure 5: Drug class breakdown ---
fig, axes = plt.subplots(1, 3, figsize=(18, 8))
for idx, (cls, name) in enumerate(ANALGESIC_CLASSES.items()):
    ax = axes[idx]
    vals = []
    for code in range(1, 48):
        surgeries = surgery_by_pref.get(code, 0)
        if surgeries > 0:
            ratio = analgesic_detail[cls][code] / surgeries
        else:
            ratio = 0
        vals.append(ratio)
    results_df[f'class_{cls}_per_surgery'] = vals
    sorted_temp = results_df.sort_values(f'class_{cls}_per_surgery')
    colors_temp = ['#e74c3c' if x == 1 else '#3498db' for x in sorted_temp['is_tohoku']]
    ax.barh(range(len(sorted_temp)), sorted_temp[f'class_{cls}_per_surgery'], color=colors_temp)
    ax.set_yticks(range(len(sorted_temp)))
    ax.set_yticklabels(sorted_temp['pref_name'], fontsize=5)
    ax.set_title(f'{cls}: {name}', fontsize=9)
    ax.axvline(x=results_df[f'class_{cls}_per_surgery'].mean(), color='gray', linestyle='--', alpha=0.5)
plt.suptitle('Drug Class Breakdown: Per-Surgery Use by Prefecture (Red = Tohoku)', fontsize=13)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig5_drug_class_breakdown.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig5_drug_class_breakdown.png")

# ============================================================
# STEP 7: Summary
# ============================================================
print("\n" + "=" * 60)
print("STEP 7: Summary for report")
print("=" * 60)

tohoku_mean = tohoku.mean()
all_region_means = results_df.groupby('region')['analgesic_per_surgery'].mean().sort_values()
tohoku_rank = list(all_region_means.index).index('東北') + 1
print(f"\nTohoku rank among {len(all_region_means)} regions: {tohoku_rank}/{len(all_region_means)} (1=lowest)")
print(f"Tohoku mean: {tohoku_mean:.2f}")
print(f"National mean: {results_df['analgesic_per_surgery'].mean():.2f}")
print(f"Difference from national: {((tohoku_mean / results_df['analgesic_per_surgery'].mean()) - 1)*100:.1f}%")

print("\nTohoku prefecture rankings:")
all_sorted = results_df.sort_values('analgesic_per_surgery').reset_index(drop=True)
for _, row in results_df[results_df['is_tohoku']==1].iterrows():
    rank = (all_sorted['pref_code'] == row['pref_code']).idxmax() + 1
    print(f"  {row['pref_name']}: {row['analgesic_per_surgery']:.2f} (rank {rank}/47)")

# Regional ranking table
print("\n=== All Regions Ranked ===")
for i, (reg, mean_val) in enumerate(all_region_means.items()):
    marker = " <-- TOHOKU" if reg == '東北' else ""
    print(f"  {i+1}. {reg}: {mean_val:.2f}{marker}")

print("\n=== ANALYSIS COMPLETE ===")
print(f"Results saved to {OUTPUT_DIR}")

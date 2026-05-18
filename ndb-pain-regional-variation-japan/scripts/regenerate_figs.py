#!/usr/bin/env python3
"""Regenerate figures with proper Japanese font support."""

import pandas as pd
import numpy as np
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
import warnings
warnings.filterwarnings('ignore')

# Set Japanese font
font_path = '/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf'
font_manager.fontManager.addfont(font_path)
jp_font = font_manager.FontProperties(fname=font_path)
plt.rcParams['font.family'] = jp_font.get_name()
plt.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = '/home/ubuntu/analysis/output/'

# Load results
results_df = pd.read_csv(OUTPUT_DIR + 'prefecture_results.csv')

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

regional_stats = results_df.groupby('region')['analgesic_per_surgery'].agg(['mean', 'std', 'median', 'count'])
regional_stats = regional_stats.sort_values('mean')

# --- Figure 1: Prefecture bar chart ---
fig, ax = plt.subplots(figsize=(16, 12))
sorted_results = results_df.sort_values('analgesic_per_surgery')
colors = ['#e74c3c' if x == 1 else '#3498db' for x in sorted_results['is_tohoku']]
bars = ax.barh(range(len(sorted_results)), sorted_results['analgesic_per_surgery'], color=colors)
ax.set_yticks(range(len(sorted_results)))
ax.set_yticklabels(sorted_results['pref_name'], fontsize=9, fontproperties=jp_font)
ax.set_xlabel('手術あたり鎮痛薬処方量（入院NSAIDs＋オピオイド）', fontsize=12, fontproperties=jp_font)
ax.set_title('都道府県別 手術あたり鎮痛薬使用量\n（NDB第10回オープンデータ 2023年4月〜2024年3月）\n赤＝東北地方', fontsize=14, fontproperties=jp_font)
ax.axvline(x=results_df['analgesic_per_surgery'].mean(), color='gray', linestyle='--', alpha=0.5, label=f'全国平均 ({results_df["analgesic_per_surgery"].mean():.1f})')
ax.legend(prop=jp_font, fontsize=11)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig1_prefecture_bar.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig1_prefecture_bar.png")

# --- Figure 2: Regional block comparison ---
fig, ax = plt.subplots(figsize=(14, 7))
region_order = regional_stats.sort_values('mean').index.tolist()
region_means = [regional_stats.loc[r, 'mean'] for r in region_order]
region_stds = [regional_stats.loc[r, 'std'] for r in region_order]
colors2 = ['#e74c3c' if r == '東北' else '#3498db' for r in region_order]
bars2 = ax.bar(range(len(region_order)), region_means, yerr=region_stds,
               color=colors2, capsize=5, alpha=0.8, edgecolor='white', linewidth=0.5)
# Add value labels on bars
for i, (m, s) in enumerate(zip(region_means, region_stds)):
    ax.text(i, m + (s if not np.isnan(s) else 0) + 0.8, f'{m:.1f}', ha='center', va='bottom', fontsize=10, fontweight='bold')
ax.set_xticks(range(len(region_order)))
ax.set_xticklabels(region_order, fontsize=11, fontproperties=jp_font)
ax.set_ylabel('手術あたり鎮痛薬処方量（平均値）', fontsize=12, fontproperties=jp_font)
ax.set_title('地域ブロック別 手術あたり鎮痛薬使用量の比較\n（エラーバー＝標準偏差、赤＝東北）', fontsize=14, fontproperties=jp_font)
ax.axhline(y=results_df['analgesic_per_surgery'].mean(), color='gray', linestyle='--', alpha=0.5, label=f'全国平均 ({results_df["analgesic_per_surgery"].mean():.1f})')
ax.legend(prop=jp_font, fontsize=11)
ax.set_ylim(0, 55)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig2_regional_comparison.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig2_regional_comparison.png")

# --- Figure 3: Scatter plot ---
fig, ax = plt.subplots(figsize=(10, 8))
tohoku_mask = results_df['is_tohoku'] == 1
ax.scatter(results_df[~tohoku_mask]['surgery_count'],
           results_df[~tohoku_mask]['analgesic_per_surgery'],
           c='#3498db', alpha=0.7, s=80, label='その他の地域', edgecolors='white')
ax.scatter(results_df[tohoku_mask]['surgery_count'],
           results_df[tohoku_mask]['analgesic_per_surgery'],
           c='#e74c3c', alpha=0.9, s=120, label='東北地方', edgecolors='white', marker='D')
for _, row in results_df[tohoku_mask].iterrows():
    ax.annotate(row['pref_name'], (row['surgery_count'], row['analgesic_per_surgery']),
                textcoords="offset points", xytext=(5, 5), fontsize=9, color='#e74c3c',
                fontproperties=jp_font)
ax.set_xlabel('手術件数（入院）', fontsize=12, fontproperties=jp_font)
ax.set_ylabel('手術あたり鎮痛薬処方量', fontsize=12, fontproperties=jp_font)
ax.set_title('手術件数 vs 手術あたり鎮痛薬使用量\n（東北地方を赤で表示）', fontsize=14, fontproperties=jp_font)
ax.legend(prop=jp_font, fontsize=11)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig3_scatter.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig3_scatter.png")

# --- Figure 4: Box plot by region ---
fig, ax = plt.subplots(figsize=(14, 7))
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
ax.set_xticklabels(region_labels, fontsize=11, fontproperties=jp_font)
ax.set_ylabel('手術あたり鎮痛薬処方量', fontsize=12, fontproperties=jp_font)
ax.set_title('地域ブロック別 手術あたり鎮痛薬使用量の分布\n（箱ひげ図、赤＝東北）', fontsize=14, fontproperties=jp_font)
ax.axhline(y=results_df['analgesic_per_surgery'].mean(), color='gray', linestyle='--', alpha=0.3, label=f'全国平均')
ax.legend(prop=jp_font)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig4_boxplot.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig4_boxplot.png")

# --- Figure 5: Drug class breakdown ---
# Need to reload drug data for class breakdown
DATA_DIR = '/home/ubuntu/analysis/data/'

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

# Re-derive drug class columns from the CSV if they exist
if 'class_114_per_surgery' in results_df.columns:
    fig, axes = plt.subplots(1, 3, figsize=(20, 10))
    class_cols = {114: 'class_114_per_surgery', 811: 'class_811_per_surgery', 821: 'class_821_per_surgery'}
    class_names = {114: 'NSAIDs/アセトアミノフェン\n(薬効分類114)', 811: 'オピオイドアルカロイド\n(薬効分類811)', 821: '合成麻薬\n(薬効分類821)'}
    
    for idx, (cls, col) in enumerate(class_cols.items()):
        ax = axes[idx]
        sorted_temp = results_df.sort_values(col)
        colors_temp = ['#e74c3c' if x == 1 else '#3498db' for x in sorted_temp['is_tohoku']]
        ax.barh(range(len(sorted_temp)), sorted_temp[col], color=colors_temp)
        ax.set_yticks(range(len(sorted_temp)))
        ax.set_yticklabels(sorted_temp['pref_name'], fontsize=6, fontproperties=jp_font)
        ax.set_title(class_names[cls], fontsize=11, fontproperties=jp_font)
        ax.axvline(x=results_df[col].mean(), color='gray', linestyle='--', alpha=0.5)
    plt.suptitle('薬効分類別 手術あたり処方量（赤＝東北地方）', fontsize=14, fontproperties=jp_font, y=1.01)
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR + 'fig5_drug_class_breakdown.png', dpi=150, bbox_inches='tight')
    plt.close()
    print("Saved fig5_drug_class_breakdown.png")
else:
    print("Drug class columns not found in CSV, skipping fig5")

print("\nAll figures regenerated with Japanese fonts!")

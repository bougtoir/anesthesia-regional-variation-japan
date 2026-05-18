#!/usr/bin/env python3
"""Regenerate fig5 with drug class data from original Excel files."""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
import warnings
warnings.filterwarnings('ignore')

font_path = '/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf'
font_manager.fontManager.addfont(font_path)
jp_font = font_manager.FontProperties(fname=font_path)
plt.rcParams['font.family'] = jp_font.get_name()
plt.rcParams['axes.unicode_minus'] = False

DATA_DIR = '/home/ubuntu/analysis/data/'
OUTPUT_DIR = '/home/ubuntu/analysis/output/'

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

ANALGESIC_CLASSES = {114: 'NSAIDs等', 811: 'オピオイド', 821: '合成麻薬'}
TOHOKU_CODES = [2, 3, 4, 5, 6, 7]

# Load surgery data
surgery_df = pd.read_excel(DATA_DIR + 'surgery_prefecture.xlsx', sheet_name='入院', header=None)
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

surgery_by_pref = {}
for pref_code, col_idx in pref_cols.items():
    col_data = surgery_data[col_idx].copy()
    col_data = col_data.replace('-', 0).replace('\u2010', 0).replace('\u2013', 0).replace('\u2014', 0).replace('\u2015', 0).replace('\u2212', 0).replace('\uff0d', 0)
    col_data = pd.to_numeric(col_data, errors='coerce').fillna(0)
    surgery_by_pref[pref_code] = col_data.sum()

# Load drug data
drug_df = pd.read_excel(DATA_DIR + 'inpatient_drugs_prefecture.xlsx', sheet_name='内服薬 入院', header=None)
drug_header = drug_df.iloc[3]
drug_pref_cols = {}
for i, val in enumerate(drug_header):
    if isinstance(val, str):
        for code, name in PREFECTURE_NAMES.items():
            if name == val:
                drug_pref_cols[code] = i
                break

drug_data = drug_df.iloc[4:].copy()
drug_data.columns = list(range(drug_data.shape[1]))

analgesic_detail = {cls: {code: 0.0 for code in range(1, 48)} for cls in ANALGESIC_CLASSES}
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
            analgesic_detail[current_class][pref_code] += val

# Build per-surgery ratios for each class
class_data = {}
for cls in ANALGESIC_CLASSES:
    ratios = []
    names = []
    is_tohoku = []
    for code in range(1, 48):
        surgeries = surgery_by_pref.get(code, 0)
        ratio = analgesic_detail[cls][code] / surgeries if surgeries > 0 else 0
        ratios.append(ratio)
        names.append(PREFECTURE_NAMES[code])
        is_tohoku.append(1 if code in TOHOKU_CODES else 0)
    class_data[cls] = pd.DataFrame({'pref_name': names, 'ratio': ratios, 'is_tohoku': is_tohoku})

# Plot
fig, axes = plt.subplots(1, 3, figsize=(22, 12))
class_titles = {
    114: 'NSAIDs/アセトアミノフェン\n（薬効分類114）',
    811: 'あへんアルカロイド系麻薬\n（薬効分類811）',
    821: '合成麻薬（フェンタニル等）\n（薬効分類821）'
}

for idx, (cls, title) in enumerate(class_titles.items()):
    ax = axes[idx]
    df = class_data[cls].sort_values('ratio')
    colors = ['#e74c3c' if x == 1 else '#3498db' for x in df['is_tohoku']]
    ax.barh(range(len(df)), df['ratio'], color=colors)
    ax.set_yticks(range(len(df)))
    ax.set_yticklabels(df['pref_name'], fontsize=7, fontproperties=jp_font)
    ax.set_title(title, fontsize=12, fontproperties=jp_font)
    ax.axvline(x=df['ratio'].mean(), color='gray', linestyle='--', alpha=0.5)

plt.suptitle('薬効分類別 手術あたり処方量（赤＝東北地方）', fontsize=15, fontproperties=jp_font, y=1.01)
plt.tight_layout()
plt.savefig(OUTPUT_DIR + 'fig5_drug_class_breakdown.png', dpi=150, bbox_inches='tight')
plt.close()
print("Saved fig5_drug_class_breakdown.png")

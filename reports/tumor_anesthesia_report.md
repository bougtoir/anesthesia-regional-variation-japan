# 悪性腫瘍手術における麻酔方法の地域差分析

## 背景と目的

悪性腫瘍手術において、全身麻酔単独よりも硬膜外麻酔・区域麻酔を併用した方が無再発生存期間(RFS)および5年生存率で有利であるとするエビデンスが蓄積されている。しかし実臨床では全身麻酔単独で施行されるケースも少なくない。

本分析では、NDBオープンデータの二次医療圏別SCR（標準化レセプト出現比）を用いて：

1. **全身麻酔(L008)と硬膜外麻酔(L002)の併用パターン**の地域分布を可視化
2. **L003（硬膜外麻酔後持続注入）**を全身麻酔+硬膜外併用の直接指標として分析
3. 地域差が**査定（審査）で説明がつくものか、構造的な差異か**を感度分析で検証

---

## データと方法

### 使用指標
| コード | 名称 | 意味 | 圏数 |
|--------|------|------|------|
| L008 | 閉鎖循環式全身麻酔 | 全身麻酔の実施量 | 334 |
| L002 | 硬膜外麻酔 | 硬膜外麻酔（単独or併用）| 307 |
| L003 | 硬膜外麻酔後持続注入 | **GA+硬膜外併用の直接指標** | 331 |
| L004 | 脊椎麻酔 | くも膜下麻酔 | 334 |

### 解析指標
- **L002/L008比率**: 全身麻酔に対する硬膜外麻酔の相対量 → 区域麻酔併用傾向
- **L003/L008比率**: 全身麻酔に対する持続硬膜外の比率 → 併用率の直接指標
- **L008+L002合計SCR**: 査定による振替効果を中和した合計量

### GISデータ
- 国土数値情報 A38-20（二次医療圏）335圏
- 北方領土は領土として白色表示（医療圏未設定）

---

## 結果

### 1. 全身麻酔+硬膜外併用の地域差

#### L003（持続硬膜外注入）= GA+硬膜外併用の直接指標

L003は全身麻酔後に硬膜外カテーテルから局所麻酔薬を持続注入する手技であり、全身麻酔+硬膜外併用を最も直接的に反映する。

| 指標 | 全体 | 大学病院圏(64) | 非大学圏(271) |
|------|------|---------------|---------------|
| L003 SCR mean | 83.5 | **126.4** | 73.2 |
| L003 SCR std | 54.2 | - | - |
| CV | **64.9%** | - | - |

- **大学病院効果**: d = 0.96（大効果）、p = 1.5e-08
- 大学病院圏は非大学圏の**1.73倍**のGA+硬膜外併用率

#### L008とL003の相関
- **r(L008, L003) = 0.753**, p = 8.4e-62
- 全身麻酔が多い地域ほど硬膜外併用も多い → 麻酔科の充実度を反映

### 2. 硬膜外/全身麻酔比率（L002/L008）

| 指標 | 全体 | 大学病院圏 | 非大学圏 |
|------|------|-----------|---------|
| L002/L008比率 mean | 1.25 | **0.98** | 1.33 |
| L002/L008比率 median | 0.89 | - | - |
| CV | **99.5%** | - | - |

- **大学病院圏の方が比率が低い** (d = -0.33, p = 0.002)
- 解釈: 大学病院圏はL008(GA)もL002(硬膜外)も絶対量が多いが、GAの増加がより顕著なため比率は低下
- 県内比較: 47県中18県(38%)で大学圏の比率が高い → 県によりパターンが異なる

### 3. 査定による振替仮説の検証

#### 3.1 相関構造テスト

「査定によりL008がL002に振り替えられる」仮説が正しければ、L008とL002は**負の相関**を示すはず。

| 相関 | r値 | p値 | 解釈 |
|------|-----|-----|------|
| r(L008, L002) | **+0.319** | 1.1e-08 | **正の相関** |
| r(L008, L003) | **+0.753** | 8.4e-62 | **強い正の相関** |

- L008とL002は正の相関 → **査定による振替が主因ではない**
- 全身麻酔が多い地域は硬膜外も多い = 麻酔科リソースの供給効果

#### 3.2 変動係数テスト

査定振替が主因なら、L008+L002合計のCVはL008単独より大幅に低下するはず。

| 指標 | CV |
|------|-----|
| L008単独 | 50.2% |
| L002単独 | 87.1% |
| **L008+L002合計** | **58.3%** |
| CV変化 | **+16.1%増加** |

- 合計してもCVは**むしろ増加** → 査定振替効果はほぼゼロ
- L008とL002が正に相関しているため、合計するとばらつきが拡大

#### 3.3 分散分解（L008+L002合計）

| 成分 | 寄与率 |
|------|--------|
| 都道府県間差 | 16.2% |
| 県内・大学病院効果 | **25.0%** |
| 残差 | 58.8% |

- 大学病院の有無（単一二値変数）が合計SCRの**25%**を説明
- 査定方針は最大でも県間差16.2%の一部しか説明できない

#### 3.4 合計SCRでの大学病院効果

| 指標 | 大学病院圏 | 非大学圏 | 差 |
|------|-----------|---------|-----|
| L008+L002合計 mean | **257.6** | 158.4 | +99.1 |
| Cohen's d | | | **1.00** |
| t検定 p値 | | | **6.7e-10** |

- 合計しても大学病院効果は巨大（d = 1.00）
- 査定を完全に除外しても、地域差は明確に残存

### 4. 結論

#### 査定仮説の検証結果

| テスト | 査定仮説の予測 | 実測 | 判定 |
|--------|--------------|------|------|
| L008-L002相関 | 負 (r < 0) | **r = +0.32** | 棄却 |
| 合計CV低下 | 大幅低下 | **むしろ増加** | 棄却 |
| 合計での大学効果消失 | 効果消失 | **d = 1.00** | 棄却 |

**結論: 悪性腫瘍手術における全身麻酔+硬膜外併用の地域差は、査定による見かけの差ではなく、構造的な臨床実践の地域差である。**

#### 臨床的含意

1. **大学病院圏**は全身麻酔・硬膜外麻酔・持続硬膜外注入のいずれも高い → 麻酔科リソースの充実が併用率を規定
2. **GA+硬膜外併用率(L003)**のCV=64.9%は極めて大きく、「同じエビデンスがあっても実践率は地域により2倍以上異なる」ことを示す
3. RFSや5年生存率への影響を考えると、この地域差は**医療の質の格差**として政策的な注目に値する
4. 大学病院・医局の教育方針が長期予後に影響しうるメカニズムとして、麻酔方法選択のバリエーションは重要な研究課題

---

## 図版一覧

### 既存図版（北方領土追加版）
1. **Figure 1**: 全身麻酔 (L008) SCR 二次医療圏別 (`map_L008_scr.png`)
2. **Figure 2**: 大学病院所在 二次医療圏 (`map_univ_presence.png`)
3. **Figure 3**: 大学病院数 二次医療圏別 (`map_n_univ.png`)
4. **Figure 4**: 脊椎麻酔 (L004) SCR (`map_L004_scr.png`)
5. **Figure 5**: 全身麻酔+脊椎麻酔 合計SCR (`map_L008_L004_combined.png`)
6. **Figure 6**: 硬膜外麻酔 (L002) SCR (`map_L002_scr.png`)

### 新規図版（悪性腫瘍手術分析）
7. **Figure 7**: 硬膜外麻酔後持続注入 (L003) SCR — GA+硬膜外併用の直接指標 (`map_L003_scr.png`)
8. **Figure 8**: 硬膜外/全身麻酔 比率 (L002/L008) — 区域麻酔併用傾向 (`map_L002_L008_ratio.png`)
9. **Figure 9**: 持続硬膜外/全身麻酔 比率 (L003/L008) — 硬膜外併用率 (`map_L003_L008_ratio.png`)
10. **Figure 10**: 全身麻酔+硬膜外麻酔 合計SCR (L008+L002) — 査定中和 (`map_L008_L002_combined.png`)

### 地図凡例
- **実線**: 都道府県境
- **点線**: 二次医療圏境
- **白色（北東）**: 北方領土（医療圏未設定、日本固有の領土）
- **赤丸**: 大学病院所在圏

---

## 技術的注記

- SCRデータ: NDBオープンデータ R04（令和4年度）二次医療圏別
- GISデータ: 国土数値情報 A38-20（二次医療圏）、N03-20240101（行政界、北方領土用）
- L003は入院のみ（in/out=1）のデータで、259圏で利用可能
- L002/L008比率・L003/L008比率は、分母(L008)が0またはデータなしの圏を除外して算出

---

# English Translation

---

# Analysis of regional differences in anesthesia methods in malignant tumor surgery

## Background and purpose

In surgery for malignant tumors, evidence is accumulating that the combination of epidural anesthesia and regional anesthesia is more advantageous in terms of recurrence-free survival (RFS) and 5-year survival rate than general anesthesia alone. However, in actual clinical practice, there are many cases in which general anesthesia alone is used.

In this analysis, we used the SCR (Standardized Receipt Occurrence Ratio) by secondary medical care area from NDB open data:

1. Visualize the regional distribution of **combination pattern of general anesthesia (L008) and epidural anesthesia (L002)**
2. Analysis of **L003 (continuous infusion after epidural anesthesia)** as a direct indicator of general anesthesia + epidural combination
3. Verify through sensitivity analysis whether regional differences can be explained by assessment (examination) or are structural differences.

---

## Data and methods

### Usage metrics
| Code | Name | Meaning | Category number |
|--------|------|------|------|
| L008 | Closed circulation general anesthesia | Amount of general anesthesia administered | 334 |
| L002 | Epidural anesthesia | Epidural anesthesia (alone or in combination) | 307 |
| L003 | Continuous infusion after epidural anesthesia | **Direct indicator of GA + epidural combination** | 331 |
| L004 | Spinal anesthesia | Subarachnoid anesthesia | 334 |

### Analysis metrics
- **L002/L008 ratio**: Relative amount of epidural anesthesia to general anesthesia → tendency to use regional anesthesia together
- **L003/L008 ratio**: Ratio of continuous epidural to general anesthesia → direct indicator of combination rate
- **L008+L002 total SCR**: Total amount after neutralizing the transfer effect due to assessment

### GIS data
- National land numerical information A38-20 (secondary medical area) 335 areas
- The Northern Territories are displayed in white as territory (medical area not set)

---

## Results

### 1. Regional differences in general anesthesia + epidural combination

#### L003 (continuous epidural infusion) = Direct indicator of GA + epidural combination

L003 is a procedure in which local anesthetic is continuously injected through an epidural catheter after general anesthesia, and most directly reflects the combination of general anesthesia and epidural.

| Indicator | Overall | University hospital area (64) | Non-university area (271) |
|------|------|---------------|--------------|
| L003 SCR mean | 83.5 | **126.4** | 73.2 |
| L003 SCR std | 54.2 | - | - |
| CV | **64.9%** | - | - |

- **University hospital effect**: d = 0.96 (large effect), p = 1.5e-08
- GA + epidural combination rate in university hospital areas is **1.73 times higher than in non-university areas
#### Correlation between L008 and L003
- **r(L008, L003) = 0.753**, p = 8.4e-62
- In regions where general anesthesia is more common, epidural combinations are also more common → reflecting the level of anesthesiology

### 2. Epidural/general anesthesia ratio (L002/L008)

| Indicator | Overall | University hospital area | Non-university area |
|------|------|------------|---------|
| L002/L008 ratio mean | 1.25 | **0.98** | 1.33 |
| L002/L008 ratio median | 0.89 | - | - |
| CV | **99.5%** | - | - |

- **The ratio is lower in university hospital areas** (d = -0.33, p = 0.002)
- Interpretation: The absolute amount of both L008 (GA) and L002 (epidural) is large in the university hospital area, but the ratio is decreasing because the increase in GA is more pronounced.
- Comparison within prefectures: 18 out of 47 prefectures (38%) have a high ratio of university areas → Patterns vary by prefecture

### 3. Verification of transfer hypothesis through assessment

#### 3.1 Correlation structure test

If the hypothesis that "L008 is transferred to L002 due to assessment" is correct, L008 and L002 should show a **negative correlation**.

| Correlation | r value | p value | Interpretation |
|------|----|------|------|
| r(L008, L002) | **+0.319** | 1.1e-08 | **Positive correlation** |
| r(L008, L003) | **+0.753** | 8.4e-62 | **Strong positive correlation** |

- L008 and L002 are positively correlated → **Transfer due to appraisal is not the main cause**
- Areas with more general anesthesia also have more epidurals = anesthesiology resource supply effect

#### 3.2 Coefficient of variation test

If appraisal transfer is the main cause, the CV of L008+L002 combined should be significantly lower than L008 alone.

| Index | CV |
|------|----|
| L008 alone | 50.2% |
| L002 alone | 87.1% |
| **L008+L002 total** | **58.3%** |
| CV change | **+16.1% increase** |

- Even in total, CV **increases** → Assessment transfer effect is almost zero
- Since L008 and L002 are positively correlated, the dispersion increases when added together

#### 3.3 Variance decomposition (L008+L002 total)

| Component | Contribution rate |
|------|---------|
| Difference between prefectures | 16.2% |
| Prefecture/university hospital effect | **25.0%** |
| Residual | 58.8% |

- Presence or absence of university hospital (single dichotomous variable) explains **25%** of total SCR
- At most, the assessment policy can only partially explain the 16.2% difference between prefectures.

#### 3.4 University hospital effect on total SCR

| Indicator | University hospital area | Non-university area | Difference |
|------|----------|---------|------|
| L008+L002 total mean | **257.6** | 158.4 | +99.1 |
| Cohen's d | | | **1.00** |
| t-test p-value | | | **6.7e-10** |

- Even in total, the university hospital effect is huge (d = 1.00)
- Even if assessment is completely excluded, regional differences clearly remain

### 4. Conclusion

#### Verification results of assessment hypothesis

| Test | Prediction of assessment hypothesis | Actual measurement | Judgment |
|--------|--------------|------|------|
| L008-L002 correlation | Negative (r < 0) | **r = +0.32** | Reject |
| Total CV decrease | Significant decrease | **Increase** | Rejection |
| Total university effect disappears | Effect disappears | **d = 1.00** | Rejected |

**Conclusion: Regional differences in the combination of general anesthesia and epidural surgery in malignant tumor surgery are not due to apparent differences in assessment, but are due to structural regional differences in clinical practice. **

#### Clinical implications
1. General anesthesia, epidural anesthesia, and continuous epidural infusion are all high in **university hospital areas** → The availability of anesthesiology resources determines the combination use rate
2. The CV=64.9% for **GA + epidural combination rate (L003)** is extremely large, indicating that “even with the same evidence, the practice rate varies by region by more than twice.”
3. Considering the impact on RFS and 5-year survival rate, this regional difference deserves policy attention as a **disparity in medical quality**
4. Variation in anesthesia method selection is an important research topic as a mechanism by which the educational policies of university hospitals and medical departments may affect long-term prognosis.

---

## List of illustrations

### Existing illustrations (Northern Territories additional edition)
1. **Figure 1**: General anesthesia (L008) SCR by secondary care area (`map_L008_scr.png`)
2. **Figure 2**: Secondary medical care area where university hospital is located (`map_univ_presence.png`)
3. **Figure 3**: Number of university hospitals by secondary healthcare area (`map_n_univ.png`)
4. **Figure 4**: Spinal anesthesia (L004) SCR (`map_L004_scr.png`)
5. **Figure 5**: General anesthesia + spinal anesthesia total SCR (`map_L008_L004_combined.png`)
6. **Figure 6**: Epidural anesthesia (L002) SCR (`map_L002_scr.png`)

### New illustration (malignant tumor surgery analysis)
7. **Figure 7**: Continuous infusion after epidural anesthesia (L003) SCR — Direct indicator of GA+epidural combination (`map_L003_scr.png`)
8. **Figure 8**: Epidural/general anesthesia ratio (L002/L008) — Trend of combined use of regional anesthesia (`map_L002_L008_ratio.png`)
9. **Figure 9**: Continuous epidural/general anesthesia ratio (L003/L008) — Epidural combination ratio (`map_L003_L008_ratio.png`)
10. **Figure 10**: General anesthesia + epidural anesthesia total SCR (L008+L002) — Assessment neutralization (`map_L008_L002_combined.png`)

### Map legend
- **Solid line**: Prefectural border
- **Dotted line**: Secondary medical care area
- **White (Northeast)**: Northern Territories (medical area not established, territory unique to Japan)
- **Red circle**: University hospital area

---

## Technical notes

- SCR data: NDB open data R04 (FY2020) by secondary medical area
- GIS data: National land numerical information A38-20 (secondary medical area), N03-20240101 (administrative boundary, for the Northern Territories)
- L003 is data for hospitalization only (in/out=1) and is available in 259 areas
- L002/L008 ratio and L003/L008 ratio are calculated by excluding areas where the denominator (L008) is 0 or no data.

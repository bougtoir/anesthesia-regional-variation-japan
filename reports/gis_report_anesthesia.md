# 麻酔の地域差 GIS可視化レポート

## 概要

二次医療圏(335圏)レベルのSCRデータと大学病院所在情報をGISデータに統合し、麻酔関連指標の地域差を地図上に可視化した。都道府県境は**実線**、二次医療圏境は**点線**で区別している。

加えて、大腿骨頸部骨折における「全身麻酔(L008) vs 脊椎麻酔(L004)」の査定による振替効果を検証するため、L008+L004合計SCRの感度分析を実施した。

---

## Figure 1: 全身麻酔 (L008) SCR

![全身麻酔 SCR](https://app.devin.ai/attachments/136a03dd-cda3-4978-b5d4-2344f34696a6/map_L008_scr.png)

- 赤い点 = 大学病院所在圏
- 大学病院所在圏に高SCR(暖色)が集中している
- 東京区中央部(5校): SCR=435.7、北海道上川(旭川医大): SCR=226.5
- 大学病院のない圏域の多くはSCR<70(寒色)

---

## Figure 2: 大学病院所在 二次医療圏

![大学病院所在](https://app.devin.ai/attachments/025fb89e-a34c-40bb-a749-9f24fa1a99cc/map_univ_presence.png)

- 赤 = 大学病院あり (64圏 / 335圏 = 19%)
- Figure 1と重ね合わせると、大学病院所在圏とL008高SCR圏域の一致が視覚的に明確

---

## Figure 3: 大学病院数 (用量反応)

![大学病院数](https://app.devin.ai/attachments/17314cd4-3e73-4adc-a621-cf3ae51e5830/map_n_univ.png)

- 0校(271圏) → 1校(51圏) → 2校(10圏) → 3校(2圏) → 5校(1圏: 東京区中央部)
- 大学病院数が増えるほどL008 SCRが高い傾向 (用量反応関係)

---

## Figure 4: 脊椎麻酔 (L004) SCR

![脊椎麻酔 SCR](https://app.devin.ai/attachments/a00b3700-622f-4e1c-aa13-6e7e171f8554/map_L004_scr.png)

- L008とは異なるパターン: 大学病院との関連は弱い (Cohen's d = 0.36)
- 東北・北陸に高SCR圏域が散在
- **L008と正の相関 (r = +0.235)**: もし査定による振替(L008→L004)が主因なら負の相関になるはずだが、実際は正 → 査定振替は主因ではない

---

## Figure 5: 全身麻酔+脊椎麻酔 合計SCR (L008+L004) -- 査定振替効果の中和

![L008+L004合計](https://app.devin.ai/attachments/7537701c-0530-4898-83fc-a83b5fdec1e4/map_L008_L004_combined.png)

**これが本レポートの核心図版。**

大腿骨頸部骨折で「脊椎麻酔+鎮静+O2 = 全身麻酔」として算定するか「脊椎麻酔のみ」とするかは、査定によって地域差が生じる可能性がある。L008+L004の合計を取ることで、この振替効果を中和した「真の地域差」を可視化している。

### 結果: 地域差は消えない

| 指標 | L008単独 | L004単独 | L008+L004合計 |
|------|----------|----------|--------------|
| CV (変動係数) | 53.3% | 56.4% | **43.4%** |
| 大学圏 mean | 130.2 | 108.6 | **238.7** |
| 非大学圏 mean | 68.0 | 90.6 | **158.6** |
| 差 | +62.2 | +17.9 | **+80.1** |
| Cohen's d | 1.53 | 0.36 | **1.16** |
| t値 | 9.69 | 2.68 | **8.18** |
| p値 | 7.4e-15 | 8.4e-3 | **1.5e-12** |

- CV reduction: 53.3% → 43.4% (削減率わずか19%)
- **合計しても大学病院効果は依然として極めて大きい (d=1.16)**
- 県内比較: 47県中**45県 (96%)** で大学圏の合計SCR > 非大学圏

### 分散分解 (L008+L004合計)

| 成分 | SS | 寄与率 |
|------|-----|--------|
| 県間差 | 358,911 | 18.9% |
| 県内・大学病院効果 | 475,738 | **25.1%** |
| 残差 | 1,060,474 | 56.0% |

大学病院の有無という単一二値変数が、合計SCRの分散の25.1%を説明する。

---

## Figure 6: 硬膜外麻酔 (L002) SCR

![硬膜外麻酔 SCR](https://app.devin.ai/attachments/0d0811fc-cff3-4e32-bd3a-c7a91114b6ea/map_L002_scr.png)

- 大学病院効果は中程度 (Cohen's d = 0.49)
- 西日本(特に九州)に高SCR圏域が集中する独自の地理的パターン

---

## 感度分析の結論

### 査定による振替仮説の検証

大腿骨頸部骨折の例のように、「同じ医療行為なのに地域によってL008で算定できたりL004でしか算定できなかったりする」場合:

1. **L008とL004は負の相関になるはず** → 実際は **r = +0.235 (正の相関)**
2. **L008+L004合計のCVはゼロに近づくはず** → 実際は **CV = 43.4% (依然として大きい)**
3. **合計でも大学病院効果は消えるはず** → 実際は **d = 1.16 (依然として極めて大きい)**

### 解釈

| シナリオ | 予測 | 観測 | 判定 |
|----------|------|------|------|
| 査定振替が地域差の全て | r(L008,L004)≈-1, CV(合計)≈0 | r=+0.24, CV=43% | **棄却** |
| 査定振替が地域差の大半 | CV大幅減少 | CV減少わずか19% | **棄却** |
| 査定振替は一部のみ、本質は臨床ポリシー差 | CV一部減少、大学効果残存 | CV19%減少、d=1.16 | **支持** |

**結論: 査定による算定振替は地域差のごく一部(最大19%)しか説明しない。残りの81%以上は、大学病院の有無と関連する本質的な臨床ポリシーの地域差である。**

ただし、この「臨床ポリシーの地域差」には:
- 大学医局の学術的方針の影響
- 都市部の病院集積効果(大学病院と交絡)
- 査定の繰り返しによる防衛的縮小請求の地域差

が含まれており、これらを分離するにはさらなる分析が必要。

---

## データソース

- SCR: 内閣府「見える化」R4年度 二次医療圏別 (https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/)
- GIS: 国土数値情報 医療圏データ R2年度 (https://nlftp.mlit.go.jp/ksj/gml/datalist/KsjTmplt-A38.html)
- 大学病院: 81校 → 64二次医療圏にマッピング (手動照合・検証済み)

---

# English Translation

---

# Regional differences in anesthesia GIS visualization report

## Overview

SCR data at the level of secondary medical care areas (335 areas) and university hospital location information were integrated with GIS data, and regional differences in anesthesia-related indicators were visualized on a map. Prefectural boundaries are distinguished by **solid lines**, and secondary medical area boundaries are distinguished by **dotted lines**.

In addition, in order to verify the transfer effect of assessing ``general anesthesia (L008) vs. spinal anesthesia (L004)'' in femoral neck fractures, we conducted a sensitivity analysis of the L008+L004 total SCR.

---

## Figure 1: General anesthesia (L008) SCR

![General anesthesia SCR](https://app.devin.ai/attachments/136a03dd-cda3-4978-b5d4-2344f34696a6/map_L008_scr.png)

- Red dot = university hospital location area
- High SCR (warm colors) are concentrated in areas where university hospitals are located
- Central Tokyo (5 schools): SCR=435.7, Hokkaido Kamikawa (Asahikawa Medical University): SCR=226.5
- Many areas without university hospitals have SCR<70 (cool colors)

---

## Figure 2: Secondary medical care area where university hospitals are located

![University Hospital Location](https://app.devin.ai/attachments/025fb89e-a34c-40bb-a749-9f24fa1a99cc/map_univ_presence.png)
- Red = University hospital available (64 areas / 335 areas = 19%)
- When superimposed on Figure 1, the correspondence between the university hospital area and the L008 high SCR area is visually clear.

---

## Figure 3: Number of university hospitals (dose response)

![Number of university hospitals](https://app.devin.ai/attachments/17314cd4-3e73-4adc-a621-cf3ae51e5830/map_n_univ.png)

- 0 schools (271 areas) → 1 school (51 areas) → 2 schools (10 areas) → 3 schools (2 areas) → 5 schools (1 area: Central Tokyo)
- L008 SCR tends to increase as the number of university hospitals increases (dose-response relationship)

---

## Figure 4: Spinal anesthesia (L004) SCR

![Spinal anesthesia SCR](https://app.devin.ai/attachments/a00b3700-622f-4e1c-aa13-6e7e171f8554/map_L004_scr.png)

- Different pattern from L008: weak association with university hospitals (Cohen's d = 0.36)
- High SCR areas are scattered in Tohoku and Hokuriku
- **Positive correlation with L008 (r = +0.235)**: If transfer due to appraisal (L008→L004) is the main cause, there should be a negative correlation, but it is actually positive → Transfer due to appraisal is not the main cause

---
## Figure 5: General anesthesia + spinal anesthesia total SCR (L008+L004) -- Neutralization of assessment transfer effect

![L008+L004 total](https://app.devin.ai/attachments/7537701c-0530-4898-83fc-a83b5fdec1e4/map_L008_L004_combined.png)

**This is the core illustration of this report. **

There may be regional differences depending on the assessment of whether to calculate ``spinal anesthesia + sedation + O2 = general anesthesia'' or ``spinal anesthesia only'' for femoral neck fractures. By taking the sum of L008+L004, we visualize the "true regional differences" that neutralize this transfer effect.

### Results: Regional differences persist

| Index | L008 alone | L004 alone | L008+L004 total |
|------|----------|----------|--------------|
| CV (coefficient of variation) | 53.3% | 56.4% | **43.4%** |
| University area mean | 130.2 | 108.6 | **238.7** |
| Non-university mean | 68.0 | 90.6 | **158.6** |
| Difference | +62.2 | +17.9 | **+80.1** |
| Cohen's d | 1.53 | 0.36 | **1.16** |
| t value | 9.69 | 2.68 | **8.18** |
| p-value | 7.4e-15 | 8.4e-3 | **1.5e-12** |

- CV reduction: 53.3% → 43.4% (reduction rate only 19%)
- **Even in total, the university hospital effect is still extremely large (d=1.16)**
- Comparison within prefectures: Total SCR in university areas > non-university areas in **45 out of 47 prefectures (96%)**

### Variance decomposition (L008+L004 total)

| Component | SS | Contribution rate |
|------|-----|---------|
| Difference between prefectures | 358,911 | 18.9% |
| Prefecture/university hospital effect | 475,738 | **25.1%** |
| Residual | 1,060,474 | 56.0% |

A single dichotomous variable, presence or absence of a university hospital, explains 25.1% of the variance in total SCR.

---

## Figure 6: Epidural anesthesia (L002) SCR

![Epidural anesthesia SCR](https://app.devin.ai/attachments/0d0811fc-cff3-4e32-bd3a-c7a91114b6ea/map_L002_scr.png)
- University hospital effect is moderate (Cohen's d = 0.49)
- A unique geographical pattern where high SCR areas are concentrated in western Japan (especially Kyushu)

---

## Conclusion of sensitivity analysis

### Verification of transfer hypothesis through appraisal

As in the case of femoral neck fracture, cases where ``even though it is the same medical procedure, it can be calculated as L008 or only L004 depending on the region'':

1. **L008 and L004 should be negatively correlated** → Actually **r = +0.235 (positive correlation)**
2. **L008+L004 total CV should be close to zero** → Actually **CV = 43.4% (still large)**
3. **In total, the university hospital effect should disappear** → Actually **d = 1.16 (still extremely large)**

### Interpretation

| Scenario | Prediction | Observation | Judgment |
|----------|------|------|------|
| Assessment transfer is all the regional difference | r(L008,L004)≈-1, CV(total)≈0 | r=+0.24, CV=43% | **Reject** |
| Assessment transfer accounts for most of the regional differences | CV significantly decreased | CV decreased by only 19% | **Rejected** |
| Appraisal transfer is only partially, the essence is clinical policy difference | CV partially decreased, university effect remains | CV 19% decrease, d=1.16 | **Support** |
**Conclusion: Assessment transfers explain only a small portion (up to 19%) of regional differences. The remaining 81% or more is due to regional differences in essential clinical policies related to the presence or absence of university hospitals. **

However, this “regional variation in clinical policy”:
- Influence of academic policies of university medical offices
- Hospital agglomeration effect in urban areas (confounded with university hospitals)
- Regional differences in defensive reduction claims due to repeated assessments

further analysis is required to separate them.

---

## Data source

- SCR: Cabinet Office “Visualization” R4 by secondary medical area (https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/)
- GIS: National Land Numerical Information Medical Area Data R2 (https://nlftp.mlit.go.jp/ksj/gml/datalist/KsjTmplt-A38.html)
- University hospitals: 81 schools → mapped to 64 secondary medical areas (manual verification and verification completed)

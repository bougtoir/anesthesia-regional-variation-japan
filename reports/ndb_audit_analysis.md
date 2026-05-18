# NDB/DPCデータと査定情報：公開データで検証可能か？

## 1. ご指摘の本質

NDBオープンデータ・DPCデータは**一次審査（査定）後**のレセプトから集計されています。藤森教授（東北大学）のSCR留意点でも明記されています：

> 「国保連合会及び支払基金からNDBに提供されるレセプトは一次審査後、すなわち審査機関における査定後のものである。その後に保険者査定が行われ、あるいは再審査請求で復活の場合もあるが、それらは反映していない。各医療機関においてどのような医療が提供されたのかというよりは、**一次審査レベルで支払いが認められた医療は何かの指標**、とも言える。」

つまり、NDBオープンデータの地域差 ＝「医師が意図した医療行為の地域差」＋「審査による査定の地域差」が混在しています。

---

## 2. 公開されている査定関連データ

### 2-1. 支払基金「審査統計」（都道府県別）
- **URL**: https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html
- **内容**: 原審査・再審査の審査状況を、**都道府県別**に件数・点数で集計
- **期間**: 平成23年度〜令和7年度（月次データあり）
- **粒度**: 都道府県別の査定件数・査定点数（医科・歯科・調剤別）。ただし**診療行為コード別の査定内訳は公開されていない**

### 2-2. 支払基金「審査の差異の可視化レポート」
- **URL**: https://www.ssk.or.jp/shinryohoshu/saikaisyou_torikumi/kashikarepo/
- **内容**: 「審査情報提供事例」「審査の一般的な取扱い」「多くの付箋がつくコンピュータチェック」について、**都道府県別に審査の差異を可視化**
- **特徴**: 特定の診療行為（例：Dダイマーの手術前スクリーニング検査等）について、都道府県ごとの取扱いの差異を具体的に示す
- **限界**: 全診療行為を網羅するものではなく、差異が顕著な事例のみ

### 2-3. 国保連合会「審査状況」
- **URL**: https://www.kokuho.or.jp/statistics/shinsa/joukyou/
- **内容**: 国保連合会における審査状況（令和2〜7年度）
- **粒度**: 全国状況の集計

### 2-4. 内閣府 SCRデータ（間接的）
- **URL**: https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/index.html
- **内容**: NDBから算出した性・年齢調整済みSCR（都道府県・二次医療圏・市区町村別）
- **留意**: SCR自体が査定後データから算出されている

---

## 3. 既知の知見：査定率の地域差

m3.comの分析（2023年）によると、支払基金の基金年報から医科の支部別審査状況を5年おきに分析した結果：

- 査定率の高い都道府県は**「西高東低」**の傾向
- **大阪府・福岡県**が上位3位以内に頻出、過去10年では**北海道**も目立つ
- 2009年データでは、請求点数1万点当たりの原審査査定点数で、最高の大阪（0.284%）と最低の宮崎（0.071%）で約4倍の開き

---

## 4. 公開データで何ができるか／できないか

### できること（実現可能なアプローチ）

| アプローチ | データソース | 検証内容 |
|---|---|---|
| **査定率の地域差の定量化** | 支払基金・審査統計 | 都道府県別の査定件数率・査定点数率の時系列分析 |
| **SCRと査定率の相関分析** | SCRデータ + 審査統計 | 査定率が高い地域でSCRが低いか（査定によりNDB上の算定件数が減少しているか）を間接的に検証 |
| **特定診療行為の差異分析** | 可視化レポート + NDBオープンデータ | 可視化レポートで差異が指摘された診療行為について、NDB上の地域差と照合 |
| **支払基金 vs 国保連の差異** | 両機関の審査統計 | 保険者種別による審査基準の違いの影響 |

### できないこと（データの限界）

- **診療行為コード別の査定内訳**: 支払基金の審査統計は総量（件数・点数）のみで、どの診療行為が査定されたかの詳細は非公開
- **査定前のレセプト情報**: 医師が実際に請求した（査定前の）データは公開されていない → NDBオープンデータとの直接差分は取れない
- **保険者査定・再審査の影響**: 一次審査後の保険者査定や再審査による復活はNDBに反映されておらず、その統計も体系的には非公開
- **個別レセプト単位の分析**: オープンデータは集計値のため、個票レベルでの査定影響は検証不可

---

## 5. 実践的な検証戦略（提案）

### Step 1: 審査統計から査定率の地域プロファイルを作成
支払基金の審査統計ZIPファイルをダウンロードし、都道府県別の査定件数率・査定点数率を算出。時系列で推移を把握。

### Step 2: SCRデータとの突合
内閣府公開のSCRデータ（都道府県別）と、Step 1の査定率を突合。査定率が高い地域でSCRが系統的に低い傾向があるかを回帰分析等で検証。

### Step 3: 可視化レポートの事例ベース分析
支払基金の「審査の差異の可視化レポート」で具体的に差異が指摘されている診療行為（例：Dダイマー）について、NDBオープンデータの当該行為の都道府県別算定件数と照合。

### Step 4: 感度分析
査定率の地域差（例：0.07%〜0.28%）がSCRにどの程度のインパクトを与えうるかをシミュレーション。査定率が最も低い地域の値を基準として、他地域のSCRを補正した場合の地域差の変動幅を推計。

---

## 6. 結論

**部分的には検証可能**です。

- 支払基金の審査統計（都道府県別・件数/点数）と可視化レポートは公開されており、査定の地域差は定量的に確認できます
- ただし、査定率の絶対水準は低く（概ね0.1〜0.3%程度）、NDBオープンデータで見られる診療行為の地域差（SCRで50〜150程度のばらつき）に比べると、**査定のみで地域差を説明するにはインパクトが小さい可能性**があります
- 一方で、特定の診療行為（審査の差異の可視化レポートで指摘されるような項目）については、査定の地域差が個別の算定件数に有意な影響を与えている可能性は否定できません
- **診療行為コード別の査定内訳**が非公開である点が最大のボトルネックです。NDB第三者提供（研究利用）でも査定前データへのアクセスは困難と思われます

SCRの留意点が指摘するように、「何が提供されたか」ではなく「何の支払いが認められたか」のデータである限り、この構造的限界は公開データだけでは完全には解消できません。

---

# English Translation

---

# NDB/DPC data and assessment information: Can it be verified with public data?

## 1. The essence of your point

NDB open data and DPC data are compiled from receipts **after the first screening (assessment)**. Professor Fujimori (Tohoku University)'s SCR points are also clearly stated:

> "Receipts provided to the NDB by the National Health Insurance Federation and Payment Fund are after the first examination, that is, after the assessment by the examination body. After that, there may be an insurer assessment, or there may be a case of reinstatement due to a request for reexamination, but these do not reflect what kind of medical care was provided at each medical institution. Rather, it can be said that it is an indicator of what kind of medical care was approved for payment at the first examination level."

In other words, regional differences in NDB open data = "regional differences in doctors' intended medical practices" + "regional differences in assessment by examination" are mixed together.

---

## 2. Publicly available assessment-related data

### 2-1. Payment Fund “Examination Statistics” (by prefecture)
- **URL**: https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html
- **Content**: The examination status of original examination and reexamination is summarized by number of cases and points by **prefecture**.
- **Period**: FY 2011 - FY 2020 (with monthly data)
- **Grainness**: Number of assessments and assessment scores by prefecture (by medical, dental, and dispensing). However, **Assessment breakdown by medical practice code is not disclosed**

### 2-2. Payment Fund “Visualization Report of Examination Differences”
- **URL**: https://www.ssk.or.jp/shinryohoshu/saikaisyou_torikumi/kashikarepo/
- **Contents**: **Visualization of differences in screening by prefecture** regarding ``examination information provision examples'', ``general handling of screening'', and ``computer checks with many sticky notes''**
- **Characteristics**: Specifies the differences in handling of specific medical procedures (e.g. pre-surgical screening tests for D-dimer, etc.) by prefecture.
- **Limitations**: Does not cover all medical procedures, only cases with notable differences

### 2-3. National Health Insurance Federation “Examination status”
- **URL**: https://www.kokuho.or.jp/statistics/shinsa/joukyou/
- **Contents**: Examination status by the National Health Insurance Federation (FY2020-2020)
- **Grainness**: Aggregation of national situation

### 2-4. Cabinet Office SCR data (indirect)
- **URL**: https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/index.html
- **Contents**: Sex and age adjusted SCR calculated from NDB (by prefecture, secondary medical area, city, ward, town, village)
- **Note**: SCR itself is calculated from post-assessment data

---

## 3. Known knowledge: Regional differences in assessment rates

According to m3.com's analysis (2023), the results of analyzing the examination status by medical branch every five years from the annual report of the payment fund:

- Prefectures with high assessment rates tend to be **high in the west and low in the east**
- **Osaka and Fukuoka prefectures** frequently appear in the top 3, and in the past 10 years **Hokkaido** has also been prominent
- According to 2009 data, the difference in original examination assessment points per 10,000 requested points is approximately 4 times between the highest in Osaka (0.284%) and the lowest in Miyazaki (0.071%).

---

## 4. What can/cannot be done with public data?

### What can be done (feasible approaches)

| Approach | Data source | Verification content |
|---|---|---|
| **Quantification of regional differences in assessment rates** | Payment funds and examination statistics | Time-series analysis of assessment rate and assessment point rate by prefecture |
| **Correlation analysis between SCR and assessment rate** | SCR data + assessment statistics | Indirectly verifying whether SCR is low in areas with high assessment rates (whether the number of cases calculated on NDB is decreasing due to assessment) |
| **Analysis of differences in specific medical practices** | Visualization report + NDB open data | Comparing the medical practices for which differences were pointed out in the visualization report with regional differences on NDB |
| **Difference between Payment Fund and National Health Insurance Federation** | Screening statistics for both organizations | Impact of differences in screening standards by type of insurer |

### What cannot be done (data limitations)

- **Assessment breakdown by medical practice code**: Payment fund examination statistics only include the total amount (number of cases/points), and details of which medical practices were assessed are not disclosed.
- **Pre-assessment receipt information**: Data actually requested by the doctor (before assessment) is not made public → Direct differences with NDB open data cannot be obtained.
- **Impact of insurer assessment/reexamination**: Insurer assessment after the initial assessment and reinstatement due to reexamination are not reflected in the NDB, and the statistics are not disclosed systematically.
- **Analysis on an individual receipt basis**: Open data is aggregated values, so assessment impact at the individual receipt level cannot be verified.

---

## 5. Practical verification strategy (proposal)

### Step 1: Create a regional profile of assessment rates from assessment statistics
Download the payment fund examination statistics ZIP file and calculate the number of appraisals and the appraisal score rate by prefecture. Understand trends over time.

### Step 2: Match with SCR data
Compare the SCR data (by prefecture) published by the Cabinet Office with the assessment rate of Step 1. Using regression analysis, etc., we verified whether SCR tends to be systematically lower in regions with higher assessment rates.

### Step 3: Case-based analysis of visualization reports
Compare the medical practices (e.g. D-dimer) for which differences are specifically pointed out in the Payment Fund's "Visualization of Examination Differences Report" with the calculated number of such practices by prefecture in NDB open data.

### Step 4: Sensitivity analysis
Simulate how much impact regional differences in assessment rates (e.g. 0.07% to 0.28%) can have on SCR. Estimate the range of variation in regional differences when the SCR of other regions is corrected using the value of the region with the lowest assessment rate as the standard.

---

## 6. Conclusion

**Partially verifiable**.

- Payment fund examination statistics (by prefecture, number of cases/points) and visualization reports are publicly available, and regional differences in evaluation can be confirmed quantitatively.
- However, the absolute level of the assessment rate is low (approximately 0.1 to 0.3%), and compared to the regional differences in medical practices seen in NDB open data (variation of about 50 to 150 in SCR), there is a possibility that assessment alone has a small impact in explaining regional differences.
- On the other hand, it cannot be denied that regional differences in assessment have a significant impact on the number of individual calculations for specific medical practices (items pointed out in the visualization of assessment differences report).
- The biggest bottleneck is that **Assessment breakdown by medical practice code** is not made public. Even if NDB is provided to a third party (for research use), it may be difficult to access pre-assessment data.

As the SCR notes point out, this structural limitation cannot be completely overcome with public data alone, as long as the data is about ``what was paid for'' rather than ``what was provided.''

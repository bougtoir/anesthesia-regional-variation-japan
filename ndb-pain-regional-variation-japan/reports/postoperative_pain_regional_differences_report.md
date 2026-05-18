# 術後疼痛の地域差・文化差：オープンデータによる検証可能性と研究デザイン提案

## はじめに

痛みの表現・知覚には地域差・文化差があることは広く認められています。この報告では、**手術あたりの術後疼痛（postoperative pain per surgery）に地域差が存在するか**を、(1) オープンデータで検証可能かどうか、(2) 研究デザインとしてどうあるべきかについて整理します。

---

## 1. 既存エビデンスの概要

### 1.1 術後疼痛の国際比較に関する先行研究

| 研究 | データソース | 対象 | 主な知見 |
|------|-------------|------|---------|
| Zaslansky et al. (2018) BJA | PAIN OUT | 整形外科手術、米国 vs 国際 | 米国患者はworst pain intensity が高く、オピオイド使用量も多い。国際比較で有意差あり |
| PAIN OUT 10カ国研究 (2022) Eur J Pain | PAIN OUT | 10カ国10,415名 | 国・術科で説明できる分散は約12%に留まり、88%は施設間・個人間変動 |
| Huang et al. (2025) Glob Health Res Policy | PAIN OUT (アジア7地域) | 5,093名 | アジア地域内でも疼痛アウトカム・オピオイド使用量にネットワーク構造の差を確認 |
| Macchia et al. (2025) Commun Med | Gallup World Poll等 | 22カ国202,898名 | 一般痛の有病割合に大きな国間差（エジプト60% vs イスラエル25%） |
| Zimmer et al. (2022) Pain | WHO WHS等 | 52カ国 | 国レベルの文脈要因（GDP、医療費等）が痛み有病率と関連 |
| Al-Hashimi et al. (2015) Br J Pain | 単施設（英国） | 民族別（白人・南アジア・黒人） | 南アジア系患者は術後NRS疼痛スコアが高く、オピオイド処方が少ない |
| Rogger et al. (2023) Curr Pain Headache Rep | レビュー | 文化的枠組みと急性痛 | 文化・民族背景が疼痛の知覚・表出・管理に強く影響 |
| Jones et al. (2025) J Clin Anesth | メタ分析 | 人種・民族別術後疼痛管理 | 非白人患者は区域麻酔を受ける可能性が18%低い |

**要点**: 術後疼痛スコアと疼痛管理に地域差・文化差が存在する強いエビデンスがある。ただし「地域差」の内訳として、(a) 疼痛の知覚・表出の差、(b) 鎮痛プロトコルの差、(c) 施設レベルの質の差 が混在しており、純粋な「文化的疼痛知覚差」の分離は困難。

---

## 2. オープンデータによる検証可能性

### 2.1 候補データソース一覧

| データソース | アクセス | 国際比較 | 術後疼痛スコア | 手術種別 | 人種/民族/国 | 評価 |
|-------------|---------|---------|---------------|---------|------------|------|
| **PAIN OUT** (Jena大学) | 共同研究契約が必要（完全オープンではない） | 40カ国以上 | NRS (IPO-Q) | 術科別 | 国・施設 | **最適** |
| **MIMIC-IV** (PhysioNet) | 認証制（CITI修了後無料） | 米国単一施設 | pain score記録あり | 手術コードあり | 人種/民族あり | **次善** |
| **eICU-CRD** (PhysioNet) | 認証制 | 米国多施設 | 限定的 | ICU患者中心 | 人種あり | 補助的 |
| **INSPIRE** (韓国) | PhysioNet経由 | 韓国単一施設 | バイタルのみ | 手術コードあり | 単一国 | 限定的 |
| **NIS/HCUP** (AHRQ) | 有料ライセンス | 米国全国 | 疼痛スコアなし | ICD手術コード | 人種あり | 不適 |
| **ACS-NSQIP** | 参加施設のみ | 米国+カナダ一部 | 疼痛スコアなし | 術式別 | 人種あり | 不適 |
| **GBD 2021** (IHME) | 無料ダウンロード | 204カ国 | 慢性疼痛DALYs | 術後に非特異的 | 国別 | 間接的 |
| **Gallup World Poll** | 有料ライセンス | 140カ国以上 | "physical pain yesterday" | 術後に非特異的 | 国別 | 間接的 |
| **QUIPS** (ドイツ) | ドイツ国内共同研究 | ドイツ国内 | NRS | 術式別 | 限定的 | ドイツ限定 |

### 2.2 各データソースの詳細評価

#### A. PAIN OUT レジストリ（最有力）
- **概要**: Jena大学が運営する国際術後疼痛レジストリ（NCT02083835）。2009年開始、目標200,000例、2030年まで継続。
- **変数**: IPO-Q（International Pain Outcomes Questionnaire）による多次元疼痛アウトカム（worst pain, time in severe pain, 機能障害, 満足度等）、鎮痛薬使用量、患者背景
- **地域**: 40カ国以上、20言語以上でバリデーション済み
- **アクセス**: 完全オープンデータではない。共同研究提案を Jena 大学 PAIN OUT チーム（ruth.zaslansky@med.uni-jena.de）に提出する必要がある。ただし、既に多数の外部共同研究論文が出版されており、アクセスのハードルは比較的低い。
- **検証可能性**: **高い**。国・地域を独立変数、術式を層別因子として、術後疼痛スコアの地域差を多水準モデルで検証可能。

#### B. MIMIC-IV（単独で検証可能）
- **概要**: Beth Israel Deaconess Medical Center（米国ボストン）のEHR、約130,000入院、ICU含む
- **変数**: pain score（chartevents内）、手術コード、人種/民族、年齢、性別、各種臨床データ
- **制限**: 単一施設・単一国のため「地域差」は直接検証不可。ただし、**同一施設内での人種/民族間差**は検証可能
- **アクセス**: PhysioNet credentialed access（CITI Human Subjects Research コース修了後、無料）
- **検証可能性**: **中程度**。「地域差」ではなく「同一医療環境下での民族間差」として部分的に検証可能。同一鎮痛プロトコル下での差を見ることで、文化的疼痛知覚差に迫れる可能性。

#### C. GBD 2021 + Gallup World Poll（生態学的研究として）
- **概要**: 国レベルの疼痛負荷（DALYs）や痛み有病率の国際比較データ
- **制限**: 術後疼痛に特異的でない。慢性疼痛（腰痛、頸部痛、変形性関節症等）が中心。
- **検証可能性**: **低い（間接的）**。「手術あたりの術後疼痛」の検証には不向きだが、背景としての国民の疼痛感受性の地域差を示す補助データとして有用。

### 2.3 総合判定

| 検証レベル | 実現可能性 | 推奨データ |
|-----------|-----------|----------|
| 国際間比較（本命） | 共同研究契約が必要 | PAIN OUT |
| 同一施設内の民族間比較 | すぐに着手可能 | MIMIC-IV |
| 国レベルの生態学的相関 | すぐに着手可能 | GBD 2021 + 手術件数データ |

**結論: オープンデータで「ある程度」検証可能だが、最も適切なデータ（PAIN OUT）は完全オープンではなく共同研究契約が必要。MIMIC-IVを用いた民族間比較は即座に実施可能。**

---

## 3. 研究デザイン提案

### 3.1 なぜ単純な国際比較では不十分か

PAIN OUT 10カ国研究(2022)が示したように、「国」変数が説明する分散は全体のわずか数%であり、施設間・個人間変動が圧倒的に大きい。これは以下の交絡因子が混在するためです：

- **鎮痛プロトコルの施設差**（最大の交絡）
- **術式構成の国間差**
- **患者背景の系統的差異**（年齢、BMI、併存疾患）
- **医療制度・アクセスの差**
- **疼痛評価スケールの言語的同等性**

### 3.2 推奨デザイン

#### デザイン案A: 多水準横断研究（PAIN OUT活用・本命）

```
研究構造（3水準ネスト構造）:
  Level 3: 国 / 地域（文化圏）
  Level 2: 施設（病院）
  Level 1: 患者（個人）
```

**方法論**:
- **目的**: 術後疼痛アウトカムの分散を「国」「施設」「個人」に分割し、国レベルの効果量を推定
- **統計モデル**: 3水準ランダム切片モデル（multilevel model / hierarchical linear model）
  - アウトカム: worst pain NRS, % time in severe pain, 機能障害, 満足度
  - Level 1 共変量: 年齢, 性別, BMI, ASA-PS, 術式, 麻酔法, 併存疾患
  - Level 2 共変量: 病院規模, 疼痛管理体制（APS有無等）, 教育病院か否か
  - Level 3 共変量: 国（文化圏プロキシ）, HDI, 医療費対GDP比, オピオイド規制レベル
- **サンプルサイズ**: Level 3 で最低15カ国（ICC推定の安定性のため）、各施設50-100例以上
- **主要解析**: ICC（級内相関係数）の分解 → 国レベルICCが有意であれば地域差の存在を示唆
- **感度分析**: 術式を層別化（例: TKA, 帝王切開, 腹部手術）して術式交絡を除去

**利点**: 交絡の階層的制御が可能、既存PAIN OUTインフラを活用
**課題**: PAIN OUTデータアクセスの交渉が必要

#### デザイン案B: 同一施設内民族間比較（MIMIC-IV・即時実施可能）

```
研究構造: 後ろ向きコホート
  対象: MIMIC-IV内の手術患者
  曝露: 人種/民族（White, Black, Asian, Hispanic, Other）
  アウトカム: 術後pain score（NRS）の時系列
```

**方法論**:
- **目的**: 同一施設・同一鎮痛環境下で、人種/民族間の術後疼痛スコアに差があるか
- **統計モデル**: 
  - 混合効果モデル（反復測定: 術後0-72hのpain score trajectory）
  - 傾向スコアマッチング or IPTW（inverse probability of treatment weighting）で術式・背景因子を調整
- **強み**: 施設差・プロトコル差を完全に除去できる（単一施設のため）
- **交絡制御**: 年齢, 性別, BMI, ASA-PS, 術式（CPTコード）, 麻酔法, 術中オピオイド使用量, 併存疾患（Elixhauser/Charlson）
- **追加分析**: オピオイド消費量の民族間差（疼痛スコアだけでなく鎮痛需要も比較）
- **サンプルサイズ**: MIMIC-IVには約130,000入院があり、手術患者はその相当割合を占める。人種別にも十分なサンプルが期待できる（White ~68%, Black ~8%, Asian ~3%, Hispanic ~3%）

**利点**: 完全オープンデータで即時開始可能、施設差の交絡を完全排除
**課題**: 単一施設（ボストン）のため一般化可能性に限界。「地域差」というよりは「民族差」の検証。

#### デザイン案C: 生態学的研究（GBD + 手術件数データ）

```
研究構造: 国レベル生態学的横断研究
  単位: 国
  曝露: 文化圏/地域/SDI
  アウトカム: 疼痛関連DALYs / 手術件数 の比
```

**方法論**:
- GBD 2021の慢性疼痛DALYsデータ（国別）と、Lancet Commission on Global Surgery の手術件数推定データを組み合わせる
- **限界**: 術後疼痛に特異的でなく、生態学的誤謬のリスク大

### 3.3 改善提案のまとめ

| 現状の課題 | 改善提案 |
|-----------|---------|
| 国と施設の効果が分離できない | 3水準マルチレベルモデルを採用し、ICCを階層分解 |
| 術式構成の交絡 | 術式を層別化（TKA, C-section等の標準手術に限定） |
| 鎮痛プロトコルの施設差 | Level 2共変量に鎮痛管理体制を含める or MIMIC-IV型の単一施設デザイン |
| NRSの言語的同等性が不明 | IPO-Qの言語バリデーション論文を確認 + DIF（Differential Item Functioning）分析を追加 |
| 「文化」の操作的定義が曖昧 | Hofstedeの文化次元（個人主義vs集団主義等）やWVS（世界価値観調査）データをLevel 3共変量に |
| サンプルサイズが不十分な地域 | アジア・アフリカ・南米の施設リクルートを優先（PAIN OUT拡張） |
| 横断研究の限界 | 可能であれば同一患者の術前疼痛閾値測定（QST: 定量的感覚テスト）を追加する前向きデザイン |

---

## 4. 具体的アクションプラン

### Phase 1: 即時実施（0-3ヶ月）— MIMIC-IV分析
1. PhysioNet認証取得（CITI course修了）
2. MIMIC-IV内の手術患者コホート抽出
3. pain scoreデータの可用性・完全性を確認
4. 人種/民族別の術後疼痛スコア比較（調整済み）
5. プレプリント公開

### Phase 2: 共同研究開始（3-6ヶ月）— PAIN OUT
1. PAIN OUTチーム（Jena大学 Zaslansky教授）に研究提案書を送付
2. データ使用契約の締結
3. 3水準マルチレベルモデルの解析プラン策定
4. SAP（Statistical Analysis Plan）の事前登録（PROSPERO or OSF）

### Phase 3: 前向き研究（6-24ヶ月）— 新規データ収集
1. 3-5カ国（例: 日本, ドイツ, ブラジル, ナイジェリア, 米国）の施設リクルート
2. 標準手術（TKA or 帝王切開）に限定
3. IPO-Q + QST（術前疼痛閾値）+ 文化的変数（言語, 信仰, 痛みの信念尺度）の統一プロトコル
4. 事前登録（ClinicalTrials.gov）
5. 本論文

---

## 5. 結論

**Q1: オープンデータで検証可能か → 部分的にYES**
- MIMIC-IV を用いた「同一施設内の民族間差」は即座に検証可能（完全オープン）
- 国際間比較の本命である PAIN OUT レジストリは完全オープンではないが、共同研究契約により比較的容易にアクセス可能
- GBD等の完全オープンデータは術後疼痛に特異的でなく、間接的なエビデンスに留まる

**Q2: デザイン改善提案**
- **最大の課題は「国差」と「施設差」の分離**。3水準マルチレベルモデルが必須
- **術式の標準化**（TKA, C-section等に限定）により、術式交絡を排除
- **疼痛評価の文化的妥当性**（DIF分析）を組み込むべき
- **文化の操作的定義**にHofstede文化次元やWVSデータを活用
- 最も強力なデザインは、**複数国の標準手術患者に対する前向き統一プロトコル研究 + QST**

---

## 参考文献（主要）

1. Zaslansky R et al. (2018) Pain after orthopaedic surgery: differences in patient reported outcomes in the United States vs internationally. *Br J Anaesth* 120(4):790-797.
2. PAIN OUT Research Group (2022) Status quo of pain-related PROs and perioperative pain management in 10,415 patients from 10 countries. *Eur J Pain* 26:2120-2140.
3. Huang Y et al. (2025) Differentiating network structures and sex differences of pain-related outcomes in seven Asian regions. *Glob Health Res Policy* 10:51.
4. Macchia L et al. (2025) Demographic variation in pain across 22 countries. *Commun Med* 5:154.
5. Rogger R et al. (2023) Cultural Framing and the Impact on Acute Pain and Pain Services. *Curr Pain Headache Rep* 27:429-436.
6. Al-Hashimi M et al. (2015) Influence of ethnicity on the perception and treatment of early post-operative pain. *Br J Pain* 9(3):167-172.
7. Jones A et al. (2025) Racial and ethnic differences in acute post-operative pain management: Systematic review and meta-analysis. *J Clin Anesth.*
8. Zimmer Z et al. (2022) A global study of pain prevalence across 52 countries. *Pain* 163(9):1740-1750.
9. Shah N et al. (2024) Unraveling the Tapestry of Pain: Ethnic Variations, Cultural Influences, and Physiological Mechanisms. *Cureus* 16(5):e60692.
10. PAIN OUT website: https://www.pain-out.eu/
11. MIMIC-IV: https://physionet.org/content/mimiciv/
12. GBD 2021: https://ghdx.healthdata.org/gbd-2021/data-input-sources

---

# English Translation

---

# Regional and cultural differences in postoperative pain: Verifiability using open data and research design proposals

## Introduction

It is widely accepted that there are regional and cultural differences in the expression and perception of pain. In this report, we will examine whether there are regional differences in postoperative pain per surgery (1) whether it can be verified using open data, and (2) what the research design should be.

---

## 1. Overview of existing evidence

### 1.1 Previous research on international comparison of postoperative pain

| Research | Data source | Subject | Main findings |
|------|-------------|------|---------|
| Zaslansky et al. (2018) BJA | PAIN OUT | Orthopedic surgery, US vs. international | US patients have higher worst pain intensity and use more opioids. Significant difference in international comparison |
| PAIN OUT 10 countries study (2022) Eur J Pain | PAIN OUT | 10,415 people from 10 countries | Only about 12% of the variance can be explained by country and surgical department, and 88% is due to inter-facility/individual variation |
| Huang et al. (2025) Glob Health Res Policy | PAIN OUT (7 Asian regions) | 5,093 people | Confirmed differences in network structure in pain outcomes and opioid usage even within the Asian region |
| Macchia et al. (2025) Commun Med | Gallup World Poll, etc. | 202,898 people from 22 countries | Large differences between countries in the prevalence of general pain (60% in Egypt vs. 25% in Israel) |
| Zimmer et al. (2022) Pain | WHO WHS, etc. | 52 countries | Country-level contextual factors (GDP, medical expenses, etc.) are associated with pain prevalence |
| Al-Hashimi et al. (2015) Br J Pain | Single center (UK) | By ethnicity (white, South Asian, black) | South Asian patients have higher postoperative NRS pain scores and fewer opioid prescriptions |
| Rogger et al. (2023) Curr Pain Headache Rep | Review | Cultural framework and acute pain | Cultural and ethnic backgrounds strongly influence pain perception, expression, and management |
| Jones et al. (2025) J Clin Anesth | Meta-analysis | Postoperative pain management by race/ethnicity | Non-white patients are 18% less likely to receive regional anesthesia |
**Key Point**: There is strong evidence of regional and cultural differences in postoperative pain scores and pain management. However, ``regional differences'' include (a) differences in pain perception and expression, (b) differences in analgesic protocols, and (c) differences in facility-level quality, making it difficult to isolate pure ``cultural differences in pain perception.''

---

## 2. Verifiability through open data

### 2.1 List of candidate data sources

| Data Source | Access | International Comparison | Postoperative Pain Score | Surgery Type | Race/Ethnicity/Country | Rating |
|-------------|---------|---------|------|---------|------------|------|
| **PAIN OUT** (Jena University) | Collaborative research agreement required (not completely open) | More than 40 countries | NRS (IPO-Q) | By surgical department | Country/facility | **Optimal** |
| **MIMIC-IV** (PhysioNet) | Certification system (free after completing CITI) | Single facility in the United States | Pain score records | Surgery codes available | Race/ethnicity available | **Next best** |
| **eICU-CRD** (PhysioNet) | Certification system | US multi-center | Limited | ICU patient-centered | Race-based | Ancillary |
| **INSPIRE** (Korea) | Via PhysioNet | Single facility in Korea | Vitals only | With surgery code | Single country | Limited |
| **NIS/HCUP** (AHRQ) | Paid license | Nationwide | No pain score | ICD procedure code | Ethnicity | Unsuitable |
| **ACS-NSQIP** | Participating facilities only | US + some parts of Canada | No pain score | By procedure | With race | Not suitable |
| **GBD 2021** (IHME) | Free download | 204 countries | Chronic pain DALYs | Non-specific postoperatively | By country | Indirect |
| **Gallup World Poll** | Paid license | Over 140 countries | "physical pain yesterday" | Non-specific after surgery | By country | Indirect |
| **QUIPS** (Germany) | Joint research in Germany | Domestic Germany | NRS | By surgical procedure | Limited | Limited to Germany |

### 2.2 Detailed evaluation of each data source

#### A. PAIN OUT registry (most likely)
- **Overview**: International Postoperative Pain Registry (NCT02083835) operated by Jena University. Started in 2009 with goal of 200,000 cases and will continue until 2030.
- **Variables**: Multidimensional pain outcomes (worst pain, time in severe pain, functional impairment, satisfaction, etc.) by IPO-Q (International Pain Outcomes Questionnaire), analgesic usage amount, patient background
- **Region**: Validated in 40+ countries and 20+ languages
- **Access**: Not completely open data. Collaborative research proposals should be submitted to the Jena University PAIN OUT team (ruth.zaslansky@med.uni-jena.de). However, many external collaborative research papers have already been published, and the hurdles to access are relatively low.
- **Verifiability**: **High**. Using country/region as an independent variable and surgical method as a stratification factor, regional differences in postoperative pain scores can be verified using a multilevel model.

#### B. MIMIC-IV (Can be verified independently)
- **Summary**: EHR of Beth Israel Deaconess Medical Center (Boston, USA), approximately 130,000 hospitalizations, including ICU
- **Variables**: pain score (in chartevents), surgery code, race/ethnicity, age, gender, various clinical data
- **Limitations**: "Regional differences" cannot be directly verified due to single facility and single country. However, **racial/ethnic differences within the same facility** can be verified.
- **Access**: PhysioNet credentialed access (free after completing CITI Human Subjects Research course)
- **Verifiability**: **Moderate**. This can be partially verified as ``interethnic differences in the same medical environment'' rather than ``regional differences.'' By looking at the differences under the same analgesic protocol, it is possible to understand cultural differences in pain perception.

#### C. GBD 2021 + Gallup World Poll (as an ecological study)
- **Summary**: International comparative data on country-level pain burden (DALYs) and pain prevalence
- **Limitations**: Not specific for postoperative pain. Mainly chronic pain (lower back pain, neck pain, osteoarthritis, etc.).
- **Verifiability**: **Low (indirect)**. Although it is not suitable for verifying ``postoperative pain per surgery,'' it is useful as supplementary data that shows regional differences in national pain sensitivity as a background.

### 2.3 Overall Judgment

| Verification level | Feasibility | Recommended data |
|------------|------------|------------|
| International comparison (favorite) | Joint research agreement required | PAIN OUT |
| Comparison between ethnic groups within the same facility | Ready to start | MIMIC-IV |
| Country-level ecological correlates | Ready to start | GBD 2021 + surgery volume data |
**Conclusion: Verification is possible to some extent with open data, but the most appropriate data (PAIN OUT) is not completely open and requires a collaborative research agreement. Ethnic comparisons using MIMIC-IV can be performed immediately. **

---

## 3. Research design proposal

### 3.1 Why simple international comparisons are not enough

As shown in the PAIN OUT 10 Country Study (2022), the “country” variable explains only a few percent of the total variance, and the variation between facilities and individuals is overwhelmingly large. This is due to the following confounding factors:

- **Institutional differences in analgesic protocols** (biggest confound)
- **Differences in surgical composition between countries**
- **Systematic differences in patient demographics** (age, BMI, comorbidities)
- **Differences in healthcare systems and access**
- **Linguistic equivalence of pain rating scales**

### 3.2 Recommended design

#### Design proposal A: Multi-level cross-sectional research (PAIN OUT utilization/favorite)

````
Research structure (3-level nested structure):
  Level 3: Country/Region (Cultural Area)
  Level 2: Facility (hospital)
  Level 1: Patient (individual)
````

**Methodology**:
- **Purpose**: Divide the variance in postoperative pain outcomes into “country,” “facility,” and “individual,” and estimate country-level effect sizes.
- **Statistical model**: 3-level random intercept model (multilevel model / hierarchical linear model)
  - Outcome: worst pain NRS, % time in severe pain, functional impairment, satisfaction
  - Level 1 covariates: age, gender, BMI, ASA-PS, surgical method, anesthesia method, comorbidities
  - Level 2 covariates: hospital size, pain management system (APS, etc.), whether it is a teaching hospital or not
  - Level 3 covariates: country (cultural proxy), HDI, medical expenditure to GDP ratio, opioid regulation level
- **Sample size**: At least 15 countries at Level 3 (for stability of ICC estimation), 50-100 cases per site
- **Main analysis**: Decomposition of ICC (intraclass correlation coefficient) → If the country-level ICC is significant, it suggests the existence of regional differences
- **Sensitivity analysis**: Stratify by surgical procedure (e.g. TKA, caesarean section, abdominal surgery) to remove surgical confounding

**Advantages**: Hierarchical control of confounding possible, leveraging existing PAIN OUT infrastructure
**Challenge**: Need to negotiate PAIN OUT data access

#### Design proposal B: Comparison between ethnic groups within the same facility (MIMIC-IV, immediate implementation possible)

````
Study structure: retrospective cohort
  Target: Surgical patients in MIMIC-IV
Exposure: Race/Ethnicity (White, Black, Asian, Hispanic, Other)
  Outcome: Time series of postoperative pain score (NRS)
````

**Methodology**:
- **Purpose**: Are there differences in postoperative pain scores between racial/ethnic groups at the same facility and under the same analgesic environment?
- **Statistical Model**:
  - Mixed effects model (repeated measurements: pain score trajectory from 0-72h postoperatively)
  - Adjust surgical method and background factors using propensity score matching or IPTW (inverse probability of treatment weighting)
- **Strengths**: Can completely eliminate facility differences and protocol differences (because it is a single facility)
- **Confounding control**: age, gender, BMI, ASA-PS, surgical method (CPT code), anesthesia method, intraoperative opioid usage, comorbidities (Elixhauser/Charlson)
- **Additional Analysis**: Ethnic differences in opioid consumption (comparing not only pain scores but also analgesic demand)
- **Sample Size**: There were approximately 130,000 hospitalizations in MIMIC-IV, a significant proportion of which were surgical patients. We can expect sufficient samples by race (White ~68%, Black ~8%, Asian ~3%, Hispanic ~3%)
**Advantages**: Completely open data allows immediate start, completely eliminating confounding by facility differences.
**Challenges**: Single institution (Boston) limits generalizability. Examination of ``ethnic differences'' rather than ``regional differences.''

#### Design proposal C: Ecological study (GBD + surgery volume data)

````
Research structure: National ecological cross-sectional study
  Unit: Country
  Exposure: Culture/Region/SDI
  Outcome: Ratio of pain-related DALYs/number of surgeries
````

**Methodology**:
- Combine chronic pain DALYs data (by country) from GBD 2021 with surgery volume estimates from the Lancet Commission on Global Surgery
- **Limitations**: Not specific to postoperative pain, high risk of ecological fallacy

### 3.3 Summary of improvement proposals

| Current issues | Improvement proposals |
|------------|---------|
| Country and facility effects cannot be separated | Adopting a three-level multilevel model and hierarchically decomposing the ICC |
| Confounding surgical procedure composition | Stratification of surgical procedures (limited to standard surgeries such as TKA, C-section, etc.) |
| Institutional differences in analgesic protocols | Including analgesic management system as a Level 2 covariate or MIMIC-IV type single-institution design |
| NRS linguistic equivalence is unclear | Check IPO-Q language validation paper + Add DIF (Differential Item Functioning) analysis |
| The operational definition of "culture" is ambiguous | Hofstede's cultural dimensions (individualism vs. collectivism, etc.) and WVS (World Values Survey) data are used as Level 3 covariates |
| Regions with insufficient sample size | Prioritize recruitment to facilities in Asia, Africa, and South America (PAIN OUT expansion) |
| Limitations of cross-sectional studies | Prospective design with additional preoperative pain threshold measurement (QST: quantitative sensory testing) in the same patient if possible |

---

## 4. Specific action plan

### Phase 1: Immediate implementation (0-3 months) — MIMIC-IV analysis
1. Obtained PhysioNet certification (CITI course completion)
2. Extraction of surgical patient cohort within MIMIC-IV
3. Check availability and completeness of pain score data
4. Postoperative pain score comparison by race/ethnicity (adjusted)
5. Preprint release

### Phase 2: Start of joint research (3-6 months) — PAIN OUT
1. Submit a research proposal to the PAIN OUT team (Professor Zaslansky, Jena University)
2. Conclusion of data usage agreement
3. Formulation of analysis plan for 3-level multilevel model
4. Pre-registration of SAP (Statistical Analysis Plan) (PROSPERO or OSF)

### Phase 3: Prospective study (6-24 months) — New data collection
1. Facility recruitment in 3-5 countries (e.g. Japan, Germany, Brazil, Nigeria, USA)
2. Limited to standard surgery (TKA or Caesarean section)
3. Uniform protocol for IPO-Q + QST (preoperative pain threshold) + cultural variables (language, beliefs, pain belief scale)
4. Pre-registration (ClinicalTrials.gov)
5. This paper

---

## 5. Conclusion

**Q1: Can it be verified using open data? → Partially YES**
- "Ethnic differences within the same facility" using MIMIC-IV can be verified immediately (completely open)
- The PAIN OUT registry, which is the favorite for international comparisons, is not completely open, but it is relatively easy to access through a joint research agreement.
- Completely open data such as GBD is not specific to postoperative pain and remains indirect evidence.

**Q2: Design improvement proposal**
- **The biggest challenge is separating "national differences" and "facility differences"**. 3-level multilevel model required
- **Standardization of surgical procedures** (limited to TKA, C-section, etc.) eliminates surgical confounds
- **Cultural validity of pain assessment** (DIF analysis) should be incorporated
- Utilize Hofstede cultural dimensions and WVS data for **operational definition of culture**
- The most powerful design is a **prospective unified protocol study of standard surgery patients in multiple countries + QST**

---

## References (main)

1. Zaslansky R et al. (2018) Pain after orthopedic surgery: differences in patient reported outcomes in the United States vs internationally. *Br J Anaesth* 120(4):790-797.
2. PAIN OUT Research Group (2022) Status quo of pain-related PROs and perioperative pain management in 10,415 patients from 10 countries. *Eur J Pain* 26:2120-2140.
3. Huang Y et al. (2025) Differentiating network structures and sex differences of pain-related outcomes in seven Asian regions. *Glob Health Res Policy* 10:51.
4. Macchia L et al. (2025) Demographic variation in pain across 22 countries. *Commun Med* 5:154.
5. Rogger R et al. (2023) Cultural Framing and the Impact on Acute Pain and Pain Services. *Curr Pain Headache Rep* 27:429-436.
6. Al-Hashimi M et al. (2015) Influence of ethnicity on the perception and treatment of early post-operative pain. *Br J Pain* 9(3):167-172.
7. Jones A et al. (2025) Racial and ethnic differences in acute post-operative pain management: Systematic review and meta-analysis. *J Clin Anesth.*
8. Zimmer Z et al. (2022) A global study of pain prevalence across 52 countries. *Pain* 163(9):1740-1750.
9. Shah N et al. (2024) Unraveling the Tapestry of Pain: Ethnic Variations, Cultural Influences, and Physiological Mechanisms. *Cureus* 16(5):e60692.
10. PAIN OUT website: https://www.pain-out.eu/
11. MIMIC-IV: https://physionet.org/content/mimiciv/
12. GBD 2021: https://ghdx.healthdata.org/gbd-2021/data-input-sources


# NDBオープンデータで「術後遷延性疼痛（CPSP）」を拾えるか

## 結論：直接は拾えないが、プロキシ指標で間接的に検討可能

---

## 1. 直接捕捉できない理由

NDBオープンデータには**医科傷病分類（ICD診断名）が含まれていない**。

- 歯科傷病（う蝕・歯周病）のみ公表
- ICD-10コード G89.28（Other chronic postprocedural pain）等で直接検索不可能
- したがって「術後遷延性疼痛」という診断を直接都道府県別に集計することは**不可能**

## 2. 利用可能なプロキシ指標

### 2A. 神経障害性疼痛治療薬（外来処方）★最有力

NDB第10回オープンデータの外来処方薬（院外）に、以下が都道府県別で収載：

| 薬剤群 | 主な品目 | 外来処方総数 | 品目数 |
|---------|----------|-------------|--------|
| プレガバリン系 | リリカ、プレガバリンOD錠（多数GE） | 約10億錠規模 | 60+ |
| ミロガバリン | タリージェ | 約4億錠規模 | 10+ |
| デュロキセチン | サインバルタ、デュロキセチンCAP/OD錠 | 約3.5億錠規模 | 30+ |
| トラマドール | トラムセット、トラマドールOD錠 | 約1.1億錠規模 | 3 |
| ノイロトロピン | ノイロトロピン錠 | 約4.6億単位 | 1 |

**強み**: 都道府県別データあり → 手術件数で割ることで「手術あたり神経障害性疼痛薬外来処方」を算出可能

**限界**: 
- プレガバリンは糖尿病性神経障害、帯状疱疹後神経痛、線維筋痛症にも使用
- デュロキセチンはうつ病、不安障害にも使用
- **術後遷延性疼痛に特異的ではない**

### 2B. 神経ブロック手技（外来）★有力

NDB第10回の麻酔（L区分）に外来の神経ブロック算定回数が都道府県別で収載：

| 手技 | 外来算定回数（全国） |
|------|---------------------|
| 仙骨部硬膜外ブロック | 1,162,351 |
| 頸胸腰傍脊椎神経ブロック | 1,050,591 |
| 腰部硬膜外ブロック | 722,614 |
| 肩甲上神経ブロック | 510,018 |
| 星状神経節ブロック | 377,671 |
| トリガーポイント注射 | 74,738 |
| 神経ブロック持続注入 | 86,793 |

**強み**: 外来の神経ブロックは慢性疼痛管理を反映する可能性が高い

**限界**: CPSP以外の慢性疼痛（腰痛、CRPS、帯状疱疹後神経痛等）も含む

### 2C. 脊髄刺激装置植込術（K190）

- 難治性慢性疼痛の最終手段
- 入院・外来ともに都道府県別データあり
- **CPSPへの特異度は比較的高い**が、件数が少なく統計的検出力に限界

## 3. 実行可能なデザイン案

### デザインA: 外来神経障害性疼痛薬 / 手術件数 比の地域差（即時実行可能）

```
指標 = Σ(プレガバリン + ミロガバリン + デュロキセチン)外来処方数量 / Σ手術件数
```

- 都道府県別に算出し、地域ブロック間で比較
- Phase 1の「入院鎮痛薬/手術」を補完する**外来慢性疼痛プロキシ**
- 「手術後に外来で神経障害性疼痛薬を多く使う地域」を間接的に検出

**解釈の注意**: この指標は「全慢性神経障害性疼痛 / 全手術」であり、CPSPに限定されない。高齢化率・糖尿病有病率等で交絡する。

### デザインB: 外来神経ブロック / 手術件数 比の地域差（即時実行可能）

```
指標 = Σ(外来神経ブロック算定回数) / Σ手術件数
```

- 外来での神経ブロックは慢性疼痛管理を強く示唆
- 入院中の急性期神経ブロックと分離可能（外来のみ集計）

### デザインC: Phase 1との統合分析

- Phase 1の「入院鎮痛薬/手術」（急性期疼痛プロキシ）
- デザインA/Bの「外来慢性疼痛薬/手術」（遷延性疼痛プロキシ）
- 両指標の相関・乖離を地域別に検討
- **「急性期は鎮痛薬を多く使うが、慢性期の神経障害性疼痛薬は少ない地域」**＝急性期疼痛管理が良好でCPSP移行を抑制？

## 4. 真にCPSPを捕捉するために必要なデータ

| レベル | データソース | 何ができるか |
|--------|-------------|-------------|
| Level 1（現在） | NDBオープンデータ | 上記プロキシ指標のみ |
| Level 2 | NDBサンプリングデータセット | 手術→3ヶ月後のプレガバリン新規処方を同一患者で追跡可能 |
| Level 3 | NDB特別抽出 | 手術レセプト＋後続外来レセプトを傷病名付きでリンケージ。G89.28等で直接同定 |
| Level 4 | DPC個票データ | 術式別・施設別にCPSP発生率を精密推定 |

## 5. 推奨

1. **即座に実行可能**: デザインA〜Cを追加解析としてPhase 1論文に補足（Supplementary Analysis）
2. **論文のDiscussionで**: 「本研究の急性期鎮痛薬指標に加え、外来神経障害性疼痛薬をCPSPプロキシとして検討したが、疾患特異性の限界がある」と記載
3. **Phase 2への橋渡し**: NDBサンプリングデータセット（臨床疫学研究推進機構が公開）を用いて、手術後の神経障害性疼痛薬新規処方を同一患者で追跡する研究を提案

---

# English Translation

---

# Can “prolonged postoperative pain (CPSP)” be detected using NDB open data?

## Conclusion: Cannot be picked up directly, but can be considered indirectly using proxy indicators

---

## 1. Reasons why direct capture is not possible

NDB open data does not include **medical disease classification (ICD diagnosis name)**.

- Only dental injuries and diseases (caries and periodontal disease) are published.
- Direct search is not possible with ICD-10 code G89.28 (Other chronic postprocedural pain) etc.
- Therefore, it is **impossible** to directly aggregate the diagnosis of "prolonged postoperative pain" by prefecture.

## 2. Available proxy indicators

### 2A. Neuropathic pain treatment drug (outpatient prescription)★Most promising

The following are listed in NDB 10th open data outpatient prescription drugs (outside the hospital) by prefecture:

| Drug groups | Main items | Total number of outpatient prescriptions | Number of items |
|---------|---------|-------------|---------|
| Pregabalin series | Lyrica, Pregabalin OD tablets (many GE) | Approximately 1 billion tablets | 60+ |
| Mirogabalin | Tarige | Approximately 400 million tablets | 10+ |
| Duloxetine | Cymbalta, Duloxetine CAP/OD tablets | Approximately 350 million tablets scale | 30+ |
| Tramadol | Tramset, Tramadol OD tablets | Approximately 110 million tablets | 3 |
| Neurotropin | Neurotropin Tablets | Approximately 460 million units | 1 |

**Strengths**: Data by prefecture is available → By dividing by the number of surgeries, it is possible to calculate "outpatient prescriptions for neuropathic pain drugs per surgery"

**Limits**:
- Pregabalin is also used in diabetic neuropathy, postherpetic neuralgia, and fibromyalgia.
- Duloxetine is also used for depression and anxiety disorders
- **Not specific for persistent postoperative pain**

### 2B. Nerve block procedure (outpatient)★Possible

The number of outpatient nerve block calculations listed in NDB 10th Anesthesia (L category) by prefecture:

| Procedure | Calculated number of outpatient visits (nationwide) |
|------|-------|
| Sacral epidural block | 1,162,351 |
| Cervicothoracic lumbar paravertebral nerve block | 1,050,591 |
| Lumbar epidural block | 722,614 |
| Suprascapular nerve block | 510,018 |
| Stellate ganglion block | 377,671 |
| Trigger point injection | 74,738 |
| Continuous nerve block injection | 86,793 |

**Strengths**: Outpatient nerve blocks are more likely to reflect chronic pain management
**Limitations**: Includes chronic pain other than CPSP (lower back pain, CRPS, postherpetic neuralgia, etc.)

### 2C. Spinal cord stimulator implantation (K190)

- Last resort for intractable chronic pain
- Data by prefecture for both inpatient and outpatient care
- **Specificity for CPSP is relatively high**, but the number of cases is small and statistical power is limited

## 3. Viable design ideas

### Design A: Regional differences in outpatient neuropathic pain medicine/surgery ratio (immediately actionable)

````
Indicator = Σ (pregabalin + mirogabalin + duloxetine) outpatient prescription quantity / Σ number of surgeries
````

- Calculated by prefecture and compared between regional blocks
- **Outpatient chronic pain proxy** to complement Phase 1 “inpatient analgesics/surgery”
- Indirectly detect areas where neuropathic pain drugs are used frequently in outpatient settings after surgery

**Interpretation Note**: This indicator is “all chronic neuropathic pain/all surgeries” and is not limited to CPSP. Confounded by aging rate, diabetes prevalence, etc.

### Design B: Regional differences in outpatient nerve block/surgery ratio (immediately possible)

````
Index = Σ(calculated number of outpatient nerve blocks) / Σnumber of surgeries
````

- Outpatient nerve blocks strongly suggest chronic pain management
- Separable from acute nerve block during hospitalization (accounted for outpatients only)

### Design C: Integrated analysis with Phase 1
- Phase 1 “hospital analgesics/surgery” (acute phase pain proxy)
- Design A/B “outpatient chronic pain medicine/surgery” (prolonged pain proxy)
- Examining the correlation and divergence between both indicators by region
- ** “Regions where many analgesics are used in the acute phase, but few drugs for neuropathic pain are used in the chronic phase” **=Is acute phase pain management good and CPSP transition suppressed?

## 4. Data needed to truly capture CPSP

| Level | Data Source | What You Can Do |
|--------|-------------|-------------|
| Level 1 (current) | NDB open data | Only the above proxy indicators |
| Level 2 | NDB sampling dataset | New prescription of pregabalin 3 months after surgery can be tracked in the same patient |
| Level 3 | NDB special extraction | Linkage of surgical receipt + subsequent outpatient receipt with injury/disease name. Direct identification with G89.28 etc. |
| Level 4 | DPC individual data | Precise estimation of CPSP incidence by surgical procedure and facility |

## 5. Recommended

1. **Ready to execute**: Designs A to C are supplementary analyzes to the Phase 1 paper (Supplementary Analysis)
2. **Discussion of the paper**: ``In addition to the acute analgesic index in this study, outpatient neuropathic pain medications were examined as a proxy for CPSP, but there are limitations in disease specificity.''
3. **Bridge to Phase 2**: Propose a study to track new prescriptions for neuropathic pain drugs after surgery in the same patient using the NDB sampling dataset (published by the Japan Clinical Epidemiology Research Promotion Agency)

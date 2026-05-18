# L008全身麻酔の名称変更は査定頻度から妥当だったか？ ブレスト

## 0. 変更の概要

| | 旧（〜令和6年） | 新（令和8年〜） |
|---|---|---|
| **名称** | マスク**又は**気管内挿管による閉鎖循環式全身麻酔 | **声門上器具使用又は**気管内挿管による閉鎖循環式全身麻酔 |
| **気道確保手段** | フェイスマスク／気管内挿管 | 声門上器具（SGA）／気管内挿管 |

**変更点**: 「マスク」→「声門上器具使用」に置換。気管内挿管はそのまま。

---

## 1. この変更が意味すること

### 1-1. 何が入り、何が出たか

| 気道確保デバイス | 旧L008 | 新L008 |
|---|---|---|
| 気管内挿管（ETT） | 対象 | 対象（変更なし） |
| ラリンジアルマスク（LMA） | グレーゾーン（「マスク」に含む？） | **明確に対象**（声門上器具） |
| i-gel等 第2世代SGA | グレーゾーン | **明確に対象** |
| フェイスマスクのみ | 名目上は対象 | **対象外の可能性**（声門上器具でも気管内挿管でもない） |

### 1-2. 臨床的インパクト

- **声門上器具（SGA）が明確にL008の対象に**: LMA・i-gel等を用いた閉鎖循環式全身麻酔が、名称上の根拠をもって算定可能に
- **フェイスマスクのみのGA**: 「マスク」が名称から消えたことで、フェイスマスクのみの閉鎖循環式全身麻酔がL008で算定できるかが新たなグレーゾーンに（あるいは意図的に除外）

---

## 2. 「マスク」のグレーゾーン問題と査定の構造

### 2-1. 旧L008における「マスク」の曖昧さ

L008の「マスク」は歴史的にはフェイスマスク（バッグマスク換気）を指していた。L008が制定された時代（1980年代以前の点数表体系）にはSGAは存在しなかった。

しかし1988年のLMA臨床導入以降、SGAは急速に普及。現在では：
- LMA/i-gelは「ラリンジアル**マスク**」「声門上**マスク**」とも呼ばれる
- 臨床現場では気管挿管を要しない全身麻酔の大半がSGA管理に移行
- 日本麻酔科学会のガイドラインでもSGAは標準的気道確保デバイスとして位置づけ

**問題**: 審査の場で「マスク」にSGAが含まれるかは**解釈の余地**があった。

### 2-2. 査定が発生しうるパターン

```
パターンA: SGA使用 → L008算定 → 審査委員「マスクでも挿管でもない」→ 査定
パターンB: SGA使用 → L008算定 → 審査委員「マスクの一種として認める」→ 通過
パターンC: フェイスマスクのみ → L008算定 → 通過（名称上問題なし）
```

**パターンAとBが都道府県ごとに異なれば、典型的な「審査の差異」になる。**

### 2-3. 審査の差異の可視化レポートとの関連

支払基金の「審査の差異の可視化レポート」は、まさにこの種の問題（同一の診療行為に対する都道府県間の取扱いの違い）を検出・是正するために設計されている。

- L008 + SGA の組み合わせが可視化レポートに掲載されていたかどうかは非公開事例もあり確認困難
- しかし、**名称変更（告示レベルの改正）で対応した**ということは、通知・疑義解釈レベルでは解消しきれない構造的な問題だった可能性を示唆

---

## 3. 査定頻度から「妥当だった」と推論できる根拠

### 仮説: 査定頻度の蓄積が改定の動機の一つだった

| 根拠 | 説明 |
|---|---|
| **SGAの普及と名称の乖離** | 臨床現場ではSGAが標準なのに点数表名称が「マスク又は気管内挿管」のまま数十年放置 → 審査現場での解釈の揺れが恒常化 |
| **告示レベルの名称変更** | 通知や疑義解釈（Q&A）ではなく、告示（法令）レベルで名称を変更したのは、問題の重大性を示唆。疑義解釈で済む問題なら名称変更は不要 |
| **支払基金の審査差異是正の文脈** | 支払基金は2019年以降、審査の差異の可視化・是正に注力。不合理な差異の解消には、点数表自体の明確化が最も根本的な手段 |
| **麻酔科学会からの要望** | 医療技術評価分科会を通じて、学会が「現行名称では臨床実態と乖離」と提案した可能性が高い（学会提案は693件中665件が学会発） |
| **コンピュータチェックとの整合** | レセプトコンピュータチェックで「SGAの材料コード + L008」の組み合わせがフラグされるケースがあれば、自動的に査定候補になる |

### 定量的に推論できること

1. **L008の算定件数**: NDBオープンデータからL008の年間算定件数は推計可能（数百万件規模）
2. **SGA使用割合**: 臨床文献では全身麻酔の20-40%がSGA管理（施設により大きく異なる）
3. **査定率の幅**: 支払基金の審査統計から、麻酔関連の全体査定率は0.1-0.3%程度
4. **仮定**: SGA使用例のうち1%でも「マスクに該当しない」として査定されていたとすれば、年間数千〜数万件のインパクト

---

## 4. 査定頻度だけでは説明できない側面

### 4-1. 臨床実態への追従（モダナイゼーション）

名称変更の最大の動機は、単純に**点数表が臨床実態に追いついていなかった**ことかもしれない：

- SGAは1988年導入、本格普及は2000年代
- 旧L008の名称はそれ以前の体系を引きずっていた
- 査定がなくても、名称と実態の乖離は改正の十分な理由

### 4-2. フェイスマスク除外の意図

「マスク」→「声門上器具」への変更は、SGAを入れると同時にフェイスマスクのみの全身麻酔を**意図的に除外**した可能性がある：

- フェイスマスクのみのGA（挿管もSGAもなし）は、気道確保の確実性が低い
- 短時間手術でマスク換気のみで行うケースはあるが、「閉鎖循環式」の要件を厳密に満たすかは議論あり
- 安全性の観点から、SGA or 挿管を要件化する政策的判断の可能性

### 4-3. 査定の絶対数は多くない可能性

- 前回の分析で示した通り、全体の査定率は0.1-0.3%と低水準
- L008固有の査定率データは非公開
- 「査定は少ないが審査現場の解釈が割れていた」という**質的問題**が動機の可能性

---

## 5. ブレストまとめ: 査定頻度との関係の整理

### 「妥当だった」と言える観点

```
[査定の量的側面]
  SGA使用 × L008算定 → 一部都道府県で査定 
  → 年間数千件規模の影響（推計）
  → 名称変更で根本解消 ✓

[査定の質的側面]  
  「マスクにSGAが含まれるか」という解釈問題
  → 都道府県間の審査差異を生む構造的原因
  → 支払基金の差異是正の方針と整合 ✓

[臨床実態との整合]
  SGA普及から約30年間、名称が未更新
  → 査定の有無にかかわらず改正の合理性あり ✓
```

### 「査定頻度だけでは説明不十分」な観点

```
[データの不在]
  L008固有の査定件数・査定理由の公開データがない
  → 「査定頻度が高かったから改正された」と断言する根拠は公開情報からは得られない

[複合的動機]
  学会要望 + 審査差異是正 + 臨床実態追従 + 安全性政策
  → 査定頻度は動機の一つだが、唯一の理由ではない

[フェイスマスク除外の説明]
  査定頻度の問題はSGA→L008方向の査定
  フェイスマスク除外は逆方向の制限であり、査定頻度では説明しにくい
  → 安全性政策・臨床標準の反映と考えるのが自然
```

---

## 6. 検証可能な次のステップ

| ステップ | データソース | 何がわかるか |
|---|---|---|
| 支払基金の可視化レポートでL008/麻酔関連事例を確認 | [可視化レポート](https://www.ssk.or.jp/shinryohoshu/saikaisyou_torikumi/kashikarepo/) | SGA+L008が差異事例として掲載されていたか |
| NDBオープンデータでSGA材料（i-gel等）の都道府県別算定件数を確認 | [NDBオープンデータ](https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177182.html) | SGA使用の地域偏在パターン |
| 医療技術評価分科会の提案書でL008関連の学会提案を確認 | [中医協資料](https://www.mhlw.go.jp/stf/newpage_70414.html) | 名称変更の提案理由・エビデンス |
| 支払基金の審査統計でL区分（麻酔）の査定推移を確認 | [審査統計](https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html) | 麻酔全体の査定率の時系列変動 |

---

## 7. 結論（暫定）

**名称変更は査定頻度の観点から妥当だった**と推論できるが、それは複合的理由の一つとしてである。

最も蓋然性の高いストーリー：

> SGAの臨床普及 → L008の「マスク」との名称乖離 → 審査現場での解釈の分裂 → **都道府県間の査定差異** → 支払基金の差異是正プロセスで問題が可視化 → 麻酔科学会からの技術評価提案 → 令和8年改定で名称変更（告示レベル）

査定頻度は「氷山の一角」であり、その背後にある**審査基準の不統一**と**臨床実態との乖離**が本質的な改正動機と考えられる。

公開データからL008固有の査定件数を直接検証することは困難だが、可視化レポート・NDBオープンデータ・学会提案書の三方向から間接的な検証は可能。

---

# English Translation

---

# L008 Was the name change for general anesthesia appropriate based on the frequency of assessment? breast

## 0. Summary of changes

| | Old (~Reiwa 6) | New (Reiwa 8~) |
|---|---|---|
| **Name** | Closed-circulation general anesthesia with mask** or **endotracheal intubation | **Closed-circulation general anesthesia with supraglottic device or **endotracheal intubation |
| **Airway management measures** | Face mask/endotracheal intubation | Supraglottic appliance (SGA)/endotracheal intubation |

**Changes**: Replaced "mask" with "use of supraglottic device". Endotracheal intubation remains in place.

---

## 1. What this change means

### 1-1. What went in and what came out

| Airway management device | Old L008 | New L008 |
|---|---|---|
| Endotracheal intubation (ETT) | Target | Target (no change) |
| Laryngeal mask (LMA) | Gray zone (included in "mask"?) | **Clearly targeted** (supraglottic appliance) |
| 2nd generation SGA such as i-gel | Gray zone | **Clearly targeted** |
| Face mask only | Nominally covered | **Possibly not covered** (not supraglottic device or endotracheal intubation) |

### 1-2. Clinical impact
- **Supraglottic appliance (SGA) is clearly subject to L008**: Closed circulation general anesthesia using LMA, i-gel, etc. can now be calculated based on the name basis
- **Face mask only GA**: With the removal of “mask” from the name, whether closed circulation general anesthesia using only a face mask can be calculated using L008 is a new gray area (or intentionally excluded)

---

## 2. “Mask” gray zone problem and assessment structure

### 2-1. Ambiguity of "mask" in old L008

Historically, the "mask" in L008 referred to a face mask (bag mask ventilation). SGA did not exist when L008 was established (score system before the 1980s).

However, since the clinical introduction of LMA in 1988, SGA has rapidly spread. Currently:
- LMA/i-gel is also called “Laryngeal **Mask**” and “Supraglottic **Mask**”
- In clinical practice, most general anesthesia that does not require tracheal intubation has shifted to SGA management
- The Japanese Society of Anesthesiologists guidelines also position SGA as a standard airway management device.

**Problem**: At the hearing, there was **room for interpretation** as to whether "mask" included SGA.

### 2-2. Patterns in which assessment may occur

````
Pattern A: Use of SGA → L008 calculation → Examiner: “Neither mask nor intubation” → Assessment
Pattern B: Use of SGA → L008 calculation → Examiner: “Accepted as a type of mask” → Passed
Pattern C: Face mask only → L008 calculation → Passed (no problem with name)
````

**If patterns A and B differ from prefecture to prefecture, this will be a typical "difference in examination." **

### 2-3. Relationship with audit difference visualization report

The Payment Fund's ``Visualization of Examination Differences Report'' is designed to detect and correct exactly this type of problem: differences in the treatment of the same medical practice between prefectures.

- It is difficult to confirm whether the combination of L008 + SGA was published in the visualization report as there are some undisclosed cases.
- However, the fact that it was addressed by changing the name (amendment at the notification level) suggests that it may have been a structural problem that could not be resolved at the notification/question interpretation level.

---

## 3. Reasons for inferring that it was “valid” from the frequency of assessment

### Hypothesis: Accumulation of assessment frequency was one of the motivations for revision.

| Basis | Explanation |
|---|---|
| **The spread of SGA and the discrepancy in names** | Although SGA is the standard in clinical practice, the score sheet name has been left as "mask or endotracheal intubation" for several decades → Interpretation at examination sites is constantly fluctuating |
| **Name change at the notification level** | The name change at the notification (law) level, rather than the notification or question and interpretation (Q&A) level, indicates the seriousness of the issue. If the problem can be interpreted as questionable, there is no need to change the name |
| **Context of correction of discrepancies in examination of payment fund** | Since 2019, payment fund has focused on visualizing and correcting discrepancies in examination. The most fundamental way to eliminate unreasonable differences is to clarify the score sheet itself. |
| **Request from the Society of Anesthesiology** | It is highly likely that the society made a proposal through the Medical Technology Evaluation Subcommittee saying that the current name is out of touch with the clinical reality (665 out of 693 proposals were submitted by the society) |
| **Consistency with computer check** | If there is a case where the combination "SGA material code + L008" is flagged in the receipt computer check, it will automatically become a candidate for assessment |

### What can be inferred quantitatively

1. **Calculated number of L008 cases**: The annual number of L008 calculations can be estimated from NDB open data (several million cases)
2. **SGA usage rate**: In the clinical literature, 20-40% of general anesthesia is managed by SGA (varies greatly by institution)
3. **Range of assessment rate**: Based on the examination statistics of the payment fund, the overall assessment rate related to anesthesia is around 0.1-0.3%.
4. **Hypothesis**: If even 1% of SGA use cases were assessed as “not applicable to masks”, the impact would be thousands to tens of thousands of cases per year.

---
## 4. Aspects that cannot be explained by assessment frequency alone

### 4-1. Follow-up to clinical reality (modernization)

The biggest motivator for the name change may simply be that scorecards had not kept up with clinical reality:

- SGA was introduced in 1988 and became fully popular in the 2000s.
- The old L008 name was inherited from the previous system.
- Even without an assessment, the discrepancy between the name and reality is sufficient reason for revision.

### 4-2. Intent of face mask exclusion

The change from "mask" to "supraglottic device" may have intentionally excluded general anesthesia with only a face mask at the same time as inserting SGA:

- Facemask-only GA (no intubation or SGA) is less reliable in securing the airway
- There are cases in which short surgery is performed using only mask ventilation, but there is debate as to whether it strictly satisfies the requirements for a "closed circulation system."
- Possibility of a policy decision to require SGA or intubation from a safety perspective

### 4-3. The absolute number of assessments may not be large.

- As shown in the previous analysis, the overall assessment rate is low at 0.1-0.3%
- L008 specific assessment rate data is not disclosed
- Possibly motivated by a **qualitative issue**: ``Although the number of appraisals was small, interpretations at the appraisal site were divided.''

---

## 5. Brain summary: Organizing the relationship with assessment frequency

### Point of view that can be said to be “reasonable”

````
[Quantitative aspects of assessment]
Use of SGA × L008 calculation → Assessed in some prefectures
  → Impact on thousands of cases per year (estimated)
  → Fundamentally resolved by name change ✓

[Qualitative aspects of assessment]
  Interpretation question: “Does the mask contain SGA?”
  → Structural causes of examination differences between prefectures
  → Consistent with the payment fund discrepancy correction policy ✓

[Consistency with clinical reality]
  The name has not been updated for about 30 years since SGA became popular.
  → Revision is reasonable regardless of whether there is an assessment ✓
````

### Perspective that “assessment frequency alone is insufficient explanation”

````
[Absence of data]
  There is no public data on the number of assessments and reasons for assessments specific to L008.
  → There is no basis for asserting that “the revision was made because the assessment frequency was high” from public information.

[Multiple motives]
  Requests from academic societies + Correction of review discrepancies + Follow-up of clinical practice + Safety policy
  → Assessment frequency is one motivator, but not the only one.

[Explanation of face mask exclusion]
  The issue of assessment frequency is assessment in the direction of SGA → L008
  Excluding face masks is a restriction in the opposite direction and is difficult to explain by assessment frequency.
  → It is natural to think that this is a reflection of safety policy and clinical standards.
````

---

## 6. Verifiable next steps

| Steps | Data Sources | What We Learn |
|---|---|---|
| Check the L008/anesthesia-related case in the payment fund visualization report | [Visualization report](https://www.ssk.or.jp/shinryohoshu/saikaisyou_torikumi/kashikarepo/) | Was SGA+L008 listed as a difference case |
| Check the calculated number of SGA materials (i-gel, etc.) by prefecture with NDB open data | [NDB open data](https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177182.html) | Regional uneven distribution pattern of SGA use |
| Confirm L008-related academic proposals in the Medical Technology Evaluation Subcommittee proposal | [Chuikyo materials](https://www.mhlw.go.jp/stf/newpage_70414.html) | Reasons and evidence for the name change proposal |
| Check the assessment trends for Category L (anesthesia) using the payment fund's assessment statistics | [Examination statistics](https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html) | Time-series fluctuations in assessment rates for overall anesthesia |

---

## 7. Conclusion (tentative)

It can be inferred that the name change was appropriate from the perspective of assessment frequency, but this is one of multiple reasons.

Most likely story:
> Clinical dissemination of SGA → Discrepancy in name from L008 “Mask” → Divided interpretation at the examination site → **Differences in assessment between prefectures** → Problems became visible in the process of correcting differences in payment funds → Technical evaluation proposal from the Society of Anesthesiologists → Name changed in 2020 revision (notification level)

The frequency of assessment is the "tip of the iceberg," and the underlying reasons for the revision are considered to be **inconsistency in screening standards** and **difference from clinical reality**.

Although it is difficult to directly verify the number of L008-specific assessments from public data, indirect verification is possible from three directions: visualization reports, NDB open data, and academic proposals.

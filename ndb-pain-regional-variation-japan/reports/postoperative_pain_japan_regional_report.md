# 日本国内における術後疼痛の地域差：オープンデータによる検証可能性と研究デザイン提案

## はじめに

「東北の人は我慢強い」という通説は広く知られていますが、術後疼痛の報告・鎮痛薬使用に地域差として現れるかは未検証です。本レポートでは、**日本国内のオープンデータ**に限定し、(1) 検証可能性、(2) 研究デザイン改善提案を整理します。

---

## 1. 日本国内の関連エビデンス

### 1.1 オピオイド処方の地域差（既に実証済み）

| 研究 | データ | 主な知見 |
|------|--------|---------|
| Matsuoka et al. (2025) Jpn J Clin Oncol | DeSC（レセプトDB）119,850例 | がん終末期オピオイド処方に**明確な地域差**あり。オキシコドン注射：東海16.4% vs 四国4.0%（4倍差）。フェンタニル貼付：九州・沖縄51.5% vs 南関東25.4%。近畿はオピオイド処方OR 0.68（南関東基準） |
| Shoji & Akazawa (2025) | NDBオープンデータ（2015-2021） | 強・弱オピオイド処方量の経年トレンドを都道府県別に分析。**NDBオープンデータで薬効分類別・都道府県別の処方量比較が可能**であることを実証 |
| Tamiya et al. (2025) J Opioid Manag | 茨城県NHIレセプト 6,041例 | 胸部手術後の持続オピオイド使用率3.3%。県単位での術後オピオイド使用実態を把握 |

**重要な示唆**: がん疼痛領域では既に日本国内の地域差が実証されている。術後疼痛領域でも同様の地域差が存在する可能性は高い。

### 1.2 日本の術後疼痛管理の現状

| 研究 | データ | 主な知見 |
|------|--------|---------|
| Kaibori et al. (2025) J Clin Med | 多施設前向き観察 21施設 | 日本初の術後疼痛に関する多施設前向き調査。術後鎮痛が不十分な実態を確認。**ただし施設は関東〜関西中心で東北の施設が少ない** |
| Yabuki et al. (2025) Eur J Pain | DPC 肺癌低侵襲手術 | 周術期鎮痛法が術後慢性鎮痛薬処方に影響。**東北大学グループがDPCデータを用いた術後疼痛研究を推進中** |
| 天谷 (2023) 京府医大誌 | レビュー | 術後疼痛管理チーム（POPS）の重要性。2022年度に術後疼痛管理チーム加算が新設 |

### 1.3 「東北の我慢強さ」に関するエビデンス

| ソース | 内容 |
|--------|------|
| ファイザー日本法人 ネット調査（2017年, n=8,924） | 「長く続く痛みを我慢するか？」→ 栃木81.6%が最高、神奈川68.3%が最低。**東北が突出して高いわけではないが、北関東・地方部が高い傾向** |
| 統計数理研究所「国民性調査」 | 東北地方で「ねばり強い」が日本人の長所として60%前後の高い選択率 |
| 竹田・鑓水 (2016) 国立国語研究所 | 痛みの言語表現「ウズク」に明確な地域差。西日本で多用、東日本では限定的。**痛みの表現方法自体に東西差** |
| 産経新聞 痛み学入門 (2025) | 痛みの方言：「ひらつく」（秋田・富山・九州）、「せく」（中四国）等。地域によって痛みの表現語彙が異なり、**医療者との疼痛コミュニケーションに影響する可能性** |
| 碓井 (2017) 新潟青陵大 | 日本人の「我慢は美徳」文化が痛みの過少報告につながる社会心理学的メカニズムを指摘 |

---

## 2. 利用可能なオープンデータソース

### 2.1 完全オープン（誰でもダウンロード可能）

#### A. NDBオープンデータ（厚生労働省）⭐最有力
- **URL**: https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177182.html
- **内容**: レセプト情報の集計値（薬効分類別処方量、診療行為別算定回数）を都道府県別・年齢別で公開
- **術後疼痛との関連変数**:
  - 処方薬：都道府県別の鎮痛薬（解熱鎮痛消炎剤、麻薬性鎮痛剤等）処方量
  - 診療行為：術後疼痛管理チーム加算の都道府県別算定回数
  - 手術：K手術コード別の都道府県別実施件数
  - 注射薬：フェンタニル、アセトアミノフェン注射液等の都道府県別使用量
- **強み**: 悉皆性（ほぼ全国民カバー）、無料、ダウンロード即可能、第1回〜第10回（2014〜2023年度）の時系列データ
- **限界**: 集計値のみ（個票なし）、手術と鎮痛薬の紐付け不可、pain scoreなし
- **検証可能な仮説**: 「同一術式の手術件数あたりの鎮痛薬使用量に都道府県差がある」

#### B. DPCオープンデータ（中医協/e-Stat）
- **URL**: https://www.e-stat.go.jp/ → DPC導入の影響評価に係る調査
- **内容**: DPC病院約1,700施設の疾患別・術式別の症例数、在院日数を施設別に公開
- **術後疼痛との関連変数**:
  - 術式別症例数（手術あり/なし）、在院日数
  - 施設所在地（都道府県）
  - DPC分類コード
- **強み**: 施設レベルデータ、無料、術式の詳細分類
- **限界**: pain score・鎮痛薬使用の詳細なし。ただし**在院日数は疼痛管理の質の間接指標**
- **検証可能な仮説**: 「同一DPC分類の在院日数に地域差がある」（疼痛管理の質の代理指標として）

#### C. 国民生活基礎調査（厚生労働省/e-Stat）
- **内容**: 3年毎の大規模調査。自覚症状（腰痛、肩こり等）の有訴率を都道府県別に公開
- **限界**: 術後疼痛に特異的でない
- **補助的利用**: 「痛みの有訴率の地域差」の背景データとして

### 2.2 認証制（申請すれば利用可能）

#### D. NDB特別抽出データ（厚生労働省）
- **内容**: NDBの個票データ（匿名化済み）。手術レセプトと鎮痛薬処方を紐付け可能
- **術後疼痛との関連変数**: 手術コード + 同一入院中の鎮痛薬使用（薬剤名・用量・日数）+ 都道府県
- **アクセス**: 大学等研究機関から厚労省に利用申請。審査あり（数ヶ月〜1年）
- **検証可能性**: **最も高い**。手術単位で鎮痛薬使用量の地域差を直接検証可能

#### E. DPCデータ（個票）
- **内容**: E/Fファイル（入院明細レコード）。入院日単位の診療行為・薬剤使用が記録
- **アクセス**: 各病院が保有。東京大学・京都大学等が大規模DPCデータベースを構築（康永秀生教授グループ等）
- **術後疼痛との関連変数**: 手術日からの日数別の鎮痛薬使用（薬剤名・用量）、在院日数、施設所在地
- **検証可能性**: **非常に高い**。Yabuki et al. (2025)が実際にDPCデータで周術期鎮痛と術後慢性鎮痛の関連を解析済み
- **限界**: 共同研究契約が必要。ただし実績のある研究グループ複数あり（東北大・東大・京大等）

### 2.3 データソース比較一覧

| データ | アクセス | 個票 | 手術×鎮痛 紐付け | 都道府県 | Pain Score | 即時開始 |
|--------|---------|------|-----------------|---------|-----------|---------|
| NDBオープンデータ | 完全無料 | × | × | ○ | × | **○** |
| DPCオープンデータ | 完全無料 | △施設単位 | × | ○ | × | **○** |
| 国民生活基礎調査 | 完全無料 | × | × | ○ | × | ○ |
| NDB特別抽出 | 申請制 | ○ | **○** | ○ | × | △ |
| DPC個票 | 共同研究 | ○ | **○** | ○ | × | △ |
| 前向き多施設研究 | 新規実施 | ○ | ○ | ○ | **○** | × |

---

## 3. 研究デザイン提案

### 3.1 最大の方法論的課題

**術後疼痛の「地域差」を検証する際の本質的問題**:

日本国内に焦点を絞ることで、国際比較で問題となる「医療制度の差」「保険制度の差」は排除できます。しかし以下の交絡は残存します：

1. **施設間の鎮痛プロトコル差**: 同じTKAでも、施設Aは硬膜外+NSAIDs、施設Bはiv-PCA+acetaminophenのように異なる
2. **麻酔科医配置の地域差**: 都市部と地方で麻酔科医密度が異なり、疼痛管理の質に直接影響
3. **患者背景の地域差**: 高齢化率、併存疾患構成が都道府県で異なる
4. **「我慢強さ」の定量化困難**: pain scoreが低いのは「痛くない」のか「痛いが報告しない」のか区別できない

### 3.2 推奨デザイン（3段階）

---

#### Phase 1: 生態学的研究（即時開始・完全オープンデータ）

**目的**: 術後鎮痛薬使用量に都道府県間差が存在するかの探索的検証

**データ**: NDBオープンデータ + DPCオープンデータ

**方法**:
```
分析単位: 都道府県（n=47）

分子: 鎮痛薬の都道府県別処方量（NDBオープンデータ）
  - 解熱鎮痛消炎剤（薬効分類114）
  - 麻薬（アセリオ注、フェンタニル注 等）
  - 術後疼痛管理チーム加算 算定回数

分母: 都道府県別手術件数（DPCオープンデータ or NDB手術コード）
  - 標準手術に限定（TKA: K0821, 帝王切開: K8982, 腹腔鏡下胆嚢摘出: K6721 等）

調整変数: 
  - 人口構成（65歳以上割合）
  - 麻酔科医密度（医師・歯科医師・薬剤師調査）
  - DPC病院密度
  - 1人あたり医療費

地域ブロック:
  東北（青森・岩手・宮城・秋田・山形・福島）
  vs 関東 vs 関西 vs 九州 等

統計手法:
  - 都道府県別「手術あたり鎮痛薬使用量」の地図可視化
  - 地域ブロック間の差の検定（Kruskal-Wallis）
  - 多変量回帰（鎮痛薬使用量 ～ 地域ブロック + 調整変数）
  - Moran's I 等による空間的自己相関の検出
```

**期待される成果**: 
- 東北地方の手術あたり鎮痛薬使用量が他地域より低い場合 → 「我慢強さ」仮説と整合
- ただし生態学的研究のため因果推論は不可。鎮痛プロトコルの差か患者の訴えの差かは分離不能

**所要期間**: 1-2ヶ月（データダウンロード〜分析〜論文化）

---

#### Phase 2: レセプト個票研究（NDB特別抽出 or DPC共同研究）

**目的**: 手術単位で鎮痛薬使用量の地域差を検証（個人レベル交絡の調整）

**データ**: NDB特別抽出データ or 大学DPCデータベース

**方法**:
```
研究デザイン: 後ろ向きコホート研究

対象: 標準手術を受けた成人患者
  - 初回TKA（人工膝関節全置換術）← 推奨：手術標準化度が高い
  - 帝王切開
  - 腹腔鏡下胆嚢摘出術
  いずれか1術式に限定（術式間の異質性排除）

曝露: 施設所在地の都道府県（→地域ブロック）

主要アウトカム:
  - 術後3日間の総鎮痛薬使用量（経口モルヒネ換算mg）
  - 術後3日間のアセトアミノフェン/NSAIDs使用有無・量
  - 術後オピオイド使用日数
  - 追加レスキュー鎮痛回数

副次アウトカム:
  - 術後在院日数
  - 30日以内再入院率
  - 術後合併症（イレウス、肺炎等）

交絡調整:
  Level 1（患者）: 年齢, 性別, BMI, ASA-PS相当（Charlson comorbidity index）, 
                    術前オピオイド使用歴, 術前NSAIDs使用歴
  Level 2（施設）: 病床数, DPC群（大学病院/特定/標準）, 
                    麻酔科常勤医数, 術後疼痛管理チーム加算届出有無

統計手法:
  - 2水準マルチレベルモデル（患者→施設）
  - ICC算出：施設レベル/地域レベルが説明する分散割合
  - 地域ブロックをLevel 2のfixed effectとして投入
  - 感度分析：施設のランダム効果を除いた場合 vs 含めた場合の地域効果の変化
```

**この設計のポイント**:
- **「国差」ではなく「県差」**: 日本の国民皆保険制度下では保険制度差はゼロ。純粋に施設・地域・文化の効果を検出可能
- **DPCデータにはpain scoreがない**: → 鎮痛薬使用量を代理アウトカムとする。これは「患者が痛いと訴えた結果の処方」であり、「我慢する患者が多い地域では低く出る」ため、仮説と整合する方向性
- **術式限定**: TKAに限定すれば、術式による交絡を完全排除。TKAは全国で年間10万件以上実施されており検出力は十分

**所要期間**: 6-12ヶ月（申請〜データ取得〜分析）

---

#### Phase 3: 前向き多施設研究（新規データ収集）

**目的**: pain score（NRS）を直接測定し、疼痛の「知覚/報告」と「鎮痛薬使用」を分離

**データ**: 新規前向きデータ

**方法**:
```
研究デザイン: 多施設前向き観察研究

施設選定（最低10施設）:
  東北ブロック: 2-3施設（秋田・岩手・宮城等）
  関東ブロック: 2-3施設（東京・神奈川等）
  関西ブロック: 2-3施設（大阪・京都等）
  九州ブロック: 2-3施設（福岡・鹿児島等）
  ※各ブロック内で大学病院1 + 市中病院1の構成が理想

対象: 初回TKA患者、連続症例登録

測定項目:
  [1] 主要アウトカム（患者報告）
    - 術後NRS（安静時・体動時）: 術後6h, 24h, 48h, 72h
    - BPI-SF（Brief Pain Inventory Short Form）日本語版
    - 患者満足度（5段階）
    - 「痛みを医療者に伝えることへの抵抗感」尺度（新規開発 or 既存尺度）
    
  [2] 鎮痛薬使用量（カルテ抽出）
    - 経口モルヒネ換算の総オピオイド使用量
    - アセトアミノフェン/NSAIDs使用量
    - レスキュー回数
    - 硬膜外/神経ブロック実施有無
    
  [3] 文化的変数（★本研究の独自性）
    - 出身地（都道府県）※施設所在地と異なる場合あり
    - 居住年数
    - Pain Catastrophizing Scale（PCS）日本語版
    - 痛みの信念尺度（Pain Beliefs Questionnaire 日本語版）
    - 「我慢は美徳」尺度（碓井ら2017のフレームワーク参照、新規開発）
    - 家族構成・社会的サポート
    
  [4] 施設変数
    - 鎮痛プロトコル内容
    - 術後疼痛管理チーム有無
    - 麻酔科医密度

サンプルサイズ:
  ICC = 0.05（施設内相関）、地域ブロック間差 NRS 0.5点を検出
  → 1施設あたり50例 × 10施設 = 500例（検出力80%, α=0.05）

統計手法:
  - 3水準マルチレベルモデル（患者→施設→地域ブロック）
  - 媒介分析（mediation analysis）:
    地域 → Pain Catastrophizing → NRS
    地域 → 「我慢」尺度 → NRS
    地域 → 鎮痛プロトコル → NRS
    → 地域差の何%が「文化的態度」で説明され、何%が「プロトコル差」で説明されるかを分解
```

**この設計の核心**:
- **NRS（痛みスコア）と鎮痛薬使用量を同時に測定**することで、「東北の患者はNRSが低いが鎮痛薬も少ない」→ 我慢の影響 vs 「NRSは同等だが鎮痛薬が少ない」→ プロトコル差 を識別可能
- **「我慢」を直接測定**する尺度を組み込むことで、「地域差→我慢尺度→NRS」の媒介効果を検証
- **出身地と居住地を区別**することで、「東北出身だが東京の病院で手術」というケースでの文化的影響を評価

**所要期間**: 2-3年（倫理審査・施設リクルート・データ収集・分析）

---

## 4. 改善提案のまとめ

| 既存研究の限界 | 改善提案 |
|---------------|---------|
| 国際比較では医療格差がノイズ | **日本国内**に限定し、保険制度差を排除 |
| pain scoreデータがない（レセプト研究） | Phase 1-2では**鎮痛薬使用量を代理アウトカム**、Phase 3では**NRS直接測定** |
| 施設差と地域差が分離できない | **マルチレベルモデル**（患者→施設→地域）でICCを階層分解 |
| 術式の交絡 | **TKA単一術式に限定**（標準化度が高く、全国で十分な症例数） |
| 「我慢強さ」の定量化がない | **Pain Catastrophizing Scale + 新規「我慢」尺度**を開発・測定 |
| 鎮痛プロトコルの交絡 | Phase 2で**施設レベル変数として調整**、Phase 3で**直接記録** |
| 東北の施設が含まれていない | **東北ブロックを明示的にリクルート**（東北大グループが実績あり） |
| 痛みの言語表現の地域差が未考慮 | Phase 3で**痛みの方言・表現習慣**を測定、NRS報告との乖離を分析 |

---

## 5. 具体的アクションプラン

### 即時着手（Phase 1）: NDBオープンデータ分析
1. NDBオープンデータ第1〜10回をダウンロード
2. 鎮痛薬（解熱鎮痛消炎剤・麻薬）の都道府県別処方量を抽出
3. DPCオープンデータからTKA・帝王切開等の都道府県別手術件数を抽出
4. 「手術あたり鎮痛薬使用量」の都道府県別マップを作成
5. 東北ブロック vs 他地域の比較
6. 論文化（Pythonでの分析コード含めGitHub公開も可）

### 中期（Phase 2）: DPCデータ共同研究
- 東北大学（山内・藤森グループ：DPCデータ活用実績あり）への共同研究提案
- NDB特別抽出の利用申請（厚労省）
- TKA限定の後ろ向きコホート設計

### 長期（Phase 3）: 前向き多施設研究
- UMIN-CTR事前登録
- 「我慢」尺度の開発・バリデーション
- 10施設リクルート（東北・関東・関西・九州各2-3施設）
- 科研費申請（基盤B or C）

---

## 6. 結論

**Q1: 日本のオープンデータで検証できるか → YES（段階的に）**
- **Phase 1（即時）**: NDBオープンデータ + DPCオープンデータで「手術あたり鎮痛薬使用量の都道府県差」は完全オープンデータで今すぐ検証可能
- **Phase 2（半年〜）**: NDB特別抽出 or DPC個票データで「個人レベルの鎮痛薬使用量の地域差」を交絡調整付きで検証可能
- **Phase 3（2-3年）**: 前向き研究でpain scoreと「我慢」尺度を直接測定し、文化的要因の媒介効果を検証

**Q2: デザイン改善提案の核心**
1. **日本国内限定**で医療制度差を排除（今回の修正ポイント）
2. **TKA等の単一標準手術に限定**して術式交絡を排除
3. **マルチレベルモデル**（患者→施設→地域）でICCを階層分解
4. **「我慢」の直接測定**（Pain Catastrophizing Scale + 新規尺度）が最大の独自性
5. **鎮痛薬使用量とpain scoreの乖離**が「文化的疼痛報告バイアス」の指標となる
6. **東北の施設を明示的にリクルート**（既存研究では不足）

---

## 参考文献

1. Matsuoka Y et al. (2025) Population-based claims study of regional and hospital function differences in opioid prescribing for cancer patients who died in hospital in Japan. *Jpn J Clin Oncol* hyaf149.
2. Shoji T & Akazawa M (2025) Trends of Strong and Weak Opioid Prescriptions in Japan: A Cross-Sectional Study Based on Open Data from the National Database. CiNii.
3. Kaibori M et al. (2025) Prospective Survey of Postoperative Pain in Japan: A Multicenter, Observational Study. *J Clin Med* 14(4):1130.
4. Yabuki S et al. (2025) Exploring the impact of perioperative analgesia on postoperative chronic analgesic prescriptions in patients with lung cancer. *Eur J Pain* 29:e4774.
5. Tamiya N et al. (2025) Prevalence and risk factors for persistent opioid use after thoracic surgery in a prefecture of Japan. *J Opioid Manag*.
6. 天谷文昌 (2023) 術後疼痛管理チーム. *京府医大誌* 132(12):833-841.
7. 竹田晃子・鑓水兼貴 (2016) 痛みを表す言語表現ウズクの地域差. 国立国語研究所.
8. 碓井真史 (2017) 社会心理学からみる日本人の「我慢は美徳」. ファイザー・エーザイ共催セミナー.
9. NDBオープンデータ: https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177182.html
10. DPCオープンデータ: https://www.e-stat.go.jp/ (DPC導入の影響評価に係る調査)
11. 臨床疫学研究推進機構 NDBオープンデータ再集計: https://icer.tokyo/materials/ndb_opendata_replication/

---

# English Translation

---

# Regional differences in postoperative pain in Japan: Verifiability using open data and research design proposals

## Introduction

Although it is widely known that people in Tohoku are more patient, it has not been verified whether there are regional differences in postoperative pain reports and analgesic use. In this report, we limit ourselves to **open data in Japan** and organize (1) verifiability and (2) research design improvement proposals.

---

## 1. Related evidence in Japan

### 1.1 Regional differences in opioid prescribing (already demonstrated)

| Research | Data | Key findings |
|------|---------|---------|
| Matsuoka et al. (2025) Jpn J Clin Oncol | DeSC (Receipt DB) 119,850 cases | There are **clear regional differences** in opioid prescriptions in terminal cancer stages. Oxycodone injection: Tokai 16.4% vs. Shikoku 4.0% (4-fold difference). Fentanyl pasting: Kyushu/Okinawa 51.5% vs. Southern Kanto 25.4%. Opioid prescription OR in Kinki is 0.68 (South Kanto standard) |
| Shoji & Akazawa (2025) | NDB Open Data (2015-2021) | Analysis of secular trends in strong and weak opioid prescription amounts by prefecture. **Proof that it is possible to compare prescription amounts by drug class and prefecture using NDB open data** |
| Tamiya et al. (2025) J Opioid Manag | Ibaraki Prefecture NHI receipt 6,041 cases | Persistent opioid use rate after thoracic surgery 3.3%. Understand the actual state of postoperative opioid use on a prefectural basis |

**Important Implications**: Regional differences within Japan have already been demonstrated in the field of cancer pain. Similar regional differences are likely to exist in the area of ​​postoperative pain.

### 1.2 Current status of postoperative pain management in Japan

| Research | Data | Key findings |
|------|---------|---------|
| Kaibori et al. (2025) J Clin Med | Multicenter prospective observation 21 facilities | Japan's first multicenter prospective survey on postoperative pain. We confirmed that postoperative analgesia was insufficient. **However, the facilities are mainly from Kanto to Kansai, and there are few facilities in Tohoku** |
| Yabuki et al. (2025) Eur J Pain | DPC Minimally Invasive Surgery for Lung Cancer | Perioperative analgesia affects postoperative chronic analgesic prescription. **Tohoku University group is promoting post-operative pain research using DPC data** |
| Amaya (2023) Journal of Kyoto Prefectural University of Medicine | Review | Importance of the Postoperative Pain Management Team (POPS). Postoperative pain management team addition will be newly established in 2022 |

### 1.3 Evidence regarding “Tohoku’s patience”

| Source | Contents |
|--------|------|
| Pfizer Japan Corporation Online Survey (2017, n=8,924) | "Would you put up with long-lasting pain?" → Tochigi answered the highest at 81.6%, and Kanagawa answered the lowest at 68.3%. **Tohoku is not particularly high, but northern Kanto and regional areas tend to be high** |
| Institute of Statistical Mathematics "National Character Survey" | A high selection rate of around 60% in the Tohoku region for "persistent" as an advantage of Japanese people |
| Takeda and Yarimizu (2016) National Institute for Japanese Language and Linguistics | Clear regional differences in the linguistic expression of pain, "uzuku." Frequently used in western Japan, limited in eastern Japan. **East and West differences in the way pain is expressed** |
| Sankei Shimbun Introduction to Pain Studies (2025) | Dialects for pain: "Hiratsuku" (Akita, Toyama, Kyushu), "Seku" (Chushikoku), etc. Vocabulary to express pain differs depending on region, **potentially affecting pain communication with healthcare providers** |
| Usui (2017) Niigata Seiryo University | Points out the social psychological mechanism by which the Japanese culture of "patience is a virtue" leads to underreporting of pain |

---

## 2. Available open data sources

### 2.1 Fully open (anyone can download)

#### A. NDB Open Data (Ministry of Health, Labor and Welfare) ⭐Most likely
- **URL**: https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177182.html
- **Contents**: Aggregated values of receipt information (prescription amount by drug efficacy category, number of calculations by medical practice) are published by prefecture and age group.
- **Variables associated with postoperative pain**:
  - Prescription drugs: Prescription amount of analgesics (antipyretic analgesics, anti-inflammatory drugs, narcotic analgesics, etc.) by prefecture
  - Medical practice: Calculation number of post-operative pain management team additions by prefecture
  - Surgery: Number of operations performed by prefecture by K surgery code
  - Injectable drugs: Usage amount of fentanyl, acetaminophen injection, etc. by prefecture
- **Strengths**: Comprehensive (covers almost all the people), free of charge, available for immediate download, time series data from 1st to 10th (2014 to 2023)
- **Limitations**: Only aggregated values (no individual data), cannot link surgery and analgesics, no pain score
- **Verifiable hypothesis**: "There is a difference between prefectures in the amount of analgesics used per number of surgeries of the same type of surgery."

#### B. DPC open data (Chuikyo/e-Stat)
- **URL**: https://www.e-stat.go.jp/ → Survey on impact assessment of DPC introduction
- **Contents**: The number of cases and length of hospital stay by disease and surgical method at approximately 1,700 DPC hospitals is disclosed by facility.
- **Variables associated with postoperative pain**:
  - Number of cases by surgical method (with/without surgery), length of hospital stay
  - Facility location (prefecture)
  - DPC classification code
- **Strengths**: Facility level data, free of charge, detailed classification of surgical procedures
- **Limitations**: No details on pain score/analgesic use. However, ** length of hospital stay is an indirect indicator of the quality of pain management **
- **Testable hypothesis**: "There are regional differences in the length of stay for the same DPC classification" (as a proxy indicator of the quality of pain management)

#### C. Basic Survey on National Living (Ministry of Health, Labor and Welfare/e-Stat)
- **Contents**: Large scale survey every 3 years. Publishing the complaint rate for subjective symptoms (back pain, stiff shoulders, etc.) by prefecture
- **Limitations**: Not specific for postoperative pain
- **Auxiliary use**: As background data for "regional differences in pain complaint rate"

### 2.2 Certification system (available upon application)

#### D. NDB special extracted data (Ministry of Health, Labor and Welfare)
- **Contents**: NDB individual data (anonymized). Possible to link surgical receipt and analgesic prescription
- **Variables related to postoperative pain**: Surgery code + analgesic use during the same hospitalization (drug name, dose, number of days) + prefecture
- **Access**: Application for use from universities and other research institutions to the Ministry of Health, Labor and Welfare. Examination required (several months to 1 year)
- **Verifiability**: **Highest**. It is possible to directly verify regional differences in analgesic usage on a surgical-by-surgical basis.

#### E. DPC data (individual sheet)
- **Contents**: E/F file (hospitalization details record). Records medical treatment and drug use by day of hospitalization
- **Access**: Owned by each hospital. The University of Tokyo, Kyoto University, etc. have built a large-scale DPC database (Professor Hideo Yasunaga group, etc.)
- **Variables related to postoperative pain**: Analgesic use (drug name and dose) by number of days from the date of surgery, length of hospital stay, facility location
- **Verifiability**: **Very High**. Yabuki et al. (2025) actually analyzed the relationship between perioperative analgesia and postoperative chronic analgesia using DPC data.
- **Limitations**: Requires joint research agreement. However, there are several research groups with proven results (Tohoku University, Tokyo University, Kyoto University, etc.)

### 2.3 Data source comparison list

| Data | Access | Individual records | Linking surgery x analgesia | Prefecture | Pain Score | Immediate start |
|---------|---------|------|------|---------|------------|---------|
| NDB open data | Completely free | × | × | ○ | × | **○** |
| DPC open data | Completely free | △Facility unit | × | ○ | × | **○** |
| National Lifestyle Basic Survey | Completely free | × | × | ○ | × | ○ |
| NDB special extraction | Application required | ○ | **○** | ○ | × | △ |
| DPC individual ticket | Joint research | ○ | **○** | ○ | × | △ |
| Prospective multicenter study | New implementation | ○ | ○ | ○ | **○** | × |

---

## 3. Research design proposal

### 3.1 Biggest methodological challenge

**Essential issues when examining "regional differences" in postoperative pain**:

By focusing on Japan, we can eliminate ``differences in medical systems'' and ``differences in insurance systems,'' which are problems in international comparisons. However, the following confounds remain:

1. **Differences in analgesic protocols between facilities**: Even with the same TKA, facility A uses epidural + NSAIDs and facility B uses iv-PCA + acetaminophen.
2. **Regional differences in anesthesiologist allocation**: The density of anesthesiologists differs between urban and rural areas, which directly affects the quality of pain management.
3. **Regional differences in patient background**: Aging rates and comorbidity composition differ by prefecture.
4. **Difficulty quantifying "patience"**: It is difficult to distinguish between low pain scores due to "no pain" or "pain but not reporting it"

### 3.2 Recommended design (3 stages)

---

#### Phase 1: Ecological research (immediate start, fully open data)

**Purpose**: Exploratory verification of whether there are differences between prefectures in the amount of postoperative analgesics used.

**Data**: NDB open data + DPC open data

**Method**:
````
Unit of analysis: prefecture (n=47)

Molecule: Prescription amount of analgesics by prefecture (NDB open data)
  - Antipyretic analgesic anti-inflammatory agent (medicinal class 114)
  - Narcotics (Acerio injection, Fentanyl injection, etc.)
  - Postoperative pain management team addition calculation count

Denominator: Number of surgeries by prefecture (DPC open data or NDB surgery code)
  - Limited to standard surgeries (TKA: K0821, Caesarean section: K8982, laparoscopic cholecystectomy: K6721, etc.)

Tuning variables:
  - Population composition (percentage of people aged 65 and over)
  - Anesthesiologist density (physician/dentist/pharmacist survey)
  - DPC Hospital Density
  - Medical costs per person

Region block:
  Tohoku (Aomori, Iwate, Miyagi, Akita, Yamagata, Fukushima)
  vs Kanto vs Kansai vs Kyushu etc.

Statistical methods:
  - Map visualization of "amount of analgesics used per surgery" by prefecture
  - Test for differences between regional blocks (Kruskal-Wallis)
  - Multivariate regression (analgesic usage ~ regional block + adjustment variables)
  - Detection of spatial autocorrelation by Moran's I etc.
````

**Expected results**:
- If the amount of analgesics used per surgery in the Tohoku region is lower than in other regions → consistent with the "patience" hypothesis
- However, because this is an ecological study, causal inference is not possible. Differences in analgesic protocols and patient complaints cannot be separated.
**Required period**: 1-2 months (data download - analysis - publication)

---

#### Phase 2: Individual receipt research (NDB special extraction or DPC joint research)

**Purpose**: Verify regional differences in analgesic usage by surgery (adjusting for individual-level confounding)

**Data**: NDB special extracted data or university DPC database

**Method**:
````
Study design: Retrospective cohort study

Target: Adult patients undergoing standard surgery.
  - Initial TKA (total knee arthroplasty) ← Recommended: High degree of surgical standardization
  - Caesarean section
  - Laparoscopic cholecystectomy
  Limited to one surgical technique (eliminating heterogeneity between surgical techniques)

Exposure: Prefecture where the facility is located (→region block)

Main outcomes:
  - Total analgesic consumption for 3 days after surgery (oral morphine equivalent mg)
  - Presence and amount of acetaminophen/NSAIDs used for 3 days after surgery
  - Days of postoperative opioid use
  - Additional rescue analgesia times

Secondary outcomes:
  - Post-operative hospital stay
  - 30-day readmission rate
  - Postoperative complications (ileus, pneumonia, etc.)

Confounding adjustment:
  Level 1 (patient): age, gender, BMI, ASA-PS equivalent (Charlson comorbidity index),
                    History of preoperative opioid use, history of preoperative NSAIDs use
Level 2 (facility): Number of beds, DPC group (university hospital/specific/standard),
                    Number of full-time anesthesiologists, presence or absence of additional notification for post-operative pain management team

Statistical methods:
  - 2-level multilevel model (patient → facility)
  - ICC calculation: Percentage of variance explained by facility level/region level
  - Introducing regional blocks as fixed effects at Level 2
  - Sensitivity analysis: Changes in regional effects when excluding vs. including facility random effect
````

**Key points of this design**:
- **Prefectural differences, not national differences**: Under Japan's universal health insurance system, there are no differences in insurance systems. It is possible to detect purely the effects of facilities, regions, and culture.
- **DPC data does not include pain score**: → Use analgesic usage as a surrogate outcome. This is a ``prescription as a result of patients complaining of pain,'' and ``in areas where there are many patients who tolerate it, the incidence is low,'' so this is a direction that is consistent with the hypothesis.
- **Limited to surgical technique**: Limiting to TKA completely eliminates confounding by surgical technique. TKA is performed over 100,000 times a year nationwide and has sufficient detection power.

**Required period**: 6-12 months (application - data acquisition - analysis)

---

#### Phase 3: Prospective multicenter study (new data collection)

**Purpose**: Directly measure pain score (NRS) and separate pain perception/report from pain medication use.
**Data**: New forward-looking data

**Method**:
````
Study design: Multicenter prospective observational study

Facility selection (minimum 10 facilities):
  Tohoku block: 2-3 facilities (Akita, Iwate, Miyagi, etc.)
  Kanto block: 2-3 facilities (Tokyo, Kanagawa, etc.)
  Kansai block: 2-3 facilities (Osaka, Kyoto, etc.)
  Kyushu block: 2-3 facilities (Fukuoka, Kagoshima, etc.)
  *Ideal composition is 1 university hospital + 1 community hospital in each block.

Target: First-time TKA patients, consecutive case registration

Measurement items:
  [1] Primary outcome (patient report)
    - Postoperative NRS (at rest/moving): 6h, 24h, 48h, 72h after surgery
    - BPI-SF (Brief Pain Inventory Short Form) Japanese version
    - Patient satisfaction (5 levels)
    - “Reluctance to communicate pain to medical personnel” scale (newly developed or existing scale)
    
  [2] Analgesic usage (extracted from medical records)
    - Total opioid usage in terms of oral morphine
    - Acetaminophen/NSAIDs usage
    - Number of rescues
    - Epidural/nerve block performed or not
    
  [3] Cultural variables (★Uniqueness of this research)
    - Place of birth (prefecture) *May be different from facility location
    - Years of residence
- Pain Catastrophizing Scale (PCS) Japanese version
    - Pain Beliefs Questionnaire (Japanese version)
    - “Patience is a virtue” scale (referring to the framework of Usui et al. 2017, newly developed)
    - Family structure/social support
    
  [4] Facility variables
    - Analgesic protocol content
    - Presence of postoperative pain management team
    - Anesthesiologist density

Sample size:
  ICC = 0.05 (intra-facility correlation), regional block difference NRS 0.5 points detected
  → 50 cases per facility × 10 facilities = 500 cases (power 80%, α=0.05)

Statistical methods:
  - 3-level multilevel model (patient → facility → regional block)
  - Mediation analysis:
    Region → Pain Catastrophizing → NRS
    Region → “Patience” scale → NRS
    Region → Analgesic Protocol → NRS
    → Break down what percentage of regional differences are explained by "cultural attitudes" and what percentage are explained by "protocol differences"
````

**The heart of this design**:
- By measuring **NRS (pain score) and analgesic usage at the same time**, it is possible to distinguish between "Patients in Tohoku have lower NRS but less analgesics" → influence of patience vs. "NRS is the same but less analgesics" → protocol differences
- By incorporating a scale that directly measures "patience", we verified the mediation effect of "regional differences → patience scale → NRS"
- **Distinguish between place of birth and place of residence** to evaluate the cultural influence in the case of ``I am from Tohoku, but I had surgery at a hospital in Tokyo.''

**Required period**: 2-3 years (ethics review, facility recruitment, data collection and analysis)

---

## 4. Summary of improvement proposals

| Limitations of existing research | Suggestions for improvement |
|--------------|---------|
| Medical disparities are noise in international comparisons | Limit to **Japan** and eliminate differences in insurance systems |
| No pain score data (receipt study) | In Phase 1-2, **analgesic usage was used as a surrogate outcome**, and in Phase 3, **NRS direct measurement** |
| Facility differences and regional differences cannot be separated | Hierarchical decomposition of ICC using **multilevel model** (patient → facility → region) |
| Confounding technique | **Limited to a single TKA technique** (high degree of standardization, sufficient number of cases nationwide) |
| There is no quantification of "patience" | **Development and measurement of **Pain Catastrophizing Scale + new "patience" scale** |
| Confounding of analgesic protocol | **adjusted as a facility-level variable** in Phase 2, **directly recorded** in Phase 3 |
| Facilities in Tohoku are not included | **Explicitly recruit Tohoku block** (Tohoku University group has a track record) |
| Regional differences in verbal expressions of pain are not taken into consideration | In Phase 3, **pain dialects and expression habits** will be measured and discrepancies with NRS reports will be analyzed |

---

## 5. Specific action plan

### Immediate start (Phase 1): NDB open data analysis
1. Download NDB Open Data Parts 1 to 10
2. Extracting prescription amounts of analgesics (antipyretic, analgesic, anti-inflammatory drugs, narcotics) by prefecture
3. Extracting the number of surgeries such as TKA and caesarean sections by prefecture from DPC open data
4. Create a map of “analgesic usage per surgery” by prefecture
5. Comparison of Tohoku block vs. other regions
6. Publication of paper (including analysis code in Python can be published on GitHub)

### Mid-term (Phase 2): DPC data joint research
- Joint research proposal to Tohoku University (Yamauchi/Fujimori group: DPC data utilization experience)
- Application for use of NDB special extraction (Ministry of Health, Labor and Welfare)
- Retrospective cohort design limited to TKA
### Long-term (Phase 3): Prospective multicenter study
- UMIN-CTR pre-registration
- Development and validation of the “Patience” scale
- Recruiting 10 facilities (2-3 facilities each in Tohoku, Kanto, Kansai, and Kyushu)
- Application for Grants-in-Aid for Scientific Research (Basic B or C)

---

## 6. Conclusion

**Q1: Can it be verified using Japanese open data? → YES (step by step)**
- **Phase 1 (immediate)**: With NDB open data + DPC open data, "prefectural differences in the amount of analgesics used per surgery" can be verified immediately with completely open data
- **Phase 2 (six months or more)**: NDB special extraction or DPC individual data can be used to verify "regional differences in individual-level analgesic usage" with adjustment for confounding.
- **Phase 3 (2-3 years)**: Directly measure pain score and "patience" scale in a prospective study and verify the mediating effect of cultural factors

**Q2: The core of design improvement proposals**
1. **Limited to Japan** to eliminate differences in the medical system (points of this revision)
2. **Limited to a single standard surgery such as TKA** to eliminate surgical type confounding
3. Hierarchical decomposition of ICC using **multilevel model** (patient → facility → region)
4. **Direct measurement of “patience”** (Pain Catastrophizing Scale + new scale) is the most unique
5. **Discrepancy between analgesic usage and pain score** is an indicator of “cultural pain reporting bias”
6. **Explicitly recruit facilities in Tohoku** (insufficient in existing research)

---

## References

1. Matsuoka Y et al. (2025) Population-based claims study of regional and hospital function differences in opioid prescribing for cancer patients who died in hospital in Japan. *Jpn J Clin Oncol* hyaf149.
2. Shoji T & Akazawa M (2025) Trends of Strong and Weak Opioid Prescriptions in Japan: A Cross-Sectional Study Based on Open Data from the National Database. CiNii.
3. Kaibori M et al. (2025) Prospective Survey of Postoperative Pain in Japan: A Multicenter, Observational Study. *J Clin Med* 14(4):1130.
4. Yabuki S et al. (2025) Exploring the impact of perioperative analgesia on postoperative chronic analgesic prescriptions in patients with lung cancer. *Eur J Pain* 29:e4774.
5. Tamiya N et al. (2025) Prevalence and risk factors for persistent opioid use after thoracic surgery in a prefecture of Japan. *J Opioid Manag*.
6. Amaya Fumimasa (2023) Postoperative pain management team. *Kyoto Prefecture Medical University Journal* 132(12):833-841.
7. Akiko Takeda and Kanetaka Yarimizu (2016) Regional differences in the linguistic expression uzuku expressing pain. National Institute for Japanese Language and Linguistics.
8. Masashi Usui (2017) “Patience is a virtue” among Japanese people as seen from social psychology. Seminar co-sponsored by Pfizer and Eisai.
9. NDB open data: https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000177182.html
10. DPC Open Data: https://www.e-stat.go.jp/ (Survey related to impact assessment of DPC introduction)
11. Japan Institute for Clinical Epidemiology Research NDB open data re-aggregation: https://icer.tokyo/materials/ndb_opendata_replication/

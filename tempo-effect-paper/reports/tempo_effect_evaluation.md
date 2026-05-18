# 「テンポ効果再考」論文アイデア評価レポート

## 1. アイデアの概要

| 項目 | 内容 |
|------|------|
| **タイトル** | Tempo effects revisited: reporting simultaneously living population for policy relevance |
| **投稿先** | The Lancet（Correspondence、400語以内） |
| **核心的主張** | テンポ効果（初産年齢AFBのシフト）を「同時在生人口」（simultaneously living population）というストック指標に翻訳し、政策立案者に直感的な数値を提供する |
| **モデル** | 内生更新モデル＋Gompertz生存、AFB=20/25/30、σ（分散）narrow/baseline/wide |
| **対象国** | 中国（低出生率・長寿）vs コンゴ民主共和国（DRC、高出生率・短寿） |
| **主要結果** | 中国：100年後に約1.08倍（差≈6,200万人）、DRC：100年後に約0.76倍（差≈1億人） |

---

## 2. 先行文献マッピング

### 2.1 テンポ効果の基礎理論（Rate measures）

| 文献 | 内容 | 本アイデアとの関係 |
|------|------|-------------------|
| **Bongaarts & Feeney (1998)** "On the Quantum and Tempo of Fertility" *Population and Development Review* | 期間TFRからテンポ歪みを除去するtempo-adjusted TFRを提案。出生タイミングのシフトが期間指標を歪める現象を定式化 | **直接の出発点**。ただしBFは出生率指標の補正に留まり、人口ストック（在生人口数）への翻訳は行っていない |
| **Bongaarts & Feeney (2006)** "The Quantum and Tempo of Life-Cycle Events" *Vienna Yearbook of Population Research* | テンポ歪みの概念を出生・死亡・婚姻など全ライフサイクルイベントに一般化 | テンポ効果の一般性を示すが、やはりフロー（率）指標の補正に焦点 |
| **Sobotka (2004)** *Postponement of Childbearing and Low Fertility in Europe* | 欧州全域での出産延期とTFR低下の関係を包括的に分析。period-cohort比較、tempo-quantum分解を実施 | テンポ効果の実証的重要性を示す主要文献。ただし人口規模への翻訳は扱っていない |
| **Yoo & Sobotka (2018)** "Ultra-low fertility in South Korea" *Demographic Research* | 韓国のTFR<1.3をtempo効果で説明。テンポ・パリティ調整TFR（TFRp*）を適用 | 東アジアの超低出生率をテンポで説明する先行事例。本アイデアの中国分析の文脈と近いが、人口規模は扱わない |
| **Mazzuco & Zanotto (2025)** "Tempo effects in period TFR" *Demographic Research* | コホート視点からテンポ歪みを分散・歪度の変化に分解する最新手法 | 最新のテンポ研究。分散パラメータσを変動させる本アイデアと部分的に呼応 |

### 2.2 テンポ効果と人口規模の接続（Stock measures）— **最も競合的な先行研究群**

| 文献 | 内容 | 本アイデアとの関係 |
|------|------|-------------------|
| **Goldstein, Lutz, & Scherbov (2003)** "Long-Term Population Decline in Europe: The Relative Importance of Tempo Effects and Generational Length" *Population and Development Review* 29(4):699-707 | **テンポ効果を人口規模に直接接続した最重要先行研究**。出生タイミング変化の人口への影響を3要因に分解：(a) テンポ効果（一時的）、(b) 世代長効果（長期的成長率への影響）、(c) テンポ-カンタム交互作用。EU15カ国で数値シミュレーション実施。結論：200年間はテンポ効果が世代長効果を圧倒する | **最大の競合文献**。テンポ→人口規模という本アイデアの核心的移行をすでに実施済み。ただし①指標は人口規模（出生数比）であり「同時在生人口」という明示的フレーミングではない、②対象はEUのみ、③AFBではなくMACをアンカーにしている |
| **Lutz, O'Neill, & Scherbov (2003)** "Europe's Population at a Turning Point" *Science* 299:1991-1992 | 欧州の負の人口モメンタムを示し、テンポ効果の継続が人口減少と高齢化を加速させることを定量化。**追加的労働年齢人口×年数（person-years）**で政策含意を表現 | 「テンポ効果の政策関連性」を強調した先駆的論文。person-yearsという政策的指標を提案しており、本アイデアの「同時在生人口」と類似した政策翻訳の試み |
| **Chen, I-Chun (2020)** "Population Growth Due to Earlier Childbearing" *Modern Economy* 11:1966-1975 | **OLGモデルで出産年齢の前倒しが総人口に与える影響を直接モデル化**。テンポ効果、ベビーブーム効果、エコー効果の3経路を識別。「多くの世代が同時に生存する」ことで人口が増加するメカニズムを明示。日本への数値適用 | **概念的に最も近い先行研究**。「多世代の同時生存＝同時在生人口」という概念をほぼ同一の形で既に提示。ただし①学術的影響力の低いジャーナル（SCIRP）、②分析は日本のみ、③Gompertz生存やσ帯などの精緻化なし |

### 2.3 人口モメンタムと世代重複の古典理論

| 文献 | 内容 | 本アイデアとの関係 |
|------|------|-------------------|
| **Keyfitz (1971)** "On the Momentum of Population Growth" *Demography* 8(1):71-80 | 年齢構造に内在する人口成長の慣性（モメンタム）を定式化。出生率が即座に置換水準になっても人口は成長し続ける | 本アイデアが扱う「世代の重なりによる人口規模変動」の理論的基盤。ただしKeyfitzは出生率水準の変化に焦点、タイミングの変化は直接扱わない |
| **Ryder (1964, 1975)** Translation formula | コホートとピリオドの出生率を「翻訳」する公式。世代長と成長率の関係を定式化 | 古典的な理論的基盤。本アイデアのモデルはRyderの翻訳公式の延長線上にある |
| **Goldstein (2002)** "Population momentum for gradual demographic transitions" *Demography* | 漸進的な出生力転換下でのモメンタムを再定式化 | 補完的な理論。AFBの漸進的変化を扱う本アイデアと方法論的に関連 |
| **Feichtinger, Rau, & Novak (2025)** "On the momentum of pseudostable populations" *Demographic Research* | 古典的モメンタム理論を擬安定人口に拡張 | 最新のモメンタム理論の進展。本アイデアのモデルとの理論的接点あり |

### 2.4 出生タイミングの世界的トレンドとWPP予測

| 文献 | 内容 | 本アイデアとの関係 |
|------|------|-------------------|
| **Sobotka et al. (2026)** "A Concentration of Reproduction to Later Ages?" *Population and Development Review* | WPP 2024を用いた出生タイミングの世界的晩産化トレンドの包括的評価 | 本アイデアの背景的動機を裏付ける最新文献 |
| **GBD 2021 / Lancet (2024)** "Global fertility in 204 countries" | CCF50（50歳時点コホート完結出生率）に基づく予測。期間TFRではなくコホート指標を基盤に | Lancet掲載の出生力予測。WPPとは異なるアプローチだがテンポ効果は明示的に扱わない |
| **Rodriguez (2006)** "Demographic translation and tempo effects" *Demographic Research* | RyderとBFの枠組みを加速故障時間モデルで統合 | 理論的橋渡し。本アイデアの更新モデルの理論的基盤を提供 |

---

## 3. 現在位置の評価（Positioning）

### 3.1 知識空間における位置づけ

```
テンポ効果の概念（BF 1998）
    │
    ├── Rate measures（TFR補正）────── 成熟した研究領域
    │     Sobotka, Yoo, Kohler & Ortega, Luy, Mazzuco...
    │
    ├── Stock measures（人口規模への翻訳）
    │     │
    │     ├── Goldstein, Lutz, Scherbov (2003) ── EU・分析的分解
    │     ├── Lutz et al. (2003, Science) ── EU・person-years
    │     ├── Chen (2020) ── 日本・OLGモデル・世代重複
    │     │
    │     └── ★ 本アイデア ── 中国・DRC比較・AFBアンカー・
    │                          Gompertz生存・「同時在生人口」の政策フレーミング
    │
    └── WPP予測との接続 ── 本アイデアが主張する「テンポ不在が予測修正の一因」
```

### 3.2 先行研究との差分

| 差分の次元 | Goldstein et al. (2003) | Chen (2020) | **本アイデア** |
|-----------|------------------------|-------------|---------------|
| **対象地域** | EU15カ国 | 日本 | **中国 vs DRC**（対照的な人口転換段階） |
| **アンカー指標** | MAC（平均出産年齢） | 出産年齢d期 | **AFB（初産年齢）** |
| **モデル** | 安定人口＋数値シミュレーション | OLGモデル（離散） | **更新モデル＋Gompertz生存（連続）** |
| **出力指標** | 出生数比・人口比 | 総人口 | **「同時在生人口」の倍率** |
| **分散の扱い** | なし | なし | **σ narrow/baseline/wideの帯** |
| **政策フレーミング** | 出生政策の時間的影響 | 早産化政策 | **「同時在生人口」という政策指標の提案** |
| **ジャーナル** | PDR（トップジャーナル） | Modern Economy（低IF） | **Lancet（最高IF医学誌）** |

---

## 4. 新規性の評価

### 4.1 新規性が認められる要素

1. **「同時在生人口」（simultaneously living population）というフレーミング**
   - テンポ効果の「翻訳先」として人口ストック指標を明示的に命名・提案する試みは、先行研究では体系的に行われていない
   - 政策立案者が直感的に理解できる「今、同時に何人が生きているか」という指標は、TFRやNRRよりもアクセシブル
   - **評価：中程度の新規性**（概念自体は暗黙的に先行研究に存在するが、明示的なフレーミングとしては新しい）

2. **AFB（初産年齢）をアンカーに使用**
   - WPPがMACを使用し、BFもMACベースであるのに対し、AFBを政策的アンカーとして採用
   - AFBは初産の意思決定に直結し、政策介入（育児支援、不妊治療助成等）との接続が明確
   - **評価：小〜中程度の新規性**（技術的にはMACの特殊ケースだが、政策的には意味のある選択）

3. **中国 vs DRC の対照的比較**
   - 人口増加局面（DRC）と減少局面（中国）でテンポ効果の方向と規模が非対称であることを示す
   - 先行研究は主に欧州・日本・韓国に集中しており、サブサハラアフリカとの対照は少ない
   - **評価：中程度の新規性**（地理的・人口転換段階的な拡張として価値あり）

4. **σ（分散）の感度分析**
   - 出産年齢の分散を narrow/baseline/wide で変動させ、不確実性の幅を視覚化
   - Mazzuco & Zanotto (2025) がコホート視点でσの重要性を示したことと呼応
   - **評価：小程度の新規性**（感度分析の手法として妥当だが、方法論的革新ではない）

### 4.2 新規性に対する懸念

1. **Goldstein et al. (2003) との重複**
   - テンポ効果→人口規模という核心的ロジックはすでにPDRで発表済み
   - 本アイデアは分析的分解（テンポ×世代長×カンタム交互作用）を行っておらず、理論的深度で劣る
   - **対策案**：Goldstein et al.を明示的に引用し、本稿が「政策翻訳」としての貢献であることを強調する

2. **Chen (2020) との概念的近接性**
   - 「多世代の同時生存で人口が増加する」というメカニズムは概念的にほぼ同一
   - ただしChenの論文はSCIRPジャーナルで学術的影響力が限定的であり、引用もごく少数
   - **対策案**：Chenを引用しつつ、本稿が連続モデル・Gompertz生存・比較分析で精緻化していることを示す

3. **「同時在生人口」は基本的に「総人口」の言い換えではないか**
   - 人口学者にとっては、ある時点の在生人口数＝総人口であり、新しい概念とは認識されない可能性
   - **対策案**：テンポ効果のカウンターファクチュアル比較（AFBが変化した場合としない場合の倍率）という分析枠組みが本稿の核であることを明確にする

### 4.3 総合評価

| 評価項目 | 評価 |
|---------|------|
| **理論的新規性** | **低〜中** — 基礎的メカニズムは先行研究で確立済み |
| **方法論的新規性** | **中** — AFBアンカー＋Gompertz＋σ帯の組み合わせは新しい |
| **実証的新規性** | **中〜高** — 中国vsRDC比較は先行研究にない |
| **政策的新規性** | **中〜高** — 「同時在生人口」フレーミングは新しい政策翻訳 |
| **Lancet Correspondenceとしての適合性** | **中〜高** — 短い形式で政策的メッセージを伝えるには適切 |

---

## 5. 推奨事項

### 5.1 論文化に向けた強化ポイント

1. **先行文献との差異を明示化する**
   - Goldstein et al. (2003) と Chen (2020) を必ず引用し、本稿の付加価値（AFBアンカー、中国/DRC比較、政策フレーミング）を明記
   - 「テンポ効果はRate measuresとして研究されてきたが、Stock measuresとして政策立案者に報告する枠組みが不足している」という問題設定が効果的

2. **「同時在生人口」の定義を厳密にする**
   - 単なる「総人口」ではなく、「AFBカウンターファクチュアル比較による人口倍率」であることを定義で明確化
   - 例：「The simultaneously living population ratio (SLPR) is defined as the ratio of total living population under a given AFB scenario to that under a reference AFB scenario, at time t」

3. **WPPとの接続をデータで示す**
   - WPP 2022→2024の修正方向（中国下方、DRC基準年上方）とAFBトレンドの対応を図表で示せれば説得力が増す
   - ただしCorrespondenceの字数制限（400語、図1点）の中では難しいため、詳細版（Lancet Public Health）で展開

4. **政策含意を具体化する**
   - 「同時在生人口の差は、ケア需要（介護、教育、医療）の時間的配分に直結する」
   - 中国：晩産化→短期的在生人口増（1.08倍）→ケア需要の前倒し
   - DRC：早産化→短期的在生人口減（0.76倍）→ケア需要の後ろ倒し

### 5.2 投稿戦略

| 戦略 | 詳細 |
|------|------|
| **Lancet Correspondence** | 政策メッセージの速報として適切。ただし査読者がGoldstein et al.との差異を問う可能性が高い |
| **代替投稿先（リジェクト時）** | Lancet Public Health（詳細版）、Population and Development Review（テンポ議論の本拠地）、Demographic Research（オープンアクセス、テンポ研究の蓄積あり） |
| **差別化戦略** | 「人口学者向けの手法論文」ではなく「政策立案者向けの翻訳」として位置づける。Lancetの読者層（医師・公衆衛生専門家）にはテンポ効果自体が新しい情報 |

---

## 6. 結論

本アイデアは、テンポ効果研究の成熟した領域において、**「Rate→Stockの政策翻訳」**という比較的新しい角度からの貢献を目指している。核心的メカニズムはGoldstein et al. (2003) やChen (2020) で先行的に示されているが、以下の点で差別化が可能：

- **AFBをアンカーとした直感的な政策フレーミング**
- **中国 vs DRC の対照的比較（人口増加局面と減少局面の非対称性）**
- **σ帯による不確実性の視覚化**

Lancet Correspondenceとしては、**テンポ効果という人口学的概念を医学・公衆衛生の読者に「翻訳」する**という価値があり、投稿の合理性は認められる。ただし、**先行文献との差異を300語以内で明確に示す必要がある**ため、本文の冒頭で Goldstein et al. (2003) への言及を含めることを強く推奨する。

---

## 参考文献一覧（主要）

1. Bongaarts J, Feeney G. On the quantum and tempo of fertility. *Population and Development Review*. 1998;24(2):271-291.
2. Bongaarts J, Feeney G. The quantum and tempo of life-cycle events. *Vienna Yearbook of Population Research*. 2006;4:115-151.
3. Bongaarts J, Sobotka T. A demographic explanation for the recent rise in European fertility. *Population and Development Review*. 2012;38(1):83-120.
4. Chen IC. Population growth due to earlier childbearing. *Modern Economy*. 2020;11(11):1966-1975.
5. Feichtinger G, Rau R, Novak AJ. On the momentum of pseudostable populations. *Demographic Research*. 2025;52(15):445-478.
6. Goldstein JR. Population momentum for gradual demographic transitions. *Demography*. 2002;39(1):65-73.
7. Goldstein JR, Lutz W, Scherbov S. Long-term population decline in Europe: the relative importance of tempo effects and generational length. *Population and Development Review*. 2003;29(4):699-707.
8. Keyfitz N. On the momentum of population growth. *Demography*. 1971;8(1):71-80.
9. Kohler HP, Ortega JA. Tempo-adjusted period parity progression measures, fertility postponement and completed cohort fertility. *Demographic Research*. 2002;6(6):91-144.
10. Lutz W, O'Neill BC, Scherbov S. Europe's population at a turning point. *Science*. 2003;299(5615):1991-1992.
11. Luy M. Tempo effects and their relevance in demographic analysis. *Comparative Population Studies*. 2010;35(3):523-548.
12. Mazzuco S, Zanotto L. Tempo effects in period TFR: inspecting the role of shape and scale variations in a cohort model. *Demographic Research*. 2025;52(19):559-588.
13. Ní Bhrolcháin M. Tempo and the TFR. *Demography*. 2011;48(3):841-861.
14. Rodriguez G. Demographic translation and tempo effects: an accelerated failure time perspective. *Demographic Research*. 2006;14(6):85-110.
15. Ryder NB. Notes on stationary populations. *Population Index*. 1975;41(1):3-28.
16. Schoen R. Fertility quantum and tempo with cubic age-specific birth rates. *Demographic Research*. 2024;51(42):1351-1370.
17. Sobotka T. *Postponement of Childbearing and Low Fertility in Europe*. Amsterdam: Dutch University Press; 2004.
18. Yoo SH, Sobotka T. Ultra-low fertility in South Korea: the role of the tempo effect. *Demographic Research*. 2018;38(22):549-576.
19. GBD 2021 Fertility and Forecasting Collaborators. Global fertility in 204 countries and territories, 1950–2021, with forecasts to 2100. *The Lancet*. 2024;403(10440):2057-2099.
20. Sobotka T et al. A concentration of reproduction to later ages? A worldwide assessment of trends in fertility timing. *Population and Development Review*. 2026.

---

# English Translation

---

# “Reconsidering the tempo effect” paper idea evaluation report

## 1. Outline of the idea

| Item | Contents |
|------|------|
| **Title** | Tempo effects revisited: reporting simultaneously living population for policy relevance |
| **Submit to** | The Lancet (Correspondence, up to 400 words) |
| **Core Claim** | Translating the tempo effect (shift in age at first birth AFB) into a stock indicator called "simultaneously living population" to provide policy makers with intuitive numbers |
| **Model** | Endogenous update model + Gompertz survival, AFB=20/25/30, σ (variance) narrow/baseline/wide |
| **Target countries** | China (low birth rate, long life expectancy) vs Democratic Republic of Congo (DRC, high birth rate, short life expectancy) |
| **Main results** | China: Approximately 1.08 times after 100 years (difference ≈ 62 million people), DRC: Approximately 0.76 times after 100 years (difference ≈ 100 million people) |

---

## 2. Prior literature mapping

### 2.1 Basic theory of tempo effects (Rate measures)

| Literature | Contents | Relationship with this idea |
|------|------|-------------------|
| **Bongaarts & Feeney (1998)** "On the Quantum and Tempo of Fertility" *Population and Development Review* | Proposed tempo-adjusted TFR to remove tempo distortion from period TFR. Formulating the phenomenon in which shifts in birth timing distort period indicators | **Direct starting point**. However, BF only corrects the birth rate index and does not translate it into population stock (number of live population).
| **Bongaarts & Feeney (2006)** "The Quantum and Tempo of Life-Cycle Events" *Vienna Yearbook of Population Research* | Generalizes the concept of tempo distortion to all life cycle events such as birth, death, and marriage | Shows the generality of tempo effects, but still focuses on correction of flow (rate) indicators |
| **Sobotka (2004)** *Postponement of Childbearing and Low Fertility in Europe* | Comprehensive analysis of the relationship between childbearing postponement and lower TFR across Europe. Perform period-cohort comparison and tempo-quantum decomposition | Key literature showing the empirical importance of tempo effects. However, it does not deal with translation to population size |
| **Yoo & Sobotka (2018)** "Ultra-low fertility in South Korea" *Demographic Research* | Explaining South Korea's TFR<1.3 using the tempo effect. Applying tempo-parity adjusted TFR (TFRp*) | A precedent example that explains the extremely low birth rate in East Asia using tempo. This idea is close to the context of China analysis, but does not deal with population size |
| **Mazzuco & Zanotto (2025)** "Tempo effects in period TFR" *Demographic Research* | The latest method to decompose tempo distortion into variance/skewness changes from a cohort perspective | The latest tempo research. Partially in line with this idea of varying the dispersion parameter σ |

### 2.2 Connection between tempo effects and population size (Stock measures) — **Most competitive prior research group**
| Literature | Contents | Relationship with this idea |
|------|------|-------------------|
| **Goldstein, Lutz, & Scherbov (2003)** "Long-Term Population Decline in Europe: The Relative Importance of Tempo Effects and Generational Length" *Population and Development Review* 29(4):699-707 | **Most important prior research that directly connects tempo effects to population size**. The impact of changes in birth timing on the population is broken down into three factors: (a) tempo effect (temporary), (b) generation length effect (effect on long-term growth rate), and (c) tempo-quantum interaction. Numerical simulations were conducted in 15 EU countries. Conclusion: Tempo effect overwhelms generation length effect for 200 years | **Biggest competing literature**. We have already implemented the core transition of this idea from tempo to population size. However, (1) the indicator is population size (ratio of births) and is not explicitly framed as "co-resident population", (2) it targets only the EU, and (3) the anchor is MAC rather than AFB.
| **Lutz, O'Neill, & Scherbov (2003)** "Europe's Population at a Turning Point" *Science* 299:1991-1992 | Demonstrates Europe's negative population momentum and quantifies how continued tempo effects accelerate population decline and aging. Expressing policy implications as **additional working-age population x person-years** | A pioneering paper that emphasizes the "policy relevance of tempo effects." We are proposing a policy indicator called person-years, which is an attempt at policy translation similar to the "co-resident population" of this idea.
| **Chen, I-Chun (2020)** "Population Growth Due to Earlier Childbearing" *Modern Economy* 11:1966-1975 | **Directly modeling the impact of earlier childbearing on the total population using the OLG model**. Three pathways were identified: tempo effect, baby boom effect, and echo effect. Clarifying the mechanism by which the population increases due to "many generations living at the same time". Numerical application to Japan | **Previous research that is conceptually closest**. The concept of "simultaneous existence of multiple generations = simultaneous living population" has already been presented in almost the same form. However, (1) a journal with low academic influence (SCIRP), (2) analysis only in Japan, and (3) no elaboration of Gompertz survival or σ band.

### 2.3 Classical theory of population momentum and overlapping generations

| Literature | Contents | Relationship with this idea |
|------|------|-------------------|
| **Keyfitz (1971)** "On the Momentum of Population Growth" *Demography* 8(1):71-80 | Formulates the inertia (momentum) of population growth inherent in the age structure. Even if the birth rate immediately reaches the replacement level, the population will continue to grow | This idea deals with the theoretical basis of ``changes in population size due to overlapping generations.'' However, Keyfitz focuses on changes in fertility levels and does not directly deal with changes in timing |
| **Ryder (1964, 1975)** Translation formula | Formula for "translating" cohort and period birth rates. Formulating the relationship between generation length and growth rate | Classical theoretical basis. The model of this idea is an extension of Ryder's translation formula |
| **Goldstein (2002)** "Population momentum for gradual demographic transitions" *Demography* | Reformulating momentum under gradual fertility transitions | Complementary theory. Methodologically related to this idea of ​​dealing with gradual changes in AFB |
| **Feichtinger, Rau, & Novak (2025)** "On the momentum of pseudostable populations" *Demographic Research* | Extending classical momentum theory to pseudostable populations | Latest advances in momentum theory. There is a theoretical connection with the model of this idea |

### 2.4 Global trends in birth timing and WPP predictions

| Literature | Contents | Relationship with this idea |
|------|------|-------------------|
| **Sobotka et al. (2026)** "A Concentration of Reproduction to Later Ages?" *Population and Development Review* | Comprehensive assessment of the global trend toward later birth timing using WPP 2024 | Latest literature supporting the motivation behind this idea |
| **GBD 2021 / Lancet (2024)** "Global fertility in 204 countries" | Forecast based on CCF50 (cohort completed birth rate at age 50). Based on cohort indicators rather than period TFR | Fertility prediction published in Lancet. A different approach than WPP, but it does not explicitly deal with tempo effects |
| **Rodriguez (2006)** "Demographic translation and tempo effects" *Demographic Research* | Integrating the Ryder and BF frameworks with an accelerated failure time model | Theoretical bridge. This idea provides a theoretical basis for the renewal model |

---

## 3. Evaluation of current position (Positioning)

### 3.1 Positioning in the knowledge space

````
The concept of tempo effects (BF 1998)
    │
    ├── Rate measures (TFR correction) ────── Mature research area
    │ Sobotka, Yoo, Kohler & Ortega, Luy, Mazzuco...
    │
    ├── Stock measures (translation to population size)
    │ │
│ ├── Goldstein, Lutz, Scherbov (2003) ── EU/analytical decomposition
    │ ├── Lutz et al. (2003, Science) ── EU・person-years
    │ ├── Chen (2020) ── Japan, OLG model, generation overlap
    │ │
    │ └── ★ This idea ── China/DRC comparison/AFB anchor/
    │ Gompertz survival/policy framing of “co-resident population”
    │
    └── Connection with WPP prediction ─ This idea claims that “absence of tempo is a factor in forecast correction”
````

### 3.2 Differences from previous research

| Difference dimension | Goldstein et al. (2003) | Chen (2020) | **Book idea** |
|-----------|------------------------|-------------|---------------|
| **Target area** | 15 EU countries | Japan | **China vs DRC** (contrasting demographic transition stages) |
| **Anchor indicator** | MAC (average age at birth) | Age at birth d | **AFB (age at first birth)** |
| **Model** | Stable population + numerical simulation | OLG model (discrete) | **Update model + Gompertz survival (continuous)** |
| **Output indicators** | Birth rate/population ratio | Total population | **“Simultaneous population” multiplier** |
| **Dispersion treatment** | None | None | **σ narrow/baseline/wide band** |
| **Policy framing** | Temporal effects of fertility policies | Preterm birth policies | **Proposal for a policy indicator called "co-resident population"** |
| **Journal** | PDR (Top Journal) | Modern Economy (Low IF) | **Lancet (Highest IF Medical Journal)** |

---

## 4. Evaluation of novelty

### 4.1 Elements that are recognized as novel

1. **Framing “simultaneously living population”**
   - Previous studies have not systematically attempted to explicitly name or propose population stock indicators as the "translation target" of the tempo effect.
   - An indicator of “how many people are alive at the same time” that policy makers can intuitively understand is more accessible than TFR or NRR.
   - **Rating: Moderate Novelty** (The concept itself implicitly exists in previous studies, but the explicit framing is new)
2. **AFB (age at first birth) used as anchor**
   - WPP uses MAC and BF is also MAC-based, while AFB is adopted as a policy anchor
   - AFB is directly linked to decision-making for first births and has clear connections to policy interventions (childcare support, infertility treatment subsidies, etc.)
   - **Rating: Small to Moderate Novelty** (Technically a special case of MAC, but a meaningful choice from a policy perspective)

3. **Contrastive comparison of China vs DRC**
   - Shows that the direction and magnitude of the tempo effect are asymmetric between periods of population growth (DRC) and population decline (China)
   - Previous studies have mainly focused on Europe, Japan, and South Korea, and there is little comparison with sub-Saharan Africa.
   - **Rating: Moderate Novelty** (Valuable as a step-by-step expansion with geographic and demographic transitions)

4. **Sensitivity analysis of σ (variance)**
   - Visualize the range of uncertainty by varying the variance of birth age by narrow/baseline/wide
   - In line with Mazzuco & Zanotto (2025) showing the importance of σ from a cohort perspective.
   - **Evaluation: Minor novelty** (valid as a method for sensitivity analysis, but not methodological innovation)

### 4.2 Concerns about novelty

1. **Overlap with Goldstein et al. (2003)**
- The core logic of tempo effect → population size has already been announced in PDR
   - This idea does not perform analytical decomposition (tempo x generation length x quantum interaction) and is inferior in theoretical depth.
   - **Countermeasure**: Explicitly cite Goldstein et al. and emphasize that this paper is a contribution as "policy translation"

2. **Conceptual proximity with Chen (2020)**
   - The mechanism of "population increasing due to simultaneous survival of multiple generations" is conceptually almost the same.
   - However, Chen's paper has limited academic influence in SCIRP journals and has only a few citations.
   - **Countermeasure**: Citing Chen, this paper shows that it has been refined with continuous models, Gompertz survival, and comparative analysis.

3. **Isn't "co-resident population" basically another way of saying "total population"?**
   - For demographers, the number of living people at a certain point in time = total population, and it may not be recognized as a new concept.
   - **Proposed countermeasure**: Clarify that the analytical framework of counterfactual comparison of tempo effects (multipliers when AFB changes and when it does not) is the core of this paper.

### 4.3 Overall evaluation

| Evaluation items | Evaluation |
|---------|------|
| **Theoretical novelty** | **Low to moderate** — The basic mechanism has been established in previous studies |
| **Methodological novelty** | **Medium** — The combination of AFB anchor + Gompertz + σ band is new |
| **Empirical novelty** | **Medium to high** — China vs. RDC comparison is not found in previous studies |
| **Policy Novelty** | **Medium to High** — “Co-resident population” framing is a new policy translation |
| **Suitability as a Lancet Correspondence** | **Medium to High** — Appropriate for conveying policy messages in short form |

---

## 5. Recommendations

### 5.1 Strengthening points for publication in papers

1. **Clarify the differences from previous documents**
   - Be sure to cite Goldstein et al. (2003) and Chen (2020) and clearly state the added value of this paper (AFB anchor, China/DRC comparison, policy framing)
   - An effective problem setting is that ``tempo effects have been studied as rate measures, but there is a lack of a framework for reporting them to policy makers as stock measures.''

2. **Stricter definition of “co-resident population”**
   - Clarified in the definition that it is not just "total population" but "population multiplier based on AFB counterfactual comparison"
- Example: "The simultaneously living population ratio (SLPR) is defined as the ratio of total living population under a given AFB scenario to that under a reference AFB scenario, at time t"

3. **Indicate connection with WPP with data**
   - It will be more persuasive if you can show the correspondence between the WPP 2022 → 2024 revision direction (downward in China, upward in the DRC base year) and the AFB trend.
   - However, it is difficult to meet the character limit of Correspondence (400 words, 1 figure), so it will be published in a detailed version (Lancet Public Health)

4. **Making policy implications concrete**
   - "Differences in the number of people living at the same time are directly linked to the temporal allocation of care demands (nursing care, education, medical care)"
   - China: Late birth → short-term population increase (1.08 times) → front-loading demand for care
   - DRC: Preterm birth → Short-term population decline (0.76 times) → Postponed demand for care

### 5.2 Posting Strategy

| Strategy | Details |
|------|------|
| **Lancet Correspondence** | Appropriate for breaking policy messages. However, there is a high possibility that reviewers will ask about the differences with Goldstein et al. |
| **Alternative submission destination (in case of rejection)** | Lancet Public Health (detailed version), Population and Development Review (home of tempo discussion), Demographic Research (open access, with accumulation of tempo research) |
| **Differentiation Strategy** | Position it as a ``translation for policy makers'' rather than a ``method paper for demographers.'' The tempo effect itself is new information for Lancet readers (physicians and public health experts) |

---

## 6. Conclusion

This idea aims to contribute to the mature field of tempo effect research from a relatively new angle: ``Rate→Stock policy translation.'' The core mechanism was previously demonstrated by Goldstein et al. (2003) and Chen (2020), but it can be differentiated in the following points:

- **Intuitive policy framing with AFB as an anchor**
- **Contrastive comparison of China vs. DRC (asymmetry between population growth and decline phases)**
- **Visualization of uncertainty by σ band**
As a Lancet Correspondence, there is value in ``translating'' the demographic concept of tempo effect to medical and public health readers, and the rationality of the submission is acknowledged. However, **Differences from previous literature must be clearly stated within 300 words**, so we strongly recommend including a reference to Goldstein et al. (2003) at the beginning of the text.

---

## List of references (main)

1. Bongaarts J, Feeney G. On the quantum and tempo of fertility. *Population and Development Review*. 1998;24(2):271-291.
2. Bongaarts J, Feeney G. The quantum and tempo of life-cycle events. *Vienna Yearbook of Population Research*. 2006;4:115-151.
3. Bongaarts J, Sobotka T. A demographic explanation for the recent rise in European fertility. *Population and Development Review*. 2012;38(1):83-120.
4. Chen IC. Population growth due to earlier childbearing. *Modern Economy*. 2020;11(11):1966-1975.
5. Feichtinger G, Rau R, Novak AJ. On the momentum of pseudostable populations. *Demographic Research*. 2025;52(15):445-478.
6. Goldstein JR. Population momentum for gradual demographic transitions. *Demography*. 2002;39(1):65-73.
7. Goldstein JR, Lutz W, Scherbov S. Long-term population decline in Europe: the relative importance of tempo effects and generational length. *Population and Development Review*. 2003;29(4):699-707.
8. Keyfitz N. On the momentum of population growth. *Demography*. 1971;8(1):71-80.
9. Kohler HP, Ortega JA. Tempo-adjusted period parity progression measures, fertility postponement and completed cohort fertility. *Demographic Research*. 2002;6(6):91-144.
10. Lutz W, O'Neill BC, Scherbov S. Europe's population at a turning point. *Science*. 2003;299(5615):1991-1992.
11. Luy M. Tempo effects and their relevance in demographic analysis. *Comparative Population Studies*. 2010;35(3):523-548.
12. Mazzuco S, Zanotto L. Tempo effects in period TFR: inspecting the role of shape and scale variations in a cohort model. *Demographic Research*. 2025;52(19):559-588.
13. Ní Bhrolcháin M. Tempo and the TFR. *Demography*. 2011;48(3):841-861.
14. Rodriguez G. Demographic translation and tempo effects: an accelerated failure time perspective. *Demographic Research*. 2006;14(6):85-110.
15. Ryder NB. Notes on stationary populations. *Population Index*. 1975;41(1):3-28.
16. Schoen R. Fertility quantum and tempo with cubic age-specific birth rates. *Demographic Research*. 2024;51(42):1351-1370.
17. Sobotka T. *Postponement of Childbearing and Low Fertility in Europe*. Amsterdam: Dutch University Press; 2004.
18. Yoo SH, Sobotka T. Ultra-low fertility in South Korea: the role of the tempo effect. *Demographic Research*. 2018;38(22):549-576.
19. GBD 2021 Fertility and Forecasting Collaborators. Global fertility in 204 countries and territories, 1950–2021, with forecasts to 2100. *The Lancet*. 2024;403(10440):2057-2099.
20. Sobotka T et al. A concentration of reproduction to later ages? A worldwide assessment of trends in fertility timing. *Population and Development Review*. 2026.


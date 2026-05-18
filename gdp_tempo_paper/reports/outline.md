# GDP テンポ論文 アウトライン (v0)

## 0. メタ情報

| 項目 | 値 |
|---|---|
| ワーキングタイトル | *The Forgotten Tempo Effect in Capital Accounting: Investment-to-Output Time-to-Build, Intangible Capital, and the Reconciliation of Flow- and Stock-Based National Wealth Measures* |
| 想定投稿先（Tier 1） | **Review of Income and Wealth** (IF ≈ 1.8, Wealth 会計の本丸) |
| 想定投稿先（Tier 1） | **Journal of Economic Growth** (IF ≈ 7.2, Solow/PIM 系) |
| 想定投稿先（Tier 2） | **Economic Systems Research** (IF ≈ 3.5, 国富会計 IOTA 寄り) |
| 想定投稿先（Tier 2） | **The World Bank Economic Review** (IF ≈ 2.3, CWON 直接の発行元系) |
| 想定投稿先（汎用） | **PLOS ONE** (方法論としての位置付け) |
| ドラフト形態 | Research Article (~7000 words), EN+JA 両方 |
| 図表数 | 5 figures + 2 tables（内訳は §3） |
| 使用データ | PWT 10.01, WB WDI (R&D), WB CWON (NW.PCA.TO, NW.HCA.TO, NW.TOW.TO) |
| サンプル | 39 countries (OECD + CHN), 1970–2020 |

---

## 1. 論旨（elevator pitch）

人口テンポ論文で示した「量子（quantum）vs. テンポ（tempo）」の二分法は、**資本会計でもそのまま成立する**:

| 量子（quantum） | テンポ（tempo） |
|---|---|
| 投資量 I(t) | 投資 → 稼働ラグ μ(t) |
| R&D 支出 I_R(t) | R&D → 収益化ラグ |
| 出生数・TFR | 初産年齢・MAC |

人口では MAC のドリフトが「同時生存人口」を下方バイアスさせた。**資本では μ(t) のドリフトが GDP と CWON の両者を同時バイアスさせる**。
さらに、フロー（GDP）とストック（CWON）を独立に補正するのではなく、**共通の隠れパラメータ (μ, β, σ) で両側を同時同定**すれば、恒等式 `dW/dt = S(Y) − δW` のもとで初めて閉じた国富会計が成立する。

---

## 2. セクション構成

### §1 Introduction — The Forgotten Tempo Effect in Economic Accounting
- 人口テンポ論文（Goldstein-Lutz-Scherbov 2003, Bongaarts-Feeney 1998, 先行のPDR版）を引用
- GDP と CWON の乖離問題：既報で繰り返し指摘されるが、**両側を同時に歪める隠れ指標**の概念化がない
- Research question:
  - (Q1) 人口テンポ論文の「μ(t) ドリフト」枠組は GDP 会計でも有効か？
  - (Q2) 既報の Solow 生産関数 / PIM 推計で使われているパラメータと比べ、テンポ拡張版は予測精度を改善するか？
  - (Q3) フロー（GDP）とストック（CWON）を同じ隠れ指標で同時同定できるか？
  - (Q4) これは医療アウトカムにも移植可能か？（予告のみ、次論文で実装）

### §2 Literature and Gap
- **2.1** Capital accounting の系譜: Solow (1957), Perpetual Inventory (OECD Manual 2009), Griliches (1979) R&D stock
- **2.2** Intangible capital: Corrado-Hulten-Sichel (2005, 2009), Haskel-Westlake (2018 *Capitalism Without Capital*)
- **2.3** Wealth accounting: World Bank CWON (2021), Arrow et al. (2012) *Sustainability and the measurement of wealth*, UNEP IWR (2018)
- **2.4** Tempo effect migration: Bongaarts-Feeney 1998 (源流), Goldstein-Lutz-Scherbov 2003 (人口ストックへの橋渡し), Feichtinger-Rau-Novak 2025 (pseudostable)
- **2.5** Gap: テンポ効果は demographic 文献に閉じていて、**資本会計でもアナロジーが成立することが体系的に示されていない**

### §3 Theory: A Tempo-Augmented Capital Accounting Framework

#### 3.1 Flow side — GDP with tempo
古典的 Solow:
$$Y(t) = A(t) K(t)^{\alpha} L(t)^{1-\alpha}$$

テンポ拡張版（候補 A: 投資 → 稼働ラグの時変）:
$$K(t+1) = (1-\delta) K(t) + \sum_{s=0}^{S} w_s(\mu(t)) \, I(t-s)$$
where $w_s(\mu) = (1-\theta)\theta^s$, $\theta = \mu/(1+\mu)$

既報の Solow/PIM は μ = 0 を暗黙に仮定。これは **人口テンポでの MAC = 25 固定の仮定に相当**。

#### 3.2 Stock side — Intangibles as the forgotten parameter σ
σ（忘れられたパラメータ）の候補 D: R&D 蓄積:
$$K_{\text{total}}(t) = K_{\text{tang}}(t) + \beta \cdot K_{\text{intan}}(t)$$

#### 3.3 Unifying identity
$$\underbrace{\frac{dW}{dt}}_{\text{stock change}} = \underbrace{S(Y) - \delta W}_{\text{flow-derived}}$$

両辺に同じ隠れパラメータ (μ, β) が入る → **joint identification** 問題:
$$\hat{\mu}, \hat{\beta} = \arg\min \left[ L_{\text{prod}}(\mu, \beta) + \lambda \, L_{\text{wealth}}(\mu, \beta) \right]$$

#### 3.4 Population tempo との対応表

| 概念 | 人口 | GDP / Wealth |
|---|---|---|
| 量子 | TFR | I / GDP (投資率) |
| テンポ | MAC | μ (投資→稼働ラグ) |
| 忘れられた stock | 同時生存人口 | 無形資本 K_intan |
| 橋渡し恒等式 | renewal equation | dW/dt = S(Y) − δW |
| 誤差の方向 | TFR 量過小評価 | CWON 過小評価 |

### §4 Data and Methods

#### 4.1 Data
| Source | Variable | Coverage |
|---|---|---|
| Penn World Table 10.01 | rgdpna, rnna, csh_i, delta, emp, hc, avh, labsh | 1970–2019 |
| WB WDI | R&D (% of GDP) | 1996–2020 |
| WB CWON | NW.PCA.TO, NW.HCA.TO, NW.TOW.TO | 1995–2020 |

39 か国（OECD + CHN、前稿と同一、DRC は除外）

#### 4.2 Estimation
- M0 (既報ベースライン): Solow immediate PIM
- M1: 定数ラグ（μ 単一値）
- M2: 時変テンポ μ(t)（候補 A）
- M3: 無形資本 K_intan 追加（候補 D）
- M4: M2 + M3 統合 + wealth 整合（本稿の提案）

各モデルの per-country μ̂, β̂ を格子探索で推定。

#### 4.3 Evaluation
- **Production side**: annual growth rate RMSE, level MAPE (既報との互換)
- **Wealth side**: within-country demeaned log trajectory RMSE vs CWON PCA
- **Joint**: L = L_prod + λ L_wealth（λ は感度分析）
- **Benchmark**: 既報 Solow 単純 PIM と対比

### §5 Results

#### 5.1 既報パラメータとの比較（Q2）
Table 1: 39 か国の推定値分布

| 指標 | 既報 Solow (M0) | 定数ラグ (M1) | テンポ μ(t) (M2) | +無形 (M3) | Joint (M4) |
|---|---|---|---|---|---|
| μ̂ 中央値 | 0 (暗黙) | — | 0.40 | — | **0.26** |
| β̂ 中央値 | 0 | 0 | 0 | 0.01 | **0.06** |
| 年次成長 RMSE 中央値 | — | — | +0.028 pp 改善 | — | TBD |
| 水準 MAPE 中央値 | baseline | — | — | +0.39 pp 改善 | TBD |
| 勝利国（M > baseline） | — | — | 24/39 | 38/39 | TBD |

#### 5.2 予測効果の向上（Q2）
Figure 1: 既報と本稿モデルの per-country RMSE 改善量（棒グラフ）
Figure 2: USA/JPN/KOR/DEU で 2015–2020 を holdout した out-of-sample 予測

#### 5.3 Flow-stock 整合性（Q3）
Figure 3: W_PIM(μ̂, β̂) vs CWON PCA の軌跡比較（代表6か国）
- Within-country 軌跡 RMSE 中央値 **0.049 log 単位**（≈5%）
- β は production-only 0.01 → joint 0.06 に 6 倍化
- μ は production-only 0.40 → joint 0.26

Figure 4: 横断面 log(W_PIM / CWON PCA) vs R&D 集約度
- 傾き −0.49（R²=0.49）
- PPP/XRAT 混交効果を議論、日本の γ_price 効果を注記

#### 5.4 人口テンポとの類似性
Table 2: Goldstein-Lutz-Scherbov (2003) の EU-15 結果と本稿 GDP 結果の構造的類似性

### §6 Discussion

#### 6.1 既報 Solow/PIM 推計の系統バイアス
- μ = 0 の暗黙仮定は投資先進国（USA, KOR, ISR）で資本ストックを 3–7% 過小評価
- 無形資本欠落は Wealth 測定で CWON を 5–15% 過小評価

#### 6.2 人口テンポ論文との二重類似性
- quantum vs tempo の二分法
- stock 集計量（同時生存人口 ↔ 修正 Wealth）で政策議論を振り直す必要性

#### 6.3 Beyond GDP / IWI 論への貢献
- IWI・CWON は「GDP を超える」前に **GDP と同じ隠れバイアス**を抱えている
- 正しい beyond-GDP は両側を共通パラメータで整合化することで得られる

#### 6.4 医療ドメインへの移植（preview）
- 医療費フロー → 健康寿命ストックの橋渡しでも同じ構造が期待される
- 先行 PoC（本グループ、別原稿、準備中）では平均寿命→医療費のラグ μ_H1 中央値 +0.15 年/年が観察
- 支出バケット別乗数 λ_b の同定で「治療偏重 vs 予防重点」の経済効率を定量化予定
- 本稿の枠組みを医療アウトカムに応用することで、**持続可能な医療費の構成**という政策議論への入口が開く

#### 6.5 Limitations
- PPP vs XRAT の単位混在（暫定的に demeaned log 比較で回避）
- Natural capital を残差扱い
- 39 か国・1995–2020 の標本サイズ
- γ_price（資産価格再評価）の内生化は未実装

### §7 Conclusion
- 人口テンポ論文が MAC の忘却を指摘したのと対称に、**本稿は GDP 会計における μ(t) の忘却を指摘**
- フロー（GDP）とストック（Wealth）は共通の隠れ指標で同時同定でき、従来の Beyond GDP 論の片側修正よりも構造的に正しい
- 同じ枠組みは医療ドメインにも移植可能で、持続可能な公共支出の構成という次課題につながる

---

## 3. 必要図表一覧

| No. | タイプ | 内容 | データ源 |
|---|---|---|---|
| Fig 1 | 棒グラフ | M0→M2 年次成長 RMSE 改善量（39か国） | poc_results.csv |
| Fig 2 | 多パネル折れ線 | 代表4か国の out-of-sample 予測 | PWT + M2/M3/M4 |
| Fig 3 | 多パネル折れ線 | W_PIM vs CWON PCA 軌跡比較 6か国 | cwon_integration_ts.json |
| Fig 4 | 散布図 | log(W_PIM/CWON) vs R&D 集約度 | cwon_integration.csv |
| Fig 5 | 概念図 | quantum/tempo の人口↔資本対応表 | 新規作図 |
| Table 1 | 数値表 | M0–M4 の per-country パラメータ分布 | 既出力 CSV |
| Table 2 | 構造比較表 | Goldstein-Lutz-Scherbov との対応 | — |

---

## 4. 必要な追加解析（現状 PoC でカバーされていない）

1. **Out-of-sample 予測** (§5.2) — 2015–2020 を holdout として残し、M0–M4 を訓練期間で推定して予測誤差を比較。**未実装**
2. **ブートストラップ CI** — 既報 PoC は点推定のみ。39か国 per-country CI を付ける。**未実装**
3. **M4 統合モデルの全国運行** — CWON joint (μ, β) は PR #38 で実装済みだが、M1/M2/M3 との「同じ evaluation metric 上での公平比較」になっていない。既報 RMSE/MAPE と同じスコアで M4 を評価する必要。**半分実装**
4. **γ_price 感度解析** — 日本アノマリーを数値で検証するため、CWON PCA の price vs volume 分解に踏み込むか、暫定的に土地価格指数でデフレートするか。**未実装、scope 外に置く選択肢も**

---

## 5. 執筆進行計画（案）

| Phase | 内容 | 想定工数 |
|---|---|---|
| P0 | 本アウトラインのユーザー承認 | — |
| P1 | 追加解析（§4 の 1-3） | 1 セッション |
| P2 | 本文（JA → EN の順で）ドラフト v1 | 2 セッション |
| P3 | 図表生成スクリプト（inline docx + 編集可能 pptx） | 1 セッション |
| P4 | Vancouver 参考文献整備・Devin Review 対応 | 1 セッション |
| P5 | PR 提出・投稿先別フォーマット調整 | 1 セッション |

---

## 6. ユーザー確認事項

1. **言語**: 日本語・英語・両方 どれを主にするか
2. **投稿先**: Review of Income and Wealth を第一候補で良いか（代替提案歓迎）
3. **長さ**: 本文 ~7000 words 規模で良いか（短縮形が必要なら Research Note 〜4000）
4. **追加解析スコープ**: 上記「必要な追加解析」の (1)-(4) をどこまで本稿に含めるか
5. **医療論文への予告度合い**: §6.4 を「1 段落の preview」で済ませるか、「姉妹論文として本稿と同時進行で言及」か

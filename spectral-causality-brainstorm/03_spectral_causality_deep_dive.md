# スペクトル因果性（Spectral Causality）深掘り

## 0. 位置づけ

前回のブレストで「スペクトル因果同定」を仮説として提示した。本稿ではこの仮説を、既存の数学的・理論的基盤と接続しながら深掘りする。

**核心的問い**: 有向グラフのスペクトル構造（固有値・固有ベクトル）から**因果的方向性**を読み取ることは可能か？ もし可能なら、LiNGAMの「非ガウス性」とは異なる、第二の因果同定原理となりうる。

---

## 1. 理論的基盤: 有向グラフのスペクトル理論

### 1-1. 無向グラフ vs 有向グラフのラプラシアン

**無向グラフラプラシアン L = D - W**:
- 対称行列 → 実固有値のみ
- 固有値はすべて非負（半正定値）
- 第2最小固有値（Fiedler値）= グラフ連結度の指標
- **方向性情報は失われている**

**有向グラフラプラシアン L_d = D_out - W**:
- 一般に**非対称**行列
- 固有値は**複素数**になりうる（Yu et al. 2025が詳細に分析）
- 複素固有値の出現条件 = **有向サイクル（directed cycles）の存在**
- つまり: 有向サイクルがない（= DAG）なら実固有値のみ、サイクルがあれば複素固有値

**重要な定理（Yu et al. 2025）**: 
> 有向グラフのラプラシアンが純実スペクトルを持つ十分条件は、「digon sign-asymmetric interaction がない」かつ「強連結な部分グラフがない」こと。逆に、有向サイクルの存在は複素固有値の主要因である。

### 1-2. 磁気ラプラシアン（Magnetic Laplacian）— 方向性のエレガントな符号化

de Resende & da Costa (2020), Zhang et al. (2021) が提案した磁気ラプラシアン:

```
L^(q) = I - D^{-1/2} H^(q) D^{-1/2}

H^(q)_{ij} = W_{ij} · exp(i · 2πq · (1_{i→j} - 1_{j→i}))
```

- q: 電荷パラメータ（0〜0.5）
- `1_{i→j}`: エッジ i→j が存在するとき1
- **エッジの方向性を複素位相として符号化**
- エルミート行列 → **実固有値を保証しつつ、固有ベクトルは複素数**

**物理的直感**: 電子が磁場中のグラフ上を移動するとき、ループを一周すると位相が回転する。この位相回転が方向性の情報を運ぶ。

**医療データへの含意**: 臨床変数間の有向関係（原因→結果）を複素位相で符号化し、固有分解すれば、方向性を保ったまま潜在テーマを抽出できる。

### 1-3. DAG上の因果フーリエ解析（Seifert, Wendler & Püschel 2023）

ETHのPüschelグループが提案した画期的フレームワーク:

**核心**: DAG上の信号に対するフーリエ変換を定義。これは:
- DAG上の**シフト演算子・畳み込み**の固有分解として定義される
- **因果構造（DAGの推移閉包）に基づくフーリエ基底**
- 古典的Möbius反転の拡張
- **「因果的フーリエスパース性」= 原因が少数**の仮定下で信号復元

```
DAG上のフーリエ変換:
  f̂(v) = Σ_{u ≤ v} μ(u,v) · f(u)

  μ: DAGの推移閉包上のMöbius関数
  u ≤ v: DAG上でuからvへの有向パスが存在
```

**時系列への拡張（Misiakos, Mihal & Püschel 2024）**:
- 時系列グラフデータ = グラフを時間方向にunrolしたDAG
- 「原因が少数」仮定 = フーリエスパース性
- DAG上の信号復元問題として因果推論を定式化

### 1-4. Hodge分解 — 因果フローの直交分解

Jiang et al. (2011) のHodgeRankから着想を得た情報フロー分解:

**グラフ上のHodge分解**:
```
任意のエッジフロー ω = ∇φ + δψ + h

  ∇φ: 勾配フロー（gradient flow）= DAG的・因果的フロー
  δψ: カールフロー（curl flow）= 局所的循環
  h:   調和フロー（harmonic flow）= 大域的循環
```

**因果性との接続**:
- **勾配成分 ∇φ**: エッジフローのうち「ポテンシャル差に駆動される部分」= **因果的な情報フロー**
- **カール成分 δψ**: 局所的なフィードバックループ
- **調和成分 h**: 大域的なサイクル構造

Sugihara et al. (2016) "Hodge Decomposition of Information Flow on Small-World Networks" で脳ネットワークの情報フローにHodge分解を適用した先行研究あり。

**Maehara & Ohkawa (2019)** は単一細胞RNAseqデータにHodge分解を適用し、細胞分化の「方向的フロー」を抽出。医療データへの直接的先行事例。

---

## 2. スペクトル因果性の数学的定式化

### 2-1. 定義体系

以下、段階的に定義を構築する。

#### 定義1: ユーティリティ有向グラフ（Utility Digraph）

```
G_U = (V, E, w, d)

V: ノード集合（横断: 変数 or 患者、経時: (変数, 時点) ペア）
E: エッジ集合
w: E → R+（ユーティリティ類似度に基づく重み）
d: E → {-1, 0, +1}（方向性符号）

方向性の付与:
  横断データ: ユーティリティ非対称性
    d(i,j) = sign(Util(i→j) - Util(j→i))
    Util(i→j) = 「iのデータがjに関する問いにどれだけ役立つか」
    
  経時データ: 時間的方向
    d(i^t, j^{t'}) = sign(t' - t)（tが先ならi→j方向）
```

#### 定義2: ユーティリティ磁気ラプラシアン

```
L^(q)_U = I - D^{-1/2} H^(q)_U D^{-1/2}

H^(q)_{U,ij} = w(i,j) · exp(i · 2πq · d(i,j))

q: 方向性感度パラメータ
  q = 0: 方向性を無視（通常のラプラシアン）
  q = 0.25: 最大方向性感度
  q = 0.5: 方向を反転
```

**利点**: エルミート行列なので実固有値を持ち、固有ベクトルの複素位相に方向性情報が符号化される。

#### 定義3: スペクトル因果結合度（Spectral Causal Coupling）

```
SCC(i, j) = Σ_k f(λ_k) · |⟨u_k, e_i⟩| · |⟨u_k, e_j⟩| · cos(arg(⟨u_k, e_i⟩) - arg(⟨u_k, e_j⟩))

  u_k: k番目の固有ベクトル（複素数値）
  λ_k: 対応する固有値（実数、磁気ラプラシアンの場合）
  e_i: ノードiの標準基底ベクトル
  f(λ_k): 固有値重み関数（例: f(λ) = λ で高周波=局所構造を重視）
  arg(·): 複素数の偏角
```

**直感的解釈**:
- `|⟨u_k, e_i⟩| · |⟨u_k, e_j⟩|`: iとjが同じEigenthemeに強く荷重
- `cos(arg差)`: **対称関数** — cos(a-b) = cos(b-a) なので、SCC(i,j) = SCC(j,i)
- → SCCは因果の**強度（方向なし）**を測る。方向は定義4のSCDで測る

#### 定義4: スペクトル因果方向（Spectral Causal Direction）

SCDはsinを用いて直接定義する。sin(a-b) = -sin(b-a) なので**反対称性**を持ち、因果方向を符号化できる:

```
SCD(i, j) = Σ_k f(λ_k) · |⟨u_k, e_i⟩| · |⟨u_k, e_j⟩| · sin(arg(⟨u_k, e_i⟩) - arg(⟨u_k, e_j⟩))

  SCD > 0: iがjの（スペクトル的な）原因
  SCD < 0: jがiの（スペクトル的な）原因
  SCD ≈ 0: 因果方向不確定

性質: SCD(i, j) = -SCD(j, i)（反対称性）
```

**SCCとSCDの関係**: SCCは因果カップリングの強さ（方向不問）、SCDは因果の方向を定量化する。両者は同じ固有分解から得られるが、cos（対称）とsin（反対称）の使い分けが鍵。

#### 定義5: Eigentheme因果強度（Eigentheme Causal Intensity）

テーマレベルの因果性:

```
ECI(Theme_a → Theme_b) = Σ_{i∈Theme_a} Σ_{j∈Theme_b} SCD(i, j) / (|Theme_a| · |Theme_b|)

  Theme_a: Eigentheme aに強く荷重するノード集合
  Theme_b: Eigentheme bに強く荷重するノード集合
```

### 2-2. Hodge分解との統合

ユーティリティグラフのエッジフローにHodge分解を適用:

```
Step 1: エッジフローの定義
  ω(i,j) = w(i,j) · d(i,j)  （重み × 方向性）

Step 2: Hodge分解
  ω = ∇φ + δψ + h

Step 3: 各成分の因果的解釈
  ∇φ（勾配成分）: 
    ポテンシャル φ(i) = ノードiの「因果的深さ」
    φ が高い → 原因側（上流）
    φ が低い → 結果側（下流）
    → LiNGAMの因果順序（causal order）に対応する概念

  δψ（カール成分）:
    局所的フィードバックの強さ
    医療: 臓器間フィードバック（腎→心→腎 など）

  h（調和成分）:
    大域的循環パターン
    医療: 全身性の恒常性維持メカニズム
```

**核心的洞察**: Hodge分解のポテンシャル φ(i) は、LiNGAMのcausal orderingの**連続値版**と見なせる。LiNGAMが離散的な順序（x3 → x0 → x2 → x1）を出力するのに対し、Hodge分解は各ノードに連続的な「因果的ポテンシャル」を割り当てる。

### 2-3. 因果フーリエ解析との統合

```
提案: ユーティリティDAG上の因果フーリエ解析

Step 1: ユーティリティグラフからDAG成分を抽出
  Hodge分解の勾配成分 → 方向的に整合するエッジのみ残す → DAG

Step 2: DAG上の因果フーリエ変換
  Seifert et al. (2023)の手法を適用
  フーリエ係数 = 各「原因ノード」の寄与

Step 3: 因果スパース性の仮定
  医療的仮定: 「患者の状態を決める根本原因は少数」
  = 因果フーリエスパース性

Step 4: スパース復元
  少数の観測から因果構造を復元
```

---

## 3. LiNGAMとの接続: 二つの因果同定原理の関係

### 3-1. LiNGAMの同定原理（復習）

```
x = Bx + e,  e ~ 非ガウス独立

ICA → W = B^{-1} を推定
非ガウス性 → 混合行列の一意的同定（Comon's theorem）
→ 因果方向の同定
```

**同定のカギ**: 非ガウス分布は線形混合に対して「向き」を持つ（ガウス分布は回転対称だが、非ガウスはそうではない）。

### 3-2. スペクトル因果性の同定原理（提案）

```
提案する同定原理:
  ユーティリティ関数 Util(i→j) は一般に非対称: Util(i→j) ≠ Util(j→i)
  
  この非対称性が磁気ラプラシアンの固有ベクトルの複素位相に符号化される
  → 位相差から因果方向を復元

同定のカギ: 
  「iのデータがjの問いに答える度合い」と「jのデータがiの問いに答える度合い」の
  非対称性は、因果関係を反映する傾向がある
  
  例: HbA1cのデータは「eGFR低下を予測できるか？」に答えるが、
      eGFRのデータは「HbA1cの変動を予測できるか？」にはあまり答えない
      → Util(HbA1c → eGFR) > Util(eGFR → HbA1c)
      → 因果方向 HbA1c → eGFR と整合
```

### 3-3. 二つの原理の比較

```
LiNGAM:                          スペクトル因果性:
  非ガウス性 → 因果方向              ユーティリティ非対称性 → 因果方向
  統計的性質（分布の形）             意味的性質（臨床的有用性の非対称性）
  データ駆動                        知識+データ駆動
  数学的保証あり                    仮説段階
  変数レベル                        テーマレベル可
```

### 3-4. 両原理が一致する条件・乖離する条件

**一致すると予想される場合**:
- 明確な因果方向がある（喫煙→肺がん）
- ユーティリティ非対称性が強い（原因の情報は結果の予測に有用だが逆は弱い）
- 非ガウス性も強い（原因変数の分布が歪んでいる）

**乖離すると予想される場合**:
- 双方向因果（フィードバック）: LiNGAMはDAG仮定で一方向を選ぶが、スペクトル因果性はフィードバックを表現可能
- ユーティリティが対称的: 互いに等しく情報提供する変数ペア → スペクトル因果性は方向不確定だが、LiNGAMは非ガウス性で方向を決めうる
- 交絡: LiNGAMは交絡を見落とす（基本版）が、ユーティリティ質問は交絡の存在を暗黙的に反映しうる

### 3-5. アンサンブル因果方向推定

```
提案: Ensemble Causal Direction (ECD)

ECD(i → j) = α · LiNGAM_direction(i,j) + β · SCD(i,j) + γ · Granger(i→j)

  α, β, γ: 重み（クロスバリデーションで決定）

メタ分析的解釈:
  - 3つの異なる原理で因果方向を推定
  - 一致すれば高信頼
  - 不一致は「因果関係の複雑さ」の指標
    → 単純因果ではなくフィードバック/交絡の存在を示唆
```

---

## 4. 医療データでの具体的検証シナリオ

### 4-1. シナリオA: 健診データ（横断）でのPoC

```
データ: 特定健診データ（横断スナップショット）
変数: BMI, 腹囲, BP, HbA1c, LDL, HDL, TG, γGT, Cr, eGFR, 尿蛋白, Hb, ...

Step 1: DirectLiNGAM → 変数間因果DAG（ベースライン）

Step 2: ユーティリティ質問生成
  - 各変数に対し「この検査値があれば何が答えられるか」をLLM生成
  - 変数ペアごとにUtil(i→j)とUtil(j→i)を計算

Step 3: ユーティリティ磁気ラプラシアン構築 → 固有分解

Step 4: スペクトル因果方向（SCD）の計算

Step 5: 比較
  - LiNGAMの因果方向 vs SCDの因果方向の一致率
  - 一致するペア = 高信頼因果関係
  - 不一致ペア = 調査価値のある候補（なぜ乖離するか？）

Step 6: Hodge分解
  - 因果的ポテンシャル φ(i) vs LiNGAM因果順序の比較
  - 一致すれば: Hodge分解がLiNGAM因果順序の連続値版として機能

既存研究との直接比較:
  Kotoku et al. (2020)の大阪府健診データ結果と照合可能
```

### 4-2. シナリオB: 縦断データ（経時）でのPoC

```
データ: 年次健診データ（複数年追跡）or MIMIC-IV（ICUデータ）
変数: 同上 + 時間軸

Step 1: VARLiNGAM → 時間遅れ因果 + 同時点因果

Step 2: 時間的ユーティリティグラフ構築
  - ノード = (変数, 時点)
  - ユーティリティ質問に時間文脈を追加
  - 時間方向のエッジに方向性付与

Step 3: 磁気ラプラシアン → 固有分解 → Eigentrajectories

Step 4: Hodge分解
  - 勾配成分 = 時間的因果フロー
  - カール成分 = フィードバックループの検出
  - 調和成分 = 恒常性維持パターン

Step 5: VARLiNGAMの結果との照合
  - VARLiNGAMのB_τ行列の非ゼロパターン vs Eigentrajectory構造
  - Hodge分解の因果ポテンシャル vs VARLiNGAMの因果順序

Okuda et al. (2025)の手法と直接比較可能:
  同じ日本健診コホートデータで、
  ワークフロー制約LiNGAM vs ユーティリティ制約スペクトル因果性
```

### 4-3. シナリオC: 単一細胞データとの類似性

Maehara & Ohkawa (2019)のHodge分解による単一細胞軌跡推定は、「細胞状態の方向的フロー」を抽出した。

```
対応関係:
  単一細胞データ               医療データ
  細胞 = ノード               患者(時点) = ノード
  遺伝子発現 = 信号            検査値 = 信号
  分化方向 = フロー            疾患進行方向 = フロー
  Hodge勾配 = 分化軌跡         Hodge勾配 = 疾患進行軌跡

示唆: 
  単一細胞で成功したHodge分解アプローチが、
  患者レベルの疾患進行分析にも転用可能
```

---

## 5. 理論的課題と限界

### 5-1. ユーティリティ非対称性は因果性を反映するか？

**最大の理論的課題**。ユーティリティ非対称性 Util(i→j) ≠ Util(j→i) が因果方向と整合する保証はない。

反例の可能性:
- **交絡による見かけの非対称性**: ZがXとYの共通原因で、Z→Xが強くZ→Yが弱い場合、Util(X→Y)が高くなるが、これはXがYの原因であることを意味しない
- **情報量の非対称性 ≠ 因果の非対称性**: 高次元変数は低次元変数より多くの問いに答えうる（情報量が多いだけ）

**対処**:
1. ユーティリティ質問を**条件付き**にする: 「他の変数を制御した上で、iが追加的にjの問いに答えられるか」→ 交絡の影響を軽減
2. 情報量の正規化: Util(i→j) を i の情報量で正規化
3. 経験的検証: 既知の因果関係（RCTで確立済み）でユーティリティ非対称性を検証

### 5-2. 磁気ラプラシアンの電荷パラメータ q の選択

q の値が結果に大きく影響する:
- q = 0: 方向性を完全無視（通常の対称ラプラシアン）
- q = 0.25: 最大方向性感度
- q = 0.5: エッジ方向を反転

**提案**: q を**ハイパーパラメータとして扱い**、LiNGAMの因果方向との一致率を最大化するqを選択。これにより、qの値自体が「データの因果構造の強さ」の指標になる。

### 5-3. スケーラビリティ

- 磁気ラプラシアンの固有分解: O(N³) or ランダム化手法でO(Nk²)（k = 上位k個の固有ベクトル）
- Hodge分解: O(|E|)（エッジ数に線形）→ スパースグラフなら効率的
- ユーティリティ質問生成: O(N × m × LLMコスト)（N=ノード数、m=質問数）→ **最大のボトルネック**

### 5-4. 識別可能性の理論的保証

LiNGAMには明確な識別可能性条件（非ガウス + 線形 + DAG + 共通原因なし）がある。
スペクトル因果性には**現時点で識別可能性の理論がない**。

**理論構築への道筋**:
1. 特殊ケース（ツリー構造DAG + 線形SEM）でユーティリティ非対称性とSCDが因果方向と一致することを証明
2. より一般的な構造への拡張
3. 識別不可能な条件（フィードバック、対称ユーティリティ等）の特定

---

## 6. 研究ロードマップ

### Phase 1: 概念実証（3-6ヶ月）
- 合成データ（既知の因果DAG + 非線形SEM）でSCDの性能評価
- LiNGAMとの因果方向一致率の測定
- Hodge分解のポテンシャルとLiNGAM因果順序の相関分析

### Phase 2: 実データ検証（6-12ヶ月）
- 公開健診データ or MIMIC-IVで横断・経時のPoC
- Kotoku et al. / Okuda et al. の結果との比較
- Eigenthemeの臨床解釈可能性評価（臨床専門家レビュー）

### Phase 3: 理論構築（12-18ヶ月）
- 識別可能性条件の導出（特殊ケースから）
- アンサンブル因果方向推定の統計的性質
- 磁気ラプラシアンの電荷パラメータqの理論的選択基準

### Phase 4: 応用・論文化
- 「Spectral Causality: Causal Direction from Utility Graph Spectra」
- 「Hodge-Causal Decomposition for Longitudinal Health Data」
- 「Ensemble Causal Discovery: Merging LiNGAM, Spectral, and Information-Theoretic Approaches」

---

## 7. 全体像の図示

```
                    ┌─────────────────────────────┐
                    │       医療データ              │
                    │  (横断 or 経時)              │
                    └──────────┬──────────────────┘
                               │
              ┌────────────────┼────────────────┐
              ▼                ▼                ▼
     ┌────────────┐  ┌──────────────┐  ┌────────────────┐
     │  LiNGAM    │  │  ユーティリティ  │  │  Granger/TE   │
     │ (変数レベル) │  │  質問生成      │  │ (時系列のみ)   │
     │ 非ガウス性   │  │  LLM + 臨床知識 │  │ 予測改善       │
     └─────┬──────┘  └──────┬───────┘  └───────┬────────┘
           │                │                   │
           ▼                ▼                   │
     ┌────────────┐  ┌──────────────────┐      │
     │ 因果DAG     │  │ ユーティリティ      │      │
     │ B行列       │  │ 磁気ラプラシアン   │      │
     │ 因果順序     │  │ L^(q)_U           │      │
     └─────┬──────┘  └──────┬───────────┘      │
           │                │                   │
           │         ┌──────┴───────────┐      │
           │         ▼                  ▼      │
           │  ┌──────────────┐  ┌──────────┐  │
           │  │ スペクトル分解  │  │ Hodge分解 │  │
           │  │ Eigenthemes  │  │ ∇φ+δψ+h │  │
           │  │ SCI/SCD      │  │ 因果ポテン │  │
           │  └──────┬───────┘  │ シャル    │  │
           │         │         └────┬─────┘  │
           │         │              │         │
           └─────────┼──────────────┼─────────┘
                     ▼              ▼
              ┌──────────────────────────┐
              │  アンサンブル因果推定       │
              │  ECD = αLiNGAM+βSCD+γGr   │
              │  一致 → 高信頼因果         │
              │  不一致 → フィードバック/交絡│
              └──────────┬───────────────┘
                         ▼
              ┌──────────────────────────┐
              │  臨床的解釈               │
              │  テーマレベル因果マップ     │
              │  疾患進行モード            │
              │  治療介入の効果テーマ       │
              └──────────────────────────┘
```

---

## 参考文献

### スペクトルグラフ理論（有向グラフ）
- Yu et al. (2025). "On Directed Graphs With Real Laplacian Spectra." arXiv:2508.05150
- Bauer (2012). "Normalized graph Laplacians for directed graphs." arXiv:1107.4847
- Marques, Segarra & Mateos (2020). "Signal Processing on Directed Graphs." IEEE SPM.

### 磁気ラプラシアン
- de Resende & da Costa (2020). "Characterization and comparison of large directed graphs through the spectra of the magnetic Laplacian." arXiv:2007.03466
- Zhang et al. (2022). "MGC: A Complex-Valued Graph Convolutional Network for Directed Graphs." arXiv:2110.07570
- Ko & Kim (2023). "A Graph Convolution for Signed Directed Graphs." arXiv:2208.11511

### 因果フーリエ解析
- Seifert, Wendler & Püschel (2023). "Causal Fourier Analysis on Directed Acyclic Graphs and Posets." arXiv:2209.07970
- Misiakos, Mihal & Püschel (2024). "Learning Signals and Graphs from Time-Series Graph Data with Few Causes." ICASSP.
- Dallakyan (2023). "On Learning Time Series Summary DAGs: A Frequency Domain Approach." arXiv:2304.08482

### Hodge分解
- Jiang et al. (2011). "Statistical ranking and combinatorial Hodge theory." Mathematical Programming.
- Sugihara et al. (2016). "Hodge Decomposition of Information Flow on Small-World Networks." Frontiers in Neural Circuits.
- Maehara & Ohkawa (2019). "Modeling latent flows on single-cell data using the Hodge decomposition." bioRxiv.

### スペクトル正則化×因果発見
- M'Charrak et al. (2025). "Connected Causal Graphs for Real-World Science." OpenReview.
- OpenReview (2023). "Learning DAGs from Fourier-Sparse Data."
- Kumar et al. (2019). "Structured Graph Learning Via Laplacian Spectral Constraints." NeurIPS.

### LiNGAM×医療データ
- Kotoku et al. (2020). "Causal relations of health indices inferred statistically using the DirectLiNGAM algorithm." PLOS ONE.
- Okuda et al. (2025). "Operationalizing Longitudinal Causal Discovery Under Real-World Workflow Constraints." arXiv:2602.23800
- Kawaguchi (2021). "Application of quantum computing to a linear non-Gaussian acyclic model for novel medical knowledge discovery." arXiv:2110.04485

### 因果推論（プロセスグラフ）
- Reiter et al. (2026). "Causal inference on process graphs." Bernoulli.

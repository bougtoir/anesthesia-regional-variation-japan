# 麻酔関連診療行為の地域差：感度分析フレームワーク

## 概要

「治療バリエーションの地域差は、査定（保険審査）の違いで全て説明されるのか、それとも本当に治療の地域差があるのか？」を麻酔関連診療行為について、公開データ（SCR: 年齢性別調整済み標準化レセプト出現比）を用いて検証するフレームワーク。

---

## 1. 使用データ

| データ | 出典 | 粒度 | 年度 |
|---|---|---|---|
| SCR（都道府県別） | [内閣府 地域差の「見える化」](https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/index.html) | 47都道府県 | 令和4年度(2022) |
| SCR（二次医療圏別） | 同上 | 335二次医療圏 | 令和4年度(2022) |
| 審査統計 | [支払基金](https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html) | 都道府県別・月次 | 令和4年度 |
| 審査の差異の可視化レポート | [支払基金](https://www.ssk.or.jp/shinryohoshu/saikaisyou_torikumi/kashikarepo/) | 診療行為別 | 随時更新 |

---

## 2. 麻酔関連コードのSCR地域差（都道府県レベル R4年度）

| コード | 名称 | 最小 | 最大 | 最大/最小 | CV(%) | 特徴 |
|---|---|---:|---:|---:|---:|---|
| L008 | 閉鎖循環式全身麻酔 | 64.0 | 126.2 | 2.0x | 16.3 | **最も安定**。2倍未満のばらつき |
| L002 | 硬膜外麻酔 | 35.4 | 233.4 | 6.6x | 44.7 | 大きなばらつき |
| L004 | 脊椎麻酔 | 50.3 | 192.6 | 3.8x | 35.0 | 中程度のばらつき |
| L005 | 下肢伝達麻酔 | 35.2 | 280.0 | 8.0x | 59.6 | 大きなばらつき |
| L001 | 静脈麻酔 | 22.4 | 564.1 | 25.2x | 85.8 | 極端なばらつき |
| L009 | 麻酔管理料1 | 44.4 | 148.4 | 3.3x | 26.8 | 専門医配置を反映 |
| L100 | 神経ブロック | 35.1 | 226.8 | 6.5x | 45.6 | ペインクリニック活動を反映 |
| L104 | トリガーポイント注射 | 48.2 | 246.9 | 5.1x | 36.4 | ペインクリニック活動を反映 |

### L008（全身麻酔）都道府県別SCR（上位・下位10）

**高SCR（全身麻酔が多い）**: 北海道(126.2), 鳥取(125.8), 福岡(124.8), 佐賀(124.1), 東京(119.4), 京都(118.9), 富山(112.7), 大阪(112.4), 鹿児島(111.2), 奈良(110.1)

**低SCR（全身麻酔が少ない）**: 岐阜(64.0), 山形(66.2), 三重(69.9), 青森(70.1), 愛知(77.0), 新潟(77.2), 岩手(80.7), 埼玉(81.5), 福島(84.4), 茨城(86.0)

---

## 3. 感度分析フレームワーク

### 3.1 帰無仮説

> **H0: 麻酔関連診療行為のSCR地域差は、全て保険審査（査定）の都道府県間差異によって説明される**

この仮説が正しい場合、以下の3つの予測が成立するはず。

### 3.2 検定1: 県内分散ゼロの予測

**ロジック**: 査定は都道府県単位（支払基金支部・国保連合会）で運用される → 査定が全てを説明するなら、同一県内の二次医療圏間には統計ノイズ以外の差がないはず。

**方法**: 二次医療圏レベルSCR (n=335) を都道府県間分散と都道府県内分散に分解。

**結果**:

| コード | 名称 | 県間分散寄与率 | 県内分散寄与率 | 判定 |
|---|---|---:|---:|---|
| L008 | 全身麻酔 | **14.2%** | **85.8%** | **棄却**（査定では説明不可） |
| L002 | 硬膜外麻酔 | 27.3% | 72.7% | **棄却** |
| L004 | 脊椎麻酔 | 42.0% | 58.0% | 判断留保 |
| L005 | 下肢伝達麻酔 | 46.1% | 53.9% | 判断留保 |
| L100 | 神経ブロック | 34.9% | 65.1% | **棄却** |
| L104 | トリガーポイント注射 | 41.9% | 58.1% | 判断留保 |

**解釈**:
- **全身麻酔(L008)**: 地域差の85.8%は同じ県内の二次医療圏間で生じている。都道府県単位の査定方針では最大でも14.2%しか説明できない。
- **硬膜外麻酔(L002)**: 同様に72.7%が県内要因。
- **脊椎麻酔(L004)・下肢伝達(L005)**: 県間寄与率が40%台あり、査定方針の影響がやや大きい可能性。ただし過半は依然県内要因。

**県内ばらつきの具体例（全身麻酔L008 SCR）**:
- 東京都内: 61.4〜458.9（397.5ポイント幅、12医療圏）
- 群馬県内: 10.7〜198.0（187.3ポイント幅、10医療圏）
- 熊本県内: 2.3〜170.9（168.6ポイント幅、10医療圏）
- 大阪府内: 73.5〜136.0（62.5ポイント幅、8医療圏）← 大阪は比較的均質

### 3.3 検定2: クロスコード相関の予測

**ロジック**: 同一の査定ロジック下にあるコードは、完全に相関するはず。逆に、代替関係（全身麻酔 vs 脊椎麻酔等）を示すなら、それは臨床選択の反映。

**結果（都道府県レベル r）**:

| ペア | r値 | 解釈 |
|---|---:|---|
| 全身麻酔 vs 麻酔管理料1 | +0.788 | 強い正の相関（専門医供給の共通要因） |
| 全身麻酔 vs 脊椎麻酔 | **-0.506** | **強い負の相関（代替関係）** |
| 全身麻酔 vs 下肢伝達 | -0.324 | 負の相関（代替関係） |
| 全身麻酔 vs 硬膜外麻酔 | +0.316 | 弱い正の相関 |
| 硬膜外 vs 脊椎 | -0.169 | 弱い負の相関（代替関係？） |
| 神経ブロック vs トリガーポイント | +0.022 | 無相関（異なる決定要因） |

**解釈**:
- **全身麻酔 vs 脊椎麻酔の r=-0.506** は臨床的に重要。全身麻酔が少ない地域で脊椎麻酔が多い → 麻酔方法の選択に関する**真の臨床的地域差**が存在する強い証拠。査定であれば両方とも同方向に増減するはず。
- 全身麻酔 vs 下肢伝達の負の相関も同様。区域麻酔を好む地域が存在する。

### 3.4 検定3: 査定率インパクトの数量的評価

**ロジック**: 査定率の都道府県間差異が、SCRの地域差を数量的に説明できるか。

**データ**:
- 支払基金審査統計による査定率の地域差: **最大0.28%（大阪）〜最小0.07%（宮崎）**
- つまり査定率の都道府県間レンジは約**0.2%ポイント**

**SCRへの最大インパクト推計**:
- 査定率 0.2%の差 = クレーム1万件中20件が査定で削除される差
- SCR 100の地域で、査定率が0.2%高い場合 → SCRは約0.2ポイント低下（100 → 99.8）

**一方、実際のSCRの地域差**:
| コード | SCR幅（ポイント） | 査定で説明可能な幅 | 査定の説明力 |
|---|---:|---:|---|
| 全身麻酔 L008 | 62 | 0.2 | **0.3%** |
| 硬膜外麻酔 L002 | 198 | 0.2 | 0.1% |
| 脊椎麻酔 L004 | 142 | 0.2 | 0.1% |
| 下肢伝達 L005 | 245 | 0.2 | 0.08% |

**結論**: **査定率差（0.2%ポイント）で SCR差（62〜245ポイント）を説明するのは数量的に不可能**。査定は地域差の1%未満しか説明しない。

> ⚠️ **注意**: これは「全ての査定率」の比較。診療行為コード別の査定率データが公開されていないため、特定コードの査定率差がもっと大きい可能性は否定できない。ただし、トータルの査定率が0.2%レンジである以上、特定コードだけ数十%差がつくことは論理的に困難。

---

## 4. 二次医療圏レベルの追加分析

### 4.1 二次医療圏レベルのSCR分布（n=335）

| コード | 名称 | P10 | P25 | P50 | P75 | P90 | CV(%) |
|---|---|---:|---:|---:|---:|---:|---:|
| L008 | 全身麻酔 | 32.3 | 49.5 | 73.2 | 102.3 | 131.5 | 54.6 |
| L002 | 硬膜外麻酔 | 7.6 | 31.8 | 73.9 | 126.5 | 184.0 | 83.2 |
| L004 | 脊椎麻酔 | 31.1 | 56.6 | 84.3 | 125.9 | 169.0 | 56.3 |
| L005 | 下肢伝達麻酔 | 16.0 | 37.7 | 80.1 | 154.0 | 222.9 | 86.8 |
| L009 | 麻酔管理料1 | 32.5 | 51.1 | 79.1 | 116.2 | 149.1 | 57.2 |
| L100 | 神経ブロック | 23.4 | 42.7 | 72.2 | 118.7 | 174.7 | 72.1 |
| L104 | トリガーポイント | 56.0 | 73.1 | 97.7 | 137.0 | 184.7 | 54.0 |

二次医療圏レベルでは、都道府県レベルよりCVが大幅に拡大（L008: 16.3% → 54.6%）。これは県内の二次医療圏間でも大きなばらつきがあることを示す。

### 4.2 大学医局の影響範囲の代理指標

麻酔管理料1（L009）は麻酔科標榜医の管理下で算定されるため、L009/L008比率は**麻酔科専門医の配置密度**の代理指標となる。

**管理料比率が高い二次医療圏（専門医が充実）**:
- 沖縄4705圏(2.23), 沖縄4704圏(2.21), 群馬1009圏(2.19), 群馬1002圏(2.11), 福岡4003圏(2.10)

**管理料比率が低い二次医療圏（専門医が希少）**:
- 秋田507圏(0.00), 島根3202圏(0.01), 山口3508圏(0.03), 岩手307圏(0.05), 茨城802圏(0.15)

**相関分析（二次医療圏レベル）**:
- 管理料比率 vs 区域麻酔指数: r=-0.044（ほぼ無相関）
- 管理料比率 vs ペインクリニック指数: r=-0.011（無相関）

→ **専門医の配置密度自体は区域麻酔の選好とは独立**。これは「専門医がいるから区域麻酔が多い」という単純な仮説を否定する。

---

## 5. ペインクリニック→区域麻酔スピルオーバー仮説の検証

### 5.1 仮説

> ペインクリニック活動が盛んな地域の麻酔科医は区域麻酔手技に習熟しており、手術麻酔でも区域麻酔を併用する傾向がある。

### 5.2 相関分析（二次医療圏レベル n=335）

| ペイン指標 | vs | r値 | 判定 |
|---|---|---:|---|
| 神経ブロック(L100) | 硬膜外麻酔(L002) | +0.166 | 弱い正の関連 |
| 神経ブロック(L100) | 脊椎麻酔(L004) | +0.114 | 弱い正の関連 |
| 神経ブロック(L100) | 下肢伝達(L005) | +0.028 | 無相関 |
| 神経ブロック(L100) | 全身麻酔(L008) | **+0.307** | **中程度の正の関連** |
| トリガーポイント(L104) | 硬膜外麻酔(L002) | +0.074 | 無相関 |
| トリガーポイント(L104) | 脊椎麻酔(L004) | -0.028 | 無相関 |
| トリガーポイント(L104) | 下肢伝達(L005) | +0.140 | 弱い正の関連 |
| トリガーポイント(L104) | 全身麻酔(L008) | -0.028 | 無相関 |
| **ペイン総合指数** | **区域麻酔総合指数** | **+0.153** | **弱い正の関連** |

### 5.3 解釈

- **直接的なスピルオーバー効果は限定的**（r=0.15程度）。ペインクリニックが盛んだからといって、手術麻酔での区域麻酔が劇的に増える傾向は確認されない。
- ただし、**神経ブロック(L100) → 全身麻酔(L008) で r=+0.307** という中程度の相関がある。これは「ペインクリニックが活発な地域 = 麻酔科のプレゼンスが高い地域 = 全身麻酔も多い」という**供給要因**の反映と解釈できる。
- トリガーポイント注射(L104)は整形外科等でも実施されるため、ペインクリニック活動の純粋な指標としては神経ブロック(L100)の方が適切。
- **仮説の修正案**: ペインクリニック活動は区域麻酔選好の直接要因ではなく、**麻酔科の全体的なプレゼンス**（マンパワー、教育体制）の代理指標として機能する。

---

## 6. 総合的結論

### 6.1 帰無仮説「査定が地域差の全てを説明する」の判定

| 検定 | 結果 | 判定 |
|---|---|---|
| 検定1: 県内分散ゼロ | 全身麻酔の地域差の85.8%が県内変動 | **H0棄却** |
| 検定2: クロスコード相関 | 全身麻酔と脊椎麻酔が逆相関(r=-0.506) | **H0棄却**（臨床選択の証拠） |
| 検定3: 査定率インパクト | 査定率差0.2%でSCR差62ポイントは説明不能 | **H0棄却** |

**結論**: **麻酔関連診療行為の地域差は、査定の違いだけでは説明できない。真の治療バリエーション（臨床的地域差）が存在する。**

### 6.2 地域差を生み出す要因の候補

感度分析結果から推定される要因の寄与度：

| 要因 | 推定寄与 | 根拠 |
|---|---|---|
| **二次医療圏固有の要因**（施設特性、医師の訓練歴、大学医局方針） | 最大 | 県内分散が支配的(58-86%) |
| **都道府県レベルの要因**（査定方針を含む） | 14-46% | 県間分散の寄与率 |
| **査定方針（狭義）** | <1% | 査定率差のSCRインパクト推計 |
| **麻酔方法の代替関係** | 有意 | 全身-脊椎のr=-0.506 |
| **ペインクリニック活動（スピルオーバー）** | 限定的 | r=0.15程度 |

### 6.3 大学病院（医局）効果の検証

→ **詳細は第7章を参照**

---

## 7. 大学病院（医局）効果の検証

### 7.1 概要と方法

81大学病院（国立44, 公立8, 私立29）を64の二次医療圏にマッピングし、大学病院所在圏 vs 非所在圏のSCRパターンを比較した。マッピングは各大学の所在地（市区町村）と二次医療圏の対応に基づく。

- **大学病院所在圏**: 64圏（335圏の19.1%）
- **非所在圏**: 271圏（80.9%）
- **複数大学がある圏**: 13圏（東京区中央部5校、区西部3校、大阪三島2校 等）

> 注: 日本は「一県一医大」政策により全47都道府県に少なくとも1校の医学部がある。よって「大学病院がない県」は存在せず、県間比較ではなく**県内比較**が本質的な分析手法となる。

### 7.2 全体比較：大学病院所在圏 vs 非所在圏

| コード | 名称 | 大学圏 mean (n) | 非大学圏 mean (n) | 差 | Cohen's d | t値 |
|---|---|---:|---:|---:|---:|---:|
| L008 | 全身麻酔 | 130.2 (64) | 67.7 (270) | **+62.4** | **+1.780** | 9.72 |
| L001 | 静脈麻酔 | 137.4 (64) | 56.3 (263) | +81.1 | +1.238 | 7.01 |
| L009 | 麻酔管理料1 | 124.8 (64) | 77.6 (250) | +47.2 | +1.021 | 6.96 |
| L002 | 硬膜外麻酔 | 127.4 (64) | 87.2 (243) | +40.2 | +0.491 | 3.75 |
| L100 | 神経ブロック(入院) | 127.1 (64) | 93.3 (271) | +33.8 | +0.444 | 3.14 |
| L100 | 神経ブロック(外来) | 115.7 (64) | 82.1 (271) | +33.6 | +0.525 | 3.59 |
| L004 | 脊椎麻酔 | 108.6 (64) | 90.3 (270) | +18.3 | +0.345 | 2.73 |
| L005 | 下肢伝達麻酔 | 120.3 (63) | 113.9 (259) | +6.4 | +0.055 | 0.44 |
| L104 | トリガーポイント(外来) | 107.2 (64) | 112.9 (271) | -5.8 | -0.096 | -0.78 |

**解釈**:
- **全身麻酔(L008)**: Cohen's d=1.78は極めて大きな効果量。大学病院圏は国全体平均より約30%高く、非大学圏は約30%低い。
- **静脈麻酔(L001)**: d=1.24。大学病院圏への集中が最も顕著なコードの一つ。
- **下肢伝達麻酔(L005)とトリガーポイント(L104)**: 差なし。これらは大学病院の有無と無関係に地域のローカル要因で決まる。

### 7.3 同一県内比較（査定方針を統制した分析）

県レベルの査定方針は全ての二次医療圏に等しく適用される。よって**同一県内**で大学病院圏と非大学圏を比較すれば、査定の影響を完全に排除した「純粋な大学病院効果」を推定できる。

| コード | 名称 | 県内差(平均) | SD | t値 | 大学圏>非大学圏 |
|---|---|---:|---:|---:|---|
| L008 | 全身麻酔 | **+61.4** | 33.7 | **12.49** | **47/47県 (100%)** |
| L001 | 静脈麻酔 | +84.5 | 75.3 | 7.70 | 42/47県 (89%) |
| L009 | 麻酔管理料1 | +46.8 | 41.4 | 7.74 | 40/47県 (85%) |
| L002 | 硬膜外麻酔 | +51.6 | 72.3 | 4.89 | 39/47県 (83%) |
| L100 | 神経ブロック(入院) | +43.9 | 64.5 | 4.66 | 34/47県 (72%) |
| L004 | 脊椎麻酔 | +20.5 | 37.5 | 3.75 | 31/47県 (66%) |

**重要な発見**:
- **全身麻酔(L008)**: **全47県**で大学病院圏のSCRが非大学圏を上回る。平均差+61.4ポイント、t=12.49は統計的に極めて有意。例外は一つもない。
- この効果は県の査定方針とは**完全に独立**。同じ査定環境下で、大学病院がある二次医療圏は一貫して麻酔関連行為のSCRが高い。
- 硬膜外麻酔(L002)でも83%の県で大学圏が高い。区域麻酔を含む幅広い手技で大学病院効果が確認される。

### 7.4 3段階分散分解

全SCR分散を「県間」「県内の大学病院効果」「残差」の3成分に分解:

| コード | 名称 | 県間 | 大学病院効果 | 残差 | 県内分散に占める大学効果 |
|---|---|---:|---:|---:|---:|
| L008 | 全身麻酔 | 14.5% | **38.5%** | 47.0% | **45.0%** |
| L001 | 静脈麻酔 | 21.3% | **32.6%** | 46.1% | 41.5% |
| L009 | 麻酔管理料1 | 20.0% | 23.7% | 56.4% | 29.6% |
| L002 | 硬膜外麻酔 | 22.4% | 14.0% | 63.7% | 18.0% |
| L100 | 神経ブロック(入院) | 28.9% | 12.7% | 58.4% | 17.8% |
| L004 | 脊椎麻酔 | 41.7% | 8.6% | 49.7% | 14.7% |

**解釈**:
- **全身麻酔(L008)**: 前回の分析で県内分散が85.5%と判明していたが、その県内分散のうち**45.0%が大学病院の有無で説明される**。全分散に対して38.5%という寄与は、単一の二値変数（大学病院がある/ない）としては非常に大きい。
- **残差47.0%**: 大学病院の有無だけでは説明しきれない分散が半分残る。これは個別病院の特性、地理的条件、人口密度等の他の要因を示唆。
- **脊椎麻酔(L004)**: 県間分散が最大（41.7%）で大学病院効果が最小（8.6%）。脊椎麻酔の地域差は大学病院よりも県レベルの要因（地域の臨床文化、査定を含む）で説明される可能性が高い。

### 7.5 注目県の詳細パターン

#### 7.5.1 熊本県（典型的な「大学病院集中」パターン）

| 二次医療圏 | SCR(L008) | 大学病院 |
|---|---:|---|
| 熊本・上益城 | **168.7** | ★ 熊本大学病院 |
| 八代 | 117.3 | - |
| 球磨 | 73.8 | - |
| 天草 | 57.0 | - |
| 菊池 | 49.7 | - |
| 有明 | 48.6 | - |
| 芦北 | 41.7 | - |
| 鹿本 | 37.3 | - |
| 宇城 | 7.5 | - |
| 阿蘇 | 2.3 | - |

→ 大学病院圏(168.7) vs 非大学圏平均(48.4) = **差120.3ポイント**。典型的な中核集中パターン。

#### 7.5.2 東京都（多数の大学病院が分散）

| 二次医療圏 | SCR(L008) | 大学病院 |
|---|---:|---|
| 区中央部 | **435.7** | ★ 東大, 東京科学大, 順天堂, 日本医大, 慈恵 (5校) |
| 区西部 | **168.5** | ★ 慶應, 東京医大, 東京女子医大 (3校) |
| 区西北部 | 83.4 | ★ 日大板橋, 帝京 (2校) |
| 区南部 | 118.9 | ★ 昭和, 東邦 (2校) |
| 北多摩南部 | 121.9 | ★ 杏林 (1校) |
| 区西南部 | 110.1 | - |
| 区東部 | 97.8 | - |
| 北多摩西部 | 88.4 | - |
| 南多摩 | 71.1 | - |
| 区東北部 | 61.8 | - |
| 北多摩北部 | 64.9 | - |
| 西多摩 | 62.5 | - |

→ 大学病院数と SCR に明確な用量反応関係が見られる（5校: 435.7 → 3校: 168.5 → 2校: 83-119 → 0校: 62-110）。

#### 7.5.3 茨城県（外れ値: つくば圏のSCRが非常に高い）

| 二次医療圏 | SCR(L008) | 大学病院 |
|---|---:|---|
| つくば | **181.6** | ★ 筑波大学附属病院 |
| 水戸 | 145.5 | - |
| 土浦 | 98.4 | - |
| 日立 | 77.0 | - |
| 取手・竜ヶ崎 | 73.0 | - |
| 古河・坂東 | 63.8 | - |
| 常陸太田・ひたちなか | 36.0 | - |
| 鹿行 | 33.8 | - |
| 筑西・下妻 | 31.5 | - |

→ つくば(181.6) vs 非大学圏平均(82.7) = 差98.9ポイント。水戸も高いが大学病院はない（独自の医療集積）。

### 7.6 結論

1. **大学病院効果は統計的に極めて有意**:
   - 全身麻酔(L008)で**全47県**で大学圏 > 非大学圏（100%の一致率）
   - Cohen's d = 1.78 は社会科学で「極めて大きい」に分類される効果量
   - 全分散の38.5%を大学病院の有無（単一二値変数）で説明

2. **これは査定では説明不可能**:
   - 同一県内比較なので査定方針は統制済み
   - t=12.49は偶然では到底説明不可能

3. **効果の強さはコードにより異なる**:
   - **最強**: 全身麻酔(L008, d=1.78), 静脈麻酔(L001, d=1.24)
   - **中程度**: 麻酔管理料(L009, d=1.02), 硬膜外(L002, d=0.49), 神経ブロック(L100, d=0.44-0.53)
   - **なし**: 下肢伝達(L005, d=0.06), トリガーポイント(L104, d=-0.10)

4. **解釈上の注意点**:
   - 大学病院は通常その県の最大都市に立地しており、「大学病院効果」と「都市集中効果」は完全には分離できない
   - 大学病院圏のSCRが高いのは、(a) 大学病院自体の高い麻酔実施件数、(b) 大学関連病院への波及、(c) 都市部の病院集積、が複合した結果と考えられる
   - 真の「医局効果」（大学の学術的方針が周辺病院の診療パターンに影響）を検証するには、同一都市規模で大学病院がある/ないを比較する追加分析が必要

---

## 8. 限界と今後の展望

### 8.1 本分析の限界

1. **査定率データの粒度不足**: 診療行為コード別の査定率が非公開のため、検定3の推計は全体査定率に基づく間接推計
2. **SCRの構造的限界**: SCR自体が査定後データから算出。ただし、感度分析で示したように査定のインパクトは数量的に小さい
3. **二次医療圏コードの匿名性**: SCRデータの二次医療圏は数値コードのみで、具体的な圏名（例: 区西北部）との対応は別途必要
4. **交絡因子**: 疾病構造の地域差（手術が必要な疾患の発生率）はSCRの年齢性別調整では除去しきれない
5. **時系列分析の不足**: 単年度（R4）のクロスセクション分析のみ。改定前後の経年変化は未検討
6. **大学病院効果 vs 都市集中効果の未分離**: 大学病院は通常県内最大都市に立地するため、「大学病院効果」と「都市部の病院集積効果」が交絡している。人口規模を統制した分析が必要

### 8.2 今後のステップ

| ステップ | 内容 | 難易度 | データ |
|---|---|---|---|
| A | 大学病院効果と都市集中効果の分離（同一人口規模での比較） | 中 | 人口統計+本分析データ |
| B | 大学の麻酔科の学術的傾向と影響圏のSCRパターンの対応 | 中 | 学会発表+本分析データ |
| C | 令和8年改定前後のSCR比較（L008名称変更の影響） | 中 | R8データ公開後 |
| D | NDB第三者提供による個票分析（査定前後比較） | 高 | 倫理審査・NDB申請 |
| E | 審査の差異の可視化レポートの麻酔関連事例の探索 | 低 | 支払基金公開資料 |

### 8.3 研究としての発展可能性

この感度分析フレームワークは、麻酔に限らず**あらゆる診療行為の地域差研究**に適用可能：
- 外科手術術式の選択（腹腔鏡 vs 開腹等）
- 薬剤処方パターン
- 検査オーダーパターン

分散分解（県間 vs 県内）とクロスコード相関は、NDB/SCR公開データのみで実施可能な汎用的手法。

---

## 付録: 分析に使用したデータソースとアクセス方法

### SCRデータ（内閣府）
```
# 都道府県別・診療行為区分
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/01_r04_t_kubun.csv

# 都道府県別・診療行為枝番
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/04_r04_t_edaban.csv

# 二次医療圏別・診療行為区分
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/02_r04_n_kubun.csv

# 二次医療圏別・診療行為枝番
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/05_r04_n_edaban.csv
```
エンコーディング: Shift_JIS

### 審査統計（支払基金）
```
https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html
```
令和4年度の年度統計・月次データがZIP形式でダウンロード可能。

---

# English Translation

---

# Regional differences in anesthesia-related medical practices: Sensitivity analysis framework

## Overview

A framework that examines the question, "Are regional differences in treatment variations entirely explained by differences in assessment (insurance screening), or are there really regional differences in treatment?" using public data (SCR: age-sex-adjusted Standardized Receipt Occurrence Ratio) regarding anesthesia-related medical practices.

---

## 1. Usage data

| Data | Source | Granularity | Year |
|---|---|---|---|
| SCR (by prefecture) | [Cabinet Office "Visualization" of regional differences](https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/index.html) | 47 prefectures | FY2022 |
| SCR (by secondary medical area) | Same as above | 335 secondary medical area | 2022 |
| Examination statistics | [Payment fund](https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html) | Monthly by prefecture | FY2020 |
| Visualization report of examination differences | [Payment fund](https://www.ssk.or.jp/shinryohoshu/saikaisyou_torikumi/kashikarepo/) | By medical practice | Updated from time to time |

---

## 2. SCR regional differences in anesthesia-related codes (prefectural level R4 year)

| Code | Name | Minimum | Maximum | Maximum/Minimum | CV(%) | Characteristics |
|---|---|---:|---:|---:|---:|---|
| L008 | Closed circulation general anesthesia | 64.0 | 126.2 | 2.0x | 16.3 | **Most stable**. Variation of less than 2 times |
| L002 | Epidural | 35.4 | 233.4 | 6.6x | 44.7 | Large variation |
| L004 | Spinal Anesthesia | 50.3 | 192.6 | 3.8x | 35.0 | Moderate Variability |
| L005 | Lower extremity delivery anesthesia | 35.2 | 280.0 | 8.0x | 59.6 | Large variability |
| L001 | Intravenous Anesthesia | 22.4 | 564.1 | 25.2x | 85.8 | Extreme Variability |
| L009 | Anesthesia management fee 1 | 44.4 | 148.4 | 3.3x | 26.8 | Reflects specialist assignment |
| L100 | Nerve Block | 35.1 | 226.8 | 6.5x | 45.6 | Reflects Pain Clinic Activities |
| L104 | Trigger Point Injection | 48.2 | 246.9 | 5.1x | 36.4 | Reflects Pain Clinic Activities |

### L008 (general anesthesia) SCR by prefecture (top/bottom 10)

**High SCR (general anesthesia is common)**: Hokkaido (126.2), Tottori (125.8), Fukuoka (124.8), Saga (124.1), Tokyo (119.4), Kyoto (118.9), Toyama (112.7), Osaka (112.4), Kagoshima (111.2), Nara(110.1)

**Low SCR (less general anesthesia)**: Gifu (64.0), Yamagata (66.2), Mie (69.9), Aomori (70.1), Aichi (77.0), Niigata (77.2), Iwate (80.7), Saitama (81.5), Fukushima (84.4), Ibaraki (86.0)

---

## 3. Sensitivity analysis framework

### 3.1 Null hypothesis

> **H0: All regional SCR differences in anesthesia-related medical practices are explained by inter-prefectural differences in insurance screening (assessment)**

If this hypothesis is correct, the following three predictions should hold.

### 3.2 Test 1: Prediction of zero within-prefecture variance
**Logic**: Assessments are operated on a prefectural basis (Payment Fund Branch, National Health Insurance Federation) → If assessments explain everything, there should be no difference other than statistical noise between secondary medical care areas within the same prefecture.

**Method**: Secondary medical area level SCR (n=335) was decomposed into inter-prefectural variance and intra-prefectural variance.

**Result**:

| Code | Name | Inter-prefecture variance contribution rate | Intra-prefecture variance contribution rate | Judgment |
|---|---|---:|---:|---|
| L008 | General anesthesia | **14.2%** | **85.8%** | **Rejected** (cannot be explained in assessment) |
| L002 | Epidural | 27.3% | 72.7% | **Rejected** |
| L004 | Spinal anesthesia | 42.0% | 58.0% | Judgment reserved |
| L005 | Lower extremity delivery anesthesia | 46.1% | 53.9% | Judgment reserved |
| L100 | Nerve block | 34.9% | 65.1% | **Rejected** |
| L104 | Trigger point injection | 41.9% | 58.1% | Reservation of judgment |

**Interpretation**:
- **General anesthesia (L008)**: 85.8% of regional differences occur between secondary medical care areas within the same prefecture. Prefectural assessment policies can only explain 14.2% at most.
- **Epidural anesthesia (L002)**: Similarly, 72.7% were due to local factors.
- **Spinal anesthesia (L004)/Lower extremity transfer (L005)**: The inter-prefectural contribution rate is in the 40% range, and the impact of the assessment policy may be somewhat large. However, the majority is still due to factors within the prefecture.

**Specific example of variation within the prefecture (general anesthesia L008 SCR)**:
- Tokyo: 61.4-458.9 (397.5 point width, 12 medical areas)
- Gunma prefecture: 10.7-198.0 (187.3 point width, 10 medical areas)
- Kumamoto prefecture: 2.3 to 170.9 (168.6 point width, 10 medical areas)
- Osaka Prefecture: 73.5 to 136.0 (62.5 point range, 8 medical areas) ← Osaka is relatively homogeneous

### 3.3 Test 2: Predicting cross-code correlation

**Logic**: Codes under the same assessment logic should be perfectly correlated. Conversely, if an alternative relationship is shown (general anesthesia vs. spinal anesthesia, etc.), it reflects clinical choice.

**Results (prefecture level r)**:

| pair | r value | interpretation |
|---|---:|---|
| General anesthesia vs anesthesia management fee1 | +0.788 | Strong positive correlation (common factor in specialist supply) |
| General anesthesia vs spinal anesthesia | **-0.506** | **Strong negative correlation (surrogate relationship)** |
| General anesthesia vs. lower extremity transfer | -0.324 | Negative correlation (alternative relationship) |
| General anesthesia vs. epidural | +0.316 | Weak positive correlation |
| Epidural vs spinal | -0.169 | Weak negative correlation (alternative relationship?) |
| Nerve block vs trigger point | +0.022 | Uncorrelated (different determinants) |

**Interpretation**:
- **r=-0.506 for general vs. spinal anesthesia** is clinically significant. More spinal anesthesia in areas with less general anesthesia → strong evidence of true clinical regional differences in the choice of anesthesia method. If it is an appraisal, both should increase or decrease in the same direction.
- Similar negative correlation between general anesthesia vs. lower extremity transmission. There are regions that prefer regional anesthesia.

### 3.4 Test 3: Quantitative evaluation of assessment rate impact

**Logic**: Can inter-prefectural differences in assessment rates quantitatively explain regional differences in SCR?

**Data**:
- Regional differences in assessment rates based on payment fund examination statistics: **Maximum 0.28% (Osaka) ~ Minimum 0.07% (Miyazaki)**
- In other words, the inter-prefectural range of assessment rates is approximately **0.2 percentage points**

**Estimated maximum impact on SCR**:
- 0.2% difference in assessment rate = difference in which 20 out of 10,000 complaints are deleted in assessment
- In an area with SCR 100, if the assessment rate is 0.2% higher → SCR will decrease by approximately 0.2 points (100 → 99.8)

**On the other hand, regional differences in actual SCR**:
| Code | SCR width (points) | Width that can be explained by assessment | Explanatory power of assessment |
|---|---:|---:|---|
| General anesthesia L008 | 62 | 0.2 | **0.3%** |
| Epidural anesthesia L002 | 198 | 0.2 | 0.1% |
| Spinal anesthesia L004 | 142 | 0.2 | 0.1% |
| Lower limb transmission L005 | 245 | 0.2 | 0.08% |

**Conclusion**: **It is quantitatively impossible to explain the SCR difference (62 to 245 points) by the assessment rate difference (0.2% point)**. The assessment explains less than 1% of regional differences.

> ⚠️ **Note**: This is a comparison of "all assessment rates". As assessment rate data by medical practice code has not been made public, it cannot be denied that the difference in assessment rates for specific codes may be even greater. However, since the total assessment rate is in the 0.2% range, it is logically difficult for a specific code to have a difference of several tens of percentage points.

---

## 4. Additional analysis at the secondary medical area level

### 4.1 SCR distribution at secondary medical area level (n=335)

| Code | Name | P10 | P25 | P50 | P75 | P90 | CV(%) |
|---|---|---:|---:|---:|---:|---:|---:|
| L008 | General anesthesia | 32.3 | 49.5 | 73.2 | 102.3 | 131.5 | 54.6 |
| L002 | Epidural | 7.6 | 31.8 | 73.9 | 126.5 | 184.0 | 83.2 |
| L004 | Spinal anesthesia | 31.1 | 56.6 | 84.3 | 125.9 | 169.0 | 56.3 |
| L005 | Lower extremity delivery anesthesia | 16.0 | 37.7 | 80.1 | 154.0 | 222.9 | 86.8 |
| L009 | Anesthesia management fee 1 | 32.5 | 51.1 | 79.1 | 116.2 | 149.1 | 57.2 |
| L100 | Nerve Block | 23.4 | 42.7 | 72.2 | 118.7 | 174.7 | 72.1 |
| L104 | Trigger Point | 56.0 | 73.1 | 97.7 | 137.0 | 184.7 | 54.0 |

At the secondary medical area level, CV increased significantly compared to the prefecture level (L008: 16.3% → 54.6%). This shows that there are large variations even among secondary medical care areas within the prefecture.

### 4.2 Proxy indicators for the scope of influence of university medical offices

Anesthesia management fee 1 (L009) is calculated under the supervision of an anesthesiologist, so the L009/L008 ratio serves as a proxy index for the density of anesthesiologist placement.

**Secondary medical care area with high management fee ratio (full of specialists)**:
- Okinawa 4705 area (2.23), Okinawa 4704 area (2.21), Gunma 1009 area (2.19), Gunma 1002 area (2.11), Fukuoka 4003 area (2.10)

**Secondary medical care area with low management fee ratio (specialists are rare)**:
- Akita 507 area (0.00), Shimane 3202 area (0.01), Yamaguchi 3508 area (0.03), Iwate 307 area (0.05), Ibaraki 802 area (0.15)

**Correlation analysis (secondary medical area level)**:
- Management fee ratio vs regional anesthesia index: r=-0.044 (nearly no correlation)
- Management fee ratio vs. pain clinic index: r=-0.011 (no correlation)

→ **Specialist density itself is independent of regional anesthesia preference**. This refutes the simple hypothesis that ``regional anesthesia is more common because there are specialists.''

---

## 5. Pain clinic → Verification of regional anesthesia spillover hypothesis

### 5.1 Hypothesis

> Anesthetists in regions with active pain clinics are proficient in regional anesthesia techniques and tend to use regional anesthesia in conjunction with surgical anesthesia.

### 5.2 Correlation analysis (secondary medical area level n=335)

| Pain index | vs | r value | Judgment |
|---|---|---:|---|
| Nerve block (L100) | Epidural anesthesia (L002) | +0.166 | Weak positive association |
| Nerve block (L100) | Spinal anesthesia (L004) | +0.114 | Weak positive association |
| Nerve block (L100) | Lower limb transmission (L005) | +0.028 | Uncorrelated |
| Nerve block (L100) | General anesthesia (L008) | **+0.307** | **Moderate positive association** |
| Trigger point (L104) | Epidural anesthesia (L002) | +0.074 | No correlation |
| Trigger point (L104) | Spinal anesthesia (L004) | -0.028 | No correlation |
| Trigger point (L104) | Lower extremity transmission (L005) | +0.140 | Weak positive association |
| Trigger point (L104) | General anesthesia (L008) | -0.028 | No correlation |
| **Pain Comprehensive Index** | **Regional Anesthesia Comprehensive Index** | **+0.153** | **Weak positive association** |

### 5.3 Interpretation

- **Direct spillover effect is limited** (about r=0.15). Even if pain clinics become more popular, there is no trend toward a dramatic increase in the use of regional anesthesia for surgical anesthesia.
- However, there is a moderate correlation of r=+0.307** for **nerve block (L100) → general anesthesia (L008). This can be interpreted as a reflection of the **supply factor**: ``A region with active pain clinics = a region with a high presence of anesthesiologists = a large number of general anesthesia patients.''
- Trigger point injections (L104) are also performed in orthopedics, so nerve blocks (L100) are more appropriate as a pure indicator of pain clinic activity.
- **Hypothesis modification**: Pain clinic activity is not a direct factor in regional anesthesia preference, but serves as a proxy indicator of **overall presence** (manpower, educational structure) of the anesthesiology department.

---

## 6. Overall conclusion

### 6.1 Determination of the null hypothesis “Assessment explains all regional differences”

| Test | Results | Judgment |
|---|---|---|
| Test 1: Zero variance within prefecture | 85.8% of regional differences in general anesthesia are within prefecture | **H0 rejected** |
| Test 2: Cross-code correlation | General anesthesia and spinal anesthesia are inversely correlated (r=-0.506) | **H0 rejected** (evidence of clinical selection) |
| Test 3: Assessment rate impact | SCR difference of 62 points cannot be explained with assessment rate difference of 0.2% | **H0 rejected** |
**Conclusion**: **Regional differences in anesthesia-related medical practices cannot be explained by differences in assessment alone. True treatment variations (clinical regional differences) exist. **

### 6.2 Candidate factors that cause regional differences

Contribution of factors estimated from sensitivity analysis results:

| Factor | Estimated contribution | Evidence |
|---|---|---|
| **Factors unique to secondary medical care areas** (facility characteristics, physician training history, university medical office policies) | Maximum | Predominantly distributed within the prefecture (58-86%) |
| **Prefectural-level factors** (including assessment policy) | 14-46% | Contribution rate of inter-prefectural variance |
| **Assessment policy (narrow sense)** | <1% | SCR impact estimation of assessment rate difference |
| **Alternative relationship of anesthesia method** | Significant | Whole body-spine r=-0.506 |
| **Pain clinic activities (spillover)** | Limited | r=0.15 approximately |

### 6.3 Verification of effectiveness of university hospital (medical office)

→ **See Chapter 7 for details**

---

## 7. Verification of university hospital (medical office) effect

### 7.1 Overview and method

We mapped 81 university hospitals (44 national, 8 public, and 29 private) into 64 secondary medical care areas, and compared the SCR patterns of areas where university hospitals are located versus areas where they are not located. The mapping is based on the location of each university (city, ward, town, village) and the correspondence of the secondary medical care area.
- **University hospital area**: 64 areas (19.1% of 335 areas)
- **Non-located area**: 271 areas (80.9%)
- **Area with multiple universities**: 13 areas (5 in central Tokyo, 3 in western ward, 2 in Mishima, Osaka, etc.)

> Note: Japan has at least one medical school in all 47 prefectures due to its "one prefecture, one medical school" policy. Therefore, there is no prefecture without a university hospital, and the essential analysis method is **intra-prefecture comparison** rather than inter-prefecture comparison.

### 7.2 Overall comparison: Areas where university hospitals are located vs areas where they are not located

| Code | Name | University area mean (n) | Non-university area mean (n) | Difference | Cohen's d | t value |
|---|---|---:|---:|---:|---:|---:|
| L008 | General anesthesia | 130.2 (64) | 67.7 (270) | **+62.4** | **+1.780** | 9.72 |
| L001 | Intravenous anesthesia | 137.4 (64) | 56.3 (263) | +81.1 | +1.238 | 7.01 |
| L009 | Anesthesia management fee 1 | 124.8 (64) | 77.6 (250) | +47.2 | +1.021 | 6.96 |
| L002 | Epidural | 127.4 (64) | 87.2 (243) | +40.2 | +0.491 | 3.75 |
| L100 | Nerve block (hospitalization) | 127.1 (64) | 93.3 (271) | +33.8 | +0.444 | 3.14 |
| L100 | Nerve block (outpatient) | 115.7 (64) | 82.1 (271) | +33.6 | +0.525 | 3.59 |
| L004 | Spinal anesthesia | 108.6 (64) | 90.3 (270) | +18.3 | +0.345 | 2.73 |
| L005 | Lower extremity delivery anesthesia | 120.3 (63) | 113.9 (259) | +6.4 | +0.055 | 0.44 |
| L104 | Trigger point (outpatient) | 107.2 (64) | 112.9 (271) | -5.8 | -0.096 | -0.78 |

**Interpretation**:
- **General anesthesia (L008)**: Cohen's d=1.78 is an extremely large effect size. University hospital areas are about 30% higher than the national average, and non-university areas are about 30% lower.
- **Intravenous anesthesia (L001)**: d=1.24. One of the most prominent codes is concentration in the university hospital area.
- **Lower extremity delivery anesthesia (L005) and trigger point (L104)**: No difference. These are determined by local factors in the region, regardless of the presence or absence of university hospitals.

### 7.3 Comparison within the same prefecture (analysis controlling assessment policy)
The prefecture-level assessment policy applies equally to all secondary health care areas. Therefore, by comparing university hospital areas and non-university areas within the same prefecture, it is possible to estimate a ``pure university hospital effect'' that completely eliminates the influence of assessments.

| Code | Name | Difference within prefecture (average) | SD | t value | University area > Non-university area |
|---|---|---:|---:|---:|---|
| L008 | General anesthesia | **+61.4** | 33.7 | **12.49** | **47/47 prefecture (100%)** |
| L001 | Intravenous anesthesia | +84.5 | 75.3 | 7.70 | 42/47 prefectures (89%) |
| L009 | Anesthesia management fee 1 | +46.8 | 41.4 | 7.74 | 40/47 prefecture (85%) |
| L002 | Epidural anesthesia | +51.6 | 72.3 | 4.89 | 39/47 prefectures (83%) |
| L100 | Nerve block (hospitalization) | +43.9 | 64.5 | 4.66 | 34/47 prefectures (72%) |
| L004 | Spinal anesthesia | +20.5 | 37.5 | 3.75 | 31/47 prefecture (66%) |

**Important findings**:
- **General anesthesia (L008)**: SCR in university hospital areas exceeds non-university areas in **all 47 prefectures**. The mean difference +61.4 points, t=12.49, is highly statistically significant. There are no exceptions.
- This effect is **completely independent** of the prefecture's assessment policy. Under the same assessment environment, secondary medical care areas where university hospitals are located consistently have high SCRs for anesthesia-related activities.
- Even for epidural anesthesia (L002), 83% of prefectures are located in university areas. The University Hospital effect has been confirmed for a wide range of procedures including regional anesthesia.

### 7.4 Three-step variance decomposition

Decomposes the total SCR variance into three components: "inter-prefecture", "university hospital effect within prefecture", and "residual":

| Code | Name | Between prefectures | University hospital effect | Residual | University effect as a percentage of within-prefecture variance |
|---|---|---:|---:|---:|---:|
| L008 | General anesthesia | 14.5% | **38.5%** | 47.0% | **45.0%** |
| L001 | Intravenous anesthesia | 21.3% | **32.6%** | 46.1% | 41.5% |
| L009 | Anesthesia management fee 1 | 20.0% | 23.7% | 56.4% | 29.6% |
| L002 | Epidural anesthesia | 22.4% | 14.0% | 63.7% | 18.0% |
| L100 | Nerve block (hospitalization) | 28.9% | 12.7% | 58.4% | 17.8% |
| L004 | Spinal anesthesia | 41.7% | 8.6% | 49.7% | 14.7% |

**Interpretation**:
- **General anesthesia (L008)**: In the previous analysis, the within-prefecture variance was found to be 85.5%, and of that within-prefecture variance, **45.0% is explained by the presence or absence of a university hospital**. The contribution of 38.5% to the total variance is quite large for a single binary variable (with/without a university hospital).
- **Residual 47.0%**: Half of the variance that cannot be explained by the presence or absence of a university hospital remains. This suggests other factors such as individual hospital characteristics, geographic conditions, and population density.
- **Spinal anesthesia (L004)**: The inter-prefectural variance was the largest (41.7%) and the university hospital effect was the smallest (8.6%). Regional differences in spinal anesthesia are more likely to be explained by prefecture-level factors (including local clinical culture and assessment) than by university hospitals.

### 7.5 Detailed pattern of notable prefectures

#### 7.5.1 Kumamoto Prefecture (typical “university hospital concentration” pattern)

| Secondary medical care area | SCR(L008) | University hospital |
|---|---:|---|
| Kumamoto/Kamimashiki | **168.7** | ★ Kumamoto University Hospital |
| Yashiro | 117.3 | - |
| Kuma | 73.8 | - |
| Amakusa | 57.0 | - |
| Kikuchi | 49.7 | - |
| Ariake | 48.6 | - |
| Ashikita | 41.7 | - |
| Shikamoto | 37.3 | - |
| Uki | 7.5 | - |
| Aso | 2.3 | - |

→ University hospital area (168.7) vs non-university area average (48.4) = **difference of 120.3 points**. Typical core concentration pattern.

#### 7.5.2 Tokyo (many university hospitals are distributed)

| Secondary medical care area | SCR(L008) | University hospital |
|---|---:|---|
| Central area | **435.7** | ★ University of Tokyo, Tokyo University of Science, Juntendo, Nippon Medical School, Jikei (5 schools) |
| Western Ward | **168.5** | ★ Keio, Tokyo Medical University, Tokyo Women's Medical University (3 schools) |
| Northwestern part of the ward | 83.4 | ★ Nihon University Itabashi, Teikyo (2 schools) |
| Southern part of the ward | 118.9 | ★ Showa, Toho (2 schools) |
| Southern Kitatama | 121.9 | ★ Kyorin (1 school) |
| Southwest Ward | 110.1 | - |
| Eastern Ward | 97.8 | - |
| Western Kitatama | 88.4 | - |
| Minamitama | 71.1 | - |
| Northeastern part of the ward | 61.8 | - |
| Northern Kitatama | 64.9 | - |
| Nishitama | 62.5 | - |
→ A clear dose-response relationship can be seen between the number of university hospitals and SCR (5 schools: 435.7 → 3 schools: 168.5 → 2 schools: 83-119 → 0 schools: 62-110).

#### 7.5.3 Ibaraki Prefecture (Outlier: Tsukuba area has very high SCR)

| Secondary medical care area | SCR(L008) | University hospital |
|---|---:|---|
| Tsukuba | **181.6** | ★ University of Tsukuba Hospital |
| Mito | 145.5 | - |
| Tsuchiura | 98.4 | - |
| Hitachi | 77.0 | - |
| Toride/Ryugasaki | 73.0 | - |
| Koga/Bando | 63.8 | - |
| Hitachiota/Hitachinaka | 36.0 | - |
| Deer row | 33.8 | - |
| Chikusei/Shimotsuma | 31.5 | - |

→ Tsukuba (181.6) vs non-university average (82.7) = difference of 98.9 points. Mito is also expensive, but there is no university hospital (unique medical cluster).

### 7.6 Conclusion

1. **University hospital effect is highly statistically significant**:
   - General anesthesia (L008) **All 47 prefectures** University area > Non-university area (100% concordance rate)
   - Cohen's d = 1.78 is an effect size classified as "extremely large" in the social sciences.
   - 38.5% of the total variance explained by the presence or absence of a university hospital (single dichotomous variable)
2. **This cannot be explained by assessment**:
   - The assessment policy is controlled because it is a comparison within the same prefecture.
   - t=12.49 cannot be explained by chance.

3. **Strength of effect varies by code**:
   - **Strongest**: General anesthesia (L008, d=1.78), intravenous anesthesia (L001, d=1.24)
   - **Moderate**: Anesthesia management fee (L009, d=1.02), epidural (L002, d=0.49), nerve block (L100, d=0.44-0.53)
   - **None**: Lower limb transmission (L005, d=0.06), trigger point (L104, d=-0.10)

4. **Interpretation notes**:
   - University hospitals are usually located in the largest cities of their prefectures, so the "university hospital effect" and the "urban concentration effect" cannot be completely separated.
   - The high SCR in the university hospital area is thought to be the result of a combination of (a) the high number of anesthesia cases performed at the university hospital itself, (b) the spread to university-affiliated hospitals, and (c) the concentration of hospitals in urban areas.
   - In order to verify the true "medical office effect" (university's academic policies influence the medical treatment patterns of surrounding hospitals), additional analysis is required to compare the presence and absence of university hospitals within the same city size.

---

## 8. Limitations and future prospects

### 8.1 Limitations of this analysis
1. **Insufficient granularity of assessment rate data**: As the assessment rate by medical practice code is not disclosed, the estimation of Test 3 is an indirect estimate based on the overall assessment rate.
2. **Structural limitations of SCR**: SCR itself is calculated from post-assessment data. However, as shown in the sensitivity analysis, the impact of the assessment is quantitatively small.
3. **Anonymity of secondary medical area code**: The secondary medical area in SCR data is only a numeric code, and correspondence with the specific area name (e.g. northwestern part of the ward) is required separately.
4. **Confounding factors**: Regional differences in disease structure (incidence of diseases requiring surgery) cannot be removed by adjusting SCR for age and sex.
5. **Lack of time series analysis**: Only cross-sectional analysis of a single year (R4). Changes over time before and after the revision have not been examined
6. **Unseparated university hospital effect vs. urban concentration effect**: Because university hospitals are usually located in the largest cities within a prefecture, the ``university hospital effect'' and the ``urban concentration effect'' are intertwined. Analysis that controls population size is necessary

### 8.2 Next steps

| Step | Content | Difficulty | Data |
|---|---|---|---|
| A | Separation of university hospital effect and urban concentration effect (comparison of the same population size) | Medium | Demographic statistics + main analysis data |
| B | Correspondence between academic trends in university anesthesiology departments and SCR patterns in the sphere of influence | Medium | Conference presentation + main analysis data |
| C | Comparison of SCR before and after revision in 2020 (impact of L008 name change) | Medium | After R8 data release |
| D | Individual data analysis provided by NDB third party (comparison before and after assessment) | High | Ethics review/NDB application |
| E | Exploration of anesthesia-related cases in examination variance visualization report | Low | Payment Fund Public Materials |

### 8.3 Possibility of development as research

This sensitivity analysis framework can be applied to research on regional differences in any clinical practice, not just anesthesia:
- Selection of surgical procedure (laparoscopic vs. open, etc.)
- Drug prescription pattern
- Inspection order pattern

Variance decomposition (inter-prefecture vs. intra-prefecture) and cross-code correlation are general-purpose methods that can be performed only with NDB/SCR public data.

---

## Appendix: Data sources and access methods used for analysis

### SCR data (Cabinet Office)
````
# By prefecture/medical practice classification
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/01_r04_t_kubun.csv

# Medical practice branch number by prefecture
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/04_r04_t_edaban.csv
# By secondary medical area/medical practice classification
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/02_r04_n_kubun.csv

# Secondary medical care area/medical practice branch number
https://www5.cao.go.jp/keizai-shimon/kaigi/special/reform/mieruka/chiikisa/r04/05_r04_n_edaban.csv
````
Encoding: Shift_JIS

### Examination Statistics (Payment Fund)
````
https://www.ssk.or.jp/tokeijoho/shinsatokei/index.html
````
Annual statistics and monthly data for 2020 can be downloaded in ZIP format.

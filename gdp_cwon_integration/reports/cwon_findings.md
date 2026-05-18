# CWON 統合解析 — flow-stock 統合会計の実証デモンストレータ

## 0. 動機

GDP PoC `reports/poc_findings.md` §9 で「隠れた指標 (μ, β, σ) を同時推定すれば
flow 型国富 (GDP) と stock 型国富 (CWON / IWI) を統合できる」という枠組を
提案した。本解析は、その枠組を **WB CWON データで実証** する最小限のデモ
ンストレータである。

事前仮説（§9 時点）:

1. CWON 公表値 `NW.PCA.TO` は無形資本 (K_intan) を欠いて過小推定。
2. 生産側 (A, D) と stock 側 (CWON) を joint identification すれば、μ・β の
   同定がシャープになる。
3. 高 R&D 国ほど CWON 欠落が大きいはず（R&D intensity と PIM/CWON ratio に
   正の相関）。

以下、**仮説 2 は部分的に確認、仮説 1・3 は反転あるいは修正必要** と判明。
これも flow-stock 統合の枠組にとって重要な情報である。

---

## 1. データと方法

39 か国（OECD + 中国、PoC A/D と同一集合）。CWON は 1995–2020、年次。
各国で：

1. `K*_tang(t; μ̂)` = PIM with geometric lag, mean μ̂ from PoC A.
2. `K*_intan(t; β̂)` = PIM on R&D flows (δ=0.15), stock scaled by β̂ from PoC D.
3. `W_PIM_produced(t) = K*_tang(t) + β̂ · K*_intan(t)`
4. CWON 公表値 `NW.PCA.TO(t)`（produced capital、2019 chained USD）

**評価は二系統**:

- **Within-country trajectory RMSE**: `rmse = stdev( log W_PIM − log NW.PCA.TO )`
  を各国で独立に計算（平均差分を除いた後）。単位・base year・PPP/XRAT の差は
  平均で吸収されるので、**系統的な対数トレンドの差を純粋にとらえる**指標。
- **Joint identification**: `L_total = L_production + 0.3 × L_wealth`
  を (μ, β) 格子で最小化。L_production は成長率残差、L_wealth は上記 RMSE²。

---

## 2. 主な数値結果

### 2.1 水準方向（内国整合性）

| 指標 | 値 |
|---|---|
| trajectory RMSE 中央値 (produced+intangible vs CWON PCA) | **0.049** log 単位 |
| trajectory RMSE 中央値 (tangible のみ) | 0.049 |
| slope of log W_PIM on log NW.PCA.TO（1に近いほど良い）| 中央値 ~0.95 |
| 観測 CWON 年数 中央値 | 25 年 (1995–2020) |

**解釈**: 各国の時系列として、PIM で再構築した (tang + intan) は CWON 発表値と
5% 前後の対数偏差で同期している。**flow 側から再構築した stock と、独立推計の
公表 stock は、水準トレンドでほぼ一致している**。これが「二帳簿が結ばれる」
という仮説のメカニカルな根拠。

### 2.2 Cross-country：R&D intensity との関係

| 指標 | 値 |
|---|---|
| log(PIM / CWON PCA) 中央値（単位調整済）| **+0.65** (比率で ~1.9)|
| log(PIM / CWON PCA) の R&D% への回帰 slope | **−0.49** |
| R² (cross-section) | 0.49 |

**仮説 3 は反転**: 予想に反し **slope が負**。高 R&D 国ほど PIM/CWON 比が
**小さい**。

原因は 2 つの交絡：

1. **PPP vs 市場レート差**: PWT の capital stock (`rnna`) は 2017 自国価格
   ベース、CWON は 2019 市場為替レートベース。開発途上国 (Mexico, Turkey,
   Colombia) では PPP/XRAT >> 1 なので、PWT 値が市場 USD 換算で過大表示。
   これが「低 R&D 国で PIM/CWON が大きい」方向に作用。
2. **SNA 2008 改定**: 2008 年以降、CWON の produced capital は R&D・ソフト
   ウェア・データベースを含む。つまり **CWON PCA にはすでに一部の無形資本が
   含まれている**。高 R&D 国ではこの包含分が大きいので、CWON PCA が大きくなり、
   結果として PIM/CWON 比が小さくなる。

したがって **仮説 1（CWON が無形を丸ごと欠落）は過度に単純化**だった。
より正確には：**CWON は SNA 2008 以降のフローからの無形を含むが、
その前の累積無形 (1995 年以前の R&D ストック) は欠落している**。この差が
PIM による累積ストックとの差に残っている。

### 2.3 Joint identification の効果

| パラメータ | Production-only | Joint (prod + wealth) |
|---|---|---|
| μ 中央値 | 0.40 | 0.26 |
| β 中央値 | **0.01** | **0.06** |

**wealth 制約を加えると β が 6 倍に跳ね上がる**（0.01 → 0.06）。これは：

- Production-only の Test B では β の識別力が弱く、小さな値が選好されやすい。
- Wealth trajectory 制約を加えると、**無形資本の成長パターンが CWON の産業分
  トレンドを fit するために β が上昇**する。
- これは「生産側と stock 側を独立に見ると β は低いが、両者を整合させると真の
  β はより高い（0.05〜0.10）」という joint identification の**識別性向上**の
  具体例。

μ は中央値として joint の方が少し小さくなる。これは wealth 制約が K_tang の
スムーズな上昇を要求するため、ラグ（テンポ）を少し圧縮する方向に効く。

### 2.4 Japan 異例：asset-price gap の定量化

図 C2 で Japan だけ「PIM が CWON を上回る」稀なケース。他の 5 か国
(USA, Germany, Israel, Mexico, Poland) は全部 CWON > PIM。

Japan では 1990 年代後半から不動産価格の持続的下落があり、CWON（市場ベース）
は実物 capital の名目価値を **下方修正**、一方 PIM（物的 flow 積分）は
下落を反映しない。この **10〜15% のギャップ** が 1995–2020 で蓄積。

→ **flow-stock 統合会計で最も興味深いのが、この "asset price writedown" が
stock 側にだけ現れる国**。資産デフレは物的資本の劣化ではないが、CWON では
劣化として記録される。これは名目主義 vs 物量主義の問題で、SEEA や IWI が
別個に扱う課題。

---

## 3. §9 枠組に対する含意

| §9 の主張 | 本解析での検証 | 結論 |
|---|---|---|
| PIM-reconstructed stock と CWON は同じ経済実態を見ている | trajectory RMSE 5% 中央値 | **支持** |
| CWON は無形を欠落、高 R&D 国で過小 | slope 逆符号 | **部分的反証** |
| Joint identification で (μ, β) がシャープになる | β 0.01 → 0.06 | **支持** |
| flow-stock 残差は未観測 stock の下限推定 | 国別に 5–15% の対数残差 | **支持、ただし source は価格改定も含む** |

**修正版の枠組**:

- **仮説 1 の修正**: CWON は SNA 2008 以降の無形を含む。本当に欠落している
  のは (i) 累積 pre-2008 intangibles, (ii) natural capital の一部, (iii)
  social/institutional capital。
- **仮説 3 の修正**: R&D intensity と PIM/CWON 比の関係は PPP/XRAT 交絡に
  支配される。真の intangible test には PWT の `cgdpo` / 現行 PPP 系列、
  または OECD-KLEMS の産業別 intangible を使うべき。
- **Joint identification の価値は確認**: 生産と stock の両方を整合させる
  (μ, β) は、生産単独の推定と系統的に異なり、**flow-stock 残差を最小化する
  方向に動く**（仮説 2 支持）。

---

## 4. 政策的メッセージの更新

§9.5 で書いたのは：

> 「Beyond GDP 運動は半分しか修正していない」

本解析を経て、より正確には：

> **「Beyond GDP (CWON/IWI) は SNA 2008 以降の無形資本を一部捕捉しているが、
> 累積過去ストックと価格改定効果を正しく扱えていない。GDP 側の μ(t) テンポ
> 歪みと CWON 側の価格改定バイアスは異なる方向に効くので、joint identification
> でこそ両者の整合を取れる。」**

特に「Japan 型」（持続的資産デフレで CWON が下方修正、PIM が物量主義で維持）
は単独どちらの指標で見ても誤解を招く。実物資本 (物量) と市場価値 (名目) の
乖離は flow-stock 統合会計における **第四の hidden parameter** と位置づけ
られる。これを加えると枠組は：

```
(μ, β, σ, γ_price) 同時推定
  μ        : 投資→稼働ラグ（テンポ）
  β        : 無形資本シェア
  σ        : 集計バイアス（年齢分散など）
  γ_price  : 物量 vs 市場価値の乖離（asset price writedown / revaluation）
```

γ_price を明示的に分離することで、Japan の 1990〜2000 年代の「失われた」
問題の stock 側への影響と、物的資本蓄積の fundamental 側を分離できる。
これは Hulten-Nakamura-type productivity decomposition 文献との接続点で、
次の自然な拡張。

---

## 5. 図

- `figures/figC1_ratio_by_rnd.png` — log(PIM/CWON) を R&D intensity に回帰。
  slope = −0.49 (R² = 0.49)、PPP/XRAT 交絡の可能性大、要注意。
- `figures/figC2_trajectories.png` — 6 か国 (USA, Japan, Germany, Israel,
  Mexico, Poland) の PIM vs CWON trajectory。Japan だけ PIM > CWON で
  異例。
- `figures/figC3_joint_vs_single.png` — (μ, β) の production-only vs joint
  比較。β が joint で系統的に上方シフト。

## 6. データ

- `data/cwon_integration.csv` — 39 か国の per-country 結果。
- `data/cwon_integration_summary.json` — 集計統計。
- `data/cwon_integration_ts.json` — 時系列（国別、PIM + CWON の 3 系列）。

---

## 7. 次ステップ

1. **PPP 交絡の除去**: PWT の `cn` (current PPPs) または `pl_gdpo` 相対価格
   水準を用いて CWON の market-XRAT 単位に整合。R&D intensity の slope を
   再推定。
2. **Pre-2008 無形ストックの明示的追加**: CHS の `K_intan` を OECD-KLEMS から
   直接取得し、1995 年前の累積分を加えた `CWON_adjusted` を再構築。
3. **γ_price パラメータの導入**: 第四の hidden parameter として、asset-
   price revaluation を同時推定。Japan のケースで検証。
4. **Natural capital の残差ベンチマーク化**: `L_accounting` 残差を SEEA の
   natural capital 推定と比較し、どの程度が natural vs intangible に
   帰属するか定量化。

*Analysis: 2026-04-21. Data: PWT 10.01 (Feenstra, Inklaar & Timmer 2015, rev. 2023);
WB CWON 2021 edition (NW.PCA.TO, NW.HCA.TO, NW.TOW.TO); WB WDI GB.XPD.RSDV.GD.ZS.*

# DRD-CT（動的実績連動型代議制）と先行文献の比較整理

## 論文タイトル案
**"A Dynamic Accountability-Based Representation System: Linking Policy Fulfillment to Electoral Influence"**
（動的説明責任型代議制：公約実現度を選挙影響力に連動させるシステム）

---

## 1. 提案システム（DRD-CT）の要約

ChatGPTでの議論を通じて構築された **DRD-CT（Dynamic Retrospective Delegation with Candidate Trust）** の核心的アイディアは以下の通り：

| 要素 | 内容 |
|------|------|
| **公約の重み付け宣言** | 候補者は立候補時に公約へ重要度点数を配分（合計100点） |
| **実現度の第三者評価** | 任期終了時に各公約の実現度（0〜1）を客観指標・市民評価・専門家評価の混合で算出 |
| **実績得点 S** | S = Σ(重み × 実現度) ∈ [0, 1] |
| **票の影響力関数 ω(S)** | 次回選挙で、前回当該候補に投票した有権者の票の重みを S の関数とする |
| **候補者中心の再定式化** | 個票ではなく「候補者の信任係数」として表現し、一票の平等原則との緊張を緩和 |
| **任期途中評価** | 指数平滑化による連続的評価（O/P/E モデル） |
| **時間責任** | 公約に期限を宣言し、期限内達成率を評価指標に含める |
| **落選候補の扱い** | 「落選＝全否定」ではなく「carry-over trust なし」の中立リセットを推奨 |
| **敵対的評価** | 遺伝的アルゴリズム等で「制度の攻略戦略」を探索し、ロバスト性を検証 |

---

## 2. 先行文献との体系的比較

### 2.1 回顧的投票理論（Retrospective Voting）

**主要文献：**
- Fiorina, M. P. (1981). *Retrospective Voting in American National Elections*. Yale University Press.
- Key, V. O. (1966). *The Responsible Electorate*. Harvard University Press.
- Healy, A. & Malhotra, N. (2013). "Retrospective Voting Reconsidered." *Annual Review of Political Science*, 16, 285–306.

**関係：**
回顧的投票理論は「有権者は過去の実績に基づいて投票行動を決める」という経験的知見を示す。DRD-CTは、この**非公式な行動パターンを制度として内蔵化**する提案である。

| | 回顧的投票理論 | DRD-CT |
|---|---|---|
| 性質 | 記述的（descriptive） | 規範的・設計的（prescriptive） |
| 評価主体 | 有権者個人の主観 | 第三者機関（客観指標＋市民＋専門家） |
| 評価対象 | 経済状況など漠然とした成果 | 事前に宣言された公約の実現度 |
| 制度的帰結 | なし（投票行動に留まる） | 票の重みに反映される制度設計 |

**DRD-CTの新規性：** 回顧的投票を「個人の判断」から「制度的メカニズム」に昇格させた点。Fiorina が記述した行動を、メカニズムデザインの言語で再構成している。

---

### 2.2 選挙的説明責任モデル（Electoral Accountability / Principal-Agent）

**主要文献：**
- Barro, R. (1973). "The control of politicians: an economic model." *Public Choice*, 14, 19–42.
- Ferejohn, J. (1986). "Incumbent Performance and Electoral Control." *Public Choice*, 50, 5–25.
- Besley, T. (2006). *Principled Agents? The Political Economy of Good Government*. Oxford University Press.
- Acemoglu, D., Golosov, M. & Tsyvinski, A. (2008). "Political Economy of Mechanisms." *Econometrica*, 76(3), 619–641.
- Przeworski, A., Stokes, S. C. & Manin, B. (1999). *Democracy, Accountability, and Representation*. Cambridge University Press.

**関係：**
Barro-Ferejohn モデルでは、有権者が「閾値ルール」（パフォーマンスが一定以上なら再選）を用いて政治家を規律付ける。DRD-CTはこの枠組みを拡張し、**二値判断（再選/落選）を連続的な信任係数 ω(S) に置き換えている**。

| | Barro-Ferejohn | DRD-CT |
|---|---|---|
| 規律付け | 再選/落選の二値 | ω(S) ∈ [0,1] の連続値 |
| 公約との関係 | 公約は明示的に扱わない | 公約の重み付け宣言が中核 |
| 評価の粒度 | 任期全体の包括評価 | 公約ごと・期中の連続評価 |
| 理論的基盤 | 繰り返しゲーム | メカニズムデザイン＋進化ゲーム |

**DRD-CTの新規性：** 選挙的説明責任を「再選か否か」の粗い二値から、公約実現度に応じた**連続的影響力調整**へと精緻化した点。Besley (2006) が指摘した「選挙による規律付けの限界」に対する具体的制度設計を提示。

---

### 2.3 メカニズムデザインと社会的選択理論

**主要文献：**
- Arrow, K. (1951/1963). *Social Choice and Individual Values*. Yale University Press.
- Gibbard, A. (1973). "Manipulation of voting schemes." *Econometrica*, 41, 587–601.
- Satterthwaite, M. (1975). "Strategy-proofness and Arrow's conditions." *Journal of Economic Theory*, 10, 187–217.
- Dasgupta, P. & Maskin, E. (2020). "Strategy-Proofness, Independence of Irrelevant Alternatives, and Majority Rule." *AER: Insights*, 2(4), 459–474.

**関係：**
Arrowの不可能性定理・Gibbard-Satterthwaiteの定理は、順序付け型投票制度の根本的限界を示す。DRD-CTは**候補者の選択方法自体を変えるのではなく、票の重みを動的に調整する次元を追加する**ため、古典的不可能性定理の直接的適用範囲外にあるが、新たな戦略的操作の可能性（公約の過少申告、評価指標のゲーミングなど）が生じる。

**DRD-CTの新規性：** 従来の社会的選択理論が「選好集約の方法」に焦点を当てるのに対し、DRD-CTは**「選好集約の前提条件（票の重み）を動的に変える」**という新しい設計空間を開拓。ただし、この新空間における不可能性・操作可能性の分析が論文の重要な課題となる。

---

### 2.4 Quadratic Voting（二次投票）

**主要文献：**
- Lalley, S. & Weyl, E. G. (2018). "Quadratic Voting: How Mechanism Design Can Radicalize Democracy." *AEA Papers and Proceedings*, 1(1).
- Posner, E. A. & Weyl, E. G. (2015). "Voting Squared: Quadratic Voting in Democratic Politics." *Vanderbilt Law Review*, 68(2), 441.
- Posner, E. A. & Weyl, E. G. (2018). *Radical Markets: Uprooting Capitalism and Democracy for a Just Society*. Princeton University Press.

**関係：**
QVは「選好の強度」を票のコスト（票数の二乗）で表現し、多数派の専制を緩和する。DRD-CTは「過去の投票判断の質」を票の重みで反映する。

| | Quadratic Voting | DRD-CT |
|---|---|---|
| 修正対象 | 選好強度の反映不足 | 説明責任の制度化不足 |
| 票の重みの決定要因 | 支払い意思（事前的） | 前回候補の実績（事後的） |
| 理論的動機 | パレート効率の達成 | 公約の誠実性インセンティブ |
| 「一人一票」との関係 | 形式的に変更（複数票購入可） | 候補者係数として再解釈可能 |
| 操作リスク | 資金力による不平等 | 評価指標のゲーミング |

**DRD-CTの新規性：** QVが「今この瞬間の選好強度」を反映するのに対し、DRD-CTは**「過去の委任判断の結果」を将来に持ち越す時間軸**を導入。両者は補完的であり、同時に導入することも理論上可能。

---

### 2.5 Liquid Democracy（流動的民主制）/ Proxy Voting

**主要文献：**
- Brill, M. et al. (2022). "Liquid Democracy with Ranked Delegations." *AAAI*.
- Christoff, Z. & Grossi, D. (2017). "Binary Voting with Delegable Proxy: An Analysis of Liquid Democracy." *TARK*.
- Kahng, A., Mackenzie, S. & Procaccia, A. D. (2021). "Liquid Democracy: An Algorithmic Perspective." *JAIR*, 70, 1223–1252.
- Valsangiacomo, C. (2021). "Political Representation in Liquid Democracy." *Frontiers in Political Science*.
- Major, G. & Preminger, J. (2023). "Democratising Democracy: Votes-Weighted Representation." *JeDEM*, 15(1), 191–218.

**関係：**
Liquid Democracy は「誰に委任するか」を流動的に変更可能にする。DRD-CTは委任先の変更ではなく、**委任結果の評価を制度化**する。

| | Liquid Democracy | DRD-CT |
|---|---|---|
| 焦点 | 委任の柔軟性 | 委任結果の説明責任 |
| 票の重み変動 | 委任の連鎖で自然発生 | 実績得点で制度的に決定 |
| 代議士の固定性 | 代議士の概念を融解 | 代議士制度を維持・強化 |
| リスク | 投票力の集中・委任サイクル | 評価の操作・固定化 |

**DRD-CTの新規性：** Liquid Democracy が「代議制の柔軟化」を志向するのに対し、DRD-CTは**「代議制の強化（実績による規律付け）」**を志向する。代議士制度を前提とした上での改善である点が実装可能性を高める。Flexible Representative Democracy (FRD)（Springer 2024）との比較も有益。

---

### 2.6 Futarchy（予測市場型統治）

**主要文献：**
- Hanson, R. (2013). "Shall We Vote on Values, But Bet on Beliefs?" *Journal of Political Philosophy*, 21(2), 151–173.
- Buterin, V. (2014). "An Introduction to Futarchy." Ethereum Foundation Blog.

**関係：**
Futarchyは「価値は投票で、信念は予測市場で」という分離原則を提唱する。DRD-CTは「公約は事前宣言で、評価は事後計測で」という分離を行う。

| | Futarchy | DRD-CT |
|---|---|---|
| 分離原則 | 価値 vs 信念 | 公約（事前） vs 実現度（事後） |
| 評価メカニズム | 予測市場 | 第三者評価機関（O/P/E） |
| 適用範囲 | 政策決定そのもの | 代議士の説明責任 |
| 実装前提 | 大規模予測市場 | 公約評価制度 |

**DRD-CTの新規性：** Futarchyが**政策決定プロセスそのもの**を市場に委ねるのに対し、DRD-CTは**既存の代議制の枠内**で説明責任を強化する。実装のハードルが低く、漸進的導入が可能。

---

### 2.7 公約実現度の実証研究

**主要文献：**
- Thomson, R. et al. (2017). "The Fulfillment of Parties' Election Pledges: A Comparative Study." *American Journal of Political Science*, 61(3), 527–542.
- Pétry, F. & Collette, B. (2009). "Measuring How Political Parties Keep Their Promises." In *Do They Walk Like They Talk?*, Springer.
- Naurin, E., Royed, T. J. & Thomson, R. (Eds.). *Party Mandates and Democracy*. University of Michigan Press.
- Bytzek, E. et al. (2024). "Do Election Pledges Matter? The Effects of Broken and Kept Election Pledges on Citizens' Trust in Government." *PVS*, 66(4), 785–804.
- Mellon, J. et al. (2023). "Which Promises Actually Matter? Election Pledge Centrality and Promissory Representation." *Political Studies*, 71(3).

**関係：**
Thomson et al. (2017) は12か国・20,000以上の公約を分析し、**与党の公約実現率は平均67%前後**であることを示した。これはDRD-CTの前提（公約実現度は計測可能）が経験的に支持されることを意味する。

**DRD-CTの新規性：** 既存研究は「公約はどの程度守られているか」を事後的に分析するが、DRD-CTは**「公約実現度を制度的インセンティブに変換する」**という処方箋を提示する。実証研究が記述に留まるところを、制度設計に昇華させている。

---

### 2.8 加重投票制度

**主要文献：**
- Baharad, R., Nitzan, S. & Segal-Halevi, E. (2022). "One person, one weight: when is weighted voting democratic?" *Social Choice and Welfare*, 59, 467–493.
- Hofstee, W. K. B. (2017). "Empowering Representative Voters." *Open Journal of Political Science*, 7, 145–156.
- Macé, A. & Treibich, R. (2019). "Inducing Cooperation through Weighted Voting and Veto Power." Working paper.

**関係：**
加重投票は歴史的に存在する（ローマ民会、EU理事会の特定多数決など）。DRD-CTの票の重み変動は、**実績に基づく動的な加重**であり、静的な加重（人口比など）とは性質が異なる。

**DRD-CTの新規性：** 従来の加重投票が**属性（人口、出資比率など）に基づく静的配分**であるのに対し、DRD-CTは**行動結果（公約実現度）に基づく動的配分**を提案。Baharad et al. (2022) の「最適加重と民主主義の重なり」の知見はDRD-CTの正当化にも援用可能。

---

### 2.9 AI・計算論的民主制

**主要文献：**
- Koster, R. et al. (2022). "Human-centred mechanism design with Democratic AI." *Nature Human Behaviour*, 6, 1398–1407.
- Helbing, D. et al. (2023). "Democracy by Design: Perspectives for Digitally Assisted, Participatory Upgrades of Society." *Journal of Computational Science*, 71, 102061.

**関係：**
DRD-CTの公約評価において「AI＋市民審査ハイブリッド」を提案している点は、計算論的民主制の潮流と合致する。Koster et al. (2022) のDemocratic AI（人間が好む再分配メカニズムをRLで設計）は方法論的に類似。

**DRD-CTの新規性：** Democratic AIが「メカニズムそのものをAIに設計させる」のに対し、DRD-CTは**「人間が設計した制度の中でAIを評価ツールとして使う」**というより保守的（＝実装可能性の高い）立場をとる。

---

### 2.10 エージェントベースモデルによる選挙シミュレーション

**主要文献：**
- Laver, M. (2011). *Party Competition: An Agent-Based Model*. Princeton University Press.
- Mitra, A. (2022). "Agent-based Simulation of District-based Elections." arXiv:2205.14400.
- Tomlinson, K. et al. (2024). "Replicating Electoral Success." arXiv:2402.17109.
- Grimm, V. et al. (2020). "The ODD protocol for describing agent-based models." *JASSS*, 23(2).

**関係：**
DRD-CTの検証手法（エージェントベースシミュレーション＋遺伝的アルゴリズムによる敵対探索）は、計算社会科学の標準的方法論に合致する。Tomlinson et al. (2024) のレプリケーター動力学に基づく候補者位置取りモデルは、DRD-CTの進化動態分析と直接的に関連。

---

## 3. DRD-CTの位置づけ：理論的マッピング

```
                    ┌─────────────────────────────────┐
                    │   社会的選択理論                  │
                    │  Arrow, Gibbard-Satterthwaite    │
                    └──────────┬──────────────────────┘
                               │ 不可能性の制約
           ┌───────────────────┼───────────────────┐
           │                   │                   │
   ┌───────▼───────┐  ┌───────▼───────┐  ┌───────▼───────┐
   │ Quadratic     │  │ Liquid        │  │ Futarchy      │
   │ Voting        │  │ Democracy     │  │               │
   │ (選好強度)     │  │ (委任の柔軟化) │  │ (市場による    │
   │               │  │               │  │  政策選択)     │
   └───────┬───────┘  └───────┬───────┘  └───────┬───────┘
           │                   │                   │
           │    ┌──────────────┼──────────────┐    │
           │    │              │              │    │
           └────┤    DRD-CT    ├──────────────┘    │
                │  (実績連動型  │                    │
                │   信任係数)   ├────────────────────┘
                │              │
                └──────┬───────┘
                       │
        ┌──────────────┼──────────────┐
        │              │              │
┌───────▼───────┐ ┌────▼────┐ ┌──────▼──────┐
│回顧的投票理論  │ │説明責任  │ │公約実現度    │
│ Fiorina       │ │モデル    │ │実証研究      │
│ (記述的基盤)   │ │Barro-   │ │Thomson et al.│
│               │ │Ferejohn │ │ (経験的基盤)  │
└───────────────┘ └─────────┘ └─────────────┘
```

---

## 4. DRD-CTの独自性（先行文献との差分のまとめ）

### 4.1 明確な新規貢献

1. **「回顧的投票の制度化」という発想自体**
   - 既存の回顧的投票研究は記述的。DRD-CTはこれを制度設計に昇華させた最初の体系的提案。

2. **公約の重み付け事前宣言＋事後評価の組み合わせ**
   - 公約実現度の研究は事後的分析に留まる。DRD-CTは事前宣言（重み配分）と事後評価（実現度）を結合して**インセンティブ構造**を生む。

3. **候補者中心の信任係数（Candidate Trust Coefficient）**
   - 「一人一票」原則との緊張を、個票ではなく候補者側の係数として再解釈する手法は先行文献にない。

4. **影響力関数 f(S) を「制度ファミリー」のパラメータとして扱う**
   - 単一の制度ではなく、f(S) の関数族として設計空間全体を分析する方法論は新しい。

5. **落選候補の扱いの理論的整理**
   - 「落選＝信任の不在」であり「否定的実績評価」ではないという区別は、選挙理論において明示的に論じられていない論点。

### 4.2 先行研究との重なり（要差別化）

1. **加重投票との区別**：DRD-CTの加重は「属性」ではなく「行動結果」に基づく。この違いを理論的に明確にする必要あり。
2. **メカニズムデザインの操作耐性**：Gibbard-Satterthwaite型の不可能性がDRD-CTの設計空間でどう変形するか、形式的分析が必要。
3. **QVとの補完性**：DRD-CTとQVは同時導入可能だが、相互作用の分析はまだ行われていない。

---

## 5. 論文執筆に向けた具体的提言

### 5.1 Related Work セクションで必ず言及すべき文献群

| 分野 | 筆頭文献 | DRD-CTとの関係 |
|------|---------|---------------|
| 回顧的投票 | Fiorina (1981); Healy & Malhotra (2013) | 記述的基盤 → 制度化の動機 |
| 選挙的説明責任 | Barro (1973); Ferejohn (1986); Besley (2006) | 二値→連続への拡張 |
| メカニズムデザイン | Arrow (1951); Maskin & Dasgupta (2020) | 不可能性制約の位置づけ |
| Quadratic Voting | Posner & Weyl (2018) | 「一人一票」修正の別アプローチ |
| Liquid Democracy | Kahng et al. (2021); Brill et al. (2022) | 委任の柔軟化 vs 説明責任強化 |
| Futarchy | Hanson (2013) | 事前/事後分離の別形態 |
| 公約実現度 | Thomson et al. (2017) | 評価可能性の経験的根拠 |
| Democratic AI | Koster et al. (2022) | 評価手法の技術的基盤 |
| ABMシミュレーション | Laver (2011); Grimm et al. (2020) | 方法論的基盤（ODD準拠） |

### 5.2 投稿先候補

| ジャーナル | 適合理由 | 強調ポイント |
|-----------|---------|------------|
| **JASSS** (Journal of Artificial Societies and Social Simulation) | ABM＋ゲーム理論、ODD準拠 | シミュレーション結果・敵対探索 |
| **Political Analysis** | 形式理論＋計量的手法 | 数理モデル・補題・候補者係数 |
| **Social Choice and Welfare** | 投票理論・制度設計 | f(S)関数族・不可能性との関係 |
| **Royal Society Open Science** | 学際的・オープンアクセス | 全体像を包括的に提示 |

### 5.3 論文の差別化ポイント（査読者への訴求）

1. **「なぜ今まで誰もやらなかったのか」への回答**：回顧的投票は「有権者がやること」であり「制度がやること」ではなかった。DRD-CTは後者を提案する。
2. **規範的議論を避けない**：「一人一票」原則との緊張を正面から論じ、候補者係数による再解釈で解消する。
3. **攻略可能性を自ら検証している**：制度提案論文で敵対的評価まで行うものは稀であり、査読者の信頼を得やすい。

---

## 6. 注意すべき潜在的弱点と対策

| 弱点 | 先行文献からの批判可能性 | 対策案 |
|------|----------------------|--------|
| 票の重みの不平等 | 憲法的「一人一票」原則への抵触（Baharad et al. 2022 参照） | 候補者係数としての再定式化を明示 |
| 評価指標のゲーミング | Goodhart's Law / Campbell's Law | 多元的評価（O/P/E）＋平滑化で緩和。限界は認める |
| 外的ショックへの対応 | 天災・国際情勢は候補者の責任外 | 外的ショック補正係数の導入 |
| 固定化リスク | 勝者の信任係数が累積的に有利に | f(S)の凹関数設計＋落選時リセットで抑制 |
| 操作耐性の理論的保証 | Gibbard-Satterthwaite の拡張の未検討 | 今後の課題として明示（正直に限界を述べる） |

---

## 参考文献一覧（本レポートで言及したもの）

1. Acemoglu, D., Golosov, M. & Tsyvinski, A. (2008). Political Economy of Mechanisms. *Econometrica*, 76(3), 619–641.
2. Arrow, K. (1951/1963). *Social Choice and Individual Values*. Yale University Press.
3. Baharad, R., Nitzan, S. & Segal-Halevi, E. (2022). One person, one weight. *Social Choice and Welfare*, 59, 467–493.
4. Barro, R. (1973). The control of politicians. *Public Choice*, 14, 19–42.
5. Besley, T. (2006). *Principled Agents?* Oxford University Press.
6. Brill, M. et al. (2022). Liquid Democracy with Ranked Delegations. *AAAI*.
7. Bytzek, E. et al. (2024). Do Election Pledges Matter? *PVS*, 66(4), 785–804.
8. Christoff, Z. & Grossi, D. (2017). Binary Voting with Delegable Proxy. *TARK*.
9. Dasgupta, P. & Maskin, E. (2020). Strategy-Proofness, IIA, and Majority Rule. *AER: Insights*, 2(4), 459–474.
10. Ferejohn, J. (1986). Incumbent Performance and Electoral Control. *Public Choice*, 50, 5–25.
11. Fiorina, M. P. (1981). *Retrospective Voting in American National Elections*. Yale University Press.
12. Gibbard, A. (1973). Manipulation of voting schemes. *Econometrica*, 41, 587–601.
13. Grimm, V. et al. (2020). The ODD protocol. *JASSS*, 23(2).
14. Hanson, R. (2013). Shall We Vote on Values, But Bet on Beliefs? *Journal of Political Philosophy*, 21(2), 151–173.
15. Healy, A. & Malhotra, N. (2013). Retrospective Voting Reconsidered. *Annual Review of Political Science*, 16, 285–306.
16. Helbing, D. et al. (2023). Democracy by Design. *Journal of Computational Science*, 71, 102061.
17. Hofstee, W. K. B. (2017). Empowering Representative Voters. *OJPS*, 7, 145–156.
18. Kahng, A., Mackenzie, S. & Procaccia, A. D. (2021). Liquid Democracy: An Algorithmic Perspective. *JAIR*, 70, 1223–1252.
19. Key, V. O. (1966). *The Responsible Electorate*. Harvard University Press.
20. Koster, R. et al. (2022). Human-centred mechanism design with Democratic AI. *Nature Human Behaviour*, 6, 1398–1407.
21. Lalley, S. & Weyl, E. G. (2018). Quadratic Voting. *AEA Papers and Proceedings*, 1(1).
22. Laver, M. (2011). *Party Competition: An Agent-Based Model*. Princeton University Press.
23. Macé, A. & Treibich, R. (2019). Inducing Cooperation through Weighted Voting and Veto Power. Working paper.
24. Major, G. & Preminger, J. (2023). Democratising Democracy: Votes-Weighted Representation. *JeDEM*, 15(1), 191–218.
25. Mellon, J. et al. (2023). Which Promises Actually Matter? *Political Studies*, 71(3).
26. Mitra, A. (2022). Agent-based Simulation of District-based Elections. arXiv:2205.14400.
27. Pétry, F. & Collette, B. (2009). Measuring How Political Parties Keep Their Promises. In *Do They Walk Like They Talk?*, Springer.
28. Posner, E. A. & Weyl, E. G. (2015). Voting Squared. *Vanderbilt Law Review*, 68(2), 441.
29. Posner, E. A. & Weyl, E. G. (2018). *Radical Markets*. Princeton University Press.
30. Przeworski, A., Stokes, S. C. & Manin, B. (1999). *Democracy, Accountability, and Representation*. Cambridge University Press.
31. Satterthwaite, M. (1975). Strategy-proofness and Arrow's conditions. *JET*, 10, 187–217.
32. Thomson, R. et al. (2017). The Fulfillment of Parties' Election Pledges. *AJPS*, 61(3), 527–542.
33. Tomlinson, K. et al. (2024). Replicating Electoral Success. arXiv:2402.17109.
34. Valsangiacomo, C. (2021). Political Representation in Liquid Democracy. *Frontiers in Political Science*.

# 定説曲線の現代的再検証 — 候補リスト

**目的:** 長年「定説」として引用されてきた曲線関係（非線形フィット）を、現代のデータとモデル選択手法（F検定、AIC/BIC、LOOCV）で再検証し、外れ値依存やサンプルサイズ不足による見かけの非線形性を同定する。

**手法（共通）:**
1. 元論文のデータ or 最新データで線形 vs 二次（or 対数/指数）のネストF検定
2. 外れ値除外の感度分析（Cook's distance上位1-3点を除外して有意性が変わるか）
3. AIC/BIC/LOOCV RMSE比較
4. 元論文のサンプルサイズ vs 現在のデータ量

---

## A. 経済学 (Economics) — 12件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 1 | **Phillips Curve** | 失業率 vs インフレ率 | Phillips 1958 | 1970年代のスタグフレーションで既に崩壊が知られる。非線形性は石油ショック期の数点に依存か |
| 2 | **Laffer Curve** | 税率 vs 税収 | Laffer 1974（ナプキン） | 逆U字は理論上自明だが、頂点の位置は国・時代で大きく異なる。クロスカントリーでの非線形性はデータで支持されるか |
| 3 | **Kuznets Curve** | 所得 vs 不平等（ジニ係数） | Kuznets 1955 | 元データは3か国のみ。現代の100か国超データで逆U字は再現するか |
| 4 | **Environmental Kuznets Curve (EKC)** | 所得 vs 環境汚染 | Grossman & Krueger 1991 | SO₂では成立するがCO₂では成立しない報告多数。汚染指標ごとに外れ値依存度が異なるか |
| 5 | **Beveridge Curve** | 欠員率 vs 失業率 | Beveridge 1944 | COVID後のシフトが注目。構造変化か外れ値か |
| 6 | **Okun's Law** | GDP成長率 vs 失業率変化 | Okun 1962 | 線形として有名だが、非線形拡張の有意性は？ 不況期のみ非対称性を示すか |
| 7 | **Engel Curve** | 所得 vs 食料支出比率 | Engel 1857 | 低所得国の数点が曲率を駆動している可能性 |
| 8 | **J-Curve (貿易)** | 為替減価後の時間 vs 貿易収支 | Magee 1973 | J字型の谷はどの程度ロバストか。特定の通貨危機イベントに依存するか |
| 9 | **Rahn Curve** | 政府支出/GDP比 vs 経済成長率 | Rahn 1996 | 逆U字の頂点は30-40%とされるが、北欧の高支出・高成長が外れ値的に曲率を決定か |
| 10 | **Gravity Model of Trade** | GDP×GDP/距離² vs 貿易量 | Tinbergen 1962 | 距離の指数が-2（二乗則）であることの非線形性。中国の台頭で構造変化か |
| 11 | **Great Gatsby Curve** | 不平等（ジニ） vs 世代間社会移動 | Corak 2013 | 米国の位置が曲率を大きく規定。米国除外でも有意か（Preston Curveと同構造） |
| 12 | **Balassa-Samuelson Effect** | 生産性格差 vs 実質為替レート | Balassa 1964 | 途上国の数点が線形関係の傾きを駆動か |

## B. 公衆衛生・疫学 (Public Health / Epidemiology) — 10件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 13 | **Preston Curve** | 所得 vs 平均寿命 | Preston 1975 | **本研究で実証済み**: 二次項は米国1点に依存（p=0.49 without US） |
| 14 | **Easterlin Paradox** | 所得 vs 主観的幸福度 | Easterlin 1974 | クロスセクションでは正、時系列では平坦。非線形性か閾値効果か |
| 15 | **McKeown Thesis Curve** | 時間 vs 死亡率低下（医療介入以前に低下） | McKeown 1976 | 英国の結核データが中心。他の感染症・他国で再現するか |
| 16 | **Dose-Response (LNT: Linear No-Threshold)** | 放射線量 vs がんリスク | ICRP 1958 | 低線量域の線形外挿は広島・長崎データの高線量域からの外挿。閾値モデル vs LNT |
| 17 | **Fries Compression of Morbidity** | 時間 vs 罹患期間 | Fries 1980 | 寿命延長に伴い罹患期間が圧縮されるという仮説。データでは拡張も圧縮も観察される |
| 18 | **Barker Hypothesis Curve** | 出生体重 vs 成人期疾患リスク | Barker 1990 | U字型（低体重も高体重もリスク増）。U字の対称性は特定コホートに依存か |
| 19 | **Omran Epidemiological Transition** | 経済発展段階 vs 疾病構造 | Omran 1971 | 3段階モデルは線形的移行を仮定。サブサハラアフリカの「二重負荷」は外れ値か新段階か |
| 20 | **Wilkinson Curve** | 所得不平等 vs 健康アウトカム | Wilkinson 1996 | 先進国間の関係は、米国＋英国の位置に大きく依存。北欧除外で変化するか |
| 21 | **BMI-Mortality J-Curve** | BMI vs 全死因死亡率 | Various | J字のナディール（最低死亡率のBMI）は喫煙者・既往症患者の含め方で大きく変動 |
| 22 | **Alcohol-Mortality J-Curve** | 飲酒量 vs 死亡率 | Klatsky 1981等 | 「適度な飲酒は有益」のJ字は、非飲酒者に元飲酒者を含めるコーディングに依存か |

## C. 人口学 (Demography) — 6件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 23 | **Demographic Transition Model** | 発展段階 vs 出生率・死亡率 | Notestein 1945 | 4-5段階モデルのS字は欧州の経験に基づく。アジア・アフリカの「圧縮された移行」は別モデルか |
| 24 | **Bongaarts-Feeney Tempo Adjustment** | MAC変化 vs TFR補正 | Bongaarts & Feeney 1998 | テンポ補正の大きさは数か国の急激なMAC上昇期に依存。安定期にはゼロに近づくか |
| 25 | **Lee-Carter Mortality Model** | 年齢 vs 死亡率改善速度 | Lee & Carter 1992 | 単一の時間トレンド（κ_t）の仮定。COVID後に構造変化が発生したか |
| 26 | **Coale-Trussell Fertility Schedule** | 年齢 vs 婚姻出生率 | Coale & Trussell 1974 | 自然出生力のモデルが現代の避妊普及社会に適用可能か |
| 27 | **Replacement Migration Curve** | 移民数 vs 人口維持 | UN 2000 | 非線形的に増大する必要移民数の推計。出生率回復シナリオで曲率は変わるか |
| 28 | **Second Demographic Transition** | 発展段階 vs 超低出生率 | Van de Kaa 1987 | 「不可逆的低下」の仮定。東アジアの反発的回復（あるか？）で検証 |

## D. 環境科学・生態学 (Environmental Science / Ecology) — 6件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 29 | **Species-Area Curve** | 面積 vs 種数 (S=cA^z) | Arrhenius 1921 | べき乗則の指数z≈0.25は島嶼データに基づく。大陸・都市生態系で同じ指数か |
| 30 | **Hubbert Peak Oil Curve** | 時間 vs 石油生産量 | Hubbert 1956 | ベル型曲線のピーク予測は技術革新（シェール）で外れた。対称性の仮定が問題か |
| 31 | **Keeling Curve (CO₂)** | 時間 vs 大気CO₂濃度 | Keeling 1960 | 曲線自体は実測値だが、将来の外挿（指数的 vs 線形的）の非線形性選択 |
| 32 | **HANPP (人間による一次生産量の占有)** | 経済発展 vs 土地利用強度 | Vitousek 1986 | EKCと類似の逆U字が仮定されるが、データでの支持は限定的 |
| 33 | **Jevons Paradox / Rebound Effect** | 技術効率 vs 資源消費 | Jevons 1865 | 効率改善が消費増につながる非線形性。部分反発 vs 完全反発のしきい値は？ |
| 34 | **Forest Transition Curve** | 経済発展 vs 森林面積 | Mather 1992 | U字型（発展初期に森林減少、後に回復）。ブラジル・インドネシアの位置が曲率を決定か |

## E. 心理学・行動科学 (Psychology / Behavioral Science) — 5件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 35 | **Yerkes-Dodson Curve** | 覚醒度 vs パフォーマンス | Yerkes & Dodson 1908 | 元実験はネズミの電気ショック。逆U字の頂点はタスク難易度に依存し、ヒトでのロバスト性は？ |
| 36 | **Ebbinghaus Forgetting Curve** | 時間 vs 記憶保持率 | Ebbinghaus 1885 | 元データは被験者1名（自分自身）。指数減衰 vs べき乗減衰の区別は可能か |
| 37 | **Weber-Fechner Law** | 刺激強度 vs 知覚強度 | Fechner 1860 | 対数関係 vs Stevens' Power Law。感覚モダリティごとに最適モデルが異なるか |
| 38 | **Dunning-Kruger Effect Curve** | 能力 vs 自己評価 | Kruger & Dunning 1999 | 統計的アーティファクト（回帰平均への回帰）という批判。非線形性は本物か |
| 39 | **Happiness U-Curve (年齢)** | 年齢 vs 幸福度 | Blanchflower & Oswald 2008 | U字の谷（40-50代）はコントロール変数の選択に敏感 |

## F. 物理学・自然科学 (Physics / Natural Sciences) — 4件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 40 | **Hubble's Law** | 距離 vs 後退速度 | Hubble 1929 | 元データは24銀河のみで散布が大きい。非線形項（加速膨張）の検出はSN Iaの数点に依存 |
| 41 | **Kleiber's Law (代謝スケーリング)** | 体重 vs 基礎代謝率 (B∝M^0.75) | Kleiber 1932 | 3/4乗則 vs 2/3乗則の論争。外温動物を含めるか否かで指数が変わる |
| 42 | **Gutenberg-Richter Law** | マグニチュード vs 頻度 (log-linear) | Gutenberg & Richter 1944 | 大地震（M≥8）での逸脱。最大マグニチュードの数点がべき乗則の成立範囲を規定 |
| 43 | **Moore's Law** | 時間 vs トランジスタ密度 | Moore 1965 | 指数関数的成長の「法則」は物理的限界に近づいている。2015年以降の減速は外れ値か構造変化か |

## G. 政治学・社会学 (Political Science / Sociology) — 5件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 44 | **Lipset Hypothesis** | 所得 vs 民主主義度 | Lipset 1959 | 石油産出国（サウジ、UAE）の「高所得・非民主主義」が非線形性を駆動か |
| 45 | **Duverger's Law Curve** | 選挙制度（区定数） vs 政党数 | Duverger 1954 | 小選挙区→二党制の関係はインド・カナダでは成立しない。外れ値の影響は |
| 46 | **Zipf's Law (都市規模)** | 順位 vs 人口 | Zipf 1949 | べき乗則の指数≈1は米国に最も当てはまる。他国では逸脱が大きい |
| 47 | **Crime-Temperature Curve** | 気温 vs 犯罪率 | Anderson 1987 | 非線形性（極端な高温で犯罪が減少）は一部データに依存か |
| 48 | **Putnam Social Capital Curve** | テレビ視聴時間 vs 社会的信頼 | Putnam 2000 | 米国固有の現象か。北欧では成立しない可能性 |

## H. 農学・栄養学 (Agriculture / Nutrition) — 4件

| # | 曲線名 | 関係 | 提唱者・年 | 再検証の仮説 |
|---|--------|------|-----------|-------------|
| 49 | **Mitscherlich Yield Curve** | 肥料投入量 vs 作物収量 | Mitscherlich 1909 | 逓減リターンの曲率は土壌タイプの数点で決まるか |
| 50 | **Borlaug Green Revolution Curve** | 時間 vs 穀物収量 | Borlaug 1960s | S字型の成長曲線は飽和しつつある。アフリカは別の曲線上にあるか |
| 51 | **Micronutrient Dose-Response (U-shaped)** | 摂取量 vs 健康リスク（ビタミンD等） | Various | U字（不足も過剰もリスク）の右側上昇はサプリメント大量摂取の少数例に依存か |
| 52 | **Body Weight Set-Point Curve** | カロリー摂取 vs 体重変化 | Keesey & Hirvonen 1997 | 非線形的な体重恒常性のモデル。カロリー制限の極端な例に依存か |

---

## 横断的パターン（予備的考察）

Preston Curveの再検証から得られた教訓を一般化すると：

1. **外れ値駆動型**: 非線形項の有意性が1-3点の外れ値に依存（Preston, Great Gatsby, Lipset, Zipf）
2. **時代依存型**: 元データの時代には成立したが、構造変化で成立しなくなった（Phillips, Moore, Hubbert）
3. **サンプル不足型**: 元論文のN<30で、現代のN>100では非線形性が消失（Kuznets, Ebbinghaus, Hubble）
4. **変数定義依存型**: 非線形性が変数のコーディングや測定法に依存（Alcohol J-Curve, Dunning-Kruger）
5. **地理的限定型**: 特定の国・地域でのみ成立し、他では成立しない（Zipf→米国, EKC→SO₂のみ）

---

## 推奨される優先順位（再検証の実現可能性）

| 優先度 | 候補 | 理由 |
|--------|------|------|
| ★★★ | #3 Kuznets, #4 EKC, #11 Great Gatsby, #14 Easterlin | クロスカントリーデータがOECD/WBで公開。F検定が直接適用可能 |
| ★★★ | #1 Phillips, #6 Okun | マクロ経済時系列が豊富。構造変化の検定も可能 |
| ★★☆ | #22 Alcohol J-Curve, #21 BMI J-Curve | メタ分析データが利用可能。コーディング依存性を検証可能 |
| ★★☆ | #35 Yerkes-Dodson, #36 Ebbinghaus, #38 Dunning-Kruger | 実験の再現性危機と接続。再現データが増加中 |
| ★☆☆ | #40 Hubble, #41 Kleiber | 専門知識が必要だが、データは公開されている |

---

*本リストは「医療費の経済効果」研究（PR#56）におけるPreston Curve過剰適合の発見を契機として作成。*
*Onishi T. 2026.*

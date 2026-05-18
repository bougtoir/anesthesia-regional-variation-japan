# Spectral Causality Brainstorm — 医療データへのGEM-RAG的因果性アプローチ

GEM-RAG（Graphical Eigen Memories for Retrieval Augmented Generation）のコアアイデア（ユーティリティ質問 → グラフ → スペクトル分解）を医療データの因果推論に応用するブレインストーミング文書群。

## 文書構成

| # | ファイル | 内容 |
|---|---|---|
| 01 | `01_gem_rag_medical_application.md` | GEM-RAGの医療データ応用（横断的スナップショット + 経時データ）。「ユーティリティ因果性」の基本提案。 |
| 02 | `02_lingam_x_utility_causality.md` | LiNGAMファミリーとの比較・統合アイデア。変数レベル因果 × テーマレベル因果の二段ロケット等。 |
| 03 | `03_spectral_causality_deep_dive.md` | スペクトル因果性の深掘り。磁気ラプラシアン、Hodge分解、因果フーリエ解析との接続。数学的定式化。 |
| 04 | `04_hill_criteria_mapping.md` | Hill の9基準 × 因果推論手法（40+手法）の包括的マッピング。既存手法のH6/H7/H9空白地帯と提案手法の位置づけ。 |
| 05 | `05_spectral_causality_explainer.md` | **専門外向け解説**。音のたとえ・水流のたとえで概念を説明。UCI心疾患データ（297名）での実解析結果付き。因果の梯子・Hill基準の統合的議論。 |
| 06 | `06_spectral_causality_academic.md` | **学術的解説**（大学生〜院生向け）。LaTeX数式、定義・命題・証明スタイル。磁気ラプラシアン・SCC/SCD・Hodge分解の厳密な定式化。CCIによる統一的理解。 |
| — | `demo_spectral_causality.py` | 解析再現用Pythonスクリプト。LiNGAM・磁気ラプラシアン・Hodge分解・SCI/SCD計算・図生成を一括実行。 |
| — | `figures/` | 解析結果の図（fig1〜fig5）。 |

## 起点記事

- [GEM-RAGが拓く「グラフ×スペクトル」な次世代RAG](https://zenn.dev/lluminai_tech/articles/cc4b62b47936b3)

## キーコンセプト

- **Utility Causality（ユーティリティ因果性）**: 「同じ臨床問いに答える能力」の時間的連続として定義される因果性
- **Spectral Causal Coupling (SCC)**: cosによる対称的因果結合度。**Spectral Causal Direction (SCD)**: sinによる反対称的因果方向
- **Complex Causal Index (CCI)**: SCC + i·SCD。複素平面上で因果の強さと方向を統一的に表現
- **Hodge-Causal Decomposition**: エッジフローの勾配成分 = 因果的フロー、カール成分 = フィードバック
- **Ensemble Causal Direction (ECD)**: LiNGAM + SCD + Granger のアンサンブル因果方向推定

## 関連研究

- Kotoku et al. (2020) — 大阪府健診データ × DirectLiNGAM
- Okuda et al. (2025) — 日本健診コホート10万人超 × Longitudinal LiNGAM
- M'Charrak et al. (2025) — スペクトル正則化 × 因果DAG学習
- Seifert, Wendler & Püschel (2023) — DAG上の因果フーリエ解析
- Maehara & Ohkawa (2019) — 単一細胞データ × Hodge分解

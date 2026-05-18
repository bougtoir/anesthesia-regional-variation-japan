# 琵琶湖考古学的探査パイプライン
# Lake Biwa Archaeological Prospection Pipeline

地形データ（DEM）から遺跡候補地を自動検出するための解析パイプライン。
国土地理院の公開タイルデータを使用し、複数の地形可視化手法と統計的異常検出を適用する。

## 背景

### 現琵琶湖の湖底
琵琶湖の水位変動により水没した集落遺跡が存在する（粟津湖底遺跡など）。
湖底の微地形異常を検出することで、未知の水没遺跡を発見できる可能性がある。

### 古琵琶湖の湖畔域
琵琶湖は約400万年前に三重県伊賀地方で誕生し、北へ移動して現在の位置に至った。
各ステージの湖畔域は当時の水辺の生活圏であり、遺跡が存在する蓋然性が高い。

| ステージ | 累層名 | 年代 (Ma) | 位置 |
|---------|--------|-----------|------|
| 大山田湖 | 上野累層 | 4.0–3.5 | 三重県伊賀市大山田 |
| 阿山湖 | 阿山累層 | 3.0–2.6 | 三重県伊賀市〜滋賀県甲賀市 |
| 甲賀湖 | 甲賀累層 | 2.6–2.0 | 滋賀県甲賀市 |
| 蒲生沼沢地 | 蒲生累層 | 2.6–1.0 | 滋賀県蒲生郡 |
| 堅田湖 | 堅田累層 | 1.0–0.4 | 滋賀県大津市堅田 |

## 使用技術

### データ取得
- 国土地理院 標高タイル (DEM 5m/10m)
- 国土地理院 湖水深タイル
- 産総研 シームレス地質図V2

### 地形可視化
| 手法 | 略称 | 目的 |
|------|------|------|
| 赤色立体地図 | RRIM | 微地形の総合的可視化 (千葉ほか, 2008) |
| 天空率 | SVF | 凹凸の方向非依存検出 |
| 傾斜量図 | Slope | 急傾斜部の強調 |
| 局所起伏モデル | LRM | 大域トレンド除去後の局所起伏 |
| 多方向陰影 | Multi-HS | 微地形の陰影強調 |

### 異常検出
- Z-score閾値法: 局所統計からの偏差を検出
- 円形構造スコア: 古墳・環濠等の円形パターン検出
- 線形構造スコア: 道路・堤防等の線形パターン検出
- マルチパラメータスコアリング: 複数特徴の統合評価

## 実行方法

```bash
# 依存パッケージ
pip install numpy scipy matplotlib pandas requests Pillow

# パイプライン実行
cd /path/to/wip
python -m biwa_archaeological_prospection.run --output biwa_archaeological_prospection/output --verbose
```

### オプション
- `--output`, `-o`: 出力ディレクトリ（デフォルト: `biwa_archaeological_prospection/output`）
- `--zoom`, `-z`: DEMタイルのズームレベル（デフォルト: 14, ~8m解像度）
- `--verbose`, `-v`: 詳細ログ出力

## 出力
- `paleo_biwa_overview.png`: 古琵琶湖変遷の概観図
- `analysis_*.png`: 各地域の解析結果（12パネル: DEM, Hillshade, Slope, SVF, LRM, RRIM, 異常スコア, 候補地マップ）
- `report.md`: Markdown形式の調査レポート

## 到達可能な精度レベル

| 段階 | 解像度 | 本パイプラインでの到達可能性 |
|------|--------|--------------------------|
| 広域スクリーニング | 5-30m | 完全に可能 |
| 詳細地形解析 | 0.5-2m | UAV-LiDARデータがあれば可能 |
| 地表面精密調査 | cm級 | GPR等の別手法が必要 |

## モジュール構成

```
biwa_archaeological_prospection/
├── __init__.py          # パッケージ定義
├── data_fetch.py        # 国土地理院タイルデータ取得
├── terrain_viz.py       # 地形可視化（RRIM, SVF, Hillshade等）
├── anomaly_detection.py # 統計的異常検出
├── paleo_biwa.py        # 古琵琶湖変遷データ
├── pipeline.py          # メインパイプライン
├── run.py               # 実行スクリプト
├── output/              # 出力ディレクトリ
└── README.md            # 本ファイル
```

## 参考文献

1. Chiba T., Kaneta S. & Suzuki Y. (2008) Red Relief Image Map: New visualisation for three dimensional data. ISPRS.
2. 里口保文 (2025) 琵琶湖博物館研究調査報告 38巻.
3. Štular B., Lozić E. & Eichert S. (2021) Airborne LiDAR-Derived DEM for Archaeology. Remote Sensing, 13(9), 1855.
4. Zakšek K., Oštir K. & Kokalj Ž. (2011) Sky-View Factor as a Relief Visualization Technique. Remote Sensing, 3(2), 398-415.
5. 横山卓雄・雨森清 (1991) 滋賀県湖東地域古琵琶湖層群地質図.

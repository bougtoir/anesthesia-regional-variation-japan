# PWV-VitalDB Analysis: Pulse Wave Velocity as ICU Monitoring Surrogate

VitalDB Open Datasetを用いた波形解析によるPWV（脈波伝播速度）の算出と、
侵襲的モニタリング指標の代替可能性およびSOFAスコアベース死亡予測モデルへの統合を検証する。

## 概要

### 研究課題
1. **波形解析手法の確立**: ECG・動脈圧・PPG波形からPWV関連指標（PAT, PTT, 2-site PWV）を連続的に算出
2. **侵襲的モニタリングの代替**: PWVがMAP, CVP等の侵襲的指標の代替となるかの検証
3. **SOFA予測モデルの強化**: SOFAスコアにPWVを追加/差し替えすることで予測精度が向上するかの検証

### 主な知見
- **PAT-MAP相関**: 個体間相関は弱い(r=-0.08)が、**個体内相関は中等度〜強い**(median r=-0.52)
- **PTT-MAP相関**: 有意な正の相関(r=0.19, p=0.006)
- **SOFA+PWVモデル**: SOFA-raw (AUROC 0.60) → SOFA+PWV (AUROC 0.60)、NRI +0.05
- **臨床モデル**: 年齢・ASA・緊急度等の臨床因子が最も強い予測因子

## パイプライン構成

```
scripts/
├── run_pipeline.py           # メインランナー
├── s01_data_acquisition.py   # VitalDBデータ取得
├── s02_pwv_calculation.py    # PWV算出（PAT, PTT, 2-site PWV）
├── s03_correlation_analysis.py  # 相関分析
├── s04_sofa_prediction_model.py # 予測モデル比較
└── s05_generate_report.py    # レポート生成（.docx）
```

## 実行方法

```bash
# 依存パッケージ
pip install vitaldb pyvital scikit-learn xgboost lightgbm pyarrow python-docx

# フルパイプライン実行
cd pwv_vitaldb_analysis
python scripts/run_pipeline.py

# クイックモード（少数ケースでテスト）
python scripts/run_pipeline.py --quick

# 特定ステップのみ
python scripts/run_pipeline.py --step 2  # PWV計算のみ
```

## データソース

[VitalDB Open Dataset](https://vitaldb.net/dataset/) - 6,388手術症例、500Hz波形データ

### 使用トラック
| Track | Description | Rate |
|-------|------------|------|
| SNUADC/ECG_II | ECG Lead II | 500Hz |
| SNUADC/ART | 橈骨動脈圧波形 | 500Hz |
| SNUADC/PLETH | PPG波形 | 500Hz |
| SNUADC/CVP | 中心静脈圧波形 | 500Hz |
| SNUADC/FEM | 大腿動脈圧波形 | 500Hz |

## PWV算出方法

1. **PAT (Pulse Arrival Time)**: ECG R-peak → ABP foot（PEPを含む）
2. **PTT (Pulse Transit Time)**: ABP foot → PPG foot（純粋な動脈伝播時間）
3. **2-site PWV**: FEM foot → ART foot / 推定距離（真のPWV）

## 参考文献

- Lee HC et al. VitalDB, a high-fidelity multi-parameter vital signs database. Sci Data. 2022;9(1):279.
- Vincent JL et al. The SOFA score. Intensive Care Med. 1996;22(7):707-710.
- Schaanning SG, Skjaervold NK. PLoS One. 2020;15(10):e0240126.

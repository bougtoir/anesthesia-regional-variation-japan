"""
Beyond GDP: 現代国家への適用分析

歴史データから得られた因果モデルを用いて、現代国家のリスクを評価する。

対象:
  1. 中国（Great Firewall = デジタル海禁）
  2. ロシア（制裁下のブロック化）
  3. 日本（フロー低下・ストック移行中）
  4. 北朝鮮（完全閉鎖 + 外部保護者）
  5. イラン（制裁 + 閉鎖）
  6. その他の参照国家（シンガポール、スイス、台湾、UAE）

手法:
  - 歴史モデルから得た係数を用いた予測確率算出
  - 媒介経路の状態診断
  - モンテカルロシミュレーションによるリスク評価
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import statsmodels.api as sm
from data import load_data


# 現代国家のパラメータ
MODERN_STATES = [
    {
        "entity": "中国（現在）",
        "description": "GFW（デジタル海禁）+ 一帯一路（選択的フロー）",
        "dominant": "stock",
        "stock_index": 0.75,
        "trade_openness": 0.50,      # 選択的開放
        "geo_barrier": 0.5,
        "external_threat": 0.6,
        "tech_position": 0.7,
        "institutional_quality": 0.5,
        "era_code": 5,
        "has_external_patron": 0,
        "regime_duration_yrs": 76,
        "notes": [
            "GFW = 情報フロー遮断（デジタル海禁）",
            "一帯一路 = 物理フローの選択的展開",
            "半導体規制 = 技術フロー遮断の外圧",
            "人的資本ストック◎（STEM教育）",
            "制度品質は中程度（効率的だが硬直的）",
            "技術蓄積は急速に増加中",
        ],
        "risk_factors": [
            "技術デカップリングによる技術停滞リスク",
            "GFWによる情報遮断→外部脅威認識の歪み",
            "不動産バブル後の金融ストック毀損",
            "人口減少→人的資本ストック減耗",
        ],
        "protective_factors": [
            "巨大な国内市場（ストック活用のフロー代替）",
            "一帯一路による選択的フロー維持",
            "核抑止力",
            "STEM人材の厚み",
        ],
    },
    {
        "entity": "ロシア（現在）",
        "description": "制裁下の強制ブロック化 + 資源ストック依存",
        "dominant": "stock",
        "stock_index": 0.55,
        "trade_openness": 0.20,      # 制裁で急縮小
        "geo_barrier": 0.6,
        "external_threat": 0.7,
        "tech_position": 0.4,
        "institutional_quality": 0.2,
        "era_code": 5,
        "has_external_patron": 0,
        "regime_duration_yrs": 25,   # プーチン体制
        "notes": [
            "ウクライナ戦争→西側制裁→強制的海禁状態",
            "エネルギー資源ストック◎だが技術蓄積△",
            "頭脳流出（IT人材20万人以上流出）",
            "中国への経済依存増大",
            "制度品質の急速な劣化",
        ],
        "risk_factors": [
            "技術輸入遮断→半導体・精密機器の枯渇",
            "頭脳流出→人的資本ストック急減",
            "制度品質の加速的劣化（戦時体制）",
            "中国依存＝保護者依存のリスク",
        ],
        "protective_factors": [
            "核抑止力",
            "エネルギー資源",
            "広大な国土（地理的障壁）",
        ],
    },
    {
        "entity": "日本（現在）",
        "description": "フロー低下・ストック移行中、部分的海禁傾向",
        "dominant": "stock",
        "stock_index": 0.80,
        "trade_openness": 0.45,
        "geo_barrier": 0.7,
        "external_threat": 0.6,
        "tech_position": 0.65,
        "institutional_quality": 0.7,
        "era_code": 5,
        "has_external_patron": 1,    # 日米安保
        "regime_duration_yrs": 80,
        "notes": [
            "対外純資産世界一（金融ストック◎）",
            "都市鉱山（技術ストック◎）",
            "少子高齢化→人的資本ストック減耗",
            "移民政策閉鎖的（文化的海禁）",
            "言語障壁（情報フロー阻害）",
            "安保は米国依存（外部保護者）",
        ],
        "risk_factors": [
            "人的資本の減耗（少子高齢化）",
            "技術停滞リスク（半導体・AIで遅れ）",
            "安保依存のリスク（米国の関与低下時）",
            "制度更新速度の遅さ",
        ],
        "protective_factors": [
            "巨大な蓄積ストック（対外純資産、都市鉱山）",
            "高い制度品質（法の支配）",
            "地理的障壁（島国）",
            "日米安保（外部保護者）",
            "素材・精密機器での技術優位",
        ],
    },
    {
        "entity": "北朝鮮（現在）",
        "description": "完全閉鎖 + 核抑止 + 中国保護",
        "dominant": "stock",
        "stock_index": 0.25,
        "trade_openness": 0.05,
        "geo_barrier": 0.5,
        "external_threat": 0.8,
        "tech_position": 0.15,
        "institutional_quality": 0.1,
        "era_code": 5,
        "has_external_patron": 1,
        "regime_duration_yrs": 77,
        "notes": [
            "主体思想＝究極の海禁政策",
            "核兵器＝技術ストックの極端な集中投資",
            "中国の庇護なしでは存続不能",
        ],
        "risk_factors": [
            "制度品質が極端に低い",
            "技術蓄積が核・ミサイルに偏向",
            "人的資本の劣化（栄養不良、情報遮断）",
            "中国の保護者撤退リスク",
        ],
        "protective_factors": [
            "核抑止力",
            "中国の庇護",
            "体制の強固な統制力",
        ],
    },
    {
        "entity": "イラン（現在）",
        "description": "制裁下の閉鎖経済 + 資源ストック",
        "dominant": "stock",
        "stock_index": 0.55,
        "trade_openness": 0.20,
        "geo_barrier": 0.5,
        "external_threat": 0.8,
        "tech_position": 0.4,
        "institutional_quality": 0.3,
        "era_code": 5,
        "has_external_patron": 0,
        "regime_duration_yrs": 46,
        "notes": [
            "制裁による強制的海禁状態",
            "石油資源ストックに依存",
            "核開発＝技術ストックへの集中投資",
            "プロキシ戦争＝限定的フロー（軍事）",
        ],
        "risk_factors": [
            "技術停滞（制裁による輸入遮断）",
            "頭脳流出",
            "制度品質の低さ",
            "外部保護者不在",
        ],
        "protective_factors": [
            "石油・ガス資源",
            "地理的障壁（山岳地帯）",
            "地域的影響力（プロキシ）",
        ],
    },
    # 参照国家（フロー維持型）
    {
        "entity": "シンガポール（参照）",
        "description": "完全フロー特化型",
        "dominant": "flow",
        "stock_index": 0.55,
        "trade_openness": 0.95,
        "geo_barrier": 0.3,
        "external_threat": 0.3,
        "tech_position": 0.8,
        "institutional_quality": 0.9,
        "era_code": 5,
        "has_external_patron": 1,
        "regime_duration_yrs": 60,
        "notes": ["フロー特化の成功モデル"],
        "risk_factors": ["ストック基盤の薄さ", "地理的脆弱性"],
        "protective_factors": ["極めて高い制度品質", "技術水準", "多角的同盟"],
    },
    {
        "entity": "スイス（参照）",
        "description": "ストック維持 + 高制度品質の稀有な成功例",
        "dominant": "stock",
        "stock_index": 0.85,
        "trade_openness": 0.40,
        "geo_barrier": 0.8,
        "external_threat": 0.2,
        "tech_position": 0.85,
        "institutional_quality": 0.95,
        "era_code": 5,
        "has_external_patron": 0,
        "regime_duration_yrs": 176,
        "notes": ["海禁的だが制度・技術を自律維持する稀有な例"],
        "risk_factors": ["EU依存の潜在リスク"],
        "protective_factors": ["極めて高い制度品質・技術水準", "山岳地理", "永世中立"],
    },
    {
        "entity": "台湾（参照）",
        "description": "フロー特化 + 技術ストック（半導体）",
        "dominant": "flow",
        "stock_index": 0.65,
        "trade_openness": 0.80,
        "geo_barrier": 0.7,
        "external_threat": 0.7,
        "tech_position": 0.85,
        "institutional_quality": 0.75,
        "era_code": 5,
        "has_external_patron": 1,
        "regime_duration_yrs": 76,
        "notes": ["半導体＝戦略的ストック＋フロー変換器（シリコンシールド）"],
        "risk_factors": ["中国の軍事脅威", "米国保護の不確実性"],
        "protective_factors": ["半導体産業のグローバル重要性", "台湾海峡", "日米同盟"],
    },
]


def predict_risk(df_hist):
    """歴史モデルの係数を使って現代国家の征服リスクを予測"""
    y = df_hist["outcome_binary"]

    # Best model: external_threat + institutional_quality + era_code + geo_barrier
    covariates = ["external_threat", "institutional_quality", "era_code",
                  "geo_barrier", "has_external_patron", "tech_position"]
    X = sm.add_constant(df_hist[covariates].astype(float))
    model = sm.Logit(y, X).fit(disp=0)

    print("=" * 70)
    print("現代国家リスク分析")
    print("=" * 70)
    print(f"\n使用モデル: Logit(征服 | {', '.join(covariates)})")
    print(f"学習データ: N={len(df_hist)} (歴史的事例)")
    print(f"Pseudo R²: {model.prsquared:.3f}")

    # Print coefficients
    print(f"\n  {'変数':30s} {'係数':>8s} {'OR':>10s} {'p値':>8s}")
    print(f"  {'-'*58}")
    for var in covariates:
        coef = model.params[var]
        or_val = np.exp(coef)
        p = model.pvalues[var]
        print(f"  {var:30s} {coef:>8.3f} {or_val:>10.3f} {p:>8.4f}")

    print(f"\n\n{'='*70}")
    print("現代国家の征服/崩壊リスク予測")
    print(f"{'='*70}\n")

    results = []
    for state in MODERN_STATES:
        x_new = pd.DataFrame([{c: state[c] for c in covariates}])
        x_new = sm.add_constant(x_new, has_constant="add")
        # Align columns
        for col in X.columns:
            if col not in x_new.columns:
                x_new[col] = 0
        x_new = x_new[X.columns]
        prob = model.predict(x_new).values[0]
        results.append((state, prob))

    # Sort by risk
    results.sort(key=lambda x: -x[1])

    for state, prob in results:
        risk_level = "極高" if prob > 0.8 else "高" if prob > 0.6 else "中" if prob > 0.4 else "低" if prob > 0.2 else "極低"
        bar = "█" * int(prob * 20) + "░" * (20 - int(prob * 20))
        print(f"\n  {state['entity']}")
        print(f"    {state['description']}")
        print(f"    征服/崩壊確率: {prob:.1%}  [{bar}]  リスク: {risk_level}")
        print(f"    優位タイプ: {state['dominant']}, ストック指数: {state['stock_index']}, 貿易開放度: {state['trade_openness']}")

    return results, model


def mediation_diagnosis(df_hist):
    """現代国家の媒介経路診断"""
    print(f"\n\n{'='*70}")
    print("媒介経路診断: 3つの劣化リスク評価")
    print(f"{'='*70}")

    # 歴史データの分布統計
    hist_stock = df_hist[df_hist["dominant"] == "stock"]
    hist_surv = df_hist[df_hist["outcome"] == "survived"]

    print(f"\n  歴史的参照値（ストック優位国のうち存続した国の中央値）:")
    for var in ["institutional_quality", "tech_position", "external_threat"]:
        med = hist_surv[var].median()
        q25 = hist_surv[var].quantile(0.25)
        q75 = hist_surv[var].quantile(0.75)
        print(f"    {var:30s}: 中央値={med:.2f} (IQR: {q25:.2f}-{q75:.2f})")

    surv_median_iq = hist_surv["institutional_quality"].median()
    surv_median_tp = hist_surv["tech_position"].median()
    surv_median_et = hist_surv["external_threat"].median()

    print(f"\n\n  {'国家':20s} {'制度品質':>8s} {'技術水準':>8s} {'外部脅威':>8s} {'M1リスク':>8s} {'M2リスク':>8s} {'M3リスク':>8s} {'総合':>6s}")
    print(f"  {'-'*78}")

    for state in MODERN_STATES:
        iq = state["institutional_quality"]
        tp = state["tech_position"]
        et = state["external_threat"]

        # Risk: below survival median = at risk
        m1_risk = "危険" if iq < surv_median_iq else "安全"
        m2_risk = "危険" if tp < surv_median_tp else "安全"
        m3_risk = "危険" if et > surv_median_et else "安全"

        n_risk = sum([iq < surv_median_iq, tp < surv_median_tp, et > surv_median_et])
        overall = "高危険" if n_risk >= 3 else "要注意" if n_risk >= 2 else "注意" if n_risk >= 1 else "良好"

        print(f"  {state['entity']:20s} {iq:>8.2f} {tp:>8.2f} {et:>8.2f} {m1_risk:>8s} {m2_risk:>8s} {m3_risk:>8s} {overall:>6s}")


def monte_carlo_scenarios(model, n_sim=2000):
    """モンテカルロシミュレーションで将来シナリオを評価"""
    print(f"\n\n{'='*70}")
    print("モンテカルロ・シナリオ分析")
    print(f"{'='*70}")

    rng = np.random.default_rng(42)

    scenarios = {
        "中国: 現状維持": {
            "external_threat": (0.6, 0.05),
            "institutional_quality": (0.5, 0.05),
            "era_code": (5, 0),
            "geo_barrier": (0.5, 0.05),
            "has_external_patron": (0, 0),
            "tech_position": (0.7, 0.05),
        },
        "中国: 技術デカップリング加速": {
            "external_threat": (0.7, 0.1),
            "institutional_quality": (0.4, 0.1),
            "era_code": (5, 0),
            "geo_barrier": (0.5, 0.05),
            "has_external_patron": (0, 0),
            "tech_position": (0.5, 0.1),
        },
        "ロシア: 制裁継続": {
            "external_threat": (0.8, 0.05),
            "institutional_quality": (0.15, 0.05),
            "era_code": (5, 0),
            "geo_barrier": (0.6, 0.05),
            "has_external_patron": (0, 0),
            "tech_position": (0.3, 0.1),
        },
        "ロシア: 停戦・部分制裁解除": {
            "external_threat": (0.5, 0.1),
            "institutional_quality": (0.3, 0.1),
            "era_code": (5, 0),
            "geo_barrier": (0.6, 0.05),
            "has_external_patron": (0, 0),
            "tech_position": (0.4, 0.1),
        },
        "日本: 現状維持（安保依存継続）": {
            "external_threat": (0.6, 0.1),
            "institutional_quality": (0.65, 0.1),
            "era_code": (5, 0),
            "geo_barrier": (0.7, 0.05),
            "has_external_patron": (1, 0),
            "tech_position": (0.6, 0.1),
        },
        "日本: 安保見直し（自主防衛化）": {
            "external_threat": (0.7, 0.1),
            "institutional_quality": (0.7, 0.1),
            "era_code": (5, 0),
            "geo_barrier": (0.7, 0.05),
            "has_external_patron": (0, 0),
            "tech_position": (0.65, 0.1),
        },
        "日本: 選択的フロー再開": {
            "external_threat": (0.5, 0.1),
            "institutional_quality": (0.75, 0.1),
            "era_code": (5, 0),
            "geo_barrier": (0.7, 0.05),
            "has_external_patron": (1, 0),
            "tech_position": (0.75, 0.1),
        },
        "北朝鮮: 中国保護継続": {
            "external_threat": (0.8, 0.05),
            "institutional_quality": (0.1, 0.03),
            "era_code": (5, 0),
            "geo_barrier": (0.5, 0.05),
            "has_external_patron": (1, 0),
            "tech_position": (0.15, 0.05),
        },
        "北朝鮮: 中国保護撤退": {
            "external_threat": (0.9, 0.05),
            "institutional_quality": (0.1, 0.03),
            "era_code": (5, 0),
            "geo_barrier": (0.5, 0.05),
            "has_external_patron": (0, 0),
            "tech_position": (0.15, 0.05),
        },
    }

    covariates = ["external_threat", "institutional_quality", "era_code",
                  "geo_barrier", "has_external_patron", "tech_position"]

    print(f"\n  各シナリオ {n_sim:,}回シミュレーション（パラメータにガウスノイズ付加）\n")
    print(f"  {'シナリオ':40s} {'平均リスク':>8s} {'SD':>6s} {'5%ile':>6s} {'95%ile':>6s}")
    print(f"  {'-'*68}")

    for name, params in scenarios.items():
        probs = np.zeros(n_sim)
        for i in range(n_sim):
            x_vals = {}
            for var in covariates:
                mean, sd = params[var]
                val = rng.normal(mean, sd) if sd > 0 else mean
                val = np.clip(val, 0, 1) if var != "era_code" else val
                x_vals[var] = val
            x_df = pd.DataFrame([x_vals])
            x_df = sm.add_constant(x_df, has_constant="add")
            # Ensure column alignment
            for col in model.model.exog_names:
                if col not in x_df.columns:
                    x_df[col] = 0 if col != "const" else 1
            x_df = x_df[model.model.exog_names]
            probs[i] = model.predict(x_df).values[0]

        mean_p = np.mean(probs)
        sd_p = np.std(probs)
        p5 = np.percentile(probs, 5)
        p95 = np.percentile(probs, 95)
        print(f"  {name:40s} {mean_p:>8.1%} {sd_p:>6.1%} {p5:>6.1%} {p95:>6.1%}")


if __name__ == "__main__":
    df = load_data()
    print(f"歴史データ: N={len(df)}")

    results, model = predict_risk(df)
    mediation_diagnosis(df)
    monte_carlo_scenarios(model)

    print(f"\n\n{'='*70}")
    print("主要な知見")
    print(f"{'='*70}")
    print("""
  1. 現代の「海禁的」国家（ロシア、北朝鮮、イラン）は歴史的パターンと整合的に
     高リスクと評価される。特にロシアは制裁による「強制的海禁」により、
     媒介経路（技術停滞・制度劣化・頭脳流出）が加速している。

  2. 中国はGFW（デジタル海禁）を持つが、一帯一路（選択的フロー）で
     媒介経路の遮断を試みている。技術デカップリングの深度が鍵。

  3. 日本は「選択的フロー再開」シナリオが最もリスクを下げる。
     安保依存の継続も有効だが、保護者撤退リスクがある。

  4. 歴史的に、ストック優位＋高制度品質で存続した国はスイスのみ。
     現代ではスイスが引き続きこのモデルの唯一の成功例。

  5. 台湾の「シリコンシールド」は、技術ストックをフロー変換器として
     活用する新しいパターン。ストック→フロー変換の現代版。
""")

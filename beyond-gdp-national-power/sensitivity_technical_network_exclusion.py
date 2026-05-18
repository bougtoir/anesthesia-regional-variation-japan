"""
感度分析: 技術的海禁政策 × outcome再分類 (Technical Maritime Ban + Disrupted)

2つの感度軸:
  軸1 — closure_type 再分類（技術的海禁）
    A) ベースライン（変更なし）
    B) 強い候補のみ再分類（5国）
    C) 全候補再分類（7国）

  軸2 — disrupted 18国の帰属
    overtaken扱い: disrupted → overtaken（体制崩壊＝征服と見なす）
    survived扱い:  disrupted → survived （国家存続＝生存と見なす）

  → 合計 3×2 = 6 シナリオ
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import scipy.stats as stats
import math
import statsmodels.api as sm
from data import load_data


# ============================================================
# 技術的海禁政策の対象エンティティ定義
# ============================================================

STRONG_CANDIDATES = [
    "漢朝（前漢〜後漢）",
    "マリ帝国",
    "クメール帝国（アンコール）",
    "キエフ大公国",
    "ティムール朝",
]

MODERATE_CANDIDATES = [
    "ササン朝ペルシア",
    "ビルマ（コンバウン朝）",
]

RATIONALE = {
    "漢朝（前漢〜後漢）": "古代東アジア。地中海世界からの定期海路なし。シルクロード陸路のみ。"
                          "海路到達は冒険的商人の散発的試みに限定",
    "マリ帝国": "中世サヘル。大西洋岸への定期航路は15世紀ポルトガル探検まで不在。"
                "外部接触はサハラ縦断キャラバンのみ",
    "クメール帝国（アンコール）": "中世内陸東南アジア。近隣のシュリーヴィジャヤは海洋国家だが、"
                               "アンコール自体は内陸に位置し外洋からの直接航路なし",
    "キエフ大公国": "中世東欧。ドニエプル川・ヴァリャーグ路（河川）が主な国際接続。"
                   "外洋からの定期航路なし",
    "ティムール朝": "近世中央アジア。完全内陸国家。海洋アクセスなし。"
                   "シルクロード陸路のみ",
    "ササン朝ペルシア": "ペルシア湾交易はあるが限定的。"
                       "インド洋・地中海との定期海路は未確立。陸路が主な国際接続",
    "ビルマ（コンバウン朝）": "沿岸部はあるが外洋定期航路は英国進出まで限定的。"
                            "内陸志向の閉鎖的経済構造",
}

# disrupted に再分類した18国（参照用）
DISRUPTED_ENTITIES = [
    "徳川日本", "オスマン帝国（後期）", "ポルトガル帝国", "オランダ共和国",
    "スペイン帝国", "1930s日本（大東亜共栄圏）", "ナチスドイツ", "ファシストイタリア",
    "ソ連", "東ドイツ", "ユーゴスラビア", "スウェーデン帝国",
    "ロシア帝国（ピョートル後）", "オーストリア＝ハンガリー帝国", "ナポレオン帝国",
    "チェコスロバキア（共産期）", "ポーランド（共産期）", "ルーマニア（共産期）",
]


def apply_technical_network_exclusion(df, candidates):
    """指定候補を technical_network_exclusion に再分類したDataFrameを返す"""
    df_new = df.copy()
    mask = df_new["entity"].isin(candidates)
    df_new.loc[mask, "closure_type"] = "technical_network_exclusion"
    return df_new


def apply_disrupted_assignment(df, mode):
    """disrupted を conquered または survived に割り当てた二値変数を設定"""
    df_new = df.copy()
    if mode == "as_conquered":
        df_new["outcome_binary"] = (df_new["outcome"].isin(["overtaken", "disrupted"])).astype(int)
        # For crosstab: remap outcome to binary label
        df_new["outcome_bin_label"] = df_new["outcome_binary"].map({1: "overtaken", 0: "survived"})
    else:  # as_survived
        df_new["outcome_binary"] = (df_new["outcome"] == "overtaken").astype(int)
        df_new["outcome_bin_label"] = df_new["outcome_binary"].map({1: "overtaken", 0: "survived"})
    return df_new


# ============================================================
# 分析関数群
# ============================================================

def compute_confusion_stats(df):
    """混同行列の統計量を計算（outcome_bin_label使用）"""
    ct = pd.crosstab(df["dominant"], df["outcome_bin_label"])
    ct = ct.reindex(index=["stock", "flow"], columns=["overtaken", "survived"], fill_value=0)

    TP = ct.loc["stock", "overtaken"]
    FP = ct.loc["stock", "survived"]
    FN = ct.loc["flow", "overtaken"]
    TN = ct.loc["flow", "survived"]
    N = TP + FP + FN + TN

    stock_rate = TP / (TP + FP) if (TP + FP) > 0 else 0
    flow_rate = FN / (FN + TN) if (FN + TN) > 0 else 0

    sensitivity = TP / (TP + FN) if (TP + FN) > 0 else 0
    specificity = TN / (TN + FP) if (TN + FP) > 0 else 0
    accuracy = (TP + TN) / N if N > 0 else 0
    ppv = TP / (TP + FP) if (TP + FP) > 0 else 0
    npv = TN / (TN + FN) if (TN + FN) > 0 else 0

    odds_ratio = (TP * TN) / (FP * FN) if (FP * FN) > 0 else float("inf")
    phi = (TP * TN - FP * FN) / math.sqrt(
        max(1, (TP + FP) * (FN + TN) * (TP + FN) * (FP + TN))
    )

    table = np.array([[TP, FP], [FN, TN]])
    _, p_fisher = stats.fisher_exact(table, alternative="greater")
    chi2, p_chi2, _, _ = stats.chi2_contingency(table, correction=True)

    return {
        "N": N, "TP": TP, "FP": FP, "FN": FN, "TN": TN,
        "stock_conquest_rate": stock_rate,
        "flow_conquest_rate": flow_rate,
        "sensitivity": sensitivity,
        "specificity": specificity,
        "accuracy": accuracy,
        "ppv": ppv, "npv": npv,
        "OR": odds_ratio, "phi": phi,
        "p_fisher": p_fisher, "chi2": chi2, "p_chi2": p_chi2,
    }


def compute_closure_analysis(df):
    """closure_type 別の征服率（二値化後のoutcome_binary使用）"""
    results = {}
    for ct in df["closure_type"].unique():
        sub = df[df["closure_type"] == ct]
        n = len(sub)
        n_conquered = int(sub["outcome_binary"].sum())
        rate = n_conquered / n if n > 0 else 0
        results[ct] = {"n": n, "overtaken": n_conquered, "rate": rate}
    return results


def compute_logistic_with_closure(df, include_closure_binary=True):
    """海禁ダミーを含む多変量ロジスティック回帰"""
    y = df["outcome_binary"]
    df_work = df.copy()

    if include_closure_binary:
        df_work["has_maritime_ban"] = (
            df_work["closure_type"].isin(["maritime_ban", "technical_network_exclusion", "sakoku"])
        ).astype(int)

    covariates_base = [
        "dominant_binary", "geo_barrier", "external_threat",
        "tech_position", "institutional_quality", "era_code",
        "has_external_patron",
    ]
    covariates_with_ban = covariates_base + ["has_maritime_ban"]

    results = {}

    for name, covs in [
        ("base", covariates_base),
        ("with_ban", covariates_with_ban if include_closure_binary else covariates_base),
    ]:
        X = sm.add_constant(df_work[covs].astype(float))
        try:
            model = sm.Logit(y, X).fit(disp=0, maxiter=200)
            coefs = {}
            for var in covs:
                coefs[var] = {
                    "coef": model.params[var],
                    "OR": np.exp(model.params[var]),
                    "p": model.pvalues[var],
                    "ci_lo": np.exp(model.conf_int().loc[var, 0]),
                    "ci_hi": np.exp(model.conf_int().loc[var, 1]),
                }
            results[name] = {
                "aic": model.aic,
                "bic": model.bic,
                "pseudo_r2": model.prsquared,
                "coefs": coefs,
                "converged": True,
            }
        except Exception as e:
            results[name] = {"converged": False, "error": str(e)}

    return results


def compute_mediation_paths(df):
    """主要な媒介パスの効果量を計算"""
    y = df["outcome_binary"]
    x = df["dominant_binary"]

    paths = {}

    X_const = sm.add_constant(x)

    # Total effect (mediator-independent)
    c = None
    try:
        mod_total = sm.Logit(y, X_const).fit(disp=0)
        c = mod_total.params.iloc[1]
    except Exception:
        pass

    for mediator_key, mediator_col in [
        ("tech", "tech_position"),
        ("inst", "institutional_quality"),
        ("trade", "trade_openness"),
    ]:
        try:
            mod_m = sm.OLS(df[mediator_col], X_const).fit()
            a = mod_m.params.iloc[1]
            XM = sm.add_constant(pd.DataFrame({"x": x, "m": df[mediator_col]}))
            mod_y = sm.Logit(y, XM).fit(disp=0, maxiter=100)
            b = mod_y.params["m"]
            c_prime = mod_y.params["x"]
            paths[mediator_key] = {"a": a, "b": b, "ab": a * b, "c": c, "c_prime": c_prime}
        except Exception:
            paths[mediator_key] = None

    return paths


# ============================================================
# メイン感度分析
# ============================================================

def run_sensitivity():
    df_base = load_data()

    # 軸1: closure_type 再分類
    closure_scenarios = {
        "baseline": {
            "label": "ベースライン",
            "df": df_base,
            "reclassified": [],
        },
        "strong": {
            "label": "+5国再分類",
            "df": apply_technical_network_exclusion(df_base, STRONG_CANDIDATES),
            "reclassified": STRONG_CANDIDATES,
        },
        "all": {
            "label": "+7国再分類",
            "df": apply_technical_network_exclusion(
                df_base, STRONG_CANDIDATES + MODERATE_CANDIDATES
            ),
            "reclassified": STRONG_CANDIDATES + MODERATE_CANDIDATES,
        },
    }

    # 軸2: disrupted の帰属
    disrupted_modes = {
        "as_conquered": "disrupted→征服扱い",
        "as_survived": "disrupted→存続扱い",
    }

    # ============================================================
    # 0. データの基本情報
    # ============================================================
    print("=" * 80)
    print("感度分析: 技術的海禁政策 × outcome再分類 (disrupted帰属)")
    print("=" * 80)

    n_conquered = len(df_base[df_base["outcome"] == "overtaken"])
    n_disrupted = len(df_base[df_base["outcome"] == "disrupted"])
    n_survived = len(df_base[df_base["outcome"] == "survived"])
    print(f"\n  データ: N={len(df_base)}")
    print(f"  outcome 3カテゴリ: overtaken={n_conquered}, disrupted={n_disrupted}, survived={n_survived}")
    print(f"\n  感度分析軸:")
    print(f"    軸1 — closure_type再分類: ベースライン / +5国 / +7国")
    print(f"    軸2 — disrupted帰属: 征服扱い(overtaken+disrupted vs survived)")
    print(f"                          存続扱い(overtaken vs disrupted+survived)")
    print(f"    → 合計 3×2 = 6 シナリオ")

    # ============================================================
    # 1. disrupted エンティティ一覧
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 1: disrupted エンティティ一覧 (18国)")
    print("=" * 80)

    disrupted_df = df_base[df_base["outcome"] == "disrupted"]
    for _, row in disrupted_df.iterrows():
        print(f"  {row['entity']:35s} | {row['period']:20s} | {row['dominant']:5s} | {row['closure_type']}")

    # ============================================================
    # 2. 技術的海禁再分類対象
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 2: 技術的海禁再分類対象エンティティ")
    print("=" * 80)

    for entity in STRONG_CANDIDATES + MODERATE_CANDIDATES:
        row = df_base[df_base["entity"] == entity].iloc[0]
        strength = "強" if entity in STRONG_CANDIDATES else "中"
        print(f"\n  [{strength}] {entity}")
        print(f"      時代: {row['era']}, 地域: {row['region']}")
        print(f"      元の closure_type: {row['closure_type']}")
        print(f"      trade_openness: {row['trade_openness']:.2f}")
        print(f"      dominant: {row['dominant']}, outcome: {row['outcome']}")
        print(f"      根拠: {RATIONALE[entity]}")

    # ============================================================
    # 3. outcome分布の変化（各シナリオ）
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 3: 二値化後の outcome 分布")
    print("=" * 80)

    for d_mode, d_label in disrupted_modes.items():
        print(f"\n  ═══ {d_label} ═══")
        for c_key, c_sc in closure_scenarios.items():
            df_prepared = apply_disrupted_assignment(c_sc["df"], d_mode)
            n_conq = int(df_prepared["outcome_binary"].sum())
            n_surv = len(df_prepared) - n_conq
            print(f"    {c_sc['label']:12s}: overtaken={n_conq}, survived={n_surv}")

            # closure_type別
            for ct in sorted(df_prepared["closure_type"].unique()):
                sub = df_prepared[df_prepared["closure_type"] == ct]
                n = len(sub)
                nc = int(sub["outcome_binary"].sum())
                rate = nc / n if n > 0 else 0
                print(f"      {ct:25s}: {n:3d}国  征服率={rate:.1%} ({nc}/{n})")

    # ============================================================
    # 4. 混同行列比較（6シナリオ）
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 4: 混同行列統計量の比較 (dominant × outcome)")
    print("=" * 80)

    all_cm = {}
    for d_mode, d_label in disrupted_modes.items():
        print(f"\n  ═══ {d_label} ═══")
        for c_key, c_sc in closure_scenarios.items():
            combo_key = f"{d_mode}__{c_key}"
            df_prepared = apply_disrupted_assignment(c_sc["df"], d_mode)
            cm = compute_confusion_stats(df_prepared)
            all_cm[combo_key] = cm

        # Print comparison table within this disrupted mode
        header = f"    {'指標':25s}"
        for c_key in closure_scenarios:
            header += f" {closure_scenarios[c_key]['label']:>14s}"
        print(f"\n{header}")
        print(f"    {'-' * 70}")

        metrics = [
            ("OR (オッズ比)", "OR", ".3f"),
            ("φ係数", "phi", ".3f"),
            ("Fisher p値(片側)", "p_fisher", ".4f"),
            ("χ² (Yates)", "chi2", ".3f"),
            ("感度", "sensitivity", ".1%"),
            ("特異度", "specificity", ".1%"),
            ("正確度", "accuracy", ".1%"),
            ("ストック征服率", "stock_conquest_rate", ".1%"),
            ("フロー征服率", "flow_conquest_rate", ".1%"),
        ]

        for label, key_m, fmt in metrics:
            line = f"    {label:25s}"
            for c_key in closure_scenarios:
                combo_key = f"{d_mode}__{c_key}"
                val = all_cm[combo_key][key_m]
                line += f" {val:>14{fmt}}"
            print(line)

    # ============================================================
    # 5. 多変量ロジスティック回帰（6シナリオ）
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 5: 多変量ロジスティック回帰（海禁ダミー追加）")
    print("=" * 80)

    for d_mode, d_label in disrupted_modes.items():
        print(f"\n  ═══ {d_label} ═══")
        for c_key, c_sc in closure_scenarios.items():
            df_prepared = apply_disrupted_assignment(c_sc["df"], d_mode)
            print(f"\n    --- {c_sc['label']} ---")
            lr = compute_logistic_with_closure(df_prepared)

            for model_name, model_label in [
                ("base", "基本モデル"),
                ("with_ban", "海禁ダミーあり"),
            ]:
                r = lr[model_name]
                if not r.get("converged", False):
                    print(f"\n      [{model_label}] 収束失敗: {r.get('error', 'N/A')}")
                    continue
                print(f"\n      [{model_label}] AIC={r['aic']:.1f}, Pseudo-R²={r['pseudo_r2']:.3f}")
                print(f"      {'変数':30s} {'OR':>8s} {'95%CI':>22s} {'p値':>8s}")
                print(f"      {'-'*70}")
                for var, v in r["coefs"].items():
                    sig = " *" if v["p"] < 0.05 else "  " if v["p"] < 0.10 else ""
                    print(f"      {var:30s} {v['OR']:>8.3f} [{v['ci_lo']:>7.3f}, {v['ci_hi']:>7.3f}] {v['p']:>8.4f}{sig}")

    # ============================================================
    # 6. 海禁タイプ別征服率（6シナリオ）
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 6: 海禁→征服 Fisher検定（6シナリオ）")
    print("=" * 80)

    for d_mode, d_label in disrupted_modes.items():
        print(f"\n  ═══ {d_label} ═══")
        for c_key, c_sc in closure_scenarios.items():
            df_prepared = apply_disrupted_assignment(c_sc["df"], d_mode)
            print(f"\n    --- {c_sc['label']} ---")

            has_ban = df_prepared["closure_type"].isin(
                ["maritime_ban", "technical_network_exclusion", "sakoku"]
            )
            ban_df = df_prepared[has_ban]
            no_ban_df = df_prepared[~has_ban]

            ban_rate = ban_df["outcome_binary"].mean() if len(ban_df) > 0 else 0
            no_rate = no_ban_df["outcome_binary"].mean() if len(no_ban_df) > 0 else 0

            print(f"      海禁あり: {len(ban_df)}国, 征服率={ban_rate:.1%}")
            print(f"      海禁なし: {len(no_ban_df)}国, 征服率={no_rate:.1%}")

            if len(ban_df) > 0 and len(no_ban_df) > 0:
                rd = ban_rate - no_rate
                rr = ban_rate / no_rate if no_rate > 0 else float("inf")
                print(f"      リスク差: {rd:+.1%}")
                print(f"      リスク比: {rr:.3f}")

                ban_conq = int(ban_df["outcome_binary"].sum())
                ban_surv = len(ban_df) - ban_conq
                no_conq = int(no_ban_df["outcome_binary"].sum())
                no_surv = len(no_ban_df) - no_conq
                table = np.array([[ban_conq, ban_surv], [no_conq, no_surv]])
                _, p = stats.fisher_exact(table, alternative="greater")
                print(f"      Fisher検定: p={p:.4f}" + (" *" if p < 0.05 else ""))

    # ============================================================
    # 7. 媒介分析（6シナリオ）
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 7: 媒介分析パス係数")
    print("=" * 80)

    path_labels = {
        "tech": "ストック優位→技術水準→征服",
        "inst": "ストック優位→制度品質→征服",
        "trade": "ストック優位→貿易開放度→征服",
    }

    for d_mode, d_label in disrupted_modes.items():
        print(f"\n  ═══ {d_label} ═══")
        med_results = {}
        for c_key, c_sc in closure_scenarios.items():
            df_prepared = apply_disrupted_assignment(c_sc["df"], d_mode)
            med_results[c_key] = compute_mediation_paths(df_prepared)

        for path_key, path_label in path_labels.items():
            print(f"\n    【{path_label}】")
            cprime_label = "c' (直接)"
            print(f"      {'シナリオ':12s} {'a (X→M)':>10s} {'b (M→Y)':>10s} {'a×b':>10s} {'c (総効果)':>12s} {cprime_label:>12s}")
            print(f"      {'-' * 72}")
            for c_key in closure_scenarios:
                r = med_results[c_key].get(path_key)
                if r is None:
                    print(f"      {closure_scenarios[c_key]['label']:12s}  (収束失敗)")
                    continue
                c_val = r['c'] if r['c'] is not None else float('nan')
                print(f"      {closure_scenarios[c_key]['label']:12s} {r['a']:>10.4f} {r['b']:>10.4f} {r['ab']:>10.4f} {c_val:>12.4f} {r['c_prime']:>12.4f}")

    # ============================================================
    # 8. ブートストラップOR（6シナリオ）
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SECTION 8: ブートストラップOR推定")
    print("=" * 80)

    rng = np.random.default_rng(42)
    n_boot = 5000

    for d_mode, d_label in disrupted_modes.items():
        print(f"\n  ═══ {d_label} ═══")
        for c_key, c_sc in closure_scenarios.items():
            df_prepared = apply_disrupted_assignment(c_sc["df"], d_mode)
            n = len(df_prepared)
            boot_ors = np.zeros(n_boot)

            for i in range(n_boot):
                idx = rng.choice(n, size=n, replace=True)
                boot_df = df_prepared.iloc[idx].reset_index(drop=True)
                try:
                    ct = pd.crosstab(boot_df["dominant"], boot_df["outcome_bin_label"])
                    ct = ct.reindex(index=["stock", "flow"], columns=["overtaken", "survived"], fill_value=0)
                    tp = ct.loc["stock", "overtaken"]
                    fp = ct.loc["stock", "survived"]
                    fn = ct.loc["flow", "overtaken"]
                    tn = ct.loc["flow", "survived"]
                    boot_ors[i] = (tp * tn) / (fp * fn) if (fp * fn) > 0 else np.nan
                except (KeyError, ZeroDivisionError):
                    boot_ors[i] = np.nan

            valid = boot_ors[~np.isnan(boot_ors)]
            if len(valid) >= 100:
                ci_lo, ci_hi = np.percentile(valid, [2.5, 97.5])
                median_or = np.median(valid)
                combo_key = f"{d_mode}__{c_key}"
                print(f"    {c_sc['label']:12s}: median={median_or:.3f}, 95%CI=[{ci_lo:.3f}, {ci_hi:.3f}]  (point={all_cm[combo_key]['OR']:.3f})")

    # ============================================================
    # 9. サマリーテーブル
    # ============================================================
    print("\n\n" + "=" * 80)
    print("SUMMARY: 全6シナリオ比較")
    print("=" * 80)

    summary_rows = []
    for d_mode, d_label in disrupted_modes.items():
        for c_key, c_sc in closure_scenarios.items():
            combo_key = f"{d_mode}__{c_key}"
            cm = all_cm[combo_key]
            df_prepared = apply_disrupted_assignment(c_sc["df"], d_mode)
            lr = compute_logistic_with_closure(df_prepared)
            ban_info = compute_closure_analysis(df_prepared)

            n_ban = sum(
                v["n"] for k, v in ban_info.items()
                if k in ["maritime_ban", "technical_network_exclusion", "sakoku"]
            )
            ban_conquered = sum(
                v["overtaken"] for k, v in ban_info.items()
                if k in ["maritime_ban", "technical_network_exclusion", "sakoku"]
            )
            ban_rate = ban_conquered / n_ban if n_ban > 0 else 0

            # Fisher for ban vs no-ban
            has_ban = df_prepared["closure_type"].isin(
                ["maritime_ban", "technical_network_exclusion", "sakoku"]
            )
            ban_df = df_prepared[has_ban]
            no_ban_df = df_prepared[~has_ban]
            ban_conq = int(ban_df["outcome_binary"].sum())
            ban_surv = len(ban_df) - ban_conq
            no_conq = int(no_ban_df["outcome_binary"].sum())
            no_surv = len(no_ban_df) - no_conq
            _, p_ban_fisher = stats.fisher_exact(
                np.array([[ban_conq, ban_surv], [no_conq, no_surv]]),
                alternative="greater"
            )

            lr_with = lr.get("with_ban", {})
            ban_or_mv = "NC"
            ban_p_mv = "NC"
            if lr_with.get("converged"):
                bc = lr_with["coefs"].get("has_maritime_ban", {})
                if bc:
                    ban_or_mv = f"{bc['OR']:.3f}"
                    ban_p_mv = f"{bc['p']:.4f}"

            summary_rows.append({
                "disrupted": d_label,
                "海禁再分類": c_sc["label"],
                "征服数": int(df_prepared["outcome_binary"].sum()),
                "OR": f"{cm['OR']:.3f}",
                "φ": f"{cm['phi']:.3f}",
                "Fisher_p": f"{cm['p_fisher']:.4f}",
                "海禁征服率": f"{ban_rate:.1%}",
                "海禁Fisher_p": f"{p_ban_fisher:.4f}",
                "海禁OR(MV)": ban_or_mv,
                "海禁p(MV)": ban_p_mv,
            })

    summary_df = pd.DataFrame(summary_rows)
    print(f"\n{summary_df.to_string(index=False)}")

    # ============================================================
    # 10. 解釈
    # ============================================================
    print("\n\n" + "=" * 80)
    print("解釈と考察")
    print("=" * 80)

    # Compare OR across disrupted modes
    or_conq = all_cm["as_conquered__baseline"]["OR"]
    or_surv = all_cm["as_survived__baseline"]["OR"]
    p_conq = all_cm["as_conquered__baseline"]["p_fisher"]
    p_surv = all_cm["as_survived__baseline"]["p_fisher"]

    print(f"""
  1. 【disrupted帰属の影響（ベースライン）】
     征服扱い: OR={or_conq:.3f}, Fisher p={p_conq:.4f}
     存続扱い: OR={or_surv:.3f}, Fisher p={p_surv:.4f}
     OR差: {abs(or_conq - or_surv):.3f}
     → disrupted の帰属は主要結果に{'大きく' if abs(or_conq - or_surv) > 0.3 else '軽微に'}影響。

  2. 【技術的海禁再分類の影響】""")

    for d_mode, d_label in disrupted_modes.items():
        or_b = all_cm[f"{d_mode}__baseline"]["OR"]
        or_s = all_cm[f"{d_mode}__strong"]["OR"]
        or_a = all_cm[f"{d_mode}__all"]["OR"]
        print(f"     {d_label}:")
        print(f"       ベースライン OR={or_b:.3f} → +5国 OR={or_s:.3f} → +7国 OR={or_a:.3f}")

    print(f"""
  3. 【三カテゴリ outcome の含意】
     overtaken: 外部勢力による植民地化・併合（明確な外部征服）
     disrupted: 内部崩壊・革命・体制変革（国家/後継国家は独立維持）
     survived:  政体・国家ともに存続

     disrupted を征服側に入れるか存続側に入れるかで結果が変化する場合、
     「征服」の操作的定義に対する分析の感度が高いことを意味する。
     逆に変化しなければ、結論は定義に対して頑健である。

  4. 【政策的海禁 vs 技術的海禁（再確認）】
     政策的海禁は「選択的閉鎖」であり、国家が意図的にフローを制限する。
     技術的海禁は「受動的閉鎖」であり、技術・地理的制約によりフローが不可能。
     両者が同様の帰結をもたらすならば、閉鎖メカニズムは意図ではなく
     フロー遮断そのものに由来することを示唆する。
""")


if __name__ == "__main__":
    run_sensitivity()

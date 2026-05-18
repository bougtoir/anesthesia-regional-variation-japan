"""
Beyond GDP: 媒介分析 (Mediation Analysis)

仮説的因果パス:
  ストック優位(X) → 制度劣化/技術停滞(M) → 征服(Y)

検証する媒介パス:
  Path 1: ストック優位 → 制度品質低下 → 征服
  Path 2: ストック優位 → 技術水準低下 → 征服
  Path 3: ストック優位 → 外部脅威脆弱性増大 → 征服
  Path 4: 貿易開放度低下 → 制度品質低下 → 征服
  Path 5: 貿易開放度低下 → 技術水準低下 → 征服

手法:
  1. Baron-Kenny (1986) の古典的媒介分析ステップ
  2. Sobel検定
  3. ブートストラップ信頼区間による間接効果の検定
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats as sp_stats
from data import load_data


def baron_kenny_step(y, x, m, x_label, m_label, y_label):
    """Baron-Kenny (1986) 4-step mediation analysis"""
    print(f"\n{'='*60}")
    print(f"媒介パス: {x_label} → {m_label} → {y_label}")
    print(f"{'='*60}")

    X = sm.add_constant(x)

    # Step 1: X → Y (total effect, c path)
    if y.nunique() == 2:
        model_xy = sm.Logit(y, X).fit(disp=0)
    else:
        model_xy = sm.OLS(y, X).fit()
    c = model_xy.params.iloc[1]
    p_c = model_xy.pvalues.iloc[1]
    print(f"\n  Step 1 (総効果 c): {x_label} → {y_label}")
    print(f"    c = {c:.4f}, p = {p_c:.4f}")

    # Step 2: X → M (a path)
    model_xm = sm.OLS(m, X).fit()
    a = model_xm.params.iloc[1]
    p_a = model_xm.pvalues.iloc[1]
    r2_xm = model_xm.rsquared
    print(f"\n  Step 2 (a path): {x_label} → {m_label}")
    print(f"    a = {a:.4f}, p = {p_a:.4f}, R² = {r2_xm:.4f}")

    # Step 3 & 4: X + M → Y (b path & c' path)
    XM = sm.add_constant(pd.DataFrame({"x": x, "m": m}))
    if y.nunique() == 2:
        model_xmy = sm.Logit(y, XM).fit(disp=0)
    else:
        model_xmy = sm.OLS(y, XM).fit()
    b = model_xmy.params["m"]
    p_b = model_xmy.pvalues["m"]
    c_prime = model_xmy.params["x"]
    p_c_prime = model_xmy.pvalues["x"]
    print(f"\n  Step 3 (b path): {m_label} → {y_label} (controlling {x_label})")
    print(f"    b = {b:.4f}, p = {p_b:.4f}")
    print(f"\n  Step 4 (直接効果 c'): {x_label} → {y_label} (controlling {m_label})")
    print(f"    c' = {c_prime:.4f}, p = {p_c_prime:.4f}")

    # Indirect effect (a*b)
    indirect = a * b
    print(f"\n  間接効果 (a×b) = {indirect:.4f}")
    print(f"  直接効果 (c')  = {c_prime:.4f}")
    print(f"  総効果 (c)     = {c:.4f}")

    # Proportion mediated
    if abs(c) > 1e-10:
        prop_med = indirect / c
        print(f"  媒介割合       = {prop_med:.1%}")
    else:
        prop_med = np.nan
        print(f"  媒介割合       = N/A (総効果≈0)")

    # Sobel test
    se_a = model_xm.bse.iloc[1]
    se_b = model_xmy.bse["m"]
    sobel_se = np.sqrt(a**2 * se_b**2 + b**2 * se_a**2)
    sobel_z = indirect / sobel_se if sobel_se > 0 else 0
    sobel_p = 2 * (1 - sp_stats.norm.cdf(abs(sobel_z)))
    print(f"\n  Sobel検定: z = {sobel_z:.3f}, p = {sobel_p:.4f}")

    # Baron-Kenny criteria check
    print(f"\n  【Baron-Kenny基準チェック】")
    step1_ok = p_c < 0.10
    step2_ok = p_a < 0.10
    step3_ok = p_b < 0.10
    mediation_type = "なし"
    if step2_ok and step3_ok:
        if p_c_prime >= 0.10:
            mediation_type = "完全媒介"
        elif abs(c_prime) < abs(c):
            mediation_type = "部分媒介"
        else:
            mediation_type = "媒介なし（抑制効果の可能性）"
    print(f"    Step1 (c有意):     {'✓' if step1_ok else '×'} (p={p_c:.4f})")
    print(f"    Step2 (a有意):     {'✓' if step2_ok else '×'} (p={p_a:.4f})")
    print(f"    Step3 (b有意):     {'✓' if step3_ok else '×'} (p={p_b:.4f})")
    print(f"    → 媒介の種類:      {mediation_type}")

    return {
        "x_label": x_label,
        "m_label": m_label,
        "y_label": y_label,
        "a": a, "p_a": p_a,
        "b": b, "p_b": p_b,
        "c": c, "p_c": p_c,
        "c_prime": c_prime, "p_c_prime": p_c_prime,
        "indirect": indirect,
        "prop_mediated": prop_med,
        "sobel_z": sobel_z, "sobel_p": sobel_p,
        "mediation_type": mediation_type,
    }


def bootstrap_indirect_effect(y, x, m, n_boot=5000, alpha=0.05):
    """Bootstrap confidence interval for indirect effect (a*b)"""
    n = len(y)
    indirect_boots = np.zeros(n_boot)
    rng = np.random.default_rng(42)

    for i in range(n_boot):
        idx = rng.choice(n, size=n, replace=True)
        x_b, m_b, y_b = x.iloc[idx].reset_index(drop=True), m.iloc[idx].reset_index(drop=True), y.iloc[idx].reset_index(drop=True)

        X_b = sm.add_constant(x_b)
        try:
            model_a = sm.OLS(m_b, X_b).fit()
            a_b = model_a.params.iloc[1]

            XM_b = sm.add_constant(pd.DataFrame({"x": x_b, "m": m_b}))
            if y_b.nunique() == 2:
                model_b = sm.Logit(y_b, XM_b).fit(disp=0, maxiter=50)
            else:
                model_b = sm.OLS(y_b, XM_b).fit()
            b_b = model_b.params["m"]
            indirect_boots[i] = a_b * b_b
        except Exception:
            indirect_boots[i] = np.nan

    indirect_boots = indirect_boots[~np.isnan(indirect_boots)]
    ci_lo = np.percentile(indirect_boots, 100 * alpha / 2)
    ci_hi = np.percentile(indirect_boots, 100 * (1 - alpha / 2))
    mean_indirect = np.mean(indirect_boots)
    # Bias-corrected CI
    p_val = np.mean(indirect_boots <= 0) * 2  # two-sided
    if p_val > 1:
        p_val = 2 - p_val

    return mean_indirect, ci_lo, ci_hi, p_val


def run_all_mediation(df):
    """Run mediation analysis for all hypothesized paths"""
    y = df["outcome_binary"]

    results = []

    # --- Path 1: ストック優位(二値) → 制度品質 → 征服 ---
    r = baron_kenny_step(
        y, df["dominant_binary"], df["institutional_quality"],
        "ストック優位(二値)", "制度品質", "征服"
    )
    mean_i, ci_lo, ci_hi, bp = bootstrap_indirect_effect(
        y, df["dominant_binary"], df["institutional_quality"]
    )
    print(f"\n  Bootstrap間接効果: {mean_i:.4f} [{ci_lo:.4f}, {ci_hi:.4f}], p≈{bp:.4f}")
    r["boot_indirect"] = mean_i
    r["boot_ci"] = (ci_lo, ci_hi)
    r["boot_p"] = bp
    results.append(r)

    # --- Path 2: ストック優位(二値) → 技術水準 → 征服 ---
    r = baron_kenny_step(
        y, df["dominant_binary"], df["tech_position"],
        "ストック優位(二値)", "技術水準", "征服"
    )
    mean_i, ci_lo, ci_hi, bp = bootstrap_indirect_effect(
        y, df["dominant_binary"], df["tech_position"]
    )
    print(f"\n  Bootstrap間接効果: {mean_i:.4f} [{ci_lo:.4f}, {ci_hi:.4f}], p≈{bp:.4f}")
    r["boot_indirect"] = mean_i
    r["boot_ci"] = (ci_lo, ci_hi)
    r["boot_p"] = bp
    results.append(r)

    # --- Path 3: ストック優位(二値) → 外部脅威 → 征服 ---
    r = baron_kenny_step(
        y, df["dominant_binary"], df["external_threat"],
        "ストック優位(二値)", "外部脅威脆弱性", "征服"
    )
    mean_i, ci_lo, ci_hi, bp = bootstrap_indirect_effect(
        y, df["dominant_binary"], df["external_threat"]
    )
    print(f"\n  Bootstrap間接効果: {mean_i:.4f} [{ci_lo:.4f}, {ci_hi:.4f}], p≈{bp:.4f}")
    r["boot_indirect"] = mean_i
    r["boot_ci"] = (ci_lo, ci_hi)
    r["boot_p"] = bp
    results.append(r)

    # --- Path 4: 貿易開放度 → 制度品質 → 征服 ---
    r = baron_kenny_step(
        y, df["trade_openness"], df["institutional_quality"],
        "貿易開放度", "制度品質", "征服"
    )
    mean_i, ci_lo, ci_hi, bp = bootstrap_indirect_effect(
        y, df["trade_openness"], df["institutional_quality"]
    )
    print(f"\n  Bootstrap間接効果: {mean_i:.4f} [{ci_lo:.4f}, {ci_hi:.4f}], p≈{bp:.4f}")
    r["boot_indirect"] = mean_i
    r["boot_ci"] = (ci_lo, ci_hi)
    r["boot_p"] = bp
    results.append(r)

    # --- Path 5: 貿易開放度 → 技術水準 → 征服 ---
    r = baron_kenny_step(
        y, df["trade_openness"], df["tech_position"],
        "貿易開放度", "技術水準", "征服"
    )
    mean_i, ci_lo, ci_hi, bp = bootstrap_indirect_effect(
        y, df["trade_openness"], df["tech_position"]
    )
    print(f"\n  Bootstrap間接効果: {mean_i:.4f} [{ci_lo:.4f}, {ci_hi:.4f}], p≈{bp:.4f}")
    r["boot_indirect"] = mean_i
    r["boot_ci"] = (ci_lo, ci_hi)
    r["boot_p"] = bp
    results.append(r)

    # --- Path 6: 貿易開放度 → 外部脅威 → 征服 ---
    r = baron_kenny_step(
        y, df["trade_openness"], df["external_threat"],
        "貿易開放度", "外部脅威脆弱性", "征服"
    )
    mean_i, ci_lo, ci_hi, bp = bootstrap_indirect_effect(
        y, df["trade_openness"], df["external_threat"]
    )
    print(f"\n  Bootstrap間接効果: {mean_i:.4f} [{ci_lo:.4f}, {ci_hi:.4f}], p≈{bp:.4f}")
    r["boot_indirect"] = mean_i
    r["boot_ci"] = (ci_lo, ci_hi)
    r["boot_p"] = bp
    results.append(r)

    # ========== Summary ==========
    print("\n\n" + "=" * 70)
    print("媒介分析サマリー")
    print("=" * 70)
    print(f"\n{'パス':50s} {'間接効果':>8s} {'Sobel p':>8s} {'Boot CI':>24s} {'Boot p':>8s} {'判定':>10s}")
    print("-" * 110)
    for r in results:
        path = f"{r['x_label']} → {r['m_label']} → {r['y_label']}"
        ci_str = f"[{r['boot_ci'][0]:.4f}, {r['boot_ci'][1]:.4f}]"
        print(f"{path:50s} {r['indirect']:>8.4f} {r['sobel_p']:>8.4f} {ci_str:>24s} {r['boot_p']:>8.4f} {r['mediation_type']:>10s}")

    # ========== Key findings ==========
    print("\n\n" + "=" * 70)
    print("主要な知見")
    print("=" * 70)

    significant = [r for r in results if r["sobel_p"] < 0.10 or r["boot_p"] < 0.10]
    if significant:
        print("\n  統計的に有意な（またはボーダーライン）媒介パス:")
        for r in significant:
            print(f"    {r['x_label']} → {r['m_label']} → {r['y_label']}")
            print(f"      間接効果={r['indirect']:.4f}, 媒介割合={r['prop_mediated']:.1%}")
            print(f"      Sobel p={r['sobel_p']:.4f}, Bootstrap p={r['boot_p']:.4f}")
            print(f"      判定: {r['mediation_type']}")
    else:
        print("\n  α=0.10で有意な媒介パスは検出されませんでした。")
        print("  考えられる理由:")
        print("    1. サンプルサイズ不足（N=54）")
        print("    2. 媒介変数の測定精度（専門家コーディングの限界）")
        print("    3. 実際に直接効果が主要なパスである可能性")

    # a-path analysis
    print("\n\n  【a-path分析: ストック優位は何に影響するか】")
    for r in results:
        if "ストック優位" in r["x_label"]:
            sig = "***" if r["p_a"] < 0.001 else "**" if r["p_a"] < 0.01 else "*" if r["p_a"] < 0.05 else "†" if r["p_a"] < 0.10 else ""
            print(f"    → {r['m_label']}: a={r['a']:.3f} (p={r['p_a']:.4f}) {sig}")

    return results


if __name__ == "__main__":
    df = load_data()
    print(f"データセット: N={len(df)}")
    results = run_all_mediation(df)

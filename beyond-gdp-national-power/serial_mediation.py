"""
Beyond GDP: 逐次媒介モデル (Serial Mediation Analysis)

仮説的因果パス（逐次）:
  ストック優位(X) → 技術停滞(M1) → 制度劣化(M2) → 外部脅威脆弱性(M3) → 征服(Y)

検証:
  1. 単純逐次（X→M1→M2→Y）の3パス組み合わせ
  2. 全逐次（X→M1→M2→M3→Y）のフルパス
  3. ブートストラップ信頼区間
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import statsmodels.api as sm
from data import load_data


def serial_2mediator(y, x, m1, m2,
                     x_lab, m1_lab, m2_lab, y_lab, n_boot=5000):
    """Serial mediation: X → M1 → M2 → Y
    Paths: a1 (X→M1), d21 (M1→M2), b2 (M2→Y|X,M1)
    Indirect (serial) = a1 × d21 × b2
    """
    print(f"\n{'='*60}")
    print(f"逐次媒介: {x_lab} → {m1_lab} → {m2_lab} → {y_lab}")
    print(f"{'='*60}")

    X_const = sm.add_constant(x)

    # a1: X → M1
    model_m1 = sm.OLS(m1, X_const).fit()
    a1 = model_m1.params.iloc[1]
    p_a1 = model_m1.pvalues.iloc[1]

    # d21 + a2: X,M1 → M2
    XM1 = sm.add_constant(pd.DataFrame({"x": x, "m1": m1}))
    model_m2 = sm.OLS(m2, XM1).fit()
    a2 = model_m2.params["x"]       # direct X→M2
    d21 = model_m2.params["m1"]      # M1→M2
    p_a2 = model_m2.pvalues["x"]
    p_d21 = model_m2.pvalues["m1"]

    # b1, b2, c': X,M1,M2 → Y
    XM1M2 = sm.add_constant(pd.DataFrame({"x": x, "m1": m1, "m2": m2}))
    if y.nunique() == 2:
        model_y = sm.Logit(y, XM1M2).fit(disp=0, maxiter=100)
    else:
        model_y = sm.OLS(y, XM1M2).fit()
    c_prime = model_y.params["x"]
    b1 = model_y.params["m1"]
    b2 = model_y.params["m2"]
    p_c_prime = model_y.pvalues["x"]
    p_b1 = model_y.pvalues["m1"]
    p_b2 = model_y.pvalues["m2"]

    # Total effect
    if y.nunique() == 2:
        model_total = sm.Logit(y, X_const).fit(disp=0)
    else:
        model_total = sm.OLS(y, X_const).fit()
    c = model_total.params.iloc[1]
    p_c = model_total.pvalues.iloc[1]

    # Indirect effects
    ind_serial = a1 * d21 * b2        # X→M1→M2→Y (serial)
    ind_m1 = a1 * b1                  # X→M1→Y (parallel through M1)
    ind_m2 = a2 * b2                  # X→M2→Y (parallel through M2)
    ind_total = ind_serial + ind_m1 + ind_m2

    print(f"\n  【パス係数】")
    print(f"    a1 ({x_lab}→{m1_lab}):        {a1:.4f}  p={p_a1:.4f}")
    print(f"    d21 ({m1_lab}→{m2_lab}):       {d21:.4f}  p={p_d21:.4f}")
    print(f"    a2 ({x_lab}→{m2_lab}):         {a2:.4f}  p={p_a2:.4f}")
    print(f"    b1 ({m1_lab}→{y_lab}):         {b1:.4f}  p={p_b1:.4f}")
    print(f"    b2 ({m2_lab}→{y_lab}):         {b2:.4f}  p={p_b2:.4f}")
    print(f"    c' ({x_lab}→{y_lab}, 直接):    {c_prime:.4f}  p={p_c_prime:.4f}")
    print(f"    c  ({x_lab}→{y_lab}, 総効果):  {c:.4f}  p={p_c:.4f}")

    print(f"\n  【間接効果の分解】")
    print(f"    逐次間接 (a1×d21×b2):       {ind_serial:.4f}  [{x_lab}→{m1_lab}→{m2_lab}→{y_lab}]")
    print(f"    並列間接 M1 (a1×b1):         {ind_m1:.4f}  [{x_lab}→{m1_lab}→{y_lab}]")
    print(f"    並列間接 M2 (a2×b2):         {ind_m2:.4f}  [{x_lab}→{m2_lab}→{y_lab}]")
    print(f"    間接効果合計:                 {ind_total:.4f}")
    print(f"    直接効果 (c'):                {c_prime:.4f}")

    # Bootstrap
    rng = np.random.default_rng(42)
    n = len(y)
    boot_serial = np.zeros(n_boot)
    boot_m1 = np.zeros(n_boot)
    boot_m2 = np.zeros(n_boot)
    boot_total_ind = np.zeros(n_boot)

    for i in range(n_boot):
        idx = rng.choice(n, size=n, replace=True)
        xb = x.iloc[idx].reset_index(drop=True)
        m1b = m1.iloc[idx].reset_index(drop=True)
        m2b = m2.iloc[idx].reset_index(drop=True)
        yb = y.iloc[idx].reset_index(drop=True)
        try:
            Xb = sm.add_constant(xb)
            mod1 = sm.OLS(m1b, Xb).fit()
            a1b = mod1.params.iloc[1]

            XM1b = sm.add_constant(pd.DataFrame({"x": xb, "m1": m1b}))
            mod2 = sm.OLS(m2b, XM1b).fit()
            a2b = mod2.params["x"]
            d21b = mod2.params["m1"]

            XM1M2b = sm.add_constant(pd.DataFrame({"x": xb, "m1": m1b, "m2": m2b}))
            if yb.nunique() == 2:
                mod3 = sm.Logit(yb, XM1M2b).fit(disp=0, maxiter=50)
            else:
                mod3 = sm.OLS(yb, XM1M2b).fit()
            b1b = mod3.params["m1"]
            b2b = mod3.params["m2"]

            boot_serial[i] = a1b * d21b * b2b
            boot_m1[i] = a1b * b1b
            boot_m2[i] = a2b * b2b
            boot_total_ind[i] = boot_serial[i] + boot_m1[i] + boot_m2[i]
        except Exception:
            boot_serial[i] = np.nan
            boot_m1[i] = np.nan
            boot_m2[i] = np.nan
            boot_total_ind[i] = np.nan

    def boot_ci(arr, alpha=0.05):
        arr = arr[~np.isnan(arr)]
        if len(arr) < 100:
            return np.nan, np.nan, np.nan
        lo = np.percentile(arr, 100 * alpha / 2)
        hi = np.percentile(arr, 100 * (1 - alpha / 2))
        p = min(np.mean(arr <= 0), np.mean(arr >= 0)) * 2
        return lo, hi, p

    s_lo, s_hi, s_p = boot_ci(boot_serial)
    m1_lo, m1_hi, m1_p = boot_ci(boot_m1)
    m2_lo, m2_hi, m2_p = boot_ci(boot_m2)
    t_lo, t_hi, t_p = boot_ci(boot_total_ind)

    print(f"\n  【Bootstrap CI (n={n_boot})】")
    print(f"    逐次間接:  {ind_serial:.4f}  [{s_lo:.4f}, {s_hi:.4f}]  p≈{s_p:.4f}")
    print(f"    並列 M1:   {ind_m1:.4f}  [{m1_lo:.4f}, {m1_hi:.4f}]  p≈{m1_p:.4f}")
    print(f"    並列 M2:   {ind_m2:.4f}  [{m2_lo:.4f}, {m2_hi:.4f}]  p≈{m2_p:.4f}")
    print(f"    間接合計:  {ind_total:.4f}  [{t_lo:.4f}, {t_hi:.4f}]  p≈{t_p:.4f}")

    sig_serial = "有意" if s_lo > 0 or s_hi < 0 else "非有意"
    print(f"\n  → 逐次間接効果: {sig_serial}")

    return {
        "path": f"{x_lab}→{m1_lab}→{m2_lab}→{y_lab}",
        "a1": a1, "d21": d21, "a2": a2, "b1": b1, "b2": b2,
        "c": c, "c_prime": c_prime,
        "ind_serial": ind_serial,
        "ind_m1": ind_m1,
        "ind_m2": ind_m2,
        "ind_total": ind_total,
        "boot_serial_ci": (s_lo, s_hi),
        "boot_serial_p": s_p,
    }


def serial_3mediator(y, x, m1, m2, m3,
                     x_lab, m1_lab, m2_lab, m3_lab, y_lab, n_boot=5000):
    """Full serial mediation: X → M1 → M2 → M3 → Y"""
    print(f"\n{'='*60}")
    print(f"全逐次媒介: {x_lab} → {m1_lab} → {m2_lab} → {m3_lab} → {y_lab}")
    print(f"{'='*60}")

    X_const = sm.add_constant(x)

    # a1: X → M1
    mod_m1 = sm.OLS(m1, X_const).fit()
    a1 = mod_m1.params.iloc[1]

    # d21, a2: X,M1 → M2
    XM1 = sm.add_constant(pd.DataFrame({"x": x, "m1": m1}))
    mod_m2 = sm.OLS(m2, XM1).fit()
    d21 = mod_m2.params["m1"]
    a2 = mod_m2.params["x"]

    # d31, d32, a3: X,M1,M2 → M3
    XM1M2 = sm.add_constant(pd.DataFrame({"x": x, "m1": m1, "m2": m2}))
    mod_m3 = sm.OLS(m3, XM1M2).fit()
    d31 = mod_m3.params["m1"]
    d32 = mod_m3.params["m2"]
    a3 = mod_m3.params["x"]

    # b1, b2, b3, c': X,M1,M2,M3 → Y
    XM_all = sm.add_constant(pd.DataFrame({"x": x, "m1": m1, "m2": m2, "m3": m3}))
    if y.nunique() == 2:
        mod_y = sm.Logit(y, XM_all).fit(disp=0, maxiter=100)
    else:
        mod_y = sm.OLS(y, XM_all).fit()
    b1 = mod_y.params["m1"]
    b2 = mod_y.params["m2"]
    b3 = mod_y.params["m3"]
    c_prime = mod_y.params["x"]

    # Total
    if y.nunique() == 2:
        mod_total = sm.Logit(y, X_const).fit(disp=0)
    else:
        mod_total = sm.OLS(y, X_const).fit()
    c = mod_total.params.iloc[1]

    # Serial indirect: X→M1→M2→M3→Y
    serial_full = a1 * d21 * d32 * b3
    # All serial 2-step paths
    serial_m1m2 = a1 * d21 * b2
    serial_m1m3 = a1 * d31 * b3
    serial_m2m3 = a2 * d32 * b3
    # Parallel
    par_m1 = a1 * b1
    par_m2 = a2 * b2
    par_m3 = a3 * b3

    ind_total = serial_full + serial_m1m2 + serial_m1m3 + serial_m2m3 + par_m1 + par_m2 + par_m3

    print(f"\n  【パス係数】")
    print(f"    a1 ({x_lab}→{m1_lab}):       {a1:.4f}")
    print(f"    d21 ({m1_lab}→{m2_lab}):      {d21:.4f}")
    print(f"    d31 ({m1_lab}→{m3_lab}):      {d31:.4f}")
    print(f"    d32 ({m2_lab}→{m3_lab}):      {d32:.4f}")
    print(f"    a2 ({x_lab}→{m2_lab}):       {a2:.4f}")
    print(f"    a3 ({x_lab}→{m3_lab}):       {a3:.4f}")
    print(f"    b1 ({m1_lab}→{y_lab}):       {b1:.4f}")
    print(f"    b2 ({m2_lab}→{y_lab}):       {b2:.4f}")
    print(f"    b3 ({m3_lab}→{y_lab}):       {b3:.4f}")
    print(f"    c' (直接):                    {c_prime:.4f}")
    print(f"    c  (総効果):                  {c:.4f}")

    print(f"\n  【間接効果の分解】")
    print(f"    全逐次 (a1×d21×d32×b3):      {serial_full:.4f}  [{x_lab}→{m1_lab}→{m2_lab}→{m3_lab}→{y_lab}]")
    print(f"    逐次 M1→M2 (a1×d21×b2):      {serial_m1m2:.4f}")
    print(f"    逐次 M1→M3 (a1×d31×b3):      {serial_m1m3:.4f}")
    print(f"    逐次 M2→M3 (a2×d32×b3):      {serial_m2m3:.4f}")
    print(f"    並列 M1 (a1×b1):              {par_m1:.4f}")
    print(f"    並列 M2 (a2×b2):              {par_m2:.4f}")
    print(f"    並列 M3 (a3×b3):              {par_m3:.4f}")
    print(f"    間接効果合計:                  {ind_total:.4f}")
    print(f"    直接効果 (c'):                 {c_prime:.4f}")

    # Bootstrap for full serial path
    rng = np.random.default_rng(42)
    n = len(y)
    boot_vals = np.zeros(n_boot)
    boot_total = np.zeros(n_boot)

    for i in range(n_boot):
        idx = rng.choice(n, size=n, replace=True)
        xb = x.iloc[idx].reset_index(drop=True)
        m1b = m1.iloc[idx].reset_index(drop=True)
        m2b = m2.iloc[idx].reset_index(drop=True)
        m3b = m3.iloc[idx].reset_index(drop=True)
        yb = y.iloc[idx].reset_index(drop=True)
        try:
            Xb = sm.add_constant(xb)
            r1 = sm.OLS(m1b, Xb).fit()

            XM1b = sm.add_constant(pd.DataFrame({"x": xb, "m1": m1b}))
            r2 = sm.OLS(m2b, XM1b).fit()

            XM12b = sm.add_constant(pd.DataFrame({"x": xb, "m1": m1b, "m2": m2b}))
            r3 = sm.OLS(m3b, XM12b).fit()

            XMab = sm.add_constant(pd.DataFrame({"x": xb, "m1": m1b, "m2": m2b, "m3": m3b}))
            if yb.nunique() == 2:
                r4 = sm.Logit(yb, XMab).fit(disp=0, maxiter=50)
            else:
                r4 = sm.OLS(yb, XMab).fit()

            boot_vals[i] = r1.params.iloc[1] * r2.params["m1"] * r3.params["m2"] * r4.params["m3"]
            # total indirect
            a1b = r1.params.iloc[1]
            d21b = r2.params["m1"]; a2b = r2.params["x"]
            d31b = r3.params["m1"]; d32b = r3.params["m2"]; a3b = r3.params["x"]
            b1b = r4.params["m1"]; b2b = r4.params["m2"]; b3b = r4.params["m3"]
            boot_total[i] = (a1b*d21b*d32b*b3b + a1b*d21b*b2b + a1b*d31b*b3b +
                             a2b*d32b*b3b + a1b*b1b + a2b*b2b + a3b*b3b)
        except Exception:
            boot_vals[i] = np.nan
            boot_total[i] = np.nan

    valid = boot_vals[~np.isnan(boot_vals)]
    valid_t = boot_total[~np.isnan(boot_total)]

    if len(valid) >= 100:
        ci_lo, ci_hi = np.percentile(valid, [2.5, 97.5])
        bp = min(np.mean(valid <= 0), np.mean(valid >= 0)) * 2
        t_lo, t_hi = np.percentile(valid_t, [2.5, 97.5])
        tp = min(np.mean(valid_t <= 0), np.mean(valid_t >= 0)) * 2
    else:
        ci_lo = ci_hi = bp = t_lo = t_hi = tp = np.nan

    print(f"\n  【Bootstrap CI (n={n_boot})】")
    print(f"    全逐次間接:  {serial_full:.4f}  [{ci_lo:.4f}, {ci_hi:.4f}]  p≈{bp:.4f}")
    print(f"    間接合計:    {ind_total:.4f}  [{t_lo:.4f}, {t_hi:.4f}]  p≈{tp:.4f}")

    sig = "有意" if (ci_lo > 0 or ci_hi < 0) else "非有意"
    print(f"\n  → 全逐次間接効果: {sig}")

    return {
        "serial_full": serial_full,
        "boot_ci": (ci_lo, ci_hi),
        "boot_p": bp,
        "ind_total": ind_total,
    }


if __name__ == "__main__":
    df = load_data()
    y = df["outcome_binary"]
    print(f"データセット: N={len(df)}")

    # --- 2-mediator serial models ---
    results_2m = []

    # Path A: ストック優位 → 技術停滞 → 制度劣化 → 征服
    r = serial_2mediator(
        y, df["dominant_binary"], df["tech_position"], df["institutional_quality"],
        "ストック優位", "技術水準", "制度品質", "征服"
    )
    results_2m.append(r)

    # Path B: ストック優位 → 制度劣化 → 外部脅威↑ → 征服
    r = serial_2mediator(
        y, df["dominant_binary"], df["institutional_quality"], df["external_threat"],
        "ストック優位", "制度品質", "外部脅威", "征服"
    )
    results_2m.append(r)

    # Path C: ストック優位 → 技術停滞 → 外部脅威↑ → 征服
    r = serial_2mediator(
        y, df["dominant_binary"], df["tech_position"], df["external_threat"],
        "ストック優位", "技術水準", "外部脅威", "征服"
    )
    results_2m.append(r)

    # --- 3-mediator full serial ---
    # ストック優位 → 技術停滞 → 制度劣化 → 外部脅威↑ → 征服
    r3 = serial_3mediator(
        y, df["dominant_binary"],
        df["tech_position"], df["institutional_quality"], df["external_threat"],
        "ストック優位", "技術水準", "制度品質", "外部脅威", "征服"
    )

    # --- Summary ---
    print("\n\n" + "=" * 70)
    print("逐次媒介分析サマリー")
    print("=" * 70)

    print(f"\n  【2媒介逐次モデル】")
    print(f"  {'パス':55s} {'逐次間接効果':>12s} {'Boot CI':>24s} {'Boot p':>8s}")
    print(f"  {'-'*100}")
    for r in results_2m:
        ci = r["boot_serial_ci"]
        print(f"  {r['path']:55s} {r['ind_serial']:>12.4f} [{ci[0]:.4f}, {ci[1]:.4f}] {r['boot_serial_p']:>8.4f}")

    print(f"\n  【3媒介全逐次モデル】")
    print(f"  ストック優位 → 技術水準 → 制度品質 → 外部脅威 → 征服")
    print(f"    全逐次間接効果: {r3['serial_full']:.4f}")
    ci3 = r3["boot_ci"]
    print(f"    Bootstrap CI:   [{ci3[0]:.4f}, {ci3[1]:.4f}]")
    print(f"    Bootstrap p:    {r3['boot_p']:.4f}")
    print(f"    間接効果合計:   {r3['ind_total']:.4f}")

    print(f"\n  【解釈】")
    if ci3[0] > 0 or ci3[1] < 0:
        print(f"    全逐次パス（技術停滞→制度劣化→外部脅威→征服）は統計的に有意。")
        print(f"    ストック優位国の征服メカニズムに明確な因果連鎖が存在する。")
    else:
        print(f"    全逐次パスは非有意。並列的な媒介経路の方が支配的である可能性。")

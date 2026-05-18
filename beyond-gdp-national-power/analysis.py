"""
Beyond GDP: Stock vs Flow National Power — Full Statistical Analysis

Step 1: 拡張データで混同行列 + 統計量
Step 2: ロジスティック回帰（連続変数化）
Step 3: 交絡因子を統制した多変量分析
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import scipy.stats as stats
from scipy.stats import norm, ncx2
import math
import statsmodels.api as sm
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from data import load_data

# ============================================================
# STEP 1: 混同行列 + 統計量
# ============================================================

def confusion_matrix_analysis(df):
    print("=" * 70)
    print("STEP 1: 混同行列分析（拡張データ N={}）".format(len(df)))
    print("=" * 70)

    ct = pd.crosstab(
        df["dominant"].map({"stock": "ストック優位", "flow": "フロー優位"}),
        df["outcome"].map({"overtaken": "征服・崩壊", "disrupted": "征服・崩壊", "survived": "存続・変容"}),
    )
    # Ensure order
    ct = ct.reindex(
        index=["ストック優位", "フロー優位"],
        columns=["征服・崩壊", "存続・変容"],
    )
    print("\n【混同行列】")
    print(ct.to_string())
    print(f"  合計: {len(df)}")

    TP = ct.loc["ストック優位", "征服・崩壊"]
    FP = ct.loc["ストック優位", "存続・変容"]
    FN = ct.loc["フロー優位", "征服・崩壊"]
    TN = ct.loc["フロー優位", "存続・変容"]
    N = TP + FP + FN + TN

    print(f"\n  TP(ストック&征服)={TP}, FP(ストック&存続)={FP}")
    print(f"  FN(フロー&征服)={FN}, TN(フロー&存続)={TN}")

    # --- Basic rates ---
    stock_rate = TP / (TP + FP)
    flow_rate = FN / (FN + TN)
    print(f"\n【基本統計】")
    print(f"  全体征服率 (Prevalence): {(TP+FN)/N:.1%} ({TP+FN}/{N})")
    print(f"  ストック優位の征服率:    {stock_rate:.1%} ({TP}/{TP+FP})")
    print(f"  フロー優位の征服率:      {flow_rate:.1%} ({FN}/{FN+TN})")

    # --- Diagnostic stats ---
    sensitivity = TP / (TP + FN)
    specificity = TN / (TN + FP)
    ppv = TP / (TP + FP)
    npv = TN / (TN + FN)
    accuracy = (TP + TN) / N
    f1 = 2 * ppv * sensitivity / (ppv + sensitivity)
    print(f"\n【診断的統計量】")
    print(f"  感度 (Sensitivity):    {sensitivity:.1%}")
    print(f"  特異度 (Specificity):  {specificity:.1%}")
    print(f"  陽性的中率 (PPV):      {ppv:.1%}")
    print(f"  陰性的中率 (NPV):      {npv:.1%}")
    print(f"  正確度 (Accuracy):     {accuracy:.1%}")
    print(f"  F1スコア:              {f1:.3f}")

    # --- Risk measures ---
    rr = stock_rate / flow_rate
    rd = stock_rate - flow_rate
    nnt = 1 / rd if rd != 0 else float("inf")
    odds_ratio = (TP * TN) / (FP * FN)
    print(f"\n【リスク指標】")
    print(f"  リスク比 (RR):         {rr:.3f}")
    print(f"  リスク差 (RD):         {rd:+.1%}")
    print(f"  NNT:                   {nnt:.1f}")
    print(f"  オッズ比 (OR):         {odds_ratio:.3f}")

    # OR CI
    log_or = math.log(odds_ratio)
    se_log_or = math.sqrt(1 / TP + 1 / FP + 1 / FN + 1 / TN)
    ci_lo = math.exp(log_or - 1.96 * se_log_or)
    ci_hi = math.exp(log_or + 1.96 * se_log_or)
    print(f"  OR 95%CI (Woolf):      [{ci_lo:.3f}, {ci_hi:.3f}]")

    # --- Tests ---
    table = np.array([[TP, FP], [FN, TN]])
    _, p_fisher_one = stats.fisher_exact(table, alternative="greater")
    _, p_fisher_two = stats.fisher_exact(table, alternative="two-sided")
    chi2_y, p_chi2_y, _, _ = stats.chi2_contingency(table, correction=True)
    chi2_n, p_chi2_n, _, _ = stats.chi2_contingency(table, correction=False)
    print(f"\n【統計的検定】")
    print(f"  Fisher正確検定(片側): p = {p_fisher_one:.4f}")
    print(f"  Fisher正確検定(両側): p = {p_fisher_two:.4f}")
    print(f"  χ²検定(Yates補正):   χ²={chi2_y:.3f}, p={p_chi2_y:.4f}")
    print(f"  χ²検定(補正なし):    χ²={chi2_n:.3f}, p={p_chi2_n:.4f}")

    # --- Effect sizes ---
    phi = (TP * TN - FP * FN) / math.sqrt(
        (TP + FP) * (FN + TN) * (TP + FN) * (FP + TN)
    )
    j = sensitivity + specificity - 1
    lr_pos = sensitivity / (1 - specificity)
    lr_neg = (1 - sensitivity) / specificity
    mcc = phi  # same for 2x2
    cohen_label = "小" if abs(phi) < 0.3 else "中" if abs(phi) < 0.5 else "大"
    print(f"\n【効果量・相関】")
    print(f"  ファイ係数 (φ):       {phi:.3f} (効果量: {cohen_label})")
    print(f"  Cramér's V:           {abs(phi):.3f}")
    print(f"  MCC:                  {mcc:.3f}")
    print(f"  Youden's J:           {j:.3f}")
    print(f"  陽性尤度比 (LR+):     {lr_pos:.3f}")
    print(f"  陰性尤度比 (LR-):     {lr_neg:.3f}")

    # --- Power ---
    w = abs(phi)
    z_alpha = norm.ppf(0.95)
    ncp = N * w ** 2
    crit = stats.chi2.ppf(0.95, df=1)
    power = 1 - ncx2.cdf(crit, df=1, nc=ncp)
    z_beta = norm.ppf(0.80)
    n_req = int(np.ceil(((z_alpha + z_beta) / w) ** 2))
    print(f"\n【検出力分析】")
    print(f"  N = {N}, 効果量 w = {w:.3f}")
    print(f"  現在の検出力 (α=0.05): {power:.3f}")
    print(f"  80%検出力に必要なN:    約{n_req}")

    return {
        "TP": TP, "FP": FP, "FN": FN, "TN": TN,
        "OR": odds_ratio, "phi": phi, "p_fisher": p_fisher_one,
    }


# ============================================================
# STEP 2: ロジスティック回帰（連続変数）
# ============================================================

def logistic_univariate(df):
    print("\n" + "=" * 70)
    print("STEP 2: ロジスティック回帰（単変量）")
    print("=" * 70)

    y = df["outcome_binary"]

    predictors = [
        ("stock_index", "ストック指数"),
        ("trade_openness", "貿易開放度"),
        ("dominant_binary", "ストック優位(二値)"),
        ("geo_barrier", "地理的障壁"),
        ("external_threat", "外部脅威"),
        ("relative_pop", "相対人口"),
        ("tech_position", "技術水準"),
        ("institutional_quality", "制度品質"),
        ("regime_duration_yrs", "体制期間(年)"),
        ("has_external_patron", "外部保護者"),
        ("era_code", "時代コード"),
    ]

    results = []
    for var, label in predictors:
        X = sm.add_constant(df[var].astype(float))
        try:
            model = sm.Logit(y, X).fit(disp=0)
            coef = model.params.iloc[1]
            p_val = model.pvalues.iloc[1]
            or_val = np.exp(coef)
            ci = model.conf_int().iloc[1]
            or_lo, or_hi = np.exp(ci[0]), np.exp(ci[1])
            aic = model.aic
            results.append({
                "変数": label,
                "係数": coef,
                "OR": or_val,
                "OR_95CI_lo": or_lo,
                "OR_95CI_hi": or_hi,
                "p値": p_val,
                "AIC": aic,
            })
        except Exception as e:
            results.append({
                "変数": label,
                "係数": np.nan,
                "OR": np.nan,
                "OR_95CI_lo": np.nan,
                "OR_95CI_hi": np.nan,
                "p値": np.nan,
                "AIC": np.nan,
            })

    res_df = pd.DataFrame(results)
    print("\n【単変量ロジスティック回帰結果】")
    print(res_df.to_string(index=False, float_format=lambda x: f"{x:.3f}"))
    print("\n注: OR>1 → 征服リスク増加, OR<1 → 存続傾向")
    return res_df


# ============================================================
# STEP 3: 多変量ロジスティック回帰
# ============================================================

def logistic_multivariate(df):
    print("\n" + "=" * 70)
    print("STEP 3: 多変量ロジスティック回帰（交絡因子統制）")
    print("=" * 70)

    y = df["outcome_binary"]

    def fit_and_print(name, covariates):
        """Fit penalized logistic regression (Firth-like via statsmodels regularized fit),
        falling back to sklearn L2 if statsmodels fails."""
        print(f"\n--- {name} ---")
        X_raw = df[covariates].astype(float)
        X_sm = sm.add_constant(X_raw)

        # Try statsmodels with regularization to handle quasi-separation
        try:
            model = sm.Logit(y, X_sm).fit_regularized(
                method="l1", alpha=0.1, disp=0, trim_mode="off"
            )
            print(f"  Method: L1-penalized Logit (α=0.1)")
            print(f"  Log-Likelihood: {model.llf:.3f}")
            print(f"  Pseudo R²: N/A (regularized)")
            print(f"  N: {len(y)}")
            print(f"\n  {'変数':30s} {'係数':>8s} {'OR':>10s}")
            print(f"  {'-'*50}")
            for var in covariates:
                coef = model.params[var]
                or_val = np.exp(coef)
                print(f"  {var:30s} {coef:>8.3f} {or_val:>10.3f}")
            return model
        except Exception:
            pass

        # Fallback: sklearn L2-regularized
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X_raw)
        clf = LogisticRegression(penalty="l2", C=1.0, max_iter=1000, solver="lbfgs")
        clf.fit(X_scaled, y)
        score = clf.score(X_scaled, y)
        print(f"  Method: L2-regularized Logit (sklearn, C=1.0)")
        print(f"  Accuracy: {score:.3f}")
        print(f"  N: {len(y)}")
        print(f"\n  {'変数':30s} {'係数(標準化)':>12s} {'OR(標準化)':>12s}")
        print(f"  {'-'*56}")
        for i, var in enumerate(covariates):
            coef = clf.coef_[0][i]
            or_val = np.exp(coef)
            print(f"  {var:30s} {coef:>12.3f} {or_val:>12.3f}")
        return clf

    # --- Model A: ストック優位(二値) + 主要交絡因子 ---
    model_a = fit_and_print("Model A: ストック優位(二値) + 交絡因子", [
        "dominant_binary", "geo_barrier", "external_threat",
        "tech_position", "institutional_quality", "era_code",
        "has_external_patron",
    ])

    # --- Model B: 貿易開放度(連続) + 交絡因子 ---
    model_b = fit_and_print("Model B: 貿易開放度(連続) + 交絡因子", [
        "trade_openness", "geo_barrier", "external_threat",
        "tech_position", "institutional_quality", "era_code",
        "has_external_patron",
    ])

    # --- Model C: ストック指数(連続) + 交絡因子 ---
    model_c = fit_and_print("Model C: ストック指数(連続) + 交絡因子", [
        "stock_index", "geo_barrier", "external_threat",
        "tech_position", "institutional_quality", "era_code",
        "has_external_patron",
    ])

    # --- Model D: Full model (regularized) ---
    model_d = fit_and_print("Model D: フルモデル (regularized)", [
        "stock_index", "trade_openness", "geo_barrier",
        "external_threat", "relative_pop", "tech_position",
        "institutional_quality", "era_code", "has_external_patron",
        "regime_duration_yrs",
    ])

    # --- Also run unpenalized statsmodels for A-C for proper p-values ---
    print("\n\n--- 非正則化モデル（p値算出用） ---")
    sm_models = {}
    for name, covs in [
        ("A", ["dominant_binary", "geo_barrier", "external_threat",
               "institutional_quality", "era_code", "has_external_patron"]),
        ("B", ["trade_openness", "geo_barrier", "external_threat",
               "institutional_quality", "era_code", "has_external_patron"]),
        ("C", ["stock_index", "geo_barrier", "external_threat",
               "institutional_quality", "era_code", "has_external_patron"]),
    ]:
        X = sm.add_constant(df[covs].astype(float))
        try:
            m = sm.Logit(y, X).fit(disp=0, maxiter=100)
            sm_models[name] = m
            print(f"\n  Model {name} (reduced, unpenalized):")
            print(f"    AIC={m.aic:.1f}, BIC={m.bic:.1f}, Pseudo-R²={m.prsquared:.3f}")
            print(f"    {'変数':30s} {'OR':>10s} {'95%CI':>24s} {'p値':>8s}")
            print(f"    {'-'*74}")
            for var in covs:
                coef = m.params[var]
                p = m.pvalues[var]
                or_val = np.exp(coef)
                ci = m.conf_int().loc[var]
                print(f"    {var:30s} {or_val:>10.3f} [{np.exp(ci[0]):>8.3f}, {np.exp(ci[1]):>8.3f}] {p:>8.4f}")
        except Exception as e:
            print(f"\n  Model {name}: convergence failed ({e})")

    if sm_models:
        print("\n\n【モデル比較 (非正則化)】")
        print(f"  {'Model':<10} {'AIC':>8} {'BIC':>8} {'Pseudo-R²':>10} {'N':>5}")
        for name, m in sm_models.items():
            print(f"  {name:<10} {m.aic:>8.1f} {m.bic:>8.1f} {m.prsquared:>10.3f} {int(m.nobs):>5}")

    return {"A": model_a, "B": model_b, "C": model_c, "D": model_d, "sm": sm_models}


# ============================================================
# Main
# ============================================================

if __name__ == "__main__":
    df = load_data()
    print(f"データセット: N={len(df)}, ストック優位={sum(df['dominant']=='stock')}, "
          f"フロー優位={sum(df['dominant']=='flow')}")
    print(f"征服={sum(df['outcome']=='overtaken')}, 崩壊={sum(df['outcome']=='disrupted')}, 存続={sum(df['outcome']=='survived')}")

    # Step 1
    cm_results = confusion_matrix_analysis(df)

    # Step 2
    uni_results = logistic_univariate(df)

    # Step 3
    multi_results = logistic_multivariate(df)

    print("\n" + "=" * 70)
    print("分析完了")
    print("=" * 70)

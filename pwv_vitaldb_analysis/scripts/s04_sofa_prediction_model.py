"""
04_sofa_prediction_model.py
SOFAスコアへのPWV組み込みと死亡退院予測モデルの比較検証

分析モデル:
1. Baseline SOFA (従来の構成要素のみ)
2. SOFA + PWV (PAT/PTTを追加特徴量として追加)
3. SOFA-CV replaced (心血管コンポーネントをPWV関連指標で置換)
4. PWV-only model (PWV関連指標のみ)

評価指標:
- AUROC (Area Under Receiver Operating Characteristic)
- AUPRC (Area Under Precision-Recall Curve)
- Calibration (Hosmer-Lemeshow / Brier Score)
- Net Reclassification Index (NRI)
"""

import os
import warnings

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.model_selection import StratifiedKFold, cross_val_predict
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import GradientBoostingClassifier, RandomForestClassifier
from sklearn.preprocessing import StandardScaler
from sklearn.impute import SimpleImputer
from sklearn.metrics import (
    roc_auc_score, average_precision_score, brier_score_loss,
    roc_curve, precision_recall_curve, classification_report,
    confusion_matrix,
)
from sklearn.pipeline import Pipeline

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

warnings.filterwarnings("ignore")

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
FIGURE_DIR = os.path.join(OUTPUT_DIR, "figures")
os.makedirs(FIGURE_DIR, exist_ok=True)


# ============================================================
# SOFA Score Calculation from VitalDB Data
# ============================================================

def calculate_partial_sofa(clinical_df, pwv_case_df):
    """VitalDB利用可能データからの部分的SOFAスコア算出

    VitalDBで利用可能なSOFAコンポーネント:
    - Cardiovascular: MAP (from ABP), vasopressors (intraop_eph, intraop_phe, intraop_epi)
    - Respiration: preop_pao2 + FiO2 (limited)
    - Coagulation: preop_plt (platelets)
    - Renal: preop_cr (creatinine)
    - Liver: preop_ast, preop_alt (proxy, not bilirubin)
    - CNS: not available (no GCS)
    """
    merged = clinical_df.merge(pwv_case_df, on="caseid", how="inner")

    # Cardiovascular SOFA
    def cv_sofa(row):
        map_val = row.get("abp_avg", np.nan)
        if pd.isna(map_val):
            return np.nan
        epi = row.get("intraop_epi", 0) or 0
        phe = row.get("intraop_phe", 0) or 0
        eph = row.get("intraop_eph", 0) or 0
        has_vasopressor = (epi > 0) or (phe > 0) or (eph > 0)

        # Check vasopressor use first (higher severity) before MAP-only
        if has_vasopressor and epi > 0.1:
            return 3
        elif has_vasopressor:
            return 2
        elif map_val < 70:
            return 1
        else:
            return 0

    # Coagulation SOFA
    def coag_sofa(plt_val):
        if pd.isna(plt_val):
            return np.nan
        plt_val = float(plt_val)
        if plt_val >= 150:
            return 0
        elif plt_val >= 100:
            return 1
        elif plt_val >= 50:
            return 2
        elif plt_val >= 20:
            return 3
        return 4

    # Renal SOFA
    def renal_sofa(cr_val):
        if pd.isna(cr_val):
            return np.nan
        cr_val = float(cr_val)
        if cr_val < 1.2:
            return 0
        elif cr_val < 2.0:
            return 1
        elif cr_val < 3.5:
            return 2
        elif cr_val < 5.0:
            return 3
        return 4

    # Respiratory SOFA (approximation)
    def resp_sofa(pao2, sao2):
        if pd.notna(pao2) and float(pao2) > 0:
            ratio = float(pao2) / 0.21  # assume room air if FiO2 not available
            if ratio > 400:
                return 0
            elif ratio > 300:
                return 1
            elif ratio > 200:
                return 2
            elif ratio > 100:
                return 3
            return 4
        elif pd.notna(sao2):
            sao2 = float(sao2)
            if sao2 >= 97:
                return 0
            elif sao2 >= 94:
                return 1
            elif sao2 >= 90:
                return 2
            return 3
        return np.nan

    merged["sofa_cv"] = merged.apply(cv_sofa, axis=1)
    merged["sofa_coag"] = merged["preop_plt"].apply(coag_sofa)
    merged["sofa_renal"] = merged["preop_cr"].apply(renal_sofa)
    merged["sofa_resp"] = merged.apply(lambda r: resp_sofa(r.get("preop_pao2"), r.get("preop_sao2")), axis=1)

    sofa_cols = ["sofa_cv", "sofa_coag", "sofa_renal", "sofa_resp"]
    merged["sofa_partial"] = merged[sofa_cols].sum(axis=1, min_count=2)

    return merged


def prepare_features(df):
    """特徴量の準備

    Uses raw clinical values (with imputation) instead of derived SOFA scores
    to maximize usable sample size.

    Returns:
        dict of feature sets for different models
    """
    # SOFA-proxy raw features (available in VitalDB clinical data)
    # CV: MAP from waveform
    # Coagulation: preop_plt
    # Renal: preop_cr
    # Respiratory: preop_pao2 or preop_sao2
    sofa_raw_features = ["abp_avg", "preop_plt", "preop_cr"]
    sofa_raw_features_ext = ["abp_avg", "preop_plt", "preop_cr", "preop_pao2"]

    # PWV-derived features
    pwv_features = ["pat_mean_avg", "pat_mean_std", "pat_median_avg"]
    if "ptt_mean_avg" in df.columns:
        pwv_features.extend(["ptt_mean_avg", "ptt_mean_std", "ptt_median_avg"])

    # Clinical covariates
    clinical_features = ["age", "bmi", "asa", "emop"]
    if "hr_avg" in df.columns:
        clinical_features.append("hr_avg")
    if "hr_std" in df.columns:
        clinical_features.append("hr_std")

    feature_sets = {
        "sofa_raw": sofa_raw_features,
        "sofa_raw_plus_pwv": sofa_raw_features + pwv_features,
        "sofa_cv_replaced": [f for f in sofa_raw_features if f != "abp_avg"] + pwv_features,
        "pwv_only": pwv_features,
        "clinical_only": clinical_features,
        "clinical_plus_pwv": clinical_features + pwv_features,
        "full_model": sofa_raw_features + pwv_features + clinical_features,
        "full_no_pwv": sofa_raw_features + clinical_features,
    }

    return feature_sets


def train_and_evaluate(df, feature_sets, outcome_col="death_inhosp"):
    """各モデルの訓練と評価"""
    print(f"\n{'='*70}")
    print(f"Model Comparison: Predicting {outcome_col}")
    print(f"{'='*70}")

    valid_mask = df[outcome_col].notna() & (df[outcome_col].isin([0, 1]))
    df_valid = df[valid_mask].copy()

    y = df_valid[outcome_col].astype(int).values
    print(f"\nOutcome distribution: 0={np.sum(y==0)}, 1={np.sum(y==1)}")

    if np.sum(y == 1) < 5:
        print("WARNING: Very few positive cases. Results may be unreliable.")

    results = {}
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)

    for model_name, features in feature_sets.items():
        available_features = [f for f in features if f in df_valid.columns]
        if len(available_features) < 2:
            print(f"\n{model_name}: insufficient features ({len(available_features)})")
            continue

        X = df_valid[available_features].values

        # Only drop rows where ALL features are NaN (let imputer handle partial missing)
        all_nan_mask = np.isnan(X).all(axis=1)
        X_usable = X[~all_nan_mask]
        y_usable = y[~all_nan_mask]

        # Also require outcome is not nan
        if len(y_usable) < 20 or np.sum(y_usable == 1) < 3:
            print(f"\n{model_name}: insufficient samples (n={len(y_usable)}, events={np.sum(y_usable==1)})")
            continue

        X_clean = X_usable
        y_clean = y_usable

        # Imputation + Scaling pipeline
        pipe = Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scaler", StandardScaler()),
            ("clf", LogisticRegression(
                class_weight="balanced",
                max_iter=1000,
                C=0.1,
                random_state=42,
            )),
        ])

        try:
            y_prob = cross_val_predict(pipe, X_clean, y_clean, cv=cv, method="predict_proba")[:, 1]
            y_pred = (y_prob >= 0.5).astype(int)

            auroc = roc_auc_score(y_clean, y_prob)
            auprc = average_precision_score(y_clean, y_prob)
            brier = brier_score_loss(y_clean, y_prob)

            results[model_name] = {
                "features": available_features,
                "n_features": len(available_features),
                "n_samples": len(y_clean),
                "n_events": int(np.sum(y_clean == 1)),
                "auroc": auroc,
                "auprc": auprc,
                "brier_score": brier,
                "y_true": y_clean,
                "y_prob": y_prob,
            }

            print(f"\n{model_name}:")
            print(f"  Features: {available_features}")
            print(f"  N={len(y_clean)}, Events={np.sum(y_clean==1)}")
            print(f"  AUROC={auroc:.4f}, AUPRC={auprc:.4f}, Brier={brier:.4f}")

        except Exception as e:
            print(f"\n{model_name}: Error - {e}")
            continue

    # Also try GBM for full model
    for model_name in ["full_model", "sofa_raw_plus_pwv"]:
        if model_name not in feature_sets:
            continue
        available_features = [f for f in feature_sets[model_name] if f in df_valid.columns]
        if len(available_features) < 2:
            continue

        X = df_valid[available_features].values
        all_nan_mask = np.isnan(X).all(axis=1)
        X_clean = X[~all_nan_mask]
        y_clean = y[~all_nan_mask]

        if len(y_clean) < 20 or np.sum(y_clean == 1) < 3:
            continue

        gbm_name = f"{model_name}_gbm"
        pipe_gbm = Pipeline([
            ("imputer", SimpleImputer(strategy="median")),
            ("scaler", StandardScaler()),
            ("clf", GradientBoostingClassifier(
                n_estimators=100,
                max_depth=3,
                learning_rate=0.1,
                random_state=42,
            )),
        ])

        try:
            y_prob = cross_val_predict(pipe_gbm, X_clean, y_clean, cv=cv, method="predict_proba")[:, 1]
            auroc = roc_auc_score(y_clean, y_prob)
            auprc = average_precision_score(y_clean, y_prob)
            brier = brier_score_loss(y_clean, y_prob)

            results[gbm_name] = {
                "features": available_features,
                "n_features": len(available_features),
                "n_samples": len(y_clean),
                "n_events": int(np.sum(y_clean == 1)),
                "auroc": auroc,
                "auprc": auprc,
                "brier_score": brier,
                "y_true": y_clean,
                "y_prob": y_prob,
            }

            print(f"\n{gbm_name} (GBM):")
            print(f"  AUROC={auroc:.4f}, AUPRC={auprc:.4f}, Brier={brier:.4f}")

        except Exception as e:
            print(f"\n{gbm_name}: Error - {e}")

    return results


def compute_nri(results, baseline_name="sofa_only", test_name="sofa_plus_pwv"):
    """Net Reclassification Index (NRI) の計算"""
    if baseline_name not in results or test_name not in results:
        return None

    y_true = results[baseline_name]["y_true"]
    p_baseline = results[baseline_name]["y_prob"]
    p_test = results[test_name]["y_prob"]

    if len(y_true) != len(results[test_name]["y_true"]):
        print("NRI: sample size mismatch, skipping")
        return None

    threshold = 0.5
    cat_baseline = (p_baseline >= threshold).astype(int)
    cat_test = (p_test >= threshold).astype(int)

    events = y_true == 1
    non_events = y_true == 0

    # Event NRI
    up_events = np.sum((cat_test[events] > cat_baseline[events]))
    down_events = np.sum((cat_test[events] < cat_baseline[events]))
    n_events = np.sum(events)
    nri_events = (up_events - down_events) / n_events if n_events > 0 else 0

    # Non-event NRI
    up_non = np.sum((cat_test[non_events] > cat_baseline[non_events]))
    down_non = np.sum((cat_test[non_events] < cat_baseline[non_events]))
    n_non = np.sum(non_events)
    nri_non = (down_non - up_non) / n_non if n_non > 0 else 0

    nri = nri_events + nri_non

    print(f"\nNRI ({test_name} vs {baseline_name}):")
    print(f"  Event NRI: {nri_events:.4f}")
    print(f"  Non-event NRI: {nri_non:.4f}")
    print(f"  Total NRI: {nri:.4f}")

    return {"nri_events": nri_events, "nri_non_events": nri_non, "nri_total": nri}


def plot_roc_comparison(results):
    """ROC曲線の比較プロット"""
    fig, axes = plt.subplots(1, 2, figsize=(14, 6))

    # ROC curves
    ax = axes[0]
    colors = plt.cm.Set1(np.linspace(0, 1, len(results)))
    for (name, res), color in zip(results.items(), colors):
        fpr, tpr, _ = roc_curve(res["y_true"], res["y_prob"])
        ax.plot(fpr, tpr, label=f'{name} (AUC={res["auroc"]:.3f})', color=color, linewidth=2)
    ax.plot([0, 1], [0, 1], "k--", alpha=0.5)
    ax.set_xlabel("False Positive Rate")
    ax.set_ylabel("True Positive Rate")
    ax.set_title("ROC Curves: Mortality Prediction Models")
    ax.legend(loc="lower right", fontsize=8)
    ax.grid(True, alpha=0.3)

    # PR curves
    ax = axes[1]
    for (name, res), color in zip(results.items(), colors):
        prec, rec, _ = precision_recall_curve(res["y_true"], res["y_prob"])
        ax.plot(rec, prec, label=f'{name} (AP={res["auprc"]:.3f})', color=color, linewidth=2)
    prevalence = results[list(results.keys())[0]]["n_events"] / results[list(results.keys())[0]]["n_samples"]
    ax.axhline(y=prevalence, color="k", linestyle="--", alpha=0.5, label=f"Prevalence ({prevalence:.3f})")
    ax.set_xlabel("Recall")
    ax.set_ylabel("Precision")
    ax.set_title("Precision-Recall Curves")
    ax.legend(loc="upper right", fontsize=8)
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    figpath = os.path.join(FIGURE_DIR, "model_comparison_roc.png")
    plt.savefig(figpath, dpi=150)
    plt.close()
    print(f"\nROC comparison plot saved to {figpath}")
    return figpath


def plot_feature_importance(df, feature_sets):
    """特徴量重要度の可視化"""
    features = feature_sets.get("full_model", feature_sets.get("sofa_plus_pwv", []))
    available = [f for f in features if f in df.columns]
    if len(available) < 2:
        return None

    valid = df[["death_inhosp"] + available].dropna()
    if len(valid) < 20:
        return None

    X = valid[available].values
    y = valid["death_inhosp"].astype(int).values

    pipe = Pipeline([
        ("imputer", SimpleImputer(strategy="median")),
        ("scaler", StandardScaler()),
        ("clf", GradientBoostingClassifier(n_estimators=100, max_depth=3, random_state=42)),
    ])
    pipe.fit(X, y)

    importances = pipe.named_steps["clf"].feature_importances_
    sorted_idx = np.argsort(importances)[::-1]

    fig, ax = plt.subplots(figsize=(10, 6))
    feature_names = [available[i] for i in sorted_idx]
    ax.barh(range(len(feature_names)), importances[sorted_idx])
    ax.set_yticks(range(len(feature_names)))
    ax.set_yticklabels(feature_names)
    ax.set_xlabel("Feature Importance")
    ax.set_title("GBM Feature Importance (Full Model)")
    ax.invert_yaxis()
    plt.tight_layout()

    figpath = os.path.join(FIGURE_DIR, "feature_importance.png")
    plt.savefig(figpath, dpi=150)
    plt.close()
    print(f"Feature importance plot saved to {figpath}")
    return figpath


def generate_summary_table(results):
    """結果サマリーテーブルの生成"""
    rows = []
    for name, res in results.items():
        rows.append({
            "Model": name,
            "N Features": res["n_features"],
            "N Samples": res["n_samples"],
            "N Events": res["n_events"],
            "AUROC": f"{res['auroc']:.4f}",
            "AUPRC": f"{res['auprc']:.4f}",
            "Brier Score": f"{res['brier_score']:.4f}",
        })

    summary_df = pd.DataFrame(rows)
    outpath = os.path.join(OUTPUT_DIR, "model_comparison_summary.csv")
    summary_df.to_csv(outpath, index=False)
    print(f"\nModel comparison summary saved to {outpath}")
    print(summary_df.to_string(index=False))
    return summary_df


# ============================================================
# ICU Admission Prediction (Secondary Outcome)
# ============================================================

def predict_icu_admission(df, feature_sets):
    """ICU入室予測（副次的アウトカム）"""
    if "icu_days" not in df.columns:
        return None

    df_icu = df.copy()
    df_icu["icu_admitted"] = (df_icu["icu_days"] > 0).astype(int)

    print(f"\n{'='*70}")
    print("Secondary Outcome: ICU Admission Prediction")
    print(f"{'='*70}")

    results = train_and_evaluate(df_icu, feature_sets, outcome_col="icu_admitted")
    return results


def main():
    print("=" * 70)
    print("SOFA + PWV Mortality Prediction Model")
    print("=" * 70)

    # Load data
    pwv_df = pd.read_csv(os.path.join(OUTPUT_DIR, "pwv_results.csv"))
    clinical_df = pd.read_csv(os.path.join(DATA_DIR, "clinical_data.csv"))

    # Aggregate PWV to case level
    case_pwv = pwv_df.groupby("caseid").agg({
        "pat_mean": ["mean", "std", "median"],
    }).reset_index()
    case_pwv.columns = ["caseid", "pat_mean_avg", "pat_mean_std", "pat_median_avg"]

    if "ptt_mean" in pwv_df.columns:
        ptt_agg = pwv_df.groupby("caseid").agg({
            "ptt_mean": ["mean", "std", "median"],
        }).reset_index()
        ptt_agg.columns = ["caseid", "ptt_mean_avg", "ptt_mean_std", "ptt_median_avg"]
        case_pwv = case_pwv.merge(ptt_agg, on="caseid", how="left")

    if "hr_mean" in pwv_df.columns:
        hr_agg = pwv_df.groupby("caseid").agg({"hr_mean": ["mean", "std"]}).reset_index()
        hr_agg.columns = ["caseid", "hr_avg", "hr_std"]
        case_pwv = case_pwv.merge(hr_agg, on="caseid", how="left")

    if "abp_mean" in pwv_df.columns:
        abp_agg = pwv_df.groupby("caseid").agg({"abp_mean": ["mean", "std"]}).reset_index()
        abp_agg.columns = ["caseid", "abp_avg", "abp_std"]
        case_pwv = case_pwv.merge(abp_agg, on="caseid", how="left")

    # Calculate partial SOFA
    df = calculate_partial_sofa(clinical_df, case_pwv)
    print(f"\nMerged dataset: {len(df)} cases")
    print(f"Death in hospital: {df['death_inhosp'].sum()}")
    print(f"ICU admission: {(df['icu_days'] > 0).sum()}")

    # Prepare feature sets
    feature_sets = prepare_features(df)

    # Primary outcome: mortality prediction
    mortality_results = train_and_evaluate(df, feature_sets, outcome_col="death_inhosp")

    if len(mortality_results) > 1:
        # NRI analysis
        nri = compute_nri(mortality_results, "sofa_raw", "sofa_raw_plus_pwv")

        # Plots
        plot_roc_comparison(mortality_results)
        plot_feature_importance(df, feature_sets)

    # Summary table
    if len(mortality_results) > 0:
        generate_summary_table(mortality_results)

    # Secondary outcome: ICU admission
    icu_results = predict_icu_admission(df, feature_sets)
    if icu_results and len(icu_results) > 0:
        icu_summary = []
        for name, res in icu_results.items():
            icu_summary.append({
                "Model": name,
                "AUROC": f"{res['auroc']:.4f}",
                "AUPRC": f"{res['auprc']:.4f}",
            })
        icu_df = pd.DataFrame(icu_summary)
        icu_df.to_csv(os.path.join(OUTPUT_DIR, "icu_prediction_summary.csv"), index=False)

    print("\nPrediction model analysis complete!")
    return mortality_results


if __name__ == "__main__":
    main()

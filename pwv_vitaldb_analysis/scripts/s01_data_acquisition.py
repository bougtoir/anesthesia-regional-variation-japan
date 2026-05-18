"""
01_data_acquisition.py
VitalDB Open Datasetからの波形・臨床データ取得スクリプト

対象データ:
- ECG lead II (500Hz) - R-peak検出用
- Arterial pressure wave (500Hz) - 動脈圧波形（橈骨動脈）
- Plethysmography wave (500Hz) - PPG波形
- CVP wave (500Hz) - 中心静脈圧波形（利用可能な場合）
- Femoral arterial pressure wave (500Hz) - 大腿動脈圧波形（利用可能な場合）
- 数値データ: MAP, SBP, DBP, HR, SpO2, CVP数値
- 臨床データ: 死亡退院、ICU日数、バイタルサイン、検査値

PWV算出方法:
1. PAT法: ECG R-peak → ABP foot (Pulse Arrival Time)
2. PTT法: ABP foot → PPG foot (Pulse Transit Time)
3. 2-site PWV: ART(橈骨) → FEM(大腿) foot-to-foot delay
"""

import os
import sys
import json
import time
import warnings
import traceback

import numpy as np
import pandas as pd
import vitaldb

warnings.filterwarnings("ignore")

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)

SAMPLE_RATE = 500  # Hz
SEGMENT_DURATION = 60  # seconds per analysis segment
MAX_CASES_PHASE1 = 300  # Phase 1: ensure death + ICU + normal balance
MAX_CASES_FULL = 1000  # Full analysis


def get_clinical_data():
    """VitalDB臨床データの取得"""
    print("Fetching clinical data from VitalDB API...")
    import requests
    resp = requests.get("https://api.vitaldb.net/cases")
    text = resp.content.decode("utf-8-sig")
    import io
    df = pd.read_csv(io.StringIO(text))
    outpath = os.path.join(DATA_DIR, "clinical_data.csv")
    df.to_csv(outpath, index=False)
    print(f"  Saved {len(df)} cases to {outpath}")
    print(f"  Columns: {list(df.columns)}")
    print(f"  death_inhosp distribution: {df['death_inhosp'].value_counts().to_dict()}")
    print(f"  ICU days > 0: {(df['icu_days'] > 0).sum()}")
    return df


def get_lab_data():
    """VitalDB検査データの取得"""
    print("Fetching lab data from VitalDB API...")
    import requests
    resp = requests.get("https://api.vitaldb.net/labs")
    text = resp.content.decode("utf-8-sig")
    import io
    df = pd.read_csv(io.StringIO(text))
    outpath = os.path.join(DATA_DIR, "lab_data.csv")
    df.to_csv(outpath, index=False)
    print(f"  Saved {len(df)} lab records to {outpath}")
    print(f"  Columns: {list(df.columns)}")
    return df


def find_eligible_cases():
    """解析対象ケースの選定"""
    print("\nFinding eligible cases...")

    # ECG + ART + PLETH (PAT/PTT calculation)
    basic_tracks = "SNUADC/ECG_II,SNUADC/ART,SNUADC/PLETH"
    cases_basic = vitaldb.find_cases(basic_tracks)
    print(f"  Cases with ECG+ART+PLETH: {len(cases_basic)}")

    # Cases with CVP waveform
    cvp_tracks = "SNUADC/ECG_II,SNUADC/ART,SNUADC/PLETH,SNUADC/CVP"
    cases_cvp = vitaldb.find_cases(cvp_tracks)
    print(f"  Cases with ECG+ART+PLETH+CVP: {len(cases_cvp)}")

    # Cases with femoral arterial (2-site PWV)
    fem_tracks = "SNUADC/ECG_II,SNUADC/ART,SNUADC/FEM"
    cases_fem = vitaldb.find_cases(fem_tracks)
    print(f"  Cases with ECG+ART+FEM (2-site PWV): {len(cases_fem)}")

    return {
        "basic": cases_basic,
        "cvp": cases_cvp,
        "femoral": cases_fem,
    }


def download_waveform_segment(caseid, tracks, duration_sec=300, offset_sec=None):
    """指定ケースの波形セグメントをダウンロード

    Args:
        caseid: VitalDB case ID
        tracks: list of track names
        duration_sec: segment duration in seconds
        offset_sec: start offset in seconds (None = middle of case)

    Returns:
        numpy array (samples x tracks), or None on failure
    """
    try:
        vf = vitaldb.VitalFile(caseid, tracks)
        df = vf.to_pandas(tracks, 1 / SAMPLE_RATE)

        if df is None or len(df) == 0:
            return None

        total_samples = len(df)
        total_sec = total_samples / SAMPLE_RATE

        if total_sec < duration_sec + 60:
            return None

        if offset_sec is None:
            offset_sec = max(60, (total_sec - duration_sec) / 2)

        start_idx = int(offset_sec * SAMPLE_RATE)
        end_idx = start_idx + int(duration_sec * SAMPLE_RATE)

        if end_idx > total_samples:
            return None

        segment = df.iloc[start_idx:end_idx]
        nan_frac = segment.isna().mean()
        if (nan_frac > 0.3).any():
            return None

        return segment

    except Exception as e:
        print(f"  Error downloading case {caseid}: {e}")
        return None


def download_numeric_data(caseid):
    """数値パラメータ（MAP, SBP, DBP, HR, SpO2, CVP）のダウンロード"""
    numeric_tracks = [
        "Solar8000/ART_MBP", "Solar8000/ART_SBP", "Solar8000/ART_DBP",
        "Solar8000/HR", "Solar8000/PLETH_SPO2", "Solar8000/CVP",
        "Solar8000/RR",
    ]
    try:
        vf = vitaldb.VitalFile(caseid, numeric_tracks)
        df = vf.to_pandas(numeric_tracks, 1)  # 1-second interval for numeric
        return df
    except Exception:
        return None


def process_batch(case_list, max_cases, batch_label="basic"):
    """バッチ処理: 波形データの取得と保存"""
    print(f"\nProcessing {batch_label} batch (max {max_cases} cases)...")

    if batch_label == "basic":
        wave_tracks = ["SNUADC/ECG_II", "SNUADC/ART", "SNUADC/PLETH"]
    elif batch_label == "cvp":
        wave_tracks = ["SNUADC/ECG_II", "SNUADC/ART", "SNUADC/PLETH", "SNUADC/CVP"]
    elif batch_label == "femoral":
        wave_tracks = ["SNUADC/ECG_II", "SNUADC/ART", "SNUADC/FEM"]
    else:
        wave_tracks = ["SNUADC/ECG_II", "SNUADC/ART", "SNUADC/PLETH"]

    results = []
    processed = 0
    failed = 0

    for i, caseid in enumerate(case_list):
        if processed >= max_cases:
            break

        if (i + 1) % 10 == 0:
            print(f"  Progress: {i+1}/{len(case_list)}, processed={processed}, failed={failed}")

        segment = download_waveform_segment(caseid, wave_tracks, duration_sec=300)
        if segment is None:
            failed += 1
            continue

        numeric = download_numeric_data(caseid)

        outpath = os.path.join(DATA_DIR, f"wave_{batch_label}_{caseid}.parquet")
        segment.to_parquet(outpath)

        if numeric is not None:
            num_path = os.path.join(DATA_DIR, f"numeric_{caseid}.parquet")
            numeric.to_parquet(num_path)

        results.append({
            "caseid": caseid,
            "batch": batch_label,
            "wave_samples": len(segment),
            "wave_duration_sec": len(segment) / SAMPLE_RATE,
            "tracks": wave_tracks,
            "has_numeric": numeric is not None,
        })
        processed += 1

    print(f"  Completed: {processed} cases, {failed} failures")
    return results


def main():
    print("=" * 70)
    print("VitalDB Data Acquisition for PWV Analysis")
    print("=" * 70)

    # Step 1: Clinical data
    clinical_df = get_clinical_data()

    # Step 2: Lab data
    lab_df = get_lab_data()

    # Step 3: Find eligible cases
    eligible = find_eligible_cases()

    # Step 4: Stratified sampling to ensure outcome representation
    death_cases = set(clinical_df[clinical_df["death_inhosp"] == 1]["caseid"].tolist())
    icu_cases = set(clinical_df[(clinical_df["icu_days"] > 0) & (clinical_df["death_inhosp"] != 1)]["caseid"].tolist())
    normal_cases = set(clinical_df[(clinical_df["icu_days"] == 0) & (clinical_df["death_inhosp"] != 1)]["caseid"].tolist())
    print(f"\nDeath cases: {len(death_cases)}, ICU-only: {len(icu_cases)}, Normal: {len(normal_cases)}")

    def stratified_order(case_list):
        """Order: all deaths first, then ICU cases, then normal cases (for balanced sampling)"""
        deaths = [c for c in case_list if c in death_cases]
        icus = [c for c in case_list if c in icu_cases]
        normals = [c for c in case_list if c in normal_cases]
        return deaths + icus + normals

    basic_ordered = stratified_order(eligible["basic"])
    cvp_ordered = stratified_order(eligible["cvp"])
    fem_ordered = stratified_order(eligible["femoral"])

    print(f"  Basic - deaths: {len([c for c in basic_ordered if c in death_cases])}, "
          f"ICU: {len([c for c in basic_ordered if c in icu_cases])}, "
          f"normal: {len([c for c in basic_ordered if c in normal_cases])}")

    # Step 5: Download waveform data in batches
    all_results = []

    # Phase 1: Basic PAT/PTT cases - ensure deaths + balanced ICU/normal
    results_basic = process_batch(basic_ordered, MAX_CASES_PHASE1, "basic")
    all_results.extend(results_basic)

    # Phase 2: CVP cases (subset)
    results_cvp = process_batch(cvp_ordered, min(100, len(cvp_ordered)), "cvp")
    all_results.extend(results_cvp)

    # Phase 3: Femoral cases (all available)
    results_fem = process_batch(fem_ordered, min(50, len(fem_ordered)), "femoral")
    all_results.extend(results_fem)

    # Save acquisition summary
    summary_df = pd.DataFrame(all_results)
    summary_path = os.path.join(DATA_DIR, "acquisition_summary.csv")
    summary_df.to_csv(summary_path, index=False)
    print(f"\nAcquisition summary saved to {summary_path}")
    print(f"Total cases acquired: {len(summary_df)}")
    print(f"  Basic: {len(summary_df[summary_df['batch']=='basic'])}")
    print(f"  CVP: {len(summary_df[summary_df['batch']=='cvp'])}")
    print(f"  Femoral: {len(summary_df[summary_df['batch']=='femoral'])}")

    # Save eligible case lists for reference
    eligible_path = os.path.join(DATA_DIR, "eligible_cases.json")
    with open(eligible_path, "w") as f:
        json.dump({k: v for k, v in eligible.items()}, f)

    print("\nData acquisition complete!")
    return summary_df


if __name__ == "__main__":
    main()

"""
02_pwv_calculation.py
波形解析によるPWV（脈波伝播速度）算出パイプライン

3つのPWV算出方法:
1. PAT法 (Pulse Arrival Time): ECG R-peak → ABP foot
   - 前射血時間(PEP)を含むため厳密にはPWVではないが、最も多くのデータで利用可能
2. PTT法 (Pulse Transit Time): ABP foot → PPG foot
   - PEPを除外した真のトランジットタイム
3. 2-site PWV: ART foot → FEM foot (橈骨→大腿動脈)
   - 真のPWV（距離/時間）
"""

import os
import sys
import warnings
import traceback

import numpy as np
import pandas as pd
from scipy import signal
from scipy.signal import find_peaks, butter, filtfilt, savgol_filter

warnings.filterwarnings("ignore")

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

SAMPLE_RATE = 500  # Hz


# ============================================================
# Signal Preprocessing
# ============================================================

def bandpass_filter(data, lowcut, highcut, fs, order=4):
    nyq = 0.5 * fs
    low = lowcut / nyq
    high = highcut / nyq
    low = max(low, 1e-5)
    high = min(high, 0.9999)
    b, a = butter(order, [low, high], btype="band")
    return filtfilt(b, a, data)


def lowpass_filter(data, cutoff, fs, order=4):
    nyq = 0.5 * fs
    normal_cutoff = min(cutoff / nyq, 0.9999)
    b, a = butter(order, normal_cutoff, btype="low")
    return filtfilt(b, a, data)


def remove_baseline_wander(ecg, fs):
    return bandpass_filter(ecg, 0.5, 40, fs)


def preprocess_abp(abp, fs):
    return lowpass_filter(abp, 20, fs)


def preprocess_ppg(ppg, fs):
    return bandpass_filter(ppg, 0.5, 8, fs)


# ============================================================
# R-peak Detection (Pan-Tompkins-like)
# ============================================================

def detect_r_peaks(ecg, fs):
    """ECG R-peak検出 (Pan-Tompkins法ベース)"""
    ecg_clean = remove_baseline_wander(ecg, fs)
    ecg_diff = np.diff(ecg_clean)
    ecg_squared = ecg_diff ** 2
    window_size = int(0.15 * fs)
    if window_size < 1:
        window_size = 1
    ecg_integrated = np.convolve(ecg_squared, np.ones(window_size) / window_size, mode="same")

    min_distance = int(0.3 * fs)  # min 300ms between R-peaks (200bpm max)
    threshold = np.mean(ecg_integrated) + 0.3 * np.std(ecg_integrated)

    peaks, properties = find_peaks(
        ecg_integrated,
        height=threshold,
        distance=min_distance,
    )

    if len(peaks) < 5:
        peaks, _ = find_peaks(
            ecg_integrated,
            height=np.percentile(ecg_integrated, 70),
            distance=min_distance,
        )

    # Refine to actual R-peak in original ECG
    refined_peaks = []
    search_window = int(0.05 * fs)  # ±50ms
    for p in peaks:
        start = max(0, p - search_window)
        end = min(len(ecg_clean), p + search_window)
        local_max = start + np.argmax(ecg_clean[start:end])
        refined_peaks.append(local_max)

    return np.array(refined_peaks)


# ============================================================
# ABP / PPG Foot Detection
# ============================================================

def detect_abp_feet(abp, fs, r_peaks=None):
    """動脈圧波形のfoot point（拡張末期点）検出

    各心拍の開始点（拡張末期の最低点）を検出する。
    R-peakのタイミングをガイドとして使用可能。
    """
    abp_smooth = preprocess_abp(abp, fs)
    abp_deriv = np.gradient(abp_smooth)

    feet = []
    if r_peaks is not None and len(r_peaks) > 1:
        for i in range(len(r_peaks) - 1):
            search_start = r_peaks[i]
            search_end = min(r_peaks[i] + int(0.5 * fs), r_peaks[i + 1])
            if search_end <= search_start:
                continue
            segment = abp_smooth[search_start:search_end]
            if len(segment) < 5:
                continue
            local_min_idx = search_start + np.argmin(segment)
            feet.append(local_min_idx)
    else:
        min_distance = int(0.4 * fs)
        inverted = -abp_smooth
        peaks_inv, _ = find_peaks(inverted, distance=min_distance)
        feet = peaks_inv.tolist()

    return np.array(feet)


def detect_ppg_feet(ppg, fs, r_peaks=None):
    """PPG波形のfoot point検出"""
    ppg_clean = preprocess_ppg(ppg, fs)

    feet = []
    if r_peaks is not None and len(r_peaks) > 1:
        for i in range(len(r_peaks) - 1):
            search_start = r_peaks[i]
            search_end = min(r_peaks[i] + int(0.6 * fs), r_peaks[i + 1])
            if search_end <= search_start:
                continue
            segment = ppg_clean[search_start:search_end]
            if len(segment) < 5:
                continue
            local_min_idx = search_start + np.argmin(segment)
            feet.append(local_min_idx)
    else:
        ppg_inverted = -ppg_clean
        min_distance = int(0.4 * fs)
        peaks_inv, _ = find_peaks(ppg_inverted, distance=min_distance)
        feet = peaks_inv.tolist()

    return np.array(feet)


# ============================================================
# PWV Calculation Methods
# ============================================================

def calculate_pat(ecg, abp, fs):
    """PAT法: ECG R-peak → ABP foot interval

    Returns:
        dict with 'pat_values' (array of PAT in ms),
        'r_peaks', 'abp_feet', 'hr_values'
    """
    r_peaks = detect_r_peaks(ecg, fs)
    abp_feet = detect_abp_feet(abp, fs, r_peaks)

    if len(r_peaks) < 5 or len(abp_feet) < 5:
        return None

    pat_values = []
    matched_r = []
    matched_abp = []

    for rp in r_peaks:
        candidates = abp_feet[(abp_feet > rp) & (abp_feet < rp + int(0.5 * fs))]
        if len(candidates) > 0:
            foot = candidates[0]
            pat_ms = (foot - rp) / fs * 1000  # ms
            if 50 < pat_ms < 400:  # physiological range
                pat_values.append(pat_ms)
                matched_r.append(rp)
                matched_abp.append(foot)

    if len(pat_values) < 5:
        return None

    rr_intervals = np.diff(np.array(matched_r)) / fs * 1000  # ms
    hr_values = 60000.0 / rr_intervals  # bpm

    return {
        "pat_values": np.array(pat_values),
        "r_peaks": np.array(matched_r),
        "abp_feet": np.array(matched_abp),
        "hr_values": hr_values,
        "pat_mean": np.mean(pat_values),
        "pat_std": np.std(pat_values),
        "pat_median": np.median(pat_values),
    }


def calculate_ptt(abp, ppg, fs, r_peaks=None):
    """PTT法: ABP foot → PPG foot interval

    Returns:
        dict with 'ptt_values' (array of PTT in ms)
    """
    if r_peaks is None:
        return None

    abp_feet = detect_abp_feet(abp, fs, r_peaks)
    ppg_feet = detect_ppg_feet(ppg, fs, r_peaks)

    if len(abp_feet) < 5 or len(ppg_feet) < 5:
        return None

    ptt_values = []
    matched_abp = []
    matched_ppg = []

    for af in abp_feet:
        candidates = ppg_feet[(ppg_feet > af) & (ppg_feet < af + int(0.3 * fs))]
        if len(candidates) > 0:
            pf = candidates[0]
            ptt_ms = (pf - af) / fs * 1000  # ms
            if 10 < ptt_ms < 300:  # physiological range
                ptt_values.append(ptt_ms)
                matched_abp.append(af)
                matched_ppg.append(pf)

    if len(ptt_values) < 5:
        return None

    return {
        "ptt_values": np.array(ptt_values),
        "abp_feet": np.array(matched_abp),
        "ppg_feet": np.array(matched_ppg),
        "ptt_mean": np.mean(ptt_values),
        "ptt_std": np.std(ptt_values),
        "ptt_median": np.median(ptt_values),
    }


def calculate_2site_pwv(art, fem, fs, r_peaks=None):
    """2-site PWV: ART(橈骨) foot → FEM(大腿) foot delay

    注: VitalDBでは橈骨動脈(ART)と大腿動脈(FEM)の波形が利用可能
    大腿動脈は心臓に近いため、通常FEM foot → ART footの順序
    PWV = 推定距離 / transit time

    Returns:
        dict with 'pwv_values' (array of PWV in m/s)
    """
    art_feet = detect_abp_feet(art, fs, r_peaks)
    fem_feet = detect_abp_feet(fem, fs, r_peaks)

    if len(art_feet) < 5 or len(fem_feet) < 5:
        return None

    # 大腿動脈→橈骨動脈のtransit time
    # 推定距離: ~0.6-0.8m (大腿→橈骨)
    ESTIMATED_DISTANCE_M = 0.7  # 推定値

    tt_values = []
    pwv_values = []

    for ff in fem_feet:
        candidates = art_feet[(art_feet > ff) & (art_feet < ff + int(0.2 * fs))]
        if len(candidates) > 0:
            af = candidates[0]
            tt_ms = (af - ff) / fs * 1000  # ms
            if 5 < tt_ms < 150:  # physiological range
                tt_values.append(tt_ms)
                pwv_ms = ESTIMATED_DISTANCE_M / (tt_ms / 1000)  # m/s
                if 3 < pwv_ms < 20:  # physiological PWV range
                    pwv_values.append(pwv_ms)

    if len(pwv_values) < 3:
        return None

    return {
        "tt_values": np.array(tt_values),
        "pwv_values": np.array(pwv_values),
        "pwv_mean": np.mean(pwv_values),
        "pwv_std": np.std(pwv_values),
        "pwv_median": np.median(pwv_values),
    }


# ============================================================
# Segment-wise Analysis
# ============================================================

def analyze_segment(segment_df, batch_type="basic"):
    """1セグメント（5分間）の波形解析

    Returns:
        dict with PWV metrics per 60-second window
    """
    cols = segment_df.columns.tolist()
    ecg = segment_df[cols[0]].values  # ECG_II
    abp = segment_df[cols[1]].values  # ART

    ecg = np.nan_to_num(ecg, nan=0.0)
    abp = np.nan_to_num(abp, nan=0.0)

    if abp.std() < 1 or ecg.std() < 0.001:
        return None

    results = []
    window_samples = 60 * SAMPLE_RATE  # 60-second windows
    n_windows = len(ecg) // window_samples

    for w in range(n_windows):
        start = w * window_samples
        end = start + window_samples
        ecg_win = ecg[start:end]
        abp_win = abp[start:end]

        window_result = {"window_idx": w, "window_start_sec": w * 60}

        # PAT calculation
        pat_result = calculate_pat(ecg_win, abp_win, SAMPLE_RATE)
        if pat_result is not None:
            window_result["pat_mean"] = pat_result["pat_mean"]
            window_result["pat_std"] = pat_result["pat_std"]
            window_result["pat_median"] = pat_result["pat_median"]
            window_result["pat_n_beats"] = len(pat_result["pat_values"])
            if len(pat_result["hr_values"]) > 0:
                window_result["hr_mean"] = np.mean(pat_result["hr_values"])

            # ABP statistics from matched beats
            abp_smooth = preprocess_abp(abp_win, SAMPLE_RATE)
            window_result["abp_mean"] = np.mean(abp_smooth)
            window_result["abp_std"] = np.std(abp_smooth)

            r_peaks = pat_result["r_peaks"]
        else:
            r_peaks = None
            window_result["pat_mean"] = np.nan

        # PTT calculation (if PPG available)
        if batch_type in ("basic", "cvp") and len(cols) >= 3:
            ppg = segment_df[cols[2]].values[start:end]
            ppg = np.nan_to_num(ppg, nan=0.0)
            if ppg.std() > 0.001:
                ptt_result = calculate_ptt(abp_win, ppg, SAMPLE_RATE, r_peaks)
                if ptt_result is not None:
                    window_result["ptt_mean"] = ptt_result["ptt_mean"]
                    window_result["ptt_std"] = ptt_result["ptt_std"]
                    window_result["ptt_median"] = ptt_result["ptt_median"]
                    window_result["ptt_n_beats"] = len(ptt_result["ptt_values"])

        # CVP analysis (if available)
        if batch_type == "cvp" and len(cols) >= 4:
            cvp = segment_df[cols[3]].values[start:end]
            cvp = np.nan_to_num(cvp, nan=0.0)
            if cvp.std() > 0.1:
                window_result["cvp_mean"] = np.mean(cvp)
                window_result["cvp_std"] = np.std(cvp)

        # 2-site PWV (if femoral available)
        if batch_type == "femoral" and len(cols) >= 3:
            fem = segment_df[cols[2]].values[start:end]
            fem = np.nan_to_num(fem, nan=0.0)
            if fem.std() > 1:
                pwv_result = calculate_2site_pwv(abp_win, fem, SAMPLE_RATE, r_peaks)
                if pwv_result is not None:
                    window_result["pwv_2site_mean"] = pwv_result["pwv_mean"]
                    window_result["pwv_2site_std"] = pwv_result["pwv_std"]
                    window_result["pwv_2site_median"] = pwv_result["pwv_median"]
                    window_result["pwv_2site_n"] = len(pwv_result["pwv_values"])

        results.append(window_result)

    return results if len(results) > 0 else None


# ============================================================
# Main Processing
# ============================================================

def process_all_cases():
    """全ケースの波形解析実行"""
    summary_path = os.path.join(DATA_DIR, "acquisition_summary.csv")
    if not os.path.exists(summary_path):
        print("Run 01_data_acquisition.py first!")
        return

    summary = pd.read_csv(summary_path)
    print(f"Processing {len(summary)} cases...")

    all_results = []

    for idx, row in summary.iterrows():
        caseid = row["caseid"]
        batch = row["batch"]

        if (idx + 1) % 20 == 0:
            print(f"  Processing {idx+1}/{len(summary)}...")

        wave_path = os.path.join(DATA_DIR, f"wave_{batch}_{caseid}.parquet")
        if not os.path.exists(wave_path):
            continue

        try:
            segment_df = pd.read_parquet(wave_path)
            results = analyze_segment(segment_df, batch)
            if results is not None:
                for r in results:
                    r["caseid"] = caseid
                    r["batch"] = batch
                all_results.extend(results)
        except Exception as e:
            print(f"  Error processing case {caseid}: {e}")
            continue

    results_df = pd.DataFrame(all_results)
    outpath = os.path.join(OUTPUT_DIR, "pwv_results.csv")
    results_df.to_csv(outpath, index=False)
    print(f"\nPWV results saved to {outpath}")
    print(f"Total windows analyzed: {len(results_df)}")
    print(f"Cases with PAT: {results_df['pat_mean'].notna().sum()}")
    if "ptt_mean" in results_df.columns:
        print(f"Cases with PTT: {results_df['ptt_mean'].notna().sum()}")
    if "pwv_2site_mean" in results_df.columns:
        print(f"Cases with 2-site PWV: {results_df['pwv_2site_mean'].notna().sum()}")

    return results_df


if __name__ == "__main__":
    process_all_cases()

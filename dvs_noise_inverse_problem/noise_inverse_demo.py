#!/usr/bin/env python3
"""
DVS Noise Inverse Problem — Demonstration on EBSSA Space Observation Data

This script demonstrates the core concept of the noise inverse problem approach:
1. Load real DVS astronomical observation data (EBSSA dataset)
2. Estimate a pixel-level noise model (simplified A5-inspired forward model)
3. Solve the inverse problem: estimate noise parameters from observed events
4. Subtract estimated noise to produce a "cleaned" residual event stream
5. Visualize the S/N improvement

Uses: EBSSA Dataset (Afshar et al. 2019) — event camera recordings of
      satellites, stars, and space objects.
"""

import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.colors import LogNorm
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# Try CJK font
try:
    import matplotlib.font_manager as fm
    cjk = [f.name for f in fm.fontManager.ttflist if 'CJK' in f.name or 'Noto Sans CJK' in f.name]
    if cjk:
        plt.rcParams['font.family'] = cjk[0]
except Exception:
    pass

OUT_DIR = Path(__file__).parent
DATA_DIR = OUT_DIR / 'data'


def load_ebssa_recording(idx=0):
    """Load a single EBSSA recording via tonic."""
    import tonic
    dataset = tonic.datasets.EBSSA(save_to=str(DATA_DIR), split='labelled')
    events, target = dataset[idx]
    return events, target, dataset


def events_to_frame(events, shape=(180, 240), time_window=None):
    """Accumulate events into a 2D event count frame."""
    if time_window is not None:
        t_min, t_max = time_window
        mask = (events['t'] >= t_min) & (events['t'] < t_max)
        events = events[mask]
    frame = np.zeros(shape, dtype=np.float64)
    np.add.at(frame, (events['y'], events['x']), 1)
    return frame


def estimate_noise_model(events, shape=(180, 240), time_bins=20):
    """
    Estimate per-pixel noise rate using temporal statistics.

    Noise model (simplified A5-inspired):
      λ_noise(x,y) = baseline dark current rate + shot noise contribution

    In the absence of auxiliary channels, we estimate noise from temporal
    statistics: pixels with constant event rate (low variance/mean ratio
    across time bins) are likely noise-dominated.
    """
    t_min, t_max = events['t'].min(), events['t'].max()
    duration_s = (t_max - t_min) / 1e6  # microseconds to seconds
    bin_edges = np.linspace(t_min, t_max, time_bins + 1)

    # Count events per pixel per time bin
    rate_cube = np.zeros((time_bins, *shape), dtype=np.float64)
    for i in range(time_bins):
        mask = (events['t'] >= bin_edges[i]) & (events['t'] < bin_edges[i + 1])
        bin_events = events[mask]
        bin_duration = (bin_edges[i + 1] - bin_edges[i]) / 1e6
        frame = np.zeros(shape, dtype=np.float64)
        np.add.at(frame, (bin_events['y'], bin_events['x']), 1)
        rate_cube[i] = frame / bin_duration  # events/second

    # Per-pixel statistics
    mean_rate = rate_cube.mean(axis=0)
    std_rate = rate_cube.std(axis=0)

    # Fano factor: variance/mean ratio. For Poisson noise, Fano ≈ 1
    # Pixels with signal will have higher Fano (bursty)
    fano = np.zeros_like(mean_rate)
    nonzero = mean_rate > 0
    fano[nonzero] = (std_rate[nonzero] ** 2) / mean_rate[nonzero]

    # Noise rate estimation:
    # Method: use median rate as noise baseline (robust to signal outliers)
    # For pixels with low Fano factor (< 2), rate ≈ noise rate
    # For pixels with high Fano factor, use median across time as noise estimate
    noise_rate = np.copy(mean_rate)

    # Pixels with high Fano likely have signal — use minimum rate as noise estimate
    high_fano = fano > 2.0
    for y in range(shape[0]):
        for x in range(shape[1]):
            if high_fano[y, x] and mean_rate[y, x] > 0:
                # Use the minimum temporal bin rate as noise floor
                noise_rate[y, x] = rate_cube[:, y, x].min()

    return {
        'noise_rate': noise_rate,        # λ_noise(x,y) [events/sec]
        'mean_rate': mean_rate,           # λ_total(x,y)
        'fano': fano,                      # Fano factor
        'duration_s': duration_s,
        'time_bins': time_bins,
        'rate_cube': rate_cube,
    }


def compute_noise_probability(events, noise_model, shape=(180, 240)):
    """
    Compute per-event noise probability P_noise(e_i).

    P_noise(e_i) = λ_noise(x_i, y_i) / λ_total(x_i, y_i)

    This is the iDQ-inspired approach: each event gets a probability
    of being noise-originated.
    """
    noise_rate = noise_model['noise_rate']
    total_rate = noise_model['mean_rate']

    p_noise = np.ones(len(events), dtype=np.float64)
    for i in range(len(events)):
        x, y = events[i]['x'], events[i]['y']
        if 0 <= y < shape[0] and 0 <= x < shape[1]:
            if total_rate[y, x] > 0:
                p_noise[i] = min(1.0, noise_rate[y, x] / total_rate[y, x])
    return p_noise


def subtract_noise(events, p_noise, threshold=0.5):
    """
    Produce residual event stream by probabilistic thinning.

    Keep events where P_noise < threshold (likely signal).
    """
    mask = p_noise < threshold
    return events[mask], p_noise[mask]


def create_demo_figures(events, noise_model, p_noise, residual_events,
                        shape=(180, 240)):
    """Generate comprehensive visualization figures."""
    t_min = events['t'].min()
    t_max = events['t'].max()
    duration = (t_max - t_min) / 1e6

    # Use first 30% of recording for visualization
    t_window = (t_min, t_min + int(0.3 * (t_max - t_min)))

    # ===== Figure 3: Noise Inverse Problem Demo (4-panel) =====
    fig, axes = plt.subplots(2, 2, figsize=(14, 10))
    fig.suptitle('DVS ノイズ逆問題デモ — EBSSA宇宙観測データ\n'
                 '(Noise Inverse Problem Demo on EBSSA Space Data)',
                 fontsize=14, fontweight='bold')

    # Panel A: Raw event accumulation
    raw_frame = events_to_frame(events, shape, t_window)
    im0 = axes[0, 0].imshow(raw_frame, cmap='hot',
                              norm=LogNorm(vmin=1, vmax=raw_frame.max() + 1))
    axes[0, 0].set_title('(A) 生イベント蓄積\nRaw Event Accumulation')
    axes[0, 0].set_xlabel('X pixel')
    axes[0, 0].set_ylabel('Y pixel')
    fig.colorbar(im0, ax=axes[0, 0], label='Event count')

    # Panel B: Estimated noise rate map
    nr = noise_model['noise_rate']
    nr_plot = np.where(nr > 0, nr, np.nan)
    im1 = axes[0, 1].imshow(nr_plot, cmap='viridis')
    axes[0, 1].set_title('(B) 推定ノイズレートマップ λ̂_noise(x,y)\n'
                          'Estimated Noise Rate [events/sec]')
    axes[0, 1].set_xlabel('X pixel')
    axes[0, 1].set_ylabel('Y pixel')
    fig.colorbar(im1, ax=axes[0, 1], label='λ_noise [evt/s]')

    # Panel C: Noise probability distribution
    axes[1, 0].hist(p_noise, bins=100, color='steelblue', alpha=0.8,
                     edgecolor='navy', linewidth=0.5)
    axes[1, 0].axvline(0.5, color='red', linestyle='--', linewidth=2,
                        label='閾値 τ=0.5')
    axes[1, 0].set_title('(C) イベントごとのノイズ確率分布\n'
                          'P_noise(e_i) Distribution (iDQ-inspired)')
    axes[1, 0].set_xlabel('P_noise(e_i)')
    axes[1, 0].set_ylabel('Count')
    axes[1, 0].legend()
    n_signal = (p_noise < 0.5).sum()
    n_noise = (p_noise >= 0.5).sum()
    axes[1, 0].text(0.95, 0.95,
                     f'信号候補: {n_signal:,} ({100*n_signal/len(p_noise):.1f}%)\n'
                     f'ノイズ: {n_noise:,} ({100*n_noise/len(p_noise):.1f}%)',
                     transform=axes[1, 0].transAxes, ha='right', va='top',
                     fontsize=10, bbox=dict(boxstyle='round', facecolor='wheat'))

    # Panel D: Residual (cleaned) event stream
    res_frame = events_to_frame(residual_events, shape, t_window)
    im3 = axes[1, 1].imshow(res_frame, cmap='hot',
                              norm=LogNorm(vmin=1, vmax=max(res_frame.max(), 2)))
    axes[1, 1].set_title('(D) 残差イベント（ノイズ差し引き後）\n'
                          'Residual Events after Noise Subtraction')
    axes[1, 1].set_xlabel('X pixel')
    axes[1, 1].set_ylabel('Y pixel')
    fig.colorbar(im3, ax=axes[1, 1], label='Event count')

    plt.tight_layout(rect=[0, 0, 1, 0.93])
    out3 = OUT_DIR / 'fig3_noise_inverse_demo.png'
    fig.savefig(out3, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f'Saved: {out3}')

    # ===== Figure 4: S/N Improvement Analysis =====
    fig, axes = plt.subplots(1, 3, figsize=(16, 5))
    fig.suptitle('S/N改善分析 — ノイズ逆問題による構造的改善\n'
                 'S/N Improvement via Noise Inverse Problem',
                 fontsize=13, fontweight='bold')

    # Panel A: Fano factor map (signal detection)
    fano = noise_model['fano']
    fano_clip = np.clip(fano, 0, 10)
    im = axes[0].imshow(fano_clip, cmap='RdYlBu_r', vmin=0, vmax=5)
    axes[0].set_title('(A) Fano因子マップ\n(>1: 信号候補)')
    axes[0].set_xlabel('X pixel')
    axes[0].set_ylabel('Y pixel')
    fig.colorbar(im, ax=axes[0], label='Fano factor')

    # Panel B: Temporal noise rate profile
    rate_cube = noise_model['rate_cube']
    mean_rates = rate_cube.mean(axis=(1, 2))
    noise_rates = np.array([noise_model['noise_rate'].mean()] * len(mean_rates))
    t_centers = np.arange(len(mean_rates))
    axes[1].plot(t_centers, mean_rates, 'b-o', label='Total rate', markersize=4)
    axes[1].plot(t_centers, noise_rates, 'r--', label='Noise model', linewidth=2)
    axes[1].fill_between(t_centers, noise_rates, mean_rates,
                          where=mean_rates > noise_rates,
                          alpha=0.3, color='green', label='Signal excess')
    axes[1].set_title('(B) 時間変動: 全レート vs ノイズモデル')
    axes[1].set_xlabel('Time bin')
    axes[1].set_ylabel('Mean event rate [evt/s/pixel]')
    axes[1].legend(fontsize=9)

    # Panel C: S/N improvement estimation
    raw_snr_per_pixel = np.zeros(shape)
    res_snr_per_pixel = np.zeros(shape)
    raw_total = events_to_frame(events, shape)
    res_total = events_to_frame(residual_events, shape)
    noise_count = noise_model['noise_rate'] * duration

    nonzero = noise_count > 0
    raw_snr_per_pixel[nonzero] = (raw_total[nonzero] - noise_count[nonzero]) / np.sqrt(noise_count[nonzero])

    # For residual, noise is reduced by (1-α) factor
    res_noise = noise_count * (1 - (p_noise < 0.5).sum() / max(len(p_noise), 1))
    res_noise_safe = np.where(res_noise > 0, res_noise, 1)
    res_snr_per_pixel[nonzero] = res_total[nonzero] / np.sqrt(res_noise_safe[nonzero])

    # Histogram of S/N
    raw_snr_flat = raw_snr_per_pixel[nonzero].flatten()
    res_snr_flat = res_snr_per_pixel[nonzero].flatten()
    raw_snr_flat = raw_snr_flat[np.isfinite(raw_snr_flat)]
    res_snr_flat = res_snr_flat[np.isfinite(res_snr_flat)]

    bins = np.linspace(-5, 20, 80)
    axes[2].hist(raw_snr_flat, bins=bins, alpha=0.6, color='red',
                  label=f'Raw (median={np.median(raw_snr_flat):.2f})')
    axes[2].hist(res_snr_flat, bins=bins, alpha=0.6, color='green',
                  label=f'Residual (median={np.median(res_snr_flat):.2f})')
    axes[2].axvline(0, color='black', linestyle='-', linewidth=0.5)
    axes[2].set_title('(C) ピクセルレベルS/N分布')
    axes[2].set_xlabel('S/N')
    axes[2].set_ylabel('Pixel count')
    axes[2].legend(fontsize=9)

    noise_removal_pct = 100 * (1 - len(residual_events) / len(events))
    fig.text(0.5, 0.01,
             f'データ: EBSSA Recording #0 | 総イベント: {len(events):,} | '
             f'ノイズ除去率: {noise_removal_pct:.1f}% | '
             f'残差イベント: {len(residual_events):,}',
             ha='center', fontsize=10,
             bbox=dict(boxstyle='round', facecolor='lightyellow'))

    plt.tight_layout(rect=[0, 0.05, 1, 0.92])
    out4 = OUT_DIR / 'fig4_sn_improvement.png'
    fig.savefig(out4, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f'Saved: {out4}')

    return out3, out4


def main():
    print('=' * 60)
    print('DVS Noise Inverse Problem Demo')
    print('Using EBSSA Space Observation Dataset')
    print('=' * 60)

    # 1. Load data
    print('\n[1/5] Loading EBSSA recording...')
    events, target, dataset = load_ebssa_recording(idx=0)
    shape = (180, 240)  # DAVIS240C sensor
    print(f'  Events: {len(events):,}')
    print(f'  Duration: {(events["t"].max() - events["t"].min()) / 1e6:.2f} sec')
    print(f'  Sensor: {shape[1]}×{shape[0]} (DAVIS240C)')

    # 2. Estimate noise model
    print('\n[2/5] Estimating per-pixel noise model (simplified A5-inspired)...')
    noise_model = estimate_noise_model(events, shape)
    nr = noise_model['noise_rate']
    print(f'  Mean noise rate: {nr[nr > 0].mean():.2f} events/sec/pixel')
    print(f'  Active pixels (noise > 0): {(nr > 0).sum()} / {shape[0]*shape[1]}')
    fano = noise_model['fano']
    print(f'  Signal candidate pixels (Fano > 2): {(fano > 2).sum()}')

    # 3. Compute noise probability per event
    print('\n[3/5] Computing per-event noise probability P_noise(e_i)...')
    p_noise = compute_noise_probability(events, noise_model, shape)
    print(f'  Mean P_noise: {p_noise.mean():.4f}')
    print(f'  Events with P_noise < 0.5 (signal candidates): {(p_noise < 0.5).sum():,}')
    print(f'  Events with P_noise >= 0.5 (noise): {(p_noise >= 0.5).sum():,}')

    # 4. Subtract noise
    print('\n[4/5] Subtracting noise (probabilistic thinning, τ=0.5)...')
    residual_events, residual_p = subtract_noise(events, p_noise, threshold=0.5)
    removal_pct = 100 * (1 - len(residual_events) / len(events))
    print(f'  Original events: {len(events):,}')
    print(f'  Residual events: {len(residual_events):,}')
    print(f'  Noise removal: {removal_pct:.1f}%')

    # 5. Generate figures
    print('\n[5/5] Generating visualization figures...')
    fig3, fig4 = create_demo_figures(events, noise_model, p_noise,
                                      residual_events, shape)

    print('\n' + '=' * 60)
    print('Demo complete!')
    print(f'  Fig 3: {fig3}')
    print(f'  Fig 4: {fig4}')
    print('=' * 60)


if __name__ == '__main__':
    main()

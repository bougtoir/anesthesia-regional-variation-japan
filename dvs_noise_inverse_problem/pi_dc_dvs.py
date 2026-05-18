#!/usr/bin/env python3
"""
PI-DC-DVS: Physics-Informed DeepClean for DVS — Simplified Implementation

This implements a simplified version of the PI-DC-DVS algorithm for evaluation
on the EBSSA and DVSNOISE20 datasets. The architecture follows the paper proposal:

  Layer 1: Physics model layer (A5-inspired parametric noise rate prediction)
  Layer 2: Spatio-temporal correlation layer (Conv2D on event context)
  Output:  Per-pixel noise rate λ_noise(x,y,t) → per-event P_noise(e_i)

Since we lack auxiliary channels in these datasets, Layer 2 (aux coupling) is
replaced by a learned temporal modulation factor.

Training: Self-supervised on dark-like regions (pixels with Fano ≈ 1).
Evaluation: Signal recovery metrics on labelled data.
"""

import numpy as np
import torch
import torch.nn as nn
import torch.optim as optim
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')


class A5PhysicsLayer(nn.Module):
    """
    Simplified A5-inspired physics model layer.

    Models per-pixel noise rate as:
      λ_noise(x,y) = I_dark(x,y) * exp(α_T * ΔT) * (1 + β * I_bg(x,y))

    Parameters learned per-pixel:
      - log_I_dark: log dark current rate (learnable)
      - threshold_mismatch: pixel threshold variation σ_θ

    For EBSSA (no aux channels), ΔT=0 and I_bg is estimated from event rates.
    """

    def __init__(self, height, width):
        super().__init__()
        self.height = height
        self.width = width
        # Per-pixel dark current (log scale for positivity)
        self.log_dark_rate = nn.Parameter(torch.zeros(height, width))
        # Temperature coefficient (global)
        self.alpha_T = nn.Parameter(torch.tensor(0.0))
        # Background light sensitivity (global)
        self.beta_bg = nn.Parameter(torch.tensor(0.1))
        # Threshold mismatch factor per pixel
        self.log_threshold_factor = nn.Parameter(torch.zeros(height, width))

    def forward(self, I_bg=None, delta_T=None):
        """
        Compute physics-based noise rate map.

        Args:
            I_bg: background illuminance map (H, W) or None
            delta_T: temperature offset (scalar) or None

        Returns:
            lambda_noise: (H, W) predicted noise rate
        """
        base_rate = torch.exp(self.log_dark_rate)
        # Temperature modulation
        if delta_T is not None:
            temp_factor = torch.exp(self.alpha_T * delta_T)
        else:
            temp_factor = 1.0
        # Background light contribution
        if I_bg is not None:
            bg_factor = 1.0 + self.beta_bg * I_bg
        else:
            bg_factor = 1.0
        # Threshold mismatch
        thresh_factor = torch.exp(self.log_threshold_factor)
        lambda_noise = base_rate * temp_factor * bg_factor * thresh_factor
        return lambda_noise


class SpatioTemporalLayer(nn.Module):
    """
    Spatio-temporal correlation layer.

    Learns inter-pixel noise correlations from local event context.
    Uses Conv2D on event rate maps to capture spatial patterns.
    """

    def __init__(self):
        super().__init__()
        self.conv = nn.Sequential(
            nn.Conv2d(1, 16, kernel_size=3, padding=1),
            nn.ReLU(),
            nn.Conv2d(16, 16, kernel_size=3, padding=1),
            nn.ReLU(),
            nn.Conv2d(16, 1, kernel_size=1),
        )

    def forward(self, event_context):
        """
        Args:
            event_context: (1, 1, H, W) event rate map

        Returns:
            delta_lambda: (H, W) correction term
        """
        out = self.conv(event_context)
        return out.squeeze(0).squeeze(0)


class TemporalModulationLayer(nn.Module):
    """
    Temporal modulation layer (substitute for aux-channel coupling).

    Without real auxiliary channels, learns a time-dependent modulation
    from the observed event rate statistics.
    """

    def __init__(self, n_features=4):
        super().__init__()
        self.mlp = nn.Sequential(
            nn.Linear(n_features, 32),
            nn.ReLU(),
            nn.Linear(32, 16),
            nn.ReLU(),
            nn.Linear(16, 1),
            nn.Tanh(),  # output in [-1, 1] for multiplicative modulation
        )

    def forward(self, temporal_features):
        """
        Args:
            temporal_features: (n_features,) — e.g., [mean_rate, std_rate, time_fraction, rate_trend]

        Returns:
            modulation: scalar multiplicative factor (1 + delta)
        """
        delta = self.mlp(temporal_features)
        return 1.0 + 0.5 * delta  # range [0.5, 1.5]


class PIDCDVS(nn.Module):
    """
    Physics-Informed DeepClean for DVS (PI-DC-DVS).

    Combines:
      1. A5 physics model layer (per-pixel parametric noise rate)
      2. Temporal modulation layer (non-stationary correction)
      3. Spatio-temporal correlation layer (inter-pixel patterns)

    Output: λ_noise(x,y,t) = λ_physics(x,y) * temporal_mod(t) + Δλ_corr(x,y,t)
    """

    def __init__(self, height, width):
        super().__init__()
        self.physics_layer = A5PhysicsLayer(height, width)
        self.temporal_layer = TemporalModulationLayer(n_features=4)
        self.spatial_layer = SpatioTemporalLayer()

    def forward(self, event_context, temporal_features, I_bg=None, delta_T=None):
        """
        Args:
            event_context: (1, 1, H, W) local event rate map
            temporal_features: (4,) temporal statistics
            I_bg: (H, W) background illuminance or None
            delta_T: scalar temperature offset or None

        Returns:
            lambda_noise: (H, W) predicted noise rate map
        """
        # Physics-based baseline
        lambda_physics = self.physics_layer(I_bg, delta_T)
        # Temporal modulation
        temp_mod = self.temporal_layer(temporal_features)
        # Spatio-temporal correction
        delta_corr = self.spatial_layer(event_context)
        # Combine: physics * modulation + correction
        lambda_noise = lambda_physics * temp_mod + delta_corr
        # Ensure non-negative
        lambda_noise = torch.relu(lambda_noise)
        return lambda_noise


def prepare_training_data(events, shape=(180, 240), n_time_bins=50):
    """
    Prepare training data from EBSSA recording.

    Strategy: Use temporal bins as training samples.
    Ground truth: Pixels with Fano ≈ 1 are noise-dominated → their rate = noise rate.

    Returns:
        time_bins: list of (rate_map, temporal_features) tuples
        noise_mask: (H, W) boolean mask of noise-dominated pixels
        fano_map: (H, W) Fano factor
    """
    t_min, t_max = events['t'].min(), events['t'].max()
    bin_edges = np.linspace(t_min, t_max, n_time_bins + 1)
    duration_total = (t_max - t_min) / 1e6

    # Compute rate maps per time bin
    rate_maps = []
    for i in range(n_time_bins):
        mask = (events['t'] >= bin_edges[i]) & (events['t'] < bin_edges[i + 1])
        bin_events = events[mask]
        bin_duration = (bin_edges[i + 1] - bin_edges[i]) / 1e6
        frame = np.zeros(shape, dtype=np.float64)
        np.add.at(frame, (bin_events['y'], bin_events['x']), 1)
        rate_map = frame / max(bin_duration, 1e-6)
        rate_maps.append(rate_map)

    rate_cube = np.array(rate_maps)

    # Fano factor for noise/signal discrimination
    mean_rate = rate_cube.mean(axis=0)
    var_rate = rate_cube.var(axis=0)
    fano_map = np.zeros_like(mean_rate)
    active = mean_rate > 0
    fano_map[active] = var_rate[active] / mean_rate[active]

    # Noise-dominated pixels: Fano < 2 and rate > 0
    noise_mask = (fano_map < 2.0) & (fano_map > 0) & (mean_rate > 0)

    # Prepare training samples (each time bin is a sample)
    training_data = []
    for i in range(n_time_bins):
        rm = rate_maps[i]
        # Temporal features: [global_mean_rate, global_std, time_fraction, rate_trend]
        global_mean = rm[noise_mask].mean() if noise_mask.any() else 0
        global_std = rm[noise_mask].std() if noise_mask.any() else 0
        time_frac = i / n_time_bins
        # Rate trend: difference from previous bin
        if i > 0:
            prev_mean = rate_maps[i-1][noise_mask].mean() if noise_mask.any() else 0
            trend = (global_mean - prev_mean) / max(global_mean, 1e-6)
        else:
            trend = 0.0
        temporal_features = np.array([global_mean, global_std, time_frac, trend],
                                     dtype=np.float32)
        training_data.append((rm.astype(np.float32), temporal_features))

    return training_data, noise_mask, fano_map, mean_rate


def train_model(model, training_data, noise_mask, n_epochs=100, lr=0.01,
                physics_weight=0.1, temporal_weight=0.01):
    """
    Train PI-DC-DVS model.

    Loss = Poisson NLL on noise pixels + physics regularisation + temporal smoothness.
    """
    optimizer = optim.Adam(model.parameters(), lr=lr)
    noise_mask_t = torch.tensor(noise_mask, dtype=torch.bool)
    losses = []

    for epoch in range(n_epochs):
        epoch_loss = 0.0
        for rate_map_np, temp_feat_np in training_data:
            optimizer.zero_grad()

            # Prepare inputs
            rate_map = torch.tensor(rate_map_np, dtype=torch.float32)
            event_context = rate_map.unsqueeze(0).unsqueeze(0)  # (1,1,H,W)
            temporal_features = torch.tensor(temp_feat_np, dtype=torch.float32)

            # Forward pass
            lambda_pred = model(event_context, temporal_features)

            # Poisson NLL loss on noise-dominated pixels only
            # L = λ_pred - rate_obs * log(λ_pred) (Poisson NLL)
            lambda_noise_pixels = lambda_pred[noise_mask_t]
            rate_obs_pixels = rate_map[noise_mask_t]

            # Avoid log(0)
            eps = 1e-8
            lambda_clamped = torch.clamp(lambda_noise_pixels, min=eps)
            poisson_loss = (lambda_clamped - rate_obs_pixels * torch.log(lambda_clamped)).mean()

            # Physics regularisation: penalise large deviations from physics baseline
            lambda_physics = model.physics_layer()
            physics_reg = ((lambda_pred - lambda_physics) ** 2).mean()

            # Temporal smoothness (small changes between time steps)
            temporal_reg = (model.temporal_layer(temporal_features) - 1.0) ** 2

            loss = poisson_loss + physics_weight * physics_reg + temporal_weight * temporal_reg
            loss.backward()
            optimizer.step()
            epoch_loss += loss.item()

        avg_loss = epoch_loss / len(training_data)
        losses.append(avg_loss)
        if (epoch + 1) % 20 == 0:
            print(f"  Epoch {epoch+1}/{n_epochs}, Loss: {avg_loss:.4f}")

    return losses


def predict_noise_probability(model, events, shape=(180, 240), n_time_bins=50,
                              noise_mask=None):
    """
    Compute per-event noise probability using trained PI-DC-DVS model.

    P_noise(e_i) = λ_noise(x_i, y_i, t_i) / λ_total(x_i, y_i, t_i)

    Args:
        noise_mask: (H, W) boolean mask of noise-dominated pixels (from training).
                    Must match the mask used during training to ensure consistent
                    temporal feature computation.
    """
    model.eval()
    t_min, t_max = events['t'].min(), events['t'].max()
    bin_edges = np.linspace(t_min, t_max, n_time_bins + 1)

    # Get predicted noise rate for each time bin
    noise_rate_maps = []
    prev_mean = None
    with torch.no_grad():
        for i in range(n_time_bins):
            mask = (events['t'] >= bin_edges[i]) & (events['t'] < bin_edges[i + 1])
            bin_events = events[mask]
            bin_duration = (bin_edges[i + 1] - bin_edges[i]) / 1e6
            frame = np.zeros(shape, dtype=np.float32)
            np.add.at(frame, (bin_events['y'], bin_events['x']), 1)
            rate_map = frame / max(bin_duration, 1e-6)

            event_context = torch.tensor(rate_map).unsqueeze(0).unsqueeze(0)
            # Match training: compute stats over noise_mask pixels only
            if noise_mask is not None and noise_mask.any():
                global_mean = rate_map[noise_mask].mean()
                global_std = rate_map[noise_mask].std()
            else:
                global_mean = rate_map.mean()
                global_std = rate_map.std()
            time_frac = i / n_time_bins
            if prev_mean is not None:
                trend = (global_mean - prev_mean) / max(global_mean, 1e-6)
            else:
                trend = 0.0
            prev_mean = global_mean
            temporal_features = torch.tensor(
                [global_mean, global_std, time_frac, trend], dtype=torch.float32)

            lambda_pred = model(event_context, temporal_features)
            noise_rate_maps.append(lambda_pred.numpy())

    # Assign noise probability to each event
    p_noise = np.ones(len(events), dtype=np.float32)
    for i in range(n_time_bins):
        mask = (events['t'] >= bin_edges[i]) & (events['t'] < bin_edges[i + 1])
        indices = np.where(mask)[0]
        if len(indices) == 0:
            continue
        bin_events = events[mask]
        lambda_noise = noise_rate_maps[i]

        # Total rate from observed data
        bin_duration = (bin_edges[i + 1] - bin_edges[i]) / 1e6
        frame = np.zeros(shape, dtype=np.float32)
        np.add.at(frame, (bin_events['y'], bin_events['x']), 1)
        lambda_total = frame / max(bin_duration, 1e-6)

        for j, idx in enumerate(indices):
            x, y = events[idx]['x'], events[idx]['y']
            if 0 <= y < shape[0] and 0 <= x < shape[1]:
                total = lambda_total[y, x]
                if total > 0:
                    p_noise[idx] = min(1.0, lambda_noise[y, x] / total)

    return p_noise


def evaluate_noise_removal(events, p_noise, threshold=0.5):
    """Compute noise removal statistics."""
    signal_mask = p_noise < threshold
    n_total = len(events)
    n_signal = signal_mask.sum()
    n_noise = n_total - n_signal
    removal_rate = n_noise / n_total
    return {
        'n_total': n_total,
        'n_signal': int(n_signal),
        'n_noise': int(n_noise),
        'removal_rate': removal_rate,
    }


def run_on_recording(events, shape=(180, 240), n_epochs=80, verbose=True):
    """
    Run full PI-DC-DVS pipeline on a single recording.

    Returns:
        results: dict with noise removal stats, model, p_noise
    """
    if verbose:
        print(f"  Events: {len(events):,}")

    # Prepare training data
    training_data, noise_mask, fano_map, mean_rate = prepare_training_data(
        events, shape=shape, n_time_bins=50)

    if not noise_mask.any():
        if verbose:
            print("  WARNING: No noise-dominated pixels found, using fallback")
        # Fallback: use all pixels
        noise_mask = mean_rate > 0

    if verbose:
        print(f"  Noise-dominated pixels: {noise_mask.sum():,} / {shape[0]*shape[1]:,}")

    # Initialise model
    model = PIDCDVS(shape[0], shape[1])

    # Initialise physics layer from observed noise rates
    with torch.no_grad():
        init_rate = mean_rate.copy()
        init_rate[init_rate <= 0] = 1e-3
        model.physics_layer.log_dark_rate.copy_(
            torch.tensor(np.log(init_rate), dtype=torch.float32))

    # Train
    if verbose:
        print("  Training PI-DC-DVS...")
    losses = train_model(model, training_data, noise_mask, n_epochs=n_epochs, lr=0.005)

    # Predict
    if verbose:
        print("  Computing noise probabilities...")
    p_noise = predict_noise_probability(model, events, shape=shape,
                                        noise_mask=noise_mask)

    # Evaluate
    stats = evaluate_noise_removal(events, p_noise, threshold=0.5)
    if verbose:
        print(f"  Results: {stats['removal_rate']*100:.1f}% noise removal "
              f"({stats['n_noise']:,} noise / {stats['n_signal']:,} signal)")

    return {
        'stats': stats,
        'p_noise': p_noise,
        'fano_map': fano_map,
        'noise_mask': noise_mask,
        'mean_rate': mean_rate,
        'losses': losses,
        'model': model,
    }


# ============================================================
# Baseline methods for comparison
# ============================================================

def baseline_temporal_filter(events, shape=(180, 240), dt_us=50000):
    """
    Baseline: Temporal neighbourhood filter (Delbruck 2008 [6]).

    An event is noise if no other event occurred at a neighbouring pixel
    within dt_us microseconds.
    """
    # Sort by time
    sorted_idx = np.argsort(events['t'])
    sorted_events = events[sorted_idx]
    n = len(sorted_events)

    is_signal = np.zeros(n, dtype=bool)

    # For efficiency, use spatial hashing
    last_event_time = np.full(shape, -np.inf)

    for i in range(n):
        x, y, t = sorted_events[i]['x'], sorted_events[i]['y'], sorted_events[i]['t']
        if not (0 <= y < shape[0] and 0 <= x < shape[1]):
            continue
        # Check 3x3 neighbourhood
        found_support = False
        for dy in range(-1, 2):
            for dx in range(-1, 2):
                if dy == 0 and dx == 0:
                    continue
                ny, nx = y + dy, x + dx
                if 0 <= ny < shape[0] and 0 <= nx < shape[1]:
                    if t - last_event_time[ny, nx] < dt_us:
                        found_support = True
                        break
            if found_support:
                break
        is_signal[i] = found_support
        last_event_time[y, x] = t

    # Map back to original order
    p_noise_sorted = (~is_signal).astype(np.float32)
    p_noise = np.ones(n, dtype=np.float32)
    p_noise[sorted_idx] = p_noise_sorted
    return p_noise


def baseline_fano_filter(events, shape=(180, 240), time_bins=20, threshold=0.5):
    """
    Baseline: Fano-factor-based filter (our previous demo).

    Pixels with Fano < 2 are noise-dominated; events from those pixels
    get P_noise = λ_noise/λ_total.
    """
    t_min, t_max = events['t'].min(), events['t'].max()
    bin_edges = np.linspace(t_min, t_max, time_bins + 1)

    rate_cube = np.zeros((time_bins, *shape), dtype=np.float64)
    for i in range(time_bins):
        mask = (events['t'] >= bin_edges[i]) & (events['t'] < bin_edges[i + 1])
        bin_events = events[mask]
        bin_duration = (bin_edges[i + 1] - bin_edges[i]) / 1e6
        frame = np.zeros(shape, dtype=np.float64)
        np.add.at(frame, (bin_events['y'], bin_events['x']), 1)
        rate_cube[i] = frame / max(bin_duration, 1e-6)

    mean_rate = rate_cube.mean(axis=0)
    var_rate = rate_cube.var(axis=0)
    fano = np.zeros_like(mean_rate)
    active = mean_rate > 0
    fano[active] = var_rate[active] / mean_rate[active]

    # Noise rate: for high-Fano pixels, use minimum temporal rate
    noise_rate = mean_rate.copy()
    high_fano = fano > 2.0
    for y in range(shape[0]):
        for x in range(shape[1]):
            if high_fano[y, x] and mean_rate[y, x] > 0:
                noise_rate[y, x] = rate_cube[:, y, x].min()

    # Per-event noise probability
    p_noise = np.ones(len(events), dtype=np.float32)
    for i in range(len(events)):
        x, y = events[i]['x'], events[i]['y']
        if 0 <= y < shape[0] and 0 <= x < shape[1]:
            if mean_rate[y, x] > 0:
                p_noise[i] = min(1.0, noise_rate[y, x] / mean_rate[y, x])
    return p_noise


if __name__ == '__main__':
    import tonic
    print("=" * 60)
    print("PI-DC-DVS: Physics-Informed DeepClean for DVS")
    print("Simplified implementation — Evaluation on EBSSA")
    print("=" * 60)

    DATA_DIR = Path(__file__).parent / 'data'
    dataset = tonic.datasets.EBSSA(save_to=str(DATA_DIR), split='labelled')

    print(f"\nDataset: EBSSA ({len(dataset)} recordings)")
    print("\nRunning on first recording...")

    events, target = dataset[0]
    results = run_on_recording(events, verbose=True)
    print("\nDone.")

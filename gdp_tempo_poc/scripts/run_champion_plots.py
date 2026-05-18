"""
Time-series comparison plots for 'champion' countries where M2 beats M0/M1 most.
"""
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from run_poc import (
    COUNTRIES, PWT_PATH,
    pim_instant, pim_lagged, pim_lagged_tempo,
    test_A_levels, test_B_growth,
)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
FIG = os.path.join(ROOT, "figures")
DATA = os.path.join(ROOT, "data")


def main():
    df = pd.read_stata(PWT_PATH)
    df = df[df["country"].isin(COUNTRIES)].copy()
    rdf = pd.read_csv(os.path.join(DATA, "poc_results.csv"))
    rdf["gain_M0_M2"] = rdf["M0_B_rmse"] - rdf["M2_B_rmse"]
    champions = rdf.sort_values("gain_M0_M2", ascending=False).head(6)["country"].tolist()

    fig, axes = plt.subplots(2, 3, figsize=(15, 9), sharex=False)
    axes = axes.flatten()

    for ax, country in zip(axes, champions):
        g = df[df["country"] == country].dropna(
            subset=["rgdpna", "rnna", "emp", "csh_i", "delta", "labsh"]
        ).sort_values("year").reset_index(drop=True)
        years = g["year"].values.astype(int)
        Y = g["rgdpna"].values.astype(float)
        Kpwt = g["rnna"].values.astype(float)
        I = Y * g["csh_i"].values.astype(float)
        delta = g["delta"].values.astype(float)
        labsh = g["labsh"].values.astype(float)
        emp = g["emp"].values.astype(float)
        avh = g["avh"].values.astype(float)
        if np.isnan(avh).any():
            avh = np.where(np.isnan(avh), np.nanmean(avh), avh)
        hc = g["hc"].values.astype(float)
        if np.isnan(hc).any():
            hc = np.where(np.isnan(hc), np.nanmean(hc), hc)
        LH = emp * avh * hc
        alpha = 1 - float(np.clip(np.mean(labsh), 0.40, 0.75))

        row = rdf[rdf["country"] == country].iloc[0]
        K0 = float(Kpwt[0])
        K_M0 = pim_instant(I, delta, K0)
        K_M1 = pim_lagged(I, delta, K0, row["mu_star"])
        K_M2 = pim_lagged_tempo(I, delta, K0, row["mu0"], row["mu1"], years)

        dY = np.diff(np.log(Y)) * 100
        dLH = np.diff(np.log(LH)) * 100

        def pred(K):
            dK = np.diff(np.log(np.where(K > 0, K, 1e-6))) * 100
            p = alpha * dK + (1 - alpha) * dLH
            g_i = np.mean(dY - p)
            return g_i + p

        ax.plot(years[1:], dY, color="black", lw=1.5, label="observed dlogY")
        ax.plot(years[1:], pred(K_M0), color="#888", lw=1, label=f"M0 (RMSE {row['M0_B_rmse']:.2f})")
        ax.plot(years[1:], pred(K_M1), color="#4c72b0", lw=1, label=f"M1 (RMSE {row['M1_B_rmse']:.2f})")
        ax.plot(years[1:], pred(K_M2), color="#dd8452", lw=1,
                label=f"M2 mu={row['mu0']:.1f}{row['mu1']:+.2f}t (RMSE {row['M2_B_rmse']:.2f})")
        ax.set_title(f"{country}")
        ax.set_ylabel("GDP growth (%/yr)")
        ax.legend(fontsize=7, loc="best")
        ax.grid(alpha=0.3)

    fig.suptitle("Growth-rate fit for 6 countries with the largest M0 -> M2 RMSE gains", y=0.995)
    fig.tight_layout(rect=[0, 0, 1, 0.97])
    fig.savefig(os.path.join(FIG, "fig5_champions.png"), dpi=140)
    plt.close(fig)
    print("Saved champions figure. Countries:", champions, flush=True)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
簡易シミュレーション: 計画放水発電 → 隣接水系地下水汲み上げサイクル
モデル地域: 小田川–高梁川流域（岡山県倉敷市真備町）

対象: 2018年西日本豪雨レベルの洪水イベント
目的: エネルギー収支・帯水層貯留容量・洪水ピーク削減率の概算

Author: (投稿者記入欄)
"""

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import rcParams

# ── 日本語フォント設定 ──
rcParams["font.family"] = "sans-serif"
rcParams["font.sans-serif"] = [
    "IPAexGothic", "IPAPGothic", "Noto Sans CJK JP",
    "DejaVu Sans",
]
rcParams["axes.unicode_minus"] = False

# ============================================================
# 1. 流域パラメータ（小田川–高梁川）
# ============================================================
# 小田川流域
A_oda = 483.0           # 流域面積 [km2]
S_oda = 1 / 2200        # 河床勾配
L_oda = 50.0            # 河川延長 [km] (概算)

# 高梁川本流
S_taka = 1 / 900        # 河床勾配
Q_taka_design = 13400   # 計画高水流量 [m3/s] (船穂地点)

# 2018年洪水データ
rain_3day = 320.0       # 3日間流域平均雨量 [mm]
V_flood = 15.3e6        # 浸水総量 [m3] (防災科研推定)
A_inundation = 8.28     # 浸水面積 [km2]
depth_mean = V_flood / (A_inundation * 1e6)  # 平均浸水深 [m]

# ============================================================
# 2. 帯水層パラメータ（真備地区周辺の沖積層）
# ============================================================
# 真備地区は沖積低地: 砂礫層 + 粘土互層
T_aquifer = 15.0        # 帯水層厚 [m] (沖積層の砂礫層)
n_e = 0.20              # 有効間隙率 [-] (砂礫)
K = 1e-3                # 透水係数 [m/s] (砂礫層)
A_managed = 20.0        # 地下水管理対象面積 [km2]

# 管理可能な地下水位低下量
dh_options = np.array([1, 2, 3, 5, 7, 10])  # [m]

# 帯水層の利用可能貯留容量
V_storage = A_managed * 1e6 * n_e * dh_options  # [m3]

# ============================================================
# 3. エネルギー収支モデル
# ============================================================

def hydropower(Q, H, eta=0.85):
    """水力発電出力 [kW]"""
    rho = 1000.0
    g = 9.81
    return rho * g * Q * H * eta / 1000  # kW

def pump_power(Q, h_lift, eta_pump=0.80):
    """揚水ポンプ所要出力 [kW]"""
    rho = 1000.0
    g = 9.81
    return rho * g * Q * h_lift / (eta_pump * 1000)  # kW

# パラメータスイープ: 有効落差 vs 放水流量
H_range = np.arange(20, 201, 10)    # 有効落差 [m]
Q_range = np.arange(5, 101, 5)      # 放水流量 [m3/s]

# 揚水側のパラメータ
h_lift = 30.0      # 地下水揚程 [m]
Q_pump = 10.0      # 揚水流量 [m3/s]
P_pump = pump_power(Q_pump, h_lift)  # kW

# 発電出力マトリクス
P_gen = np.zeros((len(H_range), len(Q_range)))
for i, H in enumerate(H_range):
    for j, Q in enumerate(Q_range):
        P_gen[i, j] = hydropower(Q, H)

# 余剰電力比率
surplus_ratio = (P_gen - P_pump) / P_gen * 100  # [%]
surplus_ratio = np.clip(surplus_ratio, 0, 100)

# ============================================================
# 4. 洪水ピーク削減モデル（簡易タンクモデル）
# ============================================================

dt = 3600.0  # タイムステップ [s] = 1時間
T_total = 72  # 総時間 [h]
t = np.arange(0, T_total + 1)

# 降雨ハイエトグラフ（2018年型を簡略再現）
# 3日間で320mm, 2回のピーク（6日夕方, 7日未明）
rain_intensity = np.zeros(T_total + 1)
# 第1ピーク: t=12-24h
for i in range(12, 24):
    rain_intensity[i] = 8.0 + 6.0 * np.sin(np.pi * (i - 12) / 12)
# 第2ピーク: t=36-54h (メインイベント)
for i in range(36, 54):
    rain_intensity[i] = 12.0 + 15.0 * np.sin(np.pi * (i - 36) / 18)
# 小雨: 残り時間
for i in range(T_total + 1):
    if rain_intensity[i] == 0 and i < 60:
        rain_intensity[i] = 2.0

# 総雨量を320mmにスケーリング
rain_total_raw = np.sum(rain_intensity)
rain_intensity = rain_intensity * (rain_3day / rain_total_raw)

# ── 流出モデル（合理式ベースの簡易モデル）──
C_runoff = 0.65   # 流出係数（現状: 市街地+農地混合）
tc = 6.0           # 集中時間 [h]

def calc_runoff(rain, C, A_km2, tc_h):
    """簡易流出ハイドログラフ"""
    n = len(rain)
    Q_out = np.zeros(n)
    tc_steps = max(1, int(tc_h))
    for i in range(n):
        # tc時間分の降雨を積算（合理式の拡張）
        start = max(0, i - tc_steps)
        avg_rain = np.mean(rain[start:i + 1])
        Q_out[i] = C * avg_rain / 1000 / 3600 * A_km2 * 1e6  # [m3/s]
    return Q_out

# シナリオ1: 現状（対策なし）
Q_no_action = calc_runoff(rain_intensity, C_runoff, A_oda, tc)

# シナリオ2: 合流点付替え（2024年完成）
# 水位低下効果: 小田川側で約4.6m → 流出係数に直接効かないが
# バックウォーター解消により排水能力が向上（ピーク流量の実効的な削減）
Q_relocation = Q_no_action * 0.70  # 30%削減（付替えによる背水効果解消）

# シナリオ3: 地下水管理型（本構想）
# 事前に帯水層に3m分の空間を創出 → 有効間隙率0.20 × 20km2 × 3m = 12Mm3
dh_managed = 3.0  # 事前地下水位低下 [m]
V_available = A_managed * 1e6 * n_e * dh_managed  # 12,000,000 m3

# 帯水層への浸透速度（上限）
# 人工涵養：散水浸透 + 注入井の組合せ
K_infiltration = 2e-3  # 浸透速度 [m/s] (人工涵養込み，砂礫層)
A_recharge = 15.0  # 涵養エリア [km2]
Q_infiltration_max = K_infiltration * A_recharge * 1e6 / 1000  # [m3/s]

# 帯水層への累積浸透量を追跡
V_infiltrated = np.zeros(T_total + 1)
Q_gw_managed = np.zeros(T_total + 1)

for i in range(1, T_total + 1):
    # 帯水層にまだ空きがあれば浸透
    if V_infiltrated[i - 1] < V_available:
        Q_inf_actual = min(Q_infiltration_max, Q_no_action[i] * 0.25)
        # 空き容量チェック
        remaining = V_available - V_infiltrated[i - 1]
        vol_step = Q_inf_actual * dt
        if vol_step > remaining:
            Q_inf_actual = remaining / dt
        V_infiltrated[i] = V_infiltrated[i - 1] + Q_inf_actual * dt
        Q_gw_managed[i] = Q_no_action[i] - Q_inf_actual
    else:
        V_infiltrated[i] = V_infiltrated[i - 1]
        Q_gw_managed[i] = Q_no_action[i]

# シナリオ4: 構造＋地下水の統合
Q_integrated = Q_gw_managed * 0.70  # 合流点付替え + 地下水管理

# ピーク流量
Q_peak_no_action = np.max(Q_no_action)
Q_peak_relocation = np.max(Q_relocation)
Q_peak_gw = np.max(Q_gw_managed)
Q_peak_integrated = np.max(Q_integrated)

# 浸水量の概算
V_total_no_action = np.sum(Q_no_action) * dt
V_total_relocation = np.sum(Q_relocation) * dt
V_total_gw = np.sum(Q_gw_managed) * dt
V_total_integrated = np.sum(Q_integrated) * dt

# ============================================================
# 5. 結果のサマリー出力
# ============================================================

print("=" * 60)
print("小田川–高梁川流域 洪水制御シミュレーション結果")
print("=" * 60)
print(f"\n【流域パラメータ】")
print(f"  小田川流域面積: {A_oda} km2")
print(f"  3日間降雨量:    {rain_3day} mm")
print(f"  2018年浸水量:   {V_flood/1e6:.1f} × 10^6 m3")
print(f"  2018年浸水面積: {A_inundation} km2")
print(f"  平均浸水深:     {depth_mean:.2f} m")

print(f"\n【帯水層貯留容量（管理面積={A_managed}km2, 間隙率={n_e}）】")
for i, dh in enumerate(dh_options):
    ratio = V_storage[i] / V_flood * 100
    print(f"  Δh={dh:2d}m → 貯留容量 {V_storage[i]/1e6:6.1f} × 10^6 m3"
          f"  （2018年浸水量の {ratio:5.1f}%）")

print(f"\n【エネルギー収支（代表値）】")
print(f"  発電: H=100m, Q=50m3/s → P = {hydropower(50, 100):.0f} kW = {hydropower(50, 100)/1000:.1f} MW")
print(f"  揚水: h={h_lift}m, Q={Q_pump}m3/s → P = {P_pump:.0f} kW = {P_pump/1000:.1f} MW")
print(f"  余剰: {(hydropower(50, 100) - P_pump)/1000:.1f} MW ({(1 - P_pump/hydropower(50, 100))*100:.0f}%)")

print(f"\n【洪水ピーク削減効果】")
print(f"  シナリオ1 (対策なし):       Qp = {Q_peak_no_action:.0f} m3/s (基準)")
print(f"  シナリオ2 (合流点付替え):   Qp = {Q_peak_relocation:.0f} m3/s ({(1-Q_peak_relocation/Q_peak_no_action)*100:.0f}%削減)")
print(f"  シナリオ3 (地下水管理型):   Qp = {Q_peak_gw:.0f} m3/s ({(1-Q_peak_gw/Q_peak_no_action)*100:.0f}%削減)")
print(f"  シナリオ4 (統合型):         Qp = {Q_peak_integrated:.0f} m3/s ({(1-Q_peak_integrated/Q_peak_no_action)*100:.0f}%削減)")

print(f"\n【帯水層浸透量】")
print(f"  累積浸透量: {V_infiltrated[-1]/1e6:.2f} × 10^6 m3")
print(f"  2018年浸水量に対する比率: {V_infiltrated[-1]/V_flood*100:.1f}%")

# ============================================================
# 6. 図表生成
# ============================================================

outdir = "/home/ubuntu/repos/wip/flood-prediction-review/figures"
import os
os.makedirs(outdir, exist_ok=True)

# ── 図-1: エネルギー収支コンター図（日本語）──
fig1, ax1 = plt.subplots(figsize=(8, 6))
QQ_grid, HH_grid = np.meshgrid(Q_range, H_range)  # QQ_grid=Q値, HH_grid=H値
cf = ax1.contourf(QQ_grid, HH_grid, P_gen / 1000,
                  levels=np.arange(0, 180, 10), cmap="YlOrRd", extend="both")
cbar = plt.colorbar(cf, ax=ax1)
cbar.set_label("発電出力 [MW]", fontsize=12)

# 揚水分岐点
for Q_val in [10, 20, 30, 50, 80]:
    H_break = P_pump * 1000 / (1000 * 9.81 * Q_val * 0.85)
    if 20 <= H_break <= 200:
        ax1.plot(Q_val, H_break, "bD", markersize=6)

# 揚水所要出力の等高線
H_breakeven = P_pump * 1000 / (1000 * 9.81 * Q_range * 0.85)
valid = (H_breakeven >= 20) & (H_breakeven <= 200)
ax1.plot(Q_range[valid], H_breakeven[valid], "b--", linewidth=2,
         label=f"揚水所要出力 = {P_pump/1000:.1f} MW")

ax1.set_xlabel("計画放水流量 Q [m³/s]", fontsize=12)
ax1.set_ylabel("有効落差 H [m]", fontsize=12)
ax1.set_title("図-1 計画放水発電の出力と揚水所要出力の関係", fontsize=13, fontweight="bold")
ax1.legend(loc="upper left", fontsize=10)
ax1.set_xlim(5, 100)
ax1.set_ylim(20, 200)
fig1.tight_layout()
fig1.savefig(f"{outdir}/fig1_energy_balance_ja.png", dpi=300, bbox_inches="tight")
plt.close(fig1)

# ── 図-2: 帯水層貯留容量 vs 2018年浸水量（日本語）──
fig2, ax2 = plt.subplots(figsize=(8, 5))
bars = ax2.bar(dh_options, V_storage / 1e6, color="#4ECDC4", edgecolor="black", width=0.6)
ax2.axhline(y=V_flood / 1e6, color="red", linestyle="--", linewidth=2,
            label=f"2018年浸水量 ({V_flood/1e6:.1f} x 10$^6$ m$^3$)")
ax2.set_xlabel("地下水位低下量 Δh [m]", fontsize=12)
ax2.set_ylabel("帯水層貯留容量 [x 10$^6$ m$^3$]", fontsize=12)
ax2.set_title("図-2 事前地下水位低下による帯水層貯留容量", fontsize=13, fontweight="bold")
ax2.legend(fontsize=10)
ax2.set_xticks(dh_options)

# バーの上に割合を表示
for bar, v in zip(bars, V_storage):
    ratio = v / V_flood * 100
    ax2.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
             f"{ratio:.0f}%", ha="center", va="bottom", fontsize=10, fontweight="bold")

fig2.tight_layout()
fig2.savefig(f"{outdir}/fig2_aquifer_storage_ja.png", dpi=300, bbox_inches="tight")
plt.close(fig2)

# ── 図-3: 洪水ハイドログラフ比較（日本語）──
fig3, (ax3a, ax3b) = plt.subplots(2, 1, figsize=(10, 8), gridspec_kw={"height_ratios": [1, 3]})

# 上段: 降雨
ax3a.bar(t, rain_intensity, color="#5B9BD5", width=0.8, label="降雨強度")
ax3a.set_ylabel("降雨強度\n[mm/h]", fontsize=11)
ax3a.set_xlim(0, T_total)
ax3a.invert_yaxis()
ax3a.legend(fontsize=9)
ax3a.set_title("図-3 洪水シナリオ別ハイドログラフ比較（小田川流域）", fontsize=13, fontweight="bold")

# 下段: ハイドログラフ
ax3b.plot(t, Q_no_action, "k-", linewidth=2, label="シナリオ1: 対策なし")
ax3b.plot(t, Q_relocation, "b-", linewidth=2, label="シナリオ2: 合流点付替え")
ax3b.plot(t, Q_gw_managed, "g-", linewidth=2, label="シナリオ3: 地下水管理型")
ax3b.plot(t, Q_integrated, "r-", linewidth=2, label="シナリオ4: 統合型")
ax3b.set_xlabel("経過時間 [h]", fontsize=12)
ax3b.set_ylabel("流出量 [m³/s]", fontsize=12)
ax3b.legend(fontsize=10, loc="upper right")
ax3b.set_xlim(0, T_total)
ax3b.set_ylim(bottom=0)

fig3.tight_layout()
fig3.savefig(f"{outdir}/fig3_hydrograph_ja.png", dpi=300, bbox_inches="tight")
plt.close(fig3)

# ── 図-4: シナリオ別ピーク流量・浸水量の棒グラフ（日本語）──
fig4, (ax4a, ax4b) = plt.subplots(1, 2, figsize=(12, 5))

scenarios = ["対策なし", "合流点\n付替え", "地下水\n管理型", "統合型"]
peaks = [Q_peak_no_action, Q_peak_relocation, Q_peak_gw, Q_peak_integrated]
colors = ["#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4"]

ax4a.bar(scenarios, peaks, color=colors, edgecolor="black")
ax4a.set_ylabel("ピーク流量 [m³/s]", fontsize=12)
ax4a.set_title("(a) ピーク流量の比較", fontsize=12, fontweight="bold")
for i, (s, p) in enumerate(zip(scenarios, peaks)):
    reduction = (1 - p / Q_peak_no_action) * 100
    ax4a.text(i, p + 10, f"-{reduction:.0f}%" if reduction > 0 else "基準",
              ha="center", fontsize=11, fontweight="bold")

volumes = [V_total_no_action / 1e6, V_total_relocation / 1e6,
           V_total_gw / 1e6, V_total_integrated / 1e6]
ax4b.bar(scenarios, volumes, color=colors, edgecolor="black")
ax4b.set_ylabel("総流出量 [x 10$^6$ m$^3$]", fontsize=12)
ax4b.set_title("(b) 総流出量の比較", fontsize=12, fontweight="bold")
for i, (s, v) in enumerate(zip(scenarios, volumes)):
    reduction = (1 - v / (V_total_no_action / 1e6)) * 100
    ax4b.text(i, v + 0.5, f"-{reduction:.0f}%" if reduction > 0 else "基準",
              ha="center", fontsize=11, fontweight="bold")

fig4.suptitle("図-4 シナリオ別の洪水制御効果", fontsize=14, fontweight="bold", y=1.02)
fig4.tight_layout()
fig4.savefig(f"{outdir}/fig4_scenario_comparison_ja.png", dpi=300, bbox_inches="tight")
plt.close(fig4)

# ============================================================
# 7. 英語版図表
# ============================================================

# ── Fig. 1 (EN): Energy balance contour ──
fig1e, ax1e = plt.subplots(figsize=(8, 6))
cf = ax1e.contourf(QQ_grid, HH_grid, P_gen / 1000,
                   levels=np.arange(0, 180, 10), cmap="YlOrRd", extend="both")
cbar = plt.colorbar(cf, ax=ax1e)
cbar.set_label("Power Output [MW]", fontsize=12)

for Q_val in [10, 20, 30, 50, 80]:
    H_break = P_pump * 1000 / (1000 * 9.81 * Q_val * 0.85)
    if 20 <= H_break <= 200:
        ax1e.plot(Q_val, H_break, "bD", markersize=6)

ax1e.plot(Q_range[valid], H_breakeven[valid], "b--", linewidth=2,
          label=f"Pumping Power = {P_pump/1000:.1f} MW")
ax1e.set_xlabel("Planned Discharge Q [m\u00b3/s]", fontsize=12)
ax1e.set_ylabel("Effective Head H [m]", fontsize=12)
ax1e.set_title("Fig. 1  Hydropower Output vs Pumping Requirement", fontsize=13, fontweight="bold")
ax1e.legend(loc="upper left", fontsize=10)
ax1e.set_xlim(5, 100)
ax1e.set_ylim(20, 200)
fig1e.tight_layout()
fig1e.savefig(f"{outdir}/fig1_energy_balance_en.png", dpi=300, bbox_inches="tight")
plt.close(fig1e)

# ── Fig. 2 (EN): Aquifer storage ──
fig2e, ax2e = plt.subplots(figsize=(8, 5))
bars_e = ax2e.bar(dh_options, V_storage / 1e6, color="#4ECDC4", edgecolor="black", width=0.6)
ax2e.axhline(y=V_flood / 1e6, color="red", linestyle="--", linewidth=2,
             label=f"2018 Flood Volume ({V_flood/1e6:.1f} x 10$^6$ m$^3$)")
ax2e.set_xlabel("Groundwater Drawdown Δh [m]", fontsize=12)
ax2e.set_ylabel("Aquifer Storage Capacity [x 10$^6$ m$^3$]", fontsize=12)
ax2e.set_title("Fig. 2  Aquifer Storage from Pre-Flood Drawdown", fontsize=13, fontweight="bold")
ax2e.legend(fontsize=10)
ax2e.set_xticks(dh_options)
for bar, v in zip(bars_e, V_storage):
    ratio = v / V_flood * 100
    ax2e.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
              f"{ratio:.0f}%", ha="center", va="bottom", fontsize=10, fontweight="bold")
fig2e.tight_layout()
fig2e.savefig(f"{outdir}/fig2_aquifer_storage_en.png", dpi=300, bbox_inches="tight")
plt.close(fig2e)

# ── Fig. 3 (EN): Hydrograph ──
fig3e, (ax3ae, ax3be) = plt.subplots(2, 1, figsize=(10, 8), gridspec_kw={"height_ratios": [1, 3]})
ax3ae.bar(t, rain_intensity, color="#5B9BD5", width=0.8, label="Rainfall Intensity")
ax3ae.set_ylabel("Rainfall\n[mm/h]", fontsize=11)
ax3ae.set_xlim(0, T_total)
ax3ae.invert_yaxis()
ax3ae.legend(fontsize=9)
ax3ae.set_title("Fig. 3  Flood Hydrograph Comparison (Oda River Basin)", fontsize=13, fontweight="bold")

ax3be.plot(t, Q_no_action, "k-", linewidth=2, label="Scenario 1: No Action")
ax3be.plot(t, Q_relocation, "b-", linewidth=2, label="Scenario 2: Confluence Relocation")
ax3be.plot(t, Q_gw_managed, "g-", linewidth=2, label="Scenario 3: GW Management")
ax3be.plot(t, Q_integrated, "r-", linewidth=2, label="Scenario 4: Integrated")
ax3be.set_xlabel("Elapsed Time [h]", fontsize=12)
ax3be.set_ylabel("Discharge [m³/s]", fontsize=12)
ax3be.legend(fontsize=10, loc="upper right")
ax3be.set_xlim(0, T_total)
ax3be.set_ylim(bottom=0)
fig3e.tight_layout()
fig3e.savefig(f"{outdir}/fig3_hydrograph_en.png", dpi=300, bbox_inches="tight")
plt.close(fig3e)

# ── Fig. 4 (EN): Scenario comparison ──
fig4e, (ax4ae, ax4be) = plt.subplots(1, 2, figsize=(12, 5))
scenarios_en = ["No Action", "Confluence\nRelocation", "GW\nManagement", "Integrated"]
ax4ae.bar(scenarios_en, peaks, color=colors, edgecolor="black")
ax4ae.set_ylabel("Peak Discharge [m³/s]", fontsize=12)
ax4ae.set_title("(a) Peak Discharge Comparison", fontsize=12, fontweight="bold")
for i, p in enumerate(peaks):
    reduction = (1 - p / Q_peak_no_action) * 100
    ax4ae.text(i, p + 10, f"-{reduction:.0f}%" if reduction > 0 else "Baseline",
               ha="center", fontsize=11, fontweight="bold")

ax4be.bar(scenarios_en, volumes, color=colors, edgecolor="black")
ax4be.set_ylabel("Total Runoff Volume [x 10$^6$ m$^3$]", fontsize=12)
ax4be.set_title("(b) Total Runoff Volume Comparison", fontsize=12, fontweight="bold")
for i, v in enumerate(volumes):
    reduction = (1 - v / (V_total_no_action / 1e6)) * 100
    ax4be.text(i, v + 0.5, f"-{reduction:.0f}%" if reduction > 0 else "Baseline",
               ha="center", fontsize=11, fontweight="bold")

fig4e.suptitle("Fig. 4  Flood Control Effectiveness by Scenario", fontsize=14, fontweight="bold", y=1.02)
fig4e.tight_layout()
fig4e.savefig(f"{outdir}/fig4_scenario_comparison_en.png", dpi=300, bbox_inches="tight")
plt.close(fig4e)

# ── 表-1: シナリオ比較サマリー（CSVで出力）──
import csv
with open(f"{outdir}/table1_summary.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["シナリオ", "ピーク流量[m3/s]", "削減率[%]",
                "総流出量[×10^6 m3]", "削減率[%]"])
    w.writerow(["1: 対策なし", f"{Q_peak_no_action:.0f}", "0",
                f"{V_total_no_action/1e6:.1f}", "0"])
    w.writerow(["2: 合流点付替え", f"{Q_peak_relocation:.0f}",
                f"{(1-Q_peak_relocation/Q_peak_no_action)*100:.0f}",
                f"{V_total_relocation/1e6:.1f}",
                f"{(1-V_total_relocation/V_total_no_action)*100:.0f}"])
    w.writerow(["3: 地下水管理型", f"{Q_peak_gw:.0f}",
                f"{(1-Q_peak_gw/Q_peak_no_action)*100:.0f}",
                f"{V_total_gw/1e6:.1f}",
                f"{(1-V_total_gw/V_total_no_action)*100:.0f}"])
    w.writerow(["4: 統合型", f"{Q_peak_integrated:.0f}",
                f"{(1-Q_peak_integrated/Q_peak_no_action)*100:.0f}",
                f"{V_total_integrated/1e6:.1f}",
                f"{(1-V_total_integrated/V_total_no_action)*100:.0f}"])

print(f"\n【出力ファイル】")
print(f"  {outdir}/fig1_energy_balance_ja.png")
print(f"  {outdir}/fig2_aquifer_storage_ja.png")
print(f"  {outdir}/fig3_hydrograph_ja.png")
print(f"  {outdir}/fig4_scenario_comparison_ja.png")
print(f"  {outdir}/fig1_energy_balance_en.png")
print(f"  {outdir}/fig2_aquifer_storage_en.png")
print(f"  {outdir}/fig3_hydrograph_en.png")
print(f"  {outdir}/fig4_scenario_comparison_en.png")
print(f"  {outdir}/table1_summary.csv")
print("\n完了")

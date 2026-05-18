import geopandas as gpd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import japanize_matplotlib
from matplotlib.colors import BoundaryNorm, ListedColormap
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
import warnings
warnings.filterwarnings('ignore')

print("Loading data...")
gdf = gpd.read_file('/home/ubuntu/gis_data/merged_enriched.gpkg')
gdf_pref = gpd.read_file('/home/ubuntu/gis_data/pref_simplified.gpkg')
gdf_nt = gpd.read_file('/home/ubuntu/gis_data/northern_territories.gpkg')

# Single-panel map bounds including Okinawa at actual position and Northern Territories
all_b = {'x': (122.5, 149.2), 'y': (24.0, 46.0)}

def make_map(gdf, gdf_pref, gdf_nt, column, title, cmap, boundaries, legend_label, filename,
             show_univ=False, categorical=False, cat_items=None, extra_legend=None):
    fig, ax = plt.subplots(1, 1, figsize=(14, 18))

    if categorical:
        for val, color, _ in cat_items:
            sub = gdf[gdf[column] == val]
            if len(sub) > 0:
                sub.plot(ax=ax, color=color, edgecolor='none')
        nodata = gdf[gdf[column].isna()]
        if len(nodata) > 0:
            nodata.plot(ax=ax, color='#e0e0e0', edgecolor='none')
    else:
        norm = BoundaryNorm(boundaries, ncolors=256)
        valid = gdf[gdf[column].notna()]
        nodata = gdf[gdf[column].isna()]
        if len(nodata) > 0:
            nodata.plot(ax=ax, color='#e0e0e0', edgecolor='none')
        if len(valid) > 0:
            valid.plot(ax=ax, column=column, cmap=cmap, norm=norm, edgecolor='none', legend=False)

    # Northern Territories: white fill with thin border
    gdf_nt.plot(ax=ax, color='white', edgecolor='#333333', linewidth=0.5)

    # Boundaries: dotted for medical areas, solid for prefectures
    gdf.boundary.plot(ax=ax, linewidth=0.3, color='#888888', linestyle=(0, (2, 3)))
    gdf_pref.boundary.plot(ax=ax, linewidth=0.8, color='#333333', linestyle='solid')

    if show_univ:
        univ = gdf[gdf['has_univ'] == 1]
        c = univ.geometry.centroid
        ax.scatter(c.x, c.y, s=18, c='red', marker='o', zorder=5,
                  linewidths=0.5, edgecolors='darkred', alpha=0.8)

    ax.set_xlim(all_b['x']); ax.set_ylim(all_b['y'])
    ax.set_aspect('equal'); ax.axis('off')

    ax.set_title(title, fontsize=15, fontweight='bold', pad=15)

    # Legend
    line_items = [
        Line2D([0],[0], color='#333333', lw=1.0, ls='solid', label='都道府県境'),
        Line2D([0],[0], color='#888888', lw=0.5, ls=(0,(2,3)), label='二次医療圏境'),
        Patch(facecolor='white', edgecolor='#333333', linewidth=0.5, label='北方領土（医療圏未設定）'),
    ]
    if show_univ:
        line_items.append(Line2D([0],[0], marker='o', color='w', markerfacecolor='red',
                                  markersize=6, label='大学病院所在圏'))

    if categorical:
        cat_legend = [Patch(facecolor=c, label=l) for _, c, l in cat_items]
        all_legend = cat_legend + line_items
    else:
        all_legend = line_items
    
    if extra_legend:
        all_legend = extra_legend + all_legend

    ax.legend(handles=all_legend, loc='lower left', fontsize=8.5, framealpha=0.95, edgecolor='gray')

    if not categorical:
        norm = BoundaryNorm(boundaries, ncolors=256)
        sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
        sm.set_array([])
        cbar_ax = fig.add_axes([0.15, 0.06, 0.55, 0.015])
        cbar = fig.colorbar(sm, cax=cbar_ax, orientation='horizontal')
        cbar.set_label(legend_label, fontsize=12)

    plt.savefig(filename, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f"Saved: {filename}")

# ========== EXISTING 6 MAPS ==========

# Figure 1: L008 SCR
print("\n--- Figure 1: L008 SCR ---")
make_map(gdf, gdf_pref, gdf_nt, 'L008_scr',
         '全身麻酔 (L008) SCR 二次医療圏別', 'RdYlBu_r',
         [0, 30, 50, 70, 90, 100, 110, 130, 160, 200, 500],
         'SCR (100 = 全国平均)',
         '/home/ubuntu/map_L008_scr.png', show_univ=True)

# Figure 2: University Hospital Presence
print("--- Figure 2: University Presence ---")
make_map(gdf, gdf_pref, gdf_nt, 'has_univ',
         '大学病院所在 二次医療圏', None, None, None,
         '/home/ubuntu/map_univ_presence.png',
         categorical=True,
         cat_items=[(1, '#e74c3c', '大学病院あり (64圏)'),
                    (0, '#d4e6f1', '大学病院なし (271圏)')])

# Figure 3: Number of university hospitals (dose-response)
print("--- Figure 3: N Univ (dose-response) ---")
dose_colors = ['#d4e6f1', '#f9e79f', '#e67e22', '#c0392b', '#7b241c']
dose_cmap = ListedColormap(dose_colors)
dose_bounds = [-0.5, 0.5, 1.5, 2.5, 3.5, 10]
dose_norm = BoundaryNorm(dose_bounds, dose_cmap.N)

fig, ax = plt.subplots(1, 1, figsize=(14, 18))
gdf.plot(ax=ax, column='n_univ', cmap=dose_cmap, norm=dose_norm, edgecolor='none')
gdf_nt.plot(ax=ax, color='white', edgecolor='#333333', linewidth=0.5)
gdf.boundary.plot(ax=ax, linewidth=0.3, color='#888888', linestyle=(0, (2, 3)))
gdf_pref.boundary.plot(ax=ax, linewidth=0.8, color='#333333', linestyle='solid')
ax.set_xlim(all_b['x']); ax.set_ylim(all_b['y'])
ax.set_aspect('equal'); ax.axis('off')
ax.set_title('大学病院数 二次医療圏別', fontsize=15, fontweight='bold', pad=15)
legend_elements = [
    Patch(facecolor=dose_colors[0], label='0校 (271圏)'),
    Patch(facecolor=dose_colors[1], label='1校 (51圏)'),
    Patch(facecolor=dose_colors[2], label='2校 (10圏)'),
    Patch(facecolor=dose_colors[3], label='3校 (2圏)'),
    Patch(facecolor=dose_colors[4], label='5校 (1圏)'),
    Line2D([0],[0], color='#333333', lw=1.0, ls='solid', label='都道府県境'),
    Line2D([0],[0], color='#888888', lw=0.5, ls=(0,(2,3)), label='二次医療圏境'),
    Patch(facecolor='white', edgecolor='#333333', linewidth=0.5, label='北方領土（医療圏未設定）'),
]
ax.legend(handles=legend_elements, loc='lower left', fontsize=8.5, framealpha=0.95, edgecolor='gray')
plt.savefig('/home/ubuntu/map_n_univ.png', dpi=200, bbox_inches='tight', facecolor='white')
plt.close()
print("Saved: map_n_univ.png")

# Figure 4: L004 Spinal Anesthesia
print("--- Figure 4: L004 SCR ---")
make_map(gdf, gdf_pref, gdf_nt, 'L004_scr',
         '脊椎麻酔 (L004) SCR 二次医療圏別', 'RdYlBu_r',
         [0, 30, 50, 70, 90, 100, 110, 130, 160, 200, 400],
         'SCR (100 = 全国平均)',
         '/home/ubuntu/map_L004_scr.png', show_univ=True)

# Figure 5: L008+L004 Combined
print("--- Figure 5: L008+L004 Combined ---")
make_map(gdf, gdf_pref, gdf_nt, 'L008_L004_combined',
         '全身麻酔+脊椎麻酔 合計SCR (L008+L004)\n査定による振替効果を中和',
         'RdYlBu_r',
         [0, 80, 120, 150, 180, 200, 230, 270, 350, 600],
         '合計SCR',
         '/home/ubuntu/map_L008_L004_combined.png', show_univ=True)

# Figure 6: L002 Epidural
print("--- Figure 6: L002 SCR ---")
make_map(gdf, gdf_pref, gdf_nt, 'L002_scr',
         '硬膜外麻酔 (L002) SCR 二次医療圏別', 'RdYlBu_r',
         [0, 30, 50, 70, 90, 100, 110, 130, 160, 200, 400],
         'SCR (100 = 全国平均)',
         '/home/ubuntu/map_L002_scr.png', show_univ=True)

# ========== NEW MAPS FOR TUMOR ANALYSIS ==========

# Figure 7: L003 Continuous Epidural (direct GA+epidural combination indicator)
print("\n--- Figure 7: L003 Continuous Epidural SCR ---")
make_map(gdf, gdf_pref, gdf_nt, 'L003_scr',
         '硬膜外麻酔後持続注入 (L003) SCR 二次医療圏別\n全身麻酔+硬膜外併用の直接指標',
         'RdYlBu_r',
         [0, 20, 40, 60, 80, 100, 120, 150, 200, 300],
         'SCR (100 = 全国平均)',
         '/home/ubuntu/map_L003_scr.png', show_univ=True)

# Figure 8: L002/L008 ratio (epidural combination tendency)
print("--- Figure 8: L002/L008 Ratio ---")
make_map(gdf, gdf_pref, gdf_nt, 'L002_L008_ratio',
         '硬膜外/全身麻酔 比率 (L002÷L008) 二次医療圏別\n区域麻酔併用傾向の指標',
         'PiYG',
         [0, 0.3, 0.5, 0.7, 0.9, 1.0, 1.2, 1.5, 2.0, 3.0, 10.0],
         'L002/L008 比率 (1.0 = 同量)',
         '/home/ubuntu/map_L002_L008_ratio.png', show_univ=True)

# Figure 9: L003/L008 ratio (direct combination rate)
print("--- Figure 9: L003/L008 Ratio ---")
make_map(gdf, gdf_pref, gdf_nt, 'L003_L008_ratio',
         '持続硬膜外/全身麻酔 比率 (L003÷L008) 二次医療圏別\n全身麻酔に対する硬膜外併用率',
         'PiYG',
         [0, 0.3, 0.5, 0.7, 0.9, 1.0, 1.2, 1.5, 2.0, 3.0, 10.0],
         'L003/L008 比率',
         '/home/ubuntu/map_L003_L008_ratio.png', show_univ=True)

# Figure 10: L008+L002 combined (GA+epidural total - audit sensitivity)
print("--- Figure 10: L008+L002 Combined ---")
make_map(gdf, gdf_pref, gdf_nt, 'L008_L002_combined',
         '全身麻酔+硬膜外麻酔 合計SCR (L008+L002)\n査定による振替効果を中和',
         'RdYlBu_r',
         [0, 60, 100, 140, 180, 200, 250, 320, 400, 700],
         '合計SCR',
         '/home/ubuntu/map_L008_L002_combined.png', show_univ=True)

print("\n=== All 10 maps generated successfully ===")

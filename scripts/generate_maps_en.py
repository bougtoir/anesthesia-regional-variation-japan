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
        Line2D([0],[0], color='#333333', lw=1.0, ls='solid', label='Prefecture boundary'),
        Line2D([0],[0], color='#888888', lw=0.5, ls=(0,(2,3)), label='SMA boundary'),
        Patch(facecolor='white', edgecolor='#333333', linewidth=0.5, label='Northern Territories (no SMA)'),
    ]
    if show_univ:
        line_items.append(Line2D([0],[0], marker='o', color='w', markerfacecolor='red',
                                  markersize=6, label='University hospital SMA'))

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

OUT = '/home/ubuntu/en_'

# Figure 1: L008 SCR
print("\n--- Figure 1: L008 SCR ---")
make_map(gdf, gdf_pref, gdf_nt, 'L008_scr',
         'General Anaesthesia (L008) SCR by Secondary Medical Area', 'RdYlBu_r',
         [0, 30, 50, 70, 90, 100, 110, 130, 160, 200, 500],
         'SCR (100 = national average)',
         OUT + 'map_L008_scr.png', show_univ=True)

# Figure 2: University Hospital Presence
print("--- Figure 2: University Presence ---")
make_map(gdf, gdf_pref, gdf_nt, 'has_univ',
         'University Hospital Presence by Secondary Medical Area', None, None, None,
         OUT + 'map_univ_presence.png',
         categorical=True,
         cat_items=[(1, '#e74c3c', 'University hospital present (64 SMAs)'),
                    (0, '#d4e6f1', 'No university hospital (271 SMAs)')])

# Figure 3: L004 Spinal Anaesthesia
print("--- Figure 3: L004 SCR ---")
make_map(gdf, gdf_pref, gdf_nt, 'L004_scr',
         'Spinal Anaesthesia (L004) SCR by Secondary Medical Area', 'RdYlBu_r',
         [0, 30, 50, 70, 90, 100, 110, 130, 160, 200, 400],
         'SCR (100 = national average)',
         OUT + 'map_L004_scr.png', show_univ=True)

# Figure 4: L008+L004 Combined
print("--- Figure 4: L008+L004 Combined ---")
make_map(gdf, gdf_pref, gdf_nt, 'L008_L004_combined',
         'GA + Spinal Anaesthesia Combined SCR (L008+L004)\nAudit substitution effect neutralised',
         'RdYlBu_r',
         [0, 80, 120, 150, 180, 200, 230, 270, 350, 600],
         'Combined SCR',
         OUT + 'map_L008_L004_combined.png', show_univ=True)

# Figure 5: L003 Continuous Epidural
print("--- Figure 5: L003 SCR ---")
make_map(gdf, gdf_pref, gdf_nt, 'L003_scr',
         'Continuous Epidural Infusion (L003) SCR by Secondary Medical Area\n'
         'Direct indicator of GA + epidural combination',
         'RdYlBu_r',
         [0, 20, 40, 60, 80, 100, 120, 150, 200, 300],
         'SCR (100 = national average)',
         OUT + 'map_L003_scr.png', show_univ=True)

# Figure 6: L003/L008 ratio (corrected)
print("--- Figure 6: L003/L008 Ratio ---")
make_map(gdf, gdf_pref, gdf_nt, 'L003_L008_ratio',
         'Epidural Combination Rate (L003/L008) by Secondary Medical Area\n'
         'Combined epidural rate per general anaesthesia',
         'PiYG',
         [0, 0.3, 0.5, 0.7, 0.9, 1.0, 1.2, 1.5, 2.0, 3.0, 10.0],
         'L003/L008 Ratio',
         OUT + 'map_L003_L008_ratio_corrected.png', show_univ=True)

print("\n=== All 6 English maps generated successfully ===")

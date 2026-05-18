#!/usr/bin/env python3
"""
Create multi-panel composite figures for RAPM submission.
Strategy: 9 individual figures -> 4 multi-panel figures (+ 2 tables = 6 total)
- Figure 1: Panel A = L008 SCR, Panel B = L004 SCR
- Figure 2: Panel A = University hospital presence, Panel B = L008+L004 combined
- Figure 3: Panel A = L003 SCR, Panel B = L003/L008 ratio corrected
- Figure 4: Single 3D extruded map (ratio by anes) - best single 3D
"""

import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import numpy as np

def create_panel_figure(panel_paths, labels, output_path, figsize=(16, 8), dpi=300):
    """Create a 1x2 panel figure from two images."""
    fig, axes = plt.subplots(1, 2, figsize=figsize, dpi=dpi)
    fig.subplots_adjust(wspace=0.02, left=0.01, right=0.99, top=0.95, bottom=0.02)
    
    for i, (path, label) in enumerate(zip(panel_paths, labels)):
        img = mpimg.imread(path)
        axes[i].imshow(img)
        axes[i].set_axis_off()
        axes[i].set_title(label, fontsize=14, fontweight='bold', pad=8)
    
    plt.savefig(output_path, dpi=dpi, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f"Saved: {output_path}")


# Figure 1: L008 SCR (Panel A) + L004 SCR (Panel B)
create_panel_figure(
    ['/home/ubuntu/en_map_L008_scr.png', '/home/ubuntu/en_map_L004_scr.png'],
    ['A. General anesthesia (L008) SCR', 'B. Spinal anesthesia (L004) SCR'],
    '/home/ubuntu/rapm_fig1_en.png'
)

# Figure 2: University hospital presence (Panel A) + Combined L008+L004 (Panel B)
create_panel_figure(
    ['/home/ubuntu/en_map_univ_presence.png', '/home/ubuntu/en_map_L008_L004_combined.png'],
    ['A. University hospital presence', 'B. Combined GA + spinal (L008+L004) SCR'],
    '/home/ubuntu/rapm_fig2_en.png'
)

# Figure 3: L003 SCR (Panel A) + L003/L008 ratio (Panel B)
create_panel_figure(
    ['/home/ubuntu/en_map_L003_scr.png', '/home/ubuntu/en_map_L003_L008_ratio_corrected.png'],
    ['A. Continuous epidural infusion (L003) SCR', 'B. L003/L008 ratio (epidural combination rate)'],
    '/home/ubuntu/rapm_fig3_en.png'
)

# Figure 4: 3D extruded map (single panel, larger)
fig, ax = plt.subplots(1, 1, figsize=(10, 8), dpi=300)
fig.subplots_adjust(left=0.01, right=0.99, top=0.95, bottom=0.02)
img = mpimg.imread('/home/ubuntu/3d_extruded/3D_ratio_by_anes_v2.png')
ax.imshow(img)
ax.set_axis_off()
ax.set_title('Three-dimensional extruded choropleth map', fontsize=14, fontweight='bold', pad=8)
plt.savefig('/home/ubuntu/rapm_fig4_en.png', dpi=300, bbox_inches='tight', facecolor='white')
plt.close()
print("Saved: /home/ubuntu/rapm_fig4_en.png")

# === Japanese versions ===
# Use JP source maps
jp_map_L008 = '/home/ubuntu/map_L008_scr.png'
jp_map_L004 = '/home/ubuntu/map_L004_scr.png'
jp_map_univ = '/home/ubuntu/map_univ_presence.png'
jp_map_combined = '/home/ubuntu/map_L008_L004_combined.png'
jp_map_L003 = '/home/ubuntu/map_L003_scr.png'
jp_map_ratio = '/home/ubuntu/map_L003_L008_ratio_corrected.png'
jp_3d = '/home/ubuntu/3d_extruded/3D_ratio_by_anes_v2_jp.png'

create_panel_figure(
    [jp_map_L008, jp_map_L004],
    ['A. 全身麻酔 (L008) SCR', 'B. 脊椎麻酔 (L004) SCR'],
    '/home/ubuntu/rapm_fig1_jp.png'
)

create_panel_figure(
    [jp_map_univ, jp_map_combined],
    ['A. 大学病院所在', 'B. 全身+脊椎麻酔合計 (L008+L004) SCR'],
    '/home/ubuntu/rapm_fig2_jp.png'
)

create_panel_figure(
    [jp_map_L003, jp_map_ratio],
    ['A. 持続硬膜外注入 (L003) SCR', 'B. L003/L008比（硬膜外併用率）'],
    '/home/ubuntu/rapm_fig3_jp.png'
)

fig, ax = plt.subplots(1, 1, figsize=(10, 8), dpi=300)
fig.subplots_adjust(left=0.01, right=0.99, top=0.95, bottom=0.02)
img = mpimg.imread(jp_3d)
ax.imshow(img)
ax.set_axis_off()
ax.set_title('三次元押出コロプレスマップ', fontsize=14, fontweight='bold', pad=8)
plt.savefig('/home/ubuntu/rapm_fig4_jp.png', dpi=300, bbox_inches='tight', facecolor='white')
plt.close()
print("Saved: /home/ubuntu/rapm_fig4_jp.png")

print("\nAll panel figures created successfully.")

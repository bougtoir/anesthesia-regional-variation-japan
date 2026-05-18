#!/usr/bin/env python3
"""Convert EN figure PNGs to 300 DPI TIFF for Lancet Regional Health – WP submission.

Only converts the 3 figures used in the manuscript (within the 5 display item limit).
"""

from PIL import Image
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SCRIPT_DIR)
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
LRH_DIR = os.path.join(OUTPUT_DIR, 'lancet_rh_wp')
os.makedirs(LRH_DIR, exist_ok=True)

figures = [
    ('fig1_neuropathic_unadjusted_en.png', 'LRH_Fig1_neuropathic_unadjusted.tiff'),
    ('fig2_confounder_correlations_en.png', 'LRH_Fig2_confounder_correlations.tiff'),
    ('fig4_region_unadj_vs_adj_en.png', 'LRH_Fig3_region_unadj_vs_adj.tiff'),
]

for src_name, dst_name in figures:
    src = os.path.join(OUTPUT_DIR, src_name)
    dst = os.path.join(LRH_DIR, dst_name)
    if os.path.exists(src):
        img = Image.open(src)
        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')
        img.save(dst, format='TIFF', dpi=(300, 300), compression='tiff_lzw')
        print(f'Saved: {dst} ({img.size[0]}x{img.size[1]}, 300 DPI)')
    else:
        print(f'WARNING: Source not found: {src}')

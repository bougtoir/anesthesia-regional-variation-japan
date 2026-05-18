#!/usr/bin/env python3
"""Screenshot plotly 3D HTML files v2 - Japanese versions."""
from playwright.sync_api import sync_playwright
import time

html_files = [
    ('/home/ubuntu/3d_extruded/3D_ratio_by_anes_v2_jp.html', '/home/ubuntu/3d_extruded/3D_ratio_by_anes_v2_jp.png'),
    ('/home/ubuntu/3d_extruded/3D_L008_by_anes_v2_jp.html', '/home/ubuntu/3d_extruded/3D_L008_by_anes_v2_jp.png'),
    ('/home/ubuntu/3d_extruded/3D_ratio_by_surgery_v2_jp.html', '/home/ubuntu/3d_extruded/3D_ratio_by_surgery_v2_jp.png'),
]

with sync_playwright() as p:
    browser = p.chromium.launch(args=['--no-sandbox', '--enable-unsafe-swiftshader'])
    for html_path, png_path in html_files:
        print(f"Processing {html_path}...")
        page = browser.new_page(viewport={'width': 1800, 'height': 1400})
        page.goto(f'file://{html_path}')
        time.sleep(6)
        page.screenshot(path=png_path, full_page=False)
        page.close()
        print(f"  Saved: {png_path}")
    browser.close()

print("All JP v2 screenshots saved!")

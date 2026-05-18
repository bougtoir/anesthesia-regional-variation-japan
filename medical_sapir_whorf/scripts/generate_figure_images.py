#!/usr/bin/env python3
"""
Generate figure PNG images for inline embedding in docx manuscripts.
Outputs: medical_sapir_whorf/output/figure1_en.png, figure1_ja.png,
         figure2_en.png, figure2_ja.png
"""

import os
import sys

# Reuse figure generation from pptx script
sys.path.insert(0, os.path.dirname(__file__))
from create_figures_pptx import (
    create_figure1_en, create_figure1_ja,
    create_figure2_en, create_figure2_ja,
)

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
os.makedirs(OUT_DIR, exist_ok=True)

for name, func in [
    ("figure1_en.png", create_figure1_en),
    ("figure2_en.png", create_figure2_en),
    ("figure1_ja.png", create_figure1_ja),
    ("figure2_ja.png", create_figure2_ja),
]:
    buf = func()
    path = os.path.join(OUT_DIR, name)
    with open(path, "wb") as f:
        f.write(buf.read())
    print(f"Saved: {path}")

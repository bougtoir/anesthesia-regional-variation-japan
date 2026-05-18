#!/usr/bin/env python3
"""
琵琶湖考古学的探査パイプライン実行スクリプト

Usage:
    python -m biwa_archaeological_prospection.run [--output OUTPUT_DIR] [--zoom ZOOM]
"""

import argparse
import logging
import os
import sys

from .pipeline import run_pipeline


def main():
    parser = argparse.ArgumentParser(
        description="琵琶湖考古学的探査パイプライン / Lake Biwa Archaeological Prospection Pipeline",
    )
    parser.add_argument(
        "--output", "-o", default="biwa_archaeological_prospection/output",
        help="出力ディレクトリ (default: biwa_archaeological_prospection/output)",
    )
    parser.add_argument(
        "--zoom", "-z", type=int, default=14,
        help="DEMタイルのズームレベル (default: 14, ~10m resolution)",
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true",
        help="詳細ログ出力",
    )
    args = parser.parse_args()

    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    print("=" * 60)
    print("琵琶湖考古学的探査パイプライン")
    print("Lake Biwa Archaeological Prospection Pipeline")
    print("=" * 60)
    print(f"Output: {output_dir}")
    print(f"Zoom level: {args.zoom} (~{2**args.zoom} tiles/axis)")
    print()

    results, files = run_pipeline(
        output_dir=output_dir,
        zoom=args.zoom,
    )

    print()
    print("=" * 60)
    print("Results Summary")
    print("=" * 60)
    for result in results:
        print(f"\n  {result.region.name}:")
        print(f"    DEM: {result.dem.shape[0]}x{result.dem.shape[1]} @ {result.cell_size:.1f}m")
        print(f"    Candidates: {len(result.candidates)}")
        for i, c in enumerate(result.candidates[:5]):
            print(f"      #{i+1}: ({c.lat:.4f}, {c.lon:.4f}) score={c.score:.3f} area={c.area_m2:.0f}m²")

    print(f"\nGenerated {len(files)} files in {output_dir}/")
    for f in files:
        print(f"  - {os.path.basename(f)}")


if __name__ == "__main__":
    main()

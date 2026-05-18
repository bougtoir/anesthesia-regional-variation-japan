"""
run_pipeline.py
PWV-VitalDB解析パイプラインのメインランナー

Usage:
    python run_pipeline.py              # Run full pipeline
    python run_pipeline.py --quick      # Quick mode (fewer cases, for testing)
    python run_pipeline.py --step N     # Run specific step only
"""

import os
import sys
import time
import argparse

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PARENT_DIR)


def run_step(step_num, step_name, module_func, **kwargs):
    print(f"\n{'#'*70}")
    print(f"# Step {step_num}: {step_name}")
    print(f"{'#'*70}")
    t0 = time.time()
    try:
        result = module_func(**kwargs)
        elapsed = time.time() - t0
        print(f"\n>>> Step {step_num} completed in {elapsed:.1f}s")
        return result
    except Exception as e:
        elapsed = time.time() - t0
        print(f"\n>>> Step {step_num} FAILED after {elapsed:.1f}s: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--quick", action="store_true", help="Quick mode for testing")
    parser.add_argument("--step", type=int, default=None, help="Run specific step")
    args = parser.parse_args()

    total_start = time.time()
    print("=" * 70)
    print("PWV-VitalDB Analysis Pipeline")
    print(f"Mode: {'Quick' if args.quick else 'Full'}")
    print("=" * 70)

    if args.quick:
        import scripts.s01_data_acquisition as s01
        s01.MAX_CASES_PHASE1 = 30
        s01.MAX_CASES_FULL = 30

    # Step 1: Data Acquisition
    if args.step is None or args.step == 1:
        from scripts.s01_data_acquisition import main as step1
        run_step(1, "Data Acquisition", step1)

    # Step 2: PWV Calculation
    if args.step is None or args.step == 2:
        from scripts.s02_pwv_calculation import process_all_cases as step2
        run_step(2, "PWV Calculation", step2)

    # Step 3: Correlation Analysis
    if args.step is None or args.step == 3:
        from scripts.s03_correlation_analysis import main as step3
        run_step(3, "Correlation Analysis", step3)

    # Step 4: Prediction Models
    if args.step is None or args.step == 4:
        from scripts.s04_sofa_prediction_model import main as step4
        run_step(4, "SOFA Prediction Models", step4)

    # Step 5: Report Generation
    if args.step is None or args.step == 5:
        from scripts.s05_generate_report import main as step5
        run_step(5, "Report Generation", step5)

    total_elapsed = time.time() - total_start
    print(f"\n{'='*70}")
    print(f"Pipeline completed in {total_elapsed:.1f}s ({total_elapsed/60:.1f}min)")
    print(f"{'='*70}")


if __name__ == "__main__":
    main()

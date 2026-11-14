"""
Main Entry Point for AI/ML-Powered STAAD.Pro Governing Force Extraction System.
Supports both Command-Line Interface (CLI) and Tkinter Graphical User Interface (GUI).
"""

import os
import sys
import argparse

from staad_processor import process_staad_file, APP_NAME, APP_VERSION, DEFAULT_TARGET_SHEETS
from gui import launch_gui


def parse_args():
    parser = argparse.ArgumentParser(
        description=f"{APP_NAME} v{APP_VERSION} - AI/ML Semantic Interpretation & Deterministic Force Extraction."
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        default=None,
        help="Path to STAAD.Pro Excel workbook (e.g. R-L, R-R, JB-F, JB-B, JB, TB).",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Path for generated output Excel file (.xlsx). Defaults to [input_name]_Governing_Forces_STEP1.xlsx",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Launch the Tkinter Graphical User Interface.",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    if args.gui or args.input_file is None:
        launch_gui()
        return

    input_path = os.path.abspath(args.input_file)
    output_path = os.path.abspath(args.output) if args.output else None

    print("=" * 80)
    print(f" {APP_NAME} (v{APP_VERSION})")
    print(f" Target Worksheets: {', '.join(DEFAULT_TARGET_SHEETS)}")
    print("=" * 80)
    print(f"Input Workbook:  {input_path}")
    if output_path:
        print(f"Output Workbook: {output_path}")

    def console_progress(pct: float, msg: str):
        print(f"[{int(pct):3d}%] {msg}")

    try:
        out_file, node_results, val_report, sheets_loaded = process_staad_file(
            input_file=input_path,
            output_file=output_path,
            progress_callback=console_progress,
        )

        print("\n" + "=" * 80)
        print(" STEP 1 EXTRACTION RESULTS SUMMARY")
        print("=" * 80)
        print(f"Source Worksheets Loaded: {', '.join(sheets_loaded)}")
        print(f"Total Unique Nodes:       {len(node_results)}")
        print(f"Generated Output File:    {out_file}")
        print("\nPer-Node Connected Members & Governing Load Cases:")
        for node_id, a in node_results.items():
            print(
                f"  • Node {node_id:>3}: Connected Members: {a.connected_members} | "
                f"Unique Governing L/Cs: {a.unique_governing_lcs}"
            )

        print("\n" + "=" * 80)
        print(" VALIDATION STATUS")
        print("=" * 80)
        print(f"Health:         {'PASSED (100% Verified)' if val_report.is_valid else 'FAILED / WARNINGS'}")
        print(f"Checks Passed:  {val_report.checks_passed}")
        print(f"Checks Failed:  {val_report.checks_failed}")
        if val_report.invalid_data_warnings:
            print("\nWarnings:")
            for w in val_report.invalid_data_warnings:
                print(f"  {w}")
        if val_report.errors:
            print("\nErrors:")
            for err in val_report.errors:
                print(f"  {err}")

        print("=" * 80)

    except Exception as e:
        print(f"\n[FATAL ERROR] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()

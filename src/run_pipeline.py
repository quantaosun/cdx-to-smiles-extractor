"""
run_pipeline.py
---------------
Convenience wrapper that runs the full pipeline in one command:
    PPTX → JSON (extract_cdx.py) → Excel (build_excel.py)

Usage
-----
    python src/run_pipeline.py path/to/presentation.pptx

    # Custom output paths
    python src/run_pipeline.py path/to/presentation.pptx \\
        --json  output/my_structures.json \\
        --excel output/my_structures.xlsx
"""

import argparse
import sys
from pathlib import Path

# Allow running from repo root
sys.path.insert(0, str(Path(__file__).parent))

from extract_cdx import process_pptx
from build_excel import build_excel


def run(pptx_path: str, json_path: str, excel_path: str):
    print("=" * 60)
    print("Step 1/2 — Extracting CDX structures and converting to SMILES")
    print("=" * 60)
    results = process_pptx(pptx_path, json_path)

    if not results:
        print("\n[ERROR] No structures were extracted. Aborting.")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("Step 2/2 — Building Excel file")
    print("=" * 60)
    build_excel(json_path, excel_path)

    print("\n" + "=" * 60)
    print("Pipeline complete.")
    print(f"  JSON:  {json_path}")
    print(f"  Excel: {excel_path}")
    print("=" * 60)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Run the full CDX extraction and SMILES conversion pipeline.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('input_pptx', help='Path to the input PPTX file.')
    parser.add_argument(
        '--json', '-j',
        default='output/structures.json',
        help='Path to intermediate JSON output (default: output/structures.json).'
    )
    parser.add_argument(
        '--excel', '-e',
        default='output/structures.xlsx',
        help='Path to final Excel output (default: output/structures.xlsx).'
    )
    args = parser.parse_args()
    run(args.input_pptx, args.json, args.excel)

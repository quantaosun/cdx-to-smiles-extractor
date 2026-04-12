"""
cdx_to_smiles.py
----------------
Convert one or more standalone CDX binary files to SMILES strings.
This is a lightweight utility for when you already have extracted CDX files
and just need the SMILES — no PPTX required.

Usage:
    # Single file
    python src/cdx_to_smiles.py path/to/structure.cdx

    # Directory of CDX files
    python src/cdx_to_smiles.py path/to/cdx_dir/ --output results.csv

    # Glob pattern
    python src/cdx_to_smiles.py "examples/input/cdx/*.cdx" --output results.csv
"""

import os
import re
import csv
import glob
import logging
import tempfile
import argparse
from pathlib import Path

logging.getLogger().setLevel(logging.ERROR)

try:
    from pycdxml import cdxml_converter
    from rdkit import Chem
except ImportError:
    print("Error: Required packages not found.")
    print("Please run: pip install -r requirements.txt")
    exit(1)


def cdx_to_smiles(cdx_path: str) -> list[str]:
    """
    Read a CDX binary file and return a list of SMILES strings.
    Returns an empty list if conversion fails.
    """
    try:
        doc = cdxml_converter.read_cdx(cdx_path)
        cdxml_str = doc.to_cdxml()
        mols = Chem.MolsFromCDXML(cdxml_str)
        # isomericSmiles=True: encode @/@@, /, \\ for defined stereocenters
        # canonical=True: Morgan-algorithm canonical atom ordering
        return [
            Chem.MolToSmiles(m, isomericSmiles=True, canonical=True)
            for m in mols if m is not None
        ]
    except Exception as e:
        print(f"  Error processing {cdx_path}: {e}")
        return []


def process_files(paths: list[str], output_csv: str | None = None):
    """Process a list of CDX file paths and print/save results."""
    results = []

    for path in paths:
        name = Path(path).stem
        smiles_list = cdx_to_smiles(path)

        if not smiles_list:
            print(f"[WARN] {name}: No SMILES generated")
            results.append({'id': name, 'smiles': '', 'file': path})
        else:
            for i, smi in enumerate(smiles_list):
                label = name if len(smiles_list) == 1 else f"{name}_{i+1}"
                print(f"[OK]   {label}: {smi}")
                results.append({'id': label, 'smiles': smi, 'file': path})

    if output_csv:
        with open(output_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['id', 'smiles', 'file'])
            writer.writeheader()
            writer.writerows(results)
        print(f"\nResults saved to: {output_csv}")

    return results


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Convert CDX binary files to SMILES strings.'
    )
    parser.add_argument(
        'input',
        help='Path to a CDX file, a directory of CDX files, or a glob pattern.'
    )
    parser.add_argument(
        '--output', '-o',
        default=None,
        help='Path to output CSV file (optional). If omitted, results are printed only.'
    )
    args = parser.parse_args()

    # Resolve input to a list of CDX file paths
    input_path = Path(args.input)
    if input_path.is_dir():
        cdx_files = sorted(input_path.glob('*.cdx'))
    elif '*' in args.input or '?' in args.input:
        cdx_files = sorted(Path(p) for p in glob.glob(args.input))
    elif input_path.is_file():
        cdx_files = [input_path]
    else:
        print(f"Error: Input path not found: {args.input}")
        exit(1)

    if not cdx_files:
        print("No CDX files found.")
        exit(1)

    print(f"Processing {len(cdx_files)} CDX file(s)...\n")
    process_files([str(p) for p in cdx_files], args.output)

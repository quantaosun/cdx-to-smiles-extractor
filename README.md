# CDX-to-SMILES Extractor

A Python pipeline to automatically extract **ChemDraw CDX structures** embedded as OLE objects inside PowerPoint (`.pptx`) files, convert each structure to a **canonical SMILES string**, and export the results — together with any embedded metadata — to a formatted **Excel spreadsheet**.

> **No ChemDraw installation required.** The pipeline uses only open-source tools: [`pycdxml`](https://github.com/kienerj/pycdxml) for CDX ↔ CDXML conversion and [`RDKit`](https://www.rdkit.org/) for SMILES generation.

---

## Table of Contents

- [Background](#background)
- [How It Works](#how-it-works)
- [Repository Structure](#repository-structure)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Script Reference](#script-reference)
- [Example Output](#example-output)
- [Troubleshooting](#troubleshooting)
- [Dependencies](#dependencies)
- [License](#license)

---

## Background

When a ChemDraw structure is **pasted into Microsoft PowerPoint**, it is stored as an OLE2 Compound Document (`oleObjectX.bin`) inside the PPTX ZIP archive under `ppt/embeddings/`. Each OLE container holds a `CONTENTS` stream with the raw CDX binary data (identifiable by the magic bytes `VjCD0100`).

This repository provides a complete pipeline to:

1. Unzip the PPTX and locate all embedded OLE objects.
2. Extract the raw CDX binary from each OLE container.
3. Convert CDX → CDXML (XML format) using `pycdxml`.
4. Generate canonical SMILES from CDXML using `RDKit`.
5. Parse embedded metadata annotations (compound ID, MW, CLogP, etc.).
6. Export everything to a structured Excel file.

---

## How It Works

```
PPTX file
  └── ppt/embeddings/oleObject1.bin   (OLE2 container)
        └── CONTENTS stream           (raw CDX binary: magic VjCD0100)
              │
              ▼  pycdxml.cdxml_converter.read_cdx()
           CDXML string (XML)
              │
              ├── RDKit.Chem.MolsFromCDXML()  →  canonical SMILES
              └── xml.etree.ElementTree       →  metadata annotations
                                                  (CpdIndex, MW, CLogP, ...)
                                                        │
                                                        ▼
                                               structures.json
                                                        │
                                                        ▼
                                               structures.xlsx
```

---

## Repository Structure

```
cdx-to-smiles-extractor/
│
├── src/
│   ├── extract_cdx.py          # Step 1: PPTX → CDX → CDXML → SMILES → JSON
│   ├── build_excel.py          # Step 2: JSON → formatted Excel
│   ├── run_pipeline.py         # Convenience wrapper: runs both steps
│   └── cdx_to_smiles.py        # Standalone: convert individual CDX files
│
├── examples/
│   ├── input/
│   │   └── cdx/                # Example CDX files (6 public drug molecules)
│   │       ├── Aspirin.cdx
│   │       ├── Caffeine.cdx
│   │       ├── Ibuprofen.cdx
│   │       ├── Paracetamol.cdx
│   │       ├── Metformin.cdx
│   │       └── Atorvastatin.cdx
│   ├── output/
│   │   ├── example_structures.json   # Example JSON output
│   │   └── example_structures.xlsx  # Example Excel output
│   ├── generate_example_cdx.py      # Regenerate example CDX files
│   └── generate_example_pptx.py     # Generate a demo PPTX with placeholders
│
├── output/                     # Default output directory (git-ignored)
├── requirements.txt
├── .gitignore
├── LICENSE
└── README.md
```

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/cdx-to-smiles-extractor.git
cd cdx-to-smiles-extractor
```

### 2. Create a virtual environment (recommended)

```bash
python -m venv .venv
# On Linux/macOS:
source .venv/bin/activate
# On Windows:
.venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

> **Note on `pycdxml`:** This package is not on PyPI and must be installed directly from GitHub. The `requirements.txt` handles this automatically via the `git+https://...` syntax. Git must be installed on your system.

> **Note on `rdkit`:** RDKit is available on PyPI as `rdkit` (version ≥ 2022). If you are using a conda environment, you can alternatively install it with `conda install -c conda-forge rdkit`.

---

## Quick Start

### Full pipeline (PPTX → Excel in one command)

```bash
python src/run_pipeline.py path/to/your_presentation.pptx
```

This produces two files in the `output/` directory:
- `output/structures.json` — intermediate data with all metadata
- `output/structures.xlsx` — final formatted Excel spreadsheet

### Step-by-step

```bash
# Step 1: Extract structures from PPTX and convert to SMILES
python src/extract_cdx.py path/to/your_presentation.pptx -o output/structures.json

# Step 2: Build the Excel file
python src/build_excel.py output/structures.json -o output/structures.xlsx
```

### Convert standalone CDX files (no PPTX needed)

```bash
# Single file
python src/cdx_to_smiles.py examples/input/cdx/Aspirin.cdx

# Whole directory → CSV
python src/cdx_to_smiles.py examples/input/cdx/ --output output/results.csv
```

---

## Script Reference

### `src/extract_cdx.py`

| Argument | Description |
|---|---|
| `input_pptx` | Path to the input `.pptx` file |
| `--output` / `-o` | Path to the output JSON file (default: `output/structures.json`) |

**Output JSON structure per record:**

```json
{
  "index": 1,
  "file": "oleObject1.bin",
  "compound_id": "Aspirin",
  "smiles": "CC(=O)Oc1ccccc1C(=O)O",
  "all_smiles": ["CC(=O)Oc1ccccc1C(=O)O"],
  "annotations": {
    "CpdIndex": "Aspirin",
    "MW": "180.16",
    "CLogP": "1.19"
  },
  "text_labels": ["Aspirin"]
}
```

### `src/build_excel.py`

| Argument | Description |
|---|---|
| `input_json` | Path to the JSON file from `extract_cdx.py` |
| `--output` / `-o` | Path to the output `.xlsx` file (default: `output/structures.xlsx`) |

The Excel file always contains `Compound_ID` and `SMILES` as the first two columns. All annotation keys found in the data (MW, CLogP, SeqCount_*, etc.) are appended as additional columns automatically.

### `src/run_pipeline.py`

| Argument | Description |
|---|---|
| `input_pptx` | Path to the input `.pptx` file |
| `--json` / `-j` | Path to intermediate JSON (default: `output/structures.json`) |
| `--excel` / `-e` | Path to final Excel (default: `output/structures.xlsx`) |

### `src/cdx_to_smiles.py`

| Argument | Description |
|---|---|
| `input` | Path to a CDX file, directory of CDX files, or glob pattern |
| `--output` / `-o` | Path to output CSV file (optional; prints to stdout if omitted) |

---

## Example Output

The `examples/output/` directory contains pre-generated output from 6 well-known public drug molecules:

| Compound_ID | SMILES |
|---|---|
| Aspirin | `CC(=O)Oc1ccccc1C(=O)O` |
| Caffeine | `Cn1c(=O)c2c(ncn2C)n(C)c1=O` |
| Ibuprofen | `CC(C)Cc1ccc(C(C)C(=O)O)cc1` |
| Paracetamol | `CC(=O)Nc1ccc(O)cc1` |
| Metformin | `CN(C)C(=N)NC(=N)N` |
| Atorvastatin | `CC(C)c1c(C(=O)Nc2ccccc2F)c(-c2ccccc2)c(-c2ccc(F)cc2)n1CCC(O)CC(O)CC(=O)O` |

To regenerate the example CDX files:

```bash
python examples/generate_example_cdx.py
```

---

## Troubleshooting

**`No embedded OLE objects found`**
The PPTX file does not contain ChemDraw structures pasted as native OLE objects. Structures that were inserted as images (PNG/EMF) cannot be processed by this pipeline.

**`Unexpected magic bytes`**
The embedded OLE object is not a CDX file (e.g., it may be an Excel chart or another OLE type). These are automatically skipped.

**`No SMILES generated`**
The CDX file was parsed successfully but RDKit could not interpret the molecule. This can happen with very complex structures, reaction arrows, biological sequences, or polymer notation. Check the intermediate CDXML file for clues.

**`pycdxml` install fails**
Ensure Git is installed and accessible from your terminal. The `git+https://...` install syntax requires Git. Alternatively, clone `pycdxml` manually and install with `pip install -e /path/to/pycdxml`.

**Multiple SMILES per CDX object**
Some CDX objects contain more than one molecule (e.g., a structure plus a reagent drawn nearby). The pipeline uses the first molecule by default. All SMILES are saved in the `all_smiles` field of the JSON for inspection.

---

## Dependencies

| Package | Version | Purpose |
|---|---|---|
| `rdkit` | ≥ 2023.03 | SMILES generation from CDXML |
| `pycdxml` | GitHub HEAD | CDX ↔ CDXML conversion |
| `olefile` | ≥ 0.46 | Reading OLE2 compound documents |
| `pandas` | ≥ 2.0 | DataFrame construction |
| `openpyxl` | ≥ 3.1 | Excel file writing and formatting |

---

## License

MIT License. See [LICENSE](LICENSE) for details.

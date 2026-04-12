"""
generate_example_cdx.py
-----------------------
Generates 6 real CDX binary files from well-known public drug molecules.
These can be used to test the CDX-to-SMILES conversion pipeline directly
without needing a PPTX file.

Run:
    python examples/generate_example_cdx.py

Output:
    examples/input/cdx/  — one .cdx file per molecule
"""

import os
import tempfile
from pathlib import Path
from rdkit import Chem
from rdkit.Chem import AllChem
from pycdxml import cdxml_converter
import xml.etree.ElementTree as ET

MOLECULES = [
    {"id": "Aspirin",      "smiles": "CC(=O)Oc1ccccc1C(=O)O"},
    {"id": "Caffeine",     "smiles": "Cn1cnc2c1c(=O)n(C)c(=O)n2C"},
    {"id": "Ibuprofen",    "smiles": "CC(C)Cc1ccc(cc1)C(C)C(=O)O"},
    {"id": "Paracetamol",  "smiles": "CC(=O)Nc1ccc(O)cc1"},
    {"id": "Metformin",    "smiles": "CN(C)C(=N)NC(=N)N"},
    {"id": "Atorvastatin", "smiles": "CC(C)c1c(C(=O)Nc2ccccc2F)c(-c2ccccc2)c(-c2ccc(F)cc2)n1CCC(O)CC(O)CC(=O)O"},
]

def smiles_to_cdx(smiles: str, compound_id: str, output_path: str):
    """Convert SMILES to a CDX binary file with an embedded CpdIndex annotation."""
    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        raise ValueError(f"Invalid SMILES: {smiles}")
    AllChem.Compute2DCoords(mol)

    # Convert mol to CDXML document
    doc = cdxml_converter.mol_to_document(mol)
    cdxml_str = doc.to_cdxml()

    # Inject CpdIndex annotation
    root = ET.fromstring(cdxml_str)
    page = root.find('page')
    if page is not None:
        grp = page.find('group')
        target = grp if grp is not None else page
        ann = ET.SubElement(target, 'annotation')
        ann.set('id', '0')
        ann.set('Keyword', 'CpdIndex')
        ann.set('Content', compound_id)

    modified_cdxml = '<?xml version="1.0" encoding="UTF-8" ?>\n' + \
                     ET.tostring(root, encoding='unicode', xml_declaration=False)

    with tempfile.NamedTemporaryFile(suffix='.cdxml', delete=False, mode='w', encoding='utf-8') as f:
        f.write(modified_cdxml)
        tmp_cdxml = f.name

    try:
        doc2 = cdxml_converter.read_cdxml(tmp_cdxml)
        cdxml_converter.write_cdx_file(doc2, output_path)
        print(f"  Saved: {output_path}")
    finally:
        os.unlink(tmp_cdxml)


if __name__ == '__main__':
    out_dir = Path(__file__).parent / 'input' / 'cdx'
    out_dir.mkdir(parents=True, exist_ok=True)

    print("Generating example CDX files...\n")
    for mol in MOLECULES:
        out_path = out_dir / f"{mol['id']}.cdx"
        smiles_to_cdx(mol['smiles'], mol['id'], str(out_path))

    print(f"\nDone. {len(MOLECULES)} CDX files written to {out_dir}")

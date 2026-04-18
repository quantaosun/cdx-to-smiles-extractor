"""
extract_cdx.py
--------------
Extract all ChemDraw CDX structures embedded as OLE objects inside a PowerPoint
(PPTX) file, convert each to SMILES using pycdxml + RDKit, and save the results
(including all embedded metadata annotations) to a JSON file.

Background
----------
When a ChemDraw structure is pasted into PowerPoint, it is stored as an OLE2
Compound Document (oleObjectX.bin) inside the PPTX ZIP archive under
`ppt/embeddings/`. Each OLE container has a `CONTENTS` stream that holds the
raw CDX binary data (magic bytes: `VjCD0100`).

Pipeline
--------
1. Unzip PPTX  →  2. Extract CDX from OLE  →  3. CDX → CDXML (pycdxml)
→  4. CDXML → SMILES (RDKit)  →  5. Parse metadata annotations  →  6. JSON

Usage
-----
    python src/extract_cdx.py path/to/presentation.pptx -o output/structures.json
"""

import os
import re
import json
import logging
import zipfile
import shutil
import tempfile
import argparse
import olefile
import xml.etree.ElementTree as ET
from pathlib import Path

# Suppress verbose pycdxml warnings about minor CDX property length mismatches
logging.getLogger().setLevel(logging.ERROR)

try:
    from pycdxml import cdxml_converter
    from rdkit import Chem
except ImportError as e:
    print(f"ImportError: {e}")
    print("\nPlease install all required dependencies:")
    print("    pip install -r requirements.txt")
    exit(1)


# ── Helpers ──────────────────────────────────────────────────────────────────

def _natural_sort_key(s: str) -> list:
    """Sort key for natural ordering of filenames (e.g. ole1 < ole2 < ole10)."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', s)]


def _row_band(y: int, band_size: int = 2_000_000) -> int:
    """
    Quantise a y-coordinate (in EMUs) into a row band.

    Structures on the same visual row in a slide have very similar y-coordinates
    but not identical ones (small alignment offsets). Quantising into bands of
    ~2 mm (2 000 000 EMU ≈ 56 pt) groups them into the same row so that
    left-to-right ordering within a row is applied correctly.
    """
    return (y // band_size) * band_size


def _build_slide_position_map(pptx_path: str, extract_dir: str) -> dict[str, tuple[int, int, int]]:
    """
    Parse all slide XMLs and return a mapping of OLE filename → (slide_num, row_band, x).

    This is used to sort extracted structures in the same visual order as they
    appear in the presentation: slide by slide, top-to-bottom row, left-to-right
    within each row.

    Parameters
    ----------
    pptx_path : str
        Path to the PPTX file (already extracted to extract_dir).
    extract_dir : str
        Directory where the PPTX was extracted.

    Returns
    -------
    dict mapping ole_filename (e.g. 'oleObject3.bin') to (slide_num, row_band, x).
    """
    P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    position_map = {}
    base = Path(extract_dir)

    try:
        prs_root = ET.parse(str(base / 'ppt' / 'presentation.xml')).getroot()
        prs_rels = ET.parse(str(base / 'ppt' / '_rels' / 'presentation.xml.rels')).getroot()
        rid_to_target = {r.get('Id'): r.get('Target') for r in prs_rels}

        slide_rids = [sl.get(f'{{{R}}}id') for sl in prs_root.iter(f'{{{P}}}sldId')]

        for slide_num, rid in enumerate(slide_rids, 1):
            target = rid_to_target.get(rid, '')
            slide_file = target.lstrip('/').lstrip('../')
            if not slide_file.startswith('ppt/'):
                slide_file = 'ppt/' + slide_file
            slide_path = base / slide_file
            if not slide_path.exists():
                continue

            slide_root = ET.parse(str(slide_path)).getroot()
            slide_name = slide_path.name
            rels_path = base / 'ppt' / 'slides' / '_rels' / f'{slide_name}.rels'
            if not rels_path.exists():
                continue

            rels_root = ET.parse(str(rels_path)).getroot()
            rid_to_ole = {
                r.get('Id'): Path(r.get('Target', '')).name
                for r in rels_root
                if 'oleObject' in r.get('Target', '')
            }

            for gf in slide_root.iter(f'{{{P}}}graphicFrame'):
                ole_el = gf.find(f'.//{{{P}}}oleObj')
                if ole_el is None:
                    continue
                frame_rid = ole_el.get(f'{{{R}}}id', '')
                ole_name = rid_to_ole.get(frame_rid, '')
                if not ole_name:
                    continue
                xfrm = gf.find(f'.//{{{A}}}xfrm')
                off = xfrm.find(f'{{{A}}}off') if xfrm is not None else None
                x = int(off.get('x', 0)) if off is not None else 0
                y = int(off.get('y', 0)) if off is not None else 0
                position_map[ole_name] = (slide_num, _row_band(y), x)

    except Exception as e:
        print(f"  [WARN] Could not build slide position map: {e}")

    return position_map


def _extract_pptx(pptx_path: str, extract_dir: str) -> list[Path]:
    """
    Unzip a PPTX file and return OLE embedding paths sorted by visual slide order.

    Ordering: slide number → top-to-bottom row band → left-to-right x position.
    This matches the visual reading order of structures in the presentation.
    Falls back to natural filename sort if position data is unavailable.

    Parameters
    ----------
    pptx_path : str
        Absolute or relative path to the .pptx file.
    extract_dir : str
        Directory to extract the PPTX contents into.

    Returns
    -------
    list[Path]
        Position-sorted list of paths to oleObjectX.bin files.
    """
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        zf.extractall(extract_dir)

    embeddings_dir = Path(extract_dir) / 'ppt' / 'embeddings'
    if not embeddings_dir.exists():
        return []

    ole_files = [f for f in os.listdir(embeddings_dir) if f.endswith('.bin')]

    # Build slide-position map and sort by (slide, row_band, x)
    position_map = _build_slide_position_map(pptx_path, extract_dir)

    def _sort_key(fname: str):
        if fname in position_map:
            return position_map[fname]          # (slide_num, row_band, x)
        return (9999, 9999, _natural_sort_key(fname)[0])  # unmapped: append at end

    ole_files_sorted = sorted(ole_files, key=_sort_key)
    return [embeddings_dir / f for f in ole_files_sorted]


def _extract_cdx_from_ole(ole_path: Path) -> bytes | None:
    """
    Extract the raw CDX binary from an OLE2 container.

    ChemDraw stores CDX data in the `CONTENTS` stream of the OLE file.

    Parameters
    ----------
    ole_path : Path
        Path to the OLE binary file.

    Returns
    -------
    bytes or None
        Raw CDX binary data, or None on failure.
    """
    try:
        ole = olefile.OleFileIO(str(ole_path))
        if ole.exists('CONTENTS'):
            data = ole.openstream('CONTENTS').read()
            return data
    except Exception as e:
        print(f"  [ERROR] Reading OLE {ole_path.name}: {e}")
    return None


def _cdx_to_cdxml(cdx_data: bytes, tmp_path: str) -> str | None:
    """
    Convert raw CDX binary to a CDXML string using pycdxml.

    Parameters
    ----------
    cdx_data : bytes
        Raw CDX binary data.
    tmp_path : str
        Temporary file path to write the CDX data to.

    Returns
    -------
    str or None
        CDXML string, or None on failure.
    """
    with open(tmp_path, 'wb') as f:
        f.write(cdx_data)
    try:
        doc = cdxml_converter.read_cdx(tmp_path)
        return doc.to_cdxml()
    except Exception as e:
        print(f"  [ERROR] CDX → CDXML conversion: {e}")
        return None


def _cdxml_to_smiles(cdxml_str: str) -> list[str]:
    """
    Convert a CDXML string to a list of canonical SMILES using RDKit.

    Generates **isomeric canonical SMILES** (``isomericSmiles=True``, the RDKit
    default), which encodes:
    - Tetrahedral stereocenters as ``@`` / ``@@``
    - E/Z double-bond geometry as ``/`` / ``\\``

    Stereocenters that exist in the molecular graph but were drawn *without a
    wedge bond* in ChemDraw are reported as ``?`` by RDKit and are intentionally
    omitted from the SMILES string — this is the correct canonical behaviour for
    undefined stereochemistry.

    Parameters
    ----------
    cdxml_str : str
        CDXML-formatted string.

    Returns
    -------
    list[str]
        List of canonical isomeric SMILES strings (one per molecule found).
    """
    try:
        mols = Chem.MolsFromCDXML(cdxml_str)
        smiles_list = []
        for m in mols:
            if m is not None:
                # isomericSmiles=True (default): preserves @/@@, /, \\
                # canonical=True (default): Morgan-algorithm canonical ordering
                smi = Chem.MolToSmiles(m, isomericSmiles=True, canonical=True)
                smiles_list.append(smi)
        return smiles_list
    except Exception as e:
        print(f"  [ERROR] CDXML → SMILES: {e}")
        return []


def _extract_metadata(cdxml_str: str) -> dict:
    """
    Parse a CDXML string and extract compound ID, text labels, and annotations.

    ChemDraw embeds metadata as `<annotation Keyword="..." Content="..."/>` tags.
    The `CpdIndex` or `CompoundIndex` keyword typically holds the compound ID.

    Parameters
    ----------
    cdxml_str : str
        CDXML-formatted string.

    Returns
    -------
    dict with keys:
        - 'compound_id' (str or None)
        - 'text_labels' (list[str])
        - 'annotations' (dict[str, str])
    """
    meta = {'compound_id': None, 'text_labels': [], 'annotations': {}}
    try:
        root = ET.fromstring(cdxml_str)
        page = root.find('page')
        if page is None:
            page = root  # fallback: search whole document

        # KEY FIX: scope to the <group> that contains the molecule (<fragment>).
        # Each CDX OLE object may carry loose page-level <annotation> tags that
        # belong to a *neighbouring* compound drawn on the same slide cell.
        # Scanning root.iter() picks those up and assigns the wrong ID.
        # By restricting to the group that owns the fragment we guarantee a
        # 1-to-1 mapping between structure and ID.
        search_scope = page
        groups_with_mol = [g for g in page.findall('group')
                           if list(g.iter('fragment'))]
        if groups_with_mol:
            search_scope = groups_with_mol[0]

        for ann in search_scope.iter('annotation'):
            kw = ann.get('Keyword', '')
            ct = ann.get('Content', '')
            if kw and ct:
                meta['annotations'][kw] = ct
                if kw in ('CpdIndex', 'CompoundIndex') and not meta['compound_id']:
                    meta['compound_id'] = ct

        for t in search_scope.iter('t'):
            for s in t.findall('s'):
                if s.text and s.text.strip():
                    text = s.text.strip()
                    if not meta['compound_id'] and '\n' not in text:
                        meta['compound_id'] = text
                    meta['text_labels'].append(text)

        # Last resort: first line of first text label
        if not meta['compound_id'] and meta['text_labels']:
            meta['compound_id'] = meta['text_labels'][0].split('\n')[0].strip()

    except Exception as e:
        print(f"  [WARN] Metadata extraction: {e}")

    return meta


def _stereo_note(smiles: str | None) -> str:
    """
    Return a human-readable stereo annotation for a SMILES string.

    Categories
    ----------
    - ``defined`` : all stereocenters are assigned (@ / @@ present)
    - ``E/Z only`` : only double-bond geometry is defined
    - ``defined + E/Z`` : both tetrahedral and double-bond stereo present
    - ``unassigned`` : stereocenters exist in topology but none are wedged
    - ``none`` : no stereocenters in the molecule
    """
    if not smiles:
        return 'none'
    has_tet = '@' in smiles
    has_dbl = '/' in smiles or chr(92) in smiles
    if has_tet and has_dbl:
        return 'defined + E/Z'
    if has_tet:
        return 'defined'
    if has_dbl:
        return 'E/Z only'
    # Check for potential unassigned stereocenters
    try:
        mol = Chem.MolFromSmiles(smiles)
        if mol:
            Chem.AssignStereochemistry(mol, cleanIt=True, force=True)
            centers = Chem.FindMolChiralCenters(mol, includeUnassigned=True)
            if any(c[1] == '?' for c in centers):
                return 'unassigned'
    except Exception:
        pass
    return 'none'


# ── Main pipeline ─────────────────────────────────────────────────────────────

def process_pptx(pptx_path: str, output_json: str) -> list[dict]:
    """
    Full pipeline: PPTX → CDX → CDXML → SMILES + metadata → JSON.

    Parameters
    ----------
    pptx_path : str
        Path to the input PPTX file.
    output_json : str
        Path to write the output JSON file.

    Returns
    -------
    list[dict]
        List of result dictionaries, one per successfully processed structure.
    """
    temp_dir = tempfile.mkdtemp()
    try:
        print(f"\nExtracting: {pptx_path}")
        ole_paths = _extract_pptx(pptx_path, temp_dir)

        if not ole_paths:
            print("No embedded OLE objects found. Is this a PPTX with pasted ChemDraw structures?")
            return []

        print(f"Found {len(ole_paths)} embedded OLE object(s).\n")

        results, failed = [], []

        for i, ole_path in enumerate(ole_paths, 1):
            fname = ole_path.name
            print(f"[{i:02d}/{len(ole_paths)}] {fname}")

            # Step 1 — Extract CDX
            cdx_data = _extract_cdx_from_ole(ole_path)
            if cdx_data is None:
                failed.append({'file': fname, 'reason': 'CDX extraction failed'})
                continue

            # Sanity check: CDX magic bytes
            if cdx_data[:8] != b'VjCD0100':
                print(f"  [WARN] Unexpected magic bytes ({cdx_data[:8].hex()}); skipping.")
                failed.append({'file': fname, 'reason': 'Not a CDX file'})
                continue

            # Step 2 — CDX → CDXML
            tmp_cdx = os.path.join(temp_dir, f'tmp_{i}.cdx')
            cdxml_str = _cdx_to_cdxml(cdx_data, tmp_cdx)
            if cdxml_str is None:
                failed.append({'file': fname, 'reason': 'CDX → CDXML failed'})
                continue

            # Step 3 — Extract metadata
            meta = _extract_metadata(cdxml_str)

            # Step 4 — CDXML → SMILES
            smiles_list = _cdxml_to_smiles(cdxml_str)
            smiles = smiles_list[0] if smiles_list else None

            if not smiles:
                print(f"  [WARN] No SMILES generated.")
            if len(smiles_list) > 1:
                print(f"  [INFO] {len(smiles_list)} molecules found; using first.")

            compound_id = meta['compound_id'] or f'Unknown_{i}'
            if compound_id in ('N', ''):
                compound_id = f'Unknown_{i}'

            # Step 5 — Annotate stereo status
            stereo_note = _stereo_note(smiles)

            print(f"  ID:     {compound_id}")
            print(f"  SMILES: {smiles}")
            print(f"  Stereo: {stereo_note}")

            results.append({
                'index':       i,
                'file':        fname,
                'compound_id': compound_id,
                'smiles':      smiles,
                'stereo_note': stereo_note,
                'all_smiles':  smiles_list,
                'annotations': meta['annotations'],
                'text_labels': meta['text_labels'],
            })

        # ── Summary ──────────────────────────────────────────────────────────
        print(f"\n{'─'*60}")
        print(f"Processed:  {len(results)}/{len(ole_paths)}")
        print(f"Failed:     {len(failed)}")
        if failed:
            for f in failed:
                print(f"  ✗ {f['file']}: {f['reason']}")

        # ── Save JSON ─────────────────────────────────────────────────────────
        Path(output_json).parent.mkdir(parents=True, exist_ok=True)
        with open(output_json, 'w', encoding='utf-8') as fh:
            json.dump({'results': results, 'failed': failed}, fh, indent=2, ensure_ascii=False)
        print(f"\nResults saved → {output_json}")

        return results

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


# ── CLI entry point ───────────────────────────────────────────────────────────

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Extract ChemDraw CDX structures from a PPTX file and convert to SMILES.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('input_pptx', help='Path to the input PPTX file.')
    parser.add_argument(
        '--output', '-o',
        default='output/structures.json',
        help='Path to the output JSON file (default: output/structures.json).'
    )
    args = parser.parse_args()
    process_pptx(args.input_pptx, args.output)

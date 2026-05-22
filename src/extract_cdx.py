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


def _gap_sort_slide(objects: list[tuple[int, int, str]]) -> list[tuple[int, int, str]]:
    """
    Sort (y, x, ole_name) tuples into visual reading order using gap-based row detection.

    Algorithm
    ---------
    1. Sort all objects by y-coordinate.
    2. Identify row boundaries wherever the y-gap between consecutive objects
       exceeds a threshold (90 pt = ~3 cm). This is robust to small misalignments
       because real row gaps in a slide are always much larger than within-row
       y-variation.
    3. Within each detected row, sort left-to-right by x-coordinate.

    This approach works correctly regardless of whether structures are perfectly
    aligned on the slide grid.
    """
    if not objects:
        return []
    # 90 pt in EMU: 90 * 914400 / 72
    ROW_GAP_THRESHOLD = int(90 * 914400 / 72)
    by_y = sorted(objects, key=lambda t: t[0])
    rows: list[list] = []
    current_row = [by_y[0]]
    for i in range(1, len(by_y)):
        if by_y[i][0] - by_y[i - 1][0] > ROW_GAP_THRESHOLD:
            rows.append(current_row)
            current_row = [by_y[i]]
        else:
            current_row.append(by_y[i])
    rows.append(current_row)
    result = []
    for row in rows:
        result.extend(sorted(row, key=lambda t: t[1]))
    return result


def _build_slide_position_map(pptx_path: str, extract_dir: str) -> dict[str, tuple[int, int]]:
    """
    Parse all slide XMLs and return a mapping of OLE filename → (global_rank,).

    Structures are ranked in visual reading order: slide by slide, then
    top-to-bottom row (detected by y-gap), then left-to-right within each row.
    The rank is a single integer used as the sort key in _extract_pptx.

    Parameters
    ----------
    pptx_path : str
        Path to the PPTX file (already extracted to extract_dir).
    extract_dir : str
        Directory where the PPTX was extracted.

    Returns
    -------
    dict mapping ole_filename → integer rank (0-based, lower = earlier).
    """
    P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    rank_map: dict[str, int] = {}
    base = Path(extract_dir)
    global_rank = 0

    try:
        prs_root = ET.parse(str(base / 'ppt' / 'presentation.xml')).getroot()
        prs_rels = ET.parse(str(base / 'ppt' / '_rels' / 'presentation.xml.rels')).getroot()
        rid_to_target = {r.get('Id'): r.get('Target') for r in prs_rels}
        slide_rids = [sl.get(f'{{{R}}}id') for sl in prs_root.iter(f'{{{P}}}sldId')]

        for rid in slide_rids:
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

            # Collect (y, x, ole_name) for every OLE on this slide
            slide_objects: list[tuple[int, int, str]] = []
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
                slide_objects.append((y, x, ole_name))

            # Sort this slide's objects into visual reading order
            for _, _, ole_name in _gap_sort_slide(slide_objects):
                rank_map[ole_name] = global_rank
                global_rank += 1

    except Exception as e:
        print(f"  [WARN] Could not build slide position map: {e}")

    return rank_map


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

    # Build slide-position rank map (visual reading order)
    rank_map = _build_slide_position_map(pptx_path, extract_dir)

    def _sort_key(fname: str):
        # rank_map values are integers; unmapped files go to the end
        return rank_map.get(fname, 999999)

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
            if data[:8] != _CDX_MAGIC:
                print(f"  [WARN] Unexpected magic bytes ({data[:8].hex()}) in {ole_path.name}; skipping.")
                return None
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

        # ------------------------------------------------------------------ #
        # ChemDraw CDX objects are found in three structural layouts:
        #
        # Layout A — group-wrapped (older ChemDraw / copy-paste):
        #   <page>
        #     <group>                ← owns BOTH molecule and metadata
        #       <fragment/>          ← the molecule
        #       <annotation Keyword="CpdIndex" Content="CORRECT-ID"/>
        #       <t><s>CORRECT-ID</s></t>
        #     </group>
        #     <annotation Content="FOREIGN-ID"/>  ← belongs to neighbour
        #   </page>
        #   Fix: scope to the group that directly contains <fragment>.
        #
        # Layout B — group-separated (newer DEL exports, group has no fragment):
        #   <page>
        #     <fragment/>            ← molecule is a direct page child
        #     <group>                ← metadata group (no fragment inside)
        #       <annotation Keyword="CpdIndex" Content="CORRECT-ID"/>
        #       <t><s>CORRECT-ID</s></t>
        #     </group>
        #     <annotation Content="STALE-ID"/>  ← stale/wrong, ignore
        #   </page>
        #   Fix: use the non-fragment group; ignore page-level annotations.
        #
        # Layout C — flat (newer DEL exports, no metadata group):
        #   <page>
        #     <fragment/>            ← molecule is a direct page child
        #     <t><s>CORRECT-ID</s></t>  ← ID is a direct page-level <t>
        #     <annotation Content="STALE-ID"/>  ← stale/wrong, ignore
        #   </page>
        #   Fix: use the page-level <t> text; ignore page-level annotations.
        #
        # Detection order:
        #   1. If any <group> directly contains a <fragment>  → Layout A
        #   2. Else if any <group> (without fragment) exists  → Layout B
        #   3. Else                                           → Layout C
        # ------------------------------------------------------------------ #

        groups_with_mol = [g for g in page.findall('group')
                           if g.find('fragment') is not None]
        groups_without_mol = [g for g in page.findall('group')
                               if g.find('fragment') is None]

        if groups_with_mol:
            # Layout A: the group owns the molecule.
            # Metadata annotations are normally inside the group, but in some
            # CDX variants (Layout A2) the group has no annotations and all
            # metadata sits at page level instead.
            # Page-level annotations that belong to a *neighbour* compound are
            # a known issue in other layouts — but in Layout A2 the page-level
            # annotations are the ONLY source of metadata, so we must use them
            # when the group itself is empty.
            # IMPORTANT: text labels inside the group are atom labels (O, NH,
            # F, etc.) drawn on the structure — NOT the compound ID.
            # Use the CpdIndex annotation as the primary ID source.
            ann_scope = groups_with_mol[0]
            group_anns = list(ann_scope.iter('annotation'))
            if group_anns:
                # Layout A (classic): annotations are inside the group
                for ann in group_anns:
                    kw = ann.get('Keyword', '')
                    ct = ann.get('Content', '')
                    if kw and ct:
                        meta['annotations'][kw] = ct
            else:
                # Layout A2: group has no annotations.
                # Page-level annotations are STALE (copied from a neighbour) —
                # do NOT use them for ID or metadata.
                # The true compound ID is embedded as a <t> text element inside
                # the group, mixed with atom labels (O, N, S, etc.).
                # We identify it by looking for a text that matches the compound
                # ID pattern (hyphenated alphanumeric, e.g. PAC-DEL1726-525-3492).
                pass
            for t in ann_scope.iter('t'):
                for s in t.findall('s'):
                    if s.text and s.text.strip():
                        meta['text_labels'].append(s.text.strip())
            # ID: for Layout A (classic) prefer CpdIndex annotation;
            #     for Layout A2 the annotation dict is empty, so fall through
            #     to the text-label search below.
            for kw in _ID_ANNOTATION_KEYS:
                if kw in meta['annotations']:
                    meta['compound_id'] = meta['annotations'][kw]
                    break
            # Layout A2 fallback: find the compound-ID-like text label.
            if not meta['compound_id']:
                for text in meta['text_labels']:
                    first_line = text.split('\n')[0].strip()
                    if _ID_PAT.match(first_line) and len(first_line) > _MIN_COMPOUND_ID_LENGTH:
                        meta['compound_id'] = first_line
                        break

        elif groups_without_mol:
            # Layout B: metadata lives in a sibling group (no fragment).
            # The group's <t> text label is the compound ID.
            # Page-level annotations are stale — ignore them.
            ann_scope = groups_without_mol[0]
            for ann in ann_scope.iter('annotation'):
                kw = ann.get('Keyword', '')
                ct = ann.get('Content', '')
                if kw and ct:
                    meta['annotations'][kw] = ct
            for t in ann_scope.iter('t'):
                for s in t.findall('s'):
                    if s.text and s.text.strip():
                        meta['text_labels'].append(s.text.strip())
            # ID: prefer CpdIndex annotation, then first short text label
            for kw in _ID_ANNOTATION_KEYS:
                if kw in meta['annotations']:
                    meta['compound_id'] = meta['annotations'][kw]
                    break
            if not meta['compound_id']:
                for text in meta['text_labels']:
                    if '\n' not in text:
                        meta['compound_id'] = text
                        break

        else:
            # Layout C: flat — fragment is a direct page child, no groups.
            # The compound ID is in a direct page-level <t> text element.
            # Page-level annotations are STALE (copied from a neighbouring
            # compound in the source PPTX) — do NOT collect them as metadata.
            # Collecting them would fabricate SequenceCount/MW/CLogP values
            # for compounds that don't own those annotations.
            # Direct page-level <t> elements hold the compound ID
            for t in page.findall('t'):
                for s in t.findall('s'):
                    if s.text and s.text.strip():
                        meta['text_labels'].append(s.text.strip())
            # ID: first short single-line direct page text
            for text in meta['text_labels']:
                if '\n' not in text:
                    meta['compound_id'] = text
                    break

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
    has_dbl = '/' in smiles or '\\' in smiles
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


# ── Multi-entry CDX handling (R-group tables / multi-group schemes) ────────────

_ID_PAT = re.compile(r'^[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]')
_ID_ANNOTATION_KEYS = ('CpdIndex', 'CompoundIndex', 'HitIndex')
_MIN_COMPOUND_ID_LENGTH = 8
_CDX_MAGIC = b'VjCD0100'


def _is_multi_entry(text_labels: list[str]) -> bool:
    """Check whether text_labels contain multiple compound IDs (R-group table)."""
    count = sum(1 for t in text_labels
                if _ID_PAT.match(t.split('\n')[0].strip())
                and len(t.split('\n')[0].strip()) > _MIN_COMPOUND_ID_LENGTH)
    return count >= 2


def _expand_multi_entries(text_labels: list[str]) -> list[dict]:
    """Parse alternating [ID, props, ID, props, ...] text labels into per-variant dicts."""
    entries: list[dict] = []
    i = 0
    while i < len(text_labels):
        first_line = text_labels[i].split('\n')[0].strip()
        if _ID_PAT.match(first_line) and len(first_line) > _MIN_COMPOUND_ID_LENGTH:
            entry = {'compound_id': first_line, 'annotations': {}, 'text_labels': text_labels}
            if i + 1 < len(text_labels) and '\n' in text_labels[i + 1]:
                for line in text_labels[i + 1].split('\n'):
                    line = line.strip()
                    if ':' in line:
                        k, v = line.split(':', 1)
                        entry['annotations'][k.strip()] = v.strip()
            entries.append(entry)
            i += 2
        else:
            i += 1
    return entries


def _parse_scope(scope) -> dict | None:
    """Extract annotations, text labels, and compound ID from a CDXML scope element."""
    annotations = {}
    text_labels = []
    for ann in scope.iter('annotation'):
        kw = ann.get('Keyword', '')
        ct = ann.get('Content', '')
        if kw and ct:
            annotations[kw] = ct
    for t in scope.iter('t'):
        for s in t.findall('s'):
            if s.text and s.text.strip():
                text_labels.append(s.text.strip())
    cid = None
    for kw in _ID_ANNOTATION_KEYS:
        if kw in annotations:
            cid = annotations[kw]
            break
    if not cid:
        for text in text_labels:
            fl = text.split('\n')[0].strip()
            if _ID_PAT.match(fl) and len(fl) > _MIN_COMPOUND_ID_LENGTH:
                cid = fl
                break
    if not cid:
        for text in text_labels:
            if '\n' not in text:
                cid = text
                break
    if cid or annotations or text_labels:
        return {'compound_id': cid, 'annotations': annotations, 'text_labels': text_labels}
    return None


def _extract_all_cdxml_entries(cdxml_str: str) -> list[dict]:
    """Scan the entire CDXML for ALL compound entries across ALL groups.

    Unlike ``_extract_metadata`` which returns the first compound found, this
    function iterates every group (with or without fragments) and returns one
    dict per compound.  This catches multi-group CDX objects where each group
    is an individual R-group variant with its own HitIndex annotation.

    Falls back to text-label alternating-pattern parsing for R-group tables
    stored inside a single group.
    """
    entries: list[dict] = []
    try:
        root = ET.fromstring(cdxml_str)
        page = root.find('page')
        if page is None:
            page = root

        # ── Groups that contain a fragment (Layout A, possibly multi-group) ──
        groups_with_mol = [g for g in page.findall('group')
                           if g.find('fragment') is not None]
        for group in groups_with_mol:
            entry = _parse_scope(group)
            if entry:
                entries.append(entry)

        # ── Groups without a fragment (Layout B metadata groups) ────────────
        groups_without_mol = [g for g in page.findall('group')
                              if g.find('fragment') is None]
        if not groups_with_mol:
            for group in groups_without_mol:
                entry = _parse_scope(group)
                if entry:
                    entries.append(entry)
        else:
            # When groups-with-mol exist, also collect groups-without-mol that
            # carry distinct compound IDs not already seen (some schemes mix both).
            for group in groups_without_mol:
                entry = _parse_scope(group)
                if entry and entry.get('compound_id'):
                    existing = {e.get('compound_id') for e in entries}
                    if entry['compound_id'] not in existing:
                        entries.append(entry)

        # ── Fallback: page-level extraction (Layout C, no groups) ───────────
        if not entries:
            entry = _parse_scope(page)
            if entry:
                entries.append(entry)

        # ── Last resort: expand via alternating text-label pattern ──────────
        if len(entries) == 1 and entries[0].get('text_labels'):
            tl = entries[0]['text_labels']
            multi = _expand_multi_entries(tl) if _is_multi_entry(tl) else []
            if len(multi) > 1:
                return multi

    except Exception as e:
        print(f"  [WARN] Multi-entry extraction: {e}")

    return entries


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

            # Step 2 — CDX → CDXML
            tmp_cdx = os.path.join(temp_dir, f'tmp_{i}.cdx')
            cdxml_str = _cdx_to_cdxml(cdx_data, tmp_cdx)
            if cdxml_str is None:
                failed.append({'file': fname, 'reason': 'CDX → CDXML failed'})
                continue

            # Step 3 — CDXML → SMILES
            smiles_list = _cdxml_to_smiles(cdxml_str)
            smiles = smiles_list[0] if smiles_list else None

            if not smiles:
                print(f"  [WARN] No SMILES generated.")
            if len(smiles_list) > 1:
                print(f"  [INFO] {len(smiles_list)} molecules found; using first.")

            stereo_note = _stereo_note(smiles)

            # ── Build result row(s) ──────────────────────────────────────────
            # Check for multi-compound CDX: multiple groups or R-group tables
            all_entries = _extract_all_cdxml_entries(cdxml_str)

            if len(all_entries) > 1:
                print(f"  [MULTI] {len(all_entries)} compound variants in one CDX")
                print(f"  SMILES: {smiles}")
                print(f"  Stereo: {stereo_note}")
                for entry in all_entries:
                    results.append({
                        'index':       len(results) + 1,
                        'file':        fname,
                        'compound_id': entry['compound_id'] or f'Unknown_{len(results) + 1}',
                        'smiles':      smiles,
                        'stereo_note': stereo_note,
                        'all_smiles':  smiles_list,
                        'annotations': entry.get('annotations', {}),
                        'text_labels': entry.get('text_labels', []),
                    })
            else:
                # Single entry — original logic (via _extract_metadata for compat)
                meta = all_entries[0] if all_entries else _extract_metadata(cdxml_str)
                compound_id = meta.get('compound_id') or f'Unknown_{i}'
                if compound_id in ('N', ''):
                    compound_id = f'Unknown_{i}'

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

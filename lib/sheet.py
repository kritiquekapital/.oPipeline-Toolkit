#!/usr/bin/env python3
"""
lib/sheet.py

Builds the V1 Track Metadata Excel sheet from crates_flat.json (crate
export) + a DJ_ColorSheet_*.csv (color/genre export).

FIXES vs the original build_v1_sheet.py
----------------------------------------
1. COLOR FILL BUG: Serato's 20-swatch palette hex codes happen to use only
   digits 0-9 (never A-F). When pandas reads the color CSV without an
   explicit dtype, it silently infers those columns as numbers instead of
   text -- so `isinstance(hexval, str)` was always False and the Color Hex
   cell fill was NEVER applied, for any track, ever. Fixed by reading the
   whole CSV as strings (dtype=str), so "335599" stays "335599" instead of
   becoming 335599.0.

2. SILENT BUCKET MERGING: if two crate assignments for the same filename
   don't share a common top-level parent -- which usually means either a
   genuine duplicate filename (two different physical files, same name)
   or a track deliberately filed into two unrelated crates -- the old code
   just joined every parent name into one cell with '; ', producing a row
   that silently claims to belong to two different buckets at once. Now
   these get flagged as "AMBIGUOUS" in Parent Crate with the raw
   candidates preserved in Sub Crate, and logged to ambiguous_matches.csv,
   instead of being silently blended.
"""

from __future__ import annotations

import glob
import json
import os
import re
import sys
from pathlib import Path, PurePosixPath

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from .crates import fix_escape

DECORATIVE_PREFIX = 'Ƹ̵̡Ӝ̵̨̄Ʒ'


def strip_decorative_prefix(name: str) -> str:
    name = name.strip()
    if name.startswith(DECORATIVE_PREFIX):
        name = name[len(DECORATIVE_PREFIX):].lstrip()
    return name


def most_specific(paths):
    paths = list(paths)
    return [p for p in paths if not any(o != p and o.startswith(p + ' > ') for o in paths)]


def top_level(path: str) -> str:
    return path.split(' > ', 1)[0]


def split_parent_sub(paths):
    """
    Given the 'most specific' crate path(s) for one filename, decide:
      - clean case (all share one top-level parent): return (parent, subs, False)
      - ambiguous case (genuinely different top-level parents): return
        (raw candidates joined for visibility, '', True)
    """
    parents = {top_level(p) for p in paths}
    if len(parents) > 1:
        return 'AMBIGUOUS - see Sub Crate', '; '.join(sorted(paths)), True

    subs = []
    for p in paths:
        if ' > ' in p:
            _parent, sub = p.split(' > ', 1)
            subs.append(strip_decorative_prefix(sub))
    parent_name = strip_decorative_prefix(next(iter(parents)))
    return parent_name, '; '.join(sorted(set(subs))), False


def find_latest_colorsheet(folder: str) -> str:
    candidates = sorted(glob.glob(os.path.join(folder, 'DJ_ColorSheet_*.csv')))
    if not candidates:
        raise FileNotFoundError(f"No DJ_ColorSheet_*.csv found in {folder}")
    return candidates[-1]


def build(crates_flat_path: str, colorsheet_path: str, out_path: str, out_dir: Path | None = None) -> dict:
    with open(crates_flat_path, encoding='utf-8') as f:
        flat = json.load(f)

    # FIX 1: dtype=str everywhere. This is what actually keeps the color
    # hex columns as text instead of pandas silently inferring numbers.
    df = pd.read_csv(colorsheet_path, dtype=str, keep_default_na=False)

    crate_map = {}
    for row in flat:
        fname = PurePosixPath(row['track_path']).name.lower()
        crate_map.setdefault(fname, set()).add(fix_escape(row['crate_path']))

    df['_fname'] = df['FileName'].str.lower()
    df['_specific'] = df['_fname'].map(lambda f: most_specific(crate_map[f]) if f in crate_map else [])

    parents, subs, ambiguous_flags = [], [], []
    for specific in df['_specific']:
        if not specific:
            parents.append('')
            subs.append('')
            ambiguous_flags.append(False)
            continue
        p, s, amb = split_parent_sub(specific)
        parents.append(p)
        subs.append(s)
        ambiguous_flags.append(amb)

    df['ParentCrate'] = parents
    df['SubCrate'] = subs
    df['_ambiguous'] = ambiguous_flags

    cols = ['FileName', 'Artist', 'Title', 'ParentCrate', 'SubCrate', 'Genre', 'ColorName', 'ColorDisplayedHex',
            'BPM', 'InitialKey', 'Grouping', 'Comment', 'FullPath']
    missing = [c for c in cols if c not in df.columns]
    if missing:
        print(f"WARNING: color sheet is missing expected columns: {missing}", file=sys.stderr)
        for c in missing:
            df[c] = ''

    out_df = df[cols + ['_ambiguous']].rename(columns={'_ambiguous': 'Ambiguous'}).copy()
    out_df = out_df.sort_values(['ParentCrate', 'SubCrate', 'FileName']).reset_index(drop=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'V1 Track Metadata'

    headers = ['File Name', 'Artist', 'Title', 'Parent Crate', 'Sub Crate', 'Genre', 'Color',
               'Color Hex', 'BPM', 'Initial Key', 'Grouping', 'Comment', 'Full Path']
    ws.append(headers)

    header_font = Font(name='Arial', bold=True, color='FFFFFF')
    header_fill = PatternFill('solid', fgColor='4472C4')
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = header_font
        cell.fill = header_fill

    body_font = Font(name='Arial', size=10)
    ambiguous_fill = PatternFill('solid', fgColor='FFC7CE')  # light red, same idea as Excel's built-in "bad" style
    for _, r in out_df.iterrows():
        ws.append([r['FileName'], r['Artist'], r['Title'], r['ParentCrate'], r['SubCrate'], r['Genre'],
                   r['ColorName'], r['ColorDisplayedHex'], r['BPM'], r['InitialKey'],
                   r['Grouping'], r['Comment'], r['FullPath']])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.font = body_font

    for i, r in enumerate(out_df.itertuples(), start=2):
        hexval = getattr(r, 'ColorDisplayedHex')
        if isinstance(hexval, str) and re.fullmatch(r'[0-9A-Fa-f]{6}', hexval):
            ws.cell(row=i, column=8).fill = PatternFill('solid', fgColor=hexval)
        if getattr(r, 'Ambiguous'):
            for col in (4, 5):  # Parent Crate, Sub Crate
                ws.cell(row=i, column=col).fill = ambiguous_fill

    widths = [34, 20, 30, 24, 26, 22, 12, 10, 8, 12, 16, 30, 55]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    # Every column is read as text (dtype=str) with keep_default_na=False, so
    # "missing" is always an empty string here -- never rely on dtype checks
    # or .notna(), which behave inconsistently across pandas versions for
    # string-typed columns (a bug that bit an earlier version of this).
    matched = (out_df['ParentCrate'] != '').sum()
    genre_set = (out_df['Genre'] != '').sum()
    colored = (out_df['ColorName'] != '(not tagged)').sum()
    ambiguous_count = int(out_df['Ambiguous'].sum())

    notes = wb.create_sheet('Notes')
    notes['A1'] = 'Notes on this export'
    notes['A1'].font = Font(name='Arial', bold=True, size=12)
    lines = [
        '',
        f'Source crate file:  {crates_flat_path}',
        f'Source color sheet: {colorsheet_path}',
        f'Built: {pd.Timestamp.now().isoformat(timespec="seconds")}',
        '',
        f'Rows: {len(out_df)} tracks.',
        f'Matched to at least one crate: {matched} of {len(out_df)} (matched by filename, not full path).',
        f'Genre set: {genre_set} of {len(out_df)}.',
        f'Color set: {colored} of {len(out_df)}.',
        f'Ambiguous (filename matches crates under more than one top-level bucket): {ambiguous_count}. '
        'These are highlighted in red and NOT guessed at -- check Sub Crate for the raw candidates and '
        'resolve by hand (usually a duplicate filename, or a track filed in two unrelated crates).',
        '',
        'Color name caveat: Serato stores no text names for its color swatches -- names here are matched to the '
        'documented anchor palette by hue order. Spot-check against Serato\'s own color picker periodically.',
    ]
    for i, line in enumerate(lines, start=2):
        notes.cell(row=i, column=1, value=line).font = Font(name='Arial', size=10)
        notes.cell(row=i, column=1).alignment = Alignment(wrap_text=True, vertical='top')
    notes.column_dimensions['A'].width = 110

    try:
        wb.save(out_path)
    except PermissionError:
        raise PermissionError(
            f"Could not save {out_path} -- it's probably open in Excel right now. Close it and try again.")

    if ambiguous_count and out_dir is not None:
        amb_path = Path(out_dir) / 'ambiguous_matches.csv'
        out_df[out_df['Ambiguous']][['FileName', 'FullPath', 'SubCrate']].rename(
            columns={'SubCrate': 'CandidateCrates'}).to_csv(amb_path, index=False)

    return {
        'out_path': out_path,
        'rows': len(out_df),
        'matched': int(matched),
        'genre_set': int(genre_set),
        'colored': int(colored),
        'ambiguous': ambiguous_count,
    }


def run(crates_flat_path: str, colorsheet_dir_or_path: str, out_path: str, out_dir: Path, is_dir: bool = True) -> dict:
    colorsheet_path = find_latest_colorsheet(colorsheet_dir_or_path) if is_dir else colorsheet_dir_or_path
    result = build(crates_flat_path, colorsheet_path, out_path, out_dir=out_dir)
    result['colorsheet_used'] = colorsheet_path
    return result


def main():
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--crates-flat', required=True)
    ap.add_argument('--colorsheet')
    ap.add_argument('--colorsheet-dir')
    ap.add_argument('--out', default='V1_Track_Metadata.xlsx')
    args = ap.parse_args()

    if not args.colorsheet and not args.colorsheet_dir:
        print("Provide either --colorsheet <file> or --colorsheet-dir <folder>", file=sys.stderr)
        sys.exit(1)

    colorsheet_path = args.colorsheet or find_latest_colorsheet(args.colorsheet_dir)
    print(f"Using color sheet: {colorsheet_path}")
    try:
        result = build(args.crates_flat, colorsheet_path, args.out, out_dir=Path(args.out).parent)
    except PermissionError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Saved {result['out_path']}")
    print(f"  {result['rows']} rows, {result['matched']} matched, {result['genre_set']} with Genre, "
          f"{result['colored']} with Color, {result['ambiguous']} ambiguous.")


if __name__ == '__main__':
    main()

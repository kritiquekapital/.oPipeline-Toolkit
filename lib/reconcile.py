#!/usr/bin/env python3
"""
lib/reconcile.py

Merges a freshly-built V1_Track_Metadata.xlsx (from sheet.py) into the
real working V1.xlsx master -- the file you actually use day to day,
with its own Library/Notes/Colors sheets, search-row helper, and a
live formula in "Current Bucket (Sub)" that parses the Sub name out of
"Current Bucket (Parent)".

WHY THIS EXISTS
    Refresh only ever produces a disposable snapshot in output\\. There
    was previously no safe way to pull that into the real master without
    doing it by hand in Excel, or via the old VBA macro (COM automation,
    fragile path-matching, gone along with the other risky pieces).

WHAT IT DOES AND DOESN'T TOUCH
    - Matches rows by File Name (same convention as everywhere else in
      this pipeline).
    - For an EXISTING row: overwrites Artist/Title/Bucket/Genre/Color/
      BPM/Key/Grouping/Comment/Path/Hex with the fresh values. These are
      all Serato- or tag-derived fields, not things you type into Excel
      by hand, so overwriting them is safe given your workflow (you edit
      in Serato, not in the spreadsheet).
    - The "Current Bucket (Sub)" column is a LIVE FORMULA that parses the
      Sub name out of "Current Bucket (Parent)" -- it is never written to
      directly, only recalculated by Excel once Parent changes.
    - A track flagged AMBIGUOUS by sheet.py is never written into Bucket
      here -- the existing value (if any) is left alone and it's counted
      separately so you can resolve it by hand.
    - A NEW filename (in the fresh export, not yet in the master) gets
      APPENDED as a new row at the bottom, copying the style of the last
      existing row so it doesn't look out of place.
    - A filename that's in the master but NOT in this fresh export is
      NEVER touched or deleted -- just counted and listed, since that
      could mean the file moved outside V1 root, got renamed, or simply
      wasn't scanned this run.
    - The Notes and Colors sheets are never modified.
    - This ALWAYS writes to a new file, never overwrites your real
      V1.xlsx directly -- the caller decides whether/when to replace it,
      after reviewing the summary.
"""

from __future__ import annotations

import copy
import re
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill

_HEX_RE = re.compile(r'[0-9A-Fa-f]{6}')

# Column numbers (1-indexed) in the master Library sheet.
M_FILENAME, M_ARTIST, M_TITLE, M_BUCKET_PARENT, M_BUCKET_SUB = 1, 2, 3, 4, 5
M_GENRE, M_COLOR, M_BPM, M_KEY, M_GROUPING, M_COMMENT, M_PATH, M_HEX = 6, 7, 8, 9, 10, 11, 12, 13

# Column numbers (1-indexed) in the fresh V1 Track Metadata sheet (sheet.py's output).
F_FILENAME, F_ARTIST, F_TITLE, F_PARENT, F_SUB = 1, 2, 3, 4, 5
F_GENRE, F_COLOR, F_HEX, F_BPM, F_KEY, F_GROUPING, F_COMMENT, F_PATH = 6, 7, 8, 9, 10, 11, 12, 13


def _read_fresh_rows(fresh_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(fresh_path, data_only=True)
    ws = wb['V1 Track Metadata']
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[F_FILENAME - 1]:
            continue
        rows.append({
            'FileName': r[F_FILENAME - 1], 'Artist': r[F_ARTIST - 1], 'Title': r[F_TITLE - 1],
            'ParentCrate': r[F_PARENT - 1] or '', 'SubCrate': r[F_SUB - 1] or '',
            'Genre': r[F_GENRE - 1], 'Color': r[F_COLOR - 1], 'Hex': r[F_HEX - 1],
            'BPM': r[F_BPM - 1], 'Key': r[F_KEY - 1], 'Grouping': r[F_GROUPING - 1],
            'Comment': r[F_COMMENT - 1], 'Path': r[F_PATH - 1],
        })
    return rows


def _apply_hex_fill(ws, row: int, hexval):
    cell = ws.cell(row=row, column=M_HEX)
    if isinstance(hexval, str) and _HEX_RE.fullmatch(hexval):
        cell.fill = PatternFill('solid', fgColor=hexval)


def _sub_formula(row: int) -> str:
    return f'=TRIM(IFERROR(MID(D{row}, FIND(" > ", D{row})+3, 255), ""))'


def merge(master_path: str, fresh_path: str, out_path: str) -> dict:
    fresh_rows = _read_fresh_rows(fresh_path)

    wb = openpyxl.load_workbook(master_path, data_only=False)
    if 'Library' not in wb.sheetnames:
        raise ValueError(f"No 'Library' sheet found in {master_path}")
    ws = wb['Library']

    # filename (lowercase) -> row number, for every existing data row.
    existing = {}
    for r in range(2, ws.max_row + 1):
        fname = ws.cell(row=r, column=M_FILENAME).value
        if fname:
            existing[str(fname).lower()] = r

    seen_master_rows = set()
    updated = 0
    appended = 0
    ambiguous_skipped = []

    last_row = ws.max_row
    # Grab a style template from the last existing data row so new rows match.
    style_template_row = last_row if last_row >= 2 and ws.cell(row=last_row, column=1).value else 3

    for row in fresh_rows:
        key = str(row['FileName']).lower()
        is_ambiguous = str(row['ParentCrate']).startswith('AMBIGUOUS')
        bucket_combined = None
        if not is_ambiguous and row['ParentCrate']:
            bucket_combined = row['ParentCrate'] if not row['SubCrate'] else f"{row['ParentCrate']} > {row['SubCrate']}"

        if key in existing:
            r = existing[key]
            seen_master_rows.add(r)
            ws.cell(row=r, column=M_ARTIST, value=row['Artist'])
            ws.cell(row=r, column=M_TITLE, value=row['Title'])
            if is_ambiguous:
                ambiguous_skipped.append(row['FileName'])
            elif bucket_combined is not None:
                ws.cell(row=r, column=M_BUCKET_PARENT, value=bucket_combined)
            ws.cell(row=r, column=M_GENRE, value=row['Genre'])
            ws.cell(row=r, column=M_COLOR, value=row['Color'])
            ws.cell(row=r, column=M_BPM, value=row['BPM'])
            ws.cell(row=r, column=M_KEY, value=row['Key'])
            ws.cell(row=r, column=M_GROUPING, value=row['Grouping'])
            ws.cell(row=r, column=M_COMMENT, value=row['Comment'])
            ws.cell(row=r, column=M_PATH, value=row['Path'])
            ws.cell(row=r, column=M_HEX, value=row['Hex'])
            _apply_hex_fill(ws, r, row['Hex'])
            updated += 1
        else:
            r = ws.max_row + 1
            for col in range(1, 14):
                template_cell = ws.cell(row=style_template_row, column=col)
                new_cell = ws.cell(row=r, column=col)
                new_cell.font = copy.copy(template_cell.font)
                new_cell.alignment = copy.copy(template_cell.alignment)
            ws.cell(row=r, column=M_FILENAME, value=row['FileName'])
            ws.cell(row=r, column=M_ARTIST, value=row['Artist'])
            ws.cell(row=r, column=M_TITLE, value=row['Title'])
            if is_ambiguous:
                ambiguous_skipped.append(row['FileName'])
            elif bucket_combined is not None:
                ws.cell(row=r, column=M_BUCKET_PARENT, value=bucket_combined)
            ws.cell(row=r, column=M_BUCKET_SUB, value=_sub_formula(r))
            ws.cell(row=r, column=M_GENRE, value=row['Genre'])
            ws.cell(row=r, column=M_COLOR, value=row['Color'])
            ws.cell(row=r, column=M_BPM, value=row['BPM'])
            ws.cell(row=r, column=M_KEY, value=row['Key'])
            ws.cell(row=r, column=M_GROUPING, value=row['Grouping'])
            ws.cell(row=r, column=M_COMMENT, value=row['Comment'])
            ws.cell(row=r, column=M_PATH, value=row['Path'])
            ws.cell(row=r, column=M_HEX, value=row['Hex'])
            _apply_hex_fill(ws, r, row['Hex'])
            appended += 1

    not_in_refresh = [
        ws.cell(row=r, column=M_FILENAME).value
        for r in existing.values() if r not in seen_master_rows
    ]

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    return {
        'out_path': out_path,
        'updated': updated,
        'appended': appended,
        'ambiguous_skipped': ambiguous_skipped,
        'not_in_refresh': not_in_refresh,
    }

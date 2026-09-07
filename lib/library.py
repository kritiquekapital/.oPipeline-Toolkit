#!/usr/bin/env python3
"""
lib/library.py

Read-only helpers over the already-built V1_Track_Metadata.xlsx --
used by the "Library health check" and "Search library" menu options.
Neither of these rescans Serato or touches any files; they just report
on whatever the last Refresh produced.
"""

from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

import openpyxl

COLS = ['FileName', 'Artist', 'Title', 'ParentCrate', 'SubCrate', 'Genre', 'ColorName',
        'ColorHex', 'BPM', 'InitialKey', 'Grouping', 'Comment', 'FullPath']


def load_rows(sheet_path: str) -> list[dict]:
    wb = openpyxl.load_workbook(sheet_path, data_only=True)
    ws = wb['V1 Track Metadata']
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        rows.append(dict(zip(COLS, r)))
    return rows


def health_summary(sheet_path: str) -> dict:
    rows = load_rows(sheet_path)
    total = len(rows)

    bucket_counts = Counter((r['ParentCrate'] or '(unmatched)') for r in rows)
    genre_counts = Counter((r['Genre'] or '(none)') for r in rows)
    color_counts = Counter((r['ColorName'] or '(not tagged)') for r in rows)

    unmatched = sum(1 for r in rows if not r['ParentCrate'])
    ambiguous = sum(1 for r in rows if (r['ParentCrate'] or '').startswith('AMBIGUOUS'))
    no_genre = sum(1 for r in rows if not r['Genre'])
    no_color = sum(1 for r in rows if (r['ColorName'] or '(not tagged)') in ('(not tagged)', ''))
    no_bpm = sum(1 for r in rows if not r['BPM'])
    no_key = sum(1 for r in rows if not r['InitialKey'])

    # Per-bucket coverage: a bucket with plenty of tracks but almost none
    # colored/genred is worth flagging on its own, not just buried in the
    # library-wide totals -- this is what your own manual audits have been
    # catching by hand (e.g. one whole bucket sitting untagged while others
    # are fully worked).
    bucket_total = Counter()
    bucket_genre = Counter()
    bucket_color = Counter()
    for r in rows:
        b = r['ParentCrate'] or '(unmatched)'
        if b.startswith('AMBIGUOUS'):
            continue
        bucket_total[b] += 1
        if r['Genre']:
            bucket_genre[b] += 1
        if (r['ColorName'] or '(not tagged)') not in ('(not tagged)', ''):
            bucket_color[b] += 1

    bucket_coverage = []
    for b, n in bucket_total.items():
        genre_pct = bucket_genre[b] / n * 100
        color_pct = bucket_color[b] / n * 100
        bucket_coverage.append((b, n, genre_pct, color_pct))
    bucket_coverage.sort(key=lambda x: x[3])  # lowest color coverage first

    return {
        'total': total,
        'unmatched': unmatched,
        'ambiguous': ambiguous,
        'no_genre': no_genre,
        'no_color': no_color,
        'no_bpm': no_bpm,
        'no_key': no_key,
        'buckets': bucket_counts.most_common(),
        'genres': genre_counts.most_common(),
        'colors': color_counts.most_common(),
        'bucket_coverage': bucket_coverage,
    }


def search(sheet_path: str, term: str, limit: int = 25) -> list[dict]:
    term_lower = term.lower()
    rows = load_rows(sheet_path)
    hits = []
    for r in rows:
        haystack = ' '.join(str(r.get(k) or '') for k in ('FileName', 'Artist', 'Title')).lower()
        if term_lower in haystack:
            hits.append(r)
            if len(hits) >= limit:
                break
    return hits


def _normalize(s) -> str:
    s = (s or '').lower()
    # \w is Unicode-aware by default in Python 3 -- do NOT restrict to
    # [a-z0-9], which silently destroys Cyrillic/CJK/etc text entirely and
    # collapses every non-Latin title down to an empty string, falsely
    # grouping unrelated tracks together under the same (artist, '') key.
    s = re.sub(r'[^\w\s]', '', s, flags=re.UNICODE)
    return re.sub(r'\s+', ' ', s).strip()


def find_duplicates(sheet_path: str) -> dict:
    """
    Read-only duplicate report over the last built sheet. Two independent
    checks, since they catch different things:
      - same FILENAME appearing more than once (often a double-import, or
        the same physical file matching more than one crate)
      - same normalized Artist+Title appearing under different filenames
        (often a re-download or re-encode of the same track)
    Never deletes or merges anything -- this only ever reports.
    """
    rows = load_rows(sheet_path)

    by_filename: dict[str, list[dict]] = {}
    by_artist_title: dict[tuple[str, str], list[dict]] = {}
    for r in rows:
        by_filename.setdefault(str(r['FileName']).lower(), []).append(r)
        key = (_normalize(r['Artist']), _normalize(r['Title']))
        if key != ('', ''):
            by_artist_title.setdefault(key, []).append(r)

    filename_dupes = [group for group in by_filename.values() if len(group) > 1]
    # Only report artist+title groups where the filenames actually differ --
    # an exact filename match is already covered above, no need to double-list it.
    artist_title_dupes = [
        group for group in by_artist_title.values()
        if len(group) > 1 and len({str(r['FileName']).lower() for r in group}) > 1
    ]

    return {'filename_dupes': filename_dupes, 'artist_title_dupes': artist_title_dupes}

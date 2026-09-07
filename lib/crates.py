#!/usr/bin/env python3
"""
lib/crates.py

Parses Serato .crate files (binary format in _Serato_/Subcrates/*.crate)
into a clean nested tree + flat list. Read-only -- never touches audio
files or the crate database itself.

This is the same parsing logic that was already working well; it's just
wrapped here as an importable function so the orchestrator can call it
directly instead of shelling out to a separate script.

CRATE NAME FORMAT
    Serato encodes crate hierarchy in the *filename*, not folder nesting:
        Hard %7c Raw %7c Bass.crate                      -> top-level crate
        Hard %7c Raw %7c Bass%%RuBass - Hardstyle.crate   -> subcrate of it
    "%%" separates hierarchy levels. This library's crates use a custom
    escape scheme for special characters instead of standard %XX encoding
    (since %% is already reserved): two literal U+241B ("SYMBOL FOR
    ESCAPE", displays as ␛) characters followed by the hex code, e.g.
    ␛␛7c -> '|'. Standard %XX decoding is also tried as a fallback.
"""

from __future__ import annotations

import json
import re
import struct
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote

_ESCAPE_RE = re.compile('\u241b\u241b([0-9a-fA-F]{2})')


def fix_escape(s: str) -> str:
    return _ESCAPE_RE.sub(lambda m: chr(int(m.group(1), 16)), s)


def parse_crate(path: Path):
    """Return (version_string, [track_path, ...]) for one .crate file."""
    data = path.read_bytes()

    if data[0:4] != b'vrsn':
        raise ValueError(f"{path.name}: doesn't start with 'vrsn' tag -- "
                          f"not a valid/unmangled Serato crate file")

    otrk_idx = data.find(b'otrk', 4)
    if otrk_idx == -1:
        vrsn_text = data[4:].decode('utf-16-be', errors='replace')
        return vrsn_text, []

    vrsn_text = data[4:otrk_idx].decode('utf-16-be', errors='replace')

    pos = otrk_idx
    tracks = []
    while pos < len(data):
        tag = data[pos:pos + 4]
        if len(data) - pos < 8:
            break
        length = struct.unpack('>I', data[pos + 4:pos + 8])[0]
        payload = data[pos + 8:pos + 8 + length]
        pos += 8 + length

        if tag == b'otrk':
            p = 0
            while p < len(payload) - 8:
                subtag = payload[p:p + 4]
                sublen = struct.unpack('>I', payload[p + 4:p + 8])[0]
                subpayload = payload[p + 8:p + 8 + sublen]
                if subtag == b'ptrk':
                    tracks.append(subpayload.decode('utf-16-be', errors='replace'))
                p += 8 + sublen

    return vrsn_text, tracks


def decode_crate_name(filename: str):
    """'Hard ␛␛7c Raw ␛␛7c Bass%%RuBass - Hardstyle.crate' -> ['Hard | Raw | Bass', 'RuBass - Hardstyle']"""
    stem = filename[:-len('.crate')] if filename.lower().endswith('.crate') else filename
    levels = stem.split('%%')

    def unescape(level: str) -> str:
        decoded = fix_escape(level)
        return unquote(decoded)

    return [unescape(level) for level in levels]


def build_tree(crate_dir: Path, prefix: str | None = None, progress_cb=None):
    """
    Returns (tree, flat_rows, errors, diagnostics). See module docstring
    for the tree/flat_rows shape -- unchanged from the original script.

    progress_cb, if given, is called as progress_cb(current, total) after
    each crate file is processed -- purely a UI hook, no effect on output.
    """
    tree = {}
    flat_rows = []
    errors = []
    skipped_by_prefix = []
    included_files = []

    all_crate_files = sorted(crate_dir.glob('*.crate'))
    total = len(all_crate_files)

    for i, crate_file in enumerate(all_crate_files, start=1):
        if prefix and not crate_file.name.startswith(prefix):
            skipped_by_prefix.append(crate_file.name)
            if progress_cb:
                progress_cb(i, total)
            continue

        try:
            _vrsn, tracks = parse_crate(crate_file)
        except Exception as e:
            errors.append({"file": crate_file.name, "error": str(e)})
            if progress_cb:
                progress_cb(i, total)
            continue

        included_files.append(crate_file.name)
        levels = decode_crate_name(crate_file.name)
        if prefix and levels and levels[0] == prefix:
            levels = levels[1:]
        if not levels:
            levels = [prefix or crate_file.stem]

        node = tree
        for i2, level_name in enumerate(levels):
            if level_name not in node:
                node[level_name] = {"tracks": [], "subcrates": {}}
            if i2 == len(levels) - 1:
                node[level_name]["tracks"] = tracks
            node = node[level_name]["subcrates"]

        crate_path_str = " > ".join(levels)
        for t in tracks:
            flat_rows.append({"crate_path": crate_path_str, "track_path": t})

        if progress_cb:
            progress_cb(i, total)

    diagnostics = {
        "crate_dir": str(crate_dir),
        "prefix_filter": prefix,
        "total_crate_files_found": len(all_crate_files),
        "files_included": len(included_files),
        "files_skipped_by_prefix": len(skipped_by_prefix),
        "files_failed_to_parse": len(errors),
        "included_file_names": included_files,
        "skipped_file_names": skipped_by_prefix,
        "run_at": datetime.now(timezone.utc).isoformat(),
    }

    return tree, flat_rows, errors, diagnostics


def run(crate_dir: str, prefix: str | None, out_dir: Path, progress_cb=None) -> dict:
    """
    Orchestrator entry point. Writes crates.json / crates_flat.json /
    crates_diagnostics.json (and crates_errors.json if any) into out_dir.
    Returns the diagnostics dict plus a 'flat_path' key for downstream steps.
    """
    crate_dir_p = Path(crate_dir)
    if not crate_dir_p.is_dir():
        raise FileNotFoundError(f"Not a directory: {crate_dir_p}")

    tree, flat_rows, errors, diagnostics = build_tree(crate_dir_p, prefix=prefix, progress_cb=progress_cb)

    out_dir.mkdir(parents=True, exist_ok=True)
    tree_path = out_dir / 'crates.json'
    flat_path = out_dir / 'crates_flat.json'
    diag_path = out_dir / 'crates_diagnostics.json'

    tree_path.write_text(json.dumps(tree, indent=2, ensure_ascii=False), encoding='utf-8')
    flat_path.write_text(json.dumps(flat_rows, indent=2, ensure_ascii=False), encoding='utf-8')
    diag_path.write_text(json.dumps(diagnostics, indent=2, ensure_ascii=False), encoding='utf-8')

    if errors:
        (out_dir / 'crates_errors.json').write_text(
            json.dumps(errors, indent=2, ensure_ascii=False), encoding='utf-8')

    diagnostics['flat_path'] = str(flat_path)
    diagnostics['tree_path'] = str(tree_path)
    diagnostics['track_count'] = len(flat_rows)
    return diagnostics


def main():
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--crate-dir', required=True)
    ap.add_argument('--out-dir', default='output')
    ap.add_argument('--prefix', default=None)
    args = ap.parse_args()

    diag = run(args.crate_dir, args.prefix, Path(args.out_dir))
    print(f"Scanned:  {diag['total_crate_files_found']} .crate file(s) in {diag['crate_dir']}")
    print(f"Included: {diag['files_included']}   Skipped (prefix): {diag['files_skipped_by_prefix']}   "
          f"Failed: {diag['files_failed_to_parse']}")
    print(f"Wrote: {diag['tree_path']}, {diag['flat_path']}  ({diag['track_count']} track entries)")
    if diag['files_failed_to_parse']:
        print(f"{diag['files_failed_to_parse']} file(s) failed to parse -- see crates_errors.json", file=sys.stderr)


if __name__ == '__main__':
    main()

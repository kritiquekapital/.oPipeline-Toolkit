#!/usr/bin/env python3
"""
lib/mover.py

Physically moves files inside the V1 folder into subfolders matching their
Serato crate/subcrate. Destructive -- always dry-run (writes move_plan.csv
only) unless apply=True is passed.

FIX vs the original move_v1_files.py
-------------------------------------
The old ambiguity check only looked at whether a filename's crate matches
spanned more than one TOP-LEVEL parent. But a track can legitimately be
filed into two different SUBcrates under the *same* parent (e.g. two
unrelated subcrates of "Hard | Raw | Bass") -- that also produces more
than one "most specific" path, and the old code just silently picked one
alphabetically with no log entry at all. Now ANY filename resolving to
more than one most-specific crate path is treated as ambiguous and logged,
regardless of whether the parents match, so nothing gets silently guessed.
"""

from __future__ import annotations

import csv
import json
import re
import shutil
import sys
from pathlib import Path

from .crates import fix_escape

ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*]')
FOLDER_PREFIX = 'Ƹ̵̡Ӝ̵̨̄Ʒ '


def sanitize_folder_name(name: str) -> str:
    name = name.replace('|', '-')
    name = ILLEGAL_CHARS.sub('_', name)
    name = name.strip().rstrip('.')
    return name or '_unnamed'


def most_specific(paths):
    paths = list(paths)
    return [p for p in paths if not any(o != p and o.startswith(p + ' > ') for o in paths)]


def split_parent_sub(path: str):
    if ' > ' in path:
        parent, sub = path.split(' > ', 1)
    else:
        parent, sub = path, ''
    return parent, sub


def plan_moves(v1_root: Path, crates_flat_path: str, extensions: list[str], progress_cb=None):
    with open(crates_flat_path, encoding='utf-8') as f:
        flat = json.load(f)

    ext_lower = {e.lower() for e in extensions}

    crate_map = {}
    for row in flat:
        fname = Path(row['track_path']).name.lower()
        crate_map.setdefault(fname, set()).add(fix_escape(row['crate_path']))

    files = [p for p in v1_root.rglob('*') if p.is_file() and p.suffix.lower() in ext_lower]
    total = len(files)

    move_plan = []
    unmatched = []
    ambiguous = []
    skipped_conflict = []
    already_in_place = 0

    for i, f in enumerate(files, start=1):
        key = f.name.lower()
        paths = crate_map.get(key)
        if not paths:
            unmatched.append(str(f))
            if progress_cb:
                progress_cb(i, total)
            continue

        specific = most_specific(paths)

        # FIX: flag ANY multi-candidate match as ambiguous, not just ones
        # that span different top-level parents.
        if len(specific) > 1:
            ambiguous.append(f"{f}  ->  {'; '.join(sorted(specific))}")
            if progress_cb:
                progress_cb(i, total)
            continue

        chosen = specific[0]
        parent, sub = split_parent_sub(chosen)
        parent_dir = FOLDER_PREFIX + sanitize_folder_name(parent)
        target_dir = v1_root / parent_dir
        if sub:
            target_dir = target_dir / sanitize_folder_name(sub)

        target_path = target_dir / f.name

        if progress_cb:
            progress_cb(i, total)

        if target_path.resolve() == f.resolve():
            already_in_place += 1
            continue

        if target_path.exists():
            skipped_conflict.append(f"{f}  ->  {target_path}  (target already exists, different file)")
            continue

        move_plan.append((f, target_path))

    return {
        'files_scanned': len(files),
        'already_in_place': already_in_place,
        'move_plan': move_plan,
        'unmatched': unmatched,
        'ambiguous': ambiguous,
        'skipped_conflict': skipped_conflict,
    }


def write_reports(plan: dict, output_dir: Path):
    output_dir.mkdir(parents=True, exist_ok=True)
    plan_path = output_dir / 'move_plan.csv'
    with open(plan_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['source', 'target'])
        for src, dst in plan['move_plan']:
            writer.writerow([str(src), str(dst)])

    written = {'move_plan.csv': plan_path}
    for name, rows in [('unmatched.txt', plan['unmatched']), ('ambiguous.txt', plan['ambiguous']),
                        ('skipped_conflict.txt', plan['skipped_conflict'])]:
        if rows:
            out = output_dir / name
            out.write_text('\n'.join(rows), encoding='utf-8')
            written[name] = out
    return written


def apply_moves(plan: dict, output_dir: Path):
    moved = 0
    errors = []
    done_pairs = []
    for src, dst in plan['move_plan']:
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
            moved += 1
            done_pairs.append((str(src), str(dst)))
        except Exception as e:
            errors.append(f"{src} -> {dst}: {e}")

    if errors:
        (output_dir / 'move_errors.txt').write_text('\n'.join(errors), encoding='utf-8')

    # Log of what actually succeeded, for Undo -- distinct from move_plan.csv,
    # which is just the PROPOSED plan and may include moves that failed.
    if done_pairs:
        with open(output_dir / 'last_move_log.csv', 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['source', 'target'])
            writer.writerows(done_pairs)

    return moved, errors


def undo_last_move(output_dir: Path) -> dict:
    """
    Reverses everything in last_move_log.csv (the log of actually-completed
    moves from the most recent --apply run). Moves each file back from its
    target to its original source. Never overwrites -- if the source path
    is occupied again (e.g. a new file with the same name landed there
    since), that entry is reported as un-doable rather than clobbered.
    """
    log_path = output_dir / 'last_move_log.csv'
    if not log_path.exists():
        return {'available': False}

    with open(log_path, encoding='utf-8') as f:
        pairs = list(csv.DictReader(f))

    reverted = 0
    failed = []
    for row in pairs:
        src = Path(row['source'])  # original location, before the move
        dst = Path(row['target'])  # where it currently is
        if not dst.exists():
            failed.append(f"{dst}: not found (already moved or renamed again since)")
            continue
        if src.exists():
            failed.append(f"{dst}: can't restore -- {src} is occupied again")
            continue
        try:
            src.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(dst), str(src))
            reverted += 1
        except Exception as e:
            failed.append(f"{dst} -> {src}: {e}")

    # Once undone, rename the log so a second Undo doesn't try to replay it.
    if reverted:
        log_path.rename(output_dir / 'last_move_log.undone.csv')

    return {'available': True, 'total': len(pairs), 'reverted': reverted, 'failed': failed}


def run(v1_root: str, crates_flat_path: str, output_dir: Path, extensions=None, apply: bool = False,
        progress_cb=None) -> dict:
    v1_root_p = Path(v1_root)
    if not v1_root_p.is_dir():
        raise FileNotFoundError(f"Not a directory: {v1_root_p}")

    plan = plan_moves(v1_root_p, crates_flat_path, extensions or ['.mp3'], progress_cb=progress_cb)
    written = write_reports(plan, output_dir)

    result = {
        'files_scanned': plan['files_scanned'],
        'already_in_place': plan['already_in_place'],
        'planned': len(plan['move_plan']),
        'unmatched': len(plan['unmatched']),
        'ambiguous': len(plan['ambiguous']),
        'skipped_conflict': len(plan['skipped_conflict']),
        'reports': {k: str(v) for k, v in written.items()},
        'applied': False,
        'moved': 0,
        'errors': 0,
    }

    if apply:
        moved, errors = apply_moves(plan, output_dir)
        result['applied'] = True
        result['moved'] = moved
        result['errors'] = len(errors)

    return result


def main():
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--v1-root', required=True)
    ap.add_argument('--crates-flat', required=True)
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--extensions', nargs='+', default=['.mp3'])
    ap.add_argument('--output-dir')
    args = ap.parse_args()

    out_dir = Path(args.output_dir) if args.output_dir else Path(args.v1_root)
    result = run(args.v1_root, args.crates_flat, out_dir, extensions=args.extensions, apply=args.apply)

    print(f"Found {result['files_scanned']} file(s) under {args.v1_root}")
    print(f"Already in place: {result['already_in_place']}   Planned: {result['planned']}   "
          f"Unmatched: {result['unmatched']}   Ambiguous: {result['ambiguous']}   "
          f"Conflicts: {result['skipped_conflict']}")
    if result['applied']:
        print(f"Moved {result['moved']} of {result['planned']}. Errors: {result['errors']}")
    else:
        print("DRY RUN -- nothing moved. Review move_plan.csv, then re-run with --apply.")


if __name__ == '__main__':
    main()

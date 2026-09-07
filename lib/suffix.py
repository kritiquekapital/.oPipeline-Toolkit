#!/usr/bin/env python3
"""
lib/suffix.py

Cleans download-site suffixes and (optionally) underscore-formatted
filenames. Logic is unchanged from the original clean_filenames.py aside
from dropping the Disney/Eurobeat album-prefix stripper (too narrow a
one-off case to carry as permanent functionality) -- just wrapped so the
menu can call it directly instead of shelling out.
"""

from __future__ import annotations

import re
from pathlib import Path

DEFAULT_MARKERS = ["_spotdown.org", "_KLICKAUD", "spotdown.org", "KLICKAUD"]

# KLICKAUD names are reliably all-underscore ("Artist_-_Song_Title_KLICKAUD.mp3");
# spotdown.org names are normal spacing already. So underscore-to-space
# conversion should follow which SERVICE a file came from, not one global
# on/off switch applied to everything -- otherwise spotdown files that
# happen to have an underscore for another reason get mangled, or KLICKAUD
# files stay unreadable because the global toggle was off.
UNDERSCORE_SOURCE_MARKERS = {'_klickaud', 'klickaud'}


def strip_markers(name: str, markers: list[str]) -> tuple[str, set[str]]:
    """Returns (cleaned_name, matched_markers) -- matched_markers is the set
    of marker strings (lowercased) that were actually found and removed."""
    matched = set()
    for marker in markers:
        new_name, n = re.subn(re.escape(marker), '', name, flags=re.IGNORECASE)
        if n:
            matched.add(marker.lower())
        name = new_name
    return name, matched


def clean_name(name: str, markers: list[str], underscores_to_spaces: bool) -> str:
    stem_path = Path(name)
    stem, ext = stem_path.stem, stem_path.suffix

    stem, matched = strip_markers(stem, markers)
    marker_matched = bool(matched)

    # Auto-convert for KLICKAUD files regardless of the toggle; the toggle
    # remains available to force it for anything else too.
    if underscores_to_spaces or (matched & UNDERSCORE_SOURCE_MARKERS):
        stem = stem.replace('_', ' ')

    # Whitespace collapse is always safe. Trimming leading/trailing '_'/'-'
    # is NOT always safe -- it exists to clean up a dangling separator left
    # behind when a marker like "-KLICKAUD" (hyphen instead of underscore)
    # gets removed, not to tidy up every filename's punctuation. Only do it
    # when a marker was actually found on THIS file.
    stem = re.sub(r'\s+', ' ', stem)
    if marker_matched:
        stem = stem.strip(' _-')

    return stem + ext


def plan_renames(folder: Path, markers: list[str], underscores_to_spaces: bool, recursive: bool = False,
                  progress_cb=None):
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Not a valid folder: {folder}")

    files = [p for p in (folder.rglob("*") if recursive else folder.iterdir()) if p.is_file()]
    total = len(files)
    renames = []
    skipped = []

    for i, path in enumerate(files, start=1):
        new_name = clean_name(path.name, markers, underscores_to_spaces)
        if new_name != path.name:
            target = path.with_name(new_name)
            if target.exists():
                skipped.append(path.name)
            else:
                renames.append((path, target))
        if progress_cb:
            progress_cb(i, total)

    return renames, skipped


def apply_renames(renames):
    for path, target in renames:
        path.rename(target)
    return len(renames)


def run(folder: str, recursive: bool, underscores_to_spaces: bool,
        extra_markers: list[str] | None = None, apply: bool = False, progress_cb=None) -> dict:
    markers = list(DEFAULT_MARKERS)
    if extra_markers:
        markers.extend(extra_markers)

    renames, skipped = plan_renames(Path(folder), markers, underscores_to_spaces, recursive=recursive,
                                     progress_cb=progress_cb)

    result = {
        'planned': [(str(a), str(b)) for a, b in renames],
        'skipped': skipped,
        'applied': False,
        'renamed': 0,
    }
    if apply:
        result['renamed'] = apply_renames(renames)
        result['applied'] = True
    return result

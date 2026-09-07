#!/usr/bin/env python3
"""
pipeline.py -- V1 DJ Library Pipeline, interactive console entry point.

Run it with no arguments and it drives you through a menu:
    python pipeline.py

Settings (folder paths) are remembered between runs in pipeline_config.json
sitting next to this script, so you don't retype them every time.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from lib import crates, sheet, mover, suffix, library, reconcile, downloader  # noqa: E402

SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = SCRIPT_DIR / 'pipeline_config.json'
OUTPUT_DIR = SCRIPT_DIR / 'output'

DEFAULT_CONFIG = {
    'v1_root': r'E:\Music\V1',
    'incoming_dir': '',  # blank = same as V1 root (old behavior). Set this to a
                          # staging subfolder if you want a dedicated drop zone
                          # for new downloads, separate from your bucket folders.
    'crate_dir': r'E:\_Serato_\Subcrates',
    'crate_prefix': 'V1',
    'colorsheet_ps1': str(SCRIPT_DIR / 'Export-DJColorSheet.ps1'),
    'sheet_out': str(SCRIPT_DIR / 'output' / 'V1_Track_Metadata.xlsx'),
    'master_sheet_path': '',  # your real working V1.xlsx -- blank = not set up yet
    'extensions': ['.mp3'],
    # Filename-cleaner defaults (option 1). This step never touches Serato
    # or the sheet -- these are just the starting answers offered each time
    # you run it, so you're not re-typing the same choices every session.
    'suffix_recursive': True,
    'suffix_underscores_to_spaces': False,
    'suffix_extra_markers': [],
    'download_format': 'mp3',  # mp3/wav/flac/m4a/opus, or 'best' to keep the
                                # original container with no re-encoding
}

WIDTH = 64


# --------------------------------------------------------------------------
# tiny console-UI helpers -- no external dependencies (rich/colorama etc),
# just clean layout so it reads well in a plain cmd/PowerShell window.
# --------------------------------------------------------------------------

def _enable_ansi_on_windows():
    if os.name == 'nt':
        os.system('')  # harmless no-op that also flips on VT100 processing in modern conhost


def banner(title: str):
    print()
    print('=' * WIDTH)
    print(title.center(WIDTH))
    print('=' * WIDTH)


def main_banner():
    inner = WIDTH - 2
    motif = ".-~~~-"
    reps = inner // len(motif) + 2
    wave_top = (motif * reps)[:inner]
    wave_bottom = wave_top.replace('.', "'")[::-1]

    print()
    print(' ' + wave_top)
    print('(' + ' ' * inner + ')')
    print('(' + '.oPipeline Toolkit'.center(inner) + ')')
    print('(' + 'by .oCam'.center(inner) + ')')
    print('(' + ' ' * inner + ')')
    print(' ' + wave_bottom)


def rule(char: str = '-'):
    print(char * WIDTH)


def step(n: int, total: int, label: str):
    print(f"\n[{n}/{total}] {label}")
    rule()


def kv(label: str, value):
    print(f"  {label:<22} {value}")


# --------------------------------------------------------------------------
# color + progress -- ANSI is already enabled via _enable_ansi_on_windows(),
# this just actually uses it. Falls back to plain text when stdout isn't a
# real terminal (piped output, redirected to a file) or NO_COLOR is set.
# --------------------------------------------------------------------------

_COLOR_OK = sys.stdout.isatty() and not os.environ.get('NO_COLOR')


def _wrap(code: str, text: str) -> str:
    return f"\033[{code}m{text}\033[0m" if _COLOR_OK else text


def red(text: str) -> str:
    return _wrap('91', text)


def yellow(text: str) -> str:
    return _wrap('93', text)


def green(text: str) -> str:
    return _wrap('92', text)


def progress(current: int, total: int, label: str = "Scanning"):
    """Single updating console line -- call with current==total once at the
    end to leave a clean newline behind. No-op if total is 0."""
    if total <= 0:
        return
    pct = current / total * 100
    end = '\n' if current >= total else ''
    print(f"\r  {label}: {current}/{total} ({pct:5.1f}%)", end=end, flush=True)


class Timer:
    """`with Timer() as t: ...` then `t.elapsed` is seconds as a float,
    or str(t) for a ready-made "12.3s" label."""
    def __enter__(self):
        self._start = time.perf_counter()
        return self

    def __exit__(self, *exc):
        self.elapsed = time.perf_counter() - self._start

    def __str__(self):
        return f"{self.elapsed:.1f}s"


def confirm(prompt: str, default_yes: bool = False) -> bool:
    suffix_txt = 'Y/n' if default_yes else 'y/N'
    ans = input(f"{prompt} ({suffix_txt}): ").strip().lower()
    if not ans:
        return default_yes
    return ans.startswith('y')


def pause():
    input("\nPress Enter to return to the menu...")


# --------------------------------------------------------------------------
# config
# --------------------------------------------------------------------------

def load_config() -> dict:
    if CONFIG_PATH.exists():
        cfg = json.loads(CONFIG_PATH.read_text(encoding='utf-8'))
        merged = {**DEFAULT_CONFIG, **cfg}
        return merged
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict):
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding='utf-8')


def edit_settings(cfg: dict) -> dict:
    banner("SETTINGS")
    print("Press Enter on any line to keep the current value.\n")
    kv('Output folder (fixed)', OUTPUT_DIR)
    print()

    print("-- Serato / library paths (used by Refresh and Move) --\n")
    fields = [
        ('v1_root', 'V1 root folder (actual audio files -- scanned recursively for everything)'),
        ('incoming_dir', 'Incoming/staging folder for new downloads (blank = same as V1 root)'),
        ('crate_dir', 'Serato Subcrates folder'),
        ('crate_prefix', 'Crate filename prefix to include'),
    ]
    for key, label in fields:
        current = cfg.get(key, '')
        new = input(f"{label}\n  [{current or '(same as V1 root)'}]: ").strip()
        if new:
            if new.lower() in ('same', 'none', '-'):
                cfg[key] = ''
            else:
                if not Path(new).is_dir():
                    print(f"  WARNING: {new} doesn't exist (or isn't a folder) right now -- "
                          f"saving it anyway, but fix it before you rely on it.")
                cfg[key] = new

    print("\n-- File types to scan (used by Refresh and Move) --\n")
    current_exts = ', '.join(cfg.get('extensions', ['.mp3']))
    new = input(f"Extensions, comma-separated (e.g. .mp3, .flac, .wav)\n  [{current_exts}]: ").strip()
    if new:
        exts = [e.strip() for e in new.split(',') if e.strip()]
        exts = [e if e.startswith('.') else f'.{e}' for e in exts]
        cfg['extensions'] = exts

    print("\n-- Your real working spreadsheet (used by option 6, Reconcile) --\n")
    current = cfg.get('master_sheet_path', '')
    new = input(f"Path to your real V1.xlsx (blank if you don't use one)\n  [{current or '(not set)'}]: ").strip()
    if new:
        if new.lower() in ('none', '-'):
            cfg['master_sheet_path'] = ''
        else:
            if not Path(new).is_file():
                print(f"  WARNING: {new} doesn't exist right now -- saving it anyway, "
                      f"but fix it before you rely on it.")
            cfg['master_sheet_path'] = new

    print("\n-- Filename cleaner defaults (used by option 2 only -- doesn't touch")
    print("   Serato or the sheet, safe to run any time) --\n")
    cfg['suffix_recursive'] = confirm(
        "Include subfolders by default?", default_yes=cfg.get('suffix_recursive', True))
    cfg['suffix_underscores_to_spaces'] = confirm(
        "Convert underscores to spaces for ALL files by default? "
        "(KLICKAUD files always get this automatically, regardless of this setting)",
        default_yes=cfg.get('suffix_underscores_to_spaces', False))

    print("\n-- Download format (used by option 1) --\n")
    print("  mp3 (default, highest quality) / wav / flac / m4a / opus / best (no re-encode)")
    valid_formats = {'mp3', 'wav', 'flac', 'm4a', 'opus', 'aac', 'best'}
    current = cfg.get('download_format', 'mp3')
    new = input(f"Format\n  [{current}]: ").strip().lower()
    if new:
        if new not in valid_formats:
            print(f"  WARNING: '{new}' isn't one of the options above -- saving it anyway, "
                  f"but yt-dlp may reject it.")
        cfg['download_format'] = new

    save_config(cfg)
    print("\nSaved.")
    return cfg


# --------------------------------------------------------------------------
# step runners
# --------------------------------------------------------------------------

def do_crate_export(cfg: dict) -> dict | None:
    step(1, 3, "Reading Serato crate structure")
    kv('Crate folder', cfg['crate_dir'])
    kv('Prefix filter', cfg['crate_prefix'])
    try:
        with Timer() as t:
            diag = crates.run(cfg['crate_dir'], cfg['crate_prefix'], OUTPUT_DIR,
                               progress_cb=lambda c, n: progress(c, n, 'Parsing crates'))
    except Exception as e:
        print(red(f"\n  FAILED: {e}"))
        return None
    print(f"\n  Scanned {diag['total_crate_files_found']} .crate file(s), "
          f"included {diag['files_included']} (prefix '{cfg['crate_prefix']}')  ({t}).")
    if diag['files_failed_to_parse']:
        print(yellow(f"  WARNING: {diag['files_failed_to_parse']} file(s) failed to parse -- see crates_errors.json"))
    print(f"  {diag['track_count']} track entries written to {diag['flat_path']}")
    return diag


def do_color_export(cfg: dict) -> Path | None:
    step(2, 3, "Reading fresh Genre/Color tags from Serato (PowerShell + TagLib)")
    kv('Scanning', cfg['v1_root'])
    ps1 = cfg['colorsheet_ps1']
    if not Path(ps1).exists():
        print(red(f"\n  FAILED: script not found: {ps1}"))
        return None

    cmd = ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps1,
           '-Root', cfg['v1_root'], '-OutputDir', str(OUTPUT_DIR),
           '-Extensions'] + cfg['extensions']
    print()
    with Timer() as t:
        result = subprocess.run(cmd)  # inherits stdout/stderr so PS progress/colors show live
    if result.returncode != 0:
        print(red(f"\n  FAILED: color export exited with code {result.returncode}"))
        return None
    print(f"  ({t})")

    candidates = sorted(OUTPUT_DIR.glob('DJ_ColorSheet_*.csv'))
    if not candidates:
        print(red("\n  FAILED: color export ran but no DJ_ColorSheet_*.csv appeared in output/"))
        return None
    return candidates[-1]


def do_sheet_build(cfg: dict, flat_path: str, colorsheet_path: Path) -> dict | None:
    step(3, 3, "Building the metadata sheet")
    out_path = cfg['sheet_out']
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        with Timer() as t:
            result = sheet.run(flat_path, str(colorsheet_path), out_path, OUTPUT_DIR, is_dir=False)
    except PermissionError as e:
        print(red(f"\n  FAILED: {e}"))
        return None
    except Exception as e:
        print(red(f"\n  FAILED: {e}"))
        return None

    print(f"\n  Saved {result['out_path']}  ({t})")
    kv('Rows', result['rows'])
    kv('Matched to a crate', result['matched'])
    kv('Genre set', result['genre_set'])
    kv('Color set', result['colored'])
    if result['ambiguous']:
        print(yellow(f"  {'Ambiguous (needs review)':<22} {result['ambiguous']} -- see output/ambiguous_matches.csv"))
    else:
        kv('Ambiguous', 0)
    return result


def refresh_library(cfg: dict):
    banner("REFRESH SHEET (SCAN SERATO + REBUILD SPREADSHEET)")
    diag = do_crate_export(cfg)
    if not diag:
        return None
    colorsheet_path = do_color_export(cfg)
    if not colorsheet_path:
        return None
    result = do_sheet_build(cfg, diag['flat_path'], colorsheet_path)
    return result


def move_files(cfg: dict):
    banner("MOVE FILES (USING LAST REFRESH)")
    flat_path = OUTPUT_DIR / 'crates_flat.json'
    if not flat_path.exists():
        print("\nNo refresh has been run yet -- use Pipeline > Refresh + Move, or Sheet Tools > Refresh sheet only, first.")
        return

    print(f"Using crate data from: last refresh, {last_refresh_label()}")
    print("If you've re-tagged anything in Serato since then, use Pipeline > Refresh + Move instead.\n")

    print("Running a dry run first -- nothing will be moved yet.\n")
    with Timer() as t:
        dry = mover.run(cfg['v1_root'], str(flat_path), OUTPUT_DIR, extensions=cfg['extensions'], apply=False,
                         progress_cb=lambda c, n: progress(c, n, 'Scanning files'))

    kv('Files scanned', f"{dry['files_scanned']}  ({t})")
    kv('Already in place', dry['already_in_place'])
    kv('Planned moves', dry['planned'])
    if dry['unmatched']:
        print(yellow(f"  {'Unmatched (not crated yet)':<22} {dry['unmatched']}"))
    else:
        kv('Unmatched (not crated yet)', 0)
    if dry['ambiguous']:
        print(yellow(f"  {'Ambiguous (skipped, see log)':<22} {dry['ambiguous']}"))
    else:
        kv('Ambiguous (skipped, see log)', 0)
    kv('Conflicts (target exists)', dry['skipped_conflict'])

    if dry['planned'] == 0:
        print("\nNothing to move.")
        return

    print(f"\nFull plan written to {dry['reports']['move_plan.csv']} -- review it before applying.")
    if confirm(f"\nMove the {dry['planned']} planned file(s) now?"):
        with Timer() as t:
            applied = mover.run(cfg['v1_root'], str(flat_path), OUTPUT_DIR, extensions=cfg['extensions'],
                                 apply=True, progress_cb=lambda c, n: progress(c, n, 'Scanning files'))
        print(green(f"\nMoved {applied['moved']} of {applied['planned']}.") + f"  ({t})")
        if applied['errors']:
            print(red(f"{applied['errors']} error(s) -- see output/move_errors.txt"))
    else:
        print("\nSkipped -- no files moved.")


def reconcile_master_menu(cfg: dict):
    banner("RECONCILE INTO V1.xlsx MASTER")
    master_path = cfg.get('master_sheet_path')
    if not master_path:
        print("\nNo master spreadsheet configured -- set 'Path to your real V1.xlsx' in Settings first.")
        return
    if not Path(master_path).is_file():
        print(f"\nConfigured master file doesn't exist: {master_path}")
        print("Fix the path in Settings.")
        return

    fresh_path = cfg['sheet_out']
    if not Path(fresh_path).is_file():
        print("\nNo fresh sheet to merge yet -- run Sheet Tools > Refresh sheet only (or Pipeline > Refresh + Move) first.")
        return

    preview_path = OUTPUT_DIR / 'V1_reconciled_preview.xlsx'
    print(f"Master:  {master_path}")
    print(f"Fresh:   {fresh_path}")
    print("\nBuilding a preview merge (your real file is not touched yet)...\n")
    try:
        with Timer() as t:
            result = reconcile.merge(master_path, fresh_path, str(preview_path))
    except Exception as e:
        print(red(f"FAILED: {e}"))
        return

    kv('Rows updated', result['updated'])
    kv('New rows appended', result['appended'])
    if result['ambiguous_skipped']:
        print(yellow(f"  {'Ambiguous (bucket left as-is)':<22} {len(result['ambiguous_skipped'])}"))
    else:
        kv('Ambiguous (bucket left as-is)', 0)
    kv('In master but not in this refresh', len(result['not_in_refresh']))
    kv('Time', str(t))
    if result['not_in_refresh']:
        print("\n  (not touched or deleted -- just not seen this run:)")
        for name in result['not_in_refresh'][:10]:
            print(f"    {name}")
        if len(result['not_in_refresh']) > 10:
            print(f"    ... and {len(result['not_in_refresh']) - 10} more")

    print(f"\nPreview written to {preview_path} -- open it and check it before replacing your real file.")
    if not confirm("\nReplace your real V1.xlsx with this merged version now? "
                    "(a timestamped backup will be made first)"):
        print("\nSkipped -- your real file is untouched. Preview left in output/ for review.")
        return

    import datetime
    import shutil as _shutil
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_path = Path(master_path).with_name(f"{Path(master_path).stem}.bak-{timestamp}{Path(master_path).suffix}")
    try:
        _shutil.copy2(master_path, backup_path)
    except Exception as e:
        print(red(f"\nFAILED to create backup, aborting -- nothing was overwritten: {e}"))
        return
    print(f"\nBackup saved: {backup_path}")

    try:
        _shutil.copy2(preview_path, master_path)
    except PermissionError:
        print(red(f"\nFAILED: could not overwrite {master_path} -- it's probably open in Excel right now."))
        print(f"Your backup is safe at {backup_path}. Close the file and try again.")
        return
    print(green(f"Done -- {master_path} updated."))


def undo_last_move_menu(cfg: dict):
    banner("UNDO LAST MOVE")
    result = mover.undo_last_move(OUTPUT_DIR)
    if not result['available']:
        print("\nNothing to undo -- no move has been applied yet (or it was already undone).")
        return

    print(f"Last apply moved {result['total']} file(s).")
    if not confirm(f"Move all {result['total']} back to where they came from?"):
        print("\nCancelled.")
        return

    print()
    print(green(f"Reverted: {result['reverted']} of {result['total']}"))
    if result['failed']:
        print(red(f"Could not revert {len(result['failed'])}:"))
        for line in result['failed']:
            print(f"  {line}")


def library_health_check(cfg: dict):
    banner("LIBRARY HEALTH CHECK")
    sheet_path = cfg['sheet_out']
    if not Path(sheet_path).exists():
        print("\nNo sheet built yet -- run a Refresh first.")
        return

    print("(Reads the last built sheet -- doesn't rescan Serato or files.)\n")
    h = library.health_summary(sheet_path)

    kv('Total tracks', h['total'])
    kv('Unmatched to a bucket', h['unmatched'])
    kv('Ambiguous', h['ambiguous'])
    kv('Missing Genre', h['no_genre'])
    kv('Missing Color', h['no_color'])
    kv('Missing BPM', h['no_bpm'])
    kv('Missing Key', h['no_key'])

    print("\nBUCKETS NEEDING ATTENTION (lowest color coverage first):")
    for b, n, gpct, cpct in h['bucket_coverage'][:8]:
        flag = '  <-- low coverage' if cpct < 20 and n >= 10 else ''
        print(f"  {n:>4} tracks   genre {gpct:5.1f}%   color {cpct:5.1f}%   {b}{flag}")

    print("\nBUCKETS (track count):")
    for name, count in h['buckets']:
        print(f"  {count:>4}  {name}")

    print("\nGENRES (top 10):")
    for name, count in h['genres'][:10]:
        print(f"  {count:>4}  {name}")

    print("\nCOLORS (in order, most-used first):")
    for i, (name, count) in enumerate(h['colors'], start=1):
        print(f"  {i:>2}. {count:>4}  {name}")


def duplicate_finder_menu(cfg: dict):
    banner("DUPLICATE FINDER")
    sheet_path = cfg['sheet_out']
    if not Path(sheet_path).exists():
        print("\nNo sheet built yet -- run a Refresh first.")
        return

    print("(Read-only -- reports only, never merges, deletes, or moves anything.)\n")
    d = library.find_duplicates(sheet_path)

    print(f"Same filename in more than one place: {len(d['filename_dupes'])} group(s)")
    for group in d['filename_dupes'][:15]:
        print(f"\n  {group[0]['FileName']}")
        for r in group:
            print(f"      {r['FullPath']}")
    if len(d['filename_dupes']) > 15:
        print(f"\n  ... and {len(d['filename_dupes']) - 15} more group(s)")

    print(f"\n\nSame Artist+Title under different filenames: {len(d['artist_title_dupes'])} group(s)")
    print("(often a re-download or re-encode of the same track -- not always a real duplicate)")
    for group in d['artist_title_dupes'][:15]:
        print(f"\n  {group[0]['Artist']} - {group[0]['Title']}")
        for r in group:
            print(f"      {r['FileName']}")
    if len(d['artist_title_dupes']) > 15:
        print(f"\n  ... and {len(d['artist_title_dupes']) - 15} more group(s)")


def view_logs_menu(cfg: dict):
    banner("VIEW LAST RUN LOGS")
    log_files = [
        ('crates_errors.json', 'Crate files that failed to parse'),
        ('unmatched.txt', 'Files not matched to any crate (move step)'),
        ('ambiguous.txt', 'Files skipped for the move -- ambiguous crate match'),
        ('ambiguous_matches.csv', 'Rows flagged ambiguous in the sheet'),
        ('skipped_conflict.txt', 'Files skipped -- target already existed'),
        ('move_errors.txt', 'Files that failed to move'),
    ]
    found_any = False
    for filename, label in log_files:
        path = OUTPUT_DIR / filename
        if not path.exists():
            continue
        found_any = True
        content = path.read_text(encoding='utf-8').strip()
        lines = content.splitlines()
        print(f"\n--- {filename} ({label}) ---")
        for line in lines[:30]:
            print(f"  {line}")
        if len(lines) > 30:
            print(f"  ... and {len(lines) - 30} more -- see output/{filename}")

    if not found_any:
        print("\nNo logs found -- either nothing's been run yet, or the last runs were clean.")


def search_library_menu(cfg: dict):
    banner("SEARCH LIBRARY")
    sheet_path = cfg['sheet_out']
    if not Path(sheet_path).exists():
        print("\nNo sheet built yet -- run a Refresh first.")
        return

    term = input("Search (filename / artist / title): ").strip()
    if not term:
        return

    hits = library.search(sheet_path, term, limit=25)
    if not hits:
        print("\nNo matches.")
        return

    print(f"\n{len(hits)} match(es){' (showing first 25)' if len(hits) == 25 else ''}:\n")
    for r in hits:
        bucket = r['ParentCrate'] or '(unmatched)'
        sub = f" > {r['SubCrate']}" if r['SubCrate'] else ''
        genre = r['Genre'] or '(none)'
        color = r['ColorName'] or '(not tagged)'
        print(f"  {r['FileName']}")
        print(f"      Artist: {r['Artist'] or '?'}   Bucket: {bucket}{sub}   Genre: {genre}   Color: {color}")


def download_menu(cfg: dict):
    banner("DOWNLOAD FROM URL")
    print("Uses yt-dlp -- works with SoundCloud, YouTube, Bandcamp, and many other sites.")
    print("Only for tracks you actually have the right to download (your own uploads,")
    print("Creative Commons, artist-enabled downloads, DJ pool content, etc).\n")
    kv('Format', f"{cfg.get('download_format', 'mp3')}  (change in Settings)")
    print()

    if not downloader.is_available():
        print("yt-dlp isn't installed.")
        if confirm("Install it now with pip?", default_yes=True):
            cmd = [sys.executable, '-m', 'pip', 'install', 'yt-dlp']
            print(f"\nRunning: {' '.join(cmd)}\n")
            result = subprocess.run(cmd)
            if result.returncode != 0:
                print("\nInstall failed. Run `pip install yt-dlp` yourself, then try again.")
                return
            print("\nInstalled.\n")
        else:
            return

    if not downloader.has_ffmpeg():
        print("WARNING: ffmpeg not found on PATH -- yt-dlp needs it to convert to mp3.")
        print("Install it from https://ffmpeg.org/download.html (add it to PATH), then try again.")
        return

    url = input("URL to download: ").strip()
    if not url:
        return

    dest_dir = cfg.get('incoming_dir') or cfg['v1_root']
    if not Path(dest_dir).is_dir():
        print(f"\nNOTE: configured Incoming folder doesn't exist: {dest_dir}")
        print("Falling back to V1 root for now -- fix it in Settings.\n")
        dest_dir = cfg['v1_root']

    print(f"\nDownloading to: {dest_dir}\n")
    with Timer() as t:
        result = downloader.download(url, dest_dir, audio_format=cfg.get('download_format', 'mp3'))
    if result['success']:
        print(green(f"\nDone.") + f"  ({t})  Run Pipeline > Clean filenames next if the title needs tidying up.")
    else:
        print(red(f"\nFAILED -- yt-dlp exited with code {result['returncode']}. See its output above."))


def clean_filenames_menu(cfg: dict):
    banner("CLEAN FILENAMES")
    print("(Stand-alone -- doesn't touch Serato, crates, or the sheet. Safe to run any time.)\n")
    default_folder = cfg.get('incoming_dir') or cfg['v1_root']
    if not Path(default_folder).is_dir():
        print(f"NOTE: configured Incoming folder doesn't exist: {default_folder}")
        print(f"Falling back to V1 root for now -- fix it in Settings.\n")
        default_folder = cfg['v1_root']
    folder = input(f"Folder to clean (Enter for {default_folder}): ").strip()
    if not folder:
        folder = default_folder

    recursive = confirm("Include subfolders?", default_yes=cfg.get('suffix_recursive', True))
    underscores = confirm("Convert underscores to spaces for ALL files? "
                           "(KLICKAUD files get this automatically either way)",
                           default_yes=cfg.get('suffix_underscores_to_spaces', False))

    print("\n--- Dry run ---\n")
    result = suffix.run(folder, recursive, underscores,
                         extra_markers=cfg.get('suffix_extra_markers'), apply=False,
                         progress_cb=lambda c, n: progress(c, n, 'Scanning'))
    for src, dst in result['planned']:
        print(f"  {Path(src).name}  ->  {Path(dst).name}")
    if result['skipped']:
        print(f"\n  {len(result['skipped'])} skipped (target already exists)")
    print(f"\nWould rename {len(result['planned'])} file(s).")

    if not result['planned']:
        return

    if confirm("\nLooked right? Actually rename now?"):
        with Timer() as t:
            applied = suffix.run(folder, recursive, underscores,
                                  extra_markers=cfg.get('suffix_extra_markers'), apply=True)
        print(green(f"\nRenamed {applied['renamed']} file(s).") + f"  ({t})")
    else:
        print("\nSkipped -- nothing renamed.")


# --------------------------------------------------------------------------
# menu
# --------------------------------------------------------------------------

def last_refresh_label() -> str:
    flat_path = OUTPUT_DIR / 'crates_flat.json'
    if not flat_path.exists():
        return 'never -- run Refresh first'
    import datetime
    age = datetime.datetime.now() - datetime.datetime.fromtimestamp(flat_path.stat().st_mtime)
    mins = int(age.total_seconds() // 60)
    if mins < 1:
        return 'just now'
    if mins < 60:
        return f'{mins} min ago'
    hours = mins // 60
    if hours < 24:
        return f'{hours} hr ago'
    return f'{hours // 24} day(s) ago'


def refresh_and_move(cfg: dict):
    result = refresh_library(cfg)
    if result is not None:
        move_files(cfg)


MAIN_MENU = """
  [1] Pipeline       -- download, clean, move, refresh+move, undo, reconcile
  [2] Sheet Tools    -- refresh sheet, health check, duplicate finder
  [3] Search & Logs  -- view logs, search library
  [4] Settings
  [Q] Quit
"""

PIPELINE_ITEMS = [
    ('1', 'Download from URL      -- yt-dlp, drops straight into Incoming', download_menu),
    ('2', 'Clean filenames        -- rename files (suffixes/prefixes)', clean_filenames_menu),
    ('3', 'Move files (last scan) -- uses the LAST refresh\'s data', move_files),
    ('4', 'Refresh + Move         -- rescan Serato, THEN move', refresh_and_move),
    ('5', 'Undo last move         -- move everything back from the last apply', undo_last_move_menu),
    ('6', 'Reconcile into V1.xlsx -- merge the last refresh into your real master', reconcile_master_menu),
]

SHEET_ITEMS = [
    ('1', 'Refresh sheet only     -- rescan Serato, rebuild the spreadsheet', refresh_library),
    ('2', 'Library health check   -- bucket/genre/color counts, no rescan', library_health_check),
    ('3', 'Duplicate finder       -- same filename or same Artist+Title, read-only', duplicate_finder_menu),
]

SEARCH_ITEMS = [
    ('1', 'View last run logs -- unmatched/ambiguous/errors, in console', view_logs_menu),
    ('2', 'Search library     -- look up a track without opening Excel', search_library_menu),
]


def run_submenu(title: str, subtitle: str, items: list, cfg: dict):
    while True:
        banner(title)
        print(subtitle + "\n")
        for key, label, _ in items:
            print(f"  [{key}] {label}")
        print("\n  [B] Back")
        print("  [Q] Quit")
        choice = input("\nChoice: ").strip().lower()

        if choice == 'b':
            return
        if choice == 'q':
            print("\nBye.")
            sys.exit(0)

        for key, _, func in items:
            if choice == key:
                func(cfg)
                pause()
                break
        else:
            print("\nNot a valid choice.")


def check_dependencies() -> bool:
    """
    First-run / new-machine check. Verifies the Python packages this tool
    needs are actually importable, and offers to install them on the spot
    instead of crashing partway through a menu option with a bare
    ModuleNotFoundError. Returns True if it's safe to continue.

    yt-dlp/ffmpeg (for Download) are checked lazily inside that menu
    instead of here, since most runs won't touch that feature at all.
    """
    missing = []
    for pkg in ('pandas', 'openpyxl'):
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)

    if missing:
        banner("FIRST-TIME SETUP")
        print(f"Missing required Python package(s): {', '.join(missing)}\n")
        if confirm(f"Install {' '.join(missing)} now with pip?", default_yes=True):
            cmd = [sys.executable, '-m', 'pip', 'install', *missing]
            print(f"\nRunning: {' '.join(cmd)}\n")
            result = subprocess.run(cmd)
            if result.returncode != 0:
                print("\nInstall failed. Run this yourself in a terminal, then relaunch:")
                print(f"  pip install {' '.join(missing)}")
                return False
            print("\nInstalled. Continuing...\n")
        else:
            print("\nCan't continue without these. Run this, then relaunch:")
            print(f"  pip install {' '.join(missing)}")
            return False

    if os.name == 'nt':
        import shutil as _shutil
        if not _shutil.which('powershell.exe'):
            print("\nWARNING: powershell.exe not found on PATH.")
            print("Pipeline's color/genre step needs it -- everything else will still work.")

    return True


def main():
    _enable_ansi_on_windows()
    if not check_dependencies():
        input("\nPress Enter to exit...")
        return
    cfg = load_config()

    while True:
        main_banner()
        kv('Last refresh', last_refresh_label())
        print(MAIN_MENU)
        choice = input("Choice: ").strip().lower()

        if choice == '1':
            run_submenu("PIPELINE", "Changes files on disk -- download, rename, move, undo, reconcile.",
                        PIPELINE_ITEMS, cfg)
        elif choice == '2':
            run_submenu("SHEET TOOLS", "Read-only or sheet-only -- never touches audio files.",
                        SHEET_ITEMS, cfg)
        elif choice == '3':
            run_submenu("SEARCH & LOGS", "Read-only lookups against the last built sheet.",
                        SEARCH_ITEMS, cfg)
        elif choice == '4':
            cfg = edit_settings(cfg)
            pause()
        elif choice == 'q':
            print("\nBye.")
            break
        else:
            print("\nNot a valid choice.")


if __name__ == '__main__':
    main()

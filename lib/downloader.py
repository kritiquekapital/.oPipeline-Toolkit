#!/usr/bin/env python3
"""
lib/downloader.py

Downloads audio from a URL via yt-dlp (SoundCloud, YouTube, Bandcamp, and
many other sites yt-dlp supports) straight into the configured Incoming
folder -- so it lands exactly where Clean Filenames already expects new
downloads to be.

This only wraps yt-dlp's ordinary "pull the audio a site serves through
its normal player/download endpoint" behavior -- the same thing a
browser's own download button would do for a track the uploader has
enabled downloads for, a Creative Commons track, or your own upload.
Nothing here is DRM-circumventing, and it's on you to only point it at
things you actually have the right to download.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


def is_available() -> bool:
    try:
        import yt_dlp  # noqa: F401
        return True
    except ImportError:
        return False


def has_ffmpeg() -> bool:
    return shutil.which('ffmpeg') is not None


def download(url: str, dest_dir: str, audio_format: str = 'mp3') -> dict:
    """
    Downloads the best-available audio from `url` into `dest_dir`,
    converting to `audio_format` (or keeping the original container if
    audio_format is 'best'), embedding thumbnail/metadata where the
    source provides it. Streams yt-dlp's own output live so progress is
    visible, and also captures it in case of failure.
    """
    Path(dest_dir).mkdir(parents=True, exist_ok=True)
    cmd = [
        sys.executable, '-m', 'yt_dlp',
        '-x', '--audio-format', audio_format, '--audio-quality', '0',
        '--embed-thumbnail', '--add-metadata',
        '-o', str(Path(dest_dir) / '%(title)s.%(ext)s'),
        url,
    ]
    result = subprocess.run(cmd)  # inherits stdout/stderr so progress shows live
    return {
        'returncode': result.returncode,
        'success': result.returncode == 0,
    }

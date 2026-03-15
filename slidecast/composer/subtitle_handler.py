"""Subtitle file validation and discovery helpers."""
from __future__ import annotations

import os
from pathlib import Path


def validate_subtitle(path: str) -> bool:
    """Basic check that the file is readable and non-empty."""
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            content = f.read(512)
        return len(content.strip()) > 0
    except Exception:
        return False


def subtitle_lang_from_filename(path: str) -> str:
    """Guess ISO 639-3 language code from filename (e.g. subtitle_kor.srt → kor)."""
    stem = Path(path).stem.lower()
    if "kor" in stem or "ko" in stem:
        return "kor"
    if "eng" in stem or "en" in stem:
        return "eng"
    if "jpn" in stem or "ja" in stem:
        return "jpn"
    return "kor"  # default to Korean

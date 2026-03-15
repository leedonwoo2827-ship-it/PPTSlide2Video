"""File path helpers and media discovery."""
from __future__ import annotations

import os
import shutil
import tempfile
from pathlib import Path


AUDIO_EXTS = (".mp3", ".wav", ".m4a", ".aac", ".ogg", ".flac")


def find_mp3(folder: str) -> str | None:
    """Return the first audio file found in folder, or None.
    Supports: mp3, wav, m4a, aac, ogg, flac
    """
    for f in sorted(Path(folder).iterdir()):
        if f.suffix.lower() in AUDIO_EXTS and not _is_numbered(f.stem):
            return str(f)
    return None


def find_per_slide_audio(folder: str, n_slides: int) -> list[str] | None:
    """
    슬라이드별 번호 음성 파일 탐색.

    지원 형식 (슬라이드 1번 예시):
      01.wav          ← 순수 번호
      1.mp3
      2-01.wav        ← 프리픽스-번호 (챕터, 날짜 등)
      chapter_01.mp3  ← 프리픽스_번호
      any-prefix-01.wav

    n_slides 만큼 순서대로 파일이 있으면 경로 리스트 반환,
    하나라도 없으면 None 반환.
    """
    folder_path = Path(folder)

    # 폴더 내 오디오 파일 전체 목록을 미리 수집
    audio_files = [
        f for f in sorted(folder_path.iterdir())
        if f.suffix.lower() in AUDIO_EXTS
    ]

    found: list[str] = []

    for i in range(1, n_slides + 1):
        match = _find_numbered_audio(audio_files, i)
        if match is None:
            return None   # 하나라도 없으면 포기
        found.append(match)

    return found if found else None


def _find_numbered_audio(audio_files: list[Path], n: int) -> str | None:
    """
    파일 목록에서 슬라이드 번호 n에 해당하는 파일을 찾는다.
    파일명이 n 또는 zero-padded n 으로 끝나면 매칭.
    예: n=1 → "01", "1", "2-01", "chapter_1", "any_prefix-01" 모두 매칭
    """
    targets = {f"{n:03d}", f"{n:02d}", f"{n}"}   # "001", "01", "1"
    for f in audio_files:
        stem = f.stem  # 확장자 제외 파일명
        # 순수 번호 (01.wav, 1.wav)
        if stem in targets:
            return str(f)
        # 구분자(-_) 뒤에 번호로 끝나는 경우 (2-01.wav, chapter_01.mp3)
        for sep in ("-", "_"):
            last = stem.rsplit(sep, 1)[-1]
            if last in targets:
                return str(f)
    return None


def _is_numbered(stem: str) -> bool:
    """파일명이 순수 숫자인지 확인 (01, 1, 02 등)."""
    return stem.lstrip("0").isdigit() or stem.isdigit()


def find_subtitle(folder: str) -> tuple[str, str] | None:
    """Return (path, ext) for first .srt or .vtt found in folder, or None."""
    for ext in (".srt", ".vtt"):
        for f in Path(folder).iterdir():
            if f.suffix.lower() == ext:
                return str(f), ext
    return None


def make_temp_dir(base: str) -> str:
    """Create a unique temp directory under base."""
    os.makedirs(base, exist_ok=True)
    return tempfile.mkdtemp(prefix="ppt2slide_", dir=base)


def cleanup_temp_dir(path: str) -> None:
    if path and os.path.exists(path):
        shutil.rmtree(path, ignore_errors=True)


def forward_slashes(path: str) -> str:
    """Ensure forward slashes for FFmpeg compatibility on Windows."""
    return path.replace("\\", "/")

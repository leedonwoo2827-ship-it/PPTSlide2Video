"""FFmpeg-based video composition: concat, audio merge, subtitle mux."""
from __future__ import annotations

import os
import subprocess
import tempfile
from pathlib import Path

from slidecast.utils.file_utils import forward_slashes


def _run_ffmpeg(args: list[str], description: str = "") -> None:
    cmd = ["ffmpeg", "-y"] + args
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if result.returncode != 0:
        raise RuntimeError(
            f"FFmpeg failed ({description}):\n{result.stderr[-3000:]}"
        )


def concat_audio_files(audio_paths: list[str], output_path: str) -> str:
    """
    슬라이드별 음성 파일(01.wav, 02.wav...)을 순서대로 이어붙여 하나의 오디오로 만든다.
    """
    list_content = "\n".join(
        f"file '{forward_slashes(os.path.abspath(p))}'" for p in audio_paths
    )
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".txt", delete=False, encoding="utf-8"
    ) as f:
        f.write(list_content)
        list_path = f.name

    try:
        _run_ffmpeg([
            "-f", "concat",
            "-safe", "0",
            "-i", list_path,
            "-c:a", "aac",
            "-b:a", "192k",
            output_path,
        ], "concat_audio")
    finally:
        os.unlink(list_path)

    return output_path


def concat_videos(video_paths: list[str], output_path: str) -> str:
    """
    Concatenate multiple WebM/MP4 video clips into one file.
    Uses FFmpeg concat demuxer (no re-encode for same codec streams).
    """
    # Write concat list to temp file
    list_content = "\n".join(
        f"file '{forward_slashes(os.path.abspath(p))}'" for p in video_paths
    )
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".txt", delete=False, encoding="utf-8"
    ) as f:
        f.write(list_content)
        list_path = f.name

    try:
        _run_ffmpeg([
            "-f", "concat",
            "-safe", "0",
            "-i", list_path,
            "-c", "copy",
            output_path,
        ], "concat")
    finally:
        os.unlink(list_path)

    return output_path


def merge_audio(video_path: str, audio_path: str, output_path: str) -> str:
    """
    Merge MP3 audio into video.
    - If audio is longer than video: video loops its last frame (via -loop 1 trick
      is not used here; instead video is extended via -shortest inverse — we pad video).
    - If audio is shorter: video is trimmed to audio length.
    Uses -shortest to sync lengths automatically.
    """
    _run_ffmpeg([
        "-i", video_path,
        "-i", audio_path,
        "-map", "0:v:0",
        "-map", "1:a:0",
        "-c:v", "libx264",      # re-encode to H264 for broad MP4 compatibility
        "-preset", "fast",
        "-crf", "18",
        "-c:a", "aac",
        "-b:a", "192k",
        "-shortest",            # trim to shorter of video/audio
        "-movflags", "+faststart",
        output_path,
    ], "merge_audio")
    return output_path


def extend_video_to_audio(video_path: str, audio_path: str, output_path: str) -> str:
    """
    Extend video to match audio duration by freezing the last frame.
    Used when audio is longer than the total slide animation time.
    """
    # Get audio duration
    audio_dur = _get_duration(audio_path)
    video_dur = _get_duration(video_path)

    if audio_dur <= video_dur:
        # Audio is shorter or equal — just merge normally
        return merge_audio(video_path, audio_path, output_path)

    # Freeze last frame: use tpad filter to extend video
    extra = audio_dur - video_dur
    _run_ffmpeg([
        "-i", video_path,
        "-i", audio_path,
        "-filter_complex",
        f"[0:v]tpad=stop_mode=clone:stop_duration={extra:.3f}[v]",
        "-map", "[v]",
        "-map", "1:a:0",
        "-c:v", "libx264",
        "-preset", "fast",
        "-crf", "18",
        "-c:a", "aac",
        "-b:a", "192k",
        "-movflags", "+faststart",
        output_path,
    ], "extend_video_to_audio")
    return output_path


def add_subtitles_soft(video_path: str, subtitle_path: str, output_path: str,
                       lang: str = "kor") -> str:
    """
    Add subtitle as a soft track (separate subtitle stream in MP4 container).
    Players that support soft subtitles can toggle them on/off.
    """
    sub_ext = Path(subtitle_path).suffix.lower()
    sub_codec = "mov_text" if sub_ext == ".srt" else "webvtt"

    _run_ffmpeg([
        "-i", video_path,
        "-i", subtitle_path,
        "-map", "0",
        "-map", "1",
        "-c:v", "copy",
        "-c:a", "copy",
        "-c:s", sub_codec,
        "-metadata:s:s:0", f"language={lang}",
        "-movflags", "+faststart",
        output_path,
    ], "add_subtitles_soft")
    return output_path


def burn_subtitles(video_path: str, subtitle_path: str, output_path: str) -> str:
    """
    Hard-burn subtitles into video pixels (universally compatible).
    Requires re-encoding; slower but works in all players.
    """
    import shutil

    # FFmpeg subtitles filter cannot handle non-ASCII or spaces in paths on Windows.
    # Copy the SRT to a guaranteed-safe temp location (ASCII-only path).
    with tempfile.NamedTemporaryFile(
        suffix=".srt", delete=False, dir=tempfile.gettempdir(), encoding=None
    ) as tmp:
        safe_srt_path = tmp.name

    shutil.copy2(subtitle_path, safe_srt_path)

    try:
        safe_sub = forward_slashes(safe_srt_path)
        # On Windows, drive letters need escaping for the subtitles filter
        safe_sub = safe_sub.replace(":", "\\:")

        _run_ffmpeg([
            "-i", video_path,
            "-vf", f"subtitles='{safe_sub}'",
            "-c:v", "libx264",
            "-preset", "fast",
            "-crf", "18",
            "-c:a", "copy",
            "-movflags", "+faststart",
            output_path,
        ], "burn_subtitles")
    finally:
        os.unlink(safe_srt_path)

    return output_path


def _get_duration(file_path: str) -> float:
    """Use ffprobe to get media duration in seconds."""
    result = subprocess.run(
        [
            "ffprobe", "-v", "error",
            "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1",
            file_path,
        ],
        capture_output=True, text=True, encoding="utf-8",
    )
    try:
        return float(result.stdout.strip())
    except Exception:
        return 0.0

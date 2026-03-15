"""Main pipeline orchestrator: PPTX → HTML → Video → MP4."""
from __future__ import annotations

import asyncio
import os

from slidecast.composer.ffmpeg_composer import (
    add_subtitles_soft,
    burn_subtitles,
    concat_audio_files,
    concat_videos,
    extend_video_to_audio,
    _run_ffmpeg,
)
from slidecast.composer.subtitle_handler import (
    subtitle_lang_from_filename,
    validate_subtitle,
)
from slidecast.composer.notes_to_srt import generate_srt_from_notes
from slidecast.composer.ffmpeg_composer import _get_duration
from slidecast.generator.html_generator import generate_html_slides
from slidecast.parser.pptx_parser import parse_pptx
from slidecast.renderer.playwright_renderer import render_all_slides
from slidecast.utils.file_utils import (
    cleanup_temp_dir,
    find_mp3,
    find_per_slide_audio,
    find_subtitle,
    make_temp_dir,
)


async def run_pipeline(
    pptx_path: str,
    media_folder: str,
    output_path: str,
    subtitle_mode: str = "hard",
    slide_hold_seconds: float = 1.0,
    progress_callback=None,
) -> str:
    """
    Full pipeline: PPTX → animated MP4 with audio and subtitles.

    Args:
        pptx_path: Absolute path to input .pptx file
        media_folder: Folder containing .mp3 and .srt/.vtt files
        output_path: Absolute path for output .mp4 file
        subtitle_mode: "soft" (subtitle track) or "hard" (burned into video)
        slide_hold_seconds: Seconds to hold each slide after animations complete
        progress_callback: Optional async callable(step: str, pct: int)

    Returns:
        Path to the produced MP4 file.
    """

    def _progress(step: str, pct: int):
        if progress_callback:
            asyncio.ensure_future(progress_callback(step, pct))

    output_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(output_dir, exist_ok=True)
    temp_dir = make_temp_dir(output_dir)

    try:
        # ── Stage 1: Parse PPTX ─────────────────────────────────────────────
        _progress("Parsing PPTX…", 5)
        presentation = parse_pptx(pptx_path, temp_dir)
        n_slides = len(presentation.slides)
        if n_slides == 0:
            raise ValueError("No slides found in the PPTX file.")

        # ── Stage 2: Generate HTML slides ───────────────────────────────────
        _progress("Generating HTML slides…", 15)
        html_paths = generate_html_slides(presentation, temp_dir)

        # ── 슬라이드별 음성 파일 탐색 (렌더링 전에 필요) ───────────────────
        per_slide = find_per_slide_audio(media_folder, n_slides)
        per_slide_durs: list[float] | None = None
        if per_slide:
            per_slide_durs = [_get_duration(p) for p in per_slide]

        # ── Stage 3: Render slides to video ─────────────────────────────────
        _progress("Rendering slides (this may take a while)…", 25)
        video_paths = await render_all_slides(
            html_paths=html_paths,
            slides_meta=presentation.slides,
            temp_dir=temp_dir,
            hold_seconds=slide_hold_seconds,
            per_slide_durations=per_slide_durs,
            pptx_path=pptx_path,
        )

        # ── Stage 4a: Concatenate slide videos ──────────────────────────────
        _progress("Concatenating slide videos…", 65)
        concat_path = os.path.join(temp_dir, "concat.mp4")
        if len(video_paths) == 1:
            # PowerPoint 네이티브 단일 MP4 → concat 불필요
            concat_path = video_paths[0]
        else:
            concat_videos(video_paths, concat_path)

        # ── Stage 4b: Merge audio ────────────────────────────────────────────
        # 우선순위: ① 슬라이드별 번호 파일(01.wav...) → ② 단일 오디오 파일
        with_audio = os.path.join(temp_dir, "with_audio.mp4")

        if per_slide:
            _progress(f"슬라이드별 음성 파일 {n_slides}개 이어붙이는 중…", 70)
            merged_audio = os.path.join(temp_dir, "merged_audio.m4a")
            concat_audio_files(per_slide, merged_audio)
            _progress("음성 합치는 중…", 75)
            extend_video_to_audio(concat_path, merged_audio, with_audio)
        else:
            mp3_path = find_mp3(media_folder)
            if mp3_path:
                _progress("Merging audio (MP3 duration sets video length)…", 75)
                extend_video_to_audio(concat_path, mp3_path, with_audio)
            else:
                # No audio: just convert WebM → MP4
                _run_ffmpeg([
                    "-i", concat_path,
                    "-c:v", "libx264", "-preset", "fast", "-crf", "18",
                    "-movflags", "+faststart",
                    with_audio,
                ], "webm_to_mp4")

        # ── Stage 5: Subtitles ───────────────────────────────────────────────
        # 우선순위: ① 미디어 폴더의 SRT/VTT 파일 → ② 슬라이드 노트 자동 변환
        sub_path: str | None = None
        sub_result = find_subtitle(media_folder)
        if sub_result and validate_subtitle(sub_result[0]):
            sub_path, _ = sub_result
            _progress("Adding subtitles (from SRT file)…", 90)
        else:
            # 슬라이드 노트에서 자막 자동 생성 — WAV 길이 기반 싱크
            total_dur = _get_duration(with_audio)
            auto_srt = os.path.join(temp_dir, "notes_auto.srt")
            sub_path = generate_srt_from_notes(
                presentation.slides, total_dur, auto_srt,
                slide_durations=per_slide_durs,
            )
            if sub_path:
                _progress("Adding subtitles (from slide notes)…", 90)

        if sub_path:
            lang = subtitle_lang_from_filename(sub_path)
            if subtitle_mode == "hard":
                burn_subtitles(with_audio, sub_path, output_path)
            else:
                add_subtitles_soft(with_audio, sub_path, output_path, lang=lang)
        else:
            _finalize(with_audio, output_path)

        _progress("Done!", 100)
        return output_path

    finally:
        cleanup_temp_dir(temp_dir)


def _finalize(src: str, dst: str) -> None:
    """Copy/move the intermediate MP4 to the final output path."""
    import shutil
    shutil.copy2(src, dst)

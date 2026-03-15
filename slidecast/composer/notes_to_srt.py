"""
슬라이드 노트 → SRT 자막 자동 변환

각 슬라이드의 노트 텍스트를 읽어서 SRT 파일로 만든다.
타이밍은 (MP3 총 길이 / 슬라이드 수) 기반으로 계산.
노트 안에서 빈 줄로 구분하면 별도 자막 엔트리로 분리된다.
"""
from __future__ import annotations

import os
from slidecast.parser.slide_model import SlideData


def _fmt_time(seconds: float) -> str:
    """초 → SRT 타임코드 (HH:MM:SS,mmm)"""
    ms = int(round(seconds * 1000))
    h  = ms // 3_600_000; ms %= 3_600_000
    m  = ms //    60_000; ms %=    60_000
    s  = ms //     1_000; ms %=     1_000
    return f"{h:02d}:{m:02d}:{s:02d},{ms:03d}"


def generate_srt_from_notes(
    slides: list[SlideData],
    total_duration: float,
    output_path: str,
    slide_durations: list[float] | None = None,
) -> str | None:
    """
    슬라이드 노트를 읽어 SRT 파일 생성.

    Args:
        slides: 파싱된 슬라이드 목록 (SlideData.notes 사용)
        total_duration: 총 재생 시간(초) — slide_durations 없을 때 균등 분배에 사용
        output_path: 출력할 .srt 파일 경로
        slide_durations: 슬라이드별 실제 길이(초) 리스트. 제공 시 정확한 싱크 적용.

    Returns:
        생성된 SRT 파일 경로, 노트가 하나도 없으면 None
    """
    # 노트가 있는 슬라이드만 수집
    note_slides = [s for s in slides if s.notes and s.notes.strip()]
    if not note_slides:
        return None

    n = len(slides)

    # 슬라이드별 시작 시간 계산
    if slide_durations and len(slide_durations) == n:
        # WAV 실제 길이 기반
        starts = []
        t = 0.0
        for d in slide_durations:
            starts.append(t)
            t += d
    else:
        # 균등 분배 fallback
        slide_dur = total_duration / n
        starts = [i * slide_dur for i in range(n)]
        slide_durations = [total_duration / n] * n

    entries: list[tuple[float, float, str]] = []  # (start, end, text)

    for slide in slides:
        if not slide.notes or not slide.notes.strip():
            continue

        idx = slide.slide_index
        slide_start = starts[idx]
        slide_duration = slide_durations[idx]
        slide_end = slide_start + slide_duration

        # 노트를 빈 줄 기준으로 분리 → 각각 별도 자막
        raw_blocks = [b.strip() for b in slide.notes.split("\n\n") if b.strip()]

        # 빈 줄 구분이 없으면 줄바꿈(엔터)으로 분리
        if len(raw_blocks) == 1:
            lines = [l.strip() for l in slide.notes.splitlines() if l.strip()]
            blocks = lines if len(lines) > 1 else raw_blocks
        else:
            blocks = raw_blocks

        if not blocks:
            continue

        # 슬라이드 시간을 블록 수로 균등 분배
        block_dur = slide_duration / len(blocks)
        for i, text in enumerate(blocks):
            start = slide_start + i * block_dur
            end   = start + block_dur - 0.1  # 0.1초 간격
            entries.append((start, end, text))

    if not entries:
        return None

    # SRT 파일 작성
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        for idx, (start, end, text) in enumerate(entries, 1):
            f.write(f"{idx}\n")
            f.write(f"{_fmt_time(start)} --> {_fmt_time(end)}\n")
            f.write(f"{text}\n\n")

    return output_path

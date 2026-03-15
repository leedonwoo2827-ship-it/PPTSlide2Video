"""PowerPoint COM API로 PPTX → 애니메이션 포함 MP4 내보내기.

PowerPoint가 직접 렌더링하므로 원본 애니메이션, 전환, 레이아웃이 모두 보존됨.
슬라이드별 WAV 길이를 AdvanceTime으로 설정하여 정확한 싱크.
"""
from __future__ import annotations

import os
import time


def export_pptx_to_video(
    pptx_path: str,
    output_path: str,
    slide_durations: list[float] | None = None,
    default_duration: float = 5.0,
    resolution: int = 1080,
) -> str:
    """
    PowerPoint COM으로 PPTX → MP4 (애니메이션 포함).

    Args:
        pptx_path: 원본 PPTX 경로
        output_path: 출력 MP4 경로
        slide_durations: 슬라이드별 시간(초) — WAV 길이 기반
        default_duration: 기본 슬라이드 시간 (slide_durations 없을 때)
        resolution: 세로 해상도 (1080 = Full HD)
    """
    import comtypes.client

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    ppt = comtypes.client.CreateObject("PowerPoint.Application")
    ppt.Visible = 1

    abs_pptx = os.path.abspath(pptx_path)
    abs_out = os.path.abspath(output_path)
    prs = ppt.Presentations.Open(abs_pptx, WithWindow=False)

    try:
        n_slides = prs.Slides.Count

        # 슬라이드별 타이밍 설정
        for i in range(1, n_slides + 1):
            slide = prs.Slides(i)
            slide.SlideShowTransition.AdvanceOnTime = True
            if slide_durations and (i - 1) < len(slide_durations):
                # AdvanceTime은 정수(초) — 반올림
                slide.SlideShowTransition.AdvanceTime = int(round(slide_durations[i - 1]))
            else:
                slide.SlideShowTransition.AdvanceTime = int(default_duration)

        # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertRes, Quality)
        prs.CreateVideo(abs_out, True, int(default_duration), int(resolution), 85)

        # 비동기 → 완료 대기 (1=진행중, 3=완료, 2=실패)
        timeout = 2400  # 최대 40분
        elapsed = 0
        while prs.CreateVideoStatus == 1 and elapsed < timeout:
            time.sleep(2)
            elapsed += 2
            if elapsed % 30 == 0:
                print(f"  [PowerPoint export] {elapsed}s elapsed...")

        status = prs.CreateVideoStatus
        if status != 3:
            raise RuntimeError(f"PowerPoint video export failed (status={status})")

    finally:
        prs.Close()
        ppt.Quit()

    return output_path


async def render_all_slides(
    html_paths: list[str],
    slides_meta: list,
    temp_dir: str,
    hold_seconds: float = 1.0,
    per_slide_durations: list[float] | None = None,
    pptx_path: str | None = None,
) -> list[str]:
    """
    PowerPoint 네이티브 MP4 내보내기.
    Returns: [단일 MP4 경로] — concat 단계 없이 바로 사용.
    """
    if not pptx_path:
        raise RuntimeError("pptx_path is required for PowerPoint rendering")

    ppt_video = os.path.join(temp_dir, "ppt_native.mp4")
    export_pptx_to_video(
        pptx_path=pptx_path,
        output_path=ppt_video,
        slide_durations=per_slide_durations,
        default_duration=hold_seconds + 3.0,
    )

    # 단일 비디오 반환 — pipeline의 concat 단계에서 그대로 통과
    return [ppt_video]

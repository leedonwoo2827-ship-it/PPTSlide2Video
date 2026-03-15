"""PowerPoint 네이티브 PPTX → 애니메이션 포함 MP4 내보내기.

Windows: COM API (comtypes)
macOS: AppleScript (osascript)

PowerPoint가 직접 렌더링하므로 원본 애니메이션, 전환, 레이아웃이 모두 보존됨.
슬라이드별 WAV 길이를 AdvanceTime으로 설정하여 정확한 싱크.
"""
from __future__ import annotations

import os
import platform
import subprocess
import time


def export_pptx_to_video(
    pptx_path: str,
    output_path: str,
    slide_durations: list[float] | None = None,
    default_duration: float = 5.0,
    resolution: int = 1080,
) -> str:
    """
    PowerPoint로 PPTX → MP4 (애니메이션 포함).
    OS에 따라 Windows COM 또는 macOS AppleScript 자동 선택.
    """
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    system = platform.system()
    if system == "Windows":
        return _export_windows(pptx_path, output_path, slide_durations, default_duration, resolution)
    elif system == "Darwin":
        return _export_macos(pptx_path, output_path, slide_durations, default_duration, resolution)
    else:
        raise RuntimeError(
            f"지원하지 않는 OS입니다: {system}\n"
            "Windows 또는 macOS에서 PowerPoint가 설치된 환경이 필요합니다."
        )


# ── Windows: COM API ─────────────────────────────────────────────────────

def _export_windows(
    pptx_path: str,
    output_path: str,
    slide_durations: list[float] | None,
    default_duration: float,
    resolution: int,
) -> str:
    import comtypes.client

    ppt = comtypes.client.CreateObject("PowerPoint.Application")
    ppt.Visible = 1

    abs_pptx = os.path.abspath(pptx_path)
    abs_out = os.path.abspath(output_path)
    prs = ppt.Presentations.Open(abs_pptx, WithWindow=False)

    try:
        n_slides = prs.Slides.Count

        for i in range(1, n_slides + 1):
            slide = prs.Slides(i)
            slide.SlideShowTransition.AdvanceOnTime = True
            if slide_durations and (i - 1) < len(slide_durations):
                # WAV 길이 + 1초 여유 (음성 끝난 뒤 잠깐 머문 후 전환)
                import math
                slide.SlideShowTransition.AdvanceTime = math.ceil(slide_durations[i - 1]) + 1
            else:
                slide.SlideShowTransition.AdvanceTime = int(default_duration)

        # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration, VertRes, Quality)
        prs.CreateVideo(abs_out, True, int(default_duration), int(resolution), 85)

        timeout = 2400  # 최대 40분
        elapsed = 0
        while prs.CreateVideoStatus == 1 and elapsed < timeout:
            time.sleep(2)
            elapsed += 2
            if elapsed % 30 == 0:
                print(f"  [PowerPoint export] {elapsed}s elapsed...", flush=True)

        status = prs.CreateVideoStatus
        if status != 3:
            raise RuntimeError(f"PowerPoint video export failed (status={status})")

    finally:
        prs.Close()
        ppt.Quit()

    return output_path


# ── macOS: AppleScript ───────────────────────────────────────────────────

def _export_macos(
    pptx_path: str,
    output_path: str,
    slide_durations: list[float] | None,
    default_duration: float,
    resolution: int,
) -> str:
    abs_pptx = os.path.abspath(pptx_path)
    abs_out = os.path.abspath(output_path)

    # 출력 확장자를 .mp4로 확보
    if not abs_out.lower().endswith(".mp4"):
        abs_out += ".mp4"

    # 슬라이드별 타이밍 설정 + 동영상 내보내기를 하나의 AppleScript로 실행
    slide_timing_lines = ""
    if slide_durations:
        timing_parts = []
        for idx, dur in enumerate(slide_durations):
            slide_num = idx + 1
            import math
            secs = math.ceil(dur) + 1  # WAV 길이 + 1초 여유
            timing_parts.append(
                f'set theTransition to slide transition of slide {slide_num} of thePresentation\n'
                f'set advance on time of theTransition to true\n'
                f'set advance time of theTransition to {secs}'
            )
        slide_timing_lines = "\n".join(timing_parts)
    else:
        slide_timing_lines = (
            f'set slideCount to count of slides of thePresentation\n'
            f'repeat with i from 1 to slideCount\n'
            f'  set theTransition to slide transition of slide i of thePresentation\n'
            f'  set advance on time of theTransition to true\n'
            f'  set advance time of theTransition to {int(default_duration)}\n'
            f'end repeat'
        )

    applescript = f'''
tell application "Microsoft PowerPoint"
    activate
    open POSIX file "{abs_pptx}"
    delay 3
    set thePresentation to active presentation

    -- 슬라이드별 타이밍 설정
    {slide_timing_lines}

    -- MP4로 내보내기
    save thePresentation in POSIX file "{abs_out}" as save as movie

    close thePresentation saving no
end tell
'''

    print("  [PowerPoint export] AppleScript로 내보내기 시작...", flush=True)

    result = subprocess.run(
        ["osascript", "-e", applescript],
        capture_output=True,
        text=True,
        timeout=2400,  # 최대 40분
    )

    if result.returncode != 0:
        error_msg = result.stderr.strip() if result.stderr else "Unknown error"
        raise RuntimeError(f"PowerPoint AppleScript export failed:\n{error_msg}")

    # PowerPoint for Mac은 내보내기 후 파일이 생성될 때까지 대기
    timeout = 2400
    elapsed = 0
    while not os.path.exists(abs_out) and elapsed < timeout:
        time.sleep(2)
        elapsed += 2
        if elapsed % 30 == 0:
            print(f"  [PowerPoint export] {elapsed}s elapsed... 파일 생성 대기 중", flush=True)

    if not os.path.exists(abs_out):
        raise RuntimeError(f"PowerPoint export file not found: {abs_out}")

    # 파일 크기가 증가를 멈출 때까지 대기 (쓰기 완료 확인)
    prev_size = -1
    while elapsed < timeout:
        curr_size = os.path.getsize(abs_out)
        if curr_size > 0 and curr_size == prev_size:
            break
        prev_size = curr_size
        time.sleep(2)
        elapsed += 2
        if elapsed % 30 == 0:
            print(f"  [PowerPoint export] {elapsed}s elapsed... ({curr_size // 1024 // 1024}MB)", flush=True)

    print(f"  [PowerPoint export] 완료! ({os.path.getsize(abs_out) // 1024 // 1024}MB)", flush=True)
    return abs_out


# ── 공통 진입점 ──────────────────────────────────────────────────────────

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

    return [ppt_video]

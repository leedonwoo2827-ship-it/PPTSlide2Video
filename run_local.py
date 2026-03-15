"""
로컬에서 직접 실행하는 PPTX → MP4 변환 스크립트.
PowerPoint COM 내보내기 — 타임아웃 없이 완료까지 대기.

사용법:
    python run_local.py                          (아래 설정값 사용)
    python run_local.py 발표자료.pptx             (같은 폴더의 WAV 자동 탐색)
    python run_local.py 발표자료.pptx D:\음성폴더  (음성 폴더 지정)
"""
import asyncio
import os
import sys
import time

# ── 기본 설정 (인자 없이 실행할 때 사용) ──────────────────────────────────
DEFAULT_PPTX = r""
DEFAULT_MEDIA = r""
DEFAULT_OUTPUT = r""
SUBTITLE_MODE = "hard"  # "hard" = 영상에 자막 합성, "soft" = 자막 트랙
# ─────────────────────────────────────────────────────────────────────────

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from slidecast.pipeline import run_pipeline


def resolve_paths(args):
    """명령줄 인자 또는 기본값에서 경로 결정."""
    if len(args) >= 2:
        pptx = os.path.abspath(args[1])
        media = os.path.abspath(args[2]) if len(args) >= 3 else os.path.dirname(pptx)
        if len(args) >= 4:
            output = os.path.abspath(args[3])
        else:
            base = os.path.splitext(pptx)[0]
            output = base + "_output.mp4"
    elif DEFAULT_PPTX:
        pptx = DEFAULT_PPTX
        media = DEFAULT_MEDIA or os.path.dirname(pptx)
        output = DEFAULT_OUTPUT or os.path.splitext(pptx)[0] + "_output.mp4"
    else:
        print("사용법:")
        print("  python run_local.py 발표자료.pptx")
        print("  python run_local.py 발표자료.pptx D:\\음성폴더")
        print("  python run_local.py 발표자료.pptx D:\\음성폴더 D:\\결과.mp4")
        print()
        print("또는 스크립트 상단의 DEFAULT_PPTX 등을 직접 수정하세요.")
        sys.exit(1)

    return pptx, media, output


async def main():
    pptx, media, output = resolve_paths(sys.argv)

    start = time.time()
    print(f"{'='*60}")
    print(f"  PPTSlide2Video — PPTX → MP4 변환")
    print(f"{'='*60}")
    print(f"  PPTX : {pptx}")
    print(f"  음성 : {media}")
    print(f"  출력 : {output}")
    print(f"  자막 : {SUBTITLE_MODE}")
    print(f"{'='*60}")
    print()
    print("  PowerPoint 네이티브 내보내기 중... (20~30분 소요될 수 있습니다)")
    print()

    try:
        async def on_progress(step, pct):
            print(f"  [{pct:3d}%] {step}", flush=True)

        result = await run_pipeline(
            pptx_path=pptx,
            media_folder=media,
            output_path=output,
            subtitle_mode=SUBTITLE_MODE,
            progress_callback=on_progress,
        )
        elapsed = time.time() - start
        size_mb = os.path.getsize(result) / (1024 * 1024)
        print()
        print(f"  [완료] {result}")
        print(f"  파일 크기: {size_mb:.1f} MB")
        print(f"  소요 시간: {elapsed:.0f}초 ({elapsed/60:.1f}분)")
    except Exception as e:
        elapsed = time.time() - start
        print(f"\n  [오류] {type(e).__name__}: {e}")
        print(f"  소요 시간: {elapsed:.0f}초")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    asyncio.run(main())

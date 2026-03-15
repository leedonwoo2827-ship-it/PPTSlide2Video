"""Quick end-to-end pipeline test."""
import sys, os, asyncio
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
OUTPUT   = os.path.join(os.path.dirname(__file__), "output", "result.mp4")

async def main():
    from slidecast.pipeline import run_pipeline

    os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

    def progress(step, pct):
        print(f"  [{pct:3d}%] {step}")

    async def async_progress(step, pct):
        progress(step, pct)

    print("Starting pipeline test...")
    result = await run_pipeline(
        pptx_path=os.path.join(FIXTURES, "sample.pptx"),
        media_folder=FIXTURES,
        output_path=OUTPUT,
        subtitle_mode="soft",
        slide_hold_seconds=0.5,
        progress_callback=async_progress,
    )
    print(f"\nResult: {result}")
    if os.path.exists(OUTPUT):
        size_mb = os.path.getsize(OUTPUT) / (1024*1024)
        print(f"File size: {size_mb:.2f} MB")
        print("SUCCESS!")
    else:
        print("FAILED - output file not found")

if __name__ == "__main__":
    asyncio.run(main())

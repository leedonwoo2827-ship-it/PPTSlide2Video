"""
PPTSlide2Video MCP Server
========================
Claude Desktop plugin that converts PPTX → animated MP4 (with audio + subtitles).

Usage (Claude Desktop claude_desktop_config.json):
{
  "mcpServers": {
    "pptslide2video": {
      "command": "python",
      "args": ["D:/00work/260312-PPT2slidedeck/server.py"]
    }
  }
}
"""
from __future__ import annotations

import asyncio
import sys
import os

# Ensure project root is on sys.path when run directly
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

from slidecast.pipeline import run_pipeline

server = Server("pptslide2video")


@server.list_tools()
async def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="convert_pptx_to_video",
            description=(
                "Convert a PowerPoint (.pptx) presentation into an animated MP4 video. "
                "Each slide gets typing and fade-in effects. "
                "Automatically adds MP3 audio and SRT/VTT subtitles from the specified folder. "
                "The final video length is determined by the MP3 audio duration."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "pptx_path": {
                        "type": "string",
                        "description": "Absolute path to the .pptx file",
                    },
                    "media_folder": {
                        "type": "string",
                        "description": (
                            "Folder containing the .mp3 audio file and optional "
                            ".srt or .vtt subtitle file. "
                            "The first .mp3 and first subtitle found are used."
                        ),
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Absolute path for the output .mp4 file",
                    },
                    "subtitle_mode": {
                        "type": "string",
                        "enum": ["soft", "hard"],
                        "default": "hard",
                        "description": (
                            "soft = subtitle as a separate selectable track; "
                            "hard = subtitles burned into the video pixels (universal compatibility)"
                        ),
                    },
                    "slide_hold_seconds": {
                        "type": "number",
                        "default": 1.0,
                        "description": "Seconds to hold each slide after all animations finish",
                    },
                },
                "required": ["pptx_path", "media_folder", "output_path"],
            },
        ),
        types.Tool(
            name="check_dependencies",
            description="Check that all required system dependencies (FFmpeg, Playwright) are installed.",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    if name == "check_dependencies":
        return await _check_dependencies()

    if name != "convert_pptx_to_video":
        return [types.TextContent(type="text", text=f"Unknown tool: {name}")]

    pptx_path = arguments.get("pptx_path", "")
    media_folder = arguments.get("media_folder", "")
    output_path = arguments.get("output_path", "")
    subtitle_mode = arguments.get("subtitle_mode", "hard")
    slide_hold_seconds = float(arguments.get("slide_hold_seconds", 1.0))

    # Basic validation
    if not os.path.isfile(pptx_path):
        return [types.TextContent(type="text",
            text=f"Error: PPTX file not found: {pptx_path}")]
    if not os.path.isdir(media_folder):
        return [types.TextContent(type="text",
            text=f"Error: Media folder not found: {media_folder}")]

    progress_log: list[str] = []

    async def on_progress(step: str, pct: int):
        progress_log.append(f"[{pct:3d}%] {step}")

    try:
        result_path = await run_pipeline(
            pptx_path=pptx_path,
            media_folder=media_folder,
            output_path=output_path,
            subtitle_mode=subtitle_mode,
            slide_hold_seconds=slide_hold_seconds,
            progress_callback=on_progress,
        )
        progress_summary = "\n".join(progress_log)
        return [types.TextContent(
            type="text",
            text=(
                f"✅ 변환 완료!\n"
                f"출력 파일: {result_path}\n\n"
                f"진행 로그:\n{progress_summary}"
            ),
        )]
    except Exception as e:
        import traceback
        return [types.TextContent(
            type="text",
            text=f"❌ 오류 발생:\n{str(e)}\n\n{traceback.format_exc()}",
        )]


async def _check_dependencies() -> list[types.TextContent]:
    import subprocess
    results = []

    # Check FFmpeg
    try:
        r = subprocess.run(["ffmpeg", "-version"], capture_output=True, text=True)
        line = r.stdout.split("\n")[0] if r.stdout else "unknown"
        results.append(f"✅ FFmpeg: {line}")
    except FileNotFoundError:
        results.append("❌ FFmpeg: NOT FOUND — install from https://ffmpeg.org/download.html")

    # Check ffprobe
    try:
        r = subprocess.run(["ffprobe", "-version"], capture_output=True, text=True)
        results.append("✅ ffprobe: OK")
    except FileNotFoundError:
        results.append("❌ ffprobe: NOT FOUND (comes with FFmpeg)")

    # Check Playwright + Chromium
    try:
        from playwright.async_api import async_playwright
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            version = browser.version
            await browser.close()
        results.append(f"✅ Playwright + Chromium: {version}")
    except Exception as e:
        results.append(
            f"❌ Playwright/Chromium: {e}\n"
            "   Run: playwright install chromium"
        )

    # Check python-pptx
    try:
        import pptx
        results.append(f"✅ python-pptx: {pptx.__version__}")
    except ImportError:
        results.append("❌ python-pptx: NOT FOUND — pip install python-pptx")

    return [types.TextContent(type="text", text="\n".join(results))]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())

"""
미리보기 서버 - PPTX를 HTML로 변환 후 브라우저에서 확인
사용법: python preview.py [pptx파일경로]
"""
import sys, os, asyncio, http.server, threading, webbrowser, tempfile, shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

PORT = 8765


def make_index_html(html_paths: list[str], output_dir: str) -> str:
    slide_links = "\n".join(
        f'<li><a href="html/slide_{i:03d}.html" target="preview">'
        f'슬라이드 {i+1}</a></li>'
        for i, _ in enumerate(html_paths)
    )
    first = f"html/slide_000.html" if html_paths else ""
    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>PPT2SlideDeck 미리보기</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', sans-serif; display: flex; height: 100vh; background: #1e1e2e; color: #cdd6f4; }}
  .sidebar {{ width: 200px; min-width: 160px; background: #181825; padding: 16px; overflow-y: auto; flex-shrink: 0; }}
  .sidebar h2 {{ font-size: 14px; color: #89b4fa; margin-bottom: 12px; letter-spacing: 1px; text-transform: uppercase; }}
  .sidebar ul {{ list-style: none; }}
  .sidebar li {{ margin-bottom: 8px; }}
  .sidebar a {{ color: #cdd6f4; text-decoration: none; font-size: 13px; display: block; padding: 6px 10px; border-radius: 6px; transition: background .15s; }}
  .sidebar a:hover {{ background: #313244; color: #89b4fa; }}
  .preview-area {{ flex: 1; display: flex; align-items: center; justify-content: center; padding: 20px; }}
  iframe {{ border: none; border-radius: 8px; box-shadow: 0 8px 32px rgba(0,0,0,.5); background: white; }}
  .badge {{ background: #a6e3a1; color: #1e1e2e; font-size: 11px; padding: 2px 8px; border-radius: 20px; margin-left: 6px; }}
</style>
</head>
<body>
<div class="sidebar">
  <h2>슬라이드 <span class="badge">{len(html_paths)}</span></h2>
  <ul>{slide_links}</ul>
</div>
<div class="preview-area">
  <iframe name="preview" src="{first}" width="960" height="540"></iframe>
</div>
</body>
</html>"""
    idx_path = os.path.join(output_dir, "index.html")
    with open(idx_path, "w", encoding="utf-8") as f:
        f.write(html)
    return idx_path


async def generate_slides(pptx_path: str) -> tuple[list[str], str]:
    from slidecast.parser.pptx_parser import parse_pptx
    from slidecast.generator.html_generator import generate_html_slides

    temp_dir = tempfile.mkdtemp(prefix="ppt2slide_preview_")
    print(f"  Parsing PPTX...")
    presentation = parse_pptx(pptx_path, temp_dir)
    print(f"  Generating HTML ({len(presentation.slides)} slides)...")
    html_paths = generate_html_slides(presentation, temp_dir)
    return html_paths, temp_dir


def serve(directory: str):
    class Handler(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=directory, **kwargs)
        def log_message(self, *_):
            pass  # suppress request logs

    server = http.server.HTTPServer(("localhost", PORT), Handler)
    server.serve_forever()


def main():
    if len(sys.argv) < 2:
        # Use sample if no argument given
        pptx_path = os.path.join(
            os.path.dirname(__file__), "tests", "fixtures", "sample.pptx"
        )
        if not os.path.exists(pptx_path):
            print("Usage: python preview.py <path_to.pptx>")
            sys.exit(1)
    else:
        pptx_path = sys.argv[1]

    if not os.path.isfile(pptx_path):
        print(f"File not found: {pptx_path}")
        sys.exit(1)

    print(f"Loading: {pptx_path}")
    html_paths, temp_dir = asyncio.run(generate_slides(pptx_path))

    # Build index page in the temp dir
    make_index_html(html_paths, temp_dir)

    # Start HTTP server in background thread
    t = threading.Thread(target=serve, args=(temp_dir,), daemon=True)
    t.start()

    url = f"http://localhost:{PORT}/index.html"
    print(f"\n  Preview ready at: {url}")
    print("  Press Ctrl+C to stop.\n")
    webbrowser.open(url)

    try:
        while True:
            pass
    except KeyboardInterrupt:
        print("\nStopping preview server...")
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()

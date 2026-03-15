"""HTML slide generator using Jinja2 templates."""
from __future__ import annotations

import json
import os
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape

from slidecast.generator.animation_mapper import get_gsap_from_vars
from slidecast.parser.slide_model import PresentationData, ShapeData

TEMPLATE_DIR = Path(__file__).parent / "templates"
STATIC_DIR = Path(__file__).parent / "static"
GSAP_LOCAL = STATIC_DIR / "gsap.min.js"


def _gsap_from_json(shape: ShapeData) -> str:
    """Jinja2 filter: returns JSON string of GSAP from-vars for a shape."""
    return json.dumps(get_gsap_from_vars(shape.animation_hint))


def generate_html_slides(presentation: PresentationData, temp_dir: str) -> list[str]:
    """Generate one HTML file per slide. Returns list of file paths."""
    html_dir = os.path.join(temp_dir, "html")
    os.makedirs(html_dir, exist_ok=True)

    # GSAP을 파일 경로 대신 소스 코드 인라인으로 삽입
    # (Chromium이 file:// cross-origin 차단하는 문제 해결)
    gsap_inline = _get_gsap_inline()
    # 한글 폰트를 @font-face로 직접 등록 (Playwright 헤드리스 환경에서도 작동)
    korean_font_css = _get_korean_font_css()

    env = Environment(
        loader=FileSystemLoader(str(TEMPLATE_DIR)),
        autoescape=select_autoescape(disabled_extensions=("j2",)),
    )
    env.filters["gsap_from_json"] = _gsap_from_json

    html_paths = []
    for slide in presentation.slides:
        out_path = os.path.join(html_dir, f"slide_{slide.slide_index:03d}.html")
        template = env.get_template("slide.html.j2")
        html = template.render(
            slide=slide,
            gsap_inline=gsap_inline,
            korean_font_css=korean_font_css,
        )
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)
        html_paths.append(out_path)

    return html_paths


def _get_korean_font_css() -> str:
    """
    Windows 시스템 폰트(맑은 고딕)를 @font-face로 등록하는 CSS 반환.
    Playwright 헤드리스 Chromium은 시스템 폰트를 못 찾을 수 있으므로
    file:// URL로 직접 지정한다.
    없으면 빈 문자열 반환 (CSS font-family fallback으로 처리).
    """
    candidates = [
        ("Malgun Gothic", r"C:/Windows/Fonts/malgun.ttf", r"C:/Windows/Fonts/malgunbd.ttf"),
        ("Nanum Gothic", r"C:/Windows/Fonts/NanumGothic.ttf", None),
    ]
    parts: list[str] = []
    for family, regular_path, bold_path in candidates:
        if os.path.isfile(regular_path):
            uri = Path(regular_path).as_uri()
            parts.append(
                f"@font-face {{\n"
                f"  font-family: '{family}';\n"
                f"  font-weight: normal;\n"
                f"  font-style: normal;\n"
                f"  src: url('{uri}') format('truetype');\n"
                f"}}"
            )
        if bold_path and os.path.isfile(bold_path):
            uri_b = Path(bold_path).as_uri()
            parts.append(
                f"@font-face {{\n"
                f"  font-family: '{family}';\n"
                f"  font-weight: bold;\n"
                f"  font-style: normal;\n"
                f"  src: url('{uri_b}') format('truetype');\n"
                f"}}"
            )
    return "\n".join(parts)


def _get_gsap_inline() -> str:
    """GSAP 소스 코드를 문자열로 반환 (없으면 다운로드)."""
    STATIC_DIR.mkdir(parents=True, exist_ok=True)
    gsap_dest = STATIC_DIR / "gsap.min.js"

    if not gsap_dest.exists():
        _download_gsap(gsap_dest)

    return gsap_dest.read_text(encoding="utf-8")


def _download_gsap(dest: Path) -> None:
    """Download GSAP minified from CDN."""
    import urllib.request
    url = "https://cdn.jsdelivr.net/npm/gsap@3/dist/gsap.min.js"
    try:
        urllib.request.urlretrieve(url, str(dest))
    except Exception as e:
        raise RuntimeError(
            f"Failed to download GSAP from CDN: {e}\n"
            "Please manually place gsap.min.js in "
            f"{dest.parent}"
        )

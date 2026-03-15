"""Dataclass models representing parsed PPTX data."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class RunData:
    text: str
    bold: bool = False
    italic: bool = False
    font_size_px: float = 16.0
    color: str = "#000000"


@dataclass
class ParagraphData:
    runs: list[RunData] = field(default_factory=list)
    alignment: str = "left"   # "left" | "center" | "right"

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class TableCellData:
    text: str = ""
    paragraphs: list[ParagraphData] = field(default_factory=list)
    fill_color: str | None = None
    col_span: int = 1
    row_span: int = 1


@dataclass
class ShapeData:
    shape_id: int
    shape_type: str          # "text_box" | "title" | "picture" | "auto_shape" | "group" | "table"
    name: str
    left_px: float
    top_px: float
    width_px: float
    height_px: float
    rotation: float = 0.0
    text: str | None = None
    paragraphs: list[ParagraphData] = field(default_factory=list)
    fill_color: str | None = None
    border_color: str | None = None
    border_width_px: float = 0.0
    image_path: str | None = None
    table_data: list[list[TableCellData]] | None = None  # rows → cells
    col_widths: list[float] | None = None  # column widths in px
    z_order: int = 0
    # "fade_in" | "fly_in_left" | "fly_in_right" | "zoom_in" | "type_in" | "none"
    animation_hint: str = "fade_in"
    delay: float = 0.0       # seconds, computed from z_order

    @property
    def css_id(self) -> str:
        return f"shape-{self.shape_id}"


@dataclass
class SlideData:
    slide_index: int
    width_px: float
    height_px: float
    background_color: str = "#FFFFFF"
    shapes: list[ShapeData] = field(default_factory=list)
    notes: str | None = None
    duration_seconds: float = 3.0


@dataclass
class PresentationData:
    slides: list[SlideData] = field(default_factory=list)
    source_path: str = ""
    temp_dir: str = ""

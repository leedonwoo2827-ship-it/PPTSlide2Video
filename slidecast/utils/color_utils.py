"""Color conversion utilities for PPTX → CSS."""
from __future__ import annotations


def rgb_to_hex(rgb) -> str:
    """Convert pptx RGBColor to CSS hex string."""
    if rgb is None:
        return "#000000"
    return f"#{rgb.red:02X}{rgb.green:02X}{rgb.blue:02X}"


def theme_color_to_hex(color) -> str | None:
    """Best-effort extraction of color from a pptx ColorFormat object."""
    try:
        if color.type is None:
            return None
        return rgb_to_hex(color.rgb)
    except Exception:
        return None

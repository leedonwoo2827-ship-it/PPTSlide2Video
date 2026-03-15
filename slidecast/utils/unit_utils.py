"""EMU (English Metric Units) to pixel conversion utilities."""

# 1 inch = 914400 EMU, 96 DPI → 1 px = 9525 EMU
EMU_PER_PX = 9525


def emu_to_px(emu: int | float) -> float:
    if emu is None:
        return 0.0
    return round(emu / EMU_PER_PX, 2)


def pt_to_px(pt: float) -> float:
    """Convert points to pixels (96 DPI: 1pt = 1.333px)."""
    if pt is None:
        return 16.0
    return round(pt * 96 / 72, 2)

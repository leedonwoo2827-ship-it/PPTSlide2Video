"""Maps ShapeData.animation_hint to GSAP tween configuration."""
from __future__ import annotations


def get_gsap_from_vars(hint: str) -> dict:
    """Returns the GSAP `from` vars dict for a given hint."""
    mapping = {
        "fade_in": {"opacity": 0, "duration": 0.6, "ease": "power2.out"},
        "fly_in_left": {"x": -150, "opacity": 0, "duration": 0.5, "ease": "power2.out"},
        "fly_in_right": {"x": 150, "opacity": 0, "duration": 0.5, "ease": "power2.out"},
        "fly_in_up": {"y": 60, "opacity": 0, "duration": 0.5, "ease": "power2.out"},
        "zoom_in": {"scale": 0.3, "opacity": 0, "duration": 0.5, "ease": "back.out(1.7)"},
        "type_in": {"opacity": 0, "duration": 0.4, "ease": "none"},  # handled separately
        "none": {},
    }
    return mapping.get(hint, mapping["fade_in"])


def needs_typing_effect(hint: str) -> bool:
    return hint == "type_in"

"""OOXML readers: XML in, unresolved source model out."""

from . import drawing, parts, shapes, source, text
from .parts import (
    read_presentation,
    read_slide,
    read_slide_layout,
    read_slide_master,
    read_theme,
)

__all__ = [
    "drawing",
    "parts",
    "read_presentation",
    "read_slide",
    "read_slide_layout",
    "read_slide_master",
    "read_theme",
    "shapes",
    "source",
    "text",
]

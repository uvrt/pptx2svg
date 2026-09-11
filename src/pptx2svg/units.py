"""OOXML unit system (ECMA-376 §20.1.2.1).

PPTX internal coordinates are EMU (English Metric Units).  A 16:9 slide is
9,144,000 x 5,143,500 EMU = 960 x 540 px at 96 DPI.

The TypeScript original uses branded types (``Emu``, ``Pt``, ``HundredthPt``) to keep
unit systems apart at compile time.  Python has no zero-cost equivalent, so the units
are documented in the names of the conversion helpers instead: anything called
``*_emu`` is EMU, ``*_pt`` is points, ``*_px`` is CSS pixels at 96 DPI.
"""

from __future__ import annotations

#: 1 inch = 914,400 EMU
EMU_PER_INCH = 914400
#: 1 pt = 12,700 EMU
EMU_PER_POINT = 12700
DEFAULT_DPI = 96
DEFAULT_OUTPUT_WIDTH = 960

#: Rotation angle unit: 1/60,000 degrees (ECMA-376 §20.1.10.3)
ROTATION_UNIT = 60000

#: OOXML percentages are stored as 1/1000 of a percent (100% == 100000).
PERCENT_UNIT = 100000

PX_PER_PT = 96 / 72


def emu_to_px(emu: float, dpi: int = DEFAULT_DPI) -> float:
    return (emu / EMU_PER_INCH) * dpi


def emu_to_pt(emu: float) -> float:
    return emu / EMU_PER_POINT


def px_to_emu(px: float, dpi: int = DEFAULT_DPI) -> float:
    return (px / dpi) * EMU_PER_INCH


def rotation_to_degrees(rotation: float) -> float:
    """1/60,000 degrees -> degrees."""
    return rotation / ROTATION_UNIT


def hundredth_pt_to_pt(value: float) -> float:
    """1/100 points -> points."""
    return value / 100


def percent_to_ratio(value: float | None) -> float:
    """OOXML 1/1000-percent -> 0..1 ratio."""
    return 0.0 if value is None else value / PERCENT_UNIT

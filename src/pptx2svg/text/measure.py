"""Text measurement: how wide is this string, and how tall is a line of it?

The default measurer uses the static tables in :mod:`pptx2svg.text.metrics`, so the same
deck lays out identically on every machine.  When a character has no entry -- and for any
font we have no table for at all -- it falls back to a per-category width ratio, which is
crude but keeps wrapping sane for exotic fonts.

:class:`FontToolsTextMeasurer` is the opt-in alternative: point it at real font files and
it reads their true advance widths with fontTools.  Use it when fidelity to a specific
machine's fonts matters more than reproducibility.
"""

from __future__ import annotations

from typing import Protocol

from ..units import PX_PER_PT
from .fontmap import metrics_for
from .metrics import FontMetrics

#: Width as a fraction of the font size, for characters with no metrics entry.
NARROW_RATIO = 0.3
NORMAL_RATIO = 0.6
WIDE_RATIO = 1.0

BOLD_FACTOR = 1.05
DEFAULT_LINE_HEIGHT_RATIO = 1.2
DEFAULT_ASCENDER_RATIO = 1.0

#: Characters noticeably narrower than the Latin average.
_NARROW_CHARS = frozenset(" !,.:;ijl1|'()[]{}")


def is_cjk(code_point: int) -> bool:
    """CJK ranges that are full-width and break between any two characters."""
    return (
        0x3000 <= code_point <= 0x9FFF  # CJK symbols, kana, unified ideographs
        or 0xF900 <= code_point <= 0xFAFF  # compatibility ideographs
        or 0xFF01 <= code_point <= 0xFF60  # full-width forms
        or 0x20000 <= code_point <= 0x2A6DF  # extension B
    )


class TextMeasurer(Protocol):
    def measure_text_width(
        self,
        text: str,
        font_size_pt: float,
        bold: bool = False,
        font_family: str | None = None,
        font_family_ea: str | None = None,
    ) -> float:
        """Width in CSS pixels."""

    def line_height_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        """``(ascender + |descender|) / unitsPerEm``."""

    def ascender_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        """``ascender / unitsPerEm`` -- the first line's baseline offset."""


class DefaultTextMeasurer:
    """Measures against the bundled metric-compatible tables."""

    def measure_text_width(
        self,
        text: str,
        font_size_pt: float,
        bold: bool = False,
        font_family: str | None = None,
        font_family_ea: str | None = None,
    ) -> float:
        base_size_px = font_size_pt * PX_PER_PT
        latin_metrics = metrics_for(font_family)
        ea_metrics = metrics_for(font_family_ea)
        total = 0.0

        for char in text:
            code_point = ord(char)
            east_asian = is_cjk(code_point)
            metrics = ea_metrics if east_asian and ea_metrics else latin_metrics
            if metrics is not None:
                width = _measure_with_metrics(char, code_point, base_size_px, metrics)
            else:
                width = base_size_px * _heuristic_ratio(char, code_point)
            if bold and not east_asian:
                width *= BOLD_FACTOR
            total += width
        return total

    def line_height_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        metrics = metrics_for(font_family) or metrics_for(font_family_ea)
        if metrics is None:
            return DEFAULT_LINE_HEIGHT_RATIO
        return (metrics.ascender + abs(metrics.descender)) / metrics.units_per_em

    def ascender_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        metrics = metrics_for(font_family) or metrics_for(font_family_ea)
        if metrics is None:
            return DEFAULT_ASCENDER_RATIO
        return metrics.ascender / metrics.units_per_em


def _measure_with_metrics(
    char: str, code_point: int, base_size_px: float, metrics: FontMetrics
) -> float:
    width = metrics.widths.get(char)
    if width is not None:
        return (width / metrics.units_per_em) * base_size_px
    if is_cjk(code_point):
        return (metrics.cjk_width / metrics.units_per_em) * base_size_px
    return (metrics.default_width / metrics.units_per_em) * base_size_px


def _heuristic_ratio(char: str, code_point: int) -> float:
    if is_cjk(code_point):
        return WIDE_RATIO
    if char in _NARROW_CHARS:
        return NARROW_RATIO
    return NORMAL_RATIO


class FontToolsTextMeasurer:
    """Measures with real font files via fontTools.

    ``font_paths`` maps a font family name to a ``.ttf``/``.otf``/``.ttc`` path.  Faces
    are opened lazily and cached; a family with no entry falls back to the static tables,
    so a partial map is fine.
    """

    def __init__(
        self,
        font_paths: dict[str, str],
        fallback: TextMeasurer | None = None,
    ) -> None:
        self._font_paths = font_paths
        self._fallback = fallback or DefaultTextMeasurer()
        self._cache: dict[str, object] = {}

    def _face(self, font_family: str | None):
        if not font_family:
            return None
        if font_family in self._cache:
            return self._cache[font_family]
        path = self._font_paths.get(font_family)
        if path is None:
            self._cache[font_family] = None
            return None
        try:
            from fontTools.ttLib import TTFont

            font = TTFont(path, fontNumber=0, lazy=True)
        except Exception:  # unreadable or unsupported font file
            font = None
        self._cache[font_family] = font
        return font

    def measure_text_width(
        self,
        text: str,
        font_size_pt: float,
        bold: bool = False,
        font_family: str | None = None,
        font_family_ea: str | None = None,
    ) -> float:
        latin = self._face(font_family)
        east_asian = self._face(font_family_ea)
        if latin is None and east_asian is None:
            return self._fallback.measure_text_width(
                text, font_size_pt, bold, font_family, font_family_ea
            )

        base_size_px = font_size_pt * PX_PER_PT
        total = 0.0
        for char in text:
            font = east_asian if is_cjk(ord(char)) and east_asian is not None else latin
            advance = _advance_width(font, char)
            if advance is None:
                total += self._fallback.measure_text_width(
                    char, font_size_pt, bold, font_family, font_family_ea
                )
                continue
            units_per_em = font["head"].unitsPerEm  # type: ignore[index]
            total += (advance / units_per_em) * base_size_px
        return total

    def line_height_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        font = self._face(font_family) or self._face(font_family_ea)
        if font is None:
            return self._fallback.line_height_ratio(font_family, font_family_ea)
        hhea = font["hhea"]  # type: ignore[index]
        units_per_em = font["head"].unitsPerEm  # type: ignore[index]
        return (hhea.ascender + abs(hhea.descender)) / units_per_em

    def ascender_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        font = self._face(font_family) or self._face(font_family_ea)
        if font is None:
            return self._fallback.ascender_ratio(font_family, font_family_ea)
        hhea = font["hhea"]  # type: ignore[index]
        units_per_em = font["head"].unitsPerEm  # type: ignore[index]
        return hhea.ascender / units_per_em


def _advance_width(font, char: str) -> float | None:
    if font is None:
        return None
    try:
        cmap = font.getBestCmap()
        glyph_name = cmap.get(ord(char))
        if glyph_name is None:
            return None
        return font["hmtx"][glyph_name][0]
    except Exception:
        return None

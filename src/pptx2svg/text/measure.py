"""Text measurement: how wide is this string, and how tall is a line of it?

The default measurer uses the static tables in :mod:`pptx2svg.text.metrics`, so the same
deck lays out identically on every machine.  When a character has no entry -- and for any
font we have no table for at all -- it falls back to a per-category width ratio, which is
crude but keeps wrapping sane for exotic fonts.

A string is the sum of its advances **and the face's ``kern`` pairs across the joins**,
which PowerPoint charges and this used not to; see :mod:`pptx2svg.text.kerning`.  The one
caller that must not is the chart engine, which lays its text out unkerned and draws it
kerned -- measured, and recorded on :func:`pptx2svg.resolve.chart.text_width`.

:class:`FontToolsTextMeasurer` is the opt-in alternative: point it at real font files and
it reads their true advance widths with fontTools.  Use it when fidelity to a specific
machine's fonts matters more than reproducibility.
"""

from __future__ import annotations

from typing import Mapping, Protocol

from ..units import PX_PER_PT
from .fontmap import family_key, metrics_for
from .metrics import FontMetrics

#: Width as a fraction of the font size, for characters with no metrics entry.
NARROW_RATIO = 0.3
NORMAL_RATIO = 0.6
WIDE_RATIO = 1.0

#: Fallback widening for bold text when the face has no bold table of its own.
#:
#: Only reached for faces we do not ship.  Where we *do* ship the bold cut the real
#: advance widths are used instead, because a single factor is not close enough to be
#: worth the simplicity: measured over representative headings the true ratio is 1.000
#: for Cousine (monospace bold is the same width), 1.023 for Carlito, 1.049 for Tinos,
#: 1.056 for Arimo, 1.089 for Caladea and 1.119 for Noto Sans JP.  1.05 is the middle of
#: that spread, which is another way of saying it is wrong for every one of them.
BOLD_FACTOR = 1.05

#: What PowerPoint's synthetic bold adds to an East Asian advance, in points per glyph.
#:
#: A CJK face with no bold cut -- which is all four MS Japanese faces; ``msgothic.ttc``
#: and ``msmincho.ttc`` contain none -- is emboldened by PowerPoint rather than left
#: light, and the emboldening widens the advance.  This is **not a ratio**, which is the
#: surprising part and the reason it is a constant here.  Read off the pen origins in
#: PowerPoint's own PDF export of ``sample.pptx``, which sets the same Japanese text bold
#: at three sizes:
#:
#:     24 pt   upright 24.000   bold 24.1248    (+0.1248)
#:     28 pt   upright 28.000   bold 28.1260    (+0.1260)
#:     32 pt   upright 32.000   bold 32.1248    (+0.1248)
#:
#: The increment does not scale with the font size, so it cannot be expressed in ems; a
#: 1/256 em model would predict +0.094 at 24 pt and the measurement says +0.125.  The
#: upright numbers are exactly the point size because MS Gothic is full-width monospaced
#: -- every CJK glyph advances 1 em -- which makes the difference unusually easy to read.
#:
#: It is small: 0.4% at 32 pt.  It is here because it is measured and free, not because
#: it rescues a layout on its own.
EAST_ASIAN_SYNTHETIC_BOLD_PT = 0.125

#: PowerPoint's line box for single spacing, as a multiple of the font size.
#:
#: This is a constant, not a property of the typeface, and that is genuinely surprising:
#: a line box is normally the font's own ascent + descent (+ line gap).  PowerPoint
#: ignores all three.  It was measured by exporting probe decks through PowerPoint and
#: reading the rendered line advance back off the raster -- Arial, Calibri, Times New
#: Roman, Courier New, Aptos, Aptos Display, Lato, Raleway, MS Gothic, Meiryo and Noto
#: Sans JP, with Latin and with Japanese text, at 14 pt and 28 pt.  Every one came back
#: at 1.2x the font size, although their real ascent+descent ranges from 1.00 em (MS
#: Gothic) to 1.45 em (Noto Sans JP).
#:
#: ``a:lnSpc`` percentages multiply *this*, not the face's metrics: 150% measured 1.8 em
#: and 90% measured 1.08 em.  ``a:spcPts`` is a literal point size and ignores it.
DEFAULT_LINE_HEIGHT_RATIO = 1.2

#: First-baseline offset when the font is unknown: the line box less a 0.2 em descent.
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
        """Width in CSS pixels, with the face's ``kern`` feature applied."""

    def kern_between(
        self,
        left: str,
        right: str,
        font_size_pt: float,
        bold: bool = False,
        font_family: str | None = None,
        font_family_ea: str | None = None,
    ) -> float:
        """The ``kern`` adjustment across the join of two adjacent strings, in pixels.

        Usually negative, and zero for every pair that does not kern.  It exists because
        the caller that most needs kerning measures the string in pieces:
        :mod:`pptx2svg.text.wrap` gives every CJK character its own token, so measuring
        each token on its own would drop *every* Japanese kern pair -- the ones this was
        written for.  ``left`` and ``right`` are whole strings and only the join between
        them is charged, so a caller can add it as it appends.
        """

    def line_height_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        """``(ascender + |descender|) / unitsPerEm``."""

    def ascender_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        """``ascender / unitsPerEm`` -- the first line's baseline offset."""


class DefaultTextMeasurer:
    """Measures against the bundled metric-compatible tables.

    ``extra_metrics`` maps a normalised family key (see
    :func:`pptx2svg.text.fontmap.family_key`) to a table that wins over the static ones.  It
    exists for fonts a deck carries itself: :mod:`pptx2svg.fonts.embedded` builds a table
    from the face it extracted from ``<p:embeddedFontLst>`` and passes it here, so the
    layout is computed from the same file the rasteriser will draw with.  Injecting a
    table rather than adding a measurer is deliberate -- every rule in this class about
    bold cuts, synthetic emboldening and East Asian advances then applies to an embedded
    face unchanged, instead of being reimplemented beside it.

    Empty by default, so a caller who constructs one by hand gets exactly the old
    behaviour.
    """

    def __init__(self, extra_metrics: "Mapping[str, FontMetrics] | None" = None) -> None:
        self._extra = dict(extra_metrics) if extra_metrics else {}

    def _metrics(self, font_family: str | None) -> FontMetrics | None:
        if self._extra and font_family:
            embedded = self._extra.get(family_key(font_family))
            if embedded is not None:
                return embedded
        return metrics_for(font_family)

    def measure_text_width(
        self,
        text: str,
        font_size_pt: float,
        bold: bool = False,
        font_family: str | None = None,
        font_family_ea: str | None = None,
    ) -> float:
        base_size_px = font_size_pt * PX_PER_PT
        latin_metrics = self._metrics(font_family)
        ea_metrics = self._metrics(font_family_ea)
        # A face with no bold design of its own gets PowerPoint's synthetic emboldening,
        # and that widens every East Asian advance by a fixed amount.  Asking once per
        # call rather than once per character: the tables are large and the answer is a
        # property of the face.
        synthetic_bold_px = (
            EAST_ASIAN_SYNTHETIC_BOLD_PT * PX_PER_PT if bold else 0.0
        )
        total = 0.0
        #: The character before this one and the table it was measured from, for the
        #: kern pair between them.  A pair whose two halves come from *different* tables
        #: gets no adjustment: kerning is a property of one face, and a shaper breaks the
        #: run at the font boundary the same way.
        previous: tuple[str, FontMetrics | None] = ("", None)

        for char in text:
            code_point = ord(char)
            east_asian = is_cjk(code_point)
            metrics = ea_metrics if east_asian and ea_metrics else latin_metrics
            if metrics is not None and metrics is previous[1]:
                total += _kern_px(previous[0], char, base_size_px, metrics, bold)
            previous = (char, metrics)
            if metrics is None:
                width = base_size_px * _heuristic_ratio(char, code_point)
                if bold and not east_asian:
                    width *= BOLD_FACTOR
            elif bold and metrics.bold_widths:
                # The rasteriser will draw this with the real bold face, so measure with
                # the real bold face.  See BOLD_FACTOR for what the alternative costs.
                width = _measure_with_metrics(char, code_point, base_size_px, metrics, True)
                if east_asian and metrics.bold_is_indistinguishable():
                    # ...unless there is no bold face to draw with.  A bold table equal
                    # to the upright one means the file had no bold cut, so PowerPoint
                    # emboldened the upright and the advance grew with it.  See
                    # EAST_ASIAN_SYNTHETIC_BOLD_PT.
                    width += synthetic_bold_px
            else:
                width = _measure_with_metrics(char, code_point, base_size_px, metrics, False)
                if bold and not east_asian:
                    width *= BOLD_FACTOR
                elif bold and east_asian:
                    width += synthetic_bold_px
            total += width
        return total

    def kern_between(
        self,
        left: str,
        right: str,
        font_size_pt: float,
        bold: bool = False,
        font_family: str | None = None,
        font_family_ea: str | None = None,
    ) -> float:
        if not left or not right:
            return 0.0
        first, second = left[-1], right[0]
        latin_metrics = self._metrics(font_family)
        ea_metrics = self._metrics(font_family_ea)

        def table(char: str) -> FontMetrics | None:
            return ea_metrics if is_cjk(ord(char)) and ea_metrics else latin_metrics

        metrics = table(first)
        if metrics is None or metrics is not table(second):
            return 0.0
        return _kern_px(first, second, font_size_pt * PX_PER_PT, metrics, bold)

    def line_height_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        # Deliberately ignores the font: see DEFAULT_LINE_HEIGHT_RATIO.  The arguments
        # stay for the protocol's sake, and because a future rule may need them.
        return DEFAULT_LINE_HEIGHT_RATIO

    def ascender_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        metrics = self._metrics(font_family) or self._metrics(font_family_ea)
        if metrics is None:
            return DEFAULT_ASCENDER_RATIO
        return _first_baseline_ratio(abs(metrics.descender) / metrics.units_per_em)


def _first_baseline_ratio(descender_ratio: float) -> float:
    """Where the first baseline sits below the top of the text, in ems.

    PowerPoint hangs the line box's *bottom* off the font's descent rather than its top
    off the font's ascent: the baseline lands one descent up from the bottom of the
    1.2 em box, and whatever is left over becomes leading above the text.  Measuring the
    baseline of an "H" in a top-anchored box with zero inset gave 14 pt for 14 pt Arial
    (descent 0.212 em -> 0.988 em), 14 pt for Times New Roman (0.216 -> 0.984) and 13 pt
    for Calibri (0.269 -> 0.931); PowerPoint rounds the offset to whole points, and all
    three round correctly.  Using the ascent instead -- the obvious reading -- puts Arial
    and Times a full point too high.

    Only single spacing is modelled.  Percentages above 100% measured close to
    ``ascent * percentage`` instead, which does not meet this formula at 100%, so the
    rule there is something else and is left alone rather than guessed at.
    """
    return max(0.0, DEFAULT_LINE_HEIGHT_RATIO - descender_ratio)


def _kern_px(
    first: str, second: str, base_size_px: float, metrics: FontMetrics, bold: bool
) -> float:
    """What the ``kern`` feature takes off the join between two characters, in pixels.

    Zero for a face with no kern table, which is how it stays free for the eighteen
    monospaced and full-width faces and for every embedded one.

    The bold question is settled the same way the advance widths settle it: the bold
    pairs are used exactly when the bold *widths* are, so a face whose bold cut we do not
    have is kerned with its upright pairs rather than not kerned at all.
    """
    table = metrics.kerning
    if table is None:
        return 0.0
    units = table.adjustment(first, second, bold and bool(metrics.bold_widths))
    if not units:
        return 0.0
    return (units / metrics.units_per_em) * base_size_px


def _measure_with_metrics(
    char: str, code_point: int, base_size_px: float, metrics: FontMetrics, bold: bool = False
) -> float:
    widths = metrics.bold_widths if bold else metrics.widths
    width = widths.get(char)
    if width is None:
        if is_cjk(code_point):
            width = metrics.bold_cjk_width if bold else metrics.cjk_width
        else:
            width = metrics.bold_default_width if bold else metrics.default_width
    return (width / metrics.units_per_em) * base_size_px


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
        self._kern_cache: dict[int, dict[str, int]] = {}

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
        previous: tuple[str, object] = ("", None)
        for char in text:
            font = east_asian if is_cjk(ord(char)) and east_asian is not None else latin
            if font is not None and font is previous[1]:
                total += self._kern_px(font, previous[0], char, base_size_px)
            previous = (char, font)
            advance = _advance_width(font, char)
            if advance is None:
                total += self._fallback.measure_text_width(
                    char, font_size_pt, bold, font_family, font_family_ea
                )
                continue
            units_per_em = font["head"].unitsPerEm  # type: ignore[index]
            total += (advance / units_per_em) * base_size_px
        return total

    def kern_between(
        self,
        left: str,
        right: str,
        font_size_pt: float,
        bold: bool = False,
        font_family: str | None = None,
        font_family_ea: str | None = None,
    ) -> float:
        if not left or not right:
            return 0.0
        first, second = left[-1], right[0]
        latin = self._face(font_family)
        east_asian = self._face(font_family_ea)
        if latin is None and east_asian is None:
            return self._fallback.kern_between(
                left, right, font_size_pt, bold, font_family, font_family_ea
            )

        def face(char: str):
            return east_asian if is_cjk(ord(char)) and east_asian is not None else latin

        font = face(first)
        if font is None or font is not face(second):
            return 0.0
        return self._kern_px(font, first, second, font_size_pt * PX_PER_PT)

    def _kern_px(self, font, first: str, second: str, base_size_px: float) -> float:
        """The real file's ``kern`` adjustment for one pair, in pixels.

        Read straight out of GPOS, with the legacy ``kern`` table as the fallback for a
        face that has one and no feature -- the same two sources
        ``tools/extract_font_metrics.py`` bakes the static tables from, so this measurer
        and the default one answer alike for a face that appears in both.

        Cached per pair rather than expanded up front: a face carries thousands of pairs
        and a deck asks about a few hundred.
        """
        pairs = self._kern_cache.get(id(font))
        if pairs is None:
            pairs = {}
            self._kern_cache[id(font)] = pairs
        key = first + second
        units = pairs.get(key)
        if units is None:
            units = _font_kern_units(font, first, second)
            pairs[key] = units
        if not units:
            return 0.0
        return (units / font["head"].unitsPerEm) * base_size_px

    def line_height_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        # The line box is PowerPoint's, not the face's, even when the face is readable.
        return DEFAULT_LINE_HEIGHT_RATIO

    def ascender_ratio(
        self, font_family: str | None = None, font_family_ea: str | None = None
    ) -> float:
        font = self._face(font_family) or self._face(font_family_ea)
        if font is None:
            return self._fallback.ascender_ratio(font_family, font_family_ea)
        hhea = font["hhea"]  # type: ignore[index]
        units_per_em = font["head"].unitsPerEm  # type: ignore[index]
        return _first_baseline_ratio(abs(hhea.descender) / units_per_em)


def _font_kern_units(font, first: str, second: str) -> int:
    """One pair's ``kern`` adjustment, in the file's own units.

    Mirrors ``tools/extract_font_metrics.py``: within one lookup the first subtable that
    *covers* the pair wins and the rest of that lookup is skipped, while separate lookups
    each get a pass and their adjustments add.  A face with no GPOS ``kern`` feature but
    a legacy ``kern`` table falls back to that.
    """
    try:
        cmap = font.getBestCmap()
        left = cmap.get(ord(first))
        right = cmap.get(ord(second))
        if left is None or right is None:
            return 0
        total = 0
        covered = False
        for lookup in _gpos_kern_lookups(font):
            for subtable in lookup:
                value = _pair_adjustment(subtable, left, right)
                if value is not None:
                    total += value
                    covered = True
                    break
        if covered:
            return total
        if "kern" not in font:
            return 0
        for subtable in font["kern"].kernTables:
            total += subtable.kernTable.get((left, right), 0)
        return total
    except Exception:  # a font whose tables will not parse simply does not kern
        return 0


def _gpos_kern_lookups(font) -> list[list]:
    """The ``kern`` feature's pair-adjustment subtables, grouped by lookup."""
    if "GPOS" not in font:
        return []
    gpos = font["GPOS"].table
    if gpos is None or gpos.FeatureList is None:
        return []
    wanted: set[int] = set()
    for record in gpos.FeatureList.FeatureRecord:
        if record.FeatureTag == "kern":
            wanted.update(record.Feature.LookupListIndex)
    groups: list[list] = []
    for index in sorted(wanted):
        lookup = gpos.LookupList.Lookup[index]
        group = []
        for subtable in lookup.SubTable:
            if lookup.LookupType == 9:  # extension positioning
                subtable = subtable.ExtSubTable
                kind = subtable.LookupType
            else:
                kind = lookup.LookupType
            if kind == 2:
                group.append(subtable)
        if group:
            groups.append(group)
    return groups


def _pair_adjustment(subtable, left: str, right: str) -> int | None:
    """``None`` when the subtable does not cover the pair, which is not the same as 0."""
    coverage = subtable.Coverage.glyphs
    if subtable.Format == 1:
        try:
            index = coverage.index(left)
        except ValueError:
            return None
        for record in subtable.PairSet[index].PairValueRecord:
            if record.SecondGlyph == right:
                return getattr(record.Value1, "XAdvance", 0) or 0
        return None
    if left not in coverage:
        return None
    first = subtable.ClassDef1.classDefs.get(left, 0)
    second = subtable.ClassDef2.classDefs.get(right, 0)
    if first >= subtable.Class1Count or second >= subtable.Class2Count:
        return None
    value = subtable.Class1Record[first].Class2Record[second].Value1
    return getattr(value, "XAdvance", 0) or 0


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

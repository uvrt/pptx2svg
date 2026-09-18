#!/usr/bin/env python3
"""Build a deck of scatter charts whose value axis spans fall just under a power of ten.

The question this exists to settle is :data:`pptx2svg.resolve.chart._DECADE_SLACK`: how
far *below* a power of ten an axis span may fall and still be treated as having reached
it.  A span is a subtraction of two authored decimals and lands a few ulps under a round
number -- ``1.13 - 1.03`` is ``0.09999999999999987`` -- and whether that span gets the
decade of 0.1 or the decade of 0.01 changes the drawn tick unit by a factor of five.

Only an **unanchored** value axis shows the difference: an axis pulled back to zero has
its extent rounded far enough that the unit choice is absorbed.  A scatter's value axes
are the only ones PowerPoint leaves unanchored, and only when the data sits past
``AXIS_ZERO_ANCHOR_RATIO`` of its own range, so every probe here keeps the data high in
its range (min/max ~= 0.91) and uses a scatter.

One chart per slide, the y axis carrying the gridlines and the x axis deleted, so the
export's page objects are unambiguous: the horizontal strokes are the y major gridlines
and the only text is the y tick labels.

**What the decks found**, in the order they were built, because each one exists to
answer what the last one raised:

``axis-decade``
    The slack is unmeasurable.  Thirty spans falling under a tenth by a relative 1e-15
    through 5e-2 were drawn with one and the same axis, so PowerPoint never asks which
    decade a span is in and no observation can bracket the boundary.
``axis-density``
    The unit moves with the frame's *height*, and the span sweep says the base unit is
    the smallest 1-2-5 step of at least a tenth of the range.
``axis-count``
    Eleven intervals accepted, twelve refused, and the label size -- not the label text --
    drives the coarsening on a short axis.
``axis-pad``
    The range is padded by **5%** at each end before any of that: bracketed to
    (4.35%, 5.04%] by probes either side of a 1-2-5 boundary.
``axis-shape``
    On an axis held at zero the headroom is 5% of the *maximum*, not of the data range.
``axis-bar``
    A real bar chart obeys the same rule, and the corpus' by-500 axis is a *coarsened*
    by-200 one, which is where the shipped decade-and-halve rule came from.
``axis-inset``
    Not the axis at all but the **plot rectangle** it is drawn in, read off the gridlines
    rather than the tick labels' centres: the top inset and the bottom category band,
    separately, over four faces and eight label sizes.  What it found is in ROADMAP.md 3.6
    -- the top inset was right to 0.002 pt and the band's gap term is two thirds of the
    face's *ascent*.  Read it with ``--insets``.

Usage -- the deck's file name picks its probe table::

    python3 tools/make_axis_probe.py ~/pptx2svg-oracle/axis-decade.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/axis-decade.pptx ~/pptx2svg-oracle/axis-decade.pdf
    python3 tools/read_axis_probe.py ~/pptx2svg-oracle/axis-decade.pdf
    python3 tools/read_axis_probe.py ~/pptx2svg-oracle/axis-inset.pdf --insets

The decks and their exports are throwaway and are **not** committed; ``~/pptx2svg-oracle``
is the directory PowerPoint is allowed to write to, and it is left holding its sixteen
corpus files.  See ROADMAP.md section 0.1 before blaming a failed export on the path.
"""

from __future__ import annotations

import math
import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SOURCE = HERE.parent / "tests/fixtures/authoring-integration.pptx"

C = "xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart'"
A = "xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'"
R = "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'"
NAMESPACES = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
)
CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"

EMPTY_TREE = (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
)

#: Frame geometry in EMU.  The tall frame is nearly the whole 720x405 pt slide; the short
#: one is the same width with a quarter of the height, which is how a *density* rule --
#: if PowerPoint has one -- gets separated from a decade rule: a decade rule cannot care
#: about the axis length and a density rule must.
FRAME_OFF = (228600, 152400)
FRAME_TALL = (8686800, 4838700)
FRAME_SHORT = (8686800, 1257300)


def high_for(low: float, decade: float, shortfall: float) -> float:
    """The largest double whose distance from *low* is under ``decade * (1 - shortfall)``.

    ``shortfall=0`` gives the smallest span that *reaches* the decade instead, which is
    the control at the other side of the boundary.
    """
    if shortfall == 0.0:
        high = low + decade
        while high - low < decade:
            high = math.nextafter(high, math.inf)
        return high
    high = low + decade * (1.0 - shortfall)
    while high - low >= decade:
        high = math.nextafter(high, -math.inf)
    return high


def _ladder(low: float, decade: float, shortfalls, tag: str, frame=FRAME_TALL):
    for shortfall in shortfalls:
        high = high_for(low, decade, shortfall)
        realised = 1.0 - (high - low) / decade
        name = "over" if shortfall == 0 else f"{realised:.0e}".replace("e-0", "e-")
        yield {
            "key": f"{tag}-{name}",
            "low": low,
            "high": high,
            "decade": decade,
            "frame": frame,
        }


#: The relative shortfalls probed.  The two smallest are what a subtraction of authored
#: decimals actually leaves (a few ulps, ~1e-15 relative); the largest, 5e-2, is a span
#: somebody *wrote* as 0.095, which any rule that reads the span's decade must put in the
#: decade below -- and which PowerPoint draws exactly like the ones a single ulp short.
SHORTFALLS = (0.0, 1.3e-15, 2e-14, 1e-13, 1e-11, 1e-9, 1e-6, 1e-4, 1e-2, 5e-2)
COARSE = (0.0, 1.3e-15, 1e-11, 1e-9, 1e-4, 5e-2)

DECADE_PROBES: list[dict] = [
    # Four pairs of *authored two-decimal* numbers a tenth apart, which is where the
    # question came from.  Sweeping every such pair whose ratio leaves the axis
    # unanchored gives exactly four residues: exact (1.00/1.10), 10 ulps under
    # (1.03/1.13 -- the pair that raised this), 26 (2.16/2.26) and 90 (7.94/8.04).
    {"key": "pair-1.00", "low": 1.00, "high": 1.10, "decade": 0.1, "frame": FRAME_TALL},
    {"key": "pair-1.03", "low": 1.03, "high": 1.13, "decade": 0.1, "frame": FRAME_TALL},
    {"key": "pair-2.16", "low": 2.16, "high": 2.26, "decade": 0.1, "frame": FRAME_TALL},
    {"key": "pair-7.94", "low": 7.94, "high": 8.04, "decade": 0.1, "frame": FRAME_TALL},
    *_ladder(1.03, 0.1, SHORTFALLS, "d.1"),
    *_ladder(10.3, 1.0, COARSE, "d1"),
    *_ladder(1030.0, 100.0, COARSE, "d100"),
    *_ladder(1.03, 0.1, (0.0, 1.3e-15, 1e-9, 5e-2), "short", frame=FRAME_SHORT),
]

#: The second deck.  The first one showed the unit on an unanchored axis moving with the
#: frame's *height* and not with the span's decade, so this one varies each of the three
#: things the unit could depend on -- height, span, and whether the axis is anchored at
#: zero -- one at a time, with everything else held.
HEIGHTS = (4838700, 3600000, 2700000, 2000000, 1500000, 1257300, 1000000, 800000, 600000)
SPANS = (0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1.0, 2.0, 5.0, 10.0)
ANCHORED = ((0.004, 0.1), (0.2, 5.0), (0.4, 9.0), (100.0, 1842.0))

DENSITY_PROBES: list[dict] = [
    # Height sweep: one dataset, nine frame heights.
    *[
        {
            "key": f"h{height // 10000}",
            "low": 1.03,
            "high": 1.13,
            "decade": 0.1,
            "frame": (FRAME_TALL[0], height),
        }
        for height in HEIGHTS
    ],
    # Span sweep at the tallest frame, every dataset unanchored (min/max = 10/11).
    *[
        {
            "key": f"s{span:g}",
            "low": 10 * span,
            "high": 11 * span,
            "decade": span,
            "frame": FRAME_TALL,
        }
        for span in SPANS
    ],
    # Anchored data -- the near end well inside AXIS_ZERO_ANCHOR_RATIO -- at the tallest
    # frame and at a short one, which is the comparison that says whether the anchored
    # axis takes its unit the same way.
    *[
        {
            "key": f"a{low:g}-{high:g}",
            "low": low,
            "high": high,
            "decade": 10 ** math.floor(math.log10(high - low)),
            "frame": FRAME_TALL,
        }
        for low, high in ANCHORED
    ],
    *[
        {
            "key": f"as{low:g}-{high:g}",
            "low": low,
            "high": high,
            "decade": 10 ** math.floor(math.log10(high - low)),
            "frame": FRAME_SHORT,
        }
        for low, high in ANCHORED[1:3]
    ],
]

#: The third deck.  The second one said the unit is the finest 1-2-5 step whose *rounded
#: extent* holds no more than about ten intervals, coarsened again while the step would
#: be too few points tall.  This one puts numbers on both halves: datasets whose finer
#: unit lands on exactly eleven intervals separate a limit of ten from one of eleven, and
#: the same short axis at three label sizes says whether the point floor is the label's
#: height or a constant.
FRAME_BAR = (2800000, 2300000)  # the frame the bar and scatter sweeps were measured on
FRAME_MID = (8686800, 1900000)

COUNT_PROBES: list[dict] = [
    # Eleven intervals at the finer unit: drawn fine means the limit is at least eleven.
    {"key": "n11-un", "low": 1.04, "high": 1.13, "decade": 0.01, "frame": FRAME_TALL},
    {"key": "n11-un2", "low": 10.4, "high": 11.3, "decade": 0.1, "frame": FRAME_TALL},
    {"key": "n11-anch", "low": 0.3, "high": 5.2, "decade": 1.0, "frame": FRAME_TALL},
    # Ten intervals at the finer unit: the control that says the limit is at least ten.
    {"key": "n10-un", "low": 1.05, "high": 1.13, "decade": 0.01, "frame": FRAME_TALL},
    {"key": "n10-anch", "low": 0.3, "high": 4.9, "decade": 1.0, "frame": FRAME_TALL},
    # The five observations AXIS_HALVING_RATIO was fitted to, on a frame four times the
    # height of the one they were measured on.
    {"key": "fit5", "low": 0.5, "high": 5.0, "decade": 1.0, "frame": FRAME_TALL},
    {"key": "fit9", "low": 0.5, "high": 9.0, "decade": 1.0, "frame": FRAME_TALL},
    {"key": "fit1842", "low": 50.0, "high": 1842.0, "decade": 1000.0, "frame": FRAME_TALL},
    {"key": "fit4285", "low": 50.0, "high": 4285.0, "decade": 1000.0, "frame": FRAME_TALL},
    # ... and on the frame they were measured on, to tie the two together.
    {"key": "bar5", "low": 0.5, "high": 5.0, "decade": 1.0, "frame": FRAME_BAR},
    {"key": "bar9", "low": 0.5, "high": 9.0, "decade": 1.0, "frame": FRAME_BAR},
    {"key": "bar1842", "low": 50.0, "high": 1842.0, "decade": 1000.0, "frame": FRAME_BAR},
    {"key": "bar4285", "low": 50.0, "high": 4285.0, "decade": 1000.0, "frame": FRAME_BAR},
    # One axis length, three label sizes.  A floor that is the label's height moves with
    # them; a constant floor does not.
    {"key": "t6", "low": 1.03, "high": 1.13, "decade": 0.1, "frame": FRAME_MID, "size": 600},
    {"key": "t10", "low": 1.03, "high": 1.13, "decade": 0.1, "frame": FRAME_MID, "size": 1000},
    {"key": "t18", "low": 1.03, "high": 1.13, "decade": 0.1, "frame": FRAME_MID, "size": 1800},
    {"key": "t28", "low": 1.03, "high": 1.13, "decade": 0.1, "frame": FRAME_MID, "size": 2800},
]

#: The fourth deck.  Everything drawn so far fits "the unit is the smallest 1-2-5 step of
#: at least a tenth of the data range *after a percentage of headroom at each end*", and
#: the headroom is what is left to measure.  Each probe here sits on one side of a
#: candidate headroom: the anchored ones switch unit at a headroom of 2.0, 4.2, 5.0, 6.4
#: and 8.7 per cent of the maximum, the unanchored ones at 5.6, 4.3, 2.5, 1.0 and 0.05 per
#: cent of the range at *each* end, and the last two say whether the extent's ends are
#: rounded from the padded range or from the data.
PAD_PROBES: list[dict] = [
    *[
        {"key": f"p-anch{high:g}", "low": 0.3, "high": high, "decade": 1.0, "frame": FRAME_TALL}
        for high in (4.9, 4.8, 4.76, 4.7, 4.6)
    ],
    *[
        {"key": f"p-un{high:g}", "low": 1.0, "high": high, "decade": 0.1, "frame": FRAME_TALL}
        for high in (1.09, 1.092, 1.0952, 1.098, 1.0999)
    ],
    {"key": "p-low", "low": 1.025, "high": 1.13, "decade": 0.1, "frame": FRAME_TALL},
    {"key": "p-high", "low": 1.03, "high": 1.119, "decade": 0.1, "frame": FRAME_TALL},
]

#: The fifth deck, on the two things the fourth left open.  The headroom on an axis held
#: at zero could be five per cent of the *range* or of the *maximum*, which only separate
#: when the data floor is well up the range: 3.0..4.9 and 3.5..4.9 fall on opposite sides
#: of the 1-2-5 boundary under one and the same side under the other.  And the second,
#: coarsening stage refused a 16.0 pt step on one probe while accepting 15.9 pt on
#: another, so the same two families are swept over a set of heights with the *only*
#: difference being what the labels say -- 103..113 draws three-digit integers where
#: 1.03..1.13 draws four-character decimals.
SHAPE_PROBES: list[dict] = [
    {"key": "basis3.0", "low": 3.0, "high": 4.9, "decade": 1.0, "frame": FRAME_TALL},
    {"key": "basis3.5", "low": 3.5, "high": 4.9, "decade": 1.0, "frame": FRAME_TALL},
    *[
        {
            "key": f"int{height // 10000}",
            "low": 103.0,
            "high": 113.0,
            "decade": 10.0,
            "frame": (FRAME_TALL[0], height),
        }
        for height in (2000000, 1500000, 1257300, 1000000)
    ],
    *[
        {
            "key": f"anch{height // 10000}",
            "low": 0.5,
            "high": 9.0,
            "decade": 1.0,
            "frame": (FRAME_TALL[0], height),
        }
        for height in (2700000, 2000000, 1500000, 1257300)
    ],
]

#: The sixth deck, on the one thing the coarsening stage kept contradicting.  Two of the
#: observations `AXIS_HALVING_RATIO` was fitted to read as though they sat on the same
#: frame and the same 10 pt labels: 0..9 took ten intervals at 14.5 pt where 0..1842 took
#: four.  The magnitudes differ, and so does the *width of the label text* -- "10" against
#: "2000" -- so these are bar charts, the type both came from, and ``tiny`` carries 0..9's
#: magnitude with 1842's label width to separate the two.  Each is drawn at two frame
#: widths as well, which changes what a label column costs without changing the label.
#:
#: **Neither mattered**: all three take ten intervals at 14.5 pt on the 145.0 pt plot, so
#: the by-500 axis in the corpus is a *coarsened* by-200 one from that deck's own smaller
#: frames, and label width and magnitude are both eliminated.
BAR_PROBES: list[dict] = [
    {"key": "bar9", "low": 0.0, "high": 9.0, "decade": 1.0, "frame": FRAME_BAR, "bar": True},
    {"key": "bar1842", "low": 0.0, "high": 1842.0, "decade": 1000.0, "frame": FRAME_BAR,
     "bar": True},
    {"key": "bartiny", "low": 0.0, "high": 0.0092, "decade": 0.001, "frame": FRAME_BAR,
     "bar": True},
    {"key": "bar9w", "low": 0.0, "high": 9.0, "decade": 1.0, "frame": (8686800, 2300000),
     "bar": True},
    {"key": "bar1842w", "low": 0.0, "high": 1842.0, "decade": 1000.0,
     "frame": (8686800, 2300000), "bar": True},
    {"key": "bartinyw", "low": 0.0, "high": 0.0092, "decade": 0.001,
     "frame": (8686800, 2300000), "bar": True},
    {"key": "bar1842t", "low": 0.0, "high": 1842.0, "decade": 1000.0,
     "frame": (2800000, 4838700), "bar": True},
]

#: The seventh deck, on the coarsening stage itself.  Every earlier deck held the chart
#: **type** fixed while varying the frame, and the readings that could not be ordered came
#: from different types: a scatter refused six intervals on a 95.9 pt plot while a bar
#: accepted ten on a 145.0 pt one, both at 10 pt labels.  So this one sweeps the same two
#: datasets over the same eight frames on **four** types, which is the comparison no deck
#: has made: if the coarsening is one mechanism, a scatter and a column chart whose plots
#: come out the same height must coarsen at the same height.
#:
#: The datasets are chosen for their *base interval count* -- 0..5 draws 0..6 by 1 (six)
#: and 0..9 draws 0..10 by 1 (ten) -- so each probe brackets the threshold from both
#: sides: the count drawn is accepted and the next finer rung of the 1-2-5 ladder was
#: refused.  Both are anchored at zero and both label with plain integers, which the
#: shape deck already showed the coarsening does not read.
COARSE_HEIGHTS = (600000, 850000, 1100000, 1400000, 1800000, 2300000, 3000000, 4200000)
COARSE_DATA = ((0.0, 5.0, "k6"), (0.0, 9.0, "k10"))

COARSE_PROBES: list[dict] = [
    {
        "key": f"{kind or 'sc'}-{tag}-{height // 10000}",
        "low": low,
        "high": high,
        "decade": 1.0,
        "frame": (8686800, height),
        "size": 1000,
        **({"kind": kind} if kind else {}),
    }
    for kind in (None, "col", "line", "area")
    for low, high, tag in COARSE_DATA
    for height in COARSE_HEIGHTS
]

#: The eighth deck, which turns the seventh's answer into numbers.  ``axis-coarse`` said
#: the interval count is a function of the **frame** height and nothing else -- four chart
#: types whose plots come out 14 pt apart coarsen at exactly the same frame -- so what is
#: left is the function.  Under ``unit = ceil125(padded_range / N)`` the drawn unit reads
#: *N* back, and five datasets read it exactly: their 1-2-5 boundaries fall at N = 2, 3
#: (A), 5, 10 (B), 4, 8 (C), 7 (D) and 9 (E), so the five units drawn together name every
#: N from 1 to 11.  That is the "N-meter": one cell of five slides per frame and font.
#:
#: ``m`` sweeps the frame at 10 pt to fit the function; ``f`` sweeps the label size at
#: three frames each to find what in it is font-driven; ``w`` varies the frame's *width*
#: with its height held, which a height rule must ignore; ``x`` puts the value axis along
#: the **bottom** of a horizontal bar chart and sweeps the width instead, and ``xb``
#: repeats it with labels a million times wider to see whether a bottom axis counts its
#: labels' width where a side axis counts nothing.
METER = (("A", 5.0), ("B", 9.0), ("C", 3.8), ("D", 6.6), ("E", 8.5))
RUNG_HEIGHTS = (60, 75, 90, 105, 120, 135, 150, 165, 180, 195)
RUNG_FONTS = ((600, (45, 70, 100)), (800, (55, 95, 135)), (1400, (95, 165, 240)),
              (2000, (130, 230, 340)), (2800, (180, 300, 385)))
RUNG_WIDTHS = (150, 250, 350, 450, 550, 680)

RUNG_PROBES: list[dict] = [
    *[
        {
            "key": f"m{height}-{tag}",
            "low": 0.0,
            "high": high,
            "decade": 1.0,
            "frame": (8686800, height * 12700),
            "size": 1000,
            "kind": "col",
        }
        for height in RUNG_HEIGHTS
        for tag, high in METER
    ],
    *[
        {
            "key": f"f{size // 100}-{height}-{tag}",
            "low": 0.0,
            "high": high,
            "decade": 1.0,
            "frame": (8686800, height * 12700),
            "size": size,
            "kind": "col",
        }
        for size, heights in RUNG_FONTS
        for height in heights
        for tag, high in METER
    ],
    *[
        {
            "key": f"w{width}-{tag}",
            "low": 0.0,
            "high": high,
            "decade": 1.0,
            "frame": (width * 12700, 150 * 12700),
            "size": 1000,
            "kind": "col",
        }
        for width in (150, 350)
        for tag, high in METER[:2]
    ],
    *[
        {
            "key": f"x{width}-{tag}",
            "low": 0.0,
            "high": high,
            "decade": 1.0,
            "frame": (width * 12700, 200 * 12700),
            "size": 1000,
            "kind": "bar",
        }
        for width in RUNG_WIDTHS
        for tag, high in METER[:2]
    ],
    *[
        {
            "key": f"xb{width}-B",
            "low": 0.0,
            "high": 9.0e6,
            "decade": 1.0,
            "frame": (width * 12700, 200 * 12700),
            "size": 1000,
            "kind": "bar",
        }
        for width in RUNG_WIDTHS
    ],
]

#: The ninth deck, which measures the two numbers in the eighth's function and asks what
#: height it is a function *of*.  ``axis-rung`` fitted
#: ``N = floor((frame - c) / K)`` with ``K`` a multiple of the label size and ``c`` an
#: offset with a fixed and a per-point part; a **transition scan** pins each of them.  At
#: the frame where ``N`` steps from one to the next, one dataset's drawn unit changes and
#: nothing else does, so a single probe per height reads the transition: ``B`` steps at
#: N = 2, 5 and 10, and ``A`` at N = 2, 3 and 6.  Each scan walks the window the fit
#: leaves open, two to four points at a time.
#:
#: ``g`` asks the other half.  Every frame so far held the chart's whole height, and the
#: area probes in ROADMAP.md say a **legend** changes the answer: two charts with the same
#: data drew ten intervals and five, differing only in that one had a bottom legend.  So
#: these repeat one cell of the N-meter with a legend below, a legend at the right and a
#: title, which separates "the frame" from "what is left of the frame after the furniture".
def _scan(size: int, rung: int, tag: str, heights) -> list[dict]:
    high = 5.0 if tag == "A" else 9.0
    return [
        {
            "key": f"s{size // 100}n{rung}-{height}",
            "low": 0.0,
            "high": high,
            "decade": 1.0,
            "frame": (8686800, int(round(height * 12700))),
            "size": size,
            "kind": "col",
        }
        for height in heights
    ]


BAND_PROBES: list[dict] = [
    *_scan(600, 2, "B", range(44, 58, 2)),
    *_scan(600, 10, "B", range(105, 117, 2)),
    *_scan(1000, 2, "B", range(62, 78, 2)),
    *_scan(1000, 10, "B", range(164, 180, 2)),
    *_scan(1400, 2, "B", range(80, 98, 2)),
    *_scan(1400, 10, "B", range(220, 242, 2)),
    *_scan(2000, 2, "B", range(106, 133, 3)),
    *_scan(2000, 10, "B", range(304, 337, 3)),
    *_scan(2800, 2, "B", range(140, 176, 4)),
    *_scan(2800, 6, "A", range(281, 305, 3)),
    *[
        {
            "key": f"g{what}-{tag}",
            "low": 0.0,
            "high": high,
            "decade": 1.0,
            "frame": (8686800, 150 * 12700),
            "size": 1000,
            "kind": "col",
            **extra,
        }
        for what, extra in (
            ("lb", {"legend": "b"}),
            ("lr", {"legend": "r"}),
            ("ti", {"title": "Probe title"}),
        )
        for tag, high in METER
    ],
]

#: The tenth deck.  ``axis-band`` left three things open and this closes two of them.
#:
#: ``r`` halves every transition bracket by scanning the one point the two-point scan
#: skipped, which is what separates a rung of ``1.2 * size`` -- PowerPoint's line box,
#: measured as a constant in :data:`text.measure.DEFAULT_LINE_HEIGHT_RATIO` -- from the
#: slightly larger one the two-point scans bracket.
#:
#: ``a`` asks whether the rung is the *face's* line box or PowerPoint's constant one, by
#: drawing Arial labels at heights that sit either side of an Aptos transition.  Arial's
#: own ascent+descent is 1.15 em against Aptos' 1.2207, so a face-driven rung moves the
#: transition by 7 pt at 10 pt labels and a constant one does not move it at all.
#:
#: ``h`` sweeps a **horizontal** value axis' frame width at three label widths -- "10",
#: "10.000" and "10.000.000" -- because the bottom axis is the one place a label's own
#: width can gate the count, and ``hh`` varies that frame's *height* instead, which a
#: width rule must ignore.
def _face_probe(size: int, height: int, face: str, tag: str, high: float) -> dict:
    return {
        "key": f"a{face[:3].lower()}{size // 100}-{height}-{tag}",
        "low": 0.0,
        "high": high,
        "decade": 1.0,
        "frame": (8686800, height * 12700),
        "size": size,
        "kind": "col",
        "face": face,
    }


def _wide(width: int, high: float, tag: str, height: int = 200, labels=None) -> dict:
    return {
        "key": f"h{tag}{width}",
        "low": 0.0,
        "high": high,
        "decade": 1.0,
        "frame": (width * 12700, height * 12700),
        "size": 1000,
        "kind": "bar",
        **({"labels": labels} if labels else {}),
    }


WIDE_CATS = ("Category number one", "Category number two", "Category number three",
             "Category number four", "Category number five")

WIDER_PROBES: list[dict] = [
    *_scan(600, 2, "B", (51,)),
    *_scan(600, 10, "B", (110,)),
    *_scan(1000, 2, "B", (71,)),
    *_scan(1000, 10, "B", (169,)),
    *_scan(1400, 2, "B", (91,)),
    *_scan(1400, 10, "B", (227,)),
    *_scan(2000, 2, "B", (119, 120)),
    *_scan(2000, 10, "B", (314, 315)),
    *_scan(2800, 2, "B", (157, 158, 159)),
    *_scan(2800, 6, "A", (294, 295)),
    *[
        _face_probe(1000, height, "Arial", "B", 9.0)
        for height in (160, 165, 169, 172, 175)
    ],
    *[
        _face_probe(2000, height, "Arial", "B", 9.0)
        for height in (300, 310, 320, 330)
    ],
    # A horizontal value axis, three label widths, the frame width swept.
    *[_wide(w, 9.0, "n", ) for w in range(160, 260, 10)],
    *[_wide(w, 9.0, "n") for w in range(360, 460, 10)],
    *[_wide(w, 9.0e6, "w") for w in range(260, 360, 10)],
    *[_wide(w, 9.0e6, "w") for w in range(460, 560, 10)],
    *[_wide(w, 9.0e3, "m") for w in range(180, 580, 40)],
    # The same bottom axis with much wider category labels, which take the band at the
    # left, and at three frame heights, which a width rule must ignore.
    *[_wide(w, 9.0, "c", labels=WIDE_CATS) for w in (200, 250, 300, 400, 450, 500)],
    *[_wide(400, 9.0, f"t{height}", height=height) for height in (80, 140, 320)],
]

#: The eleventh deck, on the **bottom** axis alone.  ``axis-wider`` found its count is a
#: function of the frame *width* and that the label's own width moves it -- "10" and
#: "10.000.000" cross the same rung 50 pt apart -- which is the one thing the side axis
#: does not do.  These scans put brackets on both: ``p`` walks the two transitions at
#: three label widths two points at a time, ``q`` repeats the width sweep at 20 pt labels
#: to see whether the rung is font-driven as the side axis' is, and ``v`` varies the
#: frame's height at a width that sits on a transition, which a width rule must ignore.
BOTTOM_PROBES: list[dict] = [
    *[_wide(w, 9.0, "p") for w in (222, 224, 226, 228)],
    *[_wide(w, 9.0, "p") for w in (422, 424, 426, 428)],
    *[_wide(w, 9.0e6, "p") for w in (272, 274, 276, 278)],
    *[_wide(w, 9.0e6, "p") for w in (482, 484, 486, 488)],
    *[_wide(w, 9.0e3, "p") for w in (225, 235, 245, 255)],
    *[_wide(w, 9.0e3, "p") for w in (425, 435, 445, 455)],
    *[{**_wide(w, 9.0, "q"), "size": 2000} for w in range(200, 700, 40)],
    *[_wide(426, 9.0, f"v{height}", height=height) for height in (80, 140, 320)],
    *[_wide(228, 9.0, f"u{height}", height=height) for height in (80, 140, 320)],
]

#: The twelfth deck: a **scatter's x axis**, which is the same bottom value axis and the
#: one :data:`HORIZONTAL_MAX_INTERVALS` was fitted to.  Those four readings never recorded
#: their frames, so they cannot be checked against a width rule; these sweep the width
#: with the same two datasets, ``0.5..4.5`` (the probe that discriminated the old cap) and
#: ``1..5``.
SCATTER_PROBES: list[dict] = [
    {
        "key": f"sx{tag}-{width}",
        "low": low,
        "high": high,
        "decade": 1.0,
        "frame": (width * 12700, 250 * 12700),
        "size": 1000,
        "read": "x",
    }
    for low, high, tag in ((0.5, 4.5, "f"), (1.0, 5.0, "i"))
    for width in (160, 200, 240, 300, 400, 684)
]

#: The thirteenth deck: a **radar's radial axis**, which is a value axis on a plot that is
#: half as tall as it looks -- the axis runs from the centre to the rim, not across the
#: frame.  ``real-financial-report.pptx``'s radar draws two rings where the Cartesian rule
#: read off the frame asks for five, so what the radial axis divides is measured here
#: rather than guessed: the same two datasets over eight frames, as ``axis-coarse`` swept
#: them for the four Cartesian types.
RADAR_PROBES: list[dict] = [
    {
        "key": f"rad-{tag}-{height // 10000}",
        "low": low,
        "high": high,
        "decade": 1.0,
        "frame": (int(height * 1.6), height),
        "size": 1000,
        "kind": "radar",
    }
    for low, high, tag in COARSE_DATA
    for height in (900000, 1200000, 1500000, 1900000, 2400000, 3000000, 3800000, 4800000)
]

#: The fourteenth deck: the N-meter on a radar, because ``axis-radar``'s two datasets only
#: bracket its count to three or four values per frame and the Cartesian rule is refuted
#: inside those brackets.  Five datasets per frame name N exactly, as they did for the side
#: axis in ``axis-rung``.
RING_PROBES: list[dict] = [
    {
        "key": f"ring{height // 10000}-{tag}",
        "low": 0.0,
        "high": high,
        "decade": 1.0,
        "frame": (int(height * 1.6), height),
        "size": 1000,
        "kind": "radar",
    }
    for height in (1200000, 1500000, 1700000, 1900000, 2200000, 2600000, 3000000, 3800000)
    for tag, high in METER
]

#: The fifteenth deck, on the one reading the radar sweep contradicts.
#: ``real-financial-report.pptx``'s radar draws **two** rings for 0..100 on a 360x150 pt
#: frame where the rule fitted to ``axis-ring`` asks for four, so this replicates that
#: chart and takes one thing off it at a time: its Japanese category labels, its 12 pt
#: size, its Arial face, its six spokes.  Whichever of them the count follows is the term
#: the sweep -- all of it 10 pt Aptos with five Latin categories -- could not see.
JAPANESE_CATS = ("売上成長率", "営業利益率", "海外売上比率", "DX投資額", "従業員満足度", "CO2削減")
LATIN_SIX = ("C1", "C2", "C3", "C4", "C5", "C6")

CORPUS_RADAR: list[dict] = [
    {"key": "cr-as-is", "size": 1200, "face": "Arial", "labels": JAPANESE_CATS},
    {"key": "cr-latin", "size": 1200, "face": "Arial", "labels": LATIN_SIX},
    {"key": "cr-10pt", "size": 1000, "face": "Arial", "labels": JAPANESE_CATS},
    {"key": "cr-aptos", "size": 1200, "face": None, "labels": JAPANESE_CATS},
    {"key": "cr-five", "size": 1200, "face": "Arial", "labels": None},
    {"key": "cr-latin10", "size": 1000, "face": "Arial", "labels": LATIN_SIX},
]
CORPUS_RADAR = [
    {
        "low": 0.0,
        "high": 100.0,
        "decade": 100.0,
        "frame": (int(360 * 12700), int(150 * 12700)),
        "kind": "radar",
        **probe,
    }
    for probe in CORPUS_RADAR
] + [
    # The same question on the sweep's own geometry, where the answer is known to be four.
    {
        "key": f"cr-ring-{tag}",
        "low": 0.0,
        "high": 100.0,
        "decade": 100.0,
        "frame": (int(1700000 * 1.6), 1700000),
        "kind": "radar",
        "size": size,
        "face": face,
        "labels": labels,
    }
    for tag, size, face, labels in (
        ("10apt", 1000, None, None),
        ("12ari", 1200, "Arial", None),
        ("12jpn", 1200, "Arial", JAPANESE_CATS),
    )
]

#: The sixteenth deck, and the only one read with ``--insets``: what the plot rectangle's
#: **top and bottom insets are separately**, rather than what their sum is.
#:
#: Every earlier reading of the rectangle came off the extreme tick labels' centres, which
#: carries whatever a PDF text rect's centre is against the tick it marks -- so the split
#: between the top inset and the bottom band was never observed, only their sum.  The
#: **major gridlines** are the rectangle itself: the topmost one is its top edge and the
#: category axis line its bottom, both drawn strokes and neither a glyph.
#:
#: The sweep is the label's size and its face, because those are the only two things
#: :meth:`~pptx2svg.resolve.chart.ChartBuilder._top_inset` and
#: :meth:`~pptx2svg.resolve.chart.ChartBuilder._bottom_label_band` read.  Aptos and Arial
#: are what separate a rule from a coincidence -- Calibri is metric-compatible with Aptos
#: to the unit, so a Calibri reading confirms nothing an Aptos one did not -- and Times
#: New Roman and Courier New add a third and fourth set of ``hhea`` numbers.  ``f`` holds
#: the frame and sweeps the size; ``h`` holds the size and sweeps the frame, which both
#: rules must ignore.
def _inset_probe(size: int, face: str | None, height: int, tag: str) -> dict:
    return {
        "key": f"{tag}{'apt' if face is None else face[:3].lower()}{size // 100}-{height}",
        "low": 0.0,
        "high": 9.0,
        "decade": 1.0,
        "frame": (8686800, height * 12700),
        "size": size,
        "kind": "col",
        **({"face": face} if face else {}),
    }


INSET_PROBES: list[dict] = [
    *[
        _inset_probe(size, face, 195, "f")
        for face in (None, "Arial")
        for size in (600, 800, 1000, 1200, 1400, 1800, 2400, 2800)
    ],
    *[
        _inset_probe(size, face, 195, "f")
        for face in ("Times New Roman", "Courier New")
        for size in (800, 1000, 1400, 2400)
    ],
    *[
        _inset_probe(size, face, height, "h")
        for face in (None, "Arial")
        for size in (1000, 2400)
        for height in (90, 120, 330)
    ],
]

DECKS = {
    "axis-inset": INSET_PROBES,
    "axis-corpusradar": CORPUS_RADAR,
    "axis-ring": RING_PROBES,
    "axis-radar": RADAR_PROBES,
    "axis-scatter": SCATTER_PROBES,
    "axis-bottom": BOTTOM_PROBES,
    "axis-decade": DECADE_PROBES,
    "axis-coarse": COARSE_PROBES,
    "axis-rung": RUNG_PROBES,
    "axis-band": BAND_PROBES,
    "axis-wider": WIDER_PROBES,
    "axis-density": DENSITY_PROBES,
    "axis-count": COUNT_PROBES,
    "axis-pad": PAD_PROBES,
    "axis-shape": SHAPE_PROBES,
    "axis-bar": BAR_PROBES,
}


def scatter_xy_xml(
    x_low: float, x_high: float, size: int | None = None
) -> bytes:
    """A scatter with **both** axes drawn, for reading the bottom one.

    The x axis a scatter gets is the same value axis a horizontal bar chart puts along the
    bottom, and :data:`HORIZONTAL_MAX_INTERVALS` was fitted to four such readings whose
    frames were never recorded.  This draws them again with the frame swept.
    """
    xs = [x_low, x_low + (x_high - x_low) * 0.4, x_low + (x_high - x_low) * 0.7, x_high]
    xs[0], xs[-1] = x_low, x_high
    ys = [3.0, 5.0, 4.0, 4.5]

    def cache(values):
        points = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
        return (
            "<c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef>"
        )

    def axis(axis_id, cross_id, position, *, gridlines):
        return (
            f"<c:valAx><c:axId val='{axis_id}'/>"
            "<c:scaling><c:orientation val='minMax'/></c:scaling>"
            f"<c:delete val='0'/><c:axPos val='{position}'/>"
            + ("<c:majorGridlines/>" if gridlines else "")
            + "<c:numFmt formatCode='General' sourceLinked='1'/>"
            "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
            "<c:tickLblPos val='nextTo'/>"
            f"<c:crossAx val='{cross_id}'/><c:crosses val='autoZero'/>"
            "<c:crossBetween val='midCat'/></c:valAx>"
        )

    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        "<c:scatterChart><c:scatterStyle val='lineMarker'/><c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Probe</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        f"<c:xVal>{cache(xs)}</c:xVal><c:yVal>{cache(ys)}</c:yVal>"
        "<c:smooth val='0'/></c:ser>"
        "<c:axId val='100002'/><c:axId val='100003'/></c:scatterChart>"
        + axis("100002", "100003", "b", gridlines=True)
        + axis("100003", "100002", "l", gridlines=False)
        + "</c:plotArea><c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        + (
            "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
            f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
            if size
            else ""
        )
        + "</c:chartSpace>"
    ).encode()


def chart_xml(low: float, high: float, size: int | None = None) -> bytes:
    """One scatter chart: five points from *low* to *high*, y gridlines, no x axis."""
    ys = [low, low + (high - low) * 0.4, low + (high - low) * 0.2, high, low]
    # The endpoints must be exactly the authored doubles, not a recomputed interpolation.
    ys[0], ys[3], ys[4] = low, high, low
    xs = [1, 2, 3, 4, 5]

    def cache(values):
        points = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
        return (
            "<c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef>"
        )

    def axis(axis_id, cross_id, position, *, gridlines, delete):
        return (
            f"<c:valAx><c:axId val='{axis_id}'/>"
            "<c:scaling><c:orientation val='minMax'/></c:scaling>"
            f"<c:delete val='{1 if delete else 0}'/><c:axPos val='{position}'/>"
            + ("<c:majorGridlines/>" if gridlines else "")
            + "<c:numFmt formatCode='General' sourceLinked='1'/>"
            "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
            "<c:tickLblPos val='nextTo'/>"
            f"<c:crossAx val='{cross_id}'/><c:crosses val='autoZero'/>"
            "<c:crossBetween val='midCat'/></c:valAx>"
        )

    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        "<c:scatterChart><c:scatterStyle val='lineMarker'/><c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Probe</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        f"<c:xVal>{cache(xs)}</c:xVal><c:yVal>{cache(ys)}</c:yVal>"
        "<c:smooth val='0'/></c:ser>"
        "<c:axId val='100002'/><c:axId val='100003'/></c:scatterChart>"
        + axis("100002", "100003", "b", gridlines=False, delete=True)
        + axis("100003", "100002", "l", gridlines=True, delete=False)
        + "</c:plotArea><c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        + (
            "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
            f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
            if size
            else ""
        )
        + "</c:chartSpace>"
    ).encode()


def bar_chart_xml(high: float, size: int | None = None) -> bytes:
    """One clustered column chart, five categories, the largest bar at *high*.

    The type the value-axis observations in :data:`AXIS_HALVING_RATIO` came from, kept as
    close to those probes as it can be: a category axis along the bottom, gridlines on the
    value axis, nothing else styled.
    """
    values = [high * f for f in (0.3, 1.0, 0.55, 0.8, 0.45)]
    values[1] = high
    points = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
    cats = "".join(f"<c:pt idx='{i}'><c:v>C{i + 1}</c:v></c:pt>" for i in range(5))
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
        "<c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Probe</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='5'/>{cats}</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='5'/>{points}</c:numCache></c:numRef></c:val></c:ser>"
        "<c:axId val='100002'/><c:axId val='100003'/></c:barChart>"
        "<c:catAx><c:axId val='100002'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='l'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        "<c:crosses val='autoZero'/><c:crossBetween val='between'/></c:valAx>"
        "</c:plotArea><c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        + (
            "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
            f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
            if size
            else ""
        )
        + "</c:chartSpace>"
    ).encode()


def category_chart_xml(
    kind: str,
    high: float,
    size: int | None = None,
    labels: tuple[str, ...] | None = None,
    legend: str | None = None,
    title: str | None = None,
    face: str | None = None,
) -> bytes:
    """One category chart of *kind* -- ``col``, ``bar``, ``line`` or ``area``.

    The same five values and the same unstyled axes as :func:`bar_chart_xml`, with the
    plot type and the category labels as the only variables.  ``bar`` puts the **value**
    axis along the bottom, which is the one case where a label's own width could gate the
    tick count; ``labels`` widens the category labels without touching anything else.
    """
    names = labels or tuple(f"C{i + 1}" for i in range(5))
    factors = (0.3, 1.0, 0.55, 0.8, 0.45)
    values = [high * factors[i % len(factors)] for i in range(len(names))]
    values[1] = high
    points = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
    cats = "".join(f"<c:pt idx='{i}'><c:v>{name}</c:v></c:pt>" for i, name in enumerate(names))
    series = (
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Probe</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='{len(names)}'/>{cats}</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef></c:val>"
        + ("<c:smooth val='0'/>" if kind == "line" else "")
        + "</c:ser>"
    )
    if kind in ("col", "bar"):
        plot = (
            f"<c:barChart><c:barDir val='{kind}'/><c:grouping val='clustered'/>"
            f"<c:varyColors val='0'/>{series}"
            "<c:axId val='100002'/><c:axId val='100003'/></c:barChart>"
        )
    elif kind == "line":
        plot = (
            "<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            f"{series}<c:marker val='1'/>"
            "<c:axId val='100002'/><c:axId val='100003'/></c:lineChart>"
        )
    elif kind == "area":
        plot = (
            "<c:areaChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            f"{series}<c:axId val='100002'/><c:axId val='100003'/></c:areaChart>"
        )
    elif kind == "radar":
        plot = (
            "<c:radarChart><c:radarStyle val='marker'/><c:varyColors val='0'/>"
            f"{series}<c:axId val='100002'/><c:axId val='100003'/></c:radarChart>"
        )
    else:
        raise SystemExit(f"unknown chart kind {kind!r}")
    # A horizontal bar chart swaps which axis is along the bottom; everything else --
    # the gridlines, the tick marks, the label positions -- is held.
    cat_pos, val_pos = ("l", "b") if kind == "bar" else ("b", "l")
    head = (
        (
            "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>"
            f"<a:t>{title}</a:t></a:r></a:p></c:rich></c:tx>"
            "<c:overlay val='0'/></c:title><c:autoTitleDeleted val='0'/>"
        )
        if title
        else "<c:autoTitleDeleted val='1'/>"
    )
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:chart>" + head + "<c:plotArea><c:layout/>"
        + plot
        + "<c:catAx><c:axId val='100002'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='0'/><c:axPos val='{cat_pos}'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='0'/><c:axPos val='{val_pos}'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        "<c:crosses val='autoZero'/>"
        + ("<c:crossBetween val='midCat'/>" if kind == "area" else
           "<c:crossBetween val='between'/>")
        + "</c:valAx>"
        "</c:plotArea>"
        + (
            f"<c:legend><c:legendPos val='{legend}'/><c:overlay val='0'/></c:legend>"
            if legend
            else ""
        )
        + "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        + (
            "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
            f"<a:defRPr sz='{size}'>"
            + (f"<a:latin typeface='{face}'/>" if face else "")
            + "</a:defRPr></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
            if size
            else ""
        )
        + "</c:chartSpace>"
    ).encode()


def slide_xml(index: int, probe: dict) -> str:
    cx, cy = probe["frame"]
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='{100 + index}' "
        f"name='{probe['key']}'/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        f"<p:xfrm><a:off x='{FRAME_OFF[0]}' y='{FRAME_OFF[1]}'/>"
        f"<a:ext cx='{cx}' cy='{cy}'/></p:xfrm>"
        "<a:graphic><a:graphicData "
        "uri='http://schemas.openxmlformats.org/drawingml/2006/chart'>"
        f"<c:chart {C} {R} r:id='rIdChart'/></a:graphicData></a:graphic></p:graphicFrame>"
        "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


def slide_rels(index: int) -> str:
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        f"<Relationship Id='rId1' Type='{LAYOUT_REL}' Target='../slideLayouts/slideLayout1.xml'/>"
        f"<Relationship Id='rIdChart' Type='{CHART_REL}' Target='../charts/probe{index}.xml'/>"
        "</Relationships>"
    )


def blank_template(xml: str) -> str:
    """Empty a layout or master so nothing but the chart is drawn."""
    xml = re.sub(r"<p:spTree>.*?</p:spTree>", EMPTY_TREE, xml, flags=re.S)
    return re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)


def probes_for(target: Path) -> list[dict]:
    """The probe table this deck's file name names."""
    for name, table in DECKS.items():
        if name in target.stem:
            return table
    raise SystemExit(f"name the deck after one of {sorted(DECKS)}, not {target.stem!r}")


def write_deck(target: Path, probes: list[dict], part_for=None) -> None:
    """One slide per probe, each holding one chart, written to *target*.

    ``part_for`` builds the chart part for a probe row and defaults to this deck's own
    :func:`chart_part`; ``tools/make_view3d_probe.py`` passes its own so the slide,
    relationship and content-type scaffolding is written in one place only.
    """
    part_for = part_for or chart_part
    source = zipfile.ZipFile(SOURCE)
    presentation = source.read("ppt/presentation.xml").decode()
    pres_rels = source.read("ppt/_rels/presentation.xml.rels").decode()
    content_types = source.read("[Content_Types].xml").decode()
    numbers = range(1, len(probes) + 1)

    presentation = re.sub(
        r"<p:sldIdLst>.*?</p:sldIdLst>",
        "<p:sldIdLst>"
        + "".join(f'<p:sldId id="{255 + n}" r:id="rIdS{n}"/>' for n in numbers)
        + "</p:sldIdLst>",
        presentation,
        flags=re.S,
    )
    pres_rels = re.sub(r'<Relationship[^>]*Type="[^"]*/slide"[^>]*/>', "", pres_rels)
    pres_rels = pres_rels.replace(
        "</Relationships>",
        "".join(
            f'<Relationship Id="rIdS{n}" Type="{SLIDE_REL}" Target="slides/slide{n}.xml"/>'
            for n in numbers
        )
        + "</Relationships>",
    )
    content_types = re.sub(
        r'<Override PartName="/ppt/(slides/slide|charts/chart)\d+\.xml"[^>]*/>', "", content_types
    )
    content_types = content_types.replace(
        "</Types>",
        "".join(
            f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{SLIDE_TYPE}"/>'
            f'<Override PartName="/ppt/charts/probe{n - 1}.xml" ContentType="{CHART_TYPE}"/>'
            for n in numbers
        )
        + "</Types>",
    )

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            name = item.filename
            if name.startswith("ppt/slides/") or name.startswith("ppt/charts/"):
                continue
            data = source.read(name)
            if name == "ppt/presentation.xml":
                data = presentation.encode()
            elif name == "ppt/_rels/presentation.xml.rels":
                data = pres_rels.encode()
            elif name == "[Content_Types].xml":
                data = content_types.encode()
            elif re.search(r"slide(Layouts|Masters)/slide\w+\d+\.xml$", name):
                data = blank_template(data.decode()).encode()
            out.writestr(name, data)
        for index, probe in enumerate(probes):
            out.writestr(f"ppt/slides/slide{index + 1}.xml", slide_xml(index, probe))
            out.writestr(f"ppt/slides/_rels/slide{index + 1}.xml.rels", slide_rels(index))
            out.writestr(f"ppt/charts/probe{index}.xml", part_for(probe))


def chart_part(probe: dict) -> bytes:
    """The chart part one probe row asks for."""
    if probe.get("read") == "x":
        return scatter_xy_xml(probe["low"], probe["high"], probe.get("size"))
    if probe.get("kind"):
        return category_chart_xml(
            probe["kind"],
            probe["high"],
            probe.get("size"),
            probe.get("labels"),
            probe.get("legend"),
            probe.get("title"),
            probe.get("face"),
        )
    if probe.get("bar"):
        return bar_chart_xml(probe["high"], probe.get("size"))
    return chart_xml(probe["low"], probe["high"], probe.get("size"))


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else Path("axis-decade.pptx")
    probes = probes_for(target)
    write_deck(target, probes)
    for index, probe in enumerate(probes):
        span = probe["high"] - probe["low"]
        print(
            f"{index + 1:3d} {probe['key']:12s} {probe['low']!r:8s}..{probe['high']!r:22s} "
            f"span={span!r:24s} short={1 - span / probe['decade']:.2e} "
            f"frame={probe['frame'][1] / 12700:.0f}pt"
        )
    print(f"\n{len(probes)} probes -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

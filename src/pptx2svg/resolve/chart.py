"""Chart layout and drawing: ``SourceChart`` in, ordinary slide elements out.

Unlike SmartArt there is nothing cached to lean on -- ``c:chartSpace`` is data plus
styling, and every position in the picture has to be computed.  The output is a plain
list of :class:`~pptx2svg.model.ShapeElement`, so the SVG writer, the text engine, the
fill resolver and the font substitution all apply to a chart exactly as they do to a
slide.  Nothing here knows about SVG.

**Every constant below was measured out of PowerPoint's own PDF export**, not reasoned
about.  Twelve probe charts differing in one input each -- title, legend position on all
four sides, font size at 8/10/14 pt, gap width, tick marks, series count -- were laid out
at an identical frame size and exported by PowerPoint 16.106; the plot rectangle, the
bars and every text baseline come out of the PDF as exact vector coordinates, so the
formulas could be fitted rather than guessed.  Each one records its residual against the
measurement.  Two font families are represented (Aptos and Arial, from the corpus), which
is what allowed the size-proportional and metric-proportional terms to be separated.

The layout PowerPoint actually performs, as far as the measurements can tell:

    frame
    ├─ title band            (only when there is a title)
    ├─ legend band           (only for legendPos t)
    │
    │   value labels │ plot area                     │ legend band (legendPos l/r)
    │
    ├─ category label band
    └─ legend band           (only for legendPos b)

Three things that look like bugs and are not:

* **The tick-mark allowance is reserved whether or not tick marks are drawn.**  A probe
  with ``majorTickMark="none"`` and one with ``"out"`` produced *byte-identical* plot
  rectangles, so the space is part of the layout rather than part of the tick.
* **The default axis and gridline colour is black, not grey**, and the default width is
  0.5 pt.  Every probe drew ``0 0 0 SC`` at ``6350 w``.  Charts written by modern
  PowerPoint carry a ``c:style`` or a chart-style part that overrides this to grey; none
  of the decks measured here does, and the measurement wins.
* **Axis labels default to 10 pt in the theme's minor font, but a ``c:rich`` title does
  not.**  In ``authoring-integration.pptx`` PowerPoint drew the axis labels in Aptos (the
  theme minor face) at 10 pt and the title, whose ``a:rPr`` names nothing at all, in
  Arial at 18 pt -- the same fallback it gives any unstyled DrawingML text.  So the title
  is resolved through the ordinary text cascade and the rest is not.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field, replace

from .. import model as m
from ..parse import chart as c
from ..parse import source as s
from ..text.fontmap import east_asian_family, metrics_for
from ..text.measure import DEFAULT_LINE_HEIGHT_RATIO, is_cjk

EMU_PER_POINT = 12700.0

_SIN_45 = math.sin(math.radians(45.0))

#: Padding between the frame edge and the outermost label block, in points.  Measured as
#: the left edge of the value-label column in all three decks and every probe: exactly
#: 6.5 pt, independent of frame size and font size.
FRAME_PADDING_PT = 6.5

#: The plot area's inset from the frame on a side with nothing on it.  Measured at
#: 10.9995 pt on the right of every probe and of ``authoring-integration.pptx``; it is
#: also the floor on the top inset (see :func:`_top_inset`).
EDGE_INSET_PT = 11.0

#: How far above its baseline a one-line label's optical centre sits, in ems.  Fitted;
#: see :attr:`FontBox.ink_centre` for the five measurements and why no exact rule emerged.
LABEL_INK_CENTRE_EM = 0.27

#: Gap between the right edge of the value labels and the value axis, over and above the
#: font's descent.  Fitted to 0.645 em across Aptos at 8/10/14 pt and Arial at 12 pt with
#: a residual under 0.05 pt in all four -- the tightest fit in this file.
VALUE_LABEL_GAP_EM = 0.645

#: Slack between the category-label line and the plot area, over and above the label's
#: own line height.  Fitted to 0.615 em over the same four cases; residual under 0.16 pt.
CATEGORY_LABEL_GAP_EM = 0.615

#: The plot area's inset above the topmost value label.  ``max(11.0, 5.0 + lineHeight/2)``
#: reproduces all four measurements to within 0.02 pt, including the 8 pt probe where the
#: floor is what binds.
TOP_INSET_BASE_PT = 5.0

#: A legend row's height, as a multiple of the line height.  Measured 24.083 pt for a
#: 10 pt Aptos legend against a 12.207 pt line height (1.973x); 2.0 is within 0.33 pt and
#: the same band was measured for ``legendPos`` ``t`` and ``b``.
LEGEND_BAND_LINES = 2.0

#: The vertical pitch between stacked legend entries, in ems.  Measured 18.0 pt for a
#: 10 pt legend on all three right-hand legends in real-financial-report.pptx -- which is
#: *not* the same as the horizontal band's height above, so the two are separate numbers.
LEGEND_ROW_PITCH_EM = 1.8

#: How much the pitch opens for each extra line once an entry wraps.  Measured once: the
#: same deck's doughnut, whose one two-line entry takes every row from 1.8 em to 3.02 em.
#: Three or more lines is **extrapolated, not measured**.
LEGEND_WRAPPED_PITCH_EM = 1.22

#: Where a horizontal legend's baseline sits, in ems from the frame edge it hugs.  Taken
#: straight off the probes rather than derived from the band: 12.913 pt above the frame
#: bottom for ``legendPos="b"`` and 17.133 pt below the frame top for ``"t"``, both at
#: 10 pt.  The two are not symmetric and no rule was found that makes them so.
LEGEND_BOTTOM_BASELINE_EM = 1.291
LEGEND_TOP_BASELINE_EM = 1.713

#: Legend swatch side and the gap after it, in ems.  Measured 5.4923 pt and 2.3711 pt at
#: 10 pt.
LEGEND_SWATCH_EM = 0.549
LEGEND_SWATCH_GAP_EM = 0.237
#: The most of the frame's width a side legend may take before its entries wrap.
#: **One measurement**: `real-financial-report.pptx`'s doughnut legends an 11-character
#: Japanese category whose natural band would be 143.96 pt, and PowerPoint reserved
#: 113.98 pt -- 40.0% of the 285 pt frame -- wrapping the entry onto two lines instead.
#: The same deck's bar chart, whose natural band is 29% of its frame, is untouched by it.
LEGEND_SIDE_MAX_FRACTION = 0.40

#: Padding either side of a side legend, and between entries in a horizontal one.
LEGEND_SIDE_LEAD_EM = 1.60
LEGEND_SIDE_TRAIL_EM = 1.01
LEGEND_ENTRY_GAP_EM = 0.5

#: A horizontal legend's run of entries is centred on the frame with this much lead-in
#: counted as part of it, which shifts the visible entries half of it to the right.
#: Measured at 1.93 pt of shift for a 10 pt legend, identically on the legend-b and
#: legend-t probes.
LEGEND_HORIZONTAL_LEAD_EM = 0.386

#: The title band, and its baseline inside it, as multiples of the line height and the
#: ascent.  Only one title was measurable (18 pt Arial, in two probes and the fixture, all
#: agreeing): the band is 29.70 pt against a 20.109 pt line height, and the baseline sits
#: 24.52 pt below the frame top against a 16.295 pt ascent.  These are the measured
#: ratios rather than the tidy 1.5 both are close to -- rounding the band cost 0.46 pt of
#: plot height, which moved every gridline by a pixel.  They are a one-font fit and should
#: be re-measured if a chart with a differently sized title ever disagrees.
TITLE_BAND_LINES = 1.4769
TITLE_BASELINE_ASCENTS = 1.5046

#: Default chart text size, in points.  ECMA-376's chart default and what PowerPoint drew
#: for every axis label and legend entry with no ``c:txPr``.
DEFAULT_CHART_FONT_PT = 10.0

#: Axis, tick and gridline defaults when no ``c:spPr`` says otherwise.
DEFAULT_AXIS_LINE_EMU = 6350.0
DEFAULT_AXIS_COLOR = "#000000"

#: A line series' stroke when its ``a:ln`` states no width -- 1.5 pt, measured.  A series
#: that *does* state one is taken literally: the real line chart in
#: ``real-financial-report.pptx`` says ``w="25400"`` and PowerPoint drew 2 pt.
DEFAULT_LINE_SERIES_WIDTH_EMU = 19050.0

#: A line series' marker outline when it states none -- 0.5 pt in the series' own colour.
DEFAULT_MARKER_OUTLINE_EMU = 6350.0

#: Marker symbols a line series cycles through when it states no ``c:marker`` of its own.
#: Measured only for series 0, which came out a **diamond** -- not the circle most
#: implementations assume.  The rest of the cycle is from the specification and is
#: **not measured**.
DEFAULT_MARKER_CYCLE = ("diamond", "square", "triangle", "x", "star", "dot")

#: Marker side when ``c:size`` is absent, in points.  **Six, not ECMA-376's seven.**
#:
#: This was carried as 7 with a note that it had never been measured either way, while the
#: radar path already used 6 from a probe whose diamond came out 5.76 pt across -- a
#: diamond, so the reading depended on the tips landing inside PowerPoint's 0.24 pt output
#: grid.  Two probes settle it with an axis-aligned shape and no such argument: the
#: scatter two-series probe's second series states no ``c:marker`` and its **square**
#: measured exactly 6.000 x 6.000, and a line chart with an explicit ``c:size val="6"``
#: square measured the same 6.000.  The line chart's own default diamond measured 5.76,
#: the radar's number, on the same export.
DEFAULT_MARKER_SIZE_PT = 6.0

#: Gap between a bar's edge and the *line box* of the data label beside it, in points.
#: Measured 4.86 pt at 10 pt and 4.70 pt at 14 pt for ``outEnd``, and 4.91 pt at 10 pt for
#: ``inBase`` -- so it is a fixed distance and **not** proportional to the font.
DATA_LABEL_GAP_PT = 4.85

#: The same for ``inEnd``, which sits closer to the bar's end.  One measurement, at 10 pt.
DATA_LABEL_INNER_GAP_PT = 4.05

#: Gap between a line chart's marker edge and its data label, in ems.  The label is
#: centred on the point vertically and sits to its right -- ECMA's ``r`` default, which is
#: what PowerPoint drew.  Measured once, with a 7 pt marker.
DATA_LABEL_LINE_GAP_EM = 0.6

#: Category labels rotate when the widest **unbreakable token** in them is wider than the
#: band it has to sit in.  Wrapping comes first: PowerPoint turns a label only when
#: breaking it would not save it.
#:
#: Two probes straddle exactly one band and both give the same ratio, one for a label with
#: no break in it and one for a label whose break does not help:
#:
#: * whole label, no space: 36.62 pt level against 38.59 pt turned on a 37.68 pt band --
#:   (0.972, 1.024];
#: * widest token, with a space after it: ``xxxxxxxi M`` at 37.22 pt wrapped and
#:   ``HHHHHi M`` at 38.33 pt turned on a 37.761 pt band -- (0.9857, 1.0151].
#:
#: One band width is the only value in both windows.  ``Fiscal MMMMMMMMMMMM 2012`` is the
#: case that separates this from every other candidate rule: it has two spaces in it and
#: PowerPoint turned it anyway, because its middle token is 99.96 pt on that same band.
#: Neither fixture carries an explicit ``rot=`` on ``a:bodyPr``, so this is PowerPoint's
#: own decision.
ROTATED_LABEL_RATIO = 1.0

# A category label breaks at a space and nowhere else.  `MMM-MM`, `MMM/MM`, `MMM,MM`,
# `MMM_MM` and `MMM<en dash>MM` were each 44-47 pt on a 37.76 pt band and PowerPoint
# turned all five rather than breaking them, and a CJK label with no space in it turned
# too -- so this is *not* Unicode line breaking, it is whitespace.  The one surprise is
# U+00A0: `MMM<nbsp>MM` broke at the no-break space exactly as a plain space did, which is
# why `wrap_label` splits with `str.split` rather than on `" "`.  Tab and the other
# whitespace `str.split` honours are **not measured**; only U+0020 and U+00A0 are.
#
# This is a category-axis rule, not a general one: a radar facing the same problem wraps
# but never turns.

#: How many lines a wrapped category label may take.  Measured linear through six: on the
#: Arial ladder the band grew by exactly one line height for every extra line up to six.
#: Past that PowerPoint stops wrapping altogether -- an eight-token label came back on
#: **two** lines, each one four band widths wide and overlapping its neighbours, and a
#: twelve- and a twenty-four-token label did the same.  No rule reproduces both the
#: linear part and that collapse, so the band stops growing where the measurements stop
#: rather than extrapolating into a corner PowerPoint does not agree with.
WRAPPED_LABEL_MAX_LINES = 6

#: And the angle it turns to.  **It snaps.**  Twelve probes from a label 1.02 band widths
#: wide to one 4.18 wide all came out at exactly 45 degrees, reading up to the right --
#: `rot="-2700000"` in DrawingML terms.  No intermediate angle appeared anywhere in that
#: range, and nothing went to 90.
ROTATED_LABEL_DEGREES = -45.0

#: The plot's bottom inset once the labels turn: this, plus the widest label's width times
#: sin 45.  Fitted to six probes across two decks, worst residual **0.03 pt** -- and the
#: residual is that small only because the *widest* label is the one that sets it, which
#: is what a 1.01 pt discrepancy on the deck whose five labels differ by one letter
#: showed.  It replaces the horizontal band's `6.5 + lineHeight + 0.615 em` entirely.
ROTATED_LABEL_INSET_PT = 21.39

#: Where the rotated baseline's far end lands, relative to the centre of its band on the
#: category axis: this far right, and this far below.  Measured on six probes, spread
#: under 0.15 pt.
ROTATED_LABEL_OFFSET_X_PT = 2.0
ROTATED_LABEL_OFFSET_Y_PT = 12.7

#: Chart kinds laid out around a centre rather than on a pair of axes.
POLAR_CHART_KINDS = frozenset({"pieChart", "doughnutChart", "radarChart"})

#: The polar kinds that are a web of spokes rather than a ring of slices.
RADAR_CHART_KINDS = frozenset({"radarChart"})

#: Filled to the zero line rather than stroked through the points.
AREA_CHART_KINDS = frozenset({"areaChart"})

#: The one Cartesian kind with **two value axes and no category axis**.
SCATTER_CHART_KINDS = frozenset({"scatterChart"})

#: ``c:crossBetween`` decides whether the points sit at the centres of the category bands
#: or on the band edges, and it is optional.  **An area chart that states none draws as
#: ``midCat``** -- the first vertex on the plot's left edge and the last on its right.
#: Measured: a probe with `<c:crossBetween
#: val="between"/>` put its three vertices at 52.473 / 115.292 / 178.072 on a
#: 21.073..209.472 plot, which are the band centres; one with `val="midCat"` and one with
#: **no element at all** were byte-identical to each other at 26.433 / 108.015 / 189.632,
#: which are the plot's edges and midpoint.  A line chart's absent case is *not* measured
#: -- every line chart in the corpus and in every probe states ``between`` -- so the
#: default below stays ``between`` for everything but an area.
DEFAULT_AREA_CROSS_BETWEEN = "midCat"

#: How far the drawn radius falls short of half the plot region, as a function of the
#: category labels' line height.  Fitted to five probes -- Aptos at 8, 10 and 14 pt and
#: Arial at 10 and 14 pt, all with one-line labels -- whose worst residual is 0.089 pt:
#:
#: ====== ==== ========= =========
#: face   size line box  reserve
#: ====== ==== ========= =========
#: Aptos   8    9.766     7.071
#: Aptos  10   12.207    10.081
#: Aptos  14   17.090    16.191
#: Arial  10   11.172     8.751
#: Arial  14   15.641    14.511
#: ====== ==== ========= =========
#:
#: The slope is **not** 1: the reserve grows faster than the line box, which is why no
#: "leave one line of room" rule reproduces the set.  A sixth probe pins the other end --
#: with ``<c:delete val="1"/>`` on the category axis, and so no labels at all, the radius
#: came out 79.44 pt against a half-region of 79.551, i.e. the reserve goes to zero.
RADAR_LABEL_RESERVE_LINES = 1.2578
RADAR_LABEL_RESERVE_PT = 5.2501

#: Gap between a polygon vertex and the category label pushed radially out from it.
#: Fitted to the four measurements taken where the direction is horizontal and the label
#: box is therefore unambiguous -- 2.775, 2.789, 2.84 and 2.980 pt at 10 pt -- residual
#: under 0.15 pt.  It is **not** proportional to the size: the same gap came out 2.81 pt
#: at 8 pt and 2.38 pt at 14 pt.  The two vertical directions are looser, 4.49 pt above
#: the top vertex and 2.29 pt below the bottom one, and the split is consistent with
#: PowerPoint's line box being about 1 pt taller than the one our metrics give -- a font
#: discrepancy rather than a second layout rule, so one constant is used for all four.
RADAR_LABEL_GAP_PT = 2.85

#: The most of the plot region's width one category label may take before it wraps onto
#: another line.  Bracketed by two probes on the same 198.47 pt region: a 43.72 pt label
#: ("Two Three") stayed on one line and a 59.10 pt one ("Category One") wrapped, which is
#: (0.2203, 0.2978]; 0.25 is the round number inside it.  Wrapping at that cap reproduces
#: PowerPoint's own break exactly -- "Category Three" came out "Category" / "Three".
RADAR_LABEL_MAX_FRACTION = 0.25

#: Where the value-axis labels sit: right-aligned, with their right edge this many widths
#: of the digit zero to the left of the twelve o'clock spoke.  Measured on six charts --
#: Aptos at 8, 10 and 14 pt, Arial at 10 and 14 pt, and the Arial 12 pt value axis of
#: ``real-financial-report.pptx``'s own radar -- and it is exactly two digits every time,
#: worst residual 0.14 pt.  Reading it as a plain em fraction does not work: it is
#: 1.069 em in Aptos and 1.112 em in Arial, and the difference is exactly twice the
#: difference between the two faces' digit widths.
RADAR_VALUE_LABEL_DIGITS = 2.0

#: A radar series' marker when it states no ``c:size``.  Measured 6.0 pt square on the
#: probe's second series; the first series' diamond measured 5.76 pt across, which is the
#: same 6 pt box with its tips falling inside PowerPoint's 0.24 pt output grid.  This is
#: **not** :data:`DEFAULT_MARKER_SIZE_PT`, which is ECMA-376's 7 and has never been
#: measured for a line chart either way.
RADAR_MARKER_SIZE_PT = 6.0

#: A line chart's legend key -- a line of the series' own stroke with its marker at the
#: middle, rather than a bar chart's square swatch -- and the gap after it, **in points**.
#: A ``standard`` or ``marker`` radar takes the same key; a ``filled`` radar legends with
#: the ordinary swatch instead, which is what ``real-financial-report.pptx``'s own radar
#: draws.
#:
#: Measured first on the radar legend probe at 10 pt, then on four line-chart legends:
#: 19.200 pt of rule and 2.025 pt before the text at 10 pt on a right-hand legend, the
#: same on a bottom one, the same for a series with no marker, and **the same at 14 pt**.
#: That last one is the refutation: the radar's single 10 pt measurement was carried as
#: 1.920 and 0.2025 *ems* on the assumption that it scaled, and it does not.
#:
#: Read off a stroked path's bounding box, which a 3 pt round cap makes 3 pt longer than
#: the rule: 22.200 pt of box is 19.200 pt of line, and the marker centre landed within
#: 0.17 pt of its midpoint.
LINE_LEGEND_KEY_PT = 19.200
LINE_LEGEND_KEY_GAP_PT = 2.025

#: ``c:holeSize`` when the element is absent.  ECMA-376 documents a default of 10; what
#: PowerPoint *draws* for a `c:doughnutChart` stating no `c:holeSize` is a **solid pie**,
#: measured on the probe -- the wedge closes through the centre with no inner arc.
DEFAULT_HOLE_SIZE = 0.0

#: Where a slice's data label sits, as a fraction of the outer radius along the slice's
#: bisector.  All four measured on one probe pie whose labels all fit comfortably:
#: ``ctr`` is exactly half the radius, the rest are means over four slices.  PowerPoint's
#: real ``bestFit`` moves a label out of the way when it does not fit, which this does not
#: reproduce and which was **not** measured.
PIE_LABEL_RADIUS = {
    "ctr": 0.500,
    "inEnd": 0.856,
    "outEnd": 1.020,
    "bestFit": 0.710,
}

#: The outline on a negative bar drawn hollow by ``c:invertIfNegative`` -- 0.75 pt.
INVERTED_BAR_OUTLINE_EMU = 9525.0

#: Floor on the divisor that turns a category band into one bar's width.  ``c:gapWidth``
#: is schema-bounded to 0..500 and files need not obey; -100 on a single series makes the
#: divisor zero exactly.
MIN_BAR_SLOTS = 0.01

#: A hand-written ``c:majorUnit`` can ask for an unbounded number of gridlines, so the
#: tick list is capped rather than trusted.
MAX_MAJOR_TICKS = 1000

#: ``c:gapWidth`` when absent, in percent of one bar's width (ECMA-376 default).
DEFAULT_GAP_WIDTH = 150.0

#: The six theme accents a series cycles through when it has no fill of its own.
ACCENT_KEYS = ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6")

#: Below this many major units of span, the plain power of ten is halved.  See
#: :func:`nice_axis_scale`; the threshold is somewhere in (1.842, 4.285] and 2 is the
#: round number inside it.
#:
#: **A ladder was written to replace this and reverted.**  A scatter probe of 120..160
#: came back 0..180 **by 20** where one halving gives 0..200 by 50, and stepping the unit
#: down the 1-2-5 ladder while the span holds fewer than ~3.5 units reproduces that *and*
#: all five observations this constant was fitted to.  Then the corpus radar refutes it:
#: `real-financial-report.pptx`'s chart5 has 65..100 of data and PowerPoint's own export
#: draws **two** rings, at radii 22.8 and 45.6 -- 0..100 by 50, a ratio of exactly 2.0
#: accepted, where the scatter refused 3.2 on the same kind of axis.  No monotone ratio
#: threshold produces both.
#:
#: What separates them is **axis length**, which makes both of these observations belong
#: to the unsolved tick-density question rather than to unit selection: the radar's axis
#: is 45.6 pt and takes 22.8 pt steps, the scatter's is 145.0 pt and takes 16.1 pt steps
#: where 36.3 pt was available.  A target band of roughly 16 to 24 pt fits those two and
#: every cell of the density table in ROADMAP.md -- and then dies on the same stacked
#: probe that killed the last candidate, which accepts 14.5 pt.  See ROADMAP.md.
AXIS_HALVING_RATIO = 2.0

#: Data that sits this far up its own range does not get an axis pulled back to zero.
#:
#: A **bar** must start at its axis -- a bar that does not is a different picture -- so
#: `nice_axis_scale` anchors at zero for every type but one.  A **scatter** does not: a
#: decade of years against a measurement would be destroyed by it, and PowerPoint agrees.
#: Measured on six probes, and the threshold is Excel's folklore 5/6 = 0.8333 with the
#: bracket [0.80, 0.84) around it: x of 40..50 came back **0..60** (40/50 = 0.80) and
#: x of 42..50 came back **40..55** (42/50 = 0.84).  The other four agree -- 2010..2020
#: -> 2005..2025, 100..104 -> 98..106, 120..160 -> 0..180, 10..50 -> 0..60 -- and the
#: all-negative case mirrors it: -50..-10 came back -60..0.
#:
#: **Only a scatter is measured.**  A line chart of temperatures has the same problem and
#: no probe has ever shown what PowerPoint does with one, so it keeps the zero anchor.
AXIS_ZERO_ANCHOR_RATIO = 5.0 / 6.0

#: A value axis **along the bottom** comes out coarser than one up the side for the same
#: data and the same axis length, so once the interval is chosen it is stepped up until
#: the axis holds no more than this many of them.  What it really is remains unknown --
#: it is not a density limit, because the horizontal probe's 0..5 data came out 0..6 by 2
#: where the identical data on a *vertical* axis of almost the same length (151.4 pt
#: against 145.0 pt) came out 0..6 by 1.
#:
#: **Four, not five.**  It was five on that one horizontal-bar observation, which only
#: bounds it below six.  Four scatter probes -- whose x axis is the same bottom axis --
#: bracket it properly, and the discriminating one is ``x-float``: x from 0.5 to 4.5
#: rounds to a 0..5 axis at unit 1, which is *five* intervals, and PowerPoint coarsened it
#: to 0..6 by 2 anyway.  ``x-neg``'s -4..4 by 2 is **four** intervals and PowerPoint kept
#: it, so the bracket is [4, 5).  The other two, 1..5 and 10..50, agree at either value.
HORIZONTAL_MAX_INTERVALS = 4


# --------------------------------------------------------------------------------------
# Axis scaling
# --------------------------------------------------------------------------------------


def nice_axis_scale(
    data_minimum: float,
    data_maximum: float,
    horizontal: bool = False,
    strict: bool = True,
    anchor_zero: bool = True,
) -> tuple[float, float, float]:
    """``(minimum, maximum, major_unit)`` for a value axis PowerPoint would draw itself.

    The major unit is the plain **power of ten** just below the span, halved when the span
    is less than :data:`AXIS_HALVING_RATIO` of it.  That is not the "aim for N ticks" rule
    every charting library uses, and the difference is not cosmetic -- N ticks cannot
    produce both of these, which PowerPoint does:

    ===========  ==============  ==========
    data         PowerPoint      intervals
    ===========  ==============  ==========
    0..5         0..6 by 1       6
    0..9         0..10 by 1      10
    -2..5        -3..6 by 1      9
    0..1842      0..2000 by 500  4
    0..4285      0..5000 by 1000 5
    ===========  ==============  ==========

    A sixth observation -- 120..160 of scatter data drawn 0..180 **by 20** on a 145 pt
    axis, where this gives 0..200 by 50 -- is **not** reproduced, and deliberately so: the
    rule that reproduces it contradicts the corpus radar.  See :data:`AXIS_HALVING_RATIO`.

    Both ends are rounded *strictly* outwards, so a series topping out at exactly 5 gets
    an axis to 6 rather than one whose last bar touches the frame.  Both bumps are
    measured: the first is what ``authoring-integration.pptx`` does, the second is the -3
    on the negative-value probe whose data floor is -2.

    ``strict=False`` turns that outward bump off, which is what a **radar** wants: the
    same 0..5 data a bar chart takes to 6 stopped at exactly 5 on every radar probe, five
    rings with the outermost passing through the largest point.  One discriminating
    observation, and it is the whole of the difference -- the unit is chosen identically.

    ``anchor_zero=False`` lets the domain leave zero out when the data sits far enough up
    its own range; see :data:`AXIS_ZERO_ANCHOR_RATIO`.  Only a **scatter** passes it, and
    only a scatter has been measured.
    """
    low, high = min(0.0, data_minimum), max(0.0, data_maximum)
    if not anchor_zero and _floats_away_from_zero(data_minimum, data_maximum):
        low, high = data_minimum, data_maximum
    anchored = low <= 0.0 <= high
    span = high - low

    if span <= 0 or not math.isfinite(span):
        # Every value zero (or unusable).  PowerPoint still draws an axis; 0..1 is the
        # smallest one that shows anything.
        return 0.0, 1.0, 1.0

    unit = 10.0 ** math.floor(math.log10(span))
    # A denormal span underflows the power of ten to zero; a span at the other end
    # overflows the strictly-outward rounding below to infinity.  Neither is a chart
    # anyone drew on purpose, and both used to raise out of the conversion.
    if not math.isfinite(unit) or unit <= 0:
        return 0.0, 1.0, 1.0
    if span / unit < AXIS_HALVING_RATIO:
        unit /= 2

    minimum, maximum = _axis_extent(
        unit, low, high, data_minimum, data_maximum, strict, anchored
    )
    if horizontal:
        # Counted on the *rounded* extent, not the data span: 0..5 of data becomes a
        # 0..6 axis, and it is the six intervals in that which PowerPoint coarsens.
        while (maximum - minimum) / unit > HORIZONTAL_MAX_INTERVALS:
            stepped = _next_nice_unit(unit)
            if not math.isfinite(stepped) or stepped <= unit:
                break
            unit = stepped
            minimum, maximum = _axis_extent(
                unit, low, high, data_minimum, data_maximum, strict, anchored
            )
    if not (math.isfinite(minimum) and math.isfinite(maximum) and maximum > minimum):
        return 0.0, 1.0, 1.0
    return minimum, maximum, unit


def _axis_extent(
    unit: float,
    low: float,
    high: float,
    data_minimum: float,
    data_maximum: float,
    strict: bool = True,
    anchored: bool = True,
) -> tuple[float, float]:
    """Round the domain outwards to whole units, strictly past the data at both ends.

    ``strict=False`` rounds to a whole unit and stops there, which is the radar rule.

    ``anchored`` says the domain is held at zero, which is what stops the strict bump
    from pushing a 0..5 axis down to -1: zero is the floor, not a datum to clear.  An
    **unanchored** axis has no such floor and its low end bumps like its high end --
    measured on the scatter probes, where 100..104 came back **98**..106 and 2010..2020
    came back **2005**..2025, both a whole unit clear of the data at each end.
    """
    maximum = math.ceil(high / unit) * unit
    if strict and maximum <= data_maximum:
        maximum += unit
    minimum = math.floor(low / unit) * unit
    bump_low = data_minimum < 0 if anchored else True
    if strict and bump_low and minimum >= data_minimum:
        minimum -= unit
    return minimum, maximum


def _floats_away_from_zero(data_minimum: float, data_maximum: float) -> bool:
    """Whether the data sits far enough up its own range to leave zero off the axis.

    Both ends must be on the same side of zero, and the near end must be past
    :data:`AXIS_ZERO_ANCHOR_RATIO` of the far one.  Data that straddles zero always keeps
    it, because zero is already inside the domain.
    """
    if not (math.isfinite(data_minimum) and math.isfinite(data_maximum)):
        return False
    if data_minimum > 0 and data_maximum > 0:
        return data_minimum > AXIS_ZERO_ANCHOR_RATIO * data_maximum
    if data_minimum < 0 and data_maximum < 0:
        return data_maximum < AXIS_ZERO_ANCHOR_RATIO * data_minimum
    return False


def _next_nice_unit(unit: float) -> float:
    """The next step up the 1-2-5 ladder from a unit already on it."""
    magnitude = 10.0 ** math.floor(math.log10(unit))
    mantissa = round(unit / magnitude, 6)
    if mantissa < 2:
        return 2 * magnitude
    if mantissa < 5:
        return 5 * magnitude
    return 10 * magnitude


# --------------------------------------------------------------------------------------
# Number formatting
# --------------------------------------------------------------------------------------


def format_number(value: float, format_code: str | None) -> str:
    """Render one number the way its ``c:formatCode`` asks.

    A small subset of the Excel format language: enough for the codes that appear on real
    axes (``General``, ``#,##0``, ``0.0%``, ``0.00``) and a graceful fall-through for the
    rest.  A full implementation is a project of its own and belongs nowhere near here.
    """
    if format_code is None or format_code in ("General", "@"):
        return _general(value)

    sections = _sections(format_code)
    # Excel's sections are positive; negative; zero; text.  A negative value uses the
    # second only when there *is* one -- with a single section it is formatted by that
    # one and keeps its own minus sign.
    negative_section = value < 0 and len(sections) > 1
    section = sections[1] if negative_section else sections[0]

    if "%" in section:
        decimals = _decimals(section)
        magnitude = abs(value) if negative_section else value
        return _wrap_negative(f"{magnitude * 100:.{decimals}f}%", section, negative_section)

    if not any(ch in section for ch in "#0"):
        return _general(value)

    decimals = _decimals(section)
    grouped = "," in _strip_literals(section)
    # The negative section states its own sign -- "(#,##0)" or "-#,##0" -- so the value
    # goes in unsigned and the section's own decoration is put back around it.
    magnitude = abs(value) if negative_section else value
    text = f"{magnitude:,.{decimals}f}" if grouped else f"{magnitude:.{decimals}f}"
    return _wrap_negative(text, section, negative_section)


def _wrap_negative(text: str, section: str, negative_section: bool) -> str:
    if not negative_section:
        return text
    if "(" in section and ")" in section:
        return f"({text})"
    return f"-{text}" if "-" in section else text


def _general(value: float) -> str:
    if value == int(value):
        return str(int(value))
    return f"{round(value, 10):g}"


def _sections(format_code: str) -> list[str]:
    """Split on ``;`` outside quotes -- positive, negative, zero, text."""
    sections: list[str] = []
    current: list[str] = []
    quoted = False
    for char in format_code:
        if char == '"':
            quoted = not quoted
        if char == ";" and not quoted:
            sections.append("".join(current))
            current = []
            continue
        current.append(char)
    sections.append("".join(current))
    return sections


def _strip_literals(section: str) -> str:
    """Drop ``[red]`` directives, ``"text"`` and escapes so only the numeric shape is left."""
    out: list[str] = []
    index = 0
    while index < len(section):
        char = section[index]
        if char == "[":
            index = section.find("]", index)
            if index < 0:
                break
            index += 1
            continue
        if char == '"':
            index = section.find('"', index + 1)
            if index < 0:
                break
            index += 1
            continue
        if char == "\\":
            index += 2
            continue
        if char in "_*":
            index += 2
            continue
        out.append(char)
        index += 1
    return "".join(out)


def _decimals(section: str) -> int:
    stripped = _strip_literals(section)
    if "." not in stripped:
        return 0
    tail = stripped.split(".", 1)[1]
    count = 0
    for char in tail:
        if char in "0#":
            count += 1
        else:
            break
    return count


# --------------------------------------------------------------------------------------
# Text metrics
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class ChartFont:
    """The face a piece of chart text is drawn in, plus its metrics at that size.

    Three things in a chart can name a face and a size independently -- the chart's own
    ``c:txPr``, an axis's, and the legend's -- and they routinely disagree.  Carrying the
    pair together is what stops a label being *measured* in one face and *drawn* in
    another, which is the mistake ``text/metrics.py`` exists to prevent.
    """

    family: str | None
    box: "FontBox"
    #: The face East Asian characters in this text resolve to, which is almost never the
    #: Latin one: ``real-financial-report.pptx``'s axes name ``<a:latin typeface="Arial"/>``
    #: and nothing else, and PowerPoint drew their Japanese labels in the theme's
    #: ``<a:font script="Jpan" typeface="游ゴシック"/>``.  See
    #: :func:`pptx2svg.text.fontmap.east_asian_family` for the cascade and the export it
    #: was read out of.
    family_ea: str | None = None
    #: That face's vertical metrics, for the reserves a CJK label's line box drives.
    box_ea: "FontBox | None" = None

    @property
    def size(self) -> float:
        return self.box.size

    def width(self, text: str) -> float:
        return text_width(text, self.family, self.box.size, self.family_ea)

    def box_for(self, *texts: str) -> "FontBox":
        """The line box of the face that will actually draw ``texts``.

        A chart reserves space for a *known* string, so it can ask which face that string
        resolves to instead of assuming the Latin one.  It matters: on
        ``real-financial-report.pptx``'s radar, Arial's line box is 1.117 em against the
        13.4 pt the fitted reserve wants it to be, and the Japanese face's is 1.448 em.

        Mixed text takes the East Asian box, because PowerPoint's line box is the tallest
        face on the line and the Japanese faces are the taller of the two in every pairing
        here.  Latin-only text takes the Latin box unchanged, which is what keeps every
        measured Latin chart in this file where it was.
        """
        if self.box_ea is not None and any(
            is_cjk(ord(char)) for text in texts for char in text
        ):
            return self.box_ea
        return self.box


@dataclass(frozen=True)
class FontBox:
    """The vertical metrics of one face at one size, in points."""

    size: float
    ascent: float
    descent: float

    @property
    def line_height(self) -> float:
        return self.ascent + self.descent

    @property
    def ink_centre(self) -> float:
        """How far above its baseline a label's optical centre sits.

        Used wherever PowerPoint centres a one-line label on something: a value-axis tick,
        or a legend swatch.  **The rule behind it was not identified.**  Measured offsets
        are 0.218 em (Aptos 10 pt), 0.2975 em (Aptos 8 pt), 0.2687 em (Aptos 14 pt),
        0.3208 em (Arial 12 pt) and 0.213 em (a 10 pt Aptos legend swatch), and that set
        is consistent with *none* of the obvious candidates -- half the cap height, half
        the x-height, half the line box, or the centre of the digits' own ink bounding
        box, each of which is out by 0.4 to 1.2 pt and in inconsistent directions.  So
        this is the fitted mean of the five, whose worst residual is 0.61 pt (about one
        pixel at the 1280 px the fidelity harness scores at).  Two of the five are 8 and
        10 pt, where PowerPoint's own 0.12 pt coordinate quantisation is +/-0.11 pt, so
        part of the spread is measurement noise rather than a missing term.
        """
        return LABEL_INK_CENTRE_EM * self.size

    @property
    def first_baseline(self) -> float:
        """Where our own text engine puts the first baseline below a box's top.

        Chart text is positioned by *baseline* here, but drawn by the ordinary text
        renderer, which positions by box.  Subtracting this converts one to the other, so
        the two stay in step even if the line-box rule changes.
        """
        return (DEFAULT_LINE_HEIGHT_RATIO - self.descent / self.size) * self.size


#: Fallback vertical metrics, as fractions of the em, for a face with no metrics table.
#: The descent is Calibri's, the commonest chart face after the theme's own; the ascent is
#: then whatever makes the line box the 1.2 em the renderer actually lays out, because a
#: shorter one would make an unmetricked title's band come out short of the text in it.
FALLBACK_DESCENT = 0.25
FALLBACK_ASCENT = DEFAULT_LINE_HEIGHT_RATIO - FALLBACK_DESCENT


def font_box(family: str | None, size: float) -> FontBox:
    metrics = metrics_for(family)
    if metrics is None:
        return FontBox(
            size=size, ascent=FALLBACK_ASCENT * size, descent=FALLBACK_DESCENT * size
        )
    units = metrics.units_per_em
    return FontBox(
        size=size,
        ascent=metrics.ascender / units * size,
        descent=abs(metrics.descender) / units * size,
    )


def wrap_label(text: str, font: ChartFont, band: float) -> list[str]:
    """One category label, broken to fit its band the way PowerPoint breaks it.

    Greedy, at whitespace only, and a token that will not fit alone is left to overflow
    rather than split -- which is the state :meth:`ChartBuilder._labels_rotate` has
    already ruled out by turning the whole axis.  See :data:`ROTATED_LABEL_RATIO` for the
    probes behind the break set: a hyphen, a slash, a comma, an underscore, an en dash
    and CJK are all *not* break opportunities, and U+00A0 is.
    """
    tokens = text.split()
    if not tokens or band <= 0:
        return [text] if text else []
    lines: list[str] = []
    current = tokens[0]
    for token in tokens[1:]:
        candidate = f"{current} {token}"
        if font.width(candidate) <= band:
            current = candidate
        else:
            lines.append(current)
            current = token
    lines.append(current)
    return lines[:WRAPPED_LABEL_MAX_LINES]


def text_width(
    text: str, family: str | None, size: float, family_ea: str | None = None
) -> float:
    """One line's advance width, in points.

    The CJK branch is not decoration: ``real-financial-report.pptx`` legends its series
    in Japanese, and measuring those with the Latin mean advance under-counted the legend
    band by 32 pt -- a quarter of the chart's width.  The rule is the same one
    :mod:`pptx2svg.text.measure` uses, so chart text is measured exactly as slide text is.

    ``family_ea`` is which face those East Asian characters resolve to, and it is
    per-character rather than per-string because one label really does mix the two:
    PowerPoint drew ``DX投資額`` with ``DX`` in ArialMT and ``投資額`` in
    YuGothic-Regular, both inside one label.  Measuring the whole string through either
    face alone gets the other half wrong -- and for a proportional Japanese face the
    error is large, ``ＭＳ Ｐゴシック`` running from 0.648 em to 1.0 across its katakana.
    """
    metrics = metrics_for(family)
    ea_metrics = metrics_for(family_ea) if family_ea else None
    if metrics is None and ea_metrics is None:
        return 0.5 * size * len(text)
    total = 0.0
    for char in text:
        east_asian = is_cjk(ord(char))
        table = ea_metrics if east_asian and ea_metrics is not None else metrics
        if table is None:
            total += 0.5 * size
            continue
        width = table.widths.get(char)
        if width is None:
            width = table.cjk_width if east_asian else table.default_width
        total += width / table.units_per_em * size
    return total


# --------------------------------------------------------------------------------------
# The resolved chart
# --------------------------------------------------------------------------------------


@dataclass
class ChartStyle:
    """Everything the drawing needs that is not geometry."""

    font_family: str | None
    font_size: float
    color: m.ResolvedColor
    accents: list[m.ResolvedColor]
    #: The theme's East Asian face -- ``<a:ea>`` if it names one, else the
    #: ``<a:font script="Jpan"/>`` beside it.  A chart's ``c:txPr`` almost never names an
    #: ``<a:ea>`` of its own, so this is what its Japanese text is drawn in.
    font_family_ea: str | None = None


@dataclass
class _Rect:
    left: float
    top: float
    right: float
    bottom: float

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top


@dataclass
class _Labels:
    """``c:dLbls`` resolved down to what actually gets printed."""

    show_value: bool = False
    show_category: bool = False
    show_series: bool = False
    show_percent: bool = False
    position: str | None = None
    number_format: str | None = None
    font: "ChartFont | None" = None
    color: m.ResolvedColor | None = None

    @property
    def anything(self) -> bool:
        return (
            self.show_value or self.show_category or self.show_series or self.show_percent
        )


@dataclass
class _Series:
    name: str | None
    values: list[float | None]
    color: m.ResolvedColor
    fill: m.Fill | None
    outline: m.Outline | None
    format_code: str | None
    invert_if_negative: bool
    #: Per-point ``c:dPt`` overrides.  Presence here also means "this point states its
    #: own formatting", which is what stops `c:invertIfNegative` overriding it.
    point_fills: dict[int, m.Fill] = field(default_factory=dict)
    point_outlines: dict[int, m.Outline] = field(default_factory=dict)
    #: ``c:varyColors`` fills, one per point.  Not a `c:dPt`, so inversion still applies.
    vary_fills: list[m.Fill] = field(default_factory=list)
    #: Line-chart only: the stroke along the points, and the marker drawn at each.
    line: m.Outline | None = None
    marker_symbol: str | None = None
    marker_size: float = DEFAULT_MARKER_SIZE_PT
    marker_fill: m.Fill | None = None
    marker_outline: m.Outline | None = None
    smooth: bool = False
    labels: _Labels | None = None
    #: ``c:dLbl`` overrides, keyed by point index.
    point_labels: dict[int, _Labels] = field(default_factory=dict)


class ChartBuilder:
    """Lowers one bar chart to slide elements, in the frame's own coordinate space.

    Coordinates are EMU with the frame's top-left at the origin, which is what
    :class:`~pptx2svg.model.GroupElement`'s child transform expects.  Internally
    everything is points, because every measured constant is.
    """

    def __init__(
        self,
        chart: c.SourceChart,
        plot: c.SourceChartPlot,
        *,
        width_pt: float,
        height_pt: float,
        style: ChartStyle,
        resolve_fill,
        resolve_outline,
        resolve_text,
        resolve_typeface=lambda typeface: typeface,
    ) -> None:
        self.chart = chart
        self.plot = plot
        self.frame = _Rect(0.0, 0.0, width_pt, height_pt)
        self.style = style
        self._resolve_fill = resolve_fill
        self._resolve_outline = resolve_outline
        self._resolve_text = resolve_text
        self._resolve_typeface = resolve_typeface
        self.elements: list[m.SlideElement] = []
        self._title_cache: "tuple[m.TextBody, FontBox] | None | object" = _UNSET

    # -- public -------------------------------------------------------------------------

    @property
    def _is_polar(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in POLAR_CHART_KINDS

    @property
    def _is_radar(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in RADAR_CHART_KINDS

    @property
    def _is_area(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in AREA_CHART_KINDS

    @property
    def _is_scatter(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in SCATTER_CHART_KINDS

    @property
    def _radar_style(self) -> str:
        """``c:radarStyle``, normalised to what PowerPoint actually draws.

        Measured: ``standard`` and ``marker`` produced *identical* output -- the same
        1.5 pt line, the same diamond at every point -- although ECMA-376 says a
        ``standard`` radar has no markers.  So only ``filled`` is a separate case.
        """
        style = (self.plot.radar_style or "standard").strip()
        return "filled" if style == "filled" else "marker"

    def build(self) -> tuple[list[m.SlideElement], m.ChartData]:
        if self._is_radar:
            return self._build_radar()
        if self._is_polar:
            return self._build_polar()
        if self._is_scatter:
            return self._build_scatter()
        return self._build_cartesian()

    def _build_polar(self) -> tuple[list[m.SlideElement], m.ChartData]:
        """A pie or doughnut: no axes, so none of the Cartesian layout applies."""
        series = self._series()
        categories = self._categories(series)
        region = self._polar_region()
        self._draw_background(region)
        self._draw_title()
        self._draw_pie(region, series, categories)
        self._draw_legend(region, series, per_point=True, categories=categories)
        return self.elements, m.ChartData(
            kind=c.flat_chart_kind(self.plot.kind),
            series=[
                m.ChartSeries(
                    name=item.name,
                    values=list(item.values),
                    categories=list(categories),
                    color=item.color,
                    format_code=item.format_code,
                )
                for item in series
            ],
            categories=list(categories),
            title=self._title_text(),
            grouping=self.plot.grouping,
            bar_direction=None,
            value_axis=None,
            legend_position=self._legend_position(),
        )

    def _build_radar(self) -> tuple[list[m.SlideElement], m.ChartData]:
        """A radar: a value axis wrapped round a ring of category spokes.

        Everything below is measured out of PowerPoint's own PDF -- eighteen probe charts
        across three decks plus ``real-financial-report.pptx``'s own radar, the only real
        one in the corpus.  The polar conventions turn out to be the pie's, and that is a
        measurement rather than an assumption:

        * **angle zero is twelve o'clock and categories run clockwise**, 360/n apart;
        * the centre is the centre of the **same plot region a pie computes** -- edge
          insets plus the legend band -- which the legend probe pins to within 0.2 pt;
        * a point sits at ``(value - minimum) / (maximum - minimum)`` of the radius, so
          the axis minimum is the centre and the outermost ring is the maximum.

        Four things the schema does not say, each from a probe that contradicts the
        obvious reading:

        * **The web is polygonal, it follows the category count, and it is drawn whether
          or not the file asks for it.**  Three, five, six and eight categories gave
          triangles, pentagons, hexagons and octagons; a probe with no ``c:majorGridlines``
          at all still drew every ring, and so did one with ``<c:delete val="1"/>`` on the
          value axis.  Only the *styling* comes from ``c:majorGridlines``.
        * **The spokes do not.**  No probe without a ``c:spPr`` on its category axis drew
          any; the corpus radar, whose category axis states ``<a:ln w="12700">`` in
          #888888, drew six in exactly that.  So the radial lines are the category axis'
          own line, and its default is none -- the opposite of a bar chart, whose default
          axis line is black at 0.5 pt.
        * **``standard`` and ``marker`` draw the same picture**, markers included, though
          ECMA-376 says a ``standard`` radar has none.
        * **``filled`` draws only the fill.**  No markers, and no outline unless the
          series states an ``a:ln`` of its own -- the probe, which states none, emits a
          bare ``f``; the corpus radar, which states ``w="25400"``, is stroked at 2 pt.
        """
        series = self._series()
        categories = self._categories(series)
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")
        scale = self._scale(series, value_axis)
        region = self._polar_region()
        category_font = self._label_font(category_axis)
        value_font = self._label_font(value_axis)

        labels = self._radar_category_labels(
            categories, category_axis, category_font, region
        )
        centre, radius = self._radar_geometry(
            region, labels, category_font, len(categories)
        )

        self._draw_background(region)
        self._draw_title()
        self._draw_radar_web(centre, radius, scale, len(categories), value_axis, category_axis)
        self._draw_radar_series(centre, radius, series, categories, scale)
        if _labels_shown(value_axis):
            self._draw_radar_value_labels(centre, radius, scale, value_axis, value_font)
        self._draw_radar_category_labels(centre, radius, labels, category_font)
        self._draw_radar_data_labels(centre, radius, series, categories, scale)
        self._draw_legend(region, series)

        return self.elements, m.ChartData(
            kind=c.flat_chart_kind(self.plot.kind),
            series=[
                m.ChartSeries(
                    name=item.name,
                    values=list(item.values),
                    categories=list(categories),
                    color=item.color,
                    format_code=item.format_code,
                )
                for item in series
            ],
            categories=list(categories),
            title=self._title_text(),
            grouping=self.plot.grouping,
            bar_direction=None,
            value_axis=m.ChartAxisScale(
                minimum=scale[0], maximum=scale[1], major_unit=scale[2]
            ),
            legend_position=self._legend_position(),
        )

    # -- radar layout -------------------------------------------------------------------

    def _radar_direction(self, index: int, count: int) -> tuple[float, float]:
        """The unit vector down the ``index``-th spoke, in frame coordinates.

        Twelve o'clock, then clockwise -- the pie's convention, measured again here on the
        four-category probe, whose vertices land due north, east, south and west.
        """
        angle = 2.0 * math.pi * index / max(count, 1)
        return math.sin(angle), -math.cos(angle)

    def _radar_category_labels(
        self,
        categories: list[str],
        axis: c.SourceChartAxis | None,
        font: ChartFont,
        region: _Rect,
    ) -> list[list[str]]:
        """Each category's label, already broken into the lines it will be drawn on.

        PowerPoint **wraps** a label too wide for its corner rather than rotating it: the
        long-label probe came back with every ``rot`` zero and "Category Three" split over
        two lines as "Category" / "Three".  The width it wraps at is bracketed by
        :data:`RADAR_LABEL_MAX_FRACTION`.
        """
        if not categories or not _labels_shown(axis):
            return []
        cap = region.width * RADAR_LABEL_MAX_FRACTION
        return [_wrap_to_width(name, font, cap) for name in categories]

    def _radar_geometry(
        self, region: _Rect, labels: list[list[str]], font: ChartFont, count: int
    ) -> tuple[tuple[float, float], float]:
        """The web's centre and radius.

        Two constraints, both measured, and the smaller wins:

        * **vertically**, the radius falls short of half the region by a reserve that is a
          function of the label's line box -- see :data:`RADAR_LABEL_RESERVE_LINES`;
        * **horizontally**, a category's label must still fit the region, so the vertex it
          hangs off can be no further out than ``half_width - gap - label_width``, which
          for a spoke at angle theta bounds the radius by that over ``|sin theta|``.

        Three probes are horizontally bound and land within 0.7 pt; five are vertically
        bound and land within 0.09 pt.

        **The reserve counts one line even when a label wraps**, and that is the
        measurement rather than an oversight: the one multi-line observation has a radius
        of 60.24 pt, which the horizontal constraint reproduces exactly and whose vertical
        reserve is therefore at most 19.31 pt -- less than the 20.16 pt that two lines of
        the fitted per-line reserve would ask for.  A rule scaling the reserve by the line
        count is contradicted by that probe, so it is not used.
        """
        centre = ((region.left + region.right) / 2, (region.top + region.bottom) / 2)
        half_width, half_height = region.width / 2, region.height / 2
        radius = min(half_width, half_height)
        if labels:
            reserve = max(
                RADAR_LABEL_RESERVE_LINES
                * font.box_for(*(line for lines in labels for line in lines)).line_height
                - RADAR_LABEL_RESERVE_PT,
                0.0,
            )
            radius = min(radius, half_height - reserve)
            for index, lines in enumerate(labels):
                sideways = abs(self._radar_direction(index, count)[0])
                if sideways < 1e-3:
                    continue
                width = max((font.width(line) for line in lines), default=0.0)
                room = half_width - RADAR_LABEL_GAP_PT - width
                radius = min(radius, room / sideways)
        return centre, max(radius, 1.0)

    def _radar_radius(
        self, radius: float, value: float, scale: tuple[float, float, float]
    ) -> float:
        minimum, maximum, _ = scale
        span = maximum - minimum
        if span <= 0:
            return 0.0
        return (value - minimum) / span * radius

    def _radar_point(
        self,
        centre: tuple[float, float],
        index: int,
        count: int,
        distance: float,
    ) -> tuple[float, float]:
        dx, dy = self._radar_direction(index, count)
        return centre[0] + dx * distance, centre[1] + dy * distance

    # -- radar drawing ------------------------------------------------------------------

    def _draw_radar_web(
        self,
        centre: tuple[float, float],
        radius: float,
        scale: tuple[float, float, float],
        count: int,
        value_axis: c.SourceChartAxis | None,
        category_axis: c.SourceChartAxis | None,
    ) -> None:
        if count <= 0:
            return
        ring_outline = self._axis_outline(
            value_axis.major_gridline_outline if value_axis is not None else None
        )
        for value in self._tick_values(scale)[1:]:
            distance = self._radar_radius(radius, value, scale)
            if distance <= 0:
                continue
            points = [
                self._radar_point(centre, index, count, distance)
                for index in range(count)
            ]
            if len(points) > 1:
                self._polyline(points + [points[0]], ring_outline, smooth=False)

        # The category axis draws the spokes, and only when it states a line of its own.
        spoke = (
            self._resolve_outline(category_axis.outline)
            if category_axis is not None and category_axis.outline is not None
            else None
        )
        if spoke is None or spoke.fill is None or isinstance(spoke.fill, m.NoFill):
            return
        for index in range(count):
            x, y = self._radar_point(centre, index, count, radius)
            self._line(centre[0], centre[1], x, y, spoke)

    def _draw_radar_series(
        self,
        centre: tuple[float, float],
        radius: float,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        count = len(categories)
        if count <= 0:
            return
        filled = self._radar_style == "filled"
        # `dispBlanksAs` on a radar is **not measured**; this mirrors the line chart, whose
        # behaviour was.  A ring with a hole in it cannot close, so it is drawn open.
        blanks = self.chart.display_blanks_as or "gap"
        for item in series:
            points: list[tuple[float, float] | None] = []
            for index in range(count):
                value = _at(item.values, index)
                if value is None:
                    if blanks == "zero":
                        value = scale[0]
                    elif blanks == "span":
                        continue
                    else:
                        points.append(None)
                        continue
                points.append(
                    self._radar_point(
                        centre,
                        index,
                        count,
                        self._radar_radius(radius, value, scale),
                    )
                )
            drawn = [point for point in points if point is not None]
            closed = len(drawn) == count and count > 2
            if filled:
                if len(drawn) > 2:
                    self._polygon(drawn, fill=item.fill, outline=item.outline)
                continue
            if item.line is not None:
                if closed:
                    self._polyline(drawn + [drawn[0]], item.line, smooth=False)
                else:
                    # A ring is a cycle, so the run either side of index 0 is one run.
                    # Rotating the list to start just after a blank is what makes
                    # `_split_runs`, which walks a straight line, see it that way.
                    for run in _split_runs(_rotate_past_blank(points)):
                        if len(run) > 1:
                            self._polyline(run, item.line, smooth=False)
            for point in points:
                if point is not None and item.marker_symbol:
                    self._marker(point, item)

    def _draw_radar_value_labels(
        self,
        centre: tuple[float, float],
        radius: float,
        scale: tuple[float, float, float],
        axis: c.SourceChartAxis | None,
        font: ChartFont,
    ) -> None:
        """The tick labels, up the twelve o'clock spoke and to the left of it.

        Measured on all six charts that draw one: right-aligned, their right edge two
        widths of the digit zero left of the spoke, each centred on its own ring.
        """
        box = font.box
        right = centre[0] - RADAR_VALUE_LABEL_DIGITS * font.width("0")
        for value, text in self._tick_texts(scale, axis):
            if not text:
                continue
            y = centre[1] - self._radar_radius(radius, value, scale)
            width = font.width(text) + box.size
            self._text(
                self._label_body(text, font, align="r"),
                left=right - width,
                width=width,
                baseline=y + box.ink_centre,
                box=box,
            )

    def _draw_radar_category_labels(
        self,
        centre: tuple[float, float],
        radius: float,
        labels: list[list[str]],
        font: ChartFont,
    ) -> None:
        """One label per spoke, its box pushed radially clear of the vertex.

        The four horizontal directions land within 0.15 pt; the two vertical ones are up
        to 1.6 pt out, which the measurements attribute to PowerPoint's line box being
        about a point taller than the one our font metrics give rather than to a second
        rule -- see :data:`RADAR_LABEL_GAP_PT`.
        """
        count = len(labels)
        box = font.box
        for index, lines in enumerate(labels):
            if not any(lines):
                continue
            dx, dy = self._radar_direction(index, count)
            anchor_x = centre[0] + dx * (radius + RADAR_LABEL_GAP_PT)
            anchor_y = centre[1] + dy * (radius + RADAR_LABEL_GAP_PT)
            width = max(font.width(line) for line in lines)
            block = box.line_height * len(lines)
            if dx > 1e-3:
                left = anchor_x
            elif dx < -1e-3:
                left = anchor_x - width
            else:
                left = anchor_x - width / 2
            if dy > 1e-3:
                top = anchor_y
            elif dy < -1e-3:
                top = anchor_y - block
            else:
                top = anchor_y - block / 2
            # The box is padded half an em either side so a wide glyph is not clipped,
            # and shifted back by the same amount so each *line* stays centred on the
            # block the anchoring above placed.
            for line_index, line in enumerate(lines):
                self._text(
                    self._label_body(line, font, align="ctr"),
                    left=left - box.size / 2,
                    width=width + box.size,
                    baseline=top + line_index * box.line_height + box.ascent,
                    box=box,
                )

    def _draw_radar_data_labels(
        self,
        centre: tuple[float, float],
        radius: float,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """Data labels, pushed radially out from their point.

        **Not measured.**  The corpus radar carries a ``c:dLbls`` block with every
        ``c:show*`` flag at 0, and no probe turned one on, so where PowerPoint puts a
        radar's data label is unknown; this places it the way a pie's ``outEnd`` sits, one
        marker clear of the point along its own spoke.  It is the one thing the radar path
        draws without a measurement behind it, and it is here rather than absent because
        silently dropping flags a file states is the worse failure.
        """
        count = len(categories)
        for item in series:
            for index in range(count):
                value = _at(item.values, index)
                if value is None:
                    continue
                labels = item.point_labels.get(index, item.labels)
                if labels is None or not labels.anything or labels.font is None:
                    continue
                parts: list[str] = []
                if labels.show_series and item.name:
                    parts.append(item.name)
                if labels.show_category and categories[index]:
                    parts.append(categories[index])
                if labels.show_value:
                    parts.append(
                        format_number(value, labels.number_format or item.format_code)
                    )
                if not parts:
                    continue
                distance = self._radar_radius(radius, value, scale) + item.marker_size
                x, y = self._radar_point(centre, index, count, distance)
                self._centred_label(parts, labels.font, x, y)

    def _build_cartesian(self) -> tuple[list[m.SlideElement], m.ChartData]:
        series = self._series()
        categories = self._categories(series)
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")

        scale = self._scale(series, value_axis)
        # Each axis carries its own `c:txPr`, and they disagree in real files.
        value_font = self._label_font(value_axis)
        category_font = self._label_font(category_axis)
        tick_texts = self._tick_texts(scale, value_axis)

        plot_rect = self._plot_rect(
            tick_texts, categories, value_font, category_font, scale
        )

        self._draw_background(plot_rect)
        self._draw_title()
        self._draw_gridlines(plot_rect, scale, value_axis)
        if self._is_area:
            self._draw_areas(plot_rect, series, categories, scale)
        elif self._is_line:
            self._draw_lines(plot_rect, series, categories, scale)
        else:
            self._draw_bars(plot_rect, series, categories, scale)
        self._draw_axis_lines(plot_rect, scale, value_axis, category_axis)
        self._draw_labels(
            plot_rect,
            scale,
            tick_texts,
            categories,
            value_axis,
            category_axis,
            value_font,
            category_font,
        )
        self._draw_data_labels(plot_rect, series, categories, scale)
        self._draw_legend(plot_rect, series)

        data = m.ChartData(
            kind=c.flat_chart_kind(self.plot.kind),
            series=[
                m.ChartSeries(
                    name=item.name,
                    values=list(item.values),
                    categories=list(categories),
                    color=item.color,
                    format_code=item.format_code,
                )
                for item in series
            ],
            categories=list(categories),
            title=self._title_text(),
            grouping=self.plot.grouping,
            bar_direction=self.plot.bar_direction or "col",
            value_axis=m.ChartAxisScale(
                minimum=scale[0], maximum=scale[1], major_unit=scale[2]
            ),
            legend_position=self._legend_position(),
        )
        return self.elements, data

    def _build_scatter(self) -> tuple[list[m.SlideElement], m.ChartData]:
        """A scatter: **both axes are value axes and there is no category axis at all.**

        That is the one structural difference from every other Cartesian chart here, and it
        is why this does not go through :meth:`_build_cartesian`.  Almost everything the
        bottom of a chart does is written for a *category* axis -- the band a label is laid
        into, the wrap that fills it, the 45 degree turn when it will not, the reserve
        those together ask for -- and none of it applies to a row of numbers.  What a
        scatter's bottom band actually is, measured, is the *horizontal bar chart's*: one
        plain line of value labels centred on their ticks, with half of the last one
        hanging past the plot's right edge.

        Every constant below is shared with a chart that already draws.  The x axis is the
        same coarse one a horizontal bar chart gets -- 1..5 of data came back 0..6 **by
        two** where the y axis over the same span takes ones -- which is
        :data:`HORIZONTAL_MAX_INTERVALS`, measured here for the second time.
        """
        series = self._series()
        x_axis, y_axis = self._scatter_axes()

        x_values = [self._x_values(index, item) for index, item in enumerate(series)]
        xs = [value for column in x_values for value in column if value is not None]
        ys = [value for item in series for value in item.values if value is not None]
        # **Neither axis is anchored at zero**, which every other type here is.  A bar has
        # to start at its axis; a scatter of years against a measurement would be destroyed
        # by it, and PowerPoint agrees -- see :data:`AXIS_ZERO_ANCHOR_RATIO`.
        x_scale = _apply_axis_limits(
            nice_axis_scale(*_span(xs), horizontal=True, anchor_zero=False), x_axis
        )
        y_scale = _apply_axis_limits(
            nice_axis_scale(*_span(ys), horizontal=False, anchor_zero=False), y_axis
        )

        x_font = self._label_font(x_axis)
        y_font = self._label_font(y_axis)
        x_ticks = self._tick_texts(x_scale, x_axis)
        y_ticks = self._tick_texts(y_scale, y_axis)

        rect = self._scatter_plot_rect(x_ticks, y_ticks, x_font, y_font, x_scale, y_scale)
        # Where each axis is drawn: at the *other* axis' zero, clamped into the plot.
        # Measured on the negative-x probe, whose value axis is drawn at 111.088 pt -- the
        # x = 0 tick -- and not at the plot's left edge 95.7 pt away.
        #
        # `c:crosses` belongs to the axis it is written on and says where **that** axis
        # crosses the perpendicular one, which is how the sibling
        # :meth:`_category_axis_position` reads it.  So the vertical line's position is a
        # question for the *y* axis even though the answer is an x coordinate.  Only
        # `autoZero` is measured; the rest follow the category axis' reading rather than a
        # second one invented here.
        cross_x = self._crossing(
            rect.left, rect.right, y_axis, x_scale, self._value_to_x, rect
        )
        cross_y = self._crossing(
            rect.bottom, rect.top, x_axis, y_scale, self._value_to_y, rect
        )

        self._draw_background(rect)
        self._draw_title()
        if y_axis is not None and y_axis.major_gridlines:
            outline = self._axis_outline(y_axis.major_gridline_outline)
            for value in self._tick_values(y_scale):
                y = self._value_to_y(rect, value, y_scale)
                if abs(y - cross_y) >= 0.01:
                    self._line(rect.left, y, rect.right, y, outline)
        if x_axis is not None and x_axis.major_gridlines:
            outline = self._axis_outline(x_axis.major_gridline_outline)
            for value in self._tick_values(x_scale):
                x = self._value_to_x(rect, value, x_scale)
                if abs(x - cross_x) >= 0.01:
                    self._line(x, rect.top, x, rect.bottom, outline)

        for index, item in enumerate(series):
            self._draw_scatter_series(rect, item, x_values[index], x_scale, y_scale)

        if y_axis is not None and not y_axis.delete:
            outline = self._axis_outline(y_axis.outline)
            self._line(cross_x, rect.top, cross_x, rect.bottom, outline)
        if x_axis is not None and not x_axis.delete:
            outline = self._axis_outline(x_axis.outline)
            self._line(rect.left, cross_y, rect.right, cross_y, outline)

        if _labels_shown(y_axis):
            placed = [
                (self._value_to_y(rect, value, y_scale), text) for value, text in y_ticks
            ]
            self._labels_down_left(rect, placed, y_font, axis_x=cross_x)
        if _labels_shown(x_axis):
            placed = [
                (self._value_to_x(rect, value, x_scale), text) for value, text in x_ticks
            ]
            self._labels_along_bottom(
                rect, placed, x_font, axis_y=cross_y, centred_on_position=True
            )

        self._draw_scatter_labels(rect, series, x_values, x_scale, y_scale)
        self._draw_legend(rect, series)

        return self.elements, m.ChartData(
            kind=c.flat_chart_kind(self.plot.kind),
            series=[
                m.ChartSeries(
                    name=item.name,
                    values=list(item.values),
                    # A scatter has no categories; the x values stand in for them, which is
                    # also what the reader caches into `c:cat` when there is no `c:cat`.
                    categories=[
                        "" if value is None else format_number(value, None)
                        for value in x_values[index]
                    ],
                    color=item.color,
                    format_code=item.format_code,
                )
                for index, item in enumerate(series)
            ],
            categories=[],
            title=self._title_text(),
            grouping=self.plot.grouping,
            bar_direction=None,
            value_axis=m.ChartAxisScale(
                minimum=y_scale[0], maximum=y_scale[1], major_unit=y_scale[2]
            ),
            legend_position=self._legend_position(),
        )

    def _scatter_axes(self) -> "tuple[c.SourceChartAxis | None, c.SourceChartAxis | None]":
        """``(x, y)``.  Both are ``c:valAx``, so they are told apart by position.

        The group's own ``c:axId`` list is authoritative and lists x first; ``c:axPos`` is
        the fallback for a file that names ids no axis declares.
        """
        x_axis, y_axis = self._axis_for(0), self._axis_for(1)
        if x_axis is None or y_axis is None:
            by_position = {axis.position: axis for axis in self.chart.axes}
            x_axis = x_axis or by_position.get("b") or by_position.get("t")
            y_axis = y_axis or by_position.get("l") or by_position.get("r")
        return x_axis, y_axis

    def _x_values(self, index: int, item: _Series) -> list[float | None]:
        """One series' ``c:xVal``, padded to the length of its ``c:yVal``.

        ``c:xVal`` is optional and a short one is legal, so the tail is filled with the
        1-based position -- which is what a spreadsheet's implicit x column would hold.
        **Not measured**: every probe states a full ``c:xVal``, and no corpus deck has a
        scatter at all.
        """
        source = self.plot.series[index] if index < len(self.plot.series) else None
        xs = list(source.x_values) if source is not None and source.x_values else []
        if len(xs) < len(item.values):
            xs += [float(position + 1) for position in range(len(xs), len(item.values))]
        return xs[: len(item.values)]

    def _crossing(
        self,
        low: float,
        high: float,
        axis: "c.SourceChartAxis | None",
        scale: tuple[float, float, float],
        to_position,
        rect: _Rect,
    ) -> float:
        """Where ``axis`` is drawn along the perpendicular axis, in frame points.

        ``low`` and ``high`` are that perpendicular axis' two ends; ``scale`` and
        ``to_position`` are *its* domain and mapping, because `c:crossesAt` is a value on
        the axis being crossed.  The reading is :meth:`_category_axis_position`'s, kept in
        step with it deliberately: a scatter's two value axes are the category axis and the
        value axis of a bar chart with the category names taken away.
        """
        if axis is not None and axis.tick_label_position == "low":
            return low
        crosses = axis.crosses if axis is not None else None
        if crosses == "max":
            return high
        if crosses == "min":
            return low
        value = axis.crosses_at if axis is not None and crosses == "val" else 0.0
        if value is None:
            value = 0.0
        along = to_position(rect, value, scale)
        return min(max(along, min(low, high)), max(low, high))

    def _scatter_plot_rect(
        self,
        x_ticks: list[tuple[float, str]],
        y_ticks: list[tuple[float, str]],
        x_font: ChartFont,
        y_font: ChartFont,
        x_scale: tuple[float, float, float],
        y_scale: tuple[float, float, float],
    ) -> _Rect:
        """The plot rectangle, from the same four measurements a bar chart's comes from.

        The left column, the top allowance, the title band and the legend band are the bar
        chart's own formulae; what is new is that **the bottom holds value labels rather
        than category ones**, so the band is one plain line and the last label overhangs
        the right edge by half its width.  Six probes: ``bare`` 21.073 / 13.670 / 11.102 /
        24.965, ``font14`` 26.907 / 14.740 / 13.545 / 32.353, ``title`` top 40.802, and
        ``x-deleted`` -- no x labels at all -- 209.472 and 11.102, all within 0.02 pt.

        Two corners mirror what a bar chart already does with negative values.  When the x
        range goes below zero the value axis floats into the plot and the y labels go with
        it, so the left column reserves nothing and the inset is the first x label's
        overhang instead: 15.373 pt measured against 15.373 predicted.  When the *y* range
        does, the x labels move up beside the zero line and the bottom band disappears.
        """
        frame = self.frame
        x_axis, y_axis = self._scatter_axes()
        show_x = _labels_shown(x_axis)
        show_y = _labels_shown(y_axis)
        x_labels = [text for _, text in x_ticks]
        y_labels = [text for _, text in y_ticks]

        left = frame.left + EDGE_INSET_PT
        if show_y and x_scale[0] >= 0:
            widest = max((y_font.width(text) for text in y_labels), default=0.0)
            left = (
                frame.left
                + FRAME_PADDING_PT
                + widest
                + y_font.box.descent
                + VALUE_LABEL_GAP_EM * y_font.size
            )
        overhang = 0.0
        if show_x and x_labels:
            left = max(left, frame.left + EDGE_INSET_PT + x_font.width(x_labels[0]) / 2)
            overhang = x_font.width(x_labels[-1]) / 2

        right = frame.right - EDGE_INSET_PT - overhang
        top = frame.top + self._top_inset(y_font.box)
        title = self._title_box()
        if title is not None:
            top += TITLE_BAND_LINES * title.line_height

        legend_bottom = 0.0
        legend = self._legend_position()
        if legend is not None and not self._legend_overlays():
            legend_font = self._legend_font()
            band = LEGEND_BAND_LINES * legend_font.box.line_height
            if legend == "b":
                legend_bottom = band
            elif legend in ("t", "tr"):
                top += band
            elif legend == "r":
                right = frame.right - self._legend_side_width(legend_font) - overhang
            elif legend == "l":
                left += self._legend_side_width(legend_font) - EDGE_INSET_PT

        if right - left < 1.0:
            right = left + 1.0
        if show_x and y_scale[0] >= 0:
            bottom = frame.bottom - legend_bottom - self._bottom_label_band(
                x_font, [], right - left
            )
        else:
            bottom = frame.bottom - legend_bottom - self._top_inset(y_font.box)
        if bottom - top < 1.0:
            bottom = top + 1.0
        return _Rect(left, top, right, bottom)

    def _scatter_points(
        self,
        rect: _Rect,
        item: _Series,
        xs: list[float | None],
        x_scale: tuple[float, float, float],
        y_scale: tuple[float, float, float],
    ) -> list[tuple[float, float] | None]:
        """The drawn points, in data order, with ``None`` where the run breaks.

        **A blank breaks a scatter exactly as it breaks a line**, and the probe says so in
        a way the vertex list alone does not: the path through a missing middle y has four
        points, which reads as unbroken, but its *segment kinds* are move, line, move,
        line -- two disjoint strokes with a gap where the blank is.  Reading coordinates
        without reading the operators is how that gets missed.

        The points are joined **in the order the file lists them**, not sorted by x: the
        unsorted-x probe's path runs 3, 1, 5, 2, 4.
        """
        blanks = self.chart.display_blanks_as or "gap"
        points: list[tuple[float, float] | None] = []
        for index, value in enumerate(item.values):
            x = xs[index] if index < len(xs) else None
            if value is None or x is None:
                if blanks == "zero" and x is not None:
                    value = 0.0
                elif blanks == "span":
                    # Unmeasured for a scatter; the line chart's reading is reused.
                    continue
                else:
                    points.append(None)
                    continue
            points.append(
                (self._value_to_x(rect, x, x_scale), self._value_to_y(rect, value, y_scale))
            )
        return points

    def _draw_scatter_series(
        self,
        rect: _Rect,
        item: _Series,
        xs: list[float | None],
        x_scale: tuple[float, float, float],
        y_scale: tuple[float, float, float],
    ) -> None:
        points = self._scatter_points(rect, item, xs, x_scale, y_scale)
        for run in _split_runs(points):
            if len(run) > 1 and item.line is not None:
                self._polyline(run, item.line, smooth=item.smooth)
        for point in points:
            if point is not None and item.marker_symbol:
                self._marker(point, item)

    def _draw_scatter_labels(
        self,
        rect: _Rect,
        series: list[_Series],
        x_values: list[list[float | None]],
        x_scale: tuple[float, float, float],
        y_scale: tuple[float, float, float],
    ) -> None:
        """``c:dLbls`` beside each point.

        All five placements ECMA-376 allows a scatter are measured on one probe each, and
        every one of them is the line chart's own geometry read off a different edge of the
        marker: ``r`` and ``l`` put the label's near edge a marker radius plus 0.6 em from
        the point (9.000 pt measured, 9.000 predicted, both sides); ``t`` and ``b`` put its
        line box a radius plus :data:`DATA_LABEL_GAP_PT` away, which lands within 0.15 pt
        above and 0.70 pt below; ``ctr`` is the ink centre, 0.23 pt out.  The 0.70 pt is
        the one loose number and it is the same slack the roadmap records elsewhere --
        PowerPoint's line box runs about a point taller than our metrics give.
        """
        categories = [
            ["" if value is None else format_number(value, None) for value in column]
            for column in x_values
        ]
        totals = _percent_totals(series)
        for order, item in enumerate(series):
            xs = x_values[order]
            for point, value in enumerate(item.values):
                labels = item.point_labels.get(point, item.labels)
                if labels is None or not labels.anything or labels.font is None:
                    continue
                x = xs[point] if point < len(xs) else None
                if value is None or x is None:
                    continue
                # The x values *are* the categories -- `parse/chart` already caches them
                # into `c:cat` when a scatter states none -- so `c:showCatName` prints
                # something rather than nothing.
                text = self._label_text(
                    labels, item, categories[order], point, value, totals
                )
                if not text:
                    continue
                centre = (
                    self._value_to_x(rect, x, x_scale),
                    self._value_to_y(rect, value, y_scale),
                )
                position = (labels.position or self._label_default()).lower()
                radius = item.marker_size / 2
                if position == "ctr":
                    geometry = (centre[0], centre[1], "centre")
                elif position == "l":
                    geometry = (centre[0] - radius, centre[1], "left")
                elif position == "t":
                    geometry = (centre[0], centre[1] - radius, "outside-y")
                elif position == "b":
                    geometry = (centre[0], centre[1] + radius, "below")
                else:
                    geometry = (centre[0] + radius, centre[1], "right")
                self._place_label(text, labels, geometry, False)

    # -- model --------------------------------------------------------------------------

    def _vary_colors(self) -> bool:
        """Whether each *point* takes its own colour rather than the series' one.

        **A pie varies by default and a bar does not.**  Measured: probe pies stating no
        ``c:varyColors`` at all came out accent1, accent2, accent3, accent4 across their
        four slices, and a two-ring doughnut cycled the same four in *both* rings -- so it
        is per point, not per series.  A bar chart with no ``c:varyColors`` draws one
        colour for the whole series, which is why the default cannot simply be true.

        For a bar it is also only meaningful on a single unstacked series; with several,
        the colours already vary by series.
        """
        if self._is_polar:
            return self.plot.vary_colors is not False
        if not self.plot.vary_colors or len(self.plot.series) != 1:
            return False
        return (self.plot.grouping or "clustered") not in ("stacked", "percentStacked")

    def _series(self) -> list[_Series]:
        out: list[_Series] = []
        for source in self.plot.series:
            fill = self._resolve_fill(source.fill)
            # `c:idx`, not the position in this list.  The list is sorted by `c:order`
            # and every probe so far had idx == order, so the two were indistinguishable
            # until `real-college-template`'s chart, whose two series are
            # (idx 2, order 0) and (idx 0, order 1).  The second states no fill;
            # PowerPoint drew it in that theme's accent1 (#C00000) and not in accent2
            # (#595959), which is what the list position would have given.
            color = self._series_color(fill, source.index)
            if fill is None:
                fill = m.SolidFill(color=color)
            outline = self._resolve_outline(source.outline)
            item = _Series(
                name=source.name.plain if source.name else None,
                values=list(source.values),
                color=color,
                fill=fill,
                outline=outline,
                format_code=source.format_code,
                # Absent means *no* inversion.  This used to default to true, reading
                # ECMA-376's CT_Boolean -- whose `val` attribute defaults to 1 -- as
                # though it also said what an absent element means.  It does not, and
                # PowerPoint disagrees: `real-college-template`'s stacked chart omits
                # `c:invertIfNegative` over a -1.0 and PowerPoint drew that bar solid in
                # the series colour.  Measured across four probes built from that chart
                # (stacked/clustered x automatic accent fill/explicit `srgbClr 2563EB`):
                # the negative bar came out solid in the series colour in all four.  The
                # same chart with `val="1"` added came out white with a dark outline in
                # both fill variants, so the hollow drawing below is right -- only the
                # default was wrong.
                invert_if_negative=bool(source.invert_if_negative),
            )
            if self._vary_colors() and source.fill is None and self.style.accents:
                # Measured on the varyColors probe: points take accent1, accent2, accent3
                # *exactly*.  pptx-renderer darkens them to 88%, which PowerPoint does not.
                # Kept apart from `point_fills`, which means "this point has a `c:dPt` of
                # its own" and is what suppresses the negative-bar inversion.
                item.vary_fills = [
                    m.SolidFill(color=self.style.accents[index % len(self.style.accents)])
                    for index in range(len(item.values))
                ]
            if (
                self._is_line
                or self._is_scatter
                or (self._is_radar and self._radar_style != "filled")
            ):
                # **A scatter series is styled exactly like a line series, and
                # ``c:scatterStyle`` decides nothing.**  Probes at `marker`, `line` and
                # `lineMarker` came back byte-identical -- line *and* markers in all three
                # -- so what turns either off is the series' own markup: `<a:ln><a:noFill/>`
                # for the line, `<c:symbol val="none"/>` for the marker, each measured.
                self._read_line_style(item, source, source.index)
                if self._is_radar and (source.marker is None or not source.marker.size):
                    # Measured on the probe: a radar series stating no `c:size` draws a
                    # 6 pt marker, not ECMA-376's 7.
                    item.marker_size = RADAR_MARKER_SIZE_PT
            item.labels = self._read_labels(source.data_labels, self.plot.data_labels)
            if source.data_labels is not None:
                for point_index, override in source.data_labels.overrides.items():
                    item.point_labels[point_index] = self._read_labels(
                        override, source.data_labels, self.plot.data_labels
                    )
            for point in source.data_points:
                point_fill = self._resolve_fill(point.fill)
                if point_fill is not None:
                    item.point_fills[point.index] = point_fill
                point_outline = self._resolve_outline(point.outline)
                if point_outline is not None:
                    item.point_outlines[point.index] = point_outline
            out.append(item)
        return out

    def _read_labels(self, *sources: c.SourceChartDataLabels | None) -> _Labels:
        """``c:dLbls`` innermost first: a point's, then its series', then the group's.

        Every flag is optional at every level, so each is resolved separately rather than
        taking the first block whole.  ``c:delete`` on a point silences it outright.
        """
        present = [source for source in sources if source is not None]
        if any(source.delete for source in present):
            return _Labels()

        def flag(name: str) -> bool:
            for source in present:
                value = getattr(source, name)
                if value is not None:
                    return value
            return False

        def first(name: str):
            for source in present:
                value = getattr(source, name)
                if value:
                    return value
            return None

        labels = _Labels(
            show_value=flag("show_value"),
            show_category=flag("show_category_name"),
            show_series=flag("show_series_name"),
            show_percent=flag("show_percent"),
            position=first("position"),
            number_format=first("number_format"),
        )
        if labels.anything:
            bodies = [source.text_properties for source in present]
            labels.font = self._font(*bodies)
            labels.color = self._text_color(*bodies) or self.style.color
        return labels

    def _text_color(self, *sources: "s.SourceTextBody | None") -> m.ResolvedColor | None:
        """The innermost ``c:txPr``'s ``a:defRPr/a:solidFill``, resolved.

        Only data labels read this so far, because that is the only place it has been
        measured: ``real-college-template``'s chart sets ``<a:schemeClr val="bg1"/>`` on
        both series' ``c:dLbls`` and PowerPoint inks the numbers white inside the bars.
        Ours came out in the chart-space default, which is dark text on a #C00000 fill.

        The same element governs axis, legend and title text and is *not* read for those;
        nothing in the corpus states one there, so there is nothing to check a change
        against.  :meth:`_font` is the place it would go.
        """
        for source in (*sources, self.chart.text_properties):
            run = _default_run(source)
            if run is not None and run.color is not None:
                fill = self._resolve_fill(s.SourceSolidFill(color=run.color))
                if isinstance(fill, m.SolidFill):
                    return fill.color
        return None

    @property
    def _is_line(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) == "lineChart"

    def _read_line_style(
        self, item: _Series, source: c.SourceChartSeries, index: int
    ) -> None:
        """A line series' stroke and marker.

        **A line's colour comes from its ``a:ln``, not from ``a:solidFill``.**  Measured:
        a probe series stating only ``<a:solidFill><a:srgbClr val="F97316"/></a:solidFill>``
        was drawn by PowerPoint in accent1, its bare fill ignored -- so a bar chart's
        colour rule cannot simply be reused here.

        ``index`` is ``c:idx``, for the same reason the bar accents are -- see
        :meth:`_series`.  The *marker* cycle below rides on the same number because it is
        the same "which series is this" question, but only the accent was measured under
        ``idx != order``; no line chart in the corpus or in any probe has one.
        """
        outline = self._resolve_outline(source.outline)
        if outline is not None and isinstance(outline.fill, m.SolidFill):
            item.color = outline.fill.color
        elif self.style.accents:
            item.color = self.style.accents[index % len(self.style.accents)]
        # **``<a:ln><a:noFill/></a:ln>`` is an explicit *no line*, not "use the default".**
        # `_resolve_outline` collapses it to None exactly as it collapses an absent
        # `c:spPr`, so the source element has to be asked -- the same trap
        # :meth:`_axis_outline` exists for.  Measured: it is how a marker-only scatter is
        # spelled, and the probe that states it drew five markers and no line at all;
        # substituting the default here drew a line PowerPoint does not.
        if isinstance(source.outline, s.SourceOutline) and isinstance(
            source.outline.fill, s.SourceNoFill
        ):
            item.line = None
            self._read_marker_style(item, source, index)
            return
        if outline is None or outline.fill is None:
            outline = m.Outline(
                width=DEFAULT_LINE_SERIES_WIDTH_EMU,
                fill=m.SolidFill(color=item.color),
            )
        if outline.line_cap is None:
            # Measured: PowerPoint strokes a line series with `1 J`, a round cap, whether
            # or not the `a:ln` says so.
            outline = replace(outline, line_cap="round")
        item.line = outline
        self._read_marker_style(item, source, index)
        # **An absent `c:smooth` smooths.**  It is a chart boolean, so the element being
        # missing is not the same as `val="0"` -- exactly the trap `parse/chart._flag`
        # exists for, and this reader used `bool(None) == False`.  Three probes: a line
        # chart with `<c:smooth val="1"/>`, one with the element absent and one with
        # `val="0"` came back as, respectively, four cubics, the *same* four cubics, and a
        # four-segment polyline.  A scatter behaves identically.  PowerPoint's own writer
        # always emits the element, so no corpus deck moves; a hand-written one does.
        item.smooth = source.smooth is not False

    def _read_marker_style(
        self, item: _Series, source: c.SourceChartSeries, index: int
    ) -> None:
        marker = source.marker
        if marker is None:
            # No `c:marker` at all: PowerPoint draws one anyway, from a per-series cycle.
            item.marker_symbol = DEFAULT_MARKER_CYCLE[index % len(DEFAULT_MARKER_CYCLE)]
        elif marker.symbol in (None, "auto"):
            item.marker_symbol = DEFAULT_MARKER_CYCLE[index % len(DEFAULT_MARKER_CYCLE)]
        elif marker.symbol != "none":
            item.marker_symbol = marker.symbol
        if marker is not None and marker.size:
            item.marker_size = marker.size
        item.marker_fill = (
            self._resolve_fill(marker.fill) if marker is not None else None
        ) or m.SolidFill(color=item.color)
        item.marker_outline = (
            self._resolve_outline(marker.outline) if marker is not None else None
        ) or m.Outline(
            width=DEFAULT_MARKER_OUTLINE_EMU, fill=m.SolidFill(color=item.color)
        )

    def _series_color(self, fill: m.Fill | None, index: int) -> m.ResolvedColor:
        """One flat colour for the series, for its legend swatch and its fallback fill.

        A series that states no fill takes the next theme accent, cycling; Office's own
        accent1 is the last resort when the theme has none.
        """
        if isinstance(fill, m.SolidFill):
            return fill.color
        if self.style.accents:
            return self.style.accents[index % len(self.style.accents)]
        return m.ResolvedColor(hex="#4472C4")

    def _categories(self, series: list[_Series]) -> list[str]:
        """One label per category, as long as the longest series.

        The bars are indexed by this list, so a series longer than the labelled one would
        lose its tail.  Series in one group need not agree on length -- and a series may
        carry no ``c:cat`` at all -- so the labelled list is padded rather than trusted for
        its length.  PowerPoint numbers unlabelled categories from 1.
        """
        longest = max((len(item.values) for item in series), default=0)
        labels: list[str] = []
        for source in self.plot.series:
            if any(source.categories):
                labels = list(source.categories)
                break
        if len(labels) >= longest:
            return labels
        return labels + [str(index + 1) for index in range(len(labels), longest)]

    def _axis_for(self, position: int) -> c.SourceChartAxis | None:
        """The axis this plot group names in its ``c:axId`` list, by position.

        Index 0 is the category axis and index 1 the value axis.  Going through the ids
        rather than through the first ``c:valAx`` in the plot area is what makes a
        secondary axis land on the right series -- and it also copes with
        ``real-financial-report.pptx``, which names a third id no axis element declares.
        """
        if position >= len(self.plot.axis_ids):
            return None
        wanted = self.plot.axis_ids[position]
        for axis in self.chart.axes:
            if axis.axis_id == wanted:
                return axis
        return None

    def _axis_of_kind(self, kind: str) -> c.SourceChartAxis | None:
        for axis in self.chart.axes:
            if axis.kind == kind:
                return axis
        return None

    def _scale(
        self, series: list[_Series], axis: c.SourceChartAxis | None
    ) -> tuple[float, float, float]:
        stacked = (self.plot.grouping or "clustered") in ("stacked", "percentStacked")
        if (self.plot.grouping or "") == "percentStacked":
            # Measured: PowerPoint labels 0%, 10% ... 100%.
            return 0.0, 1.0, 0.1

        numbers: list[float] = []
        if stacked:
            # A stacked bar reaches the sum of its category, and the positive and negative
            # halves of that category stack away from zero independently.
            length = max((len(item.values) for item in series), default=0)
            for index in range(length):
                column = [_at(item.values, index) for item in series]
                numbers.append(sum(value for value in column if value and value > 0))
                numbers.append(sum(value for value in column if value and value < 0))
        else:
            numbers = [value for item in series for value in item.values if value is not None]

        if not numbers:
            numbers = [0.0]
        minimum, maximum, unit = nice_axis_scale(
            min(numbers),
            max(numbers),
            horizontal=(self.plot.bar_direction or "col") == "bar",
            # A radar stops at the data rather than a whole unit past it: 0..5 of data
            # gives a 0..5 axis where the same data on a bar gives 0..6.
            strict=not self._is_radar,
        )

        return _apply_axis_limits((minimum, maximum, unit), axis)

    def _tick_texts(
        self, scale: tuple[float, float, float], axis: c.SourceChartAxis | None
    ) -> list[tuple[float, str]]:
        minimum, maximum, unit = scale
        format_code = self._axis_format(axis)
        if (self.plot.grouping or "") == "percentStacked":
            format_code = "0%"
        values = self._tick_values(scale)
        if not values:
            return [(minimum, format_number(minimum, format_code))]
        return [(value, format_number(value, format_code)) for value in values]

    def _axis_format(self, axis: c.SourceChartAxis | None) -> str | None:
        """The format code the value axis' labels are printed with.

        ``c:numFmt@sourceLinked`` means "take the cell's format", and the cell's format is
        what the value cache's own ``c:formatCode`` records -- so a source-linked axis
        falls through to the series.  An axis that says ``sourceLinked="0"`` has *chosen*
        its format, including when that choice is ``General``, and must not inherit.
        """
        if axis is None:
            return None
        if axis.number_format and axis.number_format != "General":
            return axis.number_format
        if axis.number_format_source_linked is False:
            return None
        for source in self.plot.series:
            if source.format_code and source.format_code != "General":
                return source.format_code
        return None

    # -- layout -------------------------------------------------------------------------

    def _font(self, *sources: "s.SourceTextBody | None") -> ChartFont:
        """The innermost ``c:txPr`` that names a size or a face wins, per property.

        Sources are given innermost first.  A ``c:txPr`` may name only one of the two --
        every axis in ``real-financial-report.pptx`` names ``Arial`` and a size while its
        legend names neither -- so size and face resolve independently rather than as a
        unit.
        """
        size = None
        typeface = None
        typeface_ea = None
        for source in (*sources, self.chart.text_properties):
            if size is None:
                size = _text_size(source)
            if typeface is None:
                typeface = _text_typeface(source)
            if typeface_ea is None:
                typeface_ea = _text_typeface_ea(source)
        family = self._resolve_typeface(typeface) if typeface else None
        family = family or self.style.font_family
        # The East Asian face is its own cascade and does not fall back to the Latin one
        # until everything else has failed -- see `pptx2svg.text.fontmap.east_asian_family`.
        family_ea = east_asian_family(
            self._resolve_typeface(typeface_ea) if typeface_ea else None,
            self.style.font_family_ea,
        )
        size = size or self.style.font_size
        return ChartFont(
            family=family,
            box=font_box(family, size),
            family_ea=family_ea,
            box_ea=font_box(family_ea, size) if family_ea else None,
        )

    def _label_font(self, axis: c.SourceChartAxis | None) -> ChartFont:
        return self._font(axis.text_properties if axis is not None else None)

    def _legend_font(self) -> ChartFont:
        legend = self.chart.legend
        return self._font(legend.text_properties if legend is not None else None)

    def _plot_rect(
        self,
        tick_texts: list[tuple[float, str]],
        categories: list[str],
        value_font: ChartFont,
        category_font: ChartFont,
        scale: tuple[float, float, float],
    ) -> _Rect:
        frame = self.frame
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")

        show_values = _labels_shown(value_axis)
        show_categories = _labels_shown(category_axis)
        horizontal = (self.plot.bar_direction or "col") == "bar"

        # The left column and the band under the plot each hold one axis' labels, and
        # `barDir` decides which.  Measured on the horizontal probe: its left inset,
        # 55.41 pt, is the same formula as a vertical chart's but fed the widest
        # *category* label instead of the widest tick.
        tick_labels = [text for _, text in tick_texts]
        down_left = categories if horizontal else tick_labels
        show_left = show_categories if horizontal else show_values
        show_bottom = show_values if horizontal else show_categories
        left_font = category_font if horizontal else value_font
        bottom_font = value_font if horizontal else category_font
        # `tickLblPos="nextTo"` means next to the *axis*, and a chart with negative values
        # has its category axis floating above the plot's lower edge.  PowerPoint then
        # reserves no band under the plot at all -- the negative probe's bottom inset is
        # 11.103 pt, the same half-label allowance as its top -- and prints the category
        # labels inside the plot, just under the zero line.
        labels_under_plot = show_bottom and (
            horizontal
            or scale[0] >= 0
            or (category_axis is not None and category_axis.tick_label_position == "low")
        )

        left = frame.left + EDGE_INSET_PT
        if show_left:
            widest = max((left_font.width(text) for text in down_left), default=0.0)
            left = (
                frame.left
                + FRAME_PADDING_PT
                + widest
                + left_font.box.descent
                + VALUE_LABEL_GAP_EM * left_font.size
            )

        # **``midCat`` centres the first and last category label on the plot's own edges**,
        # so half of each hangs outside it and the plot narrows to make room.  Measured on
        # the area probe, whose 21.073 / 209.472 plot became 26.433 / 189.632: ``Reader``
        # is 30.859 pt and half of it plus the 11.0 pt edge inset is 26.433, and
        # ``Renderer`` is 39.673 pt and half of it inside the same inset is 189.636.  Both
        # to 0.01 pt.  A line chart stating ``midCat`` is *not* measured; the same rule is
        # applied because the label placement it follows from is the same.
        overhang_left = overhang_right = 0.0
        if show_categories and categories and not horizontal and self._points_on_ticks:
            overhang_left = category_font.width(categories[0]) / 2
            overhang_right = category_font.width(categories[-1]) / 2
            left = max(left, frame.left + EDGE_INSET_PT + overhang_left)

        right = frame.right - EDGE_INSET_PT - overhang_right
        if horizontal and show_values:
            # The value axis runs along the bottom now, and its last label is centred on
            # the plot's right edge, so half of it hangs outside.  Measured 13.67 pt
            # against an 11.0 pt inset and a 5.34 pt label.
            right -= max((value_font.width(text) for text in tick_labels), default=0.0) / 2
        # Nothing overhangs the top of a horizontal chart, so it takes the plain inset.
        top = frame.top + (EDGE_INSET_PT if horizontal else self._top_inset(value_font.box))

        title = self._title_box()
        if title is not None:
            top += TITLE_BAND_LINES * title.line_height

        legend_bottom = 0.0
        legend = self._legend_position()
        # `c:overlay` draws the legend on top of the plot rather than beside it, so an
        # overlaid legend takes no space away.  Implemented from the schema; no chart
        # measured against PowerPoint sets it.
        if legend is not None and not self._legend_overlays():
            legend_font = self._legend_font()
            band = LEGEND_BAND_LINES * legend_font.box.line_height
            if legend in ("b",):
                legend_bottom = band
            elif legend in ("t", "tr"):
                top += band
            elif legend == "r":
                # The side band *replaces* the plain edge inset rather than adding to it:
                # it already ends in its own trailing pad.  Measured on
                # real-financial-report's two bar charts, whose legends are Japanese and
                # of different lengths -- both were over by exactly 11.0 pt, the inset.
                right = frame.right - self._legend_side_width(legend_font) - overhang_right
            elif legend == "l":
                # On the left the value-label column follows the legend instead of the
                # frame edge, so the band contributes one edge inset less.  Measured:
                # the legend-l probe's left inset is 85.067 pt and the legend-r probe's
                # right inset 74.994 pt for the same entry -- a difference of exactly the
                # 11.0 pt inset, with the 21.07 pt label column on top.
                left += self._legend_side_width(legend_font) - EDGE_INSET_PT

        if right - left < 1.0:
            right = left + 1.0
        # The bottom band comes last because a rotated category label's is a function of
        # the label's width against the band it has to fit, and the band is `right - left`
        # -- which the legend has only just finished moving.
        if labels_under_plot:
            # A `midCat` label is centred on a tick rather than laid into a band, so the
            # band rules -- wrapping and the 45 degree turn -- do not apply to it and the
            # reserve is one plain line.  Measured: the midCat probe's bottom inset is
            # 24.965 pt, the same one-line band as the `between` probe beside it.
            banded = [] if horizontal or self._points_on_ticks else categories
            bottom = frame.bottom - legend_bottom - self._bottom_label_band(
                bottom_font, banded, right - left
            )
        else:
            bottom = frame.bottom - legend_bottom - self._top_inset(value_font.box)
        if bottom - top < 1.0:
            bottom = top + 1.0
        return _Rect(left, top, right, bottom)

    def _labels_rotate(
        self, font: ChartFont, categories: list[str], plot_width: float
    ) -> bool:
        """Whether the category labels turn 45 degrees rather than staying level.

        **The rule is that the widest unbreakable token is wider than its own band.**
        Wrapping comes first and rotation is the last resort: ``MMM MM`` at 44.43 pt on a
        37.76 pt band came back level on two lines, while ``MMMMM`` at 41.65 pt on the
        same band turned.  The case that separates "widest token" from every other
        candidate is ``Fiscal MMMMMMMMMMMM 2012``, which has two spaces in it and turned
        anyway, because breaking it still leaves a 99.96 pt token.

        **One label that must turn turns all of them**, wrappable neighbours included: a
        chart of four ``MMM MM`` and one ``MMMMM`` came back with all five at 45 degrees
        and none of them broken.  So this is a decision for the axis, not per label.

        A radar facing the same problem wraps and never turns, so this is specifically
        what the *category* axis does.
        """
        if not categories:
            return False
        band = plot_width / len(categories)
        if band <= 0:
            return False
        widest = max(
            (font.width(token) for text in categories for token in text.split()),
            default=0.0,
        )
        return widest > band * ROTATED_LABEL_RATIO

    def _label_line_count(
        self, font: ChartFont, categories: list[str], plot_width: float
    ) -> int:
        """How many lines the tallest of the level category labels takes."""
        if not categories:
            return 1
        band = plot_width / len(categories)
        return (
            max((len(wrap_label(text, font, band)) for text in categories), default=1)
            or 1
        )

    def _bottom_label_band(
        self, font: ChartFont, categories: list[str], plot_width: float
    ) -> float:
        """How much of the frame the labels under the plot take.

        Level and on one line, that is one line box plus its gap.  Turned, it is
        :data:`ROTATED_LABEL_INSET_PT` plus the widest label's own width times sin 45,
        which fits six probes to within 0.03 pt.  **Wrapped, it is the level band plus one
        line box for every line after the first** -- and the line that sets it is the one
        needing the most lines, not the widest string, although no probe separates those
        two because in all twenty they were the same label.

        The extra line is the face's own line box, measured on a four-rung ladder in five
        faces at 10 pt.  The band grew by exactly the same amount from one line to two, two
        to three and three to four in every face, so this is a straight line and not a fit:

        ====================  ==========  ===========  =========
        face                  per line    line box     residual
        ====================  ==========  ===========  =========
        Calibri               12.205      12.207       -0.002
        Aptos                 12.205      12.207       -0.002
        Courier New           11.330      11.328       +0.002
        Times New Roman       11.075      11.074       +0.001
        Arial                 11.500      11.172       **+0.328**
        ====================  ==========  ===========  =========

        **Arial is the one refutation, and it has a name.**  PowerPoint's pitch is the
        face's full ``hhea`` line spacing -- ascender plus descender plus *lineGap* -- and
        Arial is the only one of the five whose lineGap is not zero: 67 units of 2048,
        which is 0.328 pt at 10 pt, exactly the residual above.  We cannot use that rule,
        because :mod:`pptx2svg.text.metrics` carries no lineGap and the substitute we
        would read one from disagrees with the face PowerPoint used: Tinos' is 87 where
        Office's own ``times.ttf`` is 0, so adding the gap would trade this 0.33 pt error
        on Arial for a 0.42 pt one on Times New Roman.  The line box alone is the better
        of the two, and the error it leaves is recorded rather than hidden.

        **Not capped for a turned label, and PowerPoint's is.**  A probe whose label is
        4.18 band widths wide reserved 85.63 pt where the formula asks for 113.95 -- but it
        reserved *less* than the probe one step below it, whose 2.92-band label took
        86.36 pt, so no clamp on the width reproduces both. Whatever PowerPoint does past
        about 90 pt of label was not identified, and a rule that fitted the rest and broke
        there is exactly what this file does not ship.

        A third observation narrows it without settling it.  ``real-financial-report.pptx``
        chart3 reserves 69.538 pt for a 96 pt label, which inverts through this formula to
        an implied width of 68.09 -- and its frame is 135 pt tall against the probe deck's
        181.1024.  The three implied widths are 90.85, 91.88 and 68.09, which are 0.502,
        0.507 and 0.504 of their frame heights: a straight line through them has an
        intercept of 0.01.  **Still not shipped.**  Under a pure cap the two probes would
        reserve the *same* amount and they differ by 0.73 pt, and fitting a rule on three
        points, one of them the deck it would be validated against, is how the Caladea
        claim got in.  One probe deck sweeping frame height settles it.

        The same export shows what PowerPoint does once the cap bites: it **truncates**.
        It drew ``プラット…`` for an eight-character category.  Nothing here does that, so
        our labels overflow where PowerPoint's are cut.
        """
        if self._labels_rotate(font, categories, plot_width):
            widest = max(font.width(text) for text in categories)
            return ROTATED_LABEL_INSET_PT + widest * _SIN_45
        # The line box is the *drawn* face's, which for a Japanese label is not the one
        # `<a:latin>` names -- see `ChartFont.box_for`.
        box = font.box_for(*categories)
        lines = self._label_line_count(font, categories, plot_width)
        return (
            FRAME_PADDING_PT
            + box.line_height * lines
            + CATEGORY_LABEL_GAP_EM * box.size
        )

    def _polar_region(self) -> _Rect:
        """The square-ish box a pie is drawn in.

        The same edge insets and legend band a bar chart uses -- measured on
        ``real-financial-report.pptx``'s doughnut, whose legend band came out 113.98 pt,
        the identical number that deck's bar charts reserve, and on a probe pie whose band
        matched the formula to 0.03 pt.

        A pie legends its categories and a radar its series, and the band is sized from
        whichever list it prints -- feeding a radar the category names put its centre
        13.4 pt out on the legend probe.
        """
        per_point = not self._is_radar
        frame = self.frame
        left, right = frame.left + EDGE_INSET_PT, frame.right - EDGE_INSET_PT
        top, bottom = frame.top + EDGE_INSET_PT, frame.bottom - EDGE_INSET_PT

        title = self._title_box()
        if title is not None:
            top = frame.top + TITLE_BAND_LINES * title.line_height + EDGE_INSET_PT

        legend = self._legend_position()
        if legend is not None and not self._legend_overlays():
            font = self._legend_font()
            band = LEGEND_BAND_LINES * font.box.line_height
            if legend == "b":
                bottom -= band
            elif legend in ("t", "tr"):
                top += band
            elif legend == "r":
                right = frame.right - self._legend_side_width(font, per_point=per_point)
            elif legend == "l":
                left += self._legend_side_width(font, per_point=per_point) - EDGE_INSET_PT
        if right - left < 1.0:
            right = left + 1.0
        if bottom - top < 1.0:
            bottom = top + 1.0
        return _Rect(left, top, right, bottom)

    def _draw_pie(
        self, region: _Rect, series: list[_Series], categories: list[str]
    ) -> None:
        """Slices, then their labels.

        Measured, all on PowerPoint's own export:

        * angle zero is **12 o'clock** and slices run **clockwise**, with
          ``c:firstSliceAng`` added as degrees;
        * the radius is **half the shorter side** of the region, centred in it;
        * ``c:holeSize`` is a percentage **of the outer radius**;
        * several series make concentric rings, innermost first, splitting the space
          between the hole and the outer radius evenly;
        * ``c:explosion`` shrinks the radius by ``1/(1+e)`` and offsets each slice by
          ``e`` of the *shrunk* radius along its own bisector.
        """
        if not series:
            return
        radius = min(region.width, region.height) / 2
        centre_x = (region.left + region.right) / 2
        centre_y = (region.top + region.bottom) / 2
        if radius <= 0:
            return

        explosion = max(
            (source.explosion or 0.0 for source in self.plot.series), default=0.0
        )
        for source in self.plot.series:
            for point in source.data_points:
                if point.explosion:
                    explosion = max(explosion, point.explosion)
        if explosion > 0:
            radius /= 1.0 + explosion / 100.0

        hole = (
            self.plot.hole_size if self.plot.hole_size is not None else DEFAULT_HOLE_SIZE
        ) / 100.0
        hole = min(max(hole, 0.0), 0.95)
        inner_base = radius * hole
        band = (radius - inner_base) / max(len(series), 1)
        start_angle = self.plot.first_slice_angle or 0.0

        for ring, item in enumerate(series):
            inner = inner_base + ring * band
            outer = inner + band
            total = sum(abs(value) for value in item.values if value is not None)
            if total <= 0:
                continue
            angle = start_angle
            for point in range(len(item.values)):
                value = item.values[point]
                if value is None or value == 0:
                    continue
                sweep = abs(value) / total * 360.0
                offset = self._slice_offset(item, point, angle + sweep / 2, radius)
                self._slice(
                    centre_x + offset[0],
                    centre_y + offset[1],
                    inner,
                    outer,
                    angle,
                    sweep,
                    fill=self._point_fill(item, point),
                    outline=item.point_outlines.get(point, item.outline),
                )
                angle += sweep

        self._draw_pie_labels(
            region, series, categories, centre_x, centre_y, radius, inner_base, start_angle
        )

    def _slice_offset(
        self, item: _Series, point: int, mid_angle: float, radius: float
    ) -> tuple[float, float]:
        """How far this slice is pushed out of the middle, in frame points."""
        explosion = self._point_explosion(point)
        if not explosion:
            return 0.0, 0.0
        distance = radius * explosion / 100.0
        radians = math.radians(mid_angle)
        return distance * math.sin(radians), -distance * math.cos(radians)

    def _point_explosion(self, point: int) -> float:
        for source in self.plot.series:
            for override in source.data_points:
                if override.index == point and override.explosion:
                    return override.explosion
            if source.explosion:
                return source.explosion
        return 0.0

    def _point_fill(self, item: _Series, point: int) -> m.Fill | None:
        fill = item.point_fills.get(point)
        if fill is None and point < len(item.vary_fills):
            fill = item.vary_fills[point]
        return fill if fill is not None else item.fill

    def _slice(
        self,
        cx: float,
        cy: float,
        inner: float,
        outer: float,
        start: float,
        sweep: float,
        *,
        fill: m.Fill | None,
        outline: m.Outline | None,
    ) -> None:
        commands = _ring_path(cx, cy, inner, outer, start, sweep)
        left = cx - outer
        top = cy - outer
        size = outer * 2
        local = _translate_path(commands, -left, -top)
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=left * EMU_PER_POINT,
                    offset_y=top * EMU_PER_POINT,
                    extent_width=size * EMU_PER_POINT,
                    extent_height=size * EMU_PER_POINT,
                ),
                geometry=m.CustomGeometry(
                    paths=[m.CustomGeometryPath(width=size, height=size, commands=local)]
                ),
                fill=fill,
                outline=outline,
            )
        )

    def _draw_pie_labels(
        self,
        region: _Rect,
        series: list[_Series],
        categories: list[str],
        centre_x: float,
        centre_y: float,
        radius: float,
        inner_base: float,
        start_angle: float,
    ) -> None:
        band = (radius - inner_base) / max(len(series), 1)
        for ring, item in enumerate(series):
            total = sum(abs(value) for value in item.values if value is not None)
            if total <= 0:
                continue
            shares = _percent_shares([
                abs(value) if value is not None else 0.0 for value in item.values
            ])
            angle = start_angle
            for point in range(len(item.values)):
                value = item.values[point]
                if value is None or value == 0:
                    continue
                sweep = abs(value) / total * 360.0
                mid = angle + sweep / 2
                angle += sweep
                labels = item.point_labels.get(point, item.labels)
                if labels is None or not labels.anything or labels.font is None:
                    continue
                parts: list[str] = []
                if labels.show_series and item.name:
                    parts.append(item.name)
                if labels.show_category and point < len(categories) and categories[point]:
                    parts.append(categories[point])
                if labels.show_percent:
                    parts.append(f"{shares[point]}%")
                if labels.show_value:
                    parts.append(
                        format_number(value, labels.number_format or item.format_code)
                    )
                if not parts:
                    continue
                fraction = PIE_LABEL_RADIUS.get(labels.position or "bestFit", 0.710)
                outer = inner_base + (ring + 1) * band
                distance = outer * fraction
                offset = self._slice_offset(item, point, mid, radius)
                radians = math.radians(mid)
                x = centre_x + offset[0] + distance * math.sin(radians)
                y = centre_y + offset[1] - distance * math.cos(radians)
                self._centred_label(parts, labels.font, x, y)

    def _centred_label(
        self, lines: list[str], font: ChartFont, x: float, y: float
    ) -> None:
        """A multi-line label whose block is centred on ``(x, y)``."""
        box = font.box
        width = max((font.width(line) for line in lines), default=0.0) + box.size
        block = box.line_height * len(lines)
        first = y - block / 2 + box.ascent
        for index, line in enumerate(lines):
            self._text(
                self._label_body(line, font, align="ctr"),
                left=x - width / 2,
                width=width,
                baseline=first + index * box.line_height,
                box=box,
            )

    def _top_inset(self, label: FontBox) -> float:
        """Space above the plot area for the topmost value label to sit in.

        ``max(11.0, 5.0 + lineHeight/2)`` fits Aptos at 8, 10 and 14 pt and Arial at 12 pt
        to within 0.02 pt.  The floor is what binds at 8 pt, which is why a purely
        proportional rule does not work.
        """
        return max(EDGE_INSET_PT, TOP_INSET_BASE_PT + label.line_height / 2)

    def _legend_side_width(self, font: ChartFont, *, per_point: bool = False) -> float:
        # Only the entries actually drawn: a series struck out by `c:legendEntry` would
        # otherwise reserve width for a label nobody sees, shifting the plot rectangle.
        deleted = self.chart.legend.deleted_entries if self.chart.legend else set()
        if per_point:
            # A pie legends its *categories*, not its series.
            names = self._legend_names(per_point=True)
        else:
            names = [
                source.name.plain or ""
                for index, source in enumerate(self.plot.series)
                if source.name is not None and index not in deleted
            ]
        widest = max((font.width(name) for name in names), default=0.0)
        key, key_gap = self._legend_key_size(font)
        natural = (
            widest
            + key
            + key_gap
            + (LEGEND_SIDE_LEAD_EM + LEGEND_SIDE_TRAIL_EM) * font.size
        )
        return min(natural, self.frame.width * LEGEND_SIDE_MAX_FRACTION)

    def _legend_overlays(self) -> bool:
        return self.chart.legend is not None and self.chart.legend.overlay

    def _legend_position(self) -> str | None:
        legend = self.chart.legend
        if legend is None:
            return None
        position = legend.position or "r"
        return position if position in ("b", "t", "l", "r", "tr") else "r"

    def _title(self) -> tuple[m.TextBody, FontBox] | None:
        """The title's resolved text body and the metrics of the face it will be drawn in.

        The face matters to the *layout*, not just the drawing: the title band is a
        multiple of its line height, and a 3 pt error there moves every bar.  So the body
        is resolved before the plot rectangle is computed and the metrics are read back
        off it, rather than being guessed from a default.  Resolving a chart title is a
        full trip through the text cascade, so the result is cached.
        """
        if self._title_cache is _UNSET:
            self._title_cache = self._build_title()
        return self._title_cache

    def _build_title(self) -> "tuple[m.TextBody, FontBox] | None":
        text = self._title_text()
        if text is None or self.chart.title is None:
            return None
        size = _title_size(self.chart.title)
        body = self._resolve_text(self.chart.title.rich, text, size, align="ctr")
        family, resolved_size = _first_run_font(body)
        size = resolved_size or size
        # A run that named no size would otherwise take the *renderer's* default, which
        # is a second place the number lives.  Stamping it makes the size the layout used
        # and the size drawn the same number by construction.
        for paragraph in body.paragraphs:
            for run in paragraph.runs:
                if run.properties.font_size is None:
                    run.properties.font_size = size
        return body, font_box(family, size)

    def _title_box(self) -> FontBox | None:
        title = self._title()
        return None if title is None else title[1]

    def _title_text(self) -> str | None:
        if self.chart.auto_title_deleted or self.chart.title is None:
            return None
        return self.chart.title.plain or None

    # -- drawing ------------------------------------------------------------------------

    def _draw_background(self, rect: _Rect) -> None:
        """The chart frame's own fill, then the plot rectangle's.

        A chart with no ``c:spPr`` at all is transparent -- the slide shows through, which
        is what PowerPoint drew for ``authoring-integration.pptx`` -- so an absent fill is
        not the same as a white one and nothing is emitted for it.
        """
        fill = self._resolve_fill(self.chart.fill)
        outline = self._resolve_outline(self.chart.outline)
        if (fill is not None and not isinstance(fill, m.NoFill)) or outline is not None:
            self._rect(self.frame, fill=fill, outline=outline)

        plot_fill = self._resolve_fill(self.chart.plot_area_fill)
        if plot_fill is not None and not isinstance(plot_fill, m.NoFill):
            self._rect(rect, fill=plot_fill, outline=None)

    def _draw_title(self) -> None:
        title = self._title()
        if title is None:
            return
        body, box = title
        baseline = self.frame.top + TITLE_BASELINE_ASCENTS * box.ascent
        self._text(
            body,
            left=self.frame.left,
            width=self.frame.width,
            baseline=baseline,
            box=box,
        )

    def _draw_gridlines(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        axis: c.SourceChartAxis | None,
    ) -> None:
        if axis is None or not axis.major_gridlines:
            return
        outline = self._axis_outline(axis.major_gridline_outline)
        horizontal = (self.plot.bar_direction or "col") == "bar"
        # A gridline runs *across* the value axis, so `barDir="bar"` turns them vertical.
        # Measured on the horizontal probe: three vertical lines at the ticks for 2, 4 and
        # 6, with the one for 0 left out because the category axis already draws it.
        axis_position = self._category_axis_position(rect, scale, None)
        for value in self._tick_values(scale):
            if horizontal:
                x = self._value_to_x(rect, value, scale)
                if abs(x - axis_position) < 0.01:
                    continue
                self._line(x, rect.top, x, rect.bottom, outline)
            else:
                y = self._value_to_y(rect, value, scale)
                if abs(y - axis_position) < 0.01:
                    continue
                self._line(rect.left, y, rect.right, y, outline)

    def _tick_values(self, scale: tuple[float, float, float]) -> list[float]:
        """Every major tick on the value axis, or nothing when the unit is unusable.

        `c:majorUnit` is attacker-controlled, so an interval small enough to ask for
        millions of gridlines returns an empty list rather than trying to draw them.
        """
        minimum, maximum, unit = scale
        span = maximum - minimum
        if unit <= 0 or not math.isfinite(span) or not math.isfinite(unit):
            return []
        steps = span / unit
        if not math.isfinite(steps) or steps <= 0 or steps > MAX_MAJOR_TICKS:
            return []
        return [minimum + index * unit for index in range(int(round(steps)) + 1)]

    def _draw_bars(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        if not series or not categories:
            return
        horizontal = (self.plot.bar_direction or "col") == "bar"
        grouping = self.plot.grouping or "clustered"
        stacked = grouping in ("stacked", "percentStacked")

        band = (rect.height if horizontal else rect.width) / len(categories)
        gap_width = self.plot.gap_width
        if gap_width is None:
            gap_width = DEFAULT_GAP_WIDTH
        slots = 1 if stacked else len(series)
        # `gapWidth` is the gap between category groups expressed as a percentage of one
        # bar's width, so the band holds `slots` bars plus `gapWidth/100` of one more.
        # The schema bounds it to 0..500 and a file is free to ignore that; at -100 on a
        # single series the divisor is exactly zero, which used to abort the conversion.
        bar_size = band / max(slots + gap_width / 100.0, MIN_BAR_SLOTS)
        overlap = self.plot.overlap if self.plot.overlap is not None else 0.0
        step = bar_size * (1.0 - overlap / 100.0)
        cluster = bar_size + step * (slots - 1)

        percent = grouping == "percentStacked"
        totals = _percent_totals(series) if percent else None

        for point in range(len(categories)):
            # A horizontal bar chart runs its category axis bottom-to-top, so category 0
            # is the *lowest* band.  Measured: "Reader" labels the bottom bar.
            band_start = (
                rect.bottom - (point + 1) * band if horizontal else rect.left + point * band
            )
            centre = band_start + band / 2
            positive_base = 0.0
            negative_base = 0.0
            for order, item in enumerate(series):
                value = _at(item.values, point)
                if value is None:
                    if self.chart.display_blanks_as == "zero":
                        value = 0.0
                    else:
                        continue
                if percent and totals is not None:
                    total = totals[point]
                    value = 0.0 if total == 0 else value / total

                if stacked:
                    start = positive_base if value >= 0 else negative_base
                    end = start + value
                    if value >= 0:
                        positive_base = end
                    else:
                        negative_base = end
                    slot = 0
                else:
                    start, end = 0.0, value
                    slot = order

                offset = centre - cluster / 2 + slot * step
                self._bar(rect, item, point, offset, bar_size, start, end, scale, horizontal)

    def _draw_areas(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """One filled band per series, closed back along the line beneath it.

        An area chart is a line chart that fills, and every difference from one is
        measured on an eighteen-chart probe:

        * **it closes to the zero line, not to the plot's floor.**  The negative probe's
          polygon runs along y = 117.069, which is where its value axis puts 0, with the
          plot's own bottom 53 pt lower.  Where the line crosses zero the polygon crosses
          itself, and PowerPoint leaves the bow tie exactly as it falls -- its fill rule is
          **nonzero winding**, which is what the renderer's default already is.
        * **the series are painted in series order, first at the back.**  Two probes
          settle that it is order and not size: swapping the two series' values swapped
          which one was hidden, and a probe whose first series covers the second entirely
          still emitted the first path first.  The fills are fully opaque -- every one of
          the eighteen came back at alpha 255 -- so the front series really does hide what
          is behind it, and that is PowerPoint's picture rather than an omission here.
        * **no outline unless the file asks.**  A probe stating ``<a:ln w="25400"/>`` got a
          second, stroked copy of the same closed path; the eight probes stating none got a
          bare fill.
        * **a blank splits the run and a run of one point draws nothing.**  The
          ``dispBlanksAs="gap"`` probe, whose middle value is missing, drew **no area at
          all**: both surviving runs are a single point, which has no area.

        Stacked and percent-stacked accumulate exactly as a stacked bar does, with each
        band's lower edge the running total *without* this series -- traced backwards, so
        the polygon is the strip between two lines rather than a fill to the axis.

        **A negative value inside a stack is not measured.**  A stacked *bar* runs its
        positive and negative halves away from zero independently; this keeps one running
        total, so a negative dips the band below the line beneath it.  No probe has a
        negative in a stacked area, and :meth:`_label_anchor` uses the same single total so
        the label cannot disagree with the band it sits in.
        """
        if not categories or not series:
            return
        xs = self._category_positions(rect, len(categories))
        blanks = self.chart.display_blanks_as or "gap"
        grouping = self.plot.grouping or "standard"
        stacked = grouping in ("stacked", "percentStacked")
        percent = grouping == "percentStacked"
        totals = _percent_totals(series) if percent else None
        zero = self._value_to_y(rect, 0.0, scale)

        bases = [0.0] * len(categories)
        for item in series:
            # ``(x, top, bottom)`` per point, or None where the run breaks.
            run: list[tuple[float, float, float] | None] = []
            for index in range(len(categories)):
                value = _at(item.values, index)
                if value is None:
                    if blanks == "zero":
                        value = 0.0
                    elif blanks == "span":
                        continue
                    else:
                        run.append(None)
                        continue
                if percent and totals is not None:
                    total = totals[index]
                    value = 0.0 if total == 0 else value / total
                if stacked:
                    start = bases[index]
                    bases[index] = start + value
                    top = self._value_to_y(rect, start + value, scale)
                    bottom = self._value_to_y(rect, start, scale)
                else:
                    top = self._value_to_y(rect, value, scale)
                    bottom = zero
                run.append((xs[index], top, bottom))

            for stretch in _split_runs(run):
                if len(stretch) < 2:
                    continue
                points = [(x, top) for x, top, _ in stretch]
                points += [(x, bottom) for x, _, bottom in reversed(stretch)]
                self._polygon(points, fill=item.fill, outline=item.outline)

    def _draw_lines(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """One polyline per series, with a marker at each point.

        Points sit at the **centre of their category band**, exactly where a bar would be
        -- measured on both the probe and the real line chart, whose `c:crossBetween` says
        ``between``.  ``midCat`` would put them on the band edges; no chart measured here
        uses it, so that spelling is **implemented from the schema and unverified**.
        """
        if not categories:
            return
        xs = self._category_positions(rect, len(categories))
        blanks = self.chart.display_blanks_as or "gap"

        for item in series:
            points: list[tuple[float, float] | None] = []
            for index in range(len(categories)):
                value = _at(item.values, index)
                if value is None:
                    if blanks == "zero":
                        value = 0.0
                    elif blanks == "span":
                        # "span" bridges the gap: the point is dropped and its neighbours
                        # join up, which is what leaving it out of the run does.
                        continue
                    else:
                        points.append(None)
                        continue
                points.append((xs[index], self._value_to_y(rect, value, scale)))

            for run in _split_runs(points):
                if len(run) > 1 and item.line is not None:
                    self._polyline(run, item.line, smooth=item.smooth)
            for point in points:
                if point is not None and item.marker_symbol:
                    self._marker(point, item)

    def _cross_between(self) -> str:
        """``between`` puts a point at its band's centre, ``midCat`` on the band edge.

        See :data:`DEFAULT_AREA_CROSS_BETWEEN` for why an area's default is the other one.
        """
        axis = self._axis_for(1) or self._axis_of_kind("valAx")
        stated = axis.cross_between if axis is not None else None
        if stated:
            return stated
        return DEFAULT_AREA_CROSS_BETWEEN if self._is_area else "between"

    @property
    def _points_on_ticks(self) -> bool:
        """Whether this chart's marks sit on the category ticks rather than in the bands.

        **Only a type that draws through :meth:`_category_positions` may answer yes.**  A
        bar chart still lays its bars into bands whatever `c:crossBetween` says, and Excel
        writes `midCat` on a column chart's value axis for the "Axis position: on tick
        marks" checkbox -- so reading it here without moving the bars too put every label
        up to 41 pt off the bar it names and the first one outside the plot.  What
        PowerPoint does with a `midCat` bar chart is **not measured**; leaving the bars and
        their labels in the bands together is the reading that cannot be self-contradictory.
        """
        return (self._is_area or self._is_line) and self._cross_between() == "midCat"

    def _category_positions(self, rect: _Rect, count: int) -> list[float]:
        """Where each category sits along the plot's width, in frame points."""
        if count <= 0:
            return []
        if self._points_on_ticks:
            step = rect.width / max(count - 1, 1)
            return [rect.left + index * step for index in range(count)]
        band = rect.width / count
        return [rect.left + (index + 0.5) * band for index in range(count)]

    def _marker(self, centre: tuple[float, float], item: _Series) -> None:
        """One marker, centred on its data point.

        ``c:size`` is the marker's **diameter in points** -- a 7 pt circle measured 6.96 pt
        across in the probe.
        """
        half = item.marker_size / 2
        x, y = centre
        box = _Rect(x - half, y - half, x + half, y + half)
        self._rect(
            box,
            fill=item.marker_fill,
            outline=item.marker_outline,
            preset=_MARKER_PRESETS.get(item.marker_symbol or "", "ellipse"),
        )

    def _bar(
        self,
        rect: _Rect,
        item: _Series,
        point: int,
        offset: float,
        size: float,
        start: float,
        end: float,
        scale: tuple[float, float, float],
        horizontal: bool,
    ) -> None:
        fill = item.point_fills.get(point)
        if fill is None and point < len(item.vary_fills):
            fill = item.vary_fills[point]
        if fill is None:
            fill = item.fill
        outline = item.point_outlines.get(point, item.outline)
        if end < start and item.invert_if_negative and point not in item.point_fills:
            # Measured: PowerPoint draws a negative bar white with a black 0.75 pt
            # outline, and draws it in the series colour when the file sets
            # `invertIfNegative` to 0 -- or leaves it out, which is the far more common
            # spelling; see the default above.
            #
            # The outline is `tx1` rather than literal black, which only a theme whose
            # `dk1` is not black can show: on the `probe-invert-on-fill` export the
            # hollow bar's outline came out #151515, this deck's `dk1`.  Left as
            # DEFAULT_AXIS_COLOR because that constant is shared with the axis and
            # gridline defaults and nothing has measured *those* against a non-black
            # `dk1` yet; the two should move together when one does.
            fill = m.SolidFill(color=m.ResolvedColor(hex="#FFFFFF"))
            outline = outline or m.Outline(
                width=INVERTED_BAR_OUTLINE_EMU,
                fill=m.SolidFill(color=m.ResolvedColor(hex=DEFAULT_AXIS_COLOR)),
            )

        if horizontal:
            x0 = self._value_to_x(rect, start, scale)
            x1 = self._value_to_x(rect, end, scale)
            box = _Rect(min(x0, x1), offset, max(x0, x1), offset + size)
        else:
            y0 = self._value_to_y(rect, start, scale)
            y1 = self._value_to_y(rect, end, scale)
            box = _Rect(offset, min(y0, y1), offset + size, max(y0, y1))
        if box.width <= 0 or box.height <= 0:
            return
        self._rect(box, fill=fill, outline=outline)

    def _category_axis_position(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        axis: c.SourceChartAxis | None,
    ) -> float:
        """Where the category axis crosses the value axis, in frame coordinates.

        A ``y`` for a column chart and an ``x`` for a bar one -- it is a position *along*
        the value axis either way.  ``c:crosses="autoZero"`` -- the default and what every
        chart in the corpus says -- puts it at value zero, which is the plot's low edge
        only while nothing is negative.  ``tickLblPos="low"`` pins it to the low end
        regardless; that spelling appears in ``real-financial-report.pptx`` but only over
        positive data, so its behaviour under a negative minimum is **implemented from the
        schema and not measured**.
        """
        horizontal = (self.plot.bar_direction or "col") == "bar"
        low, high = (rect.left, rect.right) if horizontal else (rect.bottom, rect.top)

        if axis is not None and axis.tick_label_position == "low":
            return low
        crosses = axis.crosses if axis is not None else None
        if crosses == "max":
            return high
        if crosses == "min":
            return low
        value = axis.crosses_at if axis is not None and crosses == "val" else 0.0
        if value is None:
            value = 0.0
        along = (
            self._value_to_x(rect, value, scale)
            if horizontal
            else self._value_to_y(rect, value, scale)
        )
        return min(max(along, min(low, high)), max(low, high))

    def _draw_axis_lines(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        value_axis: c.SourceChartAxis | None,
        category_axis: c.SourceChartAxis | None,
    ) -> None:
        # `barDir` swaps which axis is the horizontal line and which the vertical one, so
        # `c:delete` and `c:spPr` have to follow their own axis rather than a fixed side.
        # Measured on the horizontal probe: the value axis is the line along the bottom
        # and the category axis the one up the left.
        horizontal = (self.plot.bar_direction or "col") == "bar"
        crossing = self._category_axis_position(rect, scale, category_axis)

        if value_axis is not None and not value_axis.delete:
            outline = self._axis_outline(value_axis.outline)
            if horizontal:
                self._line(rect.left, rect.bottom, rect.right, rect.bottom, outline)
            else:
                self._line(rect.left, rect.top, rect.left, rect.bottom, outline)
        if category_axis is not None and not category_axis.delete:
            # The category axis sits where it crosses, which is the zero line and not the
            # plot's edge once anything is negative.
            outline = self._axis_outline(category_axis.outline)
            if horizontal:
                self._line(crossing, rect.top, crossing, rect.bottom, outline)
            else:
                self._line(rect.left, crossing, rect.right, crossing, outline)

    def _draw_labels(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        tick_texts: list[tuple[float, str]],
        categories: list[str],
        value_axis: c.SourceChartAxis | None,
        category_axis: c.SourceChartAxis | None,
        value_font: ChartFont,
        category_font: ChartFont,
    ) -> None:
        """Draw both axes' labels, on whichever side ``barDir`` puts them.

        A ``col`` chart labels values down the left and categories along the bottom; a
        ``bar`` chart does the opposite.  Both sides use the same two placements, so the
        orientation only decides which set of strings goes where.
        """
        horizontal = (self.plot.bar_direction or "col") == "bar"
        if _labels_shown(value_axis):
            if horizontal:
                self._labels_along_bottom(
                    rect,
                    [(self._value_to_x(rect, value, scale), text) for value, text in tick_texts],
                    value_font,
                    axis_y=rect.bottom,
                    centred_on_position=True,
                )
            else:
                self._labels_down_left(
                    rect,
                    [(self._value_to_y(rect, value, scale), text) for value, text in tick_texts],
                    value_font,
                )
        if _labels_shown(category_axis) and categories:
            if horizontal:
                band = rect.height / len(categories)
                self._labels_down_left(
                    rect,
                    # Category 0 is the lowest band on a horizontal chart.
                    [
                        (rect.bottom - (index + 0.5) * band, text)
                        for index, text in enumerate(categories)
                    ],
                    category_font,
                )
            else:
                band = rect.width / len(categories)
                axis_y = self._category_axis_position(rect, scale, category_axis)
                # Only labels that sit in the band *under* the plot may turn.  With
                # negative values the category axis floats up into the plot and
                # `_plot_rect` reserves no band at all, so turning them there would draw
                # into space nothing set aside -- and no probe measured what PowerPoint
                # does in that corner anyway.
                in_band = abs(axis_y - rect.bottom) < 0.01
                if self._points_on_ticks:
                    # **A ``midCat`` label is centred on its tick**, not in a band -- the
                    # area probe's three labels came back centred on the plot's left edge,
                    # its midpoint and its right edge, to 0.01 pt.  Rotation and wrapping
                    # are band rules and no probe exercises either here, so the level,
                    # unbroken placement is what is drawn.
                    self._labels_along_bottom(
                        rect,
                        list(zip(self._category_positions(rect, len(categories)), categories)),
                        category_font,
                        axis_y=axis_y,
                        centred_on_position=True,
                    )
                elif in_band and self._labels_rotate(
                    category_font, categories, rect.width
                ):
                    self._rotated_labels_along_bottom(
                        [
                            (rect.left + (index + 0.5) * band, text)
                            for index, text in enumerate(categories)
                        ],
                        category_font,
                        axis_y=axis_y,
                    )
                else:
                    self._labels_along_bottom(
                        rect,
                        [
                            (rect.left + index * band, text)
                            for index, text in enumerate(categories)
                        ],
                        category_font,
                        axis_y=axis_y,
                        width=band,
                    )

    def _rotated_labels_along_bottom(
        self,
        labels: list[tuple[float, str]],
        font: ChartFont,
        *,
        axis_y: float,
    ) -> None:
        """Category labels turned 45 degrees, reading up towards the axis.

        The measured anchor is the **far end of the rotated baseline**: it lands
        :data:`ROTATED_LABEL_OFFSET_X_PT` right of its band's centre and
        :data:`ROTATED_LABEL_OFFSET_Y_PT` below the axis, on all six probes to within
        0.15 pt.  The renderer rotates a shape about its own centre, so the box is placed
        by working that rotation backwards from the anchor rather than by rotating the
        text in place -- which is why the arithmetic below is not simply "left = x".
        """
        box = font.box
        radians = math.radians(ROTATED_LABEL_DEGREES)
        cos, sin = math.cos(radians), math.sin(radians)
        for position, text in labels:
            if not text:
                continue
            width = font.width(text) + box.size
            height = box.line_height * 1.5
            # Where the baseline's right-hand end sits inside the unrotated box, measured
            # from the box's centre.  The half em of padding is the same one every other
            # chart label carries, and the text is right-aligned against it.
            offset_x = width / 2 - box.size / 2
            offset_y = box.first_baseline - height / 2
            turned_x = offset_x * cos - offset_y * sin
            turned_y = offset_x * sin + offset_y * cos
            centre_x = position + ROTATED_LABEL_OFFSET_X_PT - turned_x
            centre_y = axis_y + ROTATED_LABEL_OFFSET_Y_PT - turned_y
            self.elements.append(
                m.ShapeElement(
                    transform=m.Transform(
                        offset_x=(centre_x - width / 2) * EMU_PER_POINT,
                        offset_y=(centre_y - height / 2) * EMU_PER_POINT,
                        extent_width=width * EMU_PER_POINT,
                        extent_height=height * EMU_PER_POINT,
                        rotation=ROTATED_LABEL_DEGREES,
                    ),
                    geometry=m.PresetGeometry(preset="rect"),
                    fill=None,
                    outline=None,
                    text_body=self._label_body(text, font, align="r"),
                )
            )

    def _labels_down_left(
        self,
        rect: _Rect,
        labels: list[tuple[float, str]],
        font: ChartFont,
        *,
        axis_x: float | None = None,
    ) -> None:
        """Right-aligned in the column left of the plot, each centred on its own y.

        ``axis_x`` is where the labels' own axis sits when it is not the plot's left edge.
        ``tickLblPos="nextTo"`` means next to the *axis*, and a scatter whose x range goes
        below zero has its value axis standing inside the plot: the negative-x probe's
        labels are right-aligned 9.23 pt left of the axis at 111.088 pt, which is the same
        ``descent + 0.645 em`` this column always uses -- just measured from the axis
        rather than from the frame.
        """
        box = font.box
        edge = rect.left if axis_x is None else axis_x
        width = edge - self.frame.left - box.descent - VALUE_LABEL_GAP_EM * box.size
        for y, text in labels:
            if not text:
                continue
            self._text(
                self._label_body(text, font, align="r"),
                left=self.frame.left,
                width=width,
                baseline=y + box.ink_centre,
                box=box,
            )

    def _labels_along_bottom(
        self,
        rect: _Rect,
        labels: list[tuple[float, str]],
        font: ChartFont,
        *,
        axis_y: float,
        width: float | None = None,
        centred_on_position: bool = False,
    ) -> None:
        """Below the axis: category labels centred in their band, ticks on theirs.

        The baseline hangs off the *category axis*, not the frame, because ``nextTo`` means
        what it says: on a chart with negative values the axis floats above the plot's
        lower edge and the labels follow it.  ``ascent + 0.615 em`` reproduces all five
        measurements -- Aptos at 8/10/14 pt, Arial at 12 pt, and the negative probe --
        with a worst residual of 0.63 pt, and is the same number as hanging the line's
        descender one frame padding above the frame whenever the axis *is* at the foot.

        A label too wide for its band is broken across lines, each one centred in the band
        under the one above.  **The block is top-aligned**, so the first baseline is where
        it would be for a one-line label however many lines follow: the Arial ladder put
        it 15.03 to 15.07 pt under the axis at one, two and three lines, and a chart whose
        one long label wrapped left every short label sitting on the *first* line with
        nothing beneath it.
        """
        box = font.box
        baseline = axis_y + box.ascent + CATEGORY_LABEL_GAP_EM * box.size
        for position, text in labels:
            if not text:
                continue
            if centred_on_position:
                # A value tick's label is centred on the tick, so the box is opened wide
                # either side of it and the text centred in that.
                span = font.width(text) + box.size
                left, box_width = position - span / 2, span
                lines = [text]
            else:
                left, box_width = position, width or box.size
                lines = wrap_label(text, font, box_width)
            for index, line in enumerate(lines):
                self._text(
                    self._label_body(line, font, align="ctr"),
                    left=left,
                    width=box_width,
                    baseline=baseline + index * box.line_height,
                    box=box,
                )

    def _draw_data_labels(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """Print `c:dLbls` beside each point.

        Positions are measured against a probe: ``outEnd`` (the bar default) puts the
        label's line box one :data:`DATA_LABEL_GAP_PT` beyond the bar's end, ``inBase``
        the same distance inside its base, ``inEnd`` just inside its end, and ``ctr`` on
        the bar's middle.  A line chart's default is ECMA's ``r``: centred on the point
        vertically, one marker radius plus a gap to its right.
        """
        if not categories:
            return
        horizontal = (self.plot.bar_direction or "col") == "bar"
        percent_totals = _percent_totals(series)

        for order, item in enumerate(series):
            for point in range(len(categories)):
                labels = item.point_labels.get(point, item.labels)
                if labels is None or not labels.anything or labels.font is None:
                    continue
                value = _at(item.values, point)
                if value is None:
                    continue
                text = self._label_text(
                    labels, item, categories, point, value, percent_totals
                )
                if not text:
                    continue
                geometry = self._label_anchor(
                    rect, series, order, item, point, value, scale, horizontal,
                    len(categories),
                )
                if geometry is None:
                    continue
                self._place_label(text, labels, geometry, horizontal)

    def _label_text(
        self,
        labels: _Labels,
        item: _Series,
        categories: list[str],
        point: int,
        value: float,
        percent_totals: list[float],
    ) -> str:
        """The label's lines, top to bottom.

        Measured on the multi-part probe: PowerPoint stacks series name, category name and
        value on **separate lines**, in that order, rather than joining them with the
        ``c:separator`` a single-line label would use.
        """
        parts: list[str] = []
        if labels.show_series and item.name:
            parts.append(item.name)
        if labels.show_category and point < len(categories) and categories[point]:
            parts.append(categories[point])
        if labels.show_percent:
            total = percent_totals[point] if point < len(percent_totals) else 0.0
            parts.append(format_number(value / total if total else 0.0, "0%"))
        if labels.show_value:
            parts.append(
                format_number(value, labels.number_format or item.format_code)
            )
        return "\n".join(parts)

    def _label_anchor(
        self,
        rect: _Rect,
        series: list[_Series],
        order: int,
        item: _Series,
        point: int,
        value: float,
        scale: tuple[float, float, float],
        horizontal: bool,
        count: int,
    ) -> "tuple[float, float, str] | None":
        """``(x, y, placement)`` for one label, in frame points.

        ``count`` is the *category* count the series were drawn against, not the longest
        series' length: a label has to land on the mark its own chart drew, and
        :meth:`_categories` pads the label list past the data when a file labels more
        categories than any series fills.
        """
        if self._is_area:
            # **An area label sits at the vertical centre of its own band** -- between the
            # series' own line and whatever is beneath it, which is the zero line for an
            # unstacked series and the running total for a stacked one.  Measured on two
            # probes: an unstacked 3/4/5 put its labels at 1.5, 2.0 and 2.5 on the value
            # axis, and a stacked pair put Beta's at 4, 6.5 and 5.5 -- the midpoints of its
            # segments, not of the stack.  Horizontally they are centred on the point.
            #
            # There is nothing to choose: **PowerPoint refuses a `c:dLblPos` on an area
            # chart outright**, opening the deck `[Repaired]` and declining to export it,
            # for `ctr` as much as for anything else.
            xs = self._category_positions(rect, count)
            if point >= len(xs):
                return None
            start = 0.0
            if (self.plot.grouping or "standard") in ("stacked", "percentStacked"):
                for earlier in series[:order]:
                    start += _at(earlier.values, point) or 0.0
            totals = _percent_totals(series)
            if (self.plot.grouping or "") == "percentStacked":
                total = totals[point] if point < len(totals) else 0.0
                if total == 0:
                    return None
                start /= total
                value = value / total
            middle = self._value_to_y(rect, start + value / 2, scale)
            return xs[point], middle, "centre"
        if self._is_line:
            xs = self._category_positions(rect, max(count, 1))
            x = xs[point] if point < len(xs) else rect.left
            return x + item.marker_size / 2, self._value_to_y(rect, value, scale), "right"

        box = self._bar_box(
            rect, series, order, item, point, value, scale, horizontal, count
        )
        if box is None:
            return None
        position = (item.labels.position if item.labels else None) or self._label_default()
        centre_x = (box.left + box.right) / 2
        centre_y = (box.top + box.bottom) / 2
        if horizontal:
            # The bar runs sideways, so "end" is its far edge in x.
            end, base = (box.right, box.left) if value >= 0 else (box.left, box.right)
            if position == "ctr":
                return centre_x, centre_y, "centre"
            if position == "inEnd":
                return end, centre_y, "inside-x"
            if position == "inBase":
                return base, centre_y, "outside-x-flip"
            return end, centre_y, "outside-x"
        end, base = (box.top, box.bottom) if value >= 0 else (box.bottom, box.top)
        if position == "ctr":
            return centre_x, centre_y, "centre"
        if position == "inEnd":
            return centre_x, end, "inside-y"
        if position == "inBase":
            return centre_x, base, "inside-base-y"
        return centre_x, end, "outside-y"

    def _label_default(self) -> str:
        """``c:dLblPos`` when the file states none.

        ``outEnd`` for a clustered bar, ``ctr`` for a stacked one.  ECMA-376 does not even
        *allow* ``outEnd`` on a stacked series -- there is no outside end to sit at, since
        the next segment starts there -- and PowerPoint agrees: on
        ``real-college-template``'s stacked chart, which states no ``c:dLblPos``, every
        label is inked inside its own segment, vertically centred, in the ``c:txPr``'s
        white.  Drawing them at ``outEnd`` put ours above the whole stack in eleven
        places.
        """
        if self._is_area:
            # The only one it has; PowerPoint repairs a file that names another.
            return "ctr"
        if self._is_scatter:
            # Measured: a scatter stating no `c:dLblPos` put its label's left edge 9.000 pt
            # right of the point, which is the marker's radius plus the line chart's own
            # 0.6 em gap -- ECMA's `r`, drawn exactly as a line chart's is.
            return "r"
        if self._is_line or self._is_polar:
            return "outEnd"
        stacked = (self.plot.grouping or "clustered") in ("stacked", "percentStacked")
        return "ctr" if stacked else "outEnd"

    def _place_label(
        self,
        text: str,
        labels: _Labels,
        geometry: tuple[float, float, str],
        horizontal: bool,
    ) -> None:
        x, y, placement = geometry
        font = labels.font
        assert font is not None
        box = font.box
        lines = text.split("\n")
        width = max((font.width(line) for line in lines), default=0.0) + box.size
        block = box.line_height * len(lines)

        if placement == "centre":
            baseline = y + box.ink_centre - block + box.line_height
            left, align = x - width / 2, "ctr"
        elif placement == "outside-y":
            baseline = y - DATA_LABEL_GAP_PT - box.descent
            left, align = x - width / 2, "ctr"
        elif placement == "inside-y":
            baseline = y + DATA_LABEL_INNER_GAP_PT + box.ascent
            left, align = x - width / 2, "ctr"
        elif placement == "inside-base-y":
            baseline = y - DATA_LABEL_GAP_PT - box.descent
            left, align = x - width / 2, "ctr"
        elif placement == "below":
            # A scatter's `b`: the mirror of `t` about the point, measured 0.70 pt loose.
            baseline = y + DATA_LABEL_GAP_PT + box.ascent
            left, align = x - width / 2, "ctr"
        elif placement == "right":
            baseline = y + box.ink_centre
            left = x + DATA_LABEL_LINE_GAP_EM * box.size
            align = "l"
        elif placement == "left":
            # A scatter's `l`: the mirror of `r`, the same 0.6 em off the marker's edge.
            baseline = y + box.ink_centre
            left = x - DATA_LABEL_LINE_GAP_EM * box.size - width
            align = "r"
        elif placement == "inside-x":
            baseline = y + box.ink_centre
            left = x - DATA_LABEL_INNER_GAP_PT - width
            align = "r"
        elif placement == "outside-x-flip":
            baseline = y + box.ink_centre
            left = x + DATA_LABEL_GAP_PT
            align = "l"
        else:  # outside-x
            baseline = y + box.ink_centre
            left = x + DATA_LABEL_GAP_PT
            align = "l"

        # Multi-line labels stack upwards from the anchor, so the *last* line is the one
        # nearest the bar; walk them in order from the first baseline.  `below` is the one
        # placement that hangs *under* its anchor, so its block grows downward and the
        # baseline computed above is already the first line's.
        first = baseline
        if placement != "below":
            first -= box.line_height * (len(lines) - 1)
        for index, line in enumerate(lines):
            self._text(
                self._label_body(line, font, align=align, color=labels.color),
                left=left,
                width=width,
                baseline=first + index * box.line_height,
                box=box,
            )

    def _bar_box(
        self,
        rect: _Rect,
        series: list[_Series],
        order: int,
        item: _Series,
        point: int,
        value: float,
        scale: tuple[float, float, float],
        horizontal: bool,
        categories: int,
    ) -> "_Rect | None":
        """The rectangle one bar occupies, recomputed for the label that sits on it.

        ``categories`` is the count :meth:`_draw_bars` laid the bands out against, so the
        two cannot drift when a file labels more categories than any series fills.
        """
        if categories <= 0:
            return None
        grouping = self.plot.grouping or "clustered"
        stacked = grouping in ("stacked", "percentStacked")
        band = (rect.height if horizontal else rect.width) / categories
        gap_width = self.plot.gap_width
        if gap_width is None:
            gap_width = DEFAULT_GAP_WIDTH
        slots = 1 if stacked else len(series)
        size = band / max(slots + gap_width / 100.0, MIN_BAR_SLOTS)
        overlap = self.plot.overlap if self.plot.overlap is not None else 0.0
        step = size * (1.0 - overlap / 100.0)
        cluster = size + step * (slots - 1)
        band_start = (rect.top if horizontal else rect.left) + (
            (categories - 1 - point) if horizontal else point
        ) * band
        centre = band_start + band / 2
        offset = centre - cluster / 2 + (0 if stacked else order) * step

        start = 0.0
        if stacked:
            for earlier in series[:order]:
                earlier_value = _at(earlier.values, point) or 0.0
                if (earlier_value >= 0) == (value >= 0):
                    start += earlier_value
        end = start + value

        if horizontal:
            x0 = self._value_to_x(rect, start, scale)
            x1 = self._value_to_x(rect, end, scale)
            return _Rect(min(x0, x1), offset, max(x0, x1), offset + size)
        y0 = self._value_to_y(rect, start, scale)
        y1 = self._value_to_y(rect, end, scale)
        return _Rect(offset, min(y0, y1), offset + size, max(y0, y1))

    def _legend_names(self, *, per_point: bool) -> list[str]:
        """Legend entry labels.

        A bar or line chart legends its **series**; a pie legends its **categories**,
        because its one series is the whole chart.  Measured: the probe pie's legend reads
        Alpha/Beta/Gamma/Delta, and its band matches the series formula fed those names.
        """
        deleted = self.chart.legend.deleted_entries if self.chart.legend else set()
        if not per_point:
            return [
                source.name.plain or ""
                for index, source in enumerate(self.plot.series)
                if source.name is not None and index not in deleted
            ]
        for source in self.plot.series:
            if any(source.categories):
                return [
                    name
                    for index, name in enumerate(source.categories)
                    if name and index not in deleted
                ]
        return []

    def _draw_legend(
        self,
        rect: _Rect,
        series: list[_Series],
        *,
        per_point: bool = False,
        categories: list[str] | None = None,
    ) -> None:
        position = self._legend_position()
        if position is None or not series:
            return
        legend = self.chart.legend
        deleted = legend.deleted_entries if legend else set()
        if per_point:
            item = series[0]
            names = categories or []
            entries = [
                (
                    index,
                    replace(item, name=name, fill=self._point_fill(item, index)),
                )
                for index, name in enumerate(names)
                if name and index not in deleted
            ]
        else:
            entries = [
                (index, item)
                for index, item in enumerate(series)
                if index not in deleted and item.name
            ]
        if not entries:
            return

        font = self._legend_font()
        box = font.box
        swatch, gap = self._legend_key_size(font)

        if position in ("b", "t", "tr"):
            widths = [swatch + gap + font.width(item.name or "") for _, item in entries]
            total = (
                sum(widths)
                + LEGEND_ENTRY_GAP_EM * box.size * (len(entries) - 1)
                + LEGEND_HORIZONTAL_LEAD_EM * box.size
            )
            baseline = (
                self.frame.bottom - LEGEND_BOTTOM_BASELINE_EM * box.size
                if position == "b"
                else self.frame.top + LEGEND_TOP_BASELINE_EM * box.size
            )
            x = (
                self.frame.left
                + (self.frame.width - total) / 2
                + LEGEND_HORIZONTAL_LEAD_EM * box.size
            )
            for (_, item), width in zip(entries, widths):
                self._legend_entry(item, x, baseline, swatch, gap, font)
                x += width + LEGEND_ENTRY_GAP_EM * box.size
            return

        # A side legend sits one lead gap outside the plot area.  Measured 15.996 pt at
        # 10 pt with the legend on the right, and the band on the left came out exactly
        # the same width, so the left case mirrors it against the frame edge.
        if position == "l":
            # Measured 10.996 pt from the frame's left edge in the legend-l probe, which
            # is the plain edge inset and not the 6.5 pt the label column starts at.
            x = self.frame.left + EDGE_INSET_PT
        else:
            x = rect.right + LEGEND_SIDE_LEAD_EM * box.size
        # A stacked legend is centred on the frame and each entry is centred in its row.
        # Measured against both bar charts in real-financial-report.pptx: baselines land
        # within 0.18 pt, where treating the row like the horizontal band's off-centre
        # line was 5.7 pt out.
        # A wrapped entry takes every row with it: PowerPoint opens the pitch rather than
        # letting two lines collide with the entry below.
        lines = max(
            (self._legend_entry_lines(item.name or "", font, x) for _, item in entries),
            default=1,
        )
        pitch = (
            LEGEND_ROW_PITCH_EM + LEGEND_WRAPPED_PITCH_EM * (lines - 1)
        ) * box.size
        y = self.frame.top + (self.frame.height - pitch * len(entries)) / 2
        for _, item in entries:
            self._legend_entry(item, x, y + pitch / 2 + box.ink_centre, swatch, gap, font)
            y += pitch

    def _legend_key_size(self, font: ChartFont) -> tuple[float, float]:
        """The legend key's width and the gap after it.

        A bar, a pie and a ``filled`` radar all take the square swatch.  A **line chart**
        and a radar drawn as lines take a line with the series' marker on it instead,
        which is 13.4 pt wider at 10 pt.

        The radar's 19.200 pt of rule and 2.025 pt of gap were measured once, at 10 pt, and
        assumed to scale with the font.  A line chart confirms the numbers and **refutes
        the scaling**: a right-hand legend gave 19.200 and 2.025 at 10 pt and the same
        19.200 and 2.025 at 14 pt, and a bottom legend the same again.  They are absolute
        points.  A series with ``c:symbol val="none"`` still gets the rule, without the
        marker.
        """
        if (
            self._is_line
            or self._is_scatter
            or (self._is_radar and self._radar_style != "filled")
        ):
            return LINE_LEGEND_KEY_PT, LINE_LEGEND_KEY_GAP_PT
        return LEGEND_SWATCH_EM * font.size, LEGEND_SWATCH_GAP_EM * font.size

    def _legend_entry_lines(self, name: str, font: ChartFont, x: float) -> int:
        """How many lines this entry needs once the band has capped its width."""
        swatch, gap = self._legend_key_size(font)
        available = self.frame.right - FRAME_PADDING_PT - (x + swatch + gap)
        natural = font.width(name)
        if available <= font.size or natural <= available:
            return 1
        return max(1, math.ceil(natural / max(available - font.size, 1.0)))

    def _legend_entry(
        self,
        item: _Series,
        x: float,
        baseline: float,
        swatch: float,
        gap: float,
        font: ChartFont,
    ) -> None:
        box = font.box
        centre = baseline - box.ink_centre
        if item.line is not None and (self._is_line or self._is_scatter or self._is_radar):
            # A line key: the stroke across the whole swatch width with the series'
            # marker centred on it.  Measured on the radar legend probe and confirmed on
            # a line chart's, where the marker's centre landed 0.17 pt off the midpoint.
            self._line(x, centre, x + swatch, centre, item.line)
            if item.marker_symbol:
                self._marker((x + swatch / 2, centre), item)
        else:
            self._rect(
                _Rect(x, centre - swatch / 2, x + swatch, centre + swatch / 2),
                fill=item.fill,
                outline=None,
            )
        # An entry wider than the band it sits in wraps rather than running out of the
        # frame.  PowerPoint wraps too -- `real-financial-report.pptx`'s doughnut legends
        # an 11-character category on two lines -- but it also opens the row pitch from
        # 1.8 em to 3.02 em to make room, which this does not; a wrapped entry therefore
        # overlaps the one below it.  Measured numbers are in ROADMAP.md.
        left = x + swatch + gap
        natural = font.width(item.name or "") + box.size
        available = self.frame.right - FRAME_PADDING_PT - left
        body = self._label_body(item.name or "", font, align="l")
        if natural > available > box.size:
            body = replace(body, body_properties=CHART_WRAPPED_TEXT_BODY)
        self._text(
            body,
            left=left,
            width=min(natural, max(available, box.size)),
            baseline=baseline,
            box=box,
        )

    # -- primitives ---------------------------------------------------------------------

    def _value_to_y(self, rect: _Rect, value: float, scale: tuple[float, float, float]) -> float:
        minimum, maximum, _ = scale
        span = maximum - minimum
        if span <= 0:
            return rect.bottom
        return rect.bottom - (value - minimum) / span * rect.height

    def _value_to_x(self, rect: _Rect, value: float, scale: tuple[float, float, float]) -> float:
        minimum, maximum, _ = scale
        span = maximum - minimum
        if span <= 0:
            return rect.left
        return rect.left + (value - minimum) / span * rect.width

    def _axis_outline(
        self, outline: m.Outline | s.SourceOutline | None
    ) -> m.Outline | None:
        """The stroke for an axis line or gridline, or ``None`` for "draw nothing".

        ``<a:ln><a:noFill/></a:ln>`` is an explicit *no line*, and it has to be told apart
        from an absent ``c:spPr``, which means "use the default".  Both arrive here as
        ``None`` out of :func:`resolve.view._resolve_outline` -- which collapses them --
        so the source element is what gets asked.  ``real-college-template``'s chart says
        it on its value axis and PowerPoint draws no line up the left of that plot; we
        drew the default black one, full plot height.
        """
        if isinstance(outline, s.SourceOutline) and isinstance(outline.fill, s.SourceNoFill):
            return None
        resolved = self._resolve_outline(outline) if outline is not None else None
        if resolved is not None and resolved.fill is not None:
            return resolved
        return m.Outline(
            width=DEFAULT_AXIS_LINE_EMU,
            fill=m.SolidFill(color=m.ResolvedColor(hex=DEFAULT_AXIS_COLOR)),
        )

    def _rect(
        self,
        box: _Rect,
        *,
        fill: m.Fill | None,
        outline: m.Outline | None,
        preset: str = "rect",
    ) -> None:
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=box.left * EMU_PER_POINT,
                    offset_y=box.top * EMU_PER_POINT,
                    extent_width=box.width * EMU_PER_POINT,
                    extent_height=box.height * EMU_PER_POINT,
                ),
                geometry=m.PresetGeometry(preset=preset),
                fill=fill,
                outline=outline,
            )
        )

    def _polyline(
        self, points: list[tuple[float, float]], outline: m.Outline, *, smooth: bool
    ) -> None:
        """A series' line, as one custom-geometry path in the frame's own space.

        Emitted as a single element rather than a segment per pair so the join between
        segments is a real line join -- drawing them separately leaves a notch at every
        vertex once the stroke is 2 pt wide, which is what a real chart uses.
        """
        left = min(x for x, _ in points)
        top = min(y for _, y in points)
        right = max(x for x, _ in points)
        bottom = max(y for _, y in points)
        width = max(right - left, 1e-6)
        height = max(bottom - top, 1e-6)

        local = [((x - left), (y - top)) for x, y in points]
        commands = _path_commands(local, smooth)
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=left * EMU_PER_POINT,
                    offset_y=top * EMU_PER_POINT,
                    extent_width=width * EMU_PER_POINT,
                    extent_height=height * EMU_PER_POINT,
                ),
                geometry=m.CustomGeometry(
                    paths=[
                        # The path's own space, which the renderer scales onto the shape
                        # box: the commands below are in points, so the extents must be
                        # too.  Giving them in EMU collapses the line to nothing.
                        m.CustomGeometryPath(
                            width=width, height=height, commands=commands
                        )
                    ]
                ),
                fill=m.NoFill(),
                outline=outline,
            )
        )

    def _polygon(
        self,
        points: list[tuple[float, float]],
        *,
        fill: m.Fill | None,
        outline: m.Outline | None,
    ) -> None:
        """A closed filled area, as one custom-geometry path in the frame's own space."""
        left = min(x for x, _ in points)
        top = min(y for _, y in points)
        width = max(max(x for x, _ in points) - left, 1e-6)
        height = max(max(y for _, y in points) - top, 1e-6)
        commands = " ".join(
            ("M" if index == 0 else "L") + f" {x - left:.4f} {y - top:.4f}"
            for index, (x, y) in enumerate(points)
        ) + " Z"
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=left * EMU_PER_POINT,
                    offset_y=top * EMU_PER_POINT,
                    extent_width=width * EMU_PER_POINT,
                    extent_height=height * EMU_PER_POINT,
                ),
                geometry=m.CustomGeometry(
                    paths=[
                        m.CustomGeometryPath(width=width, height=height, commands=commands)
                    ]
                ),
                fill=fill,
                outline=outline,
            )
        )

    def _line(
        self, x0: float, y0: float, x1: float, y1: float, outline: m.Outline | None
    ) -> None:
        if outline is None:
            # An explicit `a:noFill` stroke -- see :meth:`_axis_outline`.
            return
        self.elements.append(
            m.ConnectorElement(
                transform=m.Transform(
                    offset_x=min(x0, x1) * EMU_PER_POINT,
                    offset_y=min(y0, y1) * EMU_PER_POINT,
                    extent_width=abs(x1 - x0) * EMU_PER_POINT,
                    extent_height=abs(y1 - y0) * EMU_PER_POINT,
                ),
                geometry=m.PresetGeometry(preset="line"),
                outline=outline,
            )
        )

    def _label_body(
        self,
        text: str,
        font: ChartFont,
        *,
        align: str,
        color: m.ResolvedColor | None = None,
    ) -> m.TextBody:
        return m.TextBody(
            paragraphs=[
                m.Paragraph(
                    runs=[
                        m.TextRun(
                            text=text,
                            properties=m.RunProperties(
                                font_size=font.size,
                                font_family=font.family,
                                # Without this the renderer has one face for a label that
                                # PowerPoint drew in two, and resvg substitutes per text
                                # chunk rather than per glyph -- so the whole label would
                                # be drawn in whatever one face it settles on, at widths
                                # nothing computed.  Naming it is what makes the emitted
                                # `font-family` agree with `ChartFont.width`.
                                font_family_ea=font.family_ea,
                                color=color or self.style.color,
                            ),
                        )
                    ],
                    properties=m.ParagraphProperties(alignment=align),  # type: ignore[arg-type]
                )
            ],
            body_properties=CHART_TEXT_BODY,
        )

    def _text(
        self, body: m.TextBody, *, left: float, width: float, baseline: float, box: FontBox
    ) -> None:
        """Place a one-line text box so its baseline lands where the layout asked.

        The box's own top is derived from the renderer's first-baseline rule rather than
        assumed, so chart text stays aligned with the rest of the deck if that rule moves.
        """
        top = baseline - box.first_baseline
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=left * EMU_PER_POINT,
                    offset_y=top * EMU_PER_POINT,
                    extent_width=max(width, 1.0) * EMU_PER_POINT,
                    extent_height=(box.line_height * 1.5) * EMU_PER_POINT,
                ),
                geometry=m.PresetGeometry(preset="rect"),
                fill=None,
                outline=None,
                text_body=body,
            )
        )


#: Chart text sits in a box with no inset and no wrapping: the layout already decided
#: where every string goes, so letting the text engine re-wrap it would move it.
CHART_TEXT_BODY = m.BodyProperties(
    anchor="t",
    margin_left=0,
    margin_right=0,
    margin_top=0,
    margin_bottom=0,
    wrap="none",
)


#: Distinguishes "no title" from "not resolved yet" in the title cache.
_UNSET = object()


def _first_run_font(body: m.TextBody) -> tuple[str | None, float | None]:
    for paragraph in body.paragraphs:
        for run in paragraph.runs:
            return run.properties.font_family, run.properties.font_size
    return None, None


#: The same box, but allowed to wrap -- for a legend entry too wide for its band.
CHART_WRAPPED_TEXT_BODY = replace(CHART_TEXT_BODY, wrap="square")


def _labels_shown(axis: c.SourceChartAxis | None) -> bool:
    if axis is None:
        return True
    if axis.delete:
        return False
    return (axis.tick_label_position or "nextTo") != "none"


#: ``c:symbol`` -> the preset geometry that draws it.  PowerPoint's marker shapes are
#: simple enough that the preset catalogue already has every one.
_MARKER_PRESETS = {
    "circle": "ellipse",
    "dot": "ellipse",
    "square": "rect",
    "diamond": "diamond",
    "triangle": "triangle",
    "x": "mathMultiply",
    "plus": "mathPlus",
    "star": "star5",
    "dash": "rect",
}


def _wrap_to_width(text: str, font: "ChartFont", cap: float) -> list[str]:
    """Break ``text`` at spaces so no line is wider than ``cap``, if that is possible.

    A single word is never split: PowerPoint's own break put "Category" and "Three" on
    their own lines rather than hyphenating either.
    """
    if not text or cap <= 0 or font.width(text) <= cap:
        return [text]
    lines: list[str] = []
    current = ""
    for word in text.split(" "):
        if not word:
            continue
        candidate = f"{current} {word}" if current else word
        if current and font.width(candidate) > cap:
            lines.append(current)
            current = word
        else:
            current = candidate
    if current:
        lines.append(current)
    return lines or [text]


def _rotate_past_blank(
    points: list["tuple[float, float] | None"],
) -> list["tuple[float, float] | None"]:
    """Start the list at the point after the last blank, keeping cyclic order.

    A radar's points close into a ring, so the run running through index 0 is one run and
    not two.  Rotating makes that true of the straight list `_split_runs` walks.
    """
    blanks = [index for index, point in enumerate(points) if point is None]
    if not blanks or len(blanks) == len(points):
        return points
    start = blanks[-1]
    return points[start:] + points[:start]


def _split_runs(
    points: list["tuple[float, float] | None"],
) -> list[list[tuple[float, float]]]:
    """Split a series at its blanks, so ``dispBlanksAs="gap"`` really leaves a gap."""
    runs: list[list[tuple[float, float]]] = []
    current: list[tuple[float, float]] = []
    for point in points:
        if point is None:
            if current:
                runs.append(current)
            current = []
        else:
            current.append(point)
    if current:
        runs.append(current)
    return runs


def _path_commands(points: list[tuple[float, float]], smooth: bool) -> str:
    """SVG path data for one run of points, straight or smoothed.

    ``c:smooth`` is drawn as a Catmull-Rom spline converted to cubic Béziers, and
    **PowerPoint's tension is now measured: it is the plain 1/6**.  A smoothed five-point
    series exports as four cubics whose control points reproduce
    ``c1 = p1 + (p2 - p0) / 6`` and ``c2 = p2 - (p3 - p1) / 6`` to the 0.001 pt the PDF
    prints, on a line chart and on a scatter alike.

    **The ends were wrong and the same probe says so.**  Duplicating the terminal point --
    ``p0 = p1`` at the start -- puts the first control a *sixth* of the chord along it;
    PowerPoint's is a **third**: 87.073 against the 90.527 duplication gives, on a chord
    of 20.72 pt.  Reflecting the terminal point instead (``p0 = 2*p1 - p2``) makes the
    one-sided tangent the whole chord and reproduces both ends exactly, at 93.980 → 87.073
    and 31.822 → 59.447.
    """
    parts = [f"M {points[0][0]:.4f} {points[0][1]:.4f}"]
    if not smooth or len(points) < 3:
        for x, y in points[1:]:
            parts.append(f"L {x:.4f} {y:.4f}")
        return " ".join(parts)

    def _reflect(inner, edge):
        return (2 * edge[0] - inner[0], 2 * edge[1] - inner[1])

    for index in range(len(points) - 1):
        p1 = points[index]
        p2 = points[index + 1]
        p0 = points[index - 1] if index > 0 else _reflect(p2, p1)
        p3 = points[index + 2] if index + 2 < len(points) else _reflect(p1, p2)
        c1 = (p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6)
        c2 = (p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6)
        parts.append(
            f"C {c1[0]:.4f} {c1[1]:.4f} {c2[0]:.4f} {c2[1]:.4f} {p2[0]:.4f} {p2[1]:.4f}"
        )
    return " ".join(parts)


#: A cubic Bezier approximates a circular arc well up to a quarter turn; beyond that the
#: error becomes visible, so a sweep is cut into pieces no larger than this.
MAX_ARC_DEGREES = 90.0


def _arc_points(
    cx: float, cy: float, radius: float, start: float, sweep: float
) -> list[tuple[float, float]]:
    """A clockwise arc as ``M``-less Bezier control points, in frame coordinates.

    Angles are degrees **clockwise from 12 o'clock**, which is where PowerPoint starts and
    which way it runs -- measured on the real doughnut and on every probe pie.
    """
    pieces = max(1, math.ceil(abs(sweep) / MAX_ARC_DEGREES))
    step = sweep / pieces
    points: list[tuple[float, float]] = []
    for piece in range(pieces):
        a0 = math.radians(start + piece * step)
        a1 = math.radians(start + (piece + 1) * step)
        # Standard cubic approximation, in the (sin, -cos) frame that puts 0 at 12 o'clock.
        alpha = 4 / 3 * math.tan((a1 - a0) / 4)
        p0 = (cx + radius * math.sin(a0), cy - radius * math.cos(a0))
        p1 = (cx + radius * math.sin(a1), cy - radius * math.cos(a1))
        t0 = (radius * math.cos(a0), radius * math.sin(a0))
        t1 = (radius * math.cos(a1), radius * math.sin(a1))
        points.append((p0[0] + alpha * t0[0], p0[1] + alpha * t0[1]))
        points.append((p1[0] - alpha * t1[0], p1[1] - alpha * t1[1]))
        points.append(p1)
    return points


def _ring_path(
    cx: float, cy: float, inner: float, outer: float, start: float, sweep: float
) -> str:
    """One slice: the outer arc, then back along the inner one (or through the centre)."""
    sweep = max(min(sweep, 360.0), -360.0)
    begin = (cx + outer * math.sin(math.radians(start)),
             cy - outer * math.cos(math.radians(start)))
    parts = [f"M {begin[0]:.4f} {begin[1]:.4f}"]
    points = _arc_points(cx, cy, outer, start, sweep)
    for index in range(0, len(points), 3):
        c1, c2, end = points[index:index + 3]
        parts.append(
            f"C {c1[0]:.4f} {c1[1]:.4f} {c2[0]:.4f} {c2[1]:.4f} {end[0]:.4f} {end[1]:.4f}"
        )
    if inner <= 0:
        parts.append(f"L {cx:.4f} {cy:.4f}")
    else:
        back = (cx + inner * math.sin(math.radians(start + sweep)),
                cy - inner * math.cos(math.radians(start + sweep)))
        parts.append(f"L {back[0]:.4f} {back[1]:.4f}")
        points = _arc_points(cx, cy, inner, start + sweep, -sweep)
        for index in range(0, len(points), 3):
            c1, c2, end = points[index:index + 3]
            parts.append(
                f"C {c1[0]:.4f} {c1[1]:.4f} {c2[0]:.4f} {c2[1]:.4f} "
                f"{end[0]:.4f} {end[1]:.4f}"
            )
    parts.append("Z")
    return " ".join(parts)


def _translate_path(commands: str, dx: float, dy: float) -> str:
    """Shift an absolute path, so a slice can be drawn in its own shape box."""
    out: list[str] = []
    for token in commands.split(" "):
        out.append(token)
    numbers = [index for index, token in enumerate(out)
               if re.fullmatch(r"-?\d+\.?\d*", token)]
    for position, index in enumerate(numbers):
        value = float(out[index]) + (dx if position % 2 == 0 else dy)
        out[index] = f"{value:.4f}"
    return " ".join(out)


def _percent_shares(values: list[float]) -> list[int]:
    """Whole percentages that add up to 100.

    Measured: three equal values are labelled 34%, 33%, 33% -- rounding each on its own
    would give 33% three times and total 99.  Largest remainder, ties to the earlier
    index, reproduces it.  **One measurement**; the tie-breaking order in particular is
    only what that case shows.
    """
    total = sum(values)
    if total <= 0:
        return [0] * len(values)
    exact = [value / total * 100 for value in values]
    floors = [int(math.floor(value)) for value in exact]
    remainder = 100 - sum(floors)
    order = sorted(
        range(len(values)), key=lambda i: (-(exact[i] - floors[i]), i)
    )
    for index in order[:max(remainder, 0)]:
        floors[index] += 1
    return floors


def _span(numbers: list[float]) -> tuple[float, float]:
    """``(min, max)`` over the numbers, or ``(0, 0)`` when there are none."""
    if not numbers:
        return 0.0, 0.0
    return min(numbers), max(numbers)


def _apply_axis_limits(
    scale: tuple[float, float, float], axis: "c.SourceChartAxis | None"
) -> tuple[float, float, float]:
    """``c:min`` / ``c:max`` / ``c:majorUnit`` over a computed scale, and a usable span."""
    minimum, maximum, unit = scale
    if axis is not None:
        if axis.minimum is not None:
            minimum = axis.minimum
        if axis.maximum is not None:
            maximum = axis.maximum
        if axis.major_unit is not None and axis.major_unit > 0:
            unit = axis.major_unit
    if maximum <= minimum:
        maximum = minimum + (unit or 1.0)
    return minimum, maximum, unit


def _at(values: list[float | None], index: int) -> float | None:
    """One series' value at a category index; series need not be the same length."""
    return values[index] if index < len(values) else None


def _percent_totals(series: list[_Series]) -> list[float]:
    length = max((len(item.values) for item in series), default=0)
    totals: list[float] = []
    for index in range(length):
        total = 0.0
        for item in series:
            value = _at(item.values, index)
            if value is not None:
                total += abs(value)
        totals.append(total)
    return totals


def _default_run(body: s.SourceTextBody | None) -> s.SourceRunProperties | None:
    """``c:txPr``'s first ``a:defRPr``.

    A ``c:txPr`` is a one-paragraph text body whose only purpose is to carry defaults, so
    the first ``a:pPr/a:defRPr`` is the whole of it.
    """
    if body is None:
        return None
    for paragraph in body.paragraphs:
        properties = paragraph.properties
        if properties is not None and properties.default_run_properties is not None:
            return properties.default_run_properties
    return None


def _text_size(body: s.SourceTextBody | None) -> float | None:
    """``c:txPr``'s ``a:defRPr@sz``, in points."""
    run = _default_run(body)
    return run.font_size if run is not None and run.font_size else None


def _text_typeface(body: s.SourceTextBody | None) -> str | None:
    """``c:txPr``'s ``a:defRPr/a:latin@typeface``, unexpanded.

    Every axis in ``real-financial-report.pptx`` names ``Arial`` here while the theme's
    minor face is something else, and measuring the labels in the theme face instead put
    the plot area 1.7 pt off.  A ``+mn-lt``-style pointer comes back as-is; expanding it
    needs the theme and happens in the resolver.
    """
    run = _default_run(body)
    return run.typeface if run is not None else None


def _text_typeface_ea(body: s.SourceTextBody | None) -> str | None:
    """``c:txPr``'s ``a:defRPr/a:ea@typeface``, unexpanded.

    No chart in this corpus writes one -- every one of the five in
    ``real-financial-report.pptx`` names ``<a:latin typeface="Arial"/>`` and stops -- so
    this exists to stop a chart that *does* name an East Asian face being overridden by
    the theme.  A ``+mn-ea``-style pointer comes back as-is, the same as the Latin side.
    """
    run = _default_run(body)
    return run.typeface_ea if run is not None else None


def _title_size(title: c.SourceChartText | None) -> float:
    """A chart title's point size.

    PowerPoint draws an unstyled ``c:rich`` title at DrawingML's own default of 18 pt --
    measured, and the same size it gives a bare ``a:t`` on a slide -- not at the 10 pt the
    rest of a chart's text defaults to.
    """
    if title is not None and title.rich is not None:
        for paragraph in title.rich.paragraphs:
            for run in paragraph.runs:
                if run.properties is not None and run.properties.font_size:
                    return run.properties.font_size
            if (
                paragraph.properties is not None
                and paragraph.properties.default_run_properties is not None
                and paragraph.properties.default_run_properties.font_size
            ):
                return paragraph.properties.default_run_properties.font_size
    return 18.0


def accent_colors(resolve) -> list[m.ResolvedColor]:
    """The theme accent cycle a series falls back to when it has no fill of its own."""
    out: list[m.ResolvedColor] = []
    for key in ACCENT_KEYS:
        color = resolve(key)
        if color is not None:
            out.append(color)
    return out


def default_font_size(chart: c.SourceChart) -> float:
    return _text_size(chart.text_properties) or DEFAULT_CHART_FONT_PT


__all__ = [
    "CHART_TEXT_BODY",
    "ChartBuilder",
    "ChartStyle",
    "accent_colors",
    "default_font_size",
    "font_box",
    "format_number",
    "nice_axis_scale",
    "text_width",
]

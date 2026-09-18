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

import copy
import math
import re
from collections.abc import Sequence
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
#: own line height -- **two thirds of the face's ascent**, not a fraction of the em.
#:
#: This was ``0.615 * size`` for as long as every reading behind it was Aptos, whose
#: ascent is 0.939 em and whose two thirds is 0.626: one face cannot tell an em term from
#: an ascent term.  ``axis-inset`` does -- 24 charts, four faces, eight sizes, the plot
#: rectangle read off its own **gridlines** rather than off the tick labels' centres --
#: and the level band comes back as ``6.5 + (5/3) * ascent + descent`` with no em term in
#: it at all.  Solving each face for the coefficient on its ascent:
#:
#: ====================  =========  =========  ===============  ==========
#: face                  ascent/em  band/size  less descent/em  / ascent
#: ====================  =========  =========  ===============  ==========
#: Aptos                 0.9390     1.8466     1.5649           1.6665
#: Arial                 0.9053     1.7208     1.5089           1.6668
#: Times New Roman       0.8911     1.7015     1.4852           1.6667
#: Courier New           0.8325     1.6878     1.3875           1.6667
#: ====================  =========  =========  ===============  ==========
#:
#: -- four faces on 5/3 to four decimals, where a shared em term would have had to move
#: by 0.06 em across them.  The 6.5 pt is :data:`FRAME_PADDING_PT` exactly: differencing
#: it out leaves 6.500, 6.496, 6.499 and 6.497 over the four.
#:
#: The **drawn** baseline says the same thing independently, and says which of the band's
#: three terms the correction belongs to: the category label's baseline hangs
#: ``(5/3) * ascent`` below the axis line, measured on the same 24 charts at 1.672 +/-
#: 0.04 of the ascent, where against the *em* the same readings run 1.38 to 1.64 and are
#: no rule at all.  So the band is the padding, the baseline's own drop and the descender
#: hanging below it, and the gap is in the drop.
CATEGORY_LABEL_GAP_ASCENT = 2.0 / 3.0

#: The same gap as :data:`CATEGORY_LABEL_GAP_ASCENT`, in the em terms it was fitted in,
#: and used by the **turned** band alone -- :func:`rotated_label_anchor` and the rotated
#: branch of :meth:`ChartBuilder._bottom_label_band`.
#:
#: It stays here rather than following the level band onto the ascent because the turned
#: path's *other* constant was fitted against it: :data:`ROTATED_LABEL_HEADROOM_PT` is
#: solved from where PowerPoint cut a label, the anchor is inside that solution, and
#: moving the anchor without re-solving the headroom moves every truncation boundary.
#: The corpus has exactly one chart with turned labels -- ``real-financial-report``'s
#: chart3 -- and it sits 0.4 pt from such a boundary: on the ascent it loses the last
#: character of ``グローバル``, which PowerPoint keeps whole.
#:
#: **That chart cannot arbitrate, and the reason is worth recording.**  Its labels are
#: drawn in Yu Gothic, and pulling the font program straight out of
#: ``real-financial-report.pdf`` gives ``YuGothic-Regular`` at 0.8799 ascent, 0.2222
#: descent and a **0.5 em line gap**, where this library carries Noto Sans JP's 1.1600 and
#: 0.2880 for that name and no gap at all.  So its anchor is built from an ascent 32% too
#: large and its band from a box 1.84 pt too small, and the 0.95 pt this rule currently
#: lands from PowerPoint's 69.538 is two errors cancelling.  With the real numbers and the
#: gap carried, ``6.5 + (pitch + 60) * sin 45 + (2/3) * ascent`` gives 69.55 -- but that is
#: one reading resting on a metrics fix this change does not make.
#:
#: What would settle it: the rotated probe decks of ``tools/make_label_probe.py``, run in
#: faces whose ascents differ as ``axis-inset``'s do, re-solving the headroom and the
#: anchor together.  Until then the six readings in :data:`ROTATED_LABEL_HEADROOM_PT`'s
#: table are the only turned evidence, and they prefer the ascent -- which is why this is
#: an open question rather than a settled em.
CATEGORY_LABEL_GAP_EM = 0.615

#: The plot area's inset above the topmost value label: ``max(11.0, 5.0 + lineHeight/2)``.
#:
#: **Confirmed, not fitted.**  The four readings this was first written from were Aptos at
#: 8/10/14 and Arial at 12 pt; ``axis-inset`` puts 24 more behind it -- Aptos, Arial,
#: Times New Roman and Courier New at 6 through 28 pt -- and every one of them lands
#: within **0.002 pt** of this rule, which is the export's own coordinate quantisation.
#: The floor is real and it is :data:`EDGE_INSET_PT`: it binds for every face at 6 and
#: 8 pt and for Arial, Times New Roman and Courier New at 10 as well, and all of those
#: draw their top gridline exactly 11.000 pt below the frame.  It is the **line box** and
#: not the pitch, and Arial at eight sizes says so: half its line gap is 0.16 pt at 10 pt
#: and 0.46 at 28, both far outside the residual.
#:
#: Two readings out of the 36 do not obey the *band* beside it, and they are one thing:
#: 24 pt labels on a 90 pt frame keep this top inset and **shrink the bottom band**, by
#: 3.9 pt in Aptos and 0.2 in Arial -- see :meth:`ChartBuilder._bottom_label_band`.
TOP_INSET_BASE_PT = 5.0

#: **One legend row**, which is one number for every legend there is: the baseline-to-
#: baseline step down a side legend, the step down a wrapped horizontal one, and the height
#: a row takes out of the frame.  It is ``0.99 * lineBox + 6.0`` points -- a shade under the
#: face's line box, plus a pad in points that does not scale with anything.
#:
#: **Measured on 67 probe slides** (``tools/make_legend_probe.py``, the ``legend-side``,
#: ``legend-band`` and ``legend-face`` decks).  The precise reading is the *band*: a chart's
#: plot bottom is its category axis line, so differencing a 1-, 2- and 3-row legend against
#: a **no-legend control** of the same size and face gives the row with no other reserve in
#: it.  Four faces and five sizes come back at the same ratio to 0.0002 em:
#:
#: | face | size | row less the 6.0 pt | our line box | ratio |
#: | --- | --- | --- | --- | --- |
#: | Aptos | 8 | 9.667 | 9.766 | 0.9899 |
#: | Aptos | 10 | 12.083 | 12.207 | 0.9898 |
#: | Aptos | 18 | 21.750 | 21.973 | 0.9899 |
#: | Arial | 10 | 11.063 | 11.172 | 0.9902 |
#: | Arial | 14 | 15.484 | 15.641 | 0.9900 |
#: | Times New Roman | 10 | 10.964 | 11.074 | 0.9900 |
#: | Courier New | 10 | 11.217 | 11.328 | 0.9902 |
#:
#: **The box, not the pitch**: Arial is the face that separates them and its 0.327 pt line
#: gap is not in the row.  The 1% is carried as a ratio rather than chased: it is the same
#: 1% for every face and size here, so it is a difference between our ``hhea`` box and
#: whatever PowerPoint measures, not a per-face correction.
#:
#: **It is the Latin face's box even when no Latin is drawn.**  A legend of Japanese names
#: with ``a:latin="Arial"`` and ``a:ea="Yu Gothic"`` takes Arial's 17.06 pt row, not the
#: 20.3 pt Yu Gothic's 1.448 em box would give -- which is why
#: ``real-financial-report.pptx``'s Japanese legends measure 18.08 pt: their Latin face is
#: Calibri, and Calibri's box is Aptos' to the unit.
#:
#: Superseding ``LEGEND_ROW_PITCH_EM = 1.8`` and ``LEGEND_BAND_LINES = 2.0``, which were
#: each fitted at 10 pt in one face and are out by 4.7 and 10.2 pt at 18 pt.
LEGEND_ROW_PITCH_RATIO = 0.99
LEGEND_ROW_PITCH_PT = 6.0

#: The padding a **horizontal** legend's band carries on top of its rows, and the amount
#: that padding gives back once the legend wraps.  ``band = rows * pitch + 6.0``, less
#: 1.5 pt from the second row on -- so a wrapped legend's last row overhangs the band it
#: was given by a point and a half.  Measured at five sizes and three row counts on the
#: ``legend-band`` deck, where the 1.5 comes back identical at every one of them (14.167 vs
#: 15.668 at 8 pt, 16.583 vs 18.083 at 10, 19.003 vs 20.503 at 12, 21.419 vs 22.920 at 14,
#: 26.250 vs 27.750 at 18) and is therefore points rather than ems.  The same band was
#: measured for ``legendPos`` ``t`` and ``b``.
LEGEND_BAND_PAD_PT = 6.0
LEGEND_WRAP_BAND_TRIM_PT = 1.5

#: Legend swatch side and the gap after it, in ems.  Measured 5.4923 pt and 2.3711 pt at
#: 10 pt.
LEGEND_SWATCH_EM = 0.549
LEGEND_SWATCH_GAP_EM = 0.237
#: The most of the frame's width a side legend may take before its entries wrap.  Two
#: things come off it: **whether** an entry wraps, and the column it wraps inside once it
#: does -- see :meth:`ChartBuilder._legend_side_metrics`, where the band then shrinks back
#: to the widest line the wrap produced.
#:
#: **One measurement fixes it and one probe disagrees with it.**
#: `real-financial-report.pptx`'s doughnut legends an 11-character Japanese category whose
#: natural band would be 143.96 pt; PowerPoint reserved 113.98 pt and broke the name after
#: its eighth character, which is this fraction of the 285 pt frame less the key and the
#: pads to 0.02 pt.  The ``legend-side`` deck's frame sweep then walks the same four-word
#: Latin name across 240, 300, 360 and 480 pt frames and three of the four agree -- but at
#: 360 pt PowerPoint broke 2 + 2 where a 110.04 pt column holds three words at 107.60, so
#: that slide wants the column under 0.3933 of the frame where the doughnut wants 0.3999
#: or more.  **No single fraction fits both** and no pad or trailing-space reading closes
#: the 2.4 pt; a column of about 0.29 * frame fits all five but then has to be a second
#: constant, because at 480 pt the *unwrapped* name is 0.30 of the frame and stays on one
#: line.  The doughnut is the reading a corpus deck is scored on, so the fraction stays
#: where it measured; the Latin slide is the residual.
LEGEND_SIDE_MAX_FRACTION = 0.40

#: Padding either side of a side legend.
LEGEND_SIDE_LEAD_EM = 1.60
LEGEND_SIDE_TRAIL_EM = 1.01

#: The **cell** a horizontal legend entry puts its key in, before the name.  It is not the
#: drawn key: the key is *centred* in it, so the cell is what the layout advances by and
#: the swatch is what the eye sees.  Twice the swatch for a swatch key, and 1.25x the rule
#: for a line key -- 10.985 pt and 24.000 pt at 10 pt.  Both fall out of the same fit as
#: :data:`LEGEND_ENTRY_SLACK` below and neither is adjustable without it.
LEGEND_ENTRY_KEY_EM = 2 * LEGEND_SWATCH_EM
LINE_LEGEND_ENTRY_KEY_PT = 24.000

#: **The gap between entries in a horizontal legend, and it is not a constant.**  The run
#: is padded by a fifth of its own natural width and that slack is split into ``n + 1``
#: equal parts -- one before the first entry, one after the last, and one between each
#: pair -- so the gap grows with the names rather than shrinking, and a wider frame does
#: not touch it.
#:
#:     width(i) = key cell + advance(name i)
#:     gap      = 0.2 * sum(width) / (n + 1)
#:
#: Superseding ``LEGEND_ENTRY_GAP_EM = 0.5``, which ROADMAP.md 3.3 refuted on four charts
#: that solved for four different gaps.  Those four are what this reproduces: gallery
#: slide 1 asked for 0.77 em, slide 3 for 1.03, slide 17 for 1.12 and ``combo-legend``'s
#: ``l-bottom`` for 1.15, and one rule gives all four because the gap is a function of the
#: entries and not of the chart.
#:
#: **Measured on 88 probe slides across four decks** (``tools/make_legend_probe.py``)
#: sweeping entry count 2-7, name width, key type, frame width 200-720 pt, font size
#: 8-18 pt and ``legendPos`` ``b``/``t``, plus nine charts of ``chart-gallery.pptx`` that
#: were not fitted.  Worst residual **0.009 pt** on the gap and **0.035 pt** on the first
#: key's x.  The frame sweep is the load-bearing one: a 240 pt frame and a 720 pt frame
#: draw the same entries at the same pitch to 0.001 pt, which is what rules out the legend
#: being *distributed* across an available width.
LEGEND_ENTRY_SLACK = 0.2

#: The cap on that slack.  The run plus its ``n + 1`` gaps never exceeds this much of the
#: frame; past it the gap is whatever is left over, which is the one place the layout does
#: distribute.  Measured at exactly 0.9 on nine slides that cross the threshold at three
#: frame widths (300, 480 and 720 pt) and three entry counts.
#:
#: Once the entries alone pass 0.9 of the frame PowerPoint **wraps the legend onto more
#: rows**, at the same pitch a side legend uses, in a grid of equal columns as wide as the
#: widest entry.  The number of columns is what this cap fixes -- see
#: :meth:`ChartBuilder._legend_grid`, which is measured on the ``legend-row`` deck.
LEGEND_BAND_MAX_FRACTION = 0.9

#: How far right of the frame's centre the run's own centre lands.  Frame-independent and
#: size-independent: the same 0.75 pt at 200 pt and 720 pt of frame, and at 8 pt and 18 pt
#: of type.  Replaces ``LEGEND_HORIZONTAL_LEAD_EM = 0.386``, whose 1.93 pt of shift was
#: this constant plus the error in the gap it was fitted beside.
LEGEND_HORIZONTAL_OFFSET_PT = 0.75


def legend_row_pitch(box: "FontBox", lines: int = 1) -> float:
    """The height of one legend row whose deepest entry takes ``lines`` lines.

    ``box`` is the **Latin** face's, which is the one PowerPoint measures with; see
    :data:`LEGEND_ROW_PITCH_RATIO` for the 67 slides behind the two constants and for the
    Japanese legend that says so.

    The ``lines`` term is the same quantity again: a row that has to hold a second line is
    exactly one more ``0.99 * lineBox`` tall, measured at 10 pt for one, two, three and
    four lines (18.083, 30.12, 42.24 and 54.36 pt) and at 14 and 18 pt for three.  It is
    **not** the line the entry is drawn with: the drawn drop is the box of the face that
    draws it, which for a Japanese entry in the same 10 pt legend is 17.04 pt against this
    12.08.  So a wrapped entry's second line does not sit where its row's arithmetic puts
    it -- 0.2 pt apart for Latin, 5 pt for Japanese -- and the row is sized for the first
    of those.
    """
    return LEGEND_ROW_PITCH_RATIO * box.line_height * max(lines, 1) + LEGEND_ROW_PITCH_PT

#: The title band, and its baseline inside it, as multiples of the line height and the
#: ascent.  Only one title was measurable (18 pt Arial, in two probes and the fixture, all
#: agreeing): the band is 29.70 pt against a 20.109 pt line height, and the baseline sits
#: 24.52 pt below the frame top against a 16.295 pt ascent.  These are the measured
#: ratios rather than the tidy 1.5 both are close to -- rounding the band cost 0.46 pt of
#: plot height, which moved every gridline by a pixel.  They are a one-font fit and should
#: be re-measured if a chart with a differently sized title ever disagrees.
#:
#: **One now does, and it is left here as the honest single-font fit rather than refitted
#: to two disagreeing points.**  A region height is readable as a length wherever an
#: ``ofPieChart``'s radius is capped by it, and two such readings want two different
#: ratios: ``ofpie-clamp``'s default-size title (18 pt, a 20.109 pt line box) wants a band
#: of 30.03 against the 29.70 this gives, a ratio of 1.4931, while ``chart-gallery``
#: slide 9's 14 pt title (15.641 pt line box) wants 20.44 against 23.10, a ratio of
#: **1.3068**.  The same 2.66 pt appears independently on slide 6, whose largest bubble is
#: 64.27 pt where PowerPoint draws 63.65, because the bubble region takes the same band.
#: A straight line through the two points has a negative intercept, so this is not a
#: constant times the line height and no two-point fit is worth having.  A title-size sweep
#: read through that clamp would settle it; see ROADMAP.md 3.2b.
TITLE_BAND_LINES = 1.4769
TITLE_BASELINE_ASCENTS = 1.5046

#: Default chart text size, in points.  ECMA-376's chart default and what PowerPoint drew
#: for every axis label and legend entry with no ``c:txPr``.
DEFAULT_CHART_FONT_PT = 10.0

#: Axis, tick and gridline defaults when no ``c:spPr`` says otherwise.
DEFAULT_AXIS_LINE_EMU = 6350.0
DEFAULT_AXIS_COLOR = "#000000"

#: What a 3-D scene's **floor** is outlined with, where no black line covers it.
#:
#: The floor is a quadrilateral seen in plan and three of its four sides are drawn over by
#: something black -- the category axis along the front, and the value axis' own gridline
#: at zero back along the depth and across the back.  The fourth, the side the depth leads
#: towards, is the one place the floor's own stroke shows, and it is this grey rather than
#: the black beside it: ``#898989`` on gallery slide 13 and on the ``view3d-mesh`` probes,
#: which are two different themes, so it is a constant and not a theme colour.  Its width
#: reads as the same :data:`DEFAULT_AXIS_LINE_EMU` the lines beside it use.
VIEW_3D_FLOOR_COLOR = "#898989"

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

#: The plot's bottom inset once the labels turn is the level band with the label's own box
#: **turned through 45 degrees**: ``FRAME_PADDING_PT + (lineBox + width) * cos 45 +
#: CATEGORY_LABEL_GAP_EM * size``.  That is the same three terms the level band has, with
#: the one that runs along the baseline now running diagonally, and it is a decomposition
#: rather than a refit: the flat 21.39 pt this replaces is what it gives for the 10 pt
#: Aptos the original six probes were all drawn in, to 0.10 pt.
#:
#: **What forced the decomposition is the size sweep.**  A flat constant is right at one
#: size and nowhere else: probe charts across 6, 8, 10, 12, 14, 18, 20 and 24 pt put the
#: constant at 14.70 pt through 39.94, and the formula reproduces all nine face/size pairs
#: to 0.28 pt.  Split into the two pieces the cap needs separately:
#:
#: ====================  ==========  =========  ==========  ==========  ==========
#: face and size         pad below   anchor A   sum, drawn  sum, rule   on ascent
#: ====================  ==========  =========  ==========  ==========  ==========
#: Arial 6 pt                 7.582      7.117      14.699      14.93       14.86
#: Arial 10 pt                8.062     12.435      20.497      20.55       20.44
#: Arial 14 pt                8.302     17.643      25.945      26.17       26.01
#: Arial 20 pt                9.502     24.855      34.357      34.60       34.37
#: Arial 24 pt                9.920     30.022      39.942      40.22       39.95
#: Aptos 10 pt                8.782     12.604      21.386      21.28       21.39
#: ====================  ==========  =========  ==========  ==========  ==========
#:
#: ``A`` is the drop from the axis line to the **far end of the rotated baseline** and is
#: ``ascent * cos 45 + CATEGORY_LABEL_GAP_EM * size``; the pad under the deepest pen is
#: ``FRAME_PADDING_PT + descent * cos 45``.  Both are read straight off the export -- the
#: pen positions are in the PDF -- rather than solved for.  The unturned term is the level
#: band's own gap, which is the other half of why this reads as the same band rather than
#: a second one.
#:
#: The fifth column is what the rule gives with that term on the **ascent** instead, as
#: the level band now has it: it is better on every one of the six -- worst residual 0.16
#: pt against 0.28 -- and it is *not* what the code does, because these same readings
#: solved :data:`ROTATED_LABEL_HEADROOM_PT` with the em term inside them.  See
#: :data:`CATEGORY_LABEL_GAP_EM` for what re-solving the pair would take.

#: How far the deepest pen of a turned label may drop below the axis: **half the frame's
#: height, less this**.  Past it PowerPoint cuts the label rather than reserving more.
#:
#: This is the cap the file could not name for three observations.  Three probe decks,
#: 112 charts, exported and read back as pen positions: frame heights from 70 to 380 pt,
#: label sizes 6 to 24 pt, six frame widths, four category counts and four faces.  Every
#: chart brackets the allowance from **both** sides -- the prefix PowerPoint kept fits and
#: one more character does not -- which is what turns a drawn label into a measurement.
#:
#: **The shape is settled and the offset is not, quite.**  The allowance runs with the
#: frame's height at a slope of ``sin 45`` over twelve heights and at every size, so the
#: rule is "half the height, less a fixed drop"; what 95 two-sided readings do *not* do is
#: agree on one number for that drop.  Solved per size against the pen positions it comes
#: out at 8.6 pt for a 6 pt label, 7.6 at 8 pt, 6.4 at 10 and 12, and unconstrained above
#: 14 -- non-monotone, so it is not a linear term in the size either.  6.25 is what the
#: densest family gives (46 readings at 10 pt bracket it into (6.20, 6.31]) and it
#: reproduces **78 of the 95** exactly, drawn string and band alike.  The other 17 miss by
#: at most one and a half characters: 6 and 8 pt labels are cut about 2 pt of width later
#: than PowerPoint cuts them, and three legend probes by 0.19 for a reason that is not this
#: constant (see below).  An uncapped reserve is wrong by 30 pt on the same charts.
#:
#: **What it is not.**  Not a fixed number of points: the cap runs from 28.3 pt of band at
#: a 70 pt frame to 183.6 at a 380 pt one.  Not a fraction of the plot or of the band: six
#: frame widths from 150 to 500 pt and category counts of 3, 5 and 8 all reserved
#: *identically*, which is the same trap the tick rule fell into and is checked here the
#: same way.  Not a line count and not a character count: ``MMMM...IIII`` and
#: ``IIII...MMMM`` are the same 32 characters and the same 177.7 pt, and PowerPoint cut
#: them 2 pt apart in *width*.
#:
#: **The rival, kept because it is the other half of the residual.**  ``band <=
#: (height - EDGE_INSET_PT) / 2`` -- the plot keeping half the frame outright, with no
#: anchor term -- brackets every reading at 10 pt and below, including the 6 and 8 pt ones
#: this constant misses, and fails from 12 pt up, where it leaves one to two characters of
#: band unused.  The two rules are the same rule with the anchor counted and not counted;
#: no fraction of the anchor in between fits both ends, and neither does any affine term in
#: the label size.  Whatever PowerPoint is really doing is between them.
#:
#: The height it halves is the frame's less the furniture: a **bottom legend** and a
#: **title** each moved the cap by their own band, over three frame heights each, while a
#: *side* legend did not move it at all.  The legend probes landed 0.19 pt out for a reason
#: of their own, and **that reason is now measured**: the Arial 10 pt band those three
#: frames read as 23.02, 23.02 and 23.07 pt is 23.063 pt, one row of
#: :func:`legend_row_pitch` plus :data:`LEGEND_BAND_PAD_PT`.  The old ``LEGEND_BAND_LINES``
#: took 2 line boxes there, 22.34, which is where the 0.19 came from; the guess this
#: docstring used to make -- that the band was 2 *pitches*, 23.00 -- lands within 0.06 pt
#: of the truth for the wrong reason, since the row is the line **box** and Arial's line
#: gap is not in it.
ROTATED_LABEL_HEADROOM_PT = 6.25

#: What PowerPoint puts at the cut, as its own text object whose pen starts exactly where
#: the kept text ends: one U+2026, not three dots.  It is **inside** the allowance the
#: prefix is measured against -- ``f36``/``f37`` cut three characters earlier than a rule
#: on the prefix alone would -- and **outside** the band, hanging past the anchor towards
#: the axis, which is why the reserve follows the kept prefix and not the drawn string.
LABEL_ELLIPSIS = "…"

#: Where the rotated baseline's far end lands, relative to the centre of its band on the
#: category axis: this far right, and :func:`rotated_label_anchor` below.  Measured on six
#: probes, spread under 0.15 pt.
#:
#: **The drop is not the flat 12.7 pt this used to carry.**  That was fitted at 10 pt and
#: is the same shape of error the flat band constant was: reading the pen straight off the
#: export puts it at 7.12 pt for a 6 pt label and 30.02 for a 24 pt one.  It is now the
#: same anchor the band is built from, so a label cannot be drawn anywhere but in the space
#: reserved for it, and at 10 pt it is 12.79 against the 12.7 it replaces.
ROTATED_LABEL_OFFSET_X_PT = 2.0

#: Chart kinds laid out around a centre rather than on a pair of axes.
POLAR_CHART_KINDS = frozenset({"pieChart", "doughnutChart", "radarChart", "ofPieChart"})

#: The polar kinds that are a web of spokes rather than a ring of slices.
RADAR_CHART_KINDS = frozenset({"radarChart"})

#: A pie whose small points are pulled into a second plot beside it.
OF_PIE_CHART_KINDS = frozenset({"ofPieChart"})

#: Filled to the zero line rather than stroked through the points.
AREA_CHART_KINDS = frozenset({"areaChart"})

#: The Cartesian kinds with **two value axes and no category axis**.  A bubble is a
#: scatter with a third dimension and shares every one of its measurements -- the plot
#: rectangle, the unanchored axes, the coarse x axis -- so it goes through the same code.
SCATTER_CHART_KINDS = frozenset({"scatterChart", "bubbleChart"})

#: The scatter kind that draws a disc per point rather than a line through them.
BUBBLE_CHART_KINDS = frozenset({"bubbleChart"})

#: A stock chart **is a line chart**, and that is a measurement rather than a reading of
#: the schema: a ``c:stockChart`` with no ``c:hiLowLines`` and no ``c:upDownBars`` came
#: back from PowerPoint as one 1.5 pt polyline per series with the ordinary 6 pt marker
#: cycle on it -- byte-identical in shape to what ``lineChart`` already draws.  The lines
#: a real stock chart lacks are suppressed by the *file*, which writes
#: ``<a:ln><a:noFill/></a:ln>`` on each series; nothing in the renderer hides them.
STOCK_CHART_KINDS = frozenset({"stockChart"})

#: The surface, both spellings of which flatten to this one.  It is a 3-D scene whatever
#: its name says; see :data:`~pptx2svg.parse.chart.SURFACE_CHART_KINDS`.
SURFACE_CHART_KINDS = frozenset({"surfaceChart"})

#: **How many colours a surface's band ramp behaves as if it needed**, beyond the bands
#: it has.  The band colours are the per-point accent cycle
#: (:meth:`ChartBuilder._cycle_accent`) exactly -- plain accents inside one cycle of six,
#: the whole first cycle *darkened* once a second is needed -- but the cycle turns two
#: bands early: four bands draw the plain accents and **five** draw the darkened ones,
#: where four and five *points* both draw plain.
#:
#: Bracketed on the band legend's own vector swatches, which are the fills exactly: 3 and
#: 4 bands come back ``#4472C4 #ED7D31 #A5A5A5 #FFC000``, 5 and 6 come back
#: ``#3B64AD #D26E2A #929292 #E2AA00 #5089BC #62993E``, 7 adds a light ``#8FA2D4`` after
#: those six, and 9 and 10 bands carry the same six darkened followed by light accent1,
#: accent2, accent3 and accent4.  So the *cycle* is the point ramp's and only its
#: threshold moves, which is what this constant is and all it is: what the two extra
#: colours are for is not identified.
BAND_COLOR_CYCLE_SLACK = 2

#: The group elements a **combo** chart may be built out of, **in the order PowerPoint
#: paints them**.  A chart whose groups are all in this tuple is drawn as one picture;
#: anything else falls back to drawing the first group alone.
#:
#: **The order is a precedence by type and not the document order**, which is the one way
#: round it is easy to get backwards.  Measured on ``combo-order``: three pairs --
#: bar/line, bar/area and line/area -- were each authored twice with the two
#: ``c:*Chart`` elements swapped, and the six exports come out in *three* distinct
#: pictures, not six.  The area is under the bars in both spellings of bar+area, the line
#: is over the bars in both spellings of bar+line, and over the area in both of
#: line+area.  Two groups of the *same* type keep their document order (``g-bar-bar``
#: draws the first one first), which is what makes this a stable sort rather than a
#: reordering.
#:
#: The colours do **not** follow this order: a line group written first still takes
#: accent1 and the bar group after it accent2, so ``c:idx`` numbers the series and the
#: precedence only decides who covers whom.
COMBO_CHART_KINDS = ("areaChart", "barChart", "lineChart")

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

# -- bubbleChart ----------------------------------------------------------------------
#
# Thirty probe charts across three decks in the usual 220.4724 x 181.1024 pt frame,
# exported by PowerPoint 16.106 and read back as exact path vertices.  Every circle came
# back axis-aligned and square to 0.001 pt, so the bounding box *is* the diameter.

#: The region the largest bubble is sized against: the **frame inset by 5 pt on every
#: side**, minus the title band and the legend band, and nothing else.  It is emphatically
#: not the plot rectangle: ``font14`` and ``font8`` move all four plot edges and draw the
#: **same** 39.485 pt bubble, while a right legend (plot 137.805 x 145.035) draws 37.511
#: and a top or bottom one (185.729 x 120.952) draws 33.928.
#:
#: The 5 pt is solved rather than guessed.  The legend band is 24.083 pt and shrinks the
#: diameter by the ratio 33.928/39.485, so the height it eats into is
#: ``24.083 / (1 - 33.928/39.485) = 171.12`` -- the 181.102 pt frame less **9.98**.  The
#: side-legend probe then falls out with no further fitting: its reserve is the 47.924 pt
#: the plot gives up, and ``(210.472 - 47.924) x 0.2308`` is 37.513 against 37.511 drawn.
BUBBLE_REGION_INSET_PT = 5.0

#: ``c:bubbleScale`` does **not** scale the diameter.  Nine scales from 1 to 300 on an
#: identical chart give ``D = M * s / (s + 1000/3)`` where ``M`` is the short side of the
#: region above -- a soft clamp, linear in *s* while small and approaching the region's
#: own width as *s* grows, so a bubble can never fill more than the region:
#:
#: ===== ========= =========
#: scale PowerPoint predicted
#: ===== ========= =========
#:   1     0.512     0.5117
#:  10     4.983     4.9836
#:  25    11.937    11.9374
#:  50    22.318    22.3177
#:  75    31.427    31.4269
#: 100    39.485    39.4851
#: 150    53.101    53.1007
#: 200    64.163    64.1634
#: 300    81.048    81.0485
#: ===== ========= =========
#:
#: Worst residual **0.0005 pt**.  A linear reading is refuted at both ends: it predicts
#: 19.74 at 50 where PowerPoint drew 22.318, and 78.97 at 200 where it drew 64.163.
BUBBLE_SCALE_HALF = 1000.0 / 3.0

#: ``c:bubbleScale`` when absent, in percent.
DEFAULT_BUBBLE_SCALE = 100.0

#: How a size becomes a diameter, relative to the largest size **across every series**.
#: ``area`` (the default) is area-proportional -- sizes 1, 4, 9 drew 13.162, 26.323 and
#: 39.485, exactly 1:2:3 -- and ``w`` is diameter-proportional: the same sizes with
#: ``<c:sizeRepresents val="w"/>`` drew 4.387, 17.549 and 39.485, exactly 1:4:9.  The
#: largest is 39.485 in both, and in four probes whose size *distribution* differs
#: (1,2,3 / 1,4,9 / 5,5,5 / 1,2,100), so the reference is the maximum and not the sum.
#: A two-series probe settles that the maximum is **global**: a series topping out at 9
#: drew 27.920 beside one topping out at 18, which is 39.485 x sqrt(9/18).
DEFAULT_SIZE_REPRESENTS = "area"

#: Where a bubble's data label sits, measured from the bubble's *centre* as
#: ``radius + gap``.  ``l``/``r`` put the label box's near edge 8.494 pt out at 10 pt --
#: three bubbles of radius 6.581, 13.162 and 19.742 gave 8.504, 8.484 and 8.494 once each
#: digit's own side bearing is taken out -- and ``t``/``b`` put the line box 7.25 and
#: 6.76 pt out.  These are **not** the scatter's 6.0 and 4.85: a bubble's label stands
#: about 2.4 pt further off its mark than a marker's does.  One font size only, so
#: whether they are points or ems is unmeasured; they are carried as points because the
#: line chart's legend key turned out that way.
BUBBLE_LABEL_SIDE_GAP_PT = 8.494
BUBBLE_LABEL_EDGE_GAP_PT = 7.0

# **A bubble chart's value axis clears the bubbles, not their centres.**  There is no
# constant for it -- the rule is a linear equation and lives in :func:`_clearance_extent`
# and :meth:`ChartBuilder._build_scatter` -- but ``tools/make_bubble_probe.py``'s 40 slides
# are what settled it and the three things they refuted belong beside the sizes above:
#
# * the domain follows the **ink**.  One chart's data at nine ``c:bubbleScale`` values
#   draws 0..10 by 1 at 10, 25 and 50 and 0..12 by 2 at 75, 100 and above.  Nothing that
#   reads the values can move an axis when no value moved.
# * the room is made for the **chart's largest** radius wherever it sits, the same
#   reference :data:`DEFAULT_SIZE_REPRESENTS` takes -- ``d-inner`` and ``d-outer`` draw the
#   same axis with the big bubble in the middle and on the maximum.
# * both axes measure that radius against the plot's **height**.  38 of the 40 slides
#   agree that way and 25 the other; a square plot cannot tell them apart and these are
#   437 x 224.
#
# 38 of 40 exactly, units included.  The two misses are at ``bubbleScale`` 250 and 300 and
# are named in ROADMAP.md 3.2b.

# -- ofPieChart -----------------------------------------------------------------------
#
# Twenty-four probe charts across two decks.  The two plots are packed across the **same
# polar region a pie computes** -- the main pie's left edge and the second plot's right
# edge land on the region's own edges in every probe, and both are centred on its middle
# row to 0.001 pt.

#: ``c:splitType`` when absent, and what ``auto`` means: the **last ceil(n/3) points**
#: move to the second plot.  Measured at n = 3, 4, 6, 7 and 8, which moved 1, 2, 2, 3 and
#: 3 -- ``round(n/3)`` is refuted by n = 4 (it predicts 1) and n = 7 (it predicts 2).
DEFAULT_OF_PIE_SPLIT = "auto"
OF_PIE_AUTO_DIVISOR = 3

#: ``c:secondPieSize`` when absent, in percent of the first plot's radius.  Measured:
#: 33.079 / 44.105 is 0.75 exactly, and probes at 25, 50 and 100 reproduce their own
#: ratios to 0.001.
DEFAULT_SECOND_PIE_SIZE = 75.0

#: ``c:gapWidth`` on an ofPie, in percent of the **first plot's radius**.  With the region
#: width ``W``, ``r = W / (2 + 2s + g/100)`` for the pie form, where ``s`` is
#: ``secondPieSize`` as a fraction: 198.472 / 4.5 = 44.105 drawn 44.105; /4 = 49.618
#: drawn 49.618; /5 = 39.694 drawn 39.694; /6.5 = 30.534 drawn 30.534 at ``gapWidth=300``;
#: /3.5 = 56.706 drawn 56.706 at ``gapWidth=0`` and again at ``secondPieSize=25``.
DEFAULT_OF_PIE_GAP_WIDTH = 100.0

#: The bar form packs differently, and the divisor is **not** the pie's.  Its bar is
#: ``s*r`` wide and ``2*s*r`` tall -- the same vertical extent a second pie of that size
#: would have -- and the gap between the pie and the bar is **half** what it is between
#: two pies: ``r = W / (2 + s + g/200)``.
#:
#: **The ``/200`` is a fitted slope now.**  It used to be the natural reading of a single
#: ``gapWidth=100`` observation and said so here; ``ofpie-pack``'s 23 slides sweep
#: ``gapWidth`` over 0, 25, 50, 100, 150, 200 and 300 and ``secondPieSize`` over 25, 50,
#: 75, 100 and 125, with five cross terms and six pie-form controls, and **every one comes
#: back to 0.004 pt on the radius** -- one part in 25 000 of the divisor.  Each slide reads
#: the radius twice, because the second plot is an exact rectangle ``s*r`` wide and
#: ``2*s*r`` tall, and the two agree.  ``chart-gallery`` slide 9 was recorded in
#: ROADMAP.md 3.2a as a second reading that disagreed; it is not about this constant at
#: all -- see :meth:`ChartBuilder._of_pie_geometry`.
OF_PIE_BAR_GAP_DIVISOR = 2.0

#: The main pie is rotated so the aggregated slice is **centred at three o'clock**,
#: pointing at the second plot.  Five probes: its slice runs 72..108, 27..153, 54..126,
#: 0..180 and 54..126 degrees clockwise from twelve, every one of them centred on 90.  The
#: second plot starts at the same angle the first one does.
OF_PIE_OTHER_ANGLE = 90.0

# -- stockChart -----------------------------------------------------------------------
#
# Twelve probe charts.  The plot rectangle, the axis, the bands and the legend key are a
# line chart's in every one of them.

#: ``c:upDownBars/c:gapWidth`` when absent, in percent of one bar's width.  Measured: the
#: bar came out 14.646 pt on a 36.612 pt band, which is ``band / (1 + 150/100)``; probes
#: at 50 and 300 gave 24.410 and 9.154 against 24.408 and 9.153 predicted.
DEFAULT_UP_DOWN_GAP_WIDTH = 150.0

#: The default fills of ``c:upBars`` and ``c:downBars``.  Read off the export: #F9F9F9 and
#: #3F3F3F, with a black 0.5 pt outline on both.  They are **not** theme accents, which is
#: what the brief for this work expected, and they are measured on one theme only -- the
#: Office scheme, whose ``lt1`` is white and ``dk1`` black -- so whether they are literal
#: or derived from those two is unknown.
DEFAULT_UP_BAR_FILL = "#F9F9F9"
DEFAULT_DOWN_BAR_FILL = "#3F3F3F"

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

#: **Past six, the per-point accent cycle stops being the plain accents.**  Found on the
#: ofPie probes, which always need one colour more than they have points, and it applies
#: to any chart that colours by point: a seven-slice chart came back with accent1..accent6
#: *darkened* and the seventh a light accent1, and a nine-slice one repeated exactly the
#: same darkened six and then three light ones.  So the variation is per **cycle** of six
#: and not a function of the count.
#:
#: The two factors are exact.  Applied to the linear-light value of each channel --
#: ``L * 0.76`` for the first cycle and ``L + 0.23 * (1 - L)`` for the second --
#: they reproduce **all 27 measured channels to the byte**; 0.75 and 0.25, the round
#: numbers either side, are off by up to 1 and 5 respectively.  The conversion matters as
#: much as the factor: the same modulation in HLS on sRGB, which is what
#: :func:`resolve.color._apply_luminance` does for DrawingML's own ``lumMod``, puts
#: accent1's blue channel at 150 against the 173 PowerPoint drew.  That is a real defect in
#: the general colour transform and it is **not** fixed here -- changing it moves every
#: deck in the corpus -- so this ramp carries its own conversion and says why.
#:
#: A third cycle is **not measured**: no probe had more than twelve points.  It repeats the
#: second's tint, which is a guess and is marked as one.
VARY_COLOR_CYCLE_SHADE = 0.76
VARY_COLOR_CYCLE_TINT = 0.23

#: The most major intervals a value axis is ever divided into, and with it the base unit:
#: the major unit is the finest 1-2-5 step that fits the padded range into this many.
#:
#: **101 readings over six probe decks**, and the pattern is exact rather than fitted:
#: 0.1..0.11 draws 0.098..0.112 by 0.002, 1.0..1.1 draws 0.98..1.12 by 0.02, 100..110
#: draws 98..112 by 2.  The count is bracketed on both sides -- eleven intervals of the
#: padded range are accepted (1.04..1.13 comes back 1.03..1.14 by 0.01) and twelve are
#: refused (1.03..1.13 comes back 1.02..1.14 by 0.02) -- so the divisor is ten and the
#: extent that grows out of it may hold eleven.
#:
#: This replaces "the power of ten below the span, halved under twice it", which
#: reproduced 29 of the 94 scatter readings where this reproduces 75 exactly and every one
#: of the 94 extents; the other 19 are this unit coarsened by
#: :func:`side_axis_intervals`, every one of them on a short frame.  The five observations
#: the old ratio was fitted to are *coarsened* results: a bar chart of 0..1842 on a 145 pt
#: plot draws 0..2000 by **200**, and only on the corpus deck's own 112-150 pt frames does
#: it draw by 500.
AXIS_MAX_INTERVALS = 10

#: Headroom added at each end of the data range before the unit is chosen, as a fraction
#: of the range -- and **clamped at zero** on an axis anchored there.
#:
#: Bracketed from both sides.  Anchored 0.3..4.8 takes unit 1 and 0.3..4.76 takes 0.5,
#: which puts it in (4.17%, 5.04%] of the maximum; unanchored 1.0..1.092 takes 0.02 and
#: 1.0..1.09 takes 0.01, which puts it in (4.35%, 5.56%] of the range.  Five per cent is
#: the only round number in the intersection.  On an **anchored** axis it is five per cent
#: of the maximum and not of the data range: 3.0..4.9 and 3.5..4.9 both take unit 1, which
#: they could not do if the floor entered the range.
AXIS_HEADROOM = 0.05

#: What a value axis up the **side** reserves before its first interval, as line boxes of
#: its own tick labels: one at each end for the outermost labels' own height.
#:
#: See :func:`side_axis_intervals` for the measurement.
AXIS_END_LABEL_LINES = 2

#: What that axis reserves on top of those two line boxes, in points, independent of the
#: label size and of the typeface.  It is the frame's own top and bottom insets --
#: ``2 * EDGE_INSET_PT`` -- and the measurement says so to a tenth of a point.
#:
#: Solved from ten transition scans over five label sizes: with the rung at the label's
#: line pitch, every font puts it in (21.3, 22.7] and their intersection is
#: **(21.92, 22.03]**.  See :func:`side_axis_intervals`.
#:
#: **The Arial reading narrows that, and narrows it past this number.**  Every one of the
#: ten scans is Aptos, whose line gap is zero, so all ten bracket the reserve against the
#: same rung either way; the one Arial reading -- ten intervals at 165 pt of frame and nine
#: at 160 -- needs the reserve above 22.012, which intersects the Aptos bracket at
#: (22.012, 22.03] and leaves ``2 * EDGE_INSET_PT`` 0.012 pt outside it.  **Not moved**:
#: one reading against a constant that has a meaning, and every other reading in the sweep
#: is indifferent between 22.00 and 22.02.  An Arial scan at 161 to 164 pt settles whether
#: the reserve really is 22.02 or whether the missing hundredths are somewhere else.
AXIS_EDGE_RESERVE_PT = 2 * EDGE_INSET_PT


#: How far a 3-D scene's depth reaches across the drawing, per unit of the scene's own
#: width, per unit of ``c:depthPercent``, per unit of the sine of the rotation.  It is the
#: whole of the camera's projection, and it is one number.
#:
#: ``c:rAngAx="1"`` -- right-angle axes -- keeps the front face a true rectangle and draws
#: the depth as a fixed offset, so the projection is a pair of lengths rather than a
#: matrix: the depth lands at ``(this * depth * sin(rotY), this * depth * sin(rotX))`` of
#: the scene's width, with the horizontal component untouched by ``rotX`` and the vertical
#: one untouched by ``rotY`` except for its sign.  That separation is measured, not
#: assumed: a genuine yaw-then-pitch rotation, whose two components mix, misses the same
#: readings by 3.0 pt rms where this misses by 0.26.
#:
#: **One constant serves both axes.**  Fitted independently they come out 0.2041 and
#: 0.2083 -- a vertical and a horizontal reading of the same foreshortening -- and holding
#: them equal costs nothing.
#:
#: **What the number is, is the category count**, and nothing uses it any more: see
#: :func:`three_d_scene_depth`.  Every deck it was fitted on draws five categories,
#: ``1 / 5`` is 0.2, and the drawn depth is a bar's own width -- so the "drawn at a fifth
#: of its nominal length" this recorded as unidentified was one fifth of the plot per
#: category, and the foreshortening is not a constant at all but ``sin`` exactly.  A
#: ``line3DChart`` and an ``area3DChart`` kept it while their *aspect* was still the
#: ladder fitted against it; both are now re-derived on ``1 / categories`` over seven
#: category counts (``view3d-cat``, ``view3d-count``, ``view3d-band``), so the last
#: caller is gone and this is the name of a quantity rather than a divisor.
#:
#: Measured on 131 readings of the four ``view3d-*`` probe decks: seven frames, seven
#: depths from 20% to 2000%, nine pitches from -45 to 90 degrees, seven yaws, five stated
#: heights, five series counts and five gap depths.  Every one of them is the value axis'
#: drawn length read from its own tick labels' centres, which PowerPoint leaves vector at
#: 300 dpi beside the raster it draws the scene as.  See ROADMAP.md 3.4.
VIEW_3D_DEPTH_PROJECTION = 0.2030

#: What the scene keeps clear of its region beyond the depth, per unit of its own width.
#:
#: It survives at zero depth *and* zero rotation -- ``<c:view3D/>`` with nothing in it
#: draws a scene 0.8% of its width shorter than the region it is given -- so it is neither
#: the depth nor the projection.  A floor slab's own thickness and a wall's own edge are
#: both the right size to be it; nothing measured here separates them, and nothing here
#: needs to.  Split evenly above and below the face, which fits the drawn placement to
#: 0.98 pt rms where putting it all on one side costs 1.28.
VIEW_3D_SCENE_MARGIN = 0.0079

#: What a ``bar3DChart``'s depth used to be divided into beyond its series, and what that
#: divisor turned out to **be**: ``c:gapWidth``, as a fraction.
#:
#: The reading it records is not refuted.  Adding series really does make the scene
#: shallower -- one to five on a 195 pt frame reserved 0.0600, 0.0459, 0.0373, 0.0321 and
#: 0.0283 of its width -- and the divisor really does behave as ``series + 1.5``.  What
#: was wrong is *why*: the bar is as deep as it is **wide** (see
#: :func:`three_d_scene_depth`), so anything that narrows it makes the scene shallower by
#: the same factor, and the bar's width is the band over ``slots + gapWidth/100``.  The
#: ``1.5`` is therefore ``c:gapWidth``'s own default and not ``c:gapDepth``'s, which is
#: the question this constant's docstring recorded as unsettled and which the
#: ``view3d-mesh`` sweep of ``c:gapWidth`` settles: at 0, 50, 150 and 300% the drawn value
#: axis is 127.20, 146.16, 166.08 and 179.76 pt where the old form predicts one number
#: four times.
#:
#: Kept as the name of the default it stands for, and no longer used to divide anything.
VIEW_3D_DEPTH_ROW_GAP = DEFAULT_GAP_WIDTH / 100.0

#: ``c:gapDepth``'s default, in percent.  ECMA-376 and the measurement agree.
DEFAULT_GAP_DEPTH = 150.0

#: What a prism's faces are painted, as a multiple of the series' own fill **per sRGB
#: channel**.  Nothing draws these yet; they are what a 3-D mesh has to paint with.
#:
#: **They are constants, and that is the measurement.**  ROADMAP.md 3.4 read a top face at
#: 0.758 and a right face at 0.632 off gallery slide 13 at one camera -- ``rotX=15
#: rotY=20`` -- and could not say whether those were constants or a cosine evaluated at
#: fifteen degrees.  ``view3d-colour`` sweeps the camera and reads the faces off each
#: slide's own raster: **twelve pitches from 0 to 90 degrees, fifteen yaws from 0 to 315
#: and eight diagonals down to -45 of pitch draw every face the same byte**.  PowerPoint
#: draws a ``4472C4`` prism's top face ``#345695`` and its right face ``#2A487E`` on every
#: one of the forty-two slides that shows that face -- a degenerate camera hides one, it
#: never recolours it -- at every depth, every stated height and one, two and four series.
#: The faces change size with the camera and never colour.  (The factors reproduce those
#: hexes to a level rather than to the byte; one level is what the export's own rounding
#: moves a flat fill by, and no single factor does better across all four fills.)
#:
#: So the shading is fixed to the *box* and not to the camera -- which refutes, for
#: PowerPoint, the model the reference renderer `@silurus/ooxml` uses.  Its
#: ``meshMaterialFactor`` is a clamped affine Lambert term against a hand-chosen light,
#: evaluated on the **camera-space** normal, so its factors necessarily move with the
#: rotation; PowerPoint's do not move at all.  Its *shape* -- one flat factor per face,
#: multiplying the sRGB triple -- is right, and its numbers are not ours.
#:
#: **The multiply is in sRGB, not in linear light**, which is worth stating because this
#: project has the opposite finding recorded next door: ROADMAP.md 3.2 measured
#: PowerPoint's chart accent cycle modulating the **linear-light** value of each channel,
#: and records that doing the same in HLS on sRGB -- which `resolve/color`'s ``lumMod``
#: still does -- puts accent1's blue at 150 against the 173 PowerPoint drew.  Fitted over
#: four fills chosen for spread -- ``4472C4``, ``ED7D31``,
#: ``FF3300`` and ``103070``, twelve channels a face -- a per-channel sRGB multiply lands
#: within **1.3 levels** on every channel of every face, which is the one level the
#: export's own rounding moves a flat fill by; scaling the linear light and converting
#: back misses by up to **5.8**, most of it on the dark channels a linear
#: scaling's additive sRGB offset moves hardest.  Two different colour spaces for two
#: different operations is the finding, not an inconsistency to resolve.
#:
#: The faces, and which is which, are told apart by where they are drawn rather than by
#: their value: a top face's pixels sit above the front face's, a side face's beside them.
#: ``left`` and ``bottom`` are **one** value and not two -- a prism yawed past a half turn
#: shows its left face and one pitched from below shows its underside, and both come back
#: ``#1A2C4E`` from ``#4472C4`` -- which is the one over-determination in the set.
#:
#: Read as lighting, the four values are ``ambient + diffuse * max(0, n . light)`` with an
#: ambient of 0.395, a diffuse of 0.747 and a light at **22 degrees right of the viewer and
#: 29 degrees above** -- the left and bottom faces being the ones the light misses, which
#: is why they share the ambient exactly.  That reading used to be *consistent* rather than
#: *confirmed*, three parameters against four faces with ``left == bottom`` its only check,
#: and it asked for "a curved extrusion, which would sweep the whole cosine".
#:
#: **A ``line3DChart``'s ribbon is that experiment, and it confirms the model.**  A ribbon
#: is faceted per segment, and each segment's top face is a different normal in the same
#: scene: the five-category probe's four segments come back ``0.6201, 0.7113, 0.8018,
#: 0.8088`` where :func:`three_d_lambert` on each segment's own slope says ``0.6231,
#: 0.7144, 0.8040, 0.8126`` -- four normals spanning a quarter of the whole range, every
#: one inside 0.5%, with a two-category probe's single facet at 0.7277 against 0.7267 and
#: the same probe pitched to 45 degrees at 0.6867 and 0.7908 against 0.691 and 0.7874.
#: So the four constants below are four samples of :data:`VIEW_3D_LIGHT`, and the light
#: is what a sloped face is shaded with.  See ROADMAP.md 3.4.
VIEW_3D_FACE_SHADES = {
    "front": 1.0,
    "top": 0.7587,
    "right": 0.6364,
    "left": 0.3947,
    "bottom": 0.3947,
}

#: The light :data:`VIEW_3D_FACE_SHADES` is four samples of, as a unit vector in the
#: scene's own frame -- ``x`` right, ``y`` up, ``z`` towards the viewer.
#:
#: Solved from the three faces the light reaches: ``front`` fixes ``z``, ``top`` fixes
#: ``y`` and ``right`` fixes ``x``, each as ``(shade - ambient) / diffuse``, and the
#: diffuse below is what makes the three a unit vector.  That is 22 degrees right of the
#: viewer and 29 above, which is the reading recorded beside the constants; what makes it
#: a *model* rather than a restatement is the ribbon sweep above.
VIEW_3D_LIGHT = (0.3238, 0.4876, 0.8108)

#: The ambient term: what a face the light misses is painted, and the floor of
#: :func:`three_d_lambert`.  It is the ``left`` and ``bottom`` faces' own measurement.
VIEW_3D_AMBIENT = 0.3947

#: The diffuse term, which is ``1 - ambient`` to four places for the front face and is
#: fixed by it: the light is a unit vector, so this is the length of the three solved
#: components.
VIEW_3D_DIFFUSE = 0.7465

#: **A surface is lit by a different light, and only the direction differs.**
#:
#: The prism and the ribbon are lit from 22 degrees right of the viewer and 29 above
#: (:data:`VIEW_3D_LIGHT`); a ``surfaceChart``'s sheet is lit from the **corner** --
#: ``(1, 1, 1) / sqrt(3)``, 45 degrees right, 45 above and 45 in front -- with the same
#: ambient and the same diffuse.  That is not a guess from a round number: it is what a
#: free fit of ``ambient + u*nx + w*ny + q*nz`` to **96 facets** returns, at 0.3946,
#: 0.4345, 0.4339 and 0.4317, whose three light components agree to 0.7 per cent of each
#: other and whose ambient agrees with :data:`VIEW_3D_AMBIENT` to the fourth place.
#:
#: The instrument is the ribbon's, turned on a sheet: every band painted one and the same
#: red so that the mesh is separable from the floor by colour alone, a facet whose normal
#: is known from the drawn geometry, and ``c:hPercent`` and ``c:depthPercent`` as the two
#: levers that tilt it from level to nearly vertical in each of the scene's planes
#: without touching the data.  ``view3d-surflight`` (34 slides) and ``view3d-surflit``
#: (29) are those sweeps.  Read back with the **shared** constants -- ambient 0.3947,
#: diffuse 0.7465 -- every one of the 104 facets is inside **two 8-bit levels**, mean
#: 0.66, and a level is what the export quantises to.  The prism's own light is refuted
#: here by a wide margin: it puts a level sheet at 0.7587 where PowerPoint draws 0.8275.
#:
#: Two things the same sweeps settle:
#:
#: * **It clamps at both ends.**  ``ambient + diffuse`` is 1.1412, and the brightest facet
#:   measured is 1.0000 exactly -- so the product is clamped, not the factor.
#: * **A facet the light misses is a flat 0.4000**, measured on every back-facing sample
#:   and independent of how far it faces away.  ``max(0, n . L)`` puts it at the ambient
#:   0.3947, which is 1.4 levels darker; that is the whole of the disagreement and it is
#:   left where it is rather than given a constant of its own.
VIEW_3D_SURFACE_LIGHT = (3.0**-0.5, 3.0**-0.5, 3.0**-0.5)

#: What a ``c:wireframe`` surface's lattice is stroked at, in EMU.  Read off the wireframe
#: legend probe's own key, whose square carries a stroke of exactly this.
SURFACE_WIREFRAME_WIDTH_EMU = 6350.0

#: The hairline a filled facet is stroked with **in its own colour**, in EMU.
#:
#: Not a line PowerPoint draws: two polygons sharing an edge antialias against the paper
#: rather than against each other, so a mesh of them comes out with a pale seam along
#: every edge where PowerPoint -- which rasterises the whole scene at once -- has none.
#: The stroke closes the seam and moves the silhouette by half of it, which is a fortieth
#: of a point.
SURFACE_SEAM_WIDTH_EMU = 3175.0

#: What a title costs a **side** legend beyond its own line, in points.
#:
#: The legend block is centred in the frame less ``line_height + this``, measured on nine
#: probes at three title sizes and three frame heights: the block's centre sits 9.38,
#: 13.04 and 19.15 pt below the frame's at 8, 14 and 24 pt, which is a straight line in
#: the size with a slope of 1.2213 -- the face's own line height per em -- and this
#: intercept.  It is a constant and not a share of the frame: 120, 250 and 330 pt frames
#: all read 13.04 at 14 pt.  What the 9 pt *is* is not identified.
LEGEND_SIDE_TITLE_GAP_PT = 8.99

#: How thick a ``line3DChart``'s ribbon is, as a fraction of the scene's own **width**.
#:
#: A ribbon is a thin solid rather than a sheet: its front face is the fill exactly, and
#: it hangs *below* the value it plots -- the front face's top edge is the value, read on
#: a three-category probe where the first category's front face is painted at 5.641 value
#: units against the 4.05 it plots plus the 1.586 its row's near offset lifts it.
#:
#: Measured over **72 readings** -- eight cells, three cameras, one to three series --
#: as the front face's own pixel area over the columns it spans, which is a subpixel
#: reading of a strip the export quantises to 0.24 pt: 0.00255 of the scene's width, sd
#: 0.00015, min 0.00205 and max 0.00279.  It tracks the scene and not the frame, the
#: category band or the row: the same cell pitched to 45 degrees draws a scene 0.67 of the
#: width and a ribbon 0.67 as thick.  ``1 / 400`` is inside the spread and is what is
#: shipped; what the quantity *is* is not identified.
VIEW_3D_RIBBON_THICKNESS = 0.0025

#: What the **front** face is lifted by, in 8-bit levels added to every channel alike.
#:
#: The one thing in the scene's colour that does move with the camera, and it is not a
#: scaling: at ``rotX=15 rotY=20`` a ``103070`` prism's front face is ``#103070`` exactly,
#: and at ``rotX=-45`` it is ``#1B3B7B`` -- eleven levels up on all three channels, which
#: no multiplier produces.  It is zero to within the export's own half-level over the
#: whole ordinary camera -- every pitch from 10 to 90 degrees and every yaw from 10 to 150
#: -- and grows only as the scene degenerates: +1 at five degrees of either, +2 at zero
#: pitch or at a yaw of 0 or 180, +3 to +4 at yaws past a half turn, +6 at ``rotX=-15``,
#: +10 at ``rotX=0 rotY=0`` where no other face is drawn at all, and +11 at ``rotX=-45``.
#:
#: **What it is is not identified.**  It is not a white blend (that would move a dark
#: channel more than a bright one, and it does not), it is not a gradient (the face is one
#: flat colour over three hundred thousand pixels) and it is not compression (the rasters
#: are ``FlateDecode``).  The face is the fill at every camera anyone would author, so a
#: mesh can paint it as the fill; this is recorded because it is measured, not because it
#: is needed.
VIEW_3D_FRONT_LIFT_MAX = 11.0

#: What a **radial** axis -- a radar's, running from the centre to the rim -- can hold
#: beyond its whole line boxes, in ems of one.  Its count is
#: ``floor(radius / line_box + this)``, clamped to 1..:data:`AXIS_MAX_INTERVALS`.
#:
#: Measured on forty readings, the N-meter over eight frames: radii of 26.16 to 128.4 pt
#: at 12.207 pt of line box take 3, 3, 4, 5, 6, 7, 8 and 10 intervals, which brackets this
#: to **(0.857, 0.894]**.  Whether it is a fraction of the line box or a flat 10.7 pt is
#: not separated -- every one of those forty is 10 pt -- and nine replicas of the corpus
#: radar at 10 and 12 pt agree with either.
#:
#: **A radial axis is not the side axis seen sideways.**  The side rule on the same radius
#: is out by two to four intervals in both directions, and on the diameter by three; this
#: one has no 22 pt of frame inset in it and no two end labels, which is what the ring
#: deck says and is also what a radar looks like -- its labels stack up one spoke rather
#: than down the side of a plot.
RADIAL_AXIS_SLACK_EM = 0.875

#: The rung of a value axis along the **bottom**, in ems of its label size.
#:
#: A bottom axis is coarser than a side axis for the same data and the same length, and
#: this is why: its rung is four ems where a side axis' is one line box.  Measured by
#: scanning the frame width two points at a time either side of two transitions at 10 pt
#: and sweeping a third at 20 pt, all with :data:`BOTTOM_AXIS_RESERVE_PT` held: the
#: 10 pt scans bracket it to (4.00, 4.02] ems and the 20 pt sweep to (3.78, 4.18].
#:
#: What four ems *is* was not identified.  It is not the label -- "10" is 0.9 em wide and
#: "10.000" 2.5 em, and both cross the rung at the same frame width.
BOTTOM_AXIS_RUNG_EM = 4.0

#: What a bottom axis reserves before its first rung, in points.  Bracketed to (22, 24] by
#: the same scans -- the frame's two insets are 22.0 and sit at the very edge of it -- and
#: 23 is the middle of the bracket rather than a number with a meaning.
BOTTOM_AXIS_RESERVE_PT = 23.0

#: The gap a bottom axis leaves beside a tick label wider than
#: :data:`BOTTOM_AXIS_RUNG_EM`, in ems of the label size.
#:
#: Labels reading eight digits -- 4.3 ems wide as we measure them -- cross the rungs 50 pt
#: of frame later than "10" and "10000" do, and those two cross them together at 0.9 and
#: 2.5 ems.  So the rung is the greater of four ems and the widest label plus this gap.
#:
#: **The wide family does not fit one rung**, which is why this is a fitted number and not
#: a bracket: its N=10 crossing asks for 0.40 ems and its N=5 crossing for 0.93, and no
#: single value produces both.  0.4 is the one that reproduces the most of the sweep -- 30
#: of the 34 readings whose labels are ten characters wide -- and three of the four it
#: misses are where PowerPoint stops choosing a unit at all: 0..9e6 in a 260 pt frame came
#: back labelled 0, 4e6, 8e6 on an axis still ending at 1e7, which is a *skipped* tick
#: rather than a 1-2-5 unit.  Nothing here draws that, and in the 16 pt window where
#: PowerPoint does, ours is one rung finer.  The fourth is the N=10 crossing, 17 pt late.
BOTTOM_AXIS_LABEL_GAP_EM = 0.4

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


# --------------------------------------------------------------------------------------
# Axis scaling
# --------------------------------------------------------------------------------------


def side_axis_intervals(available_pt: float, pitch_pt: float) -> int:
    """How many major intervals a value axis **up the side** of a chart is divided into.

    ``available_pt`` is the chart frame's height less its title and less a legend above or
    below it -- what is left for the plot and its axis labels -- and ``pitch_pt`` is the
    baseline-to-baseline pitch of the face the *tick labels* are drawn in.  The count is

    ``floor((available - 2 * EDGE_INSET_PT) / pitch) - 2``, clamped to 1..10.

    **This is the coarsening stage, and it is one stage rather than two.**  There is no
    separate "step the unit up when the plot is short": the same expression that caps a
    tall chart at :data:`AXIS_MAX_INTERVALS` produces the by-500 axis on a 150 pt frame,
    because the divisor in :func:`nice_axis_scale` is this count and not always ten.

    **It is a function of the frame, not of the plot.**  That is what a decade of
    contradictory readings turned on: four chart types -- scatter, column, line and area --
    sweeping two datasets over eight frames coarsened at *exactly* the same frame heights,
    although their plot rectangles differ by 14 pt because three of them reserve a band for
    category labels and the scatter does not.  Fitted to the drawn plot instead, the 64
    readings have no solution at all: a 44.88 pt plot taking one interval and a 50.64 pt
    plot taking three force a rung under 5.8 pt, which the 74.16 pt plot's five intervals
    then contradict.

    **The rung is one line pitch of the label.**  Ten transition scans -- the frame walked
    two points at a time either side of the height where the drawn unit changes, at 6, 10,
    14, 20 and 28 pt labels -- bracket it to (1.2143, 1.225] ems, and Aptos' own pitch is
    1.2207.  Arial labels cross their transition 4 to 9 pt lower, which is a *face* ratio
    and not a constant one, so this asks the face rather than
    :data:`~pptx2svg.text.measure.DEFAULT_LINE_HEIGHT_RATIO`.

    **Pitch, not line box, and Arial is what says so.**  Every one of those ten scans is
    Aptos, whose ``hhea`` lineGap is zero, so they cannot tell the two apart; Arial's is
    67 units of 2048 and it can.  Read with the line box the Arial readings are not merely
    loose but *contradictory*: PowerPoint steps up between 160 and 165 pt of frame, and
    ``22 + 12 * lineBox`` puts the step at 156.0 to 156.1 for every reserve in the 22 pt
    bracket, which is outside that window.  With the pitch -- 2355 units of 2048, 11.499 pt
    at 10 pt -- the same expression gives 159.91 to 160.02, which is inside it.  That is
    the same lineGap ``_bottom_label_band`` measures independently on its four-rung ladder,
    arrived at from a different chart and a different quantity.

    **One reading still misses, by 0.012 pt.**  A 160 pt frame of 10 pt Arial draws ten
    intervals here and nine in PowerPoint: ``(160 - 22) / 11.499`` is 12.00104, and eleven
    would need the reserve above 22.012.  The two brackets do intersect -- the Aptos scans
    put the reserve in (21.92, 22.03] and this reading needs (22.012, ...), leaving
    (22.012, 22.03] -- so the model is consistent and it is the round
    ``2 * EDGE_INSET_PT`` that the intersection now excludes, by three hundredths of a
    point.  Not moved: that would be a one-reading fit against a constant with a meaning,
    and an Arial scan at 161 to 164 pt settles it properly.  Before the lineGap the rung
    was 0.33 pt short and this reading was out by four points of frame; it is now out by
    one part in a thousand of a rung.

    **The two line boxes and the 22 pt are separately measured.**  With the rung at the
    line box, every font's own scans put the remaining reserve at 21.3 to 22.7 pt and the
    five intersect at (21.92, 22.03] -- two frame insets -- and the 2 is what is left over
    per font once that is fixed, one line box for the label at each end of the axis.

    **Title and legend come off the top.**  A bottom legend on a 150 pt frame takes the
    count from 8 to 6 and a title takes it to 6, which is exactly what subtracting the
    bands ``_plot_rect`` already reserves for them predicts; a legend at the *right* left
    the count at 8, so it is the height the furniture eats and not the furniture itself.
    """
    if not (math.isfinite(available_pt) and math.isfinite(pitch_pt)) or pitch_pt <= 0:
        return AXIS_MAX_INTERVALS
    rungs = (available_pt - AXIS_EDGE_RESERVE_PT) / pitch_pt
    if not math.isfinite(rungs):
        return AXIS_MAX_INTERVALS
    return max(1, min(AXIS_MAX_INTERVALS, math.floor(rungs) - AXIS_END_LABEL_LINES))


def radial_axis_intervals(radial_pt: float, pitch_pt: float) -> int:
    """How many major intervals a **radial** axis -- a radar's rings -- is divided into.

    ``radial_pt`` is the drawn radius, centre to rim, and ``pitch_pt`` the line pitch of
    the ring labels' face.  See :data:`RADIAL_AXIS_SLACK_EM`; the radius is available
    before the scale because a radar's geometry is set by its *category* labels, which is
    why :meth:`ChartBuilder._build_radar` measures it first.

    **Pitch here is carried across from the side axis, not measured here.**  All forty
    ring readings are Aptos, whose ``hhea`` lineGap is zero, so they say nothing about
    which of the two this rung is, and the corpus radar comes out at four rings either
    way -- there is no neutral choice to make, only two readings of the same evidence.
    It is the pitch because a radar's rings are labelled by the same value-axis tick
    labels the side axis counts, and that is where the two were separated: see
    :func:`side_axis_intervals`.  A radar in a face with a line gap would decide it.
    """
    if not (math.isfinite(radial_pt) and math.isfinite(pitch_pt)) or pitch_pt <= 0:
        return AXIS_MAX_INTERVALS
    rungs = radial_pt / pitch_pt + RADIAL_AXIS_SLACK_EM
    if not math.isfinite(rungs):
        return AXIS_MAX_INTERVALS
    return max(1, min(AXIS_MAX_INTERVALS, math.floor(rungs)))


def bottom_axis_intervals(
    available_pt: float, font_size_pt: float, widest_label_pt: float = 0.0
) -> int:
    """How many major intervals a value axis **along the bottom** is divided into.

    The same shape as :func:`side_axis_intervals` with a wider rung and no end-label
    allowance: ``floor((available - 23) / rung)``, clamped to 1..10, where the rung is the
    greater of :data:`BOTTOM_AXIS_RUNG_EM` ems and the widest tick label plus
    :data:`BOTTOM_AXIS_LABEL_GAP_EM`.  ``available_pt`` is the frame's width less a legend
    beside it.

    **It is the width, and only the width.**  Three frame heights over a width sitting on
    a transition drew the same axis, and a horizontal bar chart and a *scatter's x axis*
    sweep the same rungs at the same widths: 0.5..4.5 of scatter x comes back 0..6 by 2 at
    160 and 200 pt of frame, 0..5 by 1 at 240, 300 and 400, and 0..5 by 0.5 at 684.  Those
    six readings are what ``HORIZONTAL_MAX_INTERVALS`` used to approximate with a flat cap
    of four: its probes all sat on 160 to 200 pt frames, where this rule also says three
    or four, and it drew four intervals on a 684 pt frame where PowerPoint draws ten.
    """
    if not (math.isfinite(available_pt) and math.isfinite(font_size_pt)) or font_size_pt <= 0:
        return AXIS_MAX_INTERVALS
    rung = max(
        BOTTOM_AXIS_RUNG_EM * font_size_pt,
        widest_label_pt + BOTTOM_AXIS_LABEL_GAP_EM * font_size_pt,
    )
    if rung <= 0 or not math.isfinite(rung):
        return AXIS_MAX_INTERVALS
    rungs = math.floor((available_pt - BOTTOM_AXIS_RESERVE_PT) / rung)
    return max(1, min(AXIS_MAX_INTERVALS, rungs))


def _clearance_extent(
    low: float, high: float, data_minimum: float, data_maximum: float, clearance: float
) -> tuple[float, float]:
    """Widen ``(low, high)`` until every mark's own *ink* fits inside it.

    ``clearance`` is a drawn radius as a fraction of the axis' own length, so the room a
    mark needs is ``clearance * (high - low)`` -- a quantity that depends on the answer.
    The circle closes in one step because the radius is fixed before the axis is (see
    :data:`BUBBLE_REGION_INSET_PT`), leaving a linear equation per end.

    Which ends are active is not known in advance: a chart whose lowest point sits well
    above its axis' floor needs no room there at all.  So each of the four cases is solved
    in closed form and the **narrowest** consistent one wins, which is the same answer a
    fixed-point iteration converges to and does not need a convergence argument.
    """
    if clearance <= 0.0 or not math.isfinite(clearance):
        return low, high
    clearance = min(clearance, 0.45)
    candidates = [(low, high)]
    if clearance < 1.0:
        candidates.append((low, (data_maximum - clearance * low) / (1.0 - clearance)))
        candidates.append(((data_minimum - clearance * high) / (1.0 - clearance), high))
    span = (data_maximum - data_minimum) / (1.0 - 2.0 * clearance)
    candidates.append((data_minimum - clearance * span, data_maximum + clearance * span))

    best: tuple[float, float] | None = None
    for candidate_low, candidate_high in candidates:
        bottom, top = min(candidate_low, low), max(candidate_high, high)
        width = top - bottom
        if not math.isfinite(width) or width <= 0:
            continue
        room = clearance * width
        # A relative tolerance, because both sides of each comparison are products of the
        # same solved width and land a few ulps either side of equality by construction.
        slack = _EXTENT_SLACK * max(width, 1.0)
        if data_minimum - room < bottom - slack or data_maximum + room > top + slack:
            continue
        if best is None or width < best[1] - best[0]:
            best = (bottom, top)
    return best if best is not None else (low, high)


def nice_axis_scale(
    data_minimum: float,
    data_maximum: float,
    *,
    intervals: int = AXIS_MAX_INTERVALS,
    strict: bool = True,
    anchor_zero: bool = True,
    clearance: float = 0.0,
) -> tuple[float, float, float]:
    """``(minimum, maximum, major_unit)`` for a value axis PowerPoint would draw itself.

    The data range is padded by :data:`AXIS_HEADROOM` at each end, the major unit is the
    finest 1-2-5 step that divides that padded range into no more than *intervals*, and
    the extent is the padded range rounded outwards to whole units.  *intervals* is
    :data:`AXIS_MAX_INTERVALS` for an axis with room for it and less for a short one; see
    :func:`side_axis_intervals` and :func:`bottom_axis_intervals`, which is where the
    frame comes in.

    ===========  ==================  =============
    data         PowerPoint          intervals
    ===========  ==================  =============
    0.1..0.11    0.098..0.112 x .002 7
    1.0..1.1     0.98..1.12 x 0.02   7
    100..110     98..112 x 2         7
    0..5         0..6 x 1            6
    0..9         0..10 x 1           10
    -2..5        -3..6 x 1           9
    0..1842      0..2000 x 200       10
    ===========  ==================  =============

    The last one is the correction this rule carries: on a 145 pt plot PowerPoint draws
    0..1842 by **200**, and the by-500 axis the old rule was fitted to is that same axis
    coarsened by a 150 pt frame.

    ``strict=False`` drops the headroom and rounds the extent from the **data** rather
    than from the padded range.  Two kinds of chart want it, and both are measured:

    * a **radar** -- the same 0..5 data a bar chart takes to 6 stopped at exactly 5 on
      every radar probe, five rings with the outermost passing through the largest point;
    * a **3-D chart** -- ``view3d-meter``'s 0..50 dataset came back 0..50 on all seven
      frames and all three 3-D group elements, never the 0..55 or 0..60 a 5% headroom
      produces, and its 0..96 dataset came back 0..100 where the headroom gives 0..120.
      Fed back through this function, **the padded rule cannot draw what PowerPoint drew
      at any interval count** -- 34 of the two decks' 52 3-D cells have no solution at all --
      while the bare rule solves every one of the 52.  The two ``barChart`` controls on
      the same deck, the same frames and the same export are the other way round: padded
      solves both and bare solves neither.  See ROADMAP.md 3.4.

    The unit is chosen identically either way, from the range this leaves.

    ``anchor_zero=False`` lets the domain leave zero out when the data sits far enough up
    its own range; see :data:`AXIS_ZERO_ANCHOR_RATIO`.  Only a **scatter** passes it, and
    only a scatter has been measured.

    ``clearance`` is a **bubble**'s radius as a fraction of this axis' drawn length, and it
    is the one thing here that pads for ink rather than for numbers: the domain is widened
    until every mark's own circle fits inside it, and the 5% headroom becomes a floor
    rather than the answer.  See :func:`_clearance_extent` and ROADMAP.md 3.2a; the two
    never *add*, which ``bubble-axis``'s scale sweep settles outright -- at
    ``bubbleScale=50`` the sum rounds a 0..10 axis to 0..12 and PowerPoint draws 0..10.
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

    # The headroom is clamped at zero on an anchored axis: a bar chart of positive data
    # does not get a strip of axis below the bars.  A **radar** has none at all, which is
    # the same measurement as its extent stopping at the data: with 5% added, its own ring
    # sweep needs eleven intervals where ten is the most any axis takes, and its 0..5 data
    # would take unit 1 where PowerPoint draws 0.5.
    headroom = AXIS_HEADROOM * span if strict else 0.0
    padded_low = low if anchored and low >= 0.0 else low - headroom
    padded_high = high if anchored and high <= 0.0 else high + headroom
    padded_low, padded_high = _clearance_extent(
        padded_low, padded_high, data_minimum, data_maximum, clearance
    )

    unit = _nice_unit((padded_high - padded_low) / max(1, intervals))
    # A denormal span underflows the power of ten to zero; a span at the other end
    # overflows the rounding below to infinity.  Neither is a chart anyone drew on
    # purpose, and both used to raise out of the conversion.
    if not math.isfinite(unit) or unit <= 0:
        return 0.0, 1.0, 1.0

    minimum, maximum = _axis_extent(unit, padded_low, padded_high, low, high, strict)
    if not (math.isfinite(minimum) and math.isfinite(maximum) and maximum > minimum):
        return 0.0, 1.0, 1.0
    if clearance > 0.0:
        # Rounding outwards is not the end of it: the room a bubble needs is a fraction of
        # the **drawn** span, which the rounding has just made larger, so a domain that
        # cleared the ink before it was rounded can fail after.  Measured: ``bubble-axis``
        # at ``bubbleScale=150`` solves to 0..10.885, rounds to 0..12, and at 0..12 the
        # bubble sitting on the data's own minimum of 2 hangs 0.08 units below the floor --
        # PowerPoint draws -2..12.  One more unit each way settles every such slide; at 300
        # it takes two.
        for _ in range(AXIS_MAX_INTERVALS):
            room = min(clearance, 0.45) * (maximum - minimum)
            below = data_minimum - room < minimum - _EXTENT_SLACK * max(maximum - minimum, 1.0)
            above = data_maximum + room > maximum + _EXTENT_SLACK * max(maximum - minimum, 1.0)
            if not (below or above):
                break
            if below:
                minimum -= unit
            if above:
                maximum += unit
    return minimum, maximum, unit


def _axis_extent(
    unit: float,
    padded_low: float,
    padded_high: float,
    low: float,
    high: float,
    strict: bool = True,
) -> tuple[float, float]:
    """Round the domain outwards to whole units, from the padded range or from the data.

    The padded range is what PowerPoint rounds, and that is measured rather than assumed:
    0.3..4.9 is already clear of a 0..5 axis and PowerPoint draws 0..**6**, which no rule
    reading the data alone produces.  All 94 extents in the scatter sweep come out of this,
    the unanchored ones included -- 100..104 comes back 98..106 and 2010..2020 comes back
    2005..2025, each a whole unit clear at both ends because 5% of their range carries them
    past one.

    ``strict=False`` rounds from the data instead, which is the radar rule: 0..5 stops at
    5 where a padded 0..5.25 would go to 6.
    """
    ceiling = padded_high if strict else high
    floor = padded_low if strict else low
    # A tolerance, because a quotient that is a whole number on paper is often a hair over
    # it in doubles -- 5.25 / 0.05 is 105.00000000000001 -- and a hair is a whole extra
    # unit of axis once it is rounded outwards.
    maximum = math.ceil(ceiling / unit - _EXTENT_SLACK) * unit
    minimum = math.floor(floor / unit + _EXTENT_SLACK) * unit
    return minimum, maximum


#: How far past a unit boundary a padded end may land and still be taken as on it.  The
#: padded range is a product and a difference of authored decimals, so a boundary it lands
#: on exactly arrives a few ulps either side of one; a tenth of one part in a billion is
#: far wider than that residue and far narrower than any authored number's distance from a
#: boundary.  Relative, because it is applied to the quotient by the unit, which is a small
#: count of intervals.
_EXTENT_SLACK = 1e-9


def _nice_unit(value: float) -> float:
    """The smallest 1-2-5 step at or above *value*, for a finite positive *value*."""
    if not math.isfinite(value) or value <= 0:
        return 0.0
    decade = _decade(value)
    if decade <= 0 or not math.isfinite(decade):
        return 0.0
    for step in (1.0, 2.0, 5.0):
        candidate = step * decade
        if value <= candidate * (1.0 + _DECADE_SLACK):
            return candidate
    return 10.0 * decade


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


#: How far *below* a power of ten a value may fall and still count as having reached it.
#:
#: Not a fudge factor for the comparison in :func:`_decade` -- that comparison is a float
#: comparison and decides exactly what it is asked -- but a statement about where axis
#: spans come from.  A span is a subtraction of two numbers somebody typed into Excel, and
#: that subtraction is not exact: ``0.24 - 0.14`` is ``0.09999999999999998`` and
#: ``1.13 - 1.03`` is ``0.09999999999999987``.  An axis that steps by 0.005 for data whose
#: span is, to any reader, a tenth is the wrong picture.
#:
#: ``log10``'s rounding used to provide a window of this kind here by accident, which is
#: why the rule has always been the forgiving one; it was just narrower and uneven -- two
#: ulps under a tenth was forgiven and ten ulps under it was not, so of those two
#: subtractions the first got the right axis and the second did not.
#:
#: One part in a trillion is wide enough for any residue a subtraction of authored
#: decimals can leave (those are a few ulps, ~1e-16 relative) and far narrower than a span
#: that misses a decade because it was *written* that way: 9.999999999999 is promoted to
#: the decade of 10, 9.9999999999 is not, and 9.99 is not by nine orders of magnitude.
#:
#: **The width is not measured, and a measurement went looking.**  Thirty scatter probes
#: were built whose unanchored value-axis span falls under a power of ten by a relative
#: 1e-15, 2e-14, 1e-13, 1e-11, 1e-9, 1e-6, 1e-4, 1e-2 and 5e-2 -- a ladder that steps clean
#: over this constant -- and exported through PowerPoint (``tools/make_axis_probe.py``,
#: deck ``axis-decade``).  **The shortfall changed nothing anywhere.**  Every probe on a
#: given base came back with the same axis as its siblings: 1.02..1.14 by 0.02, 10.2..11.4
#: by 0.2, 1020..1140 by 20.  A span written as a plain 0.095 and one ten ulps under a
#: tenth are drawn identically, so no experiment on this path can bracket a decade
#: boundary: the unit PowerPoint picks is not a function of which decade the span falls in.
#: What it *is* a function of is :func:`nice_axis_scale` and the counts feeding it, which
#: is why the axis those thirty probes drew is now the axis we draw.
#:
#: What is left for this constant is the other end of the same promise, in
#: :func:`_nice_unit`: a target that lands a few ulps *above* a 1-2-5 rung takes that rung
#: rather than the next one up.  A padded range of 5.25 over ten is 0.5250000000000001 on
#: this machine, and a unit of 1 where PowerPoint draws 0.5 is the same wrong picture from
#: the other side.
#:
#: So this number stays what it was, and stays labelled for what it is: an internal
#: promise that two spans a few ulps apart are treated alike, chosen wide enough to cover
#: any authored subtraction and narrow enough to promote nothing a reader would call short.
_DECADE_SLACK = 1e-12


def _power_of_ten(exponent: int) -> float:
    """``10 ** exponent`` as a float, from CPython's decimal parser rather than libm.

    The parser is correctly rounded and is part of the interpreter, so this is the same
    double on every platform; ``pow`` is libm and carries no such promise.  It also gives
    the ends of the range as values instead of exceptions -- ``1e309`` is ``inf`` where
    ``10.0 ** 309`` raises ``OverflowError``, and ``1e-324`` underflows to ``0.0`` --
    which is what lets :func:`_decade` walk off either end and stop.
    """
    return float(f"1e{exponent}")


def _decade(value: float) -> float:
    """The largest power of ten *value* reaches, for a **finite, positive** *value*.

    This is ``10.0 ** math.floor(math.log10(value))`` with ``log10``'s answer taken as a
    guess and then checked, because that expression has no small errors -- it is exact or
    it is out by a factor of ten.  ``log10(100.0)`` is exactly 2.0 on every mainstream
    libm, since the true value is representable and correct rounding therefore demands it;
    a merely *faithful* implementation may return ``1.9999999999999998``, and ``floor``
    turns that one ulp into a unit of 10 where PowerPoint draws 100.  One measured slide
    of the corpus re-lays out on it.  See ROADMAP.md section 0.2.

    So the guess is corrected against the input.  Comparing a value against a power of
    ten is an ordinary float comparison, decided by the same bits everywhere, where
    comparing logarithms is decided by a transcendental.  At most one of the two walks
    below can take a step: whichever direction the guess was wrong in, correcting it
    makes the other direction's test false, so this cannot oscillate.  Both terminate at
    the ends of the double range, where :func:`_power_of_ten` yields ``0.0`` (never above
    *value*) and ``inf`` (never reached).

    What it guarantees: the result depends on ``log10`` only through a guess it verifies,
    so any ``log10`` accurate to better than a whole decade -- every real one -- gives the
    same answer, on every platform.

    What it does not: the ladder it ranks *value* against is the doubles nearest the
    powers of ten, not the powers themselves, which for a negative exponent are not
    representable (``1e-3`` is a hair *above* ten cubed's reciprocal).  Values within
    :data:`_DECADE_SLACK` of a power of ten are deliberately counted as reaching it, so
    the result may be a hair larger than *value* at the bottom of a decade; that is what
    the callers want and what they already got.  At the ends of the range it returns the
    only answers available: ``0.0`` for a value below the smallest positive power of ten
    (a denormal span), and never ``inf``, since no finite value reaches ``1e309``.  A
    zero result is not a usable unit, and :func:`nice_axis_scale` checks for it.
    """
    exponent = math.floor(math.log10(value))
    while not _reaches_decade(value, exponent):
        exponent -= 1
    while _reaches_decade(value, exponent + 1):
        exponent += 1
    return _power_of_ten(exponent)


def _reaches_decade(value: float, exponent: int) -> bool:
    """Whether *value* reaches ``10 ** exponent``, give or take :data:`_DECADE_SLACK`.

    The slack scales the power rather than the value so that it means the same thing in
    every decade, and it is harmless at the ends: ``0.0`` stays ``0.0`` (every positive
    value reaches it) and ``inf`` stays ``inf`` (no finite value does).
    """
    return value >= _power_of_ten(exponent) * (1.0 - _DECADE_SLACK)


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
    #: The face's ``hhea`` line gap at this size, in points; 0.0 for a face whose gap we
    #: have not measured, which is what makes :attr:`pitch` degrade to
    #: :attr:`line_height`.
    gap: float = 0.0

    @property
    def line_height(self) -> float:
        """One line's own box: ascent + descent, with no leading.

        What PowerPoint reserves for a label standing on its own -- the one-line category
        band, the top inset over the highest value label -- and the two are separately
        measured.  :meth:`ChartBuilder._top_inset` is the box to within 0.002 pt over
        ``axis-inset``'s 24 charts, four faces and eight sizes; half of Arial's gap runs
        from 0.10 pt at 6 pt to 0.46 at 28, all of it outside that residual.  The level
        category band is the box plus :data:`CATEGORY_LABEL_GAP_ASCENT`, and the same 24
        charts put it on the box as well -- on the *ascent* for the gap term and on
        ascent + descent for the line, with no room for a third of a point of leading in
        either.  So the gap is between lines and not around them, which is what
        typesetting has always said and is here measured rather than assumed.
        """
        return self.ascent + self.descent

    @property
    def pitch(self) -> float:
        """Baseline to baseline: the line box plus the face's own ``hhea`` line gap.

        The distance PowerPoint advances by for each line *after* the first, and the rung
        its value axis counts in.  Arial is the face that separates this from
        :attr:`line_height` -- 67 units of 2048, 0.328 pt at 10 pt -- and it separates them
        twice over, in the wrapped category band and in the tick rule; see
        :func:`side_axis_intervals` and :meth:`ChartBuilder._bottom_label_band` for the two
        measurements.  Calibri, Aptos, Times New Roman and Courier New all have a zero gap,
        so for them this *is* the line box and every number measured with them stands.

        **Four places in this file were checked and deliberately left on the line box**,
        because a measurement says so rather than because nobody looked:

        * :meth:`ChartBuilder._top_inset`.  ``max(11.0, 5.0 + lineHeight/2)`` reproduces
          all 24 of ``axis-inset``'s readings to 0.002 pt, Arial and its line gap among
          them at eight sizes.  Half that gap is 0.16 pt at 10 pt, eighty times the
          residual, so the inset is measurably the box.
        * :data:`CATEGORY_LABEL_GAP_ASCENT`, the level band's gap term.  Same argument on
          the same 24 readings, and it is the **ascent** rather than the line box or the
          em; see :meth:`ChartBuilder._bottom_label_band`.
        * :data:`TITLE_BAND_LINES`.  It is a *ratio* to the line box, and the only title
          ever measured is 18 pt Arial: 29.70 pt against a 20.109 pt box.  Expressed
          against the pitch the same measurement gives 1.4350 instead of 1.4769 and
          reproduces that title identically, so the two are indistinguishable here and
          differ only for faces nobody measured.  Changing it would move every non-Arial
          title on no evidence at all.
        * The multi-line blocks in ``_place_label``, ``_centred_label`` and
          ``_draw_radar_category_labels``.  Every one of those was measured in Aptos alone,
          which cannot tell the two apart.  The legend's row is no longer among them:
          :data:`LEGEND_ROW_PITCH_RATIO` is measured in Arial as well, and it is the
          **box** -- Arial's 0.327 pt gap at 10 pt is not in the row, which is the same
          separation this docstring asks for and the answer went the other way from the
          value axis' rung.

        The general shape of it: a reserve *around* one line is the box, a step *between*
        two lines is the pitch, and everything still on the box is there because its
        probe deck had no line gap to show.
        """
        return self.ascent + self.descent + self.gap

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
        # `None` is "nobody measured this face's gap", not "the gap is zero", and the two
        # have to behave identically: a face we have not measured must lay out exactly as
        # it did before the column existed.  See `FontMetrics.line_gap`.
        gap=(metrics.line_gap or 0) / units * size,
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


def _legend_tokens(text: str) -> list[str]:
    """``text`` split at every place a legend entry may break.

    A space is a break opportunity and so is the gap between two CJK characters; a Latin
    word is one token and is never split.  The CJK half is measured rather than assumed:
    ``real-financial-report.pptx``'s doughnut legends デジタルソリューション on two lines,
    broken after the eighth character, which is exactly where its band runs out.
    """
    tokens: list[str] = []
    word = ""
    for char in text:
        if char == " " or is_cjk(ord(char)):
            if word:
                tokens.append(word)
                word = ""
            tokens.append(char)
        else:
            word += char
    if word:
        tokens.append(word)
    return tokens


def _legend_wrap(text: str, font: ChartFont, column: float) -> list[str]:
    """One legend entry, broken greedily to fit ``column``.

    Used for the line *count*, which is what opens a side legend's row pitch, and for the
    band the entry is drawn in; the drawing itself is the renderer's own wrap inside that
    width.

    **One probe slide says the column is about 2.4 pt narrower than this gets.**  A
    four-word name on a 360 pt frame has a 110.04 pt column here and three of its words are
    107.60 pt, so this keeps them together where PowerPoint broke 2 + 2.  Counting the
    space that follows the line does not explain it (109.63 still fits) and no single
    fraction of the frame does either: the Japanese doughnut's break needs a column of
    80.02 pt on a 285 pt frame, which is 0.3999 of it, where this slide needs under 0.3933.
    See :data:`LEGEND_SIDE_MAX_FRACTION`; the fraction is left where the one corpus chart
    that wraps measured it.
    """
    if not text:
        return [""]
    if column <= 0 or font.width(text) <= column:
        return [text]
    lines: list[str] = []
    current = ""
    for token in _legend_tokens(text):
        if token == " ":
            if current:
                current += " "
            continue
        candidate = current + token
        if current.strip() and font.width(candidate) > column:
            lines.append(current.rstrip())
            current = token
        else:
            current = candidate
    if current.strip():
        lines.append(current.rstrip())
    return lines or [text]


def rotated_label_anchor(box: "FontBox") -> float:
    """How far below the axis line a turned label's far end sits.

    The fixed part of the drop, and the one term the band and the drawn label must share.
    See :data:`ROTATED_LABEL_HEADROOM_PT` for the readings behind it: it is the label's
    ascent turned through 45 degrees plus the level band's own gap, which reproduces the
    pen positions of eight sizes and two faces to 0.2 pt.
    """
    return box.ascent * _SIN_45 + CATEGORY_LABEL_GAP_EM * box.size


def truncate_label(text: str, font: ChartFont, allowance: float) -> str:
    """One turned category label, cut the way PowerPoint cuts it.

    A label that fits its allowance is left alone; one that does not is cut to the longest
    **prefix whose width plus the ellipsis' own** still fits, and one U+2026 is appended.
    Both halves of that are measured, on the probe pair that separates them: a label of 36
    narrow characters was drawn whole and one of 37 came back at **32** -- three characters
    short of what the prefix alone would have allowed, which is exactly the ellipsis.

    **At least one character survives.**  A 24 pt label on a 110 pt frame has an allowance
    of 2.5 pt and PowerPoint still drew one character and an ellipsis, so the floor is a
    character rather than an empty string.

    The cut is by *width*, not by character count: ``MMMM...IIII`` and ``IIII...MMMM`` are
    the same 32 characters and the same 177.7 pt of label, and PowerPoint kept 10 of the
    first and 21 of the second.
    """
    if not text or font.width(text) <= allowance:
        return text
    ellipsis = font.width(LABEL_ELLIPSIS)
    kept = 0
    for index in range(1, len(text)):
        if font.width(text[:index]) + ellipsis > allowance:
            break
        kept = index
    return text[: max(kept, 1)] + LABEL_ELLIPSIS


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

    **It is a sum of advances with no ``kern`` term, and that is measured rather than
    left over.**  :mod:`pptx2svg.text.measure` applies the face's ``kern`` feature because
    PowerPoint applies it to slide text; the chart engine does not apply it to the text it
    *lays out*, although it does draw that text kerned.  The two are separable in
    PowerPoint's own export and the answer is not close:

    * ``chart-gallery``'s legends put twenty-odd entry names against their key positions,
      and the horizontal legend's arithmetic (:data:`LEGEND_ENTRY_SLACK`) turns each name's
      advance directly into the next key's x.  Measured against the export, the unkerned
      advance lands every one of them within **0.033 pt**; charging the same names their
      ``kern`` moves five entries on slide 9 out by 0.23 to 0.67 pt, ``Plan`` on slides 1
      and 13 by 0.18, and ``Revenue`` on slide 17 by 0.15.  Every entry that moved is one
      with a kern pair in it; the ones without (slide 8's five) do not move at all.
    * The same export *draws* those names kerned.  ``Plan`` on slide 1 is shown as
      ``[ (Pl) 83 (a) -46 (n) ]``, a net 0.37 pt of leftward adjustment against an
      unkerned advance of 19.20 pt and our kerned measure of 18.92 -- so the glyphs are
      kerned while the layout that placed them was not.

    So the chart engine measures the way GDI's ``GetTextExtent`` does and hands the string
    to a shaper afterwards.  Adding ``kern`` here is not a smaller error than leaving it
    out: it is the wrong number for this caller.
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


def three_d_lambert(
    normal: tuple[float, float, float],
    light: tuple[float, float, float] = VIEW_3D_LIGHT,
) -> float:
    """What a face of *normal* is painted, as a multiple of the fill per sRGB channel.

    ``ambient + diffuse * max(0, n . light)`` against :data:`VIEW_3D_LIGHT`, in the
    scene's own frame -- ``x`` right, ``y`` up, ``z`` towards the viewer.  The four
    :data:`VIEW_3D_FACE_SHADES` are this at the four axis-aligned normals, and a
    ``line3DChart``'s ribbon is what says it is a model and not a coincidence: see that
    constant, and :meth:`ChartBuilder._paint_slab`, which shades every sloped face with it.

    *light* is the one thing a ``surfaceChart`` does differently
    (:data:`VIEW_3D_SURFACE_LIGHT`); the ambient, the diffuse and the clamp are shared.
    The product is clamped at 1 as well as at the ambient, which the surface sweep
    measured and the prism's four faces could not have shown -- none of them is bright
    enough to reach it.
    """
    lit = sum(a * b for a, b in zip(normal, light))
    return min(1.0, VIEW_3D_AMBIENT + VIEW_3D_DIFFUSE * max(0.0, lit))


def three_d_scene_shape(
    kind: str,
    series: int = 1,
    grouping: str | None = None,
    *,
    categories: int = 5,
    across: int | None = None,
) -> "tuple[float, float | None] | None":
    """How tall and how deep this group element's scene is, or ``None`` if unmeasured.

    Returns ``(aspect, rows)``: the scene's height over its width **as a multiple of the
    one a ``bar3DChart`` would get**, and how many units of ``depthPercent`` deep it is,
    where ``None`` means the bar's own law -- as deep as one bar is wide, see
    :func:`three_d_scene_depth`.

    *across* is how many category **intervals** the front face spans, which is the
    category count for an axis crossing ``between`` and one less for ``midCat``; it
    defaults to *categories*.

    **The law is one line and it is measured over seven category counts**::

        aspect = floor((across + series) / 2) / categories
        rows   = series                    one row of depth per series

    A ``line3DChart`` and an ``area3DChart`` differ in nothing but *across*: a line's
    points sit in their bands and an area's on the ticks, so the same data gives the line
    one interval more.  ``view3d-band``'s ``crossBetween="between"`` area probes are what
    say that is the axis and not the group element -- a three-category area reads 2 where
    its ``midCat`` twin reads 1, and a five-category one 3 against 2.

    **This replaces the ladder ``0.4 + 0.2 * ceil(m / 2)``, which was that law seen at one
    category count.**  Every deck the ladder was fitted on drew five categories, where the
    depth's old form (:data:`VIEW_3D_DEPTH_PROJECTION` standing in for ``1 / categories``)
    and its corrected one agree exactly; at five the formula above *is* the ladder, which
    is why it reproduced all eight of its readings.  On three categories it does not, and
    that is where the ladder was contradicted: a three-category ``area3DChart`` reads
    0.333, 0.667 and 0.667 at one, two and three series where the ladder says 0.4, 0.6 and
    0.6, and the formula says 1/3, 2/3 and 2/3.

    Measured on ``view3d-cat`` (112 probes), ``view3d-count`` (72) and ``view3d-band``
    (80), with the **category count and the series count varied independently** -- 2, 3,
    4, 5, 6, 7 and 8 categories against one to six series -- and read off the value axis'
    own tick labels rather than the raster, which is the instrument the depth reservation
    was measured with.  Every one of the 42 cells' free fits lands within 0.03 of an
    integer ``k``; fed back as this law the drawn face height comes back within **1.21 pt
    at worst and 0.3 pt typically** over 168 probes at four cameras each.  Two of those
    four are width-bound and two height-bound, which is what separates the aspect from the
    depth.

    What the law *is* is not identified: ``k`` is half the sum of the face's intervals and
    the scene's rows, which is a plausible thing for a layout to want and is not stated
    anywhere.  It is monotone and bounded by its own inputs, though, where the ladder was
    a staircase read at four points -- so there is no longer a count to stop at, and the
    ``VIEW_3D_SHAPE_MAX_SERIES`` of 4 that capped it is gone with the ladder that needed
    it.  Five and six series were read directly, at two category counts each.

    The rest of the group elements:

    * **The depth is one row per series**, where a ``bar3DChart``'s is
      ``(1 + gapDepth) / (series + 1.5)`` of one, and a row is the category band:
      ``series / categories`` of the scene's width reproduces all 42 cells, at one to six
      series and two to eight categories, with ``depthPercent`` scaling it (500% reads
      five times the depth and the same aspect) and ``c:gapDepth`` **not** -- 0 and 500%
      divide the row without moving the scene at all.
    * **Stacked is measured now and still not drawn.**  A stacked ``area3DChart`` reads
      ``across / categories`` of the region's aspect -- 2/3 at three categories, 4/5 at
      five, 1/2 at two, 7/8 at eight, and 3/3 and 5/5 for the ``between`` pair -- on one
      shared row of depth whatever the series count, which fits its four cameras as well
      as any other cell.  What it is *for* is gallery slide 16, whose axis is already
      PowerPoint's without a camera; giving it one changes that axis, so this returns
      ``None`` and the chart keeps its flat rectangle.  See ROADMAP.md 3.4.
    * **A ``surfaceChart`` is a ``line3DChart``'s scene exactly**, both spellings of it,
      and that is measured on its own rather than borrowed: ``view3d-surfshape`` reads the
      face directly off the text -- a surface's categories sit **on** the ticks, so the
      first and last category labels stand on the drawn face's left and right edges while
      the extreme value labels stand on its top and bottom ones -- over 42 cells, two to
      eight categories against one to six series, at two cameras each.  Every cell's
      ``aspect * categories / region`` lands within 0.01 of the integer
      ``floor((across + series) / 2)``, and the depth, read as the raster's overhang past
      that face, comes back ``series / categories`` to 0.3 per cent.  ``across`` follows
      ``c:crossBetween`` here too, which the eight-category pair separates: spelled
      ``between`` it reads 5 where its ``midCat`` twin reads 4.  ``c:depthPercent`` scales
      the depth alone, a stated ``c:hPercent`` replaces the region's aspect, and
      ``c:gapDepth`` moves neither -- 0 and 500% draw the same picture to the digit.

    * **A ``pie3DChart`` has no scene box to measure.**  Its raster is the plot region
      itself at every camera and every frame -- 662.40 x 173.28 pt on a 195 pt frame
      whether the yaw is 0 or 270 -- and only the ink inside it moves.  ``rotY`` and
      ``depthPercent`` change nothing at all: 20, 45, 90, 135, 180 and 270 degrees of yaw
      and depths of 20% and 500% all drew the same 481.92 x 168.24 pt of ink to the
      hundredth of a point.  So a pie's flat rectangle is already PowerPoint's, and there
      is nothing here for a camera to correct.
    """
    if kind == "bar3DChart":
        return 1.0, None
    if kind not in ("line3DChart", "area3DChart", "surfaceChart", "surface3DChart"):
        return None
    if (grouping or "") in ("stacked", "percentStacked"):
        return None
    count = max(categories, 1)
    span = count if across is None else max(across, 1)
    rows = max(series, 1)
    return math.floor((span + rows) / 2) / count, float(rows)


def three_d_camera(
    view: "c.SourceChartView3D | None",
    kind: str,
    *,
    series: int = 1,
    grouping: str | None = None,
) -> "c.SourceChartView3D | None":
    """The ``c:view3D`` this chart's plot rectangle is laid out through, or ``None``.

    ``None`` means "keep the flat rectangle", and there are three ways to get it.

    * :func:`three_d_scene_shape` has nothing measured for this group element or this
      grouping -- a ``pie3DChart``, whose plot rectangle is already PowerPoint's; a
      stacked ``area3DChart``, whose shape is measured and whose *axis* is already right
      without one; or a 2-D group element, which has no scene at all.
    * ``c:rAngAx="0"``, which draws a perspective scene this does not model.
    * ``c:view3D`` absent altogether, which **selects that same perspective scene**
      although ECMA-376 defaults the attribute to 1: the absent probe is identical to
      0.001 pt to ``rotX=15 rotY=20 depthPercent=100 rAngAx=0``.

    A ``line3DChart`` and an ``area3DChart`` pass this gate now and did not before: their
    scenes are not a ``bar3DChart``'s -- on one frame and one view a ``bar3DChart`` drew a
    127.68 pt value axis where a ``line3DChart`` drew 88.56 and an ``area3DChart`` 59.04
    -- but the difference is now measured rather than named.  The **series count** is no
    longer one of the ways to get ``None``: the aspect is a law rather than a ladder now
    (:func:`three_d_scene_shape`), so there is nothing to run off the end of.
    """
    if view is None or view.right_angle_axes is False:
        return None
    if three_d_scene_shape(kind, series, grouping) is None:
        return None
    return view


def three_d_bar_slots(series: int, grouping: str | None) -> int:
    """How many bars stand side by side in one category band.

    The same quantity :meth:`ChartBuilder._draw_bars` divides the band by, named here
    because the **depth** is a function of it: a 3-D bar is as deep as it is wide, so
    anything that narrows the bar makes the scene shallower by the same factor.
    """
    return 1 if (grouping or "") in ("stacked", "percentStacked") else max(series, 1)


def three_d_scene_depth(
    kind: str,
    view: "c.SourceChartView3D",
    *,
    series: int = 1,
    grouping: str | None = None,
    categories: int = 5,
    gap_depth: float | None = None,
    gap_width: float | None = None,
    bar_direction: str | None = None,
    aspect: float = 1.0,
) -> float:
    """How deep the scene is, as a multiple of its own drawn **width**.

    **A 3-D bar is as deep as it is wide**, and that one sentence is the whole of the
    ``bar3DChart`` branch.  ``view3d-mesh`` reads each prism's own right face off the
    raster -- its width is the depth's horizontal projection and its top edge's slope the
    vertical one -- and against the bar's drawn width the ratio comes back **1.000 at
    every camera**: 0.990 at ``rotX=15``, 0.995 at 5 degrees, 1.000 at ``rotY=90``, over a
    pitch sweep of seven, a yaw sweep of seven, four ``c:gapWidth``, four ``c:gapDepth``
    and four series counts.  ``c:depthPercent`` scales it and nothing else does: 20, 50,
    200 and 500% read 0.184, 0.481, 1.965 and 4.993 of the bar's width.

    So the scene is ``(1 + gapDepth)`` of one bar's depth -- the bar's own row plus the
    gap around it -- and the bar's width is the *flat* rule, the category band over
    ``slots + gapWidth/100``.  Written out, as a fraction of the scene's width:

        depthPercent * (1 + gapDepth) / (categories * (slots + gapWidth/100))

    **This replaces two unidentified constants with the quantities they were standing
    in for.**  The reservation this section fitted before read
    ``VIEW_3D_DEPTH_PROJECTION * depthPercent * (1 + gapDepth) / (series + 1.5)``, and
    every deck it was fitted on had **five** categories and the default ``gapWidth`` of
    150%, which is exactly where the two forms agree: ``0.2030`` is ``1 / 5`` to 1.5% --
    ROADMAP.md 3.4's "the depth is drawn at a fifth of its nominal length and no ratio of
    the scene's own proportions produces that fifth" -- and the ``1.5`` is ``gapWidth``'s
    own default rather than ``gapDepth``'s.  The ``c:gapWidth`` sweep is what separates
    them, because the old form predicts nothing at all from it: 0, 50, 150 and 300%
    drew value axes of 127.20, 146.16, 166.08 and 179.76 pt on one frame where it says
    166.08 four times, and this form reproduces all four to **0.15 pt**.

    Two things the clustered sweep only implies, and both follow from "as deep as it is
    wide" rather than being separate rules:

    * **A stacked group takes one slot, not one per series**, so its scene is as deep as a
      single series' -- which is measured: ``m-stack2`` and ``m-stack3`` draw the same
      507.32 by 166.08 pt face, the same 67.44 pt bars and the same 23.04 by 17.28 pt
      depth vector as the one-series probe beside them.
    * **A horizontal bar is as deep as it is *thick***, and its thickness is a share of
      the scene's height rather than of its width, so *aspect* scales the whole thing.
      ``m-bardir`` draws a depth vector of 8.64 by 6.48 pt where the same chart drawn
      upright draws 23.04 by 17.28 -- about a third, which is the region's own aspect.

    **A ``line3DChart``'s and an ``area3DChart``'s row is the category band too**, and
    that is now their law rather than the fitted constant it used to be: one row per
    series, ``series / categories`` of the scene's width, with ``depthPercent`` scaling it
    and ``c:gapDepth`` dividing the row inside it rather than adding to it.  Measured over
    42 cells of ``view3d-cat``, ``view3d-count`` and ``view3d-band`` -- two to eight
    categories against one to six series -- where the free fit's ``depth * categories /
    series`` comes back 1.00 to 1.03 everywhere.  See :func:`three_d_scene_shape`, which
    the aspect had to be re-derived with at the same time: on five categories the old form
    and this one are the same number, which is why neither could be read without the
    other.
    """
    depth = (view.depth_percent if view.depth_percent is not None else 100.0) / 100.0
    gap = max((gap_depth if gap_depth is not None else DEFAULT_GAP_DEPTH) / 100.0, 0.0)
    rows = (
        three_d_scene_shape(kind, series, grouping, categories=categories) or (1.0, None)
    )[1]
    if rows is not None:
        return depth * rows / max(categories, 1)
    width = max(gap_width if gap_width is not None else DEFAULT_GAP_WIDTH, 0.0) / 100.0
    slots = three_d_bar_slots(series, grouping)
    divisor = max(categories, 1) * max(slots + width, MIN_BAR_SLOTS)
    span = aspect if (bar_direction or "col") == "bar" else 1.0
    return depth * span * (1.0 + gap) / divisor


def three_d_plot_rect(
    region: _Rect,
    view: "c.SourceChartView3D",
    *,
    series: int = 1,
    gap_depth: float | None = None,
    kind: str = "bar3DChart",
    grouping: str | None = None,
    categories: int = 5,
    gap_width: float | None = None,
    bar_direction: str | None = None,
    across: int | None = None,
) -> _Rect:
    """Where a 3-D chart's **front face** lands inside the flat plot rectangle.

    A 3-D chart draws a box, and the plot rectangle the rest of this module means -- the
    one the value axis runs up, the categories run along and the marks are drawn in -- is
    that box's front face.  The box is wider and taller than its face by the depth it is
    drawn with, so the face is displaced and shrunk to make room, and that displacement
    and shrink are the whole of what this computes.  PowerPoint rasterises the scene
    itself and leaves the text beside it vector, which is why the face is measurable to a
    quarter of a point although the picture it sits in is a photograph.

    The model, measured on the three ``view3d-*`` probe decks (ROADMAP.md 3.4):

    * The scene is a box ``w`` wide, ``w * hPercent`` high, and as deep as
      :func:`three_d_scene_depth` says -- which for a ``bar3DChart`` is one bar's own
      width, and is the one part of this that is a property of the *plot* rather than of
      the camera.
    * **``c:hPercent`` absent is the region's own aspect.**  ``region.height /
      region.width`` reproduces the seven auto readings to 0.6%.  That is what closes the
      "0.2438 measured against 0.2456" this section recorded as unexplained: the estimate
      of the region was the part that was wrong, not the rule.
    * **The group element scales both.**  *kind* and *grouping* go to
      :func:`three_d_scene_shape`, which multiplies that aspect -- by 0.6 for a
      ``line3DChart`` at one series, by 0.4 for an ``area3DChart`` -- and says how the
      depth divides: one row per series for those two where a ``bar3DChart`` shares one.
      A combination that function has nothing measured for falls back to the bar's scene
      here rather than refusing, because the refusing is :func:`three_d_camera`'s job and
      it happens first: nothing in this module reaches this function without passing it.
    * ``rAngAx="1"`` keeps the face a true rectangle whatever the rotation -- seven
      pitches from 0 to 90 degrees held its height over its width at 0.2444 +- 0.001 while
      both shrank -- and the depth projects to a fixed offset,
      :data:`VIEW_3D_DEPTH_PROJECTION`.
    * The box is scaled **isotropically** to fit the region and centred in it.  That
      ``min`` is where the second branch comes from: the height binds at ordinary pitches
      and the *width* binds at shallow ones, which is why the small-angle readings refused
      to sit on the same curve as the rest.
    * The face sits at the corner the depth leads away from.  ``rotX > 0`` tips the floor
      towards the viewer and takes its room off the top; a ``rotY`` past a half turn
      reverses that, which is measured -- 0 to 135 degrees of yaw all reserved at the top
      and 180, 270 and 340 all reserved at the bottom -- and is not ``cos(rotY)``, which
      would turn at 90.
    """
    scene = three_d_scene(
        region,
        view,
        series=series,
        gap_depth=gap_depth,
        kind=kind,
        grouping=grouping,
        categories=categories,
        gap_width=gap_width,
        bar_direction=bar_direction,
        across=across,
    )
    return region if scene is None else scene.face


def _surface_normal(
    points: list[tuple[float, float, float]],
    depth: float,
    view: tuple[float, float, float],
) -> "tuple[float, float, float] | None":
    """One facet's unit normal in the scene's frame, turned towards the viewer.

    *points* are ``(x, y, z)`` with ``x`` and ``y`` the front face's own screen
    coordinates -- ``y`` down -- and ``z`` the fraction of the scene's depth the point
    stands at.  The scene's frame is ``x`` right, ``y`` **up** and ``z`` towards the
    viewer, so the third coordinate is ``-z * depth`` in it and the second is negated.
    """
    if depth <= 0 or len(points) < 3:
        return None
    a, b, c = (
        (point[0], -point[1], -point[2] * depth) for point in points[:3]
    )
    u = tuple(second - first for first, second in zip(a, b))
    v = tuple(third - first for first, third in zip(a, c))
    normal = (
        u[1] * v[2] - u[2] * v[1],
        u[2] * v[0] - u[0] * v[2],
        u[0] * v[1] - u[1] * v[0],
    )
    length = math.hypot(*normal)
    if length <= 0:
        return None
    if sum(one * two for one, two in zip(normal, view)) < 0:
        normal = tuple(-value for value in normal)
    return tuple(value / length for value in normal)


def _clip_to_band(
    points: list[tuple[float, float, float]],
    values: list[float],
    low: float,
    high: float,
) -> list[tuple[float, float, float]]:
    """The part of a planar facet whose value lies inside ``[low, high]``.

    Sutherland-Hodgman against the two value planes, interpolating the point along with
    the value it carries -- which is exact, the facet being planar and the value linear
    across it.  That is what puts a band boundary straight across a cell rather than at
    its edges, and it is what the picture shows.
    """
    polygon = list(zip(points, values))
    for keep_above, limit in ((True, low), (False, high)):

        def inside(value: float) -> bool:
            return value >= limit if keep_above else value <= limit

        clipped: list[tuple[tuple[float, float, float], float]] = []
        for index, (point, value) in enumerate(polygon):
            previous_point, previous_value = polygon[index - 1]
            if inside(value):
                if not inside(previous_value):
                    clipped.append(
                        _lerp(previous_point, previous_value, point, value, limit)
                    )
                clipped.append((point, value))
            elif inside(previous_value):
                clipped.append(_lerp(previous_point, previous_value, point, value, limit))
        polygon = clipped
        if not polygon:
            return []
    return [point for point, _value in polygon]


def _lerp(
    first: tuple[float, float, float],
    first_value: float,
    second: tuple[float, float, float],
    second_value: float,
    limit: float,
) -> tuple[tuple[float, float, float], float]:
    """Where the segment crosses *limit*, as a point and the value it carries."""
    span = second_value - first_value
    t = 0.0 if span == 0 else (limit - first_value) / span
    t = min(max(t, 0.0), 1.0)
    return (
        tuple(a + (b - a) * t for a, b in zip(first, second)),
        limit,
    )


@dataclass(frozen=True)
class _Scene:
    """A 3-D chart's box: the front face it is laid out in, and where the back of it is.

    ``depth`` is the whole scene's depth as a **screen displacement** -- add it to a point
    on the front face and you have the point directly behind it on the back wall -- so
    every piece of geometry the scene draws is the front face plus some fraction of it.
    Its horizontal component leads the way the yaw points and its vertical component the
    way the pitch does; both are signed, and one or both can be zero at a degenerate
    camera, which is how a face disappears rather than a special case.
    """

    face: _Rect
    depth: tuple[float, float]
    #: How deep the scene is in the **scene's own units** -- the world the light lives in
    #: -- as a multiple of the face's width.  The depth vector above is that same depth
    #: *projected*, which is not the same number and cannot be un-projected: two of the
    #: three components of a normal survive the projection and the third does not.  A
    #: surface's facets are the first geometry here to need it.
    span: float = 0.0

    def at(self, x: float, y: float, z: float) -> tuple[float, float]:
        """A point on the front face, moved *z* of the way back into the scene."""
        return (x + self.depth[0] * z, y + self.depth[1] * z)


def three_d_scene(
    region: _Rect,
    view: "c.SourceChartView3D",
    *,
    series: int = 1,
    gap_depth: float | None = None,
    kind: str = "bar3DChart",
    grouping: str | None = None,
    categories: int = 5,
    gap_width: float | None = None,
    bar_direction: str | None = None,
    across: int | None = None,
) -> "_Scene | None":
    """The scene's front face and its depth vector, or ``None`` if it does not fit.

    The box is scaled isotropically to fit *region* and centred in it; see
    :func:`three_d_plot_rect`, which is this and its face.  *across* is how many category
    intervals the face spans, which :func:`three_d_scene_shape` reads.
    """
    rot_x = math.radians(view.rot_x or 0.0)
    rot_y = math.radians(view.rot_y or 0.0)
    scale = (
        three_d_scene_shape(
            kind, series, grouping, categories=categories, across=across
        )
        or (1.0, None)
    )[0]
    if region.width <= 0 or region.height <= 0:
        return None
    aspect = scale * (
        view.h_percent / 100.0
        if view.h_percent is not None
        else region.height / region.width
    )
    depth = three_d_scene_depth(
        kind,
        view,
        series=series,
        grouping=grouping,
        categories=categories,
        gap_depth=gap_depth,
        gap_width=gap_width,
        bar_direction=bar_direction,
        aspect=aspect,
    )
    if not (math.isfinite(aspect) and aspect > 0 and math.isfinite(depth) and depth >= 0):
        return None

    run = depth * abs(math.sin(rot_y))
    rise = depth * abs(math.sin(rot_x))
    width = min(
        region.width / (1.0 + VIEW_3D_SCENE_MARGIN + run),
        region.height / (aspect + VIEW_3D_SCENE_MARGIN + rise),
    )
    if not math.isfinite(width) or width <= 0:
        return None
    height = width * aspect
    margin = width * VIEW_3D_SCENE_MARGIN / 2.0
    slack_x = (region.width - width * (1.0 + VIEW_3D_SCENE_MARGIN + run)) / 2.0
    slack_y = (region.height - height - width * (VIEW_3D_SCENE_MARGIN + rise)) / 2.0
    # A half turn of yaw puts the depth in front of the face rather than behind it, and
    # the room it needs moves to the other side with it.
    reversed_ = (view.rot_y or 0.0) % 360.0 >= 180.0
    leads_left = math.sin(rot_y) < 0
    leads_up = (math.sin(rot_x) > 0) != reversed_
    left = region.left + slack_x + margin + (width * run if leads_left else 0.0)
    top = region.top + slack_y + margin + (width * rise if leads_up else 0.0)
    return _Scene(
        face=_Rect(left, top, left + width, top + height),
        depth=(
            -width * run if leads_left else width * run,
            -width * rise if leads_up else width * rise,
        ),
        span=depth,
    )


@dataclass
class _Labels:
    """``c:dLbls`` resolved down to what actually gets printed."""

    show_value: bool = False
    show_category: bool = False
    show_series: bool = False
    show_percent: bool = False
    show_bubble_size: bool = False
    position: str | None = None
    number_format: str | None = None
    font: "ChartFont | None" = None
    color: m.ResolvedColor | None = None

    @property
    def anything(self) -> bool:
        return (
            self.show_value
            or self.show_category
            or self.show_series
            or self.show_percent
            or self.show_bubble_size
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
    #: Whether this series' own group legends with a **rule and its marker** rather than
    #: with a filled swatch.  Carried on the series rather than asked of the chart,
    #: because a combo's groups disagree: the bar series of a bar-plus-line chart keeps
    #: its swatch while the line series beside it takes the rule.
    line_keyed: bool = False
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
        plots: Sequence[c.SourceChartPlot] | None = None,
    ) -> None:
        self.chart = chart
        self.plot = plot
        #: Every drawable group in the plot area, in **document** order.  ``plot`` is the
        #: first of them and stays the one the frame's own decisions -- the title, the
        #: category axis, the legend -- are taken from; see :meth:`_drawn_plots` for the
        #: shape a combo has to be for the rest to be drawn beside it.
        self.plots = list(plots) if plots else [plot]
        self.frame = _Rect(0.0, 0.0, width_pt, height_pt)
        self.style = style
        self._resolve_fill = resolve_fill
        self._resolve_outline = resolve_outline
        self._resolve_text = resolve_text
        self._resolve_typeface = resolve_typeface
        self.elements: list[m.SlideElement] = []
        self._title_cache: "tuple[m.TextBody, FontBox] | None | object" = _UNSET
        #: Forced on for every group of a combo that holds one, because a chart with a
        #: line group in it widens *every* legend key to the line key's width.  See
        #: :meth:`_line_legend_keys`.
        self.line_legend_keys = False
        #: The 3-D box this chart is drawn inside, once :meth:`_plot_rect` has solved it.
        #: ``None`` for every flat chart and for a 3-D one whose camera is refused; see
        #: :meth:`_draws_a_scene`, which is the narrower question of whether the *mesh*
        #: is drawn as well as the plot rectangle placed.
        self.scene: "_Scene | None" = None
        #: The value scale a ``surfaceChart``'s bands are cut from, stashed by
        #: :meth:`_build_cartesian` so that everything asking about the bands asks about
        #: the axis the chart drew.  See :meth:`_band_scale`.
        self.bands_scale: "tuple[float, float, float] | None" = None
        #: The prisms of the group being drawn, held back so they can be painted in depth
        #: order rather than in the order the data happens to be in.  See
        #: :meth:`_paint_prisms`.
        self._prisms: "list[tuple[tuple[float, float, float], _Rect, m.Fill | None, m.Outline | None, m.ResolvedColor]] | None" = None
        #: The 3-D group elements this chart drew **flat**, filled in by :meth:`build`.
        #: Empty when every 3-D group in it got a mesh, which is what stops
        #: ``chart-3d-flattened`` firing on a chart that has no such defect.
        self.flattened_three_d: list[str] = []

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
    def _is_bubble(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in BUBBLE_CHART_KINDS

    @property
    def _is_of_pie(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in OF_PIE_CHART_KINDS

    @property
    def _is_stock(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in STOCK_CHART_KINDS

    @property
    def _is_surface(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) in SURFACE_CHART_KINDS

    @property
    def _is_three_d(self) -> bool:
        """Whether any group in this plot area was authored as a 3-D spelling.

        Asked of the whole plot area rather than of one group because the value axis is
        the *chart's* and a combo would otherwise get two answers for one axis.  In
        practice PowerPoint will not author a 3-D group beside a 2-D one at all.
        """
        return any(c.is_three_d_kind(plot.kind) for plot in self.plots)

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
        elements, data = self._build()
        # The 3-D facts are attached here rather than at each `m.ChartData` call because
        # they are properties of the *chart*, not of the layout that drew it: whichever
        # branch ran, a `bar3DChart` is still a `bar3DChart` and its `c:view3D` is still
        # the camera it asked for.
        data.three_d = self._is_three_d
        data.view_3d = _resolve_view_3d(self.chart.view_3d)
        # Which of this chart's 3-D group elements came out of the build still flat.
        # Asked *after* drawing rather than predicted before it, because that is what the
        # warning claims: `chart-3d-flattened` says what this library did, and a chart
        # whose scene it drew has no defect to declare.
        self.flattened_three_d = sorted(
            {
                plot.kind
                for plot in self._drawn_plots()
                if c.is_three_d_kind(plot.kind)
                and not self._for_plot(plot)._draws_a_scene
            }
        )
        return elements, data

    def _build(self) -> tuple[list[m.SlideElement], m.ChartData]:
        if self._is_radar:
            return self._build_radar()
        if self._is_of_pie:
            return self._build_of_pie()
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

    # -- ofPieChart ---------------------------------------------------------------------

    def _build_of_pie(self) -> tuple[list[m.SlideElement], m.ChartData]:
        """A pie whose smallest points are pulled into a second plot beside it.

        Twenty-four probe charts.  Everything here is the ordinary pie's -- the region, the
        clockwise-from-twelve convention, the per-point accent cycle, the category legend
        -- with three things of its own, all measured:

        * **the two plots are packed across the region's full width**, the first's left
          edge and the second's right edge on the region's own edges, both centred on its
          middle row; the radius falls out of one division (:data:`DEFAULT_OF_PIE_GAP_WIDTH`);
        * **the aggregated slice is centred at three o'clock**, which fixes the rotation of
          both plots (:data:`OF_PIE_OTHER_ANGLE`);
        * **the split is ``auto`` when unstated, and ``auto`` is the last ceil(n/3)
          points** (:data:`OF_PIE_AUTO_DIVISOR`).
        """
        series = self._series()
        categories = self._categories(series)
        region = self._polar_region()
        self._draw_background(region)
        self._draw_title()
        self._draw_of_pie(region, series, categories)
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

    def _of_pie_split(self, values: list[float | None]) -> set[int]:
        """Which point indices belong to the **second** plot.

        Five spellings, four of them measured on a six-point chart of 40/25/15/10/6/4:

        * ``pos`` moves the **last** ``c:splitPos`` points -- ``val="4"`` left 40 and 25 in
          the first plot and moved the other four;
        * ``val`` moves every point **below** ``c:splitPos`` -- ``val="12"`` moved 10, 6
          and 4 and kept 15;
        * ``percent`` is the same test on the point's share of the total -- ``val="15"``
          moved 10%, 6% and 4% and kept the 15%, so the comparison is strict;
        * ``cust`` moves exactly the ``c:secondPiePt`` indices, in their original order --
          ``0`` and ``3`` moved 40 and 10 and left the rest, colours and all, in place;
        * ``auto``, which is also what an absent ``c:splitType`` means, moves the last
          ``ceil(n/3)``.  n = 3, 4, 6, 7 and 8 moved 1, 2, 2, 3 and 3 points; ``round(n/3)``
          is refuted twice over.

        A split that would move everything is drawn as PowerPoint draws it -- the first
        plot becomes one whole-circle slice -- and one that moves nothing leaves the first
        plot a plain pie.  **PowerPoint then draws a dark filled disc where the second plot
        would be**, which this does not: an empty plot is drawn as nothing.
        """
        count = len(values)
        if count <= 0:
            return set()
        kind = (self.plot.split_type or DEFAULT_OF_PIE_SPLIT).strip()
        position = self.plot.split_position

        if kind == "cust":
            return {index for index in self.plot.custom_split if 0 <= index < count}
        if kind == "val":
            threshold = position if position is not None else 0.0
            return {
                index
                for index, value in enumerate(values)
                if value is not None and value < threshold
            }
        if kind == "percent":
            total = sum(abs(value) for value in values if value is not None)
            if total <= 0:
                return set()
            threshold = position if position is not None else 0.0
            return {
                index
                for index, value in enumerate(values)
                if value is not None and abs(value) / total * 100.0 < threshold
            }
        if kind == "pos":
            moved = int(position) if position is not None else 0
        else:  # auto, and anything a file invents
            moved = -(-count // OF_PIE_AUTO_DIVISOR)
        moved = min(max(moved, 0), count)
        return set(range(count - moved, count))

    def _of_pie_geometry(self, region: _Rect) -> tuple[float, float, float, float, bool]:
        """``(first radius, first centre x, second size fraction, second centre x, is bar)``.

        The packing law, measured to the last decimal -- see
        :data:`DEFAULT_OF_PIE_GAP_WIDTH` and :data:`OF_PIE_BAR_GAP_DIVISOR`.

        **The run is centred in the region, not stretched across it.**  Those two are the
        same thing while the *width* is what limits the radius, which is every probe the
        law was fitted on: the drawn span came back as exactly the region's width, to
        0.004 pt, on all 23 slides of ``ofpie-pack``.  Once the region is short enough that
        the **height** caps the radius, they part company, and `ofpie-clamp` says which one
        PowerPoint does: on nine slides across five frames and both forms the drawn span is
        ``divisor * radius`` and its midpoint is the region's own, to 0.005 pt, where
        pinning the two plots to the region's edges puts them up to 160.6 pt out.
        ``chart-gallery`` slide 9 is on that side of the line and is what found this.
        """
        is_bar = (self.plot.of_pie_type or "pie").strip() == "bar"
        size = self.plot.second_pie_size
        if size is None:
            size = DEFAULT_SECOND_PIE_SIZE
        fraction = max(size, 0.0) / 100.0
        gap = self.plot.gap_width
        if gap is None:
            gap = DEFAULT_OF_PIE_GAP_WIDTH
        gap = max(gap, 0.0) / 100.0

        if is_bar:
            divisor = 2.0 + fraction + gap / OF_PIE_BAR_GAP_DIVISOR
        else:
            divisor = 2.0 + 2.0 * fraction + gap
        divisor = max(divisor, MIN_BAR_SLOTS)
        radius = region.width / divisor
        # The region's height caps the radius, and `ofpie-clamp` measures the cap at
        # exactly half of it: five frames whose width law asks for more all drew
        # ``region.height / 2`` to 0.005 pt, in both forms.
        radius = min(radius, region.height / 2)
        if is_bar and fraction > 0:
            # A second *bar* is ``2 * fraction * radius`` tall, so a bar wider than the pie
            # runs out of height first.  **One reading only, and it disagrees**: the one
            # probe short enough to reach this -- `ofpie-clamp`'s 520 x 220 frame at
            # ``gapWidth=300``, ``secondPieSize=125`` -- drew a pie of 49.496 where this
            # asks for 79.2, and a bar 99.0 wide where ``fraction * radius`` is 61.9, so
            # the bar and the pie stop agreeing on a radius at all.  The cap is kept
            # because without it such a chart draws a bar taller than its own region; what
            # PowerPoint replaces it with is unmeasured.  See ROADMAP.md 3.2a.
            radius = min(radius, region.height / (2.0 * fraction))
        # What the run does with the width it did not use: it **centres**, see the
        # docstring.  While the width is the binding constraint this is the identity.
        left = region.left + (region.width - divisor * radius) / 2
        right = left + divisor * radius
        second_x = (
            right - fraction * radius / 2 if is_bar else right - fraction * radius
        )
        return radius, left + radius, fraction, second_x, is_bar

    def _draw_of_pie(
        self, region: _Rect, series: list[_Series], categories: list[str]
    ) -> None:
        if not series:
            return
        item = series[0]
        values = list(item.values)
        total = sum(abs(value) for value in values if value is not None)
        if total <= 0:
            return
        moved = self._of_pie_split(values)
        radius, first_x, fraction, second_x, is_bar = self._of_pie_geometry(region)
        if radius <= 0:
            return
        middle_y = (region.top + region.bottom) / 2

        other = sum(
            abs(values[index])
            for index in sorted(moved)
            if index < len(values) and values[index] is not None
        )
        other_sweep = other / total * 360.0
        # The aggregated slice is centred at three o'clock, so the ring of real slices
        # starts where it ends -- and the second plot starts at the same angle.
        start = OF_PIE_OTHER_ANGLE + other_sweep / 2

        kept = [index for index in range(len(values)) if index not in moved]
        angle = start
        corners: list[tuple[float, float]] = []
        for index in kept:
            value = values[index]
            if value is None or value == 0:
                continue
            sweep = abs(value) / total * 360.0
            self._slice(
                first_x, middle_y, 0.0, radius, angle, sweep,
                fill=self._point_fill(item, index),
                outline=item.point_outlines.get(index, item.outline),
            )
            angle += sweep
        if other > 0:
            # The overflow slice takes the colour one past the last point, which is what
            # the probes show: six points came out accent1..accent6 and the slice accent1
            # of the next cycle, and a three-point chart's came out accent4.
            self._slice(
                first_x, middle_y, 0.0, radius, angle, other_sweep,
                fill=self._of_pie_other_fill(item, len(values)),
                outline=item.outline,
            )
            for edge in (angle, angle + other_sweep):
                radians = math.radians(edge)
                corners.append((
                    first_x + radius * math.sin(radians),
                    middle_y - radius * math.cos(radians),
                ))

        second_radius = fraction * radius
        if is_bar:
            self._draw_of_pie_bar(
                item, values, sorted(moved), other, second_x, middle_y, second_radius
            )
        else:
            self._draw_of_pie_second(
                item, values, sorted(moved), other, second_x, middle_y, second_radius, start
            )
        if self.plot.series_lines is not None and corners:
            self._draw_of_pie_connector(
                corners, second_x, middle_y, second_radius, is_bar
            )
        self._draw_of_pie_labels(
            item, values, categories, kept, sorted(moved), total, other,
            first_x, middle_y, radius, second_x, second_radius, start, is_bar,
        )

    def _of_pie_other_fill(self, item: _Series, index: int) -> m.Fill | None:
        """The aggregated slice takes the colour **one past the last point**.

        Measured on every probe: a six-point chart came out accent1..accent6 with the
        slice in the next cycle's accent1, a four-point one put it in accent5, and a
        three-point one in accent4.
        """
        if self.style.accents and self._vary_colors():
            return m.SolidFill(color=self._cycle_accent(index, index + 1))
        return item.fill

    def _draw_of_pie_second(
        self,
        item: _Series,
        values: list[float | None],
        moved: list[int],
        other: float,
        centre_x: float,
        centre_y: float,
        radius: float,
        start: float,
    ) -> None:
        if other <= 0 or radius <= 0:
            return
        angle = start
        for index in moved:
            value = values[index] if index < len(values) else None
            if value is None or value == 0:
                continue
            sweep = abs(value) / other * 360.0
            self._slice(
                centre_x, centre_y, 0.0, radius, angle, sweep,
                fill=self._point_fill(item, index),
                outline=item.point_outlines.get(index, item.outline),
            )
            angle += sweep

    def _draw_of_pie_bar(
        self,
        item: _Series,
        values: list[float | None],
        moved: list[int],
        other: float,
        centre_x: float,
        centre_y: float,
        radius: float,
    ) -> None:
        """The ``bar`` form's stack: ``s*r`` wide, ``2*s*r`` tall, first point on top.

        Both measured -- 45.801 x 91.602 at r = 61.068 and s = 0.75, and 33.078 x 66.158
        at r = 66.157 and s = 0.5 -- and the order is the probe's: the 60% segment sat
        above the 40% one, which is the order the file lists them in.
        """
        if other <= 0 or radius <= 0:
            return
        width = radius
        height = 2.0 * radius
        left = centre_x - width / 2
        top = centre_y - height / 2
        for index in moved:
            value = values[index] if index < len(values) else None
            if value is None or value == 0:
                continue
            span = abs(value) / other * height
            self._rect(
                _Rect(left, top, left + width, top + span),
                fill=self._point_fill(item, index),
                outline=item.point_outlines.get(index, item.outline),
            )
            top += span

    def _draw_of_pie_connector(
        self,
        corners: list[tuple[float, float]],
        centre_x: float,
        centre_y: float,
        radius: float,
        is_bar: bool,
    ) -> None:
        """``c:serLines`` -- two lines from the aggregated slice to the second plot.

        Measured on the pie form, and the geometry is exact: each line runs from one
        **corner of the aggregated slice** -- where its arc meets the circle -- and is
        **tangent** to the second pie, the upper corner to the upper tangent point.  The
        probe's upper line leaves (97.051, 76.922) for (168.104, 58.528), where the dot
        product of the radius and the line direction is 0.000 and the drawn length
        73.395 pt is exactly sqrt(d^2 - r^2).

        Its *presence* is the switch: a probe with no ``c:serLines`` drew no connector at
        all, and a bare one drew these two in black at 0.5 pt -- the axis default.  An
        explicit ``<a:ln w="28575">`` in red came back red at 2.25 pt.

        **The bar form is not measured** -- no probe put ``c:serLines`` on one -- so its
        lines run to the bar's two left corners, which is the natural analogue and is
        marked as a guess here rather than left undrawn.
        """
        outline = self._axis_outline(self.plot.series_lines.outline)
        if outline is None or radius <= 0:
            return
        upper, lower = sorted(corners, key=lambda point: point[1])
        if is_bar:
            left = centre_x - radius / 2
            self._line(upper[0], upper[1], left, centre_y - radius, outline)
            self._line(lower[0], lower[1], left, centre_y + radius, outline)
            return
        for corner, want_upper in ((upper, True), (lower, False)):
            point = _tangent_point(corner, (centre_x, centre_y), radius, want_upper)
            if point is not None:
                self._line(corner[0], corner[1], point[0], point[1], outline)

    def _draw_of_pie_labels(
        self,
        item: _Series,
        values: list[float | None],
        categories: list[str],
        kept: list[int],
        moved: list[int],
        total: float,
        other: float,
        first_x: float,
        middle_y: float,
        radius: float,
        second_x: float,
        second_radius: float,
        start: float,
        is_bar: bool,
    ) -> None:
        """Point labels on both plots, on the pie's own bisector rule.

        **PowerPoint shrinks both plots to make room for them** -- the label probe's first
        radius came out 37.981 against the 44.105 the same chart draws without labels, the
        same 0.861 on both plots -- and this does not: the plots keep their full size and
        the labels are laid over them.  That is a divergence, measured and recorded rather
        than fitted, because one observation does not say what the reserve is a function of.
        """
        shares = _percent_shares([
            abs(value) if value is not None else 0.0 for value in values
        ])

        def emit(index: int, x: float, y: float) -> None:
            labels = item.point_labels.get(index, item.labels)
            if labels is None or not labels.anything or labels.font is None:
                return
            parts: list[str] = []
            if labels.show_series and item.name:
                parts.append(item.name)
            if labels.show_category and index < len(categories) and categories[index]:
                parts.append(categories[index])
            if labels.show_percent:
                parts.append(f"{shares[index]}%")
            if labels.show_value:
                value = values[index]
                if value is not None:
                    parts.append(
                        format_number(value, labels.number_format or item.format_code)
                    )
            if parts:
                self._centred_label(parts, labels.font, x, y)

        fraction = PIE_LABEL_RADIUS.get(
            (item.labels.position if item.labels else None) or "bestFit", 0.710
        )
        angle = start
        for index in kept:
            value = values[index]
            if value is None or value == 0:
                continue
            sweep = abs(value) / total * 360.0
            radians = math.radians(angle + sweep / 2)
            emit(
                index,
                first_x + radius * fraction * math.sin(radians),
                middle_y - radius * fraction * math.cos(radians),
            )
            angle += sweep

        if other <= 0:
            return
        if is_bar:
            top = middle_y - second_radius
            for index in moved:
                value = values[index] if index < len(values) else None
                if value is None or value == 0:
                    continue
                span = abs(value) / other * 2.0 * second_radius
                emit(index, second_x, top + span / 2)
                top += span
            return
        angle = start
        for index in moved:
            value = values[index] if index < len(values) else None
            if value is None or value == 0:
                continue
            sweep = abs(value) / other * 360.0
            radians = math.radians(angle + sweep / 2)
            emit(
                index,
                second_x + second_radius * fraction * math.sin(radians),
                middle_y - second_radius * fraction * math.cos(radians),
            )
            angle += sweep

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
        region = self._polar_region()
        category_font = self._label_font(category_axis)
        value_font = self._label_font(value_axis)

        labels = self._radar_category_labels(
            categories, category_axis, category_font, region
        )
        # The geometry comes first here, and that is not a tidy-up: a radial axis' tick
        # count is a function of its own radius (see :func:`radial_axis_intervals`), and
        # the radius is set by the *category* labels, so it is knowable before the scale
        # is.  Every other type asks the frame instead.
        centre, radius = self._radar_geometry(
            region, labels, category_font, len(categories)
        )
        scale = self._scale(series, value_axis, radial_pt=radius)

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
        """Every group this chart draws, over one plot rectangle and up to two value axes.

        With one group this is what it always was.  With several -- a **combo** -- the
        groups are drawn in the precedence :data:`COMBO_CHART_KINDS` records, each against
        the axis its own ``c:axId`` names, and the two axes are scaled **separately**:
        measured on ``combo-domain``, where a bar group of 0..9 kept its 0..10 left axis
        while the line group beside it moved a right axis over 0..0.6, 0..6, 0..60 and
        0..600 without touching it.  The same deck's control -- both groups on the primary
        axis -- does move it, to 0..60 and 0..600, which is what says the domain follows
        the attachment and not the plot.

        **The two axes do not share a tick count.**  On ``combo-side``'s ``s150-C`` the
        left axis drew six labels and the right nine, over the same plot height and
        meeting only at the two ends.  What they do share is the *interval count the frame
        asks for*: see :meth:`_axis_scale`.
        """
        groups = [self._for_plot(plot) for plot in self._drawn_plots()]
        if self._line_legend_keys():
            # `self` draws the furniture -- the legend and the plot rectangle it reserves
            # for -- and is not one of the clones, so it has to be told as well.
            self.line_legend_keys = True
            for group in groups:
                group.line_legend_keys = True
        drawn = [(group, group._series()) for group in groups]
        series = [item for _, items in drawn for item in items]
        categories = self._categories(series)

        # `_drawn_plots` has already refused anything `_combo_axes` cannot place, so the
        # fallback here only ever runs for a single group whose own axis crosses at max.
        value_axis, second_axis = self._combo_axes() or (
            self._axis_for(1) or self._axis_of_kind("valAx"),
            None,
        )
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")

        def attached(axis: "c.SourceChartAxis | None") -> list:
            return [pair for pair in drawn if self._plot_value_axis(pair[0].plot) is axis]

        primary = attached(value_axis) or [drawn[0]]
        scale = self._axis_scale(primary, value_axis)
        # The bands are cut from this axis and several things want them before the plot
        # rectangle exists -- the legend's own band, for one.  See :meth:`_band_scale`.
        self.bands_scale = scale
        for group in groups:
            group.bands_scale = scale
        # Each axis carries its own `c:txPr`, and they disagree in real files.
        value_font = self._label_font(value_axis)
        category_font = self._label_font(category_axis)
        tick_texts = primary[0][0]._tick_texts(scale, value_axis)

        secondary = attached(second_axis) if second_axis is not None else []
        second_scale = self._axis_scale(secondary, second_axis) if secondary else None
        second_font = self._label_font(second_axis) if secondary else None
        second_texts = (
            secondary[0][0]._tick_texts(second_scale, second_axis) if secondary else None
        )
        # A `c:delete`d secondary axis still scales its own series -- it is hidden, not
        # absent -- but reserves no band and draws no labels.  Measured on ``p-secdel``,
        # whose plot runs to the same 11.0 pt right inset as the chart with no secondary
        # axis at all while its line series is still drawn across the wider plot.
        shown = bool(secondary) and _labels_shown(second_axis)

        plot_rect = self._plot_rect(
            tick_texts,
            categories,
            value_font,
            category_font,
            scale,
            second_texts=second_texts if shown else None,
            second_font=second_font if shown else None,
        )
        second_band = (
            self._value_label_band([text for _, text in second_texts], second_font)
            - EDGE_INSET_PT
            if shown and second_texts and second_font is not None
            else 0.0
        )

        def scale_for(group: "ChartBuilder") -> tuple[float, float, float]:
            if second_scale is not None and any(group is other for other, _ in secondary):
                return second_scale
            return scale

        # The clones were made before the plot rectangle was solved, and the scene is
        # solved with it: every group draws into the same box.
        for group in groups:
            group.scene = self.scene
        self._draw_background(plot_rect)
        self._draw_title()
        # The secondary axis' gridlines go under the primary's, and unlike the primary's
        # they keep the one at the crossing: measured on ``p-secgrid``, where the
        # secondary drew eleven lines and the primary ten, the secondary's first.
        if secondary and second_scale is not None:
            secondary[0][0]._draw_gridlines(
                plot_rect, second_scale, second_axis, skip_crossing=False
            )
        if self._draws_a_scene:
            # The walls and the floor go under the solids that stand on them, and the
            # category axis -- the floor's own front edge -- goes over them, which is why
            # the floor is split between here and `_draw_axis_lines` below.
            self._draw_scene_gridlines(plot_rect, scale, value_axis)
            self._draw_scene_floor(plot_rect, scale, category_axis)
        else:
            self._draw_gridlines(plot_rect, scale, value_axis)
        for group, items in drawn:
            group._draw_marks(plot_rect, items, categories, scale_for(group))
        self._draw_axis_lines(plot_rect, scale, value_axis, category_axis)
        if secondary:
            self._draw_second_axis_line(plot_rect, second_axis)
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
        if shown and second_texts and second_font is not None and second_scale is not None:
            self._labels_down_right(
                plot_rect,
                [
                    (self._value_to_y(plot_rect, value, second_scale), text)
                    for value, text in second_texts
                ],
                second_font,
            )
        for group, items in drawn:
            group._draw_data_labels(plot_rect, items, categories, scale_for(group))
        # **A side legend stands beside the *region*, not beside the scene's face.**  The
        # legend's band comes off the frame before the camera is fitted into what is left,
        # so the scene shrinking its own front face does not pull the legend in after it.
        # Measured on gallery slide 12, the one chart in the corpus with both a scene and
        # a legend at the side: PowerPoint's nine band keys stand at 685.5 pt where the
        # face's right edge plus the lead gap is 565.5.
        legend_rect = plot_rect if self.scene is None else self._three_d_region(scale)
        if self._is_surface:
            # The bands run lowest first, and a side legend stacks them **upwards**: its
            # top entry is the highest band, measured at all four legend positions.
            bands = self._band_series(scale, value_axis)
            if self._legend_position() in ("l", "r"):
                bands.reverse()
            self._draw_legend(legend_rect, bands, second_band=second_band)
        else:
            self._draw_legend(legend_rect, series, second_band=second_band)

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

    def _axis_scale(
        self, pairs: list, axis: "c.SourceChartAxis | None"
    ) -> tuple[float, float, float]:
        """The scale one value axis draws, from the groups attached to **it**.

        The domain is the union of what each attached group reaches, asked group by group
        so that a stacked group beside a clustered one contributes its sums rather than
        its individual values.  Everything else -- the interval count, the format code,
        the stated limits -- comes from the first group on the axis.

        **A secondary axis counts its intervals by the same frame rule as the primary**,
        which is measured and not assumed: ``combo-side`` put the five-dataset N-meter on
        the secondary axis at five frame heights and read the count back off the drawn
        unit.  The five readings at each height intersect at exactly one count -- 1, 3, 6,
        8 and 10 for frames of 60, 90, 120, 150 and 180 pt -- and those are the five
        counts :func:`side_axis_intervals` returns for the same frames.  Ten more slides
        put the meter on the *primary* of the same charts and read the same counts back,
        so the two axes are one rule applied twice, not one axis leading the other.
        """
        lead, lead_series = pairs[0]
        numbers = [value for group, items in pairs for value in group._axis_reach(items)]
        return lead._scale(lead_series, axis, numbers=numbers)

    def _draw_marks(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """This group's own marks: its areas, its lines or its bars.

        A ``line3DChart`` and an ``area3DChart`` whose scene is drawn take the solid
        instead of the flat mark -- no stroke and no marker for the line, which is
        PowerPoint's picture: a ribbon is the series and there is nothing drawn on it.
        """
        if self._is_surface:
            # A surface has no flat mark at all: without a scene there is nothing to put
            # in its place, so the frame stays empty and `chart-3d-flattened` says so.
            if self._draws_a_scene:
                self._draw_scene_surface(rect, series, categories, scale)
        elif self._draws_a_scene and (self._is_area or self._is_line):
            self._draw_scene_ribbons(rect, series, categories, scale)
        elif self._is_area:
            self._draw_areas(rect, series, categories, scale)
        elif self._is_line:
            self._draw_lines(rect, series, categories, scale)
            if self._is_stock:
                # Drawn over the series, which is the order PowerPoint emitted them in.
                self._draw_hi_low_lines(rect, series, categories, scale)
                self._draw_up_down_bars(rect, series, categories, scale)
        else:
            self._draw_bars(rect, series, categories, scale)

    def _draw_second_axis_line(
        self, rect: _Rect, axis: "c.SourceChartAxis | None"
    ) -> None:
        """The secondary value axis' own line, up the plot's right edge.

        Measured on ``combo-plot``: a stroke from the plot's top to its bottom at exactly
        the right edge, beside the primary axis' identical stroke up the left one.
        """
        if axis is None or axis.delete:
            return
        self._line(
            rect.right, rect.top, rect.right, rect.bottom, self._axis_outline(axis.outline)
        )

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
        :func:`bottom_axis_intervals`, measured on a scatter's own x axis over six frame
        widths.
        """
        series = self._series()
        x_axis, y_axis = self._scatter_axes()

        x_values = [self._x_values(index, item) for index, item in enumerate(series)]
        xs = [value for column in x_values for value in column if value is not None]
        ys = [value for item in series for value in item.values if value is not None]
        x_font = self._label_font(x_axis)
        y_font = self._label_font(y_axis)

        def scales(clearance: float) -> tuple[tuple, tuple]:
            # **Neither axis is anchored at zero**, which every other type here is.  A bar
            # has to start at its axis; a scatter of years against a measurement would be
            # destroyed by it, and PowerPoint agrees -- see
            # :data:`AXIS_ZERO_ANCHOR_RATIO`.  ``clearance`` is zero for everything but a
            # bubble, where it is the largest drawn radius over the plot's height.
            x_base = nice_axis_scale(*_span(xs), anchor_zero=False)
            return (
                _apply_axis_limits(
                    nice_axis_scale(
                        *_span(xs),
                        intervals=self._value_axis_intervals(x_axis, True, x_base),
                        anchor_zero=False,
                        clearance=clearance,
                    ),
                    x_axis,
                ),
                _apply_axis_limits(
                    nice_axis_scale(
                        *_span(ys),
                        intervals=self._value_axis_intervals(y_axis, False),
                        anchor_zero=False,
                        clearance=clearance,
                    ),
                    y_axis,
                ),
            )

        def plot_for(x_scale, y_scale) -> _Rect:
            return self._scatter_plot_rect(
                self._tick_texts(x_scale, x_axis),
                self._tick_texts(y_scale, y_axis),
                x_font,
                y_font,
                x_scale,
                y_scale,
            )

        x_scale, y_scale = scales(0.0)
        rect = plot_for(x_scale, y_scale)
        if self._is_bubble:
            # **A bubble chart's value axis clears the bubbles, not their centres.**  The
            # radius is a length in points and the domain is in data units, so the axis has
            # to be told the radius as a *fraction* of a length -- and that fraction is the
            # plot's **height on both axes**, which is the one genuinely surprising thing
            # here and is measured rather than reasoned: reading the x axis against the
            # plot's own width agrees with 25 of ``bubble-axis``' 40 slides and against its
            # height with 38, and the fifteen it settles are not marginal -- at
            # ``bubbleScale=150`` the width reading leaves x at 0..12 where PowerPoint draws
            # -2..14.  A square plot cannot tell the two apart; these are 437 x 224.
            #
            # The radius is the **largest in the whole chart**, wherever it sits, which is
            # the same reference the diameter takes: ``d-inner`` puts the biggest bubble in
            # the middle of both ranges and its axis is the one ``d-outer`` draws with that
            # bubble on the maximum.
            #
            # The circularity closes after exactly one step, because the radius comes from
            # the frame-derived region and not from the plot
            # (:data:`BUBBLE_REGION_INSET_PT`): the unpadded plot is enough to compute it,
            # and the padded domain's labels are then what the final rectangle is measured
            # from.  Iterating further would **oscillate** rather than converge -- a domain
            # that goes negative moves its own value labels inside the plot, which grows the
            # plot, which shrinks the clearance, which no longer needs the negative -- and
            # the single pass is what PowerPoint's own answers match.
            #
            # See :func:`_clearance_extent` and ROADMAP.md 3.2a.
            radius = self._largest_bubble() / 2.0
            x_scale, y_scale = scales(radius / max(rect.height, 1.0))
            rect = plot_for(x_scale, y_scale)

        x_ticks = self._tick_texts(x_scale, x_axis)
        y_ticks = self._tick_texts(y_scale, y_axis)
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
            self._draw_scatter_series(
                rect, index, item, x_values[index], x_scale, y_scale
            )

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
            band = self._legend_band_height(legend_font)
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
        index: int,
        item: _Series,
        xs: list[float | None],
        x_scale: tuple[float, float, float],
        y_scale: tuple[float, float, float],
    ) -> None:
        points = self._scatter_points(rect, item, xs, x_scale, y_scale)
        if self._is_bubble:
            self._draw_bubbles(index, item, points)
            return
        for run in _split_runs(points):
            if len(run) > 1 and item.line is not None:
                self._polyline(run, item.line, smooth=item.smooth)
        for point in points:
            if point is not None and item.marker_symbol:
                self._marker(point, item)

    # -- bubbles ------------------------------------------------------------------------

    def _bubble_region(self) -> _Rect:
        """What the largest bubble is sized against.

        The frame inset by :data:`BUBBLE_REGION_INSET_PT` on every side, less the title
        band and the legend band.  **Not the plot rectangle**: see the constant, where two
        probes move every plot edge without moving the bubble and two others move only the
        legend and do.  The legend's reserve is taken as the amount it takes off the
        *plot*, which for a side legend is ``_legend_side_width`` minus the edge inset --
        the side-legend probe reproduces to 0.002 pt with no constant of its own.
        """
        frame = self.frame
        inset = BUBBLE_REGION_INSET_PT
        left, right = frame.left + inset, frame.right - inset
        top, bottom = frame.top + inset, frame.bottom - inset

        title = self._title_box()
        if title is not None:
            top += TITLE_BAND_LINES * title.line_height

        legend = self._legend_position()
        if legend is not None and not self._legend_overlays():
            font = self._legend_font()
            band = self._legend_band_height(font)
            if legend == "b":
                bottom -= band
            elif legend in ("t", "tr"):
                top += band
            elif legend in ("l", "r"):
                side = self._legend_side_width(font) - EDGE_INSET_PT
                if legend == "r":
                    right -= side
                else:
                    left += side
        return _Rect(left, top, max(right, left + 1.0), max(bottom, top + 1.0))

    def _bubble_scale(self) -> float:
        scale = self.plot.bubble_scale
        if scale is None:
            scale = DEFAULT_BUBBLE_SCALE
        return max(scale, 0.0)

    def _largest_bubble(self) -> float:
        """The diameter the biggest size in the whole chart is drawn at."""
        region = self._bubble_region()
        scale = self._bubble_scale()
        if scale <= 0:
            return 0.0
        return min(region.width, region.height) * scale / (scale + BUBBLE_SCALE_HALF)

    def _bubble_reference(self) -> float:
        """The largest ``c:bubbleSize`` across **every** series, by magnitude."""
        sizes = [
            abs(size)
            for source in self.plot.series
            for size in (source.bubble_sizes or [])
            if size is not None
        ]
        return max(sizes, default=0.0)

    def _bubble_diameter(self, size: float | None, reference: float, largest: float) -> float:
        if size is None or reference <= 0 or largest <= 0:
            return 0.0
        share = min(abs(size) / reference, 1.0)
        if (self.plot.size_represents or DEFAULT_SIZE_REPRESENTS) == "w":
            return largest * share
        return largest * math.sqrt(share)

    def _bubble_sizes(self, index: int, count: int) -> list[float | None]:
        """One series' ``c:bubbleSize``, padded to the length of its ``c:yVal``.

        A short or absent list is legal; the tail reads as 1, which is what a spreadsheet
        column of blanks would leave.  **Not measured** -- every probe states a full list.
        """
        source = self.plot.series[index] if index < len(self.plot.series) else None
        sizes = list(source.bubble_sizes) if source is not None and source.bubble_sizes else []
        if len(sizes) < count:
            sizes += [1.0] * (count - len(sizes))
        return sizes[:count]

    def _draw_bubbles(
        self, index: int, item: _Series, points: list[tuple[float, float] | None]
    ) -> None:
        """One disc per point, centred on it.

        Three things the schema does not say, each measured:

        * **a size of zero draws nothing** -- the ``-4, 0, 9`` probe emitted two circles;
        * **a negative size draws its magnitude, white with a black 0.75 pt outline**, the
          same drawing a negative bar gets, and ``<c:showNegBubbles val="0"/>`` removes it
          outright while the element being absent or ``1`` draws it;
        * **no outline otherwise** -- every probe disc came back filled and not stroked.
        """
        largest = self._largest_bubble()
        reference = self._bubble_reference()
        sizes = self._bubble_sizes(index, len(points))
        show_negative = self.plot.show_negative_bubbles
        for point, centre in enumerate(points):
            if centre is None:
                continue
            size = sizes[point] if point < len(sizes) else None
            if not size:
                continue
            if size < 0 and show_negative is False:
                continue
            diameter = self._bubble_diameter(size, reference, largest)
            if diameter <= 0:
                continue
            fill = self._point_fill(item, point)
            outline = item.point_outlines.get(point, item.outline)
            if size < 0 and point not in item.point_fills:
                fill = m.SolidFill(color=m.ResolvedColor(hex="#FFFFFF"))
                outline = outline or m.Outline(
                    width=INVERTED_BAR_OUTLINE_EMU,
                    fill=m.SolidFill(color=m.ResolvedColor(hex=DEFAULT_AXIS_COLOR)),
                )
            half = diameter / 2
            self._rect(
                _Rect(centre[0] - half, centre[1] - half, centre[0] + half, centre[1] + half),
                fill=fill,
                outline=outline,
                preset="ellipse",
            )

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
        sizes = [self._bubble_sizes(order, len(item.values)) for order, item in enumerate(series)]
        reference = self._bubble_reference() if self._is_bubble else 0.0
        largest = self._largest_bubble() if self._is_bubble else 0.0
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
                size = sizes[order][point] if point < len(sizes[order]) else None
                text = self._label_text(
                    labels, item, categories[order], point, value, totals,
                    bubble_size=size if self._is_bubble else None,
                )
                if not text:
                    continue
                centre = (
                    self._value_to_x(rect, x, x_scale),
                    self._value_to_y(rect, value, y_scale),
                )
                position = (labels.position or self._label_default()).lower()
                radius = item.marker_size / 2
                side, edge = DATA_LABEL_LINE_GAP_EM * labels.font.size, DATA_LABEL_GAP_PT
                if self._is_bubble:
                    # The label stands off the *disc*, and further off it than a marker's
                    # does: 8.494 pt against 0.6 em = 6.0, and 7.0 pt against 4.85.
                    radius = self._bubble_diameter(size, reference, largest) / 2
                    side, edge = BUBBLE_LABEL_SIDE_GAP_PT, BUBBLE_LABEL_EDGE_GAP_PT
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
                self._place_label(text, labels, geometry, False, side=side, edge=edge)

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
                count = self._point_color_count(len(item.values))
                item.vary_fills = [
                    m.SolidFill(color=self._cycle_accent(index, count))
                    for index in range(len(item.values))
                ]
            if (
                self._is_line
                # A bubble series has neither: PowerPoint drew a bare disc per point on
                # every one of the thirty probes, with no connecting stroke and no marker.
                or (self._is_scatter and not self._is_bubble)
                or (self._is_radar and self._radar_style != "filled")
            ):
                # **A scatter series is styled exactly like a line series, and
                # ``c:scatterStyle`` decides nothing.**  Probes at `marker`, `line` and
                # `lineMarker` came back byte-identical -- line *and* markers in all three
                # -- so what turns either off is the series' own markup: `<a:ln><a:noFill/>`
                # for the line, `<c:symbol val="none"/>` for the marker, each measured.
                self._read_line_style(item, source, source.index)
                # The key shape follows the series, not the chart: a combo's bar series
                # keeps its swatch while the line beside it takes a rule.
                #
                # **A ``line3DChart`` keys with a swatch**, because its mark is a solid and
                # not a stroke: PowerPoint draws gallery slide 14's legend as two filled
                # squares where a flat line chart of the same data gets the rule and its
                # marker.  Gated on the spelling rather than on whether the scene is drawn,
                # because it is PowerPoint that draws the ribbon either way.
                item.line_keyed = self.plot.kind != "line3DChart"
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
            show_bubble_size=flag("show_bubble_size"),
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
        """Whether the marks are a stroke through the points with markers on it.

        **A stock chart answers yes**, measured: one with neither ``c:hiLowLines`` nor
        ``c:upDownBars`` came back from PowerPoint as a plain line chart -- 1.5 pt strokes,
        the 6 pt diamond/square/triangle marker cycle, the line-chart legend key.  Its two
        decorations are drawn on top of that, and the missing lines a real stock chart has
        are the file's own ``<a:ln><a:noFill/></a:ln>``, not a rule in the renderer.
        """
        kind = c.flat_chart_kind(self.plot.kind)
        return kind == "lineChart" or kind in STOCK_CHART_KINDS

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

    def _point_color_count(self, points: int) -> int:
        """How many colours this chart cycles through, which decides the ramp.

        An ofPie needs one more than it has points: its aggregated slice takes a colour of
        its own, and a six-point ofPie -- seven slices -- is already ramped.
        """
        return points + 1 if self._is_of_pie else points

    def _cycle_accent(self, index: int, count: int) -> m.ResolvedColor:
        """Accent ``index`` as drawn, given how many colours the chart needs in total.

        Six or fewer and it is the plain accent; past that each cycle of six takes a
        luminance shift.  See :data:`VARY_COLOR_CYCLE_SHADE`.
        """
        accents = self.style.accents
        base = accents[index % len(accents)]
        # The cycle is six long because a theme has six accents, not because *this* style
        # carries six: a caller that supplies fewer still cycles PowerPoint's six.
        if count <= len(ACCENT_KEYS):
            return base
        shifted = _cycle_shift(base.hex, index // len(ACCENT_KEYS))
        return m.ResolvedColor(hex=shifted, alpha=base.alpha)

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
        # Every drawn group, in document order: a combo's groups share one category axis
        # and need not all carry `c:cat` -- the gallery's line group does, but a group
        # that did not would otherwise number its categories from 1 beside labelled bars.
        for plot in self.plots:
            for source in plot.series:
                if any(source.categories):
                    labels = list(source.categories)
                    break
            if labels:
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

    # -- combo charts -----------------------------------------------------------------

    def _for_plot(self, plot: c.SourceChartPlot) -> "ChartBuilder":
        """This builder, reading a different group of the same chart.

        A shallow copy, so ``elements`` is the *same* list: every clone draws into one
        output in the order it is asked to.  It is what keeps the seventy-odd
        ``self.plot`` readers in this class working unchanged on a chart with several
        groups -- each group is drawn by a builder for which it *is* ``self.plot``.
        """
        clone = copy.copy(self)
        clone.plot = plot
        return clone

    def _plot_value_axis(self, plot: c.SourceChartPlot) -> "c.SourceChartAxis | None":
        """The value axis *this* group is attached to, by its own ``c:axId``."""
        return self._for_plot(plot)._axis_for(1) or self._axis_of_kind("valAx")

    def _combo_axes(self) -> "tuple[c.SourceChartAxis | None, c.SourceChartAxis | None] | None":
        """The chart's value axes as ``(left, right)``, or ``None`` if they cannot be placed.

        **``c:crosses`` decides the side and ``c:axPos`` decides nothing.**  Measured on
        ``combo-plot``: a secondary axis written ``axPos="l"`` with ``crosses="max"`` came
        out on the right, in the same place as the ``axPos="r"`` one beside it, and the
        same axis with ``crosses="autoZero"`` came out on the **left**, inside the primary
        axis, with the plot narrowing to make room for two label columns.

        That last shape -- two value axes on the same side -- is drawn by PowerPoint and
        is *not* drawn here: its inner label column measured 21.42 pt against the outer
        one's 26.41 pt for the same label, and one reading is not a rule.  This returns
        ``None`` for it, which sends the chart back to drawing one group.
        """
        axes: list[c.SourceChartAxis | None] = []
        for plot in self.plots:
            axis = self._plot_value_axis(plot)
            if not any(axis is seen for seen in axes):
                axes.append(axis)
        if len(axes) == 1:
            return axes[0], None
        right = [axis for axis in axes if axis is not None and axis.crosses == "max"]
        left = [axis for axis in axes if not any(axis is other for other in right)]
        if len(left) == 1 and len(right) == 1:
            return left[0], right[0]
        return None

    def _drawn_plots(self) -> list[c.SourceChartPlot]:
        """Every group drawn on this chart, in paint order.

        A combo is drawn whole only when its groups are the three that share a category
        axis (:data:`COMBO_CHART_KINDS`), every bar among them runs the same way up, and
        their value axes can be placed on opposite sides of the plot.  Anything else --
        a pie beside a bar, a horizontal bar in a combo, two value axes on one side --
        falls back to the one group this renderer picked before combos were drawn at all,
        which is a worse picture than PowerPoint's but not a wrong one.
        """
        if len(self.plots) < 2:
            return [self.plot]
        kinds = [c.flat_chart_kind(plot.kind) for plot in self.plots]
        if any(kind not in COMBO_CHART_KINDS for kind in kinds):
            return [self.plot]
        # `barDir="bar"` turns the value axis along the bottom and the secondary axis
        # along the top, which no probe has measured.
        if any(
            (plot.bar_direction or "col") != "col"
            for plot, kind in zip(self.plots, kinds)
            if kind == "barChart"
        ):
            return [self.plot]
        if self._combo_axes() is None:
            return [self.plot]
        # A stable sort, so two groups of one type keep the order the file wrote them in.
        return sorted(self.plots, key=lambda plot: COMBO_CHART_KINDS.index(
            c.flat_chart_kind(plot.kind)
        ))

    def _line_legend_keys(self) -> bool:
        """Whether this chart's legend keys are rules rather than swatches.

        **One line group turns every key into a rule**, which is measured rather than
        assumed: on ``combo-legend`` the bar series of a bar-plus-line chart came back
        with a key 19.200 pt wide -- the line key's width, not the 5.49 pt swatch -- and
        5.49 pt tall.  So the *width* is a decision for the chart and the *shape* one for
        the series.  A chart with no line group is unaffected, which is why a plain bar
        chart's key is unchanged.
        """
        return any(self._for_plot(plot)._is_line_keyed for plot in self._drawn_plots())

    @property
    def _is_line_keyed(self) -> bool:
        """Whether *this* group's series take a line key rather than a swatch.

        A group whose shape asks for a rule but which has **no rule to draw** does not
        count: see :meth:`_draws_a_rule`.
        """
        return bool(
            (
                self._is_line
                or (self._is_scatter and not self._is_bubble)
                or (self._is_radar and self._radar_style != "filled")
            )
            and self._draws_a_rule()
        )

    def _draws_a_rule(self) -> bool:
        """Whether any series in this group has a stroke the legend could show.

        ``<a:ln><a:noFill/></a:ln>`` is an explicit *no line* -- a stock chart states it on
        every series -- and a group where every series says so legends with the **swatch
        cell**, not the 24.0 pt line cell.  Measured on ``legend-nokey``, sixteen slides:

        * three bare line series draw **no key path at all** and lay out on a 1.0984 em
          cell, the swatch cell to 0.0001 em, at three frame widths and at two, three,
          four and five entries;
        * the same three with a marker draw the **marker alone**, centred in that same
          cell to 0.07 pt -- so what is lost is the rule, not the key;
        * **one** series keeping its rule puts the whole chart back on the 24.002 pt cell
          and the bare entries beside it simply draw nothing, which is what says the width
          is a decision for the chart and the ink one for the series;
        * a real ``c:stockChart`` reads the same as the line chart, both ways round.

        The cell is solved twice per slide -- from the pitch between labels and from the
        run's centring -- and the two agree: every slide's first label lands 0.366 pt left
        of the predicted advance origin, the same side bearing 3.5 already carries, where
        reading the line cell here is 8.5 pt out and reading no cell at all 7.5 pt the
        other way.

        ``_resolve_outline`` collapses "no line" and "no ``c:spPr``" to the same ``None``,
        so the source element is what gets asked -- the same trap :meth:`_axis_outline`
        and :meth:`_read_line_style` exist for.
        """
        return any(
            not (
                isinstance(source.outline, s.SourceOutline)
                and isinstance(source.outline.fill, s.SourceNoFill)
            )
            for source in self.plot.series
        )

    def _axis_reach(self, series: list[_Series]) -> list[float]:
        """The values this group's own marks reach on its value axis.

        Split out of :meth:`_scale` because a **combo** puts several groups on one axis
        and each reaches its own way: a stacked group reaches the sum of its category
        where a clustered one beside it reaches its tallest bar, so the axis' domain is
        the union of what each group answers here rather than one sum over all of them.
        """
        if (self.plot.grouping or "clustered") in ("stacked", "percentStacked"):
            # A stacked bar reaches the sum of its category, and the positive and negative
            # halves of that category stack away from zero independently.
            numbers: list[float] = []
            length = max((len(item.values) for item in series), default=0)
            for index in range(length):
                column = [_at(item.values, index) for item in series]
                numbers.append(sum(value for value in column if value and value > 0))
                numbers.append(sum(value for value in column if value and value < 0))
            return numbers
        return [value for item in series for value in item.values if value is not None]

    def _scale(
        self,
        series: list[_Series],
        axis: c.SourceChartAxis | None,
        radial_pt: float | None = None,
        numbers: list[float] | None = None,
    ) -> tuple[float, float, float]:
        if (self.plot.grouping or "") == "percentStacked":
            # Measured: PowerPoint labels 0%, 10% ... 100%.
            return 0.0, 1.0, 0.1

        if numbers is None:
            numbers = self._axis_reach(series)
        else:
            numbers = list(numbers)

        if not numbers:
            numbers = [0.0]
        horizontal = (self.plot.bar_direction or "col") == "bar"

        def scaled(intervals: int) -> tuple[float, float, float]:
            return nice_axis_scale(
                min(numbers),
                max(numbers),
                intervals=intervals,
                # A radar stops at the data rather than a whole unit past it: 0..5 of data
                # gives a 0..5 axis where the same data on a bar gives 0..6.  **So does a
                # 3-D chart**, and that is measured rather than borrowed: see
                # :func:`nice_axis_scale`.
                strict=not (self._is_radar or self._is_three_d),
            )

        # The finest axis the data could take, which is the answer for a frame with room
        # for ten intervals and the label estimate a narrower one is measured against.
        provisional = scaled(AXIS_MAX_INTERVALS)
        intervals = self._value_axis_intervals(axis, horizontal, provisional, radial_pt)
        scale = provisional if intervals >= AXIS_MAX_INTERVALS else scaled(intervals)
        return _apply_axis_limits(scale, axis)

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

    @property
    def _three_d_view(self) -> "c.SourceChartView3D | None":
        """The camera this group's plot rectangle is laid out through, or ``None``."""
        return three_d_camera(
            self.chart.view_3d,
            self.plot.kind,
            series=len(self.plot.series) or 1,
            grouping=self.plot.grouping,
        )

    def _three_d_region(self, scale: tuple[float, float, float]) -> _Rect:
        """The **flat** plot rectangle, which is the region the 3-D scene is fitted into.

        Measured: PowerPoint lays the scene out inside the rectangle the same chart drawn
        flat would get.  The ``flat*`` controls on ``view3d-meter`` give that rectangle
        directly -- a 195 pt frame's is inset 10.01 pt at the top and 25.87 at the bottom
        -- and the scene's own extent fills its height exactly at every ordinary pitch.
        """
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")
        return self._plot_rect(
            self._tick_texts(scale, value_axis),
            self._categories(self._series()),
            self._label_font(value_axis),
            self._label_font(category_axis),
            scale,
            flat=True,
        )

    def _three_d_reservation(self, scale: "tuple[float, float, float] | None") -> float:
        """What the scene's depth takes off the height the value axis has to divide.

        This is the whole of :func:`side_axis_intervals`' missing input.  The count was
        never a function of the *frame* for a 3-D chart -- six cells on one and the same
        120 pt frame need one, three to four and five to eight intervals -- and it is not a
        function of the drawn plot either, because a flat chart's own furniture still comes
        off the frame on top of it.  It is the frame less this, which reproduced all 52
        cells of the two meter decks when the reservation was still being measured by hand.

        *scale* is the finest axis the data could take, drawn at
        :data:`AXIS_MAX_INTERVALS`.  The region depends on the tick labels and the labels
        depend on the count, so the circle has to be cut somewhere; it is cut here, where
        the error is a fraction of one label's width against a region six hundred points
        wide.
        """
        view = self._three_d_view
        if view is None or scale is None:
            return 0.0
        region = self._three_d_region(scale)
        return max(0.0, region.height - self._three_d_face(region, view).height)

    def _three_d_face(self, region: _Rect, view: "c.SourceChartView3D") -> _Rect:
        """:func:`three_d_plot_rect`, told how deep this plot's own scene is."""
        scene = self._three_d_scene(region, view)
        return region if scene is None else scene.face

    def _three_d_scene(
        self, region: _Rect, view: "c.SourceChartView3D"
    ) -> "_Scene | None":
        """:func:`three_d_scene`, told the plot's series, categories and two gaps.

        The category count is an input to the depth and not only to the layout: a 3-D bar
        is as deep as it is wide and its width is a share of one category band, so the
        same chart drawn over three categories stands in a deeper scene than over five.
        How many *intervals* those categories span -- one less when the marks sit on the
        ticks rather than in the bands -- is an input to a ``line3DChart``'s and an
        ``area3DChart``'s **aspect** in the same way; see :func:`three_d_scene_shape`.
        """
        count = len(self._categories(self._series())) or 1
        return three_d_scene(
            region,
            view,
            series=len(self.plot.series) or 1,
            gap_depth=self.plot.gap_depth,
            kind=self.plot.kind,
            grouping=self.plot.grouping,
            categories=count,
            gap_width=self.plot.gap_width,
            bar_direction=self.plot.bar_direction,
            across=max(count - 1, 1) if self._points_on_ticks else count,
        )

    def _axis_band_height(self, scale: "tuple[float, float, float] | None" = None) -> float:
        """The frame height a value axis up the side has to divide.

        The title and a legend above or below come off it, in exactly the bands
        :meth:`_plot_rect` reserves for them: measured on a 150 pt frame, where a bottom
        legend took the interval count from 8 to 6 and a title took it to 6, both of which
        those bands predict.  A legend at the *side* left the count alone, so nothing is
        taken off for one.  See :func:`side_axis_intervals`.

        A 3-D chart's scene comes off it as well; see :meth:`_three_d_reservation`.
        """
        height = self.frame.height
        title = self._title_box()
        if title is not None:
            height -= TITLE_BAND_LINES * title.line_height
        legend = self._legend_position()
        if legend in ("b", "t", "tr") and not self._legend_overlays():
            height -= self._legend_band_height(self._legend_font())
        return height - self._three_d_reservation(scale)

    def _axis_band_width(self) -> float:
        """The frame width a value axis along the bottom has to divide.

        The mirror of :meth:`_axis_band_height`, and the legend half of it is **not
        measured**: a bottom legend is what was shown to come off the height, and no probe
        put a side legend on a chart whose value axis runs along the bottom.  The band a
        side legend takes is the one the plot already loses, so this is that same
        measurement applied to the other axis rather than a second guess.
        """
        width = self.frame.width
        legend = self._legend_position()
        if legend in ("l", "r") and not self._legend_overlays():
            width -= self._legend_side_width(self._legend_font())
        return width

    def _value_axis_intervals(
        self,
        axis: "c.SourceChartAxis | None",
        horizontal: bool,
        provisional: "tuple[float, float, float] | None" = None,
        radial_pt: float | None = None,
    ) -> int:
        """How many major intervals this chart's value axis is divided into.

        A bottom axis' rung depends on how wide its labels are, its labels depend on the
        unit and the unit depends on the rung.  PowerPoint faces the same circle; this
        settles it by measuring the labels of the finest axis the data could take -- the
        *provisional* one, drawn at :data:`AXIS_MAX_INTERVALS` -- and asking the rung about
        those.  A side axis does not read its labels at all, which is measured: three-digit
        integers and four-character decimals coarsen at exactly the same frame heights.
        """
        font = self._label_font(axis)
        if radial_pt is not None:
            return radial_axis_intervals(radial_pt, font.box.pitch)
        if not horizontal:
            return side_axis_intervals(self._axis_band_height(provisional), font.box.pitch)
        widest = 0.0
        if provisional is not None:
            texts = [text for _, text in self._tick_texts(provisional, axis)]
            widest = max((font.width(text) for text in texts), default=0.0)
        return bottom_axis_intervals(self._axis_band_width(), font.size, widest)

    def _label_font(self, axis: c.SourceChartAxis | None) -> ChartFont:
        return self._font(axis.text_properties if axis is not None else None)

    def _legend_font(self) -> ChartFont:
        legend = self.chart.legend
        return self._font(legend.text_properties if legend is not None else None)

    def _value_label_band(self, texts: list[str], font: ChartFont) -> float:
        """What a column of value-axis tick labels takes off the frame's edge, in points.

        **The band on the right of a secondary axis is this same formula**, fed that axis'
        own labels -- measured to 0.04 pt on ``combo-domain`` and ``combo-plot`` over five
        label widths: a right-hand axis labelled 0..6 reserved 21.07 pt, one labelled
        0..10 or 0..60 reserved 26.41 and one labelled 0..600 reserved 31.75, which are
        the 21.11 / 26.45 / 31.79 this returns.  The same deck's left-hand axis labelled
        0..600 reserved 31.75 as well, so the two sides are one rule and not two.
        """
        widest = max((font.width(text) for text in texts), default=0.0)
        return FRAME_PADDING_PT + widest + font.box.descent + VALUE_LABEL_GAP_EM * font.size

    def _plot_rect(
        self,
        tick_texts: list[tuple[float, str]],
        categories: list[str],
        value_font: ChartFont,
        category_font: ChartFont,
        scale: tuple[float, float, float],
        *,
        second_texts: list[tuple[float, str]] | None = None,
        second_font: ChartFont | None = None,
        flat: bool = False,
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
            left = frame.left + self._value_label_band(down_left, left_font)

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

        # A secondary value axis takes its own label column off the *right* edge, in
        # place of the plain inset.  Measured on ``combo-plot``: with no secondary axis
        # the plot's right inset is 11.0 pt and with one it is the label band above.  A
        # secondary axis that states `c:delete` takes nothing -- ``p-secdel``'s plot runs
        # to the same 11.0 pt inset as the chart with no second axis at all, with the
        # series still drawn.
        second_band = (
            self._value_label_band([text for _, text in second_texts], second_font)
            if second_texts and second_font is not None
            else EDGE_INSET_PT
        )
        right = frame.right - second_band - overhang_right
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
            band = self._legend_band_height(legend_font)
            if legend in ("b",):
                legend_bottom = band
            elif legend in ("t", "tr"):
                top += band
            elif legend == "r":
                # The side band *replaces* the plain edge inset rather than adding to it:
                # it already ends in its own trailing pad.  Measured on
                # real-financial-report's two bar charts, whose legends are Japanese and
                # of different lengths -- both were over by exactly 11.0 pt, the inset.
                #
                # A secondary axis' labels go *between* the plot and that band, so the two
                # reserves compose: the legend keeps the place it would have had on its
                # own -- measured to 0.04 pt on ``combo-legend``'s ``l-right`` -- and the
                # plot gives up the label column on top of it.  **One reading, residual
                # -0.31 pt**: the drawn plot ended at 399.08 pt where this predicts
                # 398.77.  A sweep of the secondary label width with a right legend is
                # what would settle the third of a point.
                right = (
                    frame.right
                    - self._legend_side_width(legend_font)
                    - (second_band - EDGE_INSET_PT)
                    - overhang_right
                )
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
        region = _Rect(left, top, right, bottom)
        # A 3-D chart's plot rectangle is its scene's **front face**, which is this
        # rectangle displaced and shrunk to make room for the depth.  `flat` asks for the
        # region itself, which is what the camera is fitted into.
        view = None if flat else self._three_d_view
        if view is None:
            return region
        scene = self._three_d_scene(region, view)
        if scene is None:
            return region
        self.scene = scene
        return scene.face

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

    def _furniture_height(self) -> float:
        """The frame's height, less the bands a title and a top or bottom legend take.

        This is what the turned label's allowance is half of.  A **side** legend is not in
        it: the probe with one reserved and cut identically to the probe with none, which
        is what says the allowance is a height rule rather than an area one.  A bottom
        legend was probed at three frame heights, a title at three, and a top legend at
        one -- the top legend cut at exactly the same character as the bottom one, which
        is why both are here rather than only the one the band is drawn under.
        """
        height = self.frame.height
        title = self._title_box()
        if title is not None:
            height -= TITLE_BAND_LINES * title.line_height
        legend = self._legend_position()
        if legend in ("b", "t", "tr") and not self._legend_overlays():
            height -= self._legend_band_height(self._legend_font())
        return height

    def _rotated_allowance(self, box: FontBox) -> float:
        """How wide a turned category label may be before PowerPoint cuts it.

        The deepest pen of a turned label may drop at most
        ``height / 2 - ROTATED_LABEL_HEADROOM_PT`` below the axis line; the anchor takes
        the fixed part of that drop and ``sin 45`` turns what is left back into a width.
        """
        headroom = (
            self._furniture_height() / 2
            - ROTATED_LABEL_HEADROOM_PT
            - rotated_label_anchor(box)
        )
        return max(headroom / _SIN_45, 0.0)

    def _bottom_label_band(
        self, font: ChartFont, categories: list[str], plot_width: float
    ) -> float:
        """How much of the frame the labels under the plot take.

        Level and on one line it is ``FRAME_PADDING_PT + (5/3) * ascent + descent``: the
        padding, the baseline's own drop below the axis line, and the descender hanging
        off it.  **The gap term is two thirds of the ascent and not a fraction of the
        em**, which is :data:`CATEGORY_LABEL_GAP_ASCENT`'s own table -- 24 charts in four
        faces off ``axis-inset``'s gridlines, where the previous four readings were one
        face and could not separate the two.  For Aptos, the face all of them were drawn
        in, the change is +0.011 em; for Courier New it is +0.06.

        **A short frame breaks it**, and that is unresolved: 24 pt labels on a 90 pt frame
        take a bottom band of 46.876 pt in Aptos and 47.578 in Arial where this rule --
        and the same labels on 120, 195 and 330 pt frames -- give 50.818 and 47.793.  The
        top inset does not move, the plot keeps a little over 20 pt of height, and nothing
        in the sweep says what the floor is.  Ours is 3.7 pt too deep there and correct on
        every frame that is not that short.

        Turned, it is the same three terms with the line box **and** the label's own
        width turned through 45 degrees -- its gap term still
        :data:`CATEGORY_LABEL_GAP_EM`, for the reason recorded there, and the width the
        one PowerPoint *draws* rather than the one the deck authored; see
        :data:`ROTATED_LABEL_HEADROOM_PT` and :func:`truncate_label`.
        **Wrapped, it is the level band plus one
        line box for every line after the first** -- and the line that sets it is the one
        needing the most lines, not the widest string, although no probe separates those
        two because in all twenty they were the same label.

        The extra line is the face's own line **pitch**, measured on a four-rung ladder in
        five faces at 10 pt.  The band grew by exactly the same amount from one line to
        two, two to three and three to four in every face, so this is a straight line and
        not a fit:

        ====================  ==========  ===========  ===========  =========
        face                  per line    line box     + lineGap    residual
        ====================  ==========  ===========  ===========  =========
        Calibri               12.205      12.207       12.207       -0.002
        Aptos                 12.205      12.207       12.207       -0.002
        Courier New           11.330      11.328       11.328       +0.002
        Times New Roman       11.075      11.074       11.074       +0.001
        Arial                 11.500      11.172       **11.499**   +0.001
        ====================  ==========  ===========  ===========  =========

        **Arial is what identified the term.**  PowerPoint's pitch is the face's full
        ``hhea`` line spacing -- ascender plus descender plus *lineGap* -- and Arial is the
        only one of the five whose lineGap is not zero: 67 units of 2048, 0.328 pt at
        10 pt, exactly the residual the fourth column removes.  The other four are
        unmoved because their gap is genuinely zero, which is why this table reads as a
        confirmation rather than a refit.

        **The gap it carries is Arial's, not Arimo's**, and that distinction is the whole
        reason the column could be added at all: Tinos' lineGap is 87 where Office's own
        ``times.ttf`` is 0, so reading the substitute would have traded Arial's 0.33 pt
        error for a 0.42 pt one on Times New Roman.  :class:`~pptx2svg.text.metrics.FontMetrics`
        now carries the *Office* face's gap as a measured number, on the footing the Aptos
        advance widths already stand on.

        **One line is the line box, not the pitch**, and that is measured too rather than
        assumed: a ladder gives only the slope, but the level band itself is
        ``6.5 + (5/3) * ascent + descent`` over ``axis-inset``'s four faces and eight
        sizes, and Arial is in it at every size with its 0.33 pt gap at 10 pt nowhere in
        the residual.  Leading goes between lines, not above the first.

        **Turned, the band is capped, and the cap is a truncation.**  This used to record
        an unidentified clamp: a label 4.18 band widths wide reserved 85.63 pt where the
        uncapped formula asks for 113.95, and reserved *less* than the 2.92-band label one
        step below it, which took 86.36 -- so no clamp on the width could produce both.
        Two probe decks, 88 readings, settle it, and the 0.73 pt inversion is the part that
        names the mechanism rather than the part that resisted it:

        * the 2.92-band label is **91.86 pt and fits** the 101.13 pt allowance its frame
          gives it, so it is drawn whole and reserves ``21.28 + 91.86 sin 45`` = 86.23;
        * the 4.18-band one is 130.90 pt, does **not** fit, and is cut to the longest
          prefix that does -- ``CategoryLongerStillA`` at 90.83 pt -- which reserves 85.51.

        A cut prefix is necessarily a hair *narrower* than the whole label that just fits,
        so the wider category reserves less.  Both magnitudes land within 0.12 pt of what
        PowerPoint drew and the gap between them is 0.728 against a measured 0.728, on two
        numbers that were not in the fit.

        :data:`ROTATED_LABEL_HEADROOM_PT` carries the sweep the cap came out of and
        :func:`truncate_label` the cut.  The allowance is on the label's **width**, not on
        the band: an untruncated label may reserve more than the cap would allow a cut one,
        which is what separates this from the half-the-frame-height rule the three
        observations suggested and refutes that rule outright.

        ``real-financial-report.pptx``'s chart3 comes out of this without being in it: its
        allowance is 62.3 pt, which keeps ``グローバル`` (60) whole and cuts
        ``プラットフォーム`` (96) to ``プラット…`` exactly as PowerPoint did, for a band of
        68.6 against its measured **69.538**.

        The decks are built and read by ``tools/make_label_probe.py`` and
        ``tools/read_label_probe.py``, and the reader's second line per chart is this rule's
        own prediction, so a change here is checked against the exports rather than argued
        about.  What it does not reproduce is in :data:`ROTATED_LABEL_HEADROOM_PT`: 17 of
        95 readings are cut within one and a half characters of PowerPoint rather than on
        it, and the CJK probes are not a test of any of this, because the probe deck's
        Japanese fell back to **MS Gothic**, whose metrics this library does not carry.
        """
        if self._labels_rotate(font, categories, plot_width):
            box = font.box_for(*categories)
            allowance = self._rotated_allowance(box)
            # The ellipsis hangs *past* the anchor, towards the axis, so the band is sized
            # from the kept prefix rather than from the drawn string.
            widest = max(
                font.width(truncate_label(text, font, allowance).removesuffix(LABEL_ELLIPSIS))
                for text in categories
            )
            return (
                FRAME_PADDING_PT
                + (box.line_height + widest) * _SIN_45
                + CATEGORY_LABEL_GAP_EM * box.size
            )
        # The line box is the *drawn* face's, which for a Japanese label is not the one
        # `<a:latin>` names -- see `ChartFont.box_for`.
        box = font.box_for(*categories)
        lines = self._label_line_count(font, categories, plot_width)
        return (
            FRAME_PADDING_PT
            + box.line_height
            + box.pitch * (lines - 1)
            + CATEGORY_LABEL_GAP_ASCENT * box.ascent
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
            band = self._legend_band_height(font, per_point=per_point)
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

        ``max(11.0, 5.0 + lineHeight/2)``, and it is **right**: ``axis-inset`` reads the
        plot's top edge off its own topmost gridline in four faces at eight sizes and this
        rule lands within 0.002 pt of every one of the 24.  The floor binds at 6 and 8 pt
        in every face and at 10 pt in all but Aptos, and each of those draws its gridline
        at exactly 11.000 pt, so :data:`EDGE_INSET_PT` is the floor as a measurement and
        not as a guard.

        This was for a while suspected of putting the whole plot rectangle 3.7 pt low, on
        a reading that differenced our tick labels' **boxes** against PowerPoint's
        **ink**; see ``tools/read_view3d_probe.ours``, where that is now fixed, and
        ROADMAP.md 3.4.  The gridlines were what settled it: they are the rectangle, and
        they are where PowerPoint puts them.
        """
        return max(EDGE_INSET_PT, TOP_INSET_BASE_PT + label.line_height / 2)

    def _legend_grid(self, widths: list[float]) -> tuple[int, int]:
        """A horizontal legend's ``(rows, columns)``.

        Past :data:`LEGEND_BAND_MAX_FRACTION` of the frame the entries no longer fit on one
        line and PowerPoint lays them out as a **grid of equal columns**, each as wide as
        the widest entry and packed with no gap at all.  The arithmetic is measured on the
        ``legend-row`` deck's 40 slides and is not the greedy fill it looks like:

            columns that fit = floor(0.9 * frame / widest entry)
            rows             = ceil(entries / columns that fit)
            columns drawn    = ceil(entries / rows)

        The second step is what a greedy fill gets wrong.  Six equal entries with four to a
        row came back **3 + 3**, not 4 + 2; seven entries with three to a row came back
        3 + 3 + 1, which is the same rule and not a balanced split either.  Both fall out
        of taking the row count first and then dividing the entries over it.

        A chart whose entries fit on one row returns ``(1, n)`` and is laid out by the gap
        rule in :data:`LEGEND_ENTRY_SLACK` instead, which this leaves untouched.
        """
        if not widths:
            return (1, 0)
        cap = LEGEND_BAND_MAX_FRACTION * self.frame.width
        if sum(widths) <= cap:
            return (1, len(widths))
        column = max(widths)
        fits = max(1, int(cap // column)) if column > 0 else 1
        rows = math.ceil(len(widths) / fits)
        return (rows, math.ceil(len(widths) / rows))

    def _legend_entry_widths(
        self, font: ChartFont, *, per_point: bool = False
    ) -> list[float]:
        """Each horizontal legend entry's width: its key cell plus its name."""
        cell = self._legend_key_cell(font)
        return [
            cell + font.width(name)
            for name in self._legend_names(per_point=per_point)
            if name
        ]

    def _legend_band_height(self, font: ChartFont, *, per_point: bool = False) -> float:
        """What a legend along the top or the bottom takes out of the frame.

        One row of :func:`legend_row_pitch` plus :data:`LEGEND_BAND_PAD_PT`, and one more
        pitch for every extra row less the 1.5 pt a wrapped band gives back.  Measured at
        five sizes against a no-legend control; see :data:`LEGEND_BAND_PAD_PT`.
        """
        rows, _ = self._legend_grid(self._legend_entry_widths(font, per_point=per_point))
        pad = LEGEND_BAND_PAD_PT - (LEGEND_WRAP_BAND_TRIM_PT if rows > 1 else 0.0)
        return rows * legend_row_pitch(font.box) + pad

    def _legend_side_metrics(
        self, font: ChartFont, *, per_point: bool = False
    ) -> tuple[float, float]:
        """A side legend's ``(band width, text column)``.

        The band is the widest **drawn line** plus the key and the two pads.  For a legend
        whose entries fit that is the widest name, which is the rule
        :data:`LEGEND_SIDE_MAX_FRACTION` was measured against; for one that does not, the
        cap fixes the column the names wrap inside and **the band then shrinks back to
        whatever the wrapped lines actually need**.  Measured on the ``legend-side`` deck:
        a three-word name on a 300 pt frame wraps to two words and one, and its band comes
        back 105.0 pt where the cap alone would reserve 120.  ``real-financial-report``'s
        doughnut cannot tell the two apart -- its wrapped Japanese line is 80.0 pt against
        a column of 80.04, so the shrunken band and the cap agree there to 0.04 pt, which
        is why this looked like a plain cap when it was first measured.

        Only the entries actually drawn count: a series struck out by ``c:legendEntry``
        would otherwise reserve width for a label nobody sees, shifting the plot
        rectangle.  A pie legends its *categories*, not its series.
        """
        names = [name for name in self._legend_names(per_point=per_point) if name]
        key, key_gap = self._legend_key_size(font)
        pads = key + key_gap + (LEGEND_SIDE_LEAD_EM + LEGEND_SIDE_TRAIL_EM) * font.size
        widest = max((font.width(name) for name in names), default=0.0)
        cap = self.frame.width * LEGEND_SIDE_MAX_FRACTION
        if widest + pads <= cap:
            return widest + pads, widest
        column = max(cap - pads, font.size)
        drawn = max(
            (
                font.width(line)
                for name in names
                for line in _legend_wrap(name, font, column)
            ),
            default=column,
        )
        return min(drawn + pads, cap), column

    def _legend_text_column(self, font: ChartFont, *, per_point: bool = False) -> float:
        """The width a **side** legend leaves for an entry's name.

        What an entry wraps inside once :data:`LEGEND_SIDE_MAX_FRACTION` has capped the
        band.  Measured on the ``legend-side`` deck's frame sweep: the same four-word name
        takes four lines on a 240 pt frame, two on a 360 pt one and one on a 480 pt one,
        which is this column filled greedily and is not what a share of the *frame*
        predicts.
        """
        return self._legend_side_metrics(font, per_point=per_point)[1]

    def _legend_side_width(self, font: ChartFont, *, per_point: bool = False) -> float:
        return self._legend_side_metrics(font, per_point=per_point)[0]

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

    # -- the 3-D scene ------------------------------------------------------------------

    @property
    def _draws_a_scene(self) -> bool:
        """Whether this group's solid is **drawn** rather than left as a flat rectangle.

        A narrower question than :meth:`_three_d_view`, which only asks whether the plot
        rectangle is the scene's front face.  What keeps a chart whose camera is applied
        from having its mesh drawn as well is a thing that would be a *wrong* picture
        rather than a simplified one:

        * **a ``pie3DChart``**, which is not a box at all -- and whose rim sweeps more than
          a hundred colours, so it is not the prism's faces with a different outline
          either.

        A ``surfaceChart`` always draws its scene when it has a camera, and draws its
        **lattice** into it only when it has two rows to stretch one between: a
        one-series surface is a scene with nothing in it, which is what PowerPoint draws
        as well -- see :meth:`_draw_scene_surface`.
        * **``c:grouping="standard"`` on a ``bar3DChart``**, which is the 3-D-only grouping
          that puts each series in its own row of depth rather than side by side across
          the width.  The flat fallback draws those side by side, so the mesh would stand
          them in the wrong place; nothing here has measured the bar's row layout.  It is
          the *default* for a ``line3DChart`` and an ``area3DChart``, whose rows are
          measured, and means nothing else there.
        * **``c:shape`` is not a box**, which only a ``bar3DChart`` states.  A cylinder, a
          cone or a pyramid is a different solid and drawing a box in its place is a
          different chart.

        A stacked ``area3DChart`` never gets here, because :func:`three_d_camera` gives it
        no scene to draw into.
        """
        if self.scene is None:
            return False
        if self.plot.kind in ("line3DChart", "area3DChart"):
            return True
        if self._is_surface:
            return True
        if self.plot.kind != "bar3DChart":
            return False
        if (self.plot.grouping or "clustered") == "standard":
            return False
        return (self.plot.shape or "box") == "box"

    def _scene_depth_span(self) -> tuple[float, float]:
        """Where in the scene's depth a bar stands, as two fractions of the whole.

        **Measured, and it is ``c:gapDepth`` and nothing else.**  A bar is as deep as it
        is wide (:func:`three_d_scene_depth`) and the scene is ``1 + gapDepth`` of that,
        so the bar fills ``1 / (1 + gapDepth)`` of its scene, centred.  ``view3d-mesh``
        reads the near gap and the solid's own depth off each prism's front and right
        faces at five gaps -- 0, 50, 150, 300 and 500% -- and the drawn share comes back
        0.985, 0.689, 0.406, 0.252 and 0.166 against this rule's 1.0, 0.667, 0.4, 0.25
        and 0.167.  At the default 150% that is the **0.3 to 0.7** ROADMAP.md 3.4 read off
        gallery slide 13 by hand, now with the law behind it rather than the one reading.
        """
        gap = max(
            (
                self.plot.gap_depth
                if self.plot.gap_depth is not None
                else DEFAULT_GAP_DEPTH
            )
            / 100.0,
            0.0,
        )
        near = gap / 2.0 / (1.0 + gap)
        return near, near + 1.0 / (1.0 + gap)

    def _scene_rows(self, count: int) -> list[tuple[float, float]]:
        """Where each series' solid stands in the depth, as fractions of the whole scene.

        A ``line3DChart``'s ribbons and an ``area3DChart``'s slabs take **one row of depth
        per series** where a clustered ``bar3DChart``'s share one, and inside its row each
        solid fills ``1 / (1 + gapDepth)`` centred -- the same gap rule as
        :meth:`_scene_depth_span`, applied to a row rather than to the whole scene.

        Measured on ``view3d-mesh``: the second and third rows' own near offsets come back
        48.65 and 86.33 pt against this rule's 48.98 and 86.66, and a ``c:gapDepth`` sweep
        of 0 and 500% divides the row without moving the scene at all.  The near offset is
        a reading of the rule on its own -- a three-category one-series probe paints its
        first category's front face at 5.641 value units where the value is 4.05 and the
        row's near offset lifts it by 1.586.

        **Series one stands nearest the viewer**, which is read off the raster rather than
        assumed: on the two-series slab probe the first series' fill is the one drawn over
        the second, and its front face is the lower of the two.
        """
        gap = max(
            (
                self.plot.gap_depth
                if self.plot.gap_depth is not None
                else DEFAULT_GAP_DEPTH
            )
            / 100.0,
            0.0,
        )
        rows = max(count, 1)
        span = 1.0 / rows
        near = span * gap / 2.0 / (1.0 + gap)
        deep = span / (1.0 + gap)
        return [
            (index * span + near, index * span + near + deep) for index in range(rows)
        ]

    def _draw_scene_ribbons(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """A ``line3DChart``'s ribbons and an ``area3DChart``'s slabs, back row first.

        Both are the same solid: a strip swept along the data in the front plane and
        extruded through its own row of depth.  What differs is where its underside is --
        the zero line for a slab, the value less :data:`VIEW_3D_RIBBON_THICKNESS` for a
        ribbon -- so both go through :meth:`_paint_slab`.

        The runs, the blanks and the polygon a slab closes to are the flat drawing's, which
        is deliberate: a 3-D area chart that breaks its run somewhere else than the flat
        one would be two bugs rather than one.  A run of a single point draws nothing, as
        flat.
        """
        scene = self.scene
        if scene is None or not series or not categories:
            return
        xs = self._category_positions(rect, len(categories))
        blanks = self.chart.display_blanks_as or "gap"
        zero = self._value_to_y(rect, 0.0, scale)
        rows = self._scene_rows(len(series))
        thickness = scene.face.width * VIEW_3D_RIBBON_THICKNESS
        ribbon = self._is_line
        # Back to front: the first series stands nearest the viewer, so it is painted
        # last and hides what stands behind it -- which is PowerPoint's picture.
        for order in reversed(range(len(series))):
            item = series[order]
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
                y = self._value_to_y(rect, value, scale)
                run.append((xs[index], y, y + thickness if ribbon else zero))
            for stretch in _split_runs(run):
                if len(stretch) < 2:
                    continue
                self._paint_slab(
                    [(x, top) for x, top, _ in stretch],
                    [(x, bottom) for x, _, bottom in stretch],
                    rows[order],
                    item.fill,
                    item.color,
                )

    def _value_bands(
        self, scale: tuple[float, float, float], axis: "c.SourceChartAxis | None" = None
    ) -> list[tuple[float, float, str]]:
        """The value bands a surface is coloured by: one per **major interval**.

        Measured on ``view3d-surfband``, whose ramps run the value linearly across
        thirteen categories so that every band the axis holds is in the picture and whose
        band legend names each one: the boundaries are the value axis' own major ticks,
        every time, over eleven axes from ``-20..20 by 10`` to ``0..1,4 by 0,2``.  There
        is no band count of its own and no rule of its own -- coarsen the axis and the
        bands coarsen with it, which is why a short frame draws five bands where a tall
        one draws nine on the same data.

        The label is the two ticks' own text joined by a hyphen, which is
        :meth:`_tick_texts` and therefore the axis' own number format: ``0,00-2,00`` under
        ``0.00``, ``0%-200%`` under ``0%``, and ``-4--2`` at the bottom of a signed axis,
        where PowerPoint joins the two strings and leaves the two signs where they fall.
        """
        texts = self._tick_texts(scale, axis)
        return [
            (first, second, f"{first_text}-{second_text}")
            for (first, first_text), (second, second_text) in zip(texts, texts[1:])
        ]

    def _band_fill(self, index: int, count: int) -> m.Fill | None:
        """What band *index* of *count* is painted, before the light reaches it.

        ``c:bandFmts`` first -- it numbers the bands from the axis' minimum up, measured
        on a probe stating band 0 red and band 2 green, which painted the lowest band and
        the third and left the ramp on the rest -- and otherwise the per-point accent
        cycle two colours early.  See :data:`BAND_COLOR_CYCLE_SLACK`.
        """
        stated = self.plot.band_fills.get(index)
        if stated is not None:
            fill = self._resolve_fill(stated)
            if fill is not None:
                return fill
        if not self.style.accents:
            return None
        return m.SolidFill(color=self._cycle_accent(index, count + BAND_COLOR_CYCLE_SLACK))

    def _band_color(self, index: int, count: int) -> m.ResolvedColor:
        """The band's flat colour, for its legend key and as the shading's base."""
        fill = self._band_fill(index, count)
        if isinstance(fill, m.SolidFill):
            return fill.color
        return m.ResolvedColor(hex="#4472C4")

    def _band_outline(self, index: int, count: int) -> m.Outline | None:
        """What a ``c:wireframe`` surface strokes band *index* with.

        The band's own ``a:ln`` if it states one, and otherwise the band's colour: the
        mesh deck's wireframe probe states a red *fill* per band through ``c:bandFmts``
        and draws its lattice in the **accent** ramp all the same, so the stroke follows
        the line and not the fill.  The colour is the band's flat one rather than a lit
        one -- a wireframe's strokes come back at the accent exactly, where every filled
        facet beside them is shaded.
        """
        stated = self.plot.band_outlines.get(index)
        if stated is not None:
            outline = self._resolve_outline(stated)
            if outline is not None:
                return outline
        return m.Outline(
            fill=m.SolidFill(color=self._band_color(index, count)),
            width=SURFACE_WIREFRAME_WIDTH_EMU,
        )

    def _band_series(
        self, scale: tuple[float, float, float], axis: "c.SourceChartAxis | None"
    ) -> list[_Series]:
        """A surface's legend entries: one per band, lowest first.

        **A band legend is a legend of value ranges and not of series**, which no other
        chart type here has, and it is drawn through the ordinary legend all the same:
        every entry is a name and a key, so a band is a ``_Series`` with the band's name
        and the band's fill and nothing else in it.

        The order is measured at all four positions: a legend at the **side** runs the
        highest band at the **top**, and one along the bottom or the top runs the lowest
        at the **left** -- which is one order, bands ascending, stacked upwards in the
        first case and rightwards in the second.  The caller reverses for a side legend;
        this returns them ascending.

        A ``c:wireframe`` surface keys with an **unfilled square outlined in the band's
        colour**, measured on the wireframe legend probe, where the five keys are the
        same 5.49 pt squares with a stroke and no fill at all.
        """
        bands = self._value_bands(scale, axis)
        wire = bool(self.plot.wireframe)
        return [
            _Series(
                name=name,
                values=[],
                color=self._band_color(index, len(bands)),
                fill=None if wire else self._band_fill(index, len(bands)),
                outline=self._band_outline(index, len(bands)) if wire else None,
                format_code=None,
                invert_if_negative=False,
            )
            for index, (_low, _high, name) in enumerate(bands)
        ]

    def _surface_view(self) -> tuple[float, float, float]:
        """The direction the scene is looked **from**, in the scene's own frame.

        ``rAngAx="1"`` is an oblique projection, so one direction projects to nothing and
        that direction is the view: solving ``x + dx*z`` and ``y + dy*z`` for the
        displacement that moves neither gives ``(dx / D, -dy / D, 1)``, where ``(dx, dy)``
        is the scene's drawn depth vector and ``D`` its depth in the scene's own units.
        It decides which side of a facet is the one drawn and, with it, the painter's
        order -- both of which a surface needs and a prism never did, its faces being
        axis-aligned.

        Read back on the probes, it is the camera: ``(sin rotY, sin rotX, 1)``.  The sheet
        that falls away at exactly the pitch is the check -- ``sm-c5-n2-down`` tilts its
        two rows 15.0 degrees at ``rotX=15`` and PowerPoint draws its **underside**, at
        the ambient, which is the edge-on case this predicts to a hundredth.
        """
        scene = self.scene
        if scene is None:
            return (0.0, 0.0, 1.0)
        depth = scene.span * scene.face.width
        if depth <= 0:
            return (0.0, 0.0, 1.0)
        return (scene.depth[0] / depth, -scene.depth[1] / depth, 1.0)

    def _draw_scene_surface(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """A ``surfaceChart``'s lattice: quads over (category, series, value), lit and
        banded.

        **The lattice.**  A point sits on the category tick (`c:crossBetween` is `midCat`
        by default here, measured) and at ``row / (rows - 1)`` of the scene's depth --
        series one at the front plane exactly, the last series on the back wall exactly,
        the rest spread evenly between.  Measured on ``view3d-surfmesh``, whose every band
        is painted one red so the sheet is separable from the floor by colour alone and
        whose rows are flat at a value each, so a row's own scanline gives its front-left
        corner: two, three, four and five rows come back at ``0, 1``, ``0, 1/2, 1``,
        ``0, 1/3, 2/3, 1`` and ``0, 1/4, 1/2, 3/4, 1`` of a depth that is itself
        ``series / categories`` of the face's width.  ``c:gapDepth`` does nothing at all
        here -- 0 and 500% draw the identical picture -- where it divides a ribbon's row,
        and ``c:depthPercent`` scales the whole depth with the rows still at its ends.
        **Series one is the front row** whatever its values: the descending-rows probe
        puts its highest row at the front and its lowest at the back.

        **The facets.**  Each cell of the lattice is two triangles split along the
        diagonal from its near-left corner to its far-right one, which is measured and not
        chosen: a cell with three corners level and the fourth pulled down draws **two**
        tones, one of them the level tone its neighbour draws, which is what that diagonal
        gives and what the other one cannot -- it would split the same cell into two
        sloped triangles and draw three tones.  Predicted, the two tones land within one
        8-bit level at three heights.

        **The bands.**  A facet is not one colour: the value varies across it and the band
        boundary cuts through it, so each triangle is clipped by the two planes of every
        band it reaches and each piece drawn in that band's own fill.  The cut is straight
        because the triangle is planar, which is what the picture shows -- a contour line
        running clean across a cell.

        **The order** is the painter's, along the view direction
        (:meth:`_surface_view`): a facet whose centroid sits further from the viewer is
        painted first.  That is a true depth sort and not the two-key screen sort a prism
        gets, because a surface folds -- a near row can stand behind a far one wherever
        the sheet climbs.
        """
        scene = self.scene
        if scene is None or len(series) < 2 or len(categories) < 2:
            # One row stretches no sheet, and PowerPoint draws none: its one-series probes
            # come back an empty scene with a degenerate axis.  The scene, its floor and
            # its walls are still drawn here, which that degenerate axis is not.
            return
        xs = self._category_positions(rect, len(categories))
        blanks = self.chart.display_blanks_as or "gap"
        bands = self._value_bands(scale)
        rows = len(series)
        depth = scene.span * scene.face.width
        view = self._surface_view()
        wire = bool(self.plot.wireframe)

        def value_at(row: int, col: int) -> float | None:
            value = _at(series[row].values, col)
            if value is None and blanks == "zero":
                return 0.0
            return value

        # Each entry is one band's piece of one triangle: its screen polygon, which
        # band it belongs to, the light on the whole triangle, and the depth key the
        # painter's order runs on.
        faces: list[tuple[list[tuple[float, float]], int, float, float]] = []
        for row in range(rows - 1):
            for col in range(len(categories) - 1):
                corners = [
                    (row, col), (row, col + 1), (row + 1, col + 1), (row + 1, col)
                ]
                values = [value_at(r, c) for r, c in corners]
                if any(value is None for value in values):
                    # A gap in the data takes the whole cell with it: a lattice cell needs
                    # all four of its corners and there is nothing to interpolate from.
                    continue
                points = [
                    (
                        xs[c],
                        self._value_to_y(rect, value, scale),
                        z / (rows - 1),
                    )
                    for (r, c), value, z in zip(
                        corners, values, (row, row, row + 1, row + 1)
                    )
                ]
                # The diagonal runs corner 0 to corner 2 -- near-left to far-right.
                for triangle in ((0, 1, 2), (0, 2, 3)):
                    corner_points = [points[index] for index in triangle]
                    corner_values = [values[index] for index in triangle]
                    normal = _surface_normal(corner_points, depth, view)
                    if normal is None:
                        continue
                    shade = three_d_lambert(normal, VIEW_3D_SURFACE_LIGHT)
                    key = sum(
                        point[0] * view[0]
                        - point[1] * view[1]
                        - point[2] * depth * view[2]
                        for point in corner_points
                    ) / 3.0
                    for index, (low, high, _name) in enumerate(bands):
                        piece = _clip_to_band(corner_points, corner_values, low, high)
                        if len(piece) < 3:
                            continue
                        faces.append(
                            ([scene.at(*point) for point in piece], index, shade, key)
                        )
        for points, index, shade, _key in sorted(faces, key=lambda row: row[3]):
            if wire:
                self._polygon(
                    points, fill=None, outline=self._band_outline(index, len(bands))
                )
                continue
            fill = self._shade_factor(
                self._band_fill(index, len(bands)),
                shade,
                self._band_color(index, len(bands)),
            )
            self._polygon(
                points,
                fill=fill,
                # The seam, not a line of PowerPoint's: see `SURFACE_SEAM_WIDTH_EMU`.
                outline=(
                    m.Outline(fill=fill, width=SURFACE_SEAM_WIDTH_EMU)
                    if isinstance(fill, m.SolidFill)
                    else None
                ),
            )

    def _paint_slab(
        self,
        top: list[tuple[float, float]],
        bottom: list[tuple[float, float]],
        span: tuple[float, float],
        fill: m.Fill | None,
        base: m.ResolvedColor,
    ) -> None:
        """One ribbon or slab: its two ends, its skin per segment, and its front face.

        Every face but the front one is **lit rather than looked up**: a segment's top
        surface is tilted by its own slope, so the four :data:`VIEW_3D_FACE_SHADES` are
        not enough and :func:`three_d_lambert` is what paints it.  That is the model the
        prism's four constants are four samples of, and the ribbon is what confirms it --
        see :data:`VIEW_3D_FACE_SHADES`.

        **Which faces are drawn is the camera's arithmetic and not a case to enumerate.**
        A face is visible when its outward normal points towards the viewer, and for an
        oblique projection that test is ``n . (-depth_x, depth_y, -depth) < 0`` -- so a
        segment rising more steeply than the depth leads shows its underside instead of
        its top, which is what PowerPoint draws: at ``rotX=0`` a rising segment's belly
        comes back at the ambient 0.3939 where its falling neighbour's top is 0.7908.

        The front face is at the near edge and is painted last, over everything the solid
        has behind it; between segments the paint runs the way the depth points, which is
        the rule :meth:`_paint_prisms` measured.
        """
        scene = self.scene
        if scene is None or len(top) < 2:
            return
        near, far = span
        faces: list[tuple[float, list[tuple[float, float]], tuple[float, float, float]]] = []

        def visible(normal: tuple[float, float, float]) -> bool:
            return -normal[0] * scene.depth[0] + normal[1] * scene.depth[1] < 0

        def quad(a, b, normal) -> None:
            if not visible(normal):
                return
            faces.append(
                (
                    (a[0] + b[0]) / 2.0 * math.copysign(1.0, scene.depth[0]),
                    [
                        scene.at(*a, near),
                        scene.at(*b, near),
                        scene.at(*b, far),
                        scene.at(*a, far),
                    ],
                    normal,
                )
            )

        # The two ends, whose normals are the scene's own x axis -- and which therefore
        # come out at the prism's `left` and `right` constants exactly.
        quad(top[0], bottom[0], (-1.0, 0.0, 0.0))
        quad(top[-1], bottom[-1], (1.0, 0.0, 0.0))
        for index in range(len(top) - 1):
            for edge in (top, bottom):
                first, second = edge[index], edge[index + 1]
                run = second[0] - first[0]
                rise = second[1] - first[1]
                length = math.hypot(run, rise)
                if length <= 0:
                    continue
                # Screen y runs down, so the surface's upward normal is (dy, dx) in it.
                normal = (rise / length, run / length, 0.0)
                quad(first, second, normal if edge is top else tuple(-v for v in normal))
        for _, points, normal in sorted(faces, key=lambda row: row[0]):
            self._polygon(
                points, fill=self._shade_normal(fill, normal, base), outline=None
            )
        self._polygon(
            [scene.at(*point, near) for point in top]
            + [scene.at(*point, near) for point in reversed(bottom)],
            fill=fill,
            outline=None,
        )

    def _shade_normal(
        self, fill: m.Fill | None, normal: tuple[float, float, float], base: m.ResolvedColor
    ) -> m.Fill | None:
        """:meth:`_shade`, given a normal rather than one of the four named faces."""
        return self._shade_factor(fill, three_d_lambert(normal), base)

    def _draw_scene_gridlines(
        self, rect: _Rect, scale: tuple[float, float, float], axis: c.SourceChartAxis | None
    ) -> None:
        """A value tick's gridline in a scene: **up the side wall, then across the back**.

        Read off gallery slide 13's raster: each gridline is a polyline of two segments,
        front-left to back-left along the depth vector and then horizontally across the
        back wall, rather than the single line across the plot a flat chart draws.  The
        one at the crossing is left out here exactly as it is when flat, because the floor
        (:meth:`_draw_scene_floor`) draws it along with the rest of the floor's outline.
        """
        scene = self.scene
        if scene is None or axis is None or not axis.major_gridlines:
            return
        outline = self._axis_outline(axis.major_gridline_outline)
        if outline is None:
            return
        # The side wall is the one the depth leads *away* from, so a yaw past a half turn
        # puts it on the right and the back wall's run the other way.
        near_x = rect.left if scene.depth[0] >= 0 else rect.right
        far_x = rect.right if scene.depth[0] >= 0 else rect.left
        near_y = rect.bottom if scene.depth[1] <= 0 else rect.top
        far_y = rect.top if scene.depth[1] <= 0 else rect.bottom
        horizontal = (self.plot.bar_direction or "col") == "bar"
        for value in self._tick_values(scale):
            # **The one at the crossing is drawn**, where a flat chart leaves it out: flat,
            # the category axis already draws that line, and in a scene the category axis
            # is the floor's *front* edge while this is its back one.  Measured on gallery
            # slide 13, where the floor's back edge and its left side are the same black
            # as the gridlines above them and only the fourth side is the floor's own grey.
            if horizontal:
                # `barDir="bar"` turns the gridlines across the value axis into the other
                # pair of walls: back along the **floor** and then up the back wall, which
                # is the same two segments read in the scene's other plane.
                x = self._value_to_x(rect, value, scale)
                points = [(x, near_y), scene.at(x, near_y, 1.0), scene.at(x, far_y, 1.0)]
            else:
                y = self._value_to_y(rect, value, scale)
                points = [(near_x, y), scene.at(near_x, y, 1.0), scene.at(far_x, y, 1.0)]
            self._polyline(points, outline, smooth=False)

    def _draw_scene_floor(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        axis: c.SourceChartAxis | None,
    ) -> None:
        """The floor, as the quadrilateral it is: four sides in its own grey.

        Three of the four are drawn over by something black -- the category axis along the
        front, the value axis' gridline at zero back along the depth and across the back --
        and the fourth is where the floor's own stroke shows.  It is drawn whole rather
        than as that one side, because which side is left over is a property of the camera
        and of whether the axis has gridlines at all, not something to hard-code: see
        :data:`VIEW_3D_FLOOR_COLOR`, measured at ``#898989`` on two decks with different
        themes.
        """
        scene = self.scene
        if scene is None or axis is None or axis.delete:
            return
        crossing = self._category_axis_position(rect, scale, axis)
        # The floor is the plane the categories stand on, so a horizontal bar chart -- whose
        # categories run up the side -- stands them on the *left* wall instead, and what
        # this outlines is that wall.
        if (self.plot.bar_direction or "col") == "bar":
            near, far = (crossing, rect.top), (crossing, rect.bottom)
        else:
            near, far = (rect.left, crossing), (rect.right, crossing)
        self._polygon(
            [
                near,
                scene.at(*near, 1.0),
                scene.at(*far, 1.0),
                far,
            ],
            fill=m.NoFill(),
            outline=m.Outline(
                width=DEFAULT_AXIS_LINE_EMU,
                fill=m.SolidFill(color=m.ResolvedColor(hex=VIEW_3D_FLOOR_COLOR)),
            ),
        )

    def _shade(self, fill: m.Fill | None, face: str, base: m.ResolvedColor) -> m.Fill | None:
        """One face's paint: the series' own fill, multiplied per sRGB channel.

        The factors are :data:`VIEW_3D_FACE_SHADES` and they are **constants** -- measured
        over 42 cameras, which is what refutes reading them as a Lambert term against the
        drawn normal.  A fill that is not a flat colour has no channel to multiply, so the
        front face keeps it and the shaded ones fall back to the series' resolved colour,
        which is the same colour its legend swatch draws.
        """
        return self._shade_factor(fill, VIEW_3D_FACE_SHADES.get(face, 1.0), base)

    def _shade_factor(
        self, fill: m.Fill | None, factor: float, base: m.ResolvedColor
    ) -> m.Fill | None:
        """*fill* multiplied by *factor* per sRGB channel, which is where the shade lands."""
        if factor >= 1.0:
            return fill
        color = fill.color if isinstance(fill, m.SolidFill) else base
        channels = tuple(
            min(255, max(0, round(int(color.hex[index : index + 2], 16) * factor)))
            for index in (1, 3, 5)
        )
        return m.SolidFill(
            color=m.ResolvedColor(
                hex="#%02X%02X%02X" % channels,
                alpha=color.alpha,
            )
        )

    def _paint_prism(
        self,
        box: _Rect,
        fill: m.Fill | None,
        outline: m.Outline | None,
        base: m.ResolvedColor,
        span: tuple[float, float],
    ) -> None:
        """One solid, as its **three visible faces** and no more.

        A box seen through an oblique projection shows exactly three of its six faces --
        the front, one cap and one side -- and which cap and which side is the sign of the
        depth vector's two components rather than a case to enumerate: the depth leading
        up shows the top and leading down shows the bottom, the same for right and left.
        A degenerate camera zeroes one component and the face it would have shown collapses
        to nothing, which is what PowerPoint draws too -- there is no top at ``rotX=0`` and
        no side at ``rotY=0``.

        The faces meet with no stroke between them, which is measured: a colour census of
        slide 13's scene finds the two series' fills, two shades of each, and nothing in
        between.
        """
        scene = self.scene
        if scene is None:
            return
        near, far = span
        cap_y, cap_face = (box.top, "top") if scene.depth[1] < 0 else (box.bottom, "bottom")
        side_x, side_face = (
            (box.right, "right") if scene.depth[0] > 0 else (box.left, "left")
        )
        faces = [
            (
                side_face,
                [
                    scene.at(side_x, box.top, near),
                    scene.at(side_x, box.top, far),
                    scene.at(side_x, box.bottom, far),
                    scene.at(side_x, box.bottom, near),
                ],
            ),
            (
                cap_face,
                [
                    scene.at(box.left, cap_y, near),
                    scene.at(box.right, cap_y, near),
                    scene.at(box.right, cap_y, far),
                    scene.at(box.left, cap_y, far),
                ],
            ),
            (
                "front",
                [
                    scene.at(box.left, box.top, near),
                    scene.at(box.right, box.top, near),
                    scene.at(box.right, box.bottom, near),
                    scene.at(box.left, box.bottom, near),
                ],
            ),
        ]
        for face, points in faces:
            if abs(points[0][0] - points[2][0]) < 1e-9 or abs(points[0][1] - points[2][1]) < 1e-9:
                # A face the camera has turned edge-on, which draws nothing at all.
                continue
            self._polygon(points, fill=self._shade(fill, face, base), outline=outline)

    def _paint_prisms(self) -> None:
        """Every prism this group collected, painted **back to front**.

        SVG's paint model is the painter's algorithm, so the depth sort is the whole of
        the hidden-surface problem here -- but a sort on average depth is not it, because
        every bar of a clustered ``bar3DChart`` stands in the *same* row of depth and they
        still hide each other.  What decides the order between two solids at one depth is
        the direction the depth vector points on screen: a bar's side face runs that way,
        so anything further along it is nearer the viewer and has to be painted over it.
        The same argument settles a stacked column, whose segments share a position and
        differ only up the value axis and whose cap runs the same way.

        Measured rather than reasoned: on gallery slide 13 a bar's right face is drawn
        only where it is taller than its right-hand neighbour -- Q1 and Q4 for the blue
        series, which are exactly the two categories where it outruns the orange -- and
        the top face of a shorter bar stops dead at the next bar's left edge.  That is
        left-to-right painting and nothing else.
        """
        prisms, self._prisms = self._prisms, None
        if not prisms:
            return
        span = self._scene_depth_span()
        for _, box, fill, outline, base in sorted(prisms, key=lambda row: row[0]):
            self._paint_prism(box, fill, outline, base, span)

    def _prism_order(self, box: _Rect) -> tuple[float, float, float]:
        """Where one solid sits in the paint order: furthest first.

        The three keys are the three ways a box can be behind another one -- deeper into
        the scene, and then along each of the depth vector's two screen components.  The
        first is always zero here because a clustered ``bar3DChart`` shares one row of
        depth; it is written down because the sort is the thing most likely to be wrong
        when this grows a group element that does not.
        """
        scene = self.scene
        if scene is None:
            return (0.0, 0.0, 0.0)
        return (
            0.0,
            (box.left + box.right) / 2.0 * math.copysign(1.0, scene.depth[0]),
            (box.top + box.bottom) / 2.0 * math.copysign(1.0, scene.depth[1]),
        )

    def _draw_gridlines(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        axis: c.SourceChartAxis | None,
        *,
        skip_crossing: bool = True,
    ) -> None:
        """``skip_crossing`` leaves out the gridline the category axis already draws.

        True for the primary axis and **False for a secondary one**, which is measured:
        on ``p-secgrid`` the secondary axis drew eleven gridlines for its eleven ticks
        where the primary drew ten for the same eleven, the missing one being the
        category axis' own line.
        """
        if axis is None or not axis.major_gridlines:
            return
        outline = self._axis_outline(axis.major_gridline_outline)
        horizontal = (self.plot.bar_direction or "col") == "bar"
        # A gridline runs *across* the value axis, so `barDir="bar"` turns them vertical.
        # Measured on the horizontal probe: three vertical lines at the ticks for 2, 4 and
        # 6, with the one for 0 left out because the category axis already draws it.
        axis_position = (
            self._category_axis_position(rect, scale, None) if skip_crossing else None
        )
        for value in self._tick_values(scale):
            if horizontal:
                x = self._value_to_x(rect, value, scale)
                if axis_position is not None and abs(x - axis_position) < 0.01:
                    continue
                self._line(x, rect.top, x, rect.bottom, outline)
            else:
                y = self._value_to_y(rect, value, scale)
                if axis_position is not None and abs(y - axis_position) < 0.01:
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
        overlap = self.plot.overlap if self.plot.overlap is not None else 0.0
        # `gapWidth` is the gap between category groups expressed as a percentage of one
        # bar's width, so the band holds `slots` bars plus `gapWidth/100` of one more --
        # **less what `c:overlap` takes back**, because overlapping bars share their
        # neighbours' width and the cluster gets wider bars in the same band.
        #
        # The overlap term was missing here and two charts say so, both to inside the
        # 0.24 pt PowerPoint quantises bar widths to: ``chart-gallery`` slide 1 (five
        # categories, two series, gapWidth 150, overlap -27 on a 250.59 pt plot) drew
        # 13.2 pt bars where the divisor without it asks for 14.32 and with it for 13.29,
        # and the ``combo-bar`` probe ``g-overlap`` (the same settings on a 427.18 pt
        # plot) drew 22.56 against 24.41 and 22.66.  A stacked group is unaffected --
        # ``slots`` is 1, so the term is zero -- which is why
        # ``real-college-template``'s ``overlap=100`` chart does not move.
        #
        # The schema bounds `gapWidth` to 0..500 and a file is free to ignore that; at
        # -100 on a single series the divisor is exactly zero, which used to abort the
        # conversion, and an overlap past 100 can drive it negative the same way.
        slack = slots + gap_width / 100.0 - (slots - 1) * overlap / 100.0
        bar_size = band / max(slack, MIN_BAR_SLOTS)
        step = bar_size * (1.0 - overlap / 100.0)
        cluster = bar_size + step * (slots - 1)

        percent = grouping == "percentStacked"
        totals = _percent_totals(series) if percent else None
        if self._draws_a_scene:
            self._prisms = []

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
        self._paint_prisms()

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

    def _draw_hi_low_lines(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """``c:hiLowLines`` -- the vertical range at each category.

        Measured: the probe's line runs from the largest value in the category to the
        smallest -- 18 down to 8 on a 0..20 axis, drawn at 25.606 and 98.123 pt, both on
        the axis to 0.001 pt -- in black at 0.5 pt, which is the axis default, and in an
        explicit ``<a:ln w="28575">`` red at 2.25 pt when the file states one.  The
        element's *presence* is the switch: with it absent nothing is drawn.

        Only well-formed data was measured, so "the largest and smallest of every series"
        and "the second and third series" are not separated by any probe here; the former
        is what is implemented, because it is the one that cannot pick the wrong pair when
        the series are ordered differently.
        """
        if self.plot.hi_low_lines is None or len(series) < 2 or not categories:
            return
        outline = self._axis_outline(self.plot.hi_low_lines.outline)
        if outline is None:
            return
        xs = self._category_positions(rect, len(categories))
        for index in range(len(categories)):
            values = [
                value
                for item in series
                if (value := _at(item.values, index)) is not None
            ]
            if len(values) < 2:
                continue
            top = self._value_to_y(rect, max(values), scale)
            bottom = self._value_to_y(rect, min(values), scale)
            if abs(top - bottom) < 1e-6:
                continue
            self._line(xs[index], top, xs[index], bottom, outline)

    def _draw_up_down_bars(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """``c:upDownBars`` -- the body between the first and last series.

        **The series order carries the meaning and the labels carry none**, which is the
        measurement the brief asked for: a three-series High/Low/Close chart with
        ``c:upDownBars`` drew all five bars *down*, from each category's High to its
        Close, so the pair is ``series[0]`` and ``series[-1]`` and not "open and close by
        name".  Four series drew three up bars and two down, which is where close > open
        and where it does not.

        The width is ``band / (1 + gapWidth/100)`` with the default
        :data:`DEFAULT_UP_DOWN_GAP_WIDTH`, measured at 50, 150 and 300, and the fills are
        :data:`DEFAULT_UP_BAR_FILL` / :data:`DEFAULT_DOWN_BAR_FILL` with a black 0.5 pt
        outline on both.
        """
        bars = self.plot.up_down_bars
        if bars is None or len(series) < 2 or not categories:
            return
        gap = bars.gap_width if bars.gap_width is not None else DEFAULT_UP_DOWN_GAP_WIDTH
        band = rect.width / max(len(categories), 1)
        width = band / max(1.0 + max(gap, 0.0) / 100.0, MIN_BAR_SLOTS)
        xs = self._category_positions(rect, len(categories))

        default_outline = m.Outline(
            width=DEFAULT_AXIS_LINE_EMU,
            fill=m.SolidFill(color=m.ResolvedColor(hex=DEFAULT_AXIS_COLOR)),
        )
        up_fill = self._resolve_fill(bars.up_fill) or m.SolidFill(
            color=m.ResolvedColor(hex=DEFAULT_UP_BAR_FILL)
        )
        down_fill = self._resolve_fill(bars.down_fill) or m.SolidFill(
            color=m.ResolvedColor(hex=DEFAULT_DOWN_BAR_FILL)
        )
        up_outline = self._resolve_outline(bars.up_outline) or default_outline
        down_outline = self._resolve_outline(bars.down_outline) or default_outline

        first, last = series[0], series[-1]
        for index in range(len(categories)):
            opening = _at(first.values, index)
            closing = _at(last.values, index)
            if opening is None or closing is None or opening == closing:
                continue
            rising = closing > opening
            top = self._value_to_y(rect, max(opening, closing), scale)
            bottom = self._value_to_y(rect, min(opening, closing), scale)
            self._rect(
                _Rect(xs[index] - width / 2, top, xs[index] + width / 2, bottom),
                fill=up_fill if rising else down_fill,
                outline=up_outline if rising else down_outline,
            )

    def _cross_between(self) -> str:
        """``between`` puts a point at its band's centre, ``midCat`` on the band edge.

        See :data:`DEFAULT_AREA_CROSS_BETWEEN` for why an area's default is the other one.
        """
        axis = self._axis_for(1) or self._axis_of_kind("valAx")
        stated = axis.cross_between if axis is not None else None
        if stated:
            return stated
        # **A surface's default is `midCat` too**, and it is measured rather than taken
        # from the area beside it: the probe with the attribute left out altogether draws
        # the same face, the same lattice and the same depth as its `midCat` twin to the
        # digit, where its `between` twin draws a face a category wider.
        return (
            DEFAULT_AREA_CROSS_BETWEEN
            if (self._is_area or self._is_surface)
            else "between"
        )

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
        return (
            self._is_area or self._is_line or self._is_surface
        ) and self._cross_between() == "midCat"

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
        if self._prisms is not None:
            # Held back rather than drawn: a 3-D bar is a solid, and which solid hides
            # which is a question about all of them at once.  See :meth:`_paint_prisms`.
            self._prisms.append(
                (self._prism_order(box), box, fill, outline, item.color)
            )
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
                    # The same allowance `_bottom_label_band` sized the band with, so a
                    # label is drawn exactly as long as the space set aside for it.
                    allowance = self._rotated_allowance(category_font.box_for(*categories))
                    self._rotated_labels_along_bottom(
                        [
                            (
                                rect.left + (index + 0.5) * band,
                                truncate_label(text, category_font, allowance),
                            )
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
        :func:`rotated_label_anchor` below the axis.  The renderer rotates a shape about
        its own centre, so the box is placed
        by working that rotation backwards from the anchor rather than by rotating the
        text in place -- which is why the arithmetic below is not simply "left = x".
        """
        # The *drawn* face's box, which is the one the band was built from too.
        box = font.box_for(*(text for _, text in labels))
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
            #
            # **A cut label is aligned on its prefix, not on its ellipsis.**  In
            # `real-financial-report.pptx`'s export the ellipsis' own pen starts where
            # `プラット` ends, which is the anchor, so the ellipsis hangs past it -- the
            # same reason `_bottom_label_band` sizes the band from the prefix.
            overhang = (
                font.width(LABEL_ELLIPSIS) if text.endswith(LABEL_ELLIPSIS) else 0.0
            )
            offset_x = width / 2 - box.size / 2 - overhang
            offset_y = box.first_baseline - height / 2
            turned_x = offset_x * cos - offset_y * sin
            turned_y = offset_x * sin + offset_y * cos
            centre_x = position + ROTATED_LABEL_OFFSET_X_PT - turned_x
            centre_y = axis_y + rotated_label_anchor(box) - turned_y
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

    def _labels_down_right(
        self,
        rect: _Rect,
        labels: list[tuple[float, str]],
        font: ChartFont,
    ) -> None:
        """A secondary value axis' labels: left-aligned in the column right of the plot.

        The mirror of :meth:`_labels_down_left`, and measured to be exactly that.  On
        ``combo-plot`` the right-hand column began 9.70 pt past the plot's right edge at
        three label widths -- 481.29 against a plot ending at 471.59, 449.25 against
        439.55, 465.27 against 455.57 -- and the left-hand column on the same slides ended
        9.70 pt before the plot began.  One gap, two sides.
        """
        box = font.box
        left = rect.right + box.descent + VALUE_LABEL_GAP_EM * box.size
        width = max(self.frame.right - left, box.size)
        for y, text in labels:
            if not text:
                continue
            self._text(
                self._label_body(text, font, align="l"),
                left=left,
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
        lower edge and the labels follow it.  The drop is ``(5/3) * ascent`` and it is
        **read off the drawn baselines**, not inferred from the band: ``axis-inset``'s
        category labels have no descender, so a PDF text rect's own floor is the baseline,
        and 24 of them in four faces give 1.672 +/- 0.04 of the ascent.  Against the em
        the same readings run 1.38 to 1.64 and are no rule at all, which is what says the
        drop follows the face rather than the size -- and it is the same number as hanging
        the line's descender one frame padding above the frame whenever the axis *is* at
        the foot, so this and :meth:`_bottom_label_band` stay one measurement.

        A label too wide for its band is broken across lines, each one centred in the band
        under the one above.  **The block is top-aligned**, so the first baseline is where
        it would be for a one-line label however many lines follow: the Arial ladder put
        it 15.03 to 15.07 pt under the axis at one, two and three lines, and a chart whose
        one long label wrapped left every short label sitting on the *first* line with
        nothing beneath it.
        """
        box = font.box
        baseline = axis_y + box.ascent + CATEGORY_LABEL_GAP_ASCENT * box.ascent
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
                    # The rows advance by the same pitch the band is reserved in -- see
                    # `_bottom_label_band`.  They have to: the band is sized to hold
                    # exactly these rows, so spacing them by the line box instead would
                    # leave an Arial block sitting 0.33 pt a line above its own floor.
                    baseline=baseline + index * box.pitch,
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
        bubble_size: float | None = None,
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
        if labels.show_bubble_size and bubble_size is not None:
            parts.append(format_number(bubble_size, labels.number_format))
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
        *,
        side: float | None = None,
        edge: float | None = None,
    ) -> None:
        """``side`` and ``edge`` override the two gaps a bubble measures differently."""
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
            baseline = y - (DATA_LABEL_GAP_PT if edge is None else edge) - box.descent
            left, align = x - width / 2, "ctr"
        elif placement == "inside-y":
            baseline = y + DATA_LABEL_INNER_GAP_PT + box.ascent
            left, align = x - width / 2, "ctr"
        elif placement == "inside-base-y":
            baseline = y - DATA_LABEL_GAP_PT - box.descent
            left, align = x - width / 2, "ctr"
        elif placement == "below":
            # A scatter's `b`: the mirror of `t` about the point, measured 0.70 pt loose.
            baseline = y + (DATA_LABEL_GAP_PT if edge is None else edge) + box.ascent
            left, align = x - width / 2, "ctr"
        elif placement == "right":
            baseline = y + box.ink_centre
            left = x + (DATA_LABEL_LINE_GAP_EM * box.size if side is None else side)
            align = "l"
        elif placement == "left":
            # A scatter's `l`: the mirror of `r`, the same 0.6 em off the marker's edge.
            baseline = y + box.ink_centre
            left = x - (DATA_LABEL_LINE_GAP_EM * box.size if side is None else side) - width
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
        overlap = self.plot.overlap if self.plot.overlap is not None else 0.0
        # The same divisor `_draw_bars` uses, overlap term and all; the two drift apart
        # into a label that no longer sits on its bar if either changes alone.
        size = band / max(
            slots + gap_width / 100.0 - (slots - 1) * overlap / 100.0, MIN_BAR_SLOTS
        )
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
        # Every drawn group, in **paint** order, which is the order the entries come out
        # in: measured on ``combo-legend``, where a line group written *first* still
        # legends after the bar group written second, in the same place and with the same
        # widths as the deck that writes them the other way round.
        if self._is_surface:
            # A surface legends its **value bands**.  The scale they come from is the one
            # `_build_cartesian` solved, stashed there before anything reads this; the
            # provisional axis stands in for the one call that comes *before* that -- the
            # band a legend along the top or the bottom reserves, which is measured from
            # the axis and so cannot wait for it.  The two differ only in how coarse the
            # unit is, and what is being measured here is the width of `0-2` against
            # `0-10`.
            return [name for _low, _high, name in self._value_bands(self._band_scale())]
        sources = [source for plot in self._drawn_plots() for source in plot.series]
        if not per_point:
            return [
                source.name.plain or ""
                for index, source in enumerate(sources)
                if source.name is not None and index not in deleted
            ]
        for source in sources:
            if any(source.categories):
                return [
                    name
                    for index, name in enumerate(source.categories)
                    if name and index not in deleted
                ]
        return []

    def _band_scale(self) -> tuple[float, float, float]:
        """The value scale this chart's bands are cut from.

        Solved once by :meth:`_build_cartesian` and stashed, because the bands are the
        value axis' own intervals and every reader of them wants the same axis the chart
        drew.  Before that -- there is exactly one such caller, the band a top or bottom
        legend reserves, which the axis' own interval count depends on in turn -- the
        finest axis the data could take stands in for it, which is the same stand-in
        :meth:`_value_axis_intervals` makes for the same circle.
        """
        if self.bands_scale is not None:
            return self.bands_scale
        numbers = self._axis_reach(self._series()) or [0.0]
        return nice_axis_scale(
            min(numbers), max(numbers), intervals=AXIS_MAX_INTERVALS, strict=False
        )

    def _draw_legend(
        self,
        rect: _Rect,
        series: list[_Series],
        *,
        per_point: bool = False,
        categories: list[str] | None = None,
        second_band: float = 0.0,
    ) -> None:
        """``second_band`` is what a secondary value axis takes beyond the plain edge
        inset, which a legend at the **side** has to clear.  Measured once, on
        ``combo-legend``'s ``l-right``: the legend key landed 430.14 pt from the frame's
        left edge, 0.04 pt from where it lands with no secondary axis at all, while the
        plot itself gave up a further 15.1 pt.  So the band moves the plot and not the
        legend, and the lead gap is measured from the far side of the label column."""
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
            # The layout advances by a **key cell** and the name; the drawn key is centred
            # in the cell, so the swatch is inset by half the difference.  See
            # :data:`LEGEND_ENTRY_SLACK` for where the gap comes from and what measured it.
            cell = self._legend_key_cell(font)
            widths = [cell + font.width(item.name or "") for _, item in entries]
            total = sum(widths)
            rows, columns = self._legend_grid(widths)
            pitch = legend_row_pitch(box)
            # Every row's first baseline sits the same distance below its row's top as a
            # side legend's does -- the row's own line, optically centred.  What differs is
            # which edge the block is anchored to, and the two are not symmetric: a bottom
            # legend hangs its **last** row off the frame's bottom and a top legend pins
            # its **first** row below the frame's top.  Both are measured at five sizes on
            # the ``legend-band`` deck and at 8 to 18 pt on ``legend-row``.
            inside = pitch / 2 + box.ink_centre
            if position == "b":
                pad = LEGEND_BAND_PAD_PT - (
                    LEGEND_WRAP_BAND_TRIM_PT if rows > 1 else 0.0
                )
                first = (
                    self.frame.bottom - pad - pitch + inside - (rows - 1) * pitch
                )
            else:
                first = self.frame.top + LEGEND_BAND_PAD_PT + inside
            if rows > 1:
                # **Past the cap the run becomes a grid**: equal columns as wide as the
                # widest entry, packed with no gap, the block centred on the frame.  See
                # :meth:`_legend_grid` for the row and column counts and what refuted the
                # greedy fill.
                column = max(widths)
                start = (
                    self.frame.left
                    + (self.frame.width - columns * column) / 2
                    + LEGEND_HORIZONTAL_OFFSET_PT
                )
                for index, (_, item) in enumerate(entries):
                    row, slot = divmod(index, columns)
                    self._legend_entry(
                        item,
                        start + slot * column + (cell - swatch) / 2,
                        first + row * pitch,
                        swatch,
                        gap,
                        font,
                    )
                return
            slack = min(
                LEGEND_ENTRY_SLACK * total,
                LEGEND_BAND_MAX_FRACTION * self.frame.width - total,
            )
            entry_gap = max(slack, 0.0) / (len(entries) + 1)
            run = total + entry_gap * (len(entries) - 1)
            x = (
                self.frame.left
                + (self.frame.width - run) / 2
                + LEGEND_HORIZONTAL_OFFSET_PT
            )
            for (_, item), width in zip(entries, widths):
                self._legend_entry(
                    item, x + (cell - swatch) / 2, first, swatch, gap, font
                )
                x += width + entry_gap
            return

        # A side legend sits one lead gap outside the plot area.  Measured 15.996 pt at
        # 10 pt with the legend on the right, and the band on the left came out exactly
        # the same width, so the left case mirrors it against the frame edge.
        if position == "l":
            # Measured 10.996 pt from the frame's left edge in the legend-l probe, which
            # is the plain edge inset and not the 6.5 pt the label column starts at.
            x = self.frame.left + EDGE_INSET_PT
        else:
            x = rect.right + second_band + LEGEND_SIDE_LEAD_EM * box.size
        # A stacked legend is centred on the frame and each entry is centred in its row.
        # Measured against both bar charts in real-financial-report.pptx: baselines land
        # within 0.18 pt, where treating the row like the horizontal band's off-centre
        # line was 5.7 pt out.
        #
        # **A wrapped entry takes every row with it, and the extra height hangs below the
        # line rather than round it.**  The block is still the rows centred on the frame,
        # but an entry's first baseline stays where a one-line row would have put it --
        # ``legend_row_pitch(box) / 2`` below its row's top, not half of the *opened*
        # pitch.  Measured on 20 slides of the ``legend-side`` deck at four wrap depths,
        # four frame heights, four entry counts and four sizes, and on
        # ``real-financial-report.pptx``'s doughnut, whose five baselines this puts within
        # 0.2 pt where centring the opened row was 5.9 pt low.
        column = self._legend_text_column(font, per_point=per_point)
        lines = max(
            (len(_legend_wrap(item.name or "", font, column)) for _, item in entries),
            default=1,
        )
        pitch = legend_row_pitch(box, lines)
        inside = legend_row_pitch(box) / 2 + box.ink_centre
        # **A title moves it down by half its own band**, which is measured rather than
        # inherited from the band the axis count already subtracts: nine probes on
        # ``view3d-surfrecon`` -- 3-D and flat, three frame heights and three title sizes
        # -- put the block's centre 9.38, 13.04 and 19.15 pt below the frame's at 8, 14
        # and 24 pt of title, and unmoved at 137.00 on every one of the thirteen titleless
        # probes beside them whatever the camera does to the scene.  Twice those shifts is
        # ``line_height + 8.99`` exactly at all three sizes, and it is a *constant* of the
        # title and not a share of the frame: 120, 250 and 330 pt frames all read 13.04.
        band = self._legend_title_band()
        y = (
            self.frame.top
            + band
            + (self.frame.height - band - pitch * len(entries)) / 2
        )
        for _, item in entries:
            self._legend_entry(item, x, y + inside, swatch, gap, font, column=column)
            y += pitch

    def _legend_title_band(self) -> float:
        """What the title takes off the height a **side** legend centres itself in.

        Not :data:`TITLE_BAND_LINES` times the line, which is the band the *axis* interval
        count is measured against: that reads 14.42, 25.24 and 43.27 pt at 8, 14 and 24 pt
        of title where the legend's own band reads 18.76, 26.08 and 38.30.  The two agree
        at 14 pt and nowhere else, which is why one of them could stand for the other
        until a size sweep was run.  See :data:`LEGEND_SIDE_TITLE_GAP_PT`.
        """
        title = self._title_box()
        if title is None:
            return 0.0
        return title.line_height + LEGEND_SIDE_TITLE_GAP_PT

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

        **A bubble takes the swatch**, which is the one place its legend parts company
        with the scatter it otherwise copies: the side- and bottom-legend probes both drew
        a 5.492 pt disc, and reading the line key there put our plot 13.4 pt narrow and the
        drawn bubble 3.1 pt small, because the legend reserve also feeds the region the
        largest bubble is sized against.
        """
        if self._line_legend_key():
            return LINE_LEGEND_KEY_PT, LINE_LEGEND_KEY_GAP_PT
        return LEGEND_SWATCH_EM * font.size, LEGEND_SWATCH_GAP_EM * font.size

    def _line_legend_key(self) -> bool:
        """Whether this chart's legend keys are rules rather than swatches.

        ``line_legend_keys`` is the combo's answer, set on every group by
        :meth:`_line_legend_keys`; a chart that never went through that path asks its own
        group.  Both are :attr:`_is_line_keyed`, which refuses a group with no rule to
        draw -- a stock chart's -- so such a chart lays its legend out on the swatch cell.
        """
        return self.line_legend_keys or self._is_line_keyed

    def _legend_key_cell(self, font: ChartFont) -> float:
        """The width a horizontal legend entry's key **advances**, key plus its padding.

        Wider than the drawn key, which is centred in it -- see
        :data:`LEGEND_ENTRY_KEY_EM`.  Like the key itself, the line form is absolute points
        and the swatch form scales with the type.
        """
        if self._line_legend_key():
            return LINE_LEGEND_ENTRY_KEY_PT
        return LEGEND_ENTRY_KEY_EM * font.size

    def _legend_entry(
        self,
        item: _Series,
        x: float,
        baseline: float,
        swatch: float,
        gap: float,
        font: ChartFont,
        *,
        column: float | None = None,
    ) -> None:
        box = font.box
        centre = baseline - box.ink_centre
        if item.line_keyed and item.line is not None:
            # A line key: the stroke across the whole swatch width with the series'
            # marker centred on it.  Measured on the radar legend probe and confirmed on
            # a line chart's, where the marker's centre landed 0.17 pt off the midpoint.
            self._line(x, centre, x + swatch, centre, item.line)
            if item.marker_symbol:
                self._marker((x + swatch / 2, centre), item)
        elif item.line_keyed:
            # **A series that legends with a rule and has no rule draws no rule**, and
            # nothing takes its place: `legend-nokey`'s three bare line series came back
            # with no key path on the page at all.  What it does keep is its marker, drawn
            # alone and centred in the key's own width -- measured 0.07 pt off that centre
            # on `o-mark` -- so the key loses its stroke and not its slot.
            #
            # The slot is narrower than a line key's whenever the *chart* has no rule
            # anywhere, because :meth:`_line_legend_key` then reads the swatch; a bare
            # series beside one that keeps its rule stays on the wide cell and simply
            # leaves it empty, which is the `o-mix` reading.  See ROADMAP.md 3.2a.
            if item.marker_symbol:
                self._marker((x + swatch / 2, centre), item)
        else:
            # **A swatch in a combo is as wide as the chart's key and as tall as a
            # swatch.**  Measured on ``combo-legend``: the bar keys of a bar-plus-line
            # chart came back 19.200 pt wide -- the line key's width, not the 5.49 pt
            # swatch -- and 5.49 pt tall.  On a chart with no line group the two numbers
            # are the same and this is the square every bar and pie legend measured.
            self._rect(
                _Rect(
                    x,
                    centre - LEGEND_SWATCH_EM * font.size / 2,
                    x + swatch,
                    centre + LEGEND_SWATCH_EM * font.size / 2,
                ),
                fill=item.fill,
                # **A key with no fill keeps its stroke**, which is what a `c:wireframe`
                # surface's band key is: the same square, outlined in the band's colour
                # and empty.  Measured on the wireframe legend probe, whose five keys
                # carry a stroke and no fill at all.  An entry that has a fill draws no
                # outline, which is every other chart's key and is left alone.
                outline=item.outline if item.fill is None else None,
            )
        # An entry wider than the band it sits in wraps rather than running out of the
        # frame.  ``column`` is the width a **side** legend leaves for the name -- see
        # :meth:`_legend_text_column` -- and the row pitch above has already been opened
        # for the lines it produces, so the entry below is clear of them.  A horizontal
        # legend passes no column: its entries are what the band's width was fitted to and
        # none of them wraps, so the fallback is the old overflow guard against the frame.
        left = x + swatch + gap
        name = item.name or ""
        natural = font.width(name) + box.size
        edge = self.frame.right - FRAME_PADDING_PT - left
        body = self._label_body(name, font, align="l")
        if column is None:
            # A horizontal legend: the band's width was fitted to these entries and none
            # of them wraps, so this is only the guard against one running off the frame.
            width = min(natural, max(edge, box.size))
            if natural > edge > box.size:
                body = replace(body, body_properties=CHART_WRAPPED_TEXT_BODY)
        elif font.width(name) > column:
            # A side legend's entry that does not fit its column.  The row pitch above has
            # already been opened for the lines this produces.
            width = column
            body = replace(body, body_properties=CHART_WRAPPED_TEXT_BODY)
        else:
            # One that does fit: the box is the name's own advance and a little slack, so
            # a hundredth of a point between our measurement and the renderer's cannot
            # push it onto a second line.
            width = min(natural, max(edge, box.size))
        self._text(body, left=left, width=width, baseline=baseline, box=box)

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


def _resolve_view_3d(view: "c.SourceChartView3D | None") -> m.Chart3DView | None:
    """``c:view3D`` into the render model, field for field and value for value.

    Nothing is defaulted on the way through.  A ``None`` here is the file's silence, and
    that silence has a meaning of its own: it selects the perspective scene
    :func:`three_d_camera` refuses.  The layout reads
    :attr:`~pptx2svg.parse.chart.SourceChart.view_3d` directly; this is the copy callers
    of ``convert_pptx_to_model`` get.
    """
    if view is None:
        return None
    return m.Chart3DView(
        rot_x=view.rot_x,
        rot_y=view.rot_y,
        depth_percent=view.depth_percent,
        h_percent=view.h_percent,
        right_angle_axes=view.right_angle_axes,
        perspective=view.perspective,
    )


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


def _to_linear(channel: float) -> float:
    """sRGB 0..255 to linear light.  The IEC 61966-2-1 transfer, toe included."""
    value = channel / 255.0
    return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4


def _to_srgb(value: float) -> int:
    value = max(0.0, min(1.0, value))
    encoded = value * 12.92 if value <= 0.0031308 else 1.055 * value ** (1 / 2.4) - 0.055
    return int(round(encoded * 255))


def _cycle_shift(hex_color: str, cycle: int) -> str:
    """The luminance variation a per-point accent takes in its *n*-th cycle of six.

    See :data:`VARY_COLOR_CYCLE_SHADE`.  Cycle 0 is the plain accent when the chart needs
    six colours or fewer; the caller decides that, because this cannot see the count.
    """
    text = hex_color.lstrip("#")
    if len(text) != 6:
        return hex_color
    channels = [int(text[index : index + 2], 16) for index in (0, 2, 4)]
    out = []
    for channel in channels:
        linear = _to_linear(channel)
        if cycle <= 0:
            linear *= VARY_COLOR_CYCLE_SHADE
        else:
            linear += VARY_COLOR_CYCLE_TINT * (1.0 - linear)
        out.append(_to_srgb(linear))
    return "#" + "".join(f"{value:02x}" for value in out)


def _tangent_point(
    origin: tuple[float, float],
    centre: tuple[float, float],
    radius: float,
    upper: bool,
) -> tuple[float, float] | None:
    """Where a line from ``origin`` touches the circle, on the upper or lower side.

    An ofPie's connector is tangent to its second pie -- measured, see
    :meth:`ChartBuilder._draw_of_pie_connector`.  ``None`` when the origin is inside the
    circle, which no drawable layout produces but a hand-written `c:gapWidth` can.
    """
    dx, dy = centre[0] - origin[0], centre[1] - origin[1]
    distance = math.hypot(dx, dy)
    if distance <= radius or radius <= 0:
        return None
    length = math.sqrt(distance * distance - radius * radius)
    base = math.atan2(dy, dx)
    spread = math.asin(min(radius / distance, 1.0))
    candidates = [
        (origin[0] + length * math.cos(base + sign * spread),
         origin[1] + length * math.sin(base + sign * spread))
        for sign in (1.0, -1.0)
    ]
    candidates.sort(key=lambda point: point[1])
    return candidates[0] if upper else candidates[-1]


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
    "bottom_axis_intervals",
    "default_font_size",
    "font_box",
    "format_number",
    "nice_axis_scale",
    "radial_axis_intervals",
    "side_axis_intervals",
    "text_width",
]

#!/usr/bin/env python3
"""Build ``tests/fixtures/chart-gallery.pptx`` -- one chart type per slide.

Why this deck exists
--------------------

Charts are the largest body of measured behaviour in this project: ten drawable types, an
axis rule settled over 616 probe readings, legend and plot-area layout.  **No deck in the
committed corpus is chart-heavy.**  ``real-financial-report.pptx`` has four charts and is
*skipped* by ``tools/fidelity.py`` for want of Noto Sans JP; ``authoring-integration.pptx``
has one.  So the one subsystem with the most measurement behind it had the least fixture
coverage, and a regression in it could only be caught by a unit test somebody thought to
write.

This deck is the missing input.  Every ``c:*Chart`` group element the reader recognises
appears on a slide of its own -- the ten that draw, the four 3-D spellings that degrade to
them, ``surfaceChart`` and its band legend -- plus a combo chart, and the whole
thing is generated from this file so that anyone can regenerate it and read what went in.

Authored to be *scorable*
-------------------------

The property that made this worth building is that ``tools/fidelity.py`` can score it.
Five of the seven decks already in ``tests/fixtures/`` are skipped by the oracle, four of
them because PowerPoint itself substitutes a face on the reference machine (Noto Sans JP,
ＭＳ Ｐゴシック) -- which turns any comparison into a measurement of font availability
rather than of this library.  So:

* **No CJK text anywhere**, on the slides or in the charts.
* Every face is the theme's, and the theme is the Aptos one carried by
  ``authoring-integration.pptx`` -- the deck the oracle already scores.  Nothing here
  writes an ``a:latin`` of its own, so ``fidelity.requested_faces`` sees exactly
  ``Aptos`` and ``Aptos Display``, both of which PowerPoint draws natively here.

Where the deck comes from
-------------------------

Entirely ours.  The package skeleton -- theme, master, layout, ``presentation.xml`` --
is lifted from ``tests/fixtures/authoring-integration.pptx``, which this repository
generates from its own authoring API, and the master and layout are then **emptied** so
nothing but the chart and its caption is drawn.  Every ``ppt/charts/*.xml`` part is
written by the functions below.  No third-party deck and no copied content.

The charts carry no ``c:externalData``: the numbers live in the ``c:numCache`` /
``c:strCache`` that PowerPoint would draw from anyway, which is the same shape the axis
probe decks use and which PowerPoint has opened and exported several hundred times.

Authoring hazards this file works around
----------------------------------------

* ``<c:bubble3D val="1"/>`` **hangs** PowerPoint 16.x on macOS, reproducibly, so the
  bubble slide writes no ``c:bubble3D`` at all.
* OOXML child order is enforced, not advisory.  A group-level ``c:marker`` belongs after
  every ``c:ser``; a series-level one belongs before its ``c:dLbls``.  A deck that gets
  this wrong does not render wrongly, it hangs the export -- which reads as evidence
  about a chart type and is not.

What PowerPoint would not author: a ChartEx
-------------------------------------------

There is **no ``cx:chartSpace`` slide here**, and that is a measurement rather than an
omission.  Four hand-written ChartEx parts were tried -- a treemap with
``numDim type="size"``, the same with ``type="val"``, a minimal waterfall with one
numeric dimension, and a treemap naming an embedded workbook through
``cx:externalData`` -- at both ``ppt/charts/chart18.xml`` and the ``chartEx18.xml``
spelling PowerPoint itself uses.  **Every one of them hangs PowerPoint inside ``open``**,
and the whole 17-slide deck plus any one of them exports to nothing at all: the
AppleScript ``save ... as save as PDF`` returns success and writes no file, then every
later export fails ``-9074`` until ``pkill``.  The same 17 slides without it export in
seconds.

So a hand-authored ChartEx cannot be a fixture on this machine -- "the acceptance test is
that PowerPoint opens it", and PowerPoint does not.  The warning path it would have
pinned is covered by ``tests/test_chart.py`` instead --
``test_an_office_2016_chartex_frame_says_what_it_is`` for the ``cx:`` branch and
``test_a_chart_type_that_is_not_implemented_says_so`` for the ``c:`` one, both on decks
that never go near PowerPoint.  No slide here refuses any more: ``surfaceChart`` was the
last group element out and is drawn now.  Getting a real ChartEx into the corpus needs a
deck Office wrote, not one we wrote.

Usage::

    python3 tools/make_chart_gallery.py                       # rewrite the fixture
    python3 tools/make_chart_gallery.py ~/pptx2svg-oracle/chart-gallery.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/chart-gallery.pptx ~/pptx2svg-oracle/chart-gallery.pdf

Then rebaseline the snapshots and the oracle::

    python3 -m pytest tests/test_vrt.py --update-snapshots
    python3 tools/fidelity.py --update
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SOURCE = ROOT / "tests/fixtures/authoring-integration.pptx"
TARGET = ROOT / "tests/fixtures/chart-gallery.pptx"

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

EMU_PER_POINT = 12700

#: The slide is 720 x 405 pt.  The caption sits on the first band, the chart below it.
CAPTION_OFF = (36 * EMU_PER_POINT, 14 * EMU_PER_POINT)
CAPTION_EXT = (648 * EMU_PER_POINT, 22 * EMU_PER_POINT)
FRAME_OFF = (36 * EMU_PER_POINT, 46 * EMU_PER_POINT)

EMPTY_TREE = (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
)

#: Every chart states its own text size rather than leaving PowerPoint to scale one to the
#: frame.  An explicit ``c:txPr`` is what the 616 axis-probe readings were taken under, so
#: a difference between this deck's render and its export is a layout difference and not a
#: disagreement about how big the labels are.
TEXT_SIZE = 1000


# --------------------------------------------------------------------------------------
# Data
# --------------------------------------------------------------------------------------

QUARTERS = ("Q1", "Q2", "Q3", "Q4")

#: Deliberately long, unbreakable single tokens.  On the 288 pt frame slide 1 gives them
#: the widest will not fit in its band, which is the condition that turns the whole axis
#: 45 degrees -- the rotated-label path, and the one the corpus says has no cap on its
#: reserve (ROADMAP.md, *What the corpus says is wrong now*).
LONG_CATEGORIES = (
    "Infrastructure",
    "Administration",
    "Transportation",
    "Manufacturing",
    "Communications",
)

SEGMENTS = ("Direct", "Partner", "Online", "Retail", "Wholesale")


def _points(values, tag="c:v") -> str:
    return "".join(f"<c:pt idx='{i}'><{tag}>{v}</{tag}></c:pt>" for i, v in enumerate(values))


def cats(names) -> str:
    """``c:cat`` holding a cached string literal per category."""
    inner = _points([escape(name) for name in names])
    return (
        "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='{len(names)}'/>{inner}"
        "</c:strCache></c:strRef></c:cat>"
    )


def nums(tag: str, values, fmt: str = "General") -> str:
    """``c:val`` / ``c:xVal`` / ``c:yVal`` / ``c:bubbleSize`` holding cached numbers."""
    inner = _points([repr(float(v)) for v in values])
    return (
        f"<c:{tag}><c:numRef><c:numCache><c:formatCode>{fmt}</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{inner}"
        f"</c:numCache></c:numRef></c:{tag}>"
    )


def series_name(name: str) -> str:
    return (
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        f"<c:pt idx='0'><c:v>{escape(name)}</c:v></c:pt>"
        "</c:strCache></c:strRef></c:tx>"
    )


# --------------------------------------------------------------------------------------
# Building blocks
# --------------------------------------------------------------------------------------


def cat_axis(
    ax_id: int,
    cross_id: int,
    *,
    position: str = "b",
    delete: bool = False,
    gridlines: bool = False,
) -> str:
    """``c:catAx``.  Child order is the schema's: id, scaling, delete, pos, grid, ..."""
    return (
        f"<c:catAx><c:axId val='{ax_id}'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if delete else 0}'/><c:axPos val='{position}'/>"
        + ("<c:majorGridlines/>" if gridlines else "")
        + "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        f"<c:crossAx val='{cross_id}'/><c:crosses val='autoZero'/>"
        "<c:auto val='1'/><c:lblAlgn val='ctr'/><c:lblOffset val='100'/>"
        "<c:noMultiLvlLbl val='0'/></c:catAx>"
    )


def val_axis(
    ax_id: int,
    cross_id: int,
    *,
    position: str = "l",
    delete: bool = False,
    gridlines: bool = True,
    cross_between: str = "between",
    crosses: str = "autoZero",
) -> str:
    """``c:valAx``.  ``crosses='max'`` is what puts a secondary axis on the far side."""
    return (
        f"<c:valAx><c:axId val='{ax_id}'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if delete else 0}'/><c:axPos val='{position}'/>"
        + ("<c:majorGridlines/>" if gridlines else "")
        + "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        f"<c:crossAx val='{cross_id}'/><c:crosses val='{crosses}'/>"
        f"<c:crossBetween val='{cross_between}'/></c:valAx>"
    )


def ser_axis(ax_id: int, cross_id: int) -> str:
    """``c:serAx`` -- the depth axis a 3-D group needs to be schema-valid.

    ``line3DChart`` and ``surfaceChart`` require exactly three ``c:axId`` children, and
    PowerPoint writes a real ``c:serAx`` for each.  Deleted, so nothing is drawn for it.
    """
    return (
        f"<c:serAx><c:axId val='{ax_id}'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='1'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        f"<c:crossAx val='{cross_id}'/><c:crosses val='autoZero'/></c:serAx>"
    )


def title(text: str, size: int = 1400) -> str:
    return (
        "<c:title><c:tx><c:rich><a:bodyPr rot='0' spcFirstLastPara='1' "
        "vertOverflow='ellipsis' vert='horz' wrap='square' anchor='ctr' anchorCtr='1'/>"
        "<a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{size}' b='0'/></a:pPr>"
        f"<a:r><a:rPr lang='en-US' sz='{size}' b='0'/><a:t>{escape(text)}</a:t></a:r>"
        "</a:p></c:rich></c:tx><c:overlay val='0'/></c:title>"
        "<c:autoTitleDeleted val='0'/>"
    )


NO_TITLE = "<c:autoTitleDeleted val='1'/>"


def legend(position: str) -> str:
    return f"<c:legend><c:legendPos val='{position}'/><c:overlay val='0'/></c:legend>"


#: ``c:view3D`` as PowerPoint writes it for a default 3-D chart.  Nothing here reads it --
#: the 3-D groups are drawn flat -- which is the point: it has to be *ignored*, not
#: choked on.
VIEW_3D = (
    "<c:view3D><c:rotX val='15'/><c:rotY val='20'/><c:depthPercent val='100'/>"
    "<c:rAngAx val='1'/></c:view3D>"
)


def chart_space(body: str, *, size: int = TEXT_SIZE) -> bytes:
    """Wrap a ``c:chart`` body in a ``c:chartSpace`` with an explicit text size."""
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
        f"<c:chart>{body}</c:chart>"
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
        "</c:chartSpace>"
    ).encode()


def plot_area(groups: str, axes: str) -> str:
    """``c:plotArea``: layout, then every chart group, then every axis.  In that order."""
    return f"<c:plotArea><c:layout/>{groups}{axes}</c:plotArea>"


def tail(legend_xml: str = "") -> str:
    return f"{legend_xml}<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/>"


# --------------------------------------------------------------------------------------
# One function per slide.  Each returns the bytes of a chart part.
# --------------------------------------------------------------------------------------


def bar_column() -> bytes:
    """Slide 1 -- ``barChart``, clustered columns, two series, rotated category labels.

    The widest category will not fit its band on this frame, which is the condition that
    turns the axis 45 degrees.  Title present, legend below.
    """
    groups = (
        "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
        "<c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Plan")
        + cats(LONG_CATEGORIES)
        + nums("val", (42, 58, 31, 66, 49))
        + "</c:ser>"
        "<c:ser><c:idx val='1'/><c:order val='1'/>"
        + series_name("Actual")
        + cats(LONG_CATEGORIES)
        + nums("val", (37, 61, 44, 52, 58))
        + "</c:ser>"
        "<c:gapWidth val='150'/><c:overlap val='-27'/>"
        "<c:axId val='111000'/><c:axId val='111001'/></c:barChart>"
    )
    axes = cat_axis(111000, 111001) + val_axis(111001, 111000)
    return chart_space(
        title("Plan against actual") + plot_area(groups, axes) + tail(legend("b"))
    )


def bar_horizontal() -> bytes:
    """Slide 2 -- ``barChart``, horizontal bars, **value axis along the bottom**.

    The one arrangement where a value label's own width can gate the tick count, and the
    only slide in the deck that carries data labels.  No title, no legend.
    """
    groups = (
        "<c:barChart><c:barDir val='bar'/><c:grouping val='clustered'/>"
        "<c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Units")
        + cats(SEGMENTS)
        + nums("val", (1240, 860, 1980, 640, 1510))
        + "</c:ser>"
        "<c:dLbls><c:showLegendKey val='0'/><c:showVal val='1'/>"
        "<c:showCatName val='0'/><c:showSerName val='0'/><c:showPercent val='0'/>"
        "<c:showBubbleSize val='0'/></c:dLbls>"
        "<c:gapWidth val='150'/>"
        "<c:axId val='112000'/><c:axId val='112001'/></c:barChart>"
    )
    axes = cat_axis(112000, 112001, position="l") + val_axis(
        112001, 112000, position="b"
    )
    return chart_space(NO_TITLE + plot_area(groups, axes) + tail())


def line() -> bytes:
    """Slide 3 -- ``lineChart``, three series with markers, legend at the **top**.

    ``c:marker`` appears twice and in two different places: on each ``c:ser`` before its
    ``c:dLbls`` slot, and once on the group *after* the last ``c:ser``.  Both positions
    are the schema's; getting either wrong hangs the export rather than failing to parse.
    """
    values = ((12, 19, 24, 21), (8, 11, 17, 26), (22, 18, 14, 11))
    names = ("North", "South", "Export")
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + "<c:marker><c:symbol val='circle'/><c:size val='5'/></c:marker>"
        + cats(QUARTERS)
        + nums("val", row)
        + "<c:smooth val='0'/></c:ser>"
        for i, (name, row) in enumerate(zip(names, values))
    )
    groups = (
        "<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>"
        + sers
        + "<c:marker val='1'/>"
        "<c:axId val='113000'/><c:axId val='113001'/></c:lineChart>"
    )
    axes = cat_axis(113000, 113001) + val_axis(113001, 113000)
    return chart_space(NO_TITLE + plot_area(groups, axes) + tail(legend("t")))


def area() -> bytes:
    """Slide 4 -- ``areaChart``, stacked, legend at the **right**, title present.

    States no ``c:crossBetween``, which for an area is the case
    :data:`resolve.chart.DEFAULT_AREA_CROSS_BETWEEN` was measured for: the first vertex
    lands on the plot's left edge rather than at the first band's centre.
    """
    names = ("Licences", "Services", "Support")
    values = ((30, 34, 39, 45), (18, 22, 21, 27), (9, 11, 15, 14))
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + cats(QUARTERS)
        + nums("val", row)
        + "</c:ser>"
        for i, (name, row) in enumerate(zip(names, values))
    )
    groups = (
        "<c:areaChart><c:grouping val='stacked'/><c:varyColors val='0'/>"
        + sers
        + "<c:axId val='114000'/><c:axId val='114001'/></c:areaChart>"
    )
    # No `c:crossBetween` on the value axis: see the docstring.
    axes = cat_axis(114000, 114001) + (
        "<c:valAx><c:axId val='114001'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='l'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        "<c:crossAx val='114000'/><c:crosses val='autoZero'/></c:valAx>"
    )
    return chart_space(
        title("Revenue mix by quarter") + plot_area(groups, axes) + tail(legend("r"))
    )


def scatter() -> bytes:
    """Slide 5 -- ``scatterChart``, two value axes and no category axis."""
    sers = (
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Cohort A")
        + "<c:marker><c:symbol val='circle'/><c:size val='6'/></c:marker>"
        + nums("xVal", (1.2, 2.4, 3.1, 4.7, 5.9, 7.2))
        + nums("yVal", (3.4, 4.9, 4.1, 6.8, 6.2, 8.5))
        + "<c:smooth val='0'/></c:ser>"
        "<c:ser><c:idx val='1'/><c:order val='1'/>"
        + series_name("Cohort B")
        + "<c:marker><c:symbol val='square'/><c:size val='6'/></c:marker>"
        + nums("xVal", (1.0, 2.9, 3.8, 5.1, 6.4, 7.8))
        + nums("yVal", (1.9, 2.6, 3.9, 4.2, 5.6, 5.1))
        + "<c:smooth val='0'/></c:ser>"
    )
    groups = (
        "<c:scatterChart><c:scatterStyle val='lineMarker'/><c:varyColors val='0'/>"
        + sers
        + "<c:axId val='115000'/><c:axId val='115001'/></c:scatterChart>"
    )
    axes = val_axis(
        115000, 115001, position="b", gridlines=False, cross_between="midCat"
    ) + val_axis(115001, 115000, position="l", cross_between="midCat")
    return chart_space(NO_TITLE + plot_area(groups, axes) + tail(legend("r")))


def bubble() -> bytes:
    """Slide 6 -- ``bubbleChart``.

    **No ``c:bubble3D`` anywhere.**  ``<c:bubble3D val="1"/>`` hangs PowerPoint 16.x on
    macOS reproducibly, and the element is optional, so the honest way to keep this deck
    exportable is to leave it out rather than to state the false half of it.
    """
    sers = (
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Region")
        + nums("xVal", (2.0, 4.0, 6.0, 8.0, 10.0))
        + nums("yVal", (5.0, 8.0, 4.0, 9.0, 6.0))
        + nums("bubbleSize", (14.0, 32.0, 9.0, 40.0, 21.0))
        + "</c:ser>"
        "<c:ser><c:idx val='1'/><c:order val='1'/>"
        + series_name("Channel")
        + nums("xVal", (3.0, 5.5, 7.0, 9.5))
        + nums("yVal", (2.0, 6.5, 7.5, 3.5))
        + nums("bubbleSize", (25.0, 12.0, 30.0, 18.0))
        + "</c:ser>"
    )
    groups = (
        "<c:bubbleChart><c:varyColors val='0'/>"
        + sers
        + "<c:bubbleScale val='100'/><c:showNegBubbles val='0'/>"
        "<c:axId val='116000'/><c:axId val='116001'/></c:bubbleChart>"
    )
    axes = val_axis(
        116000, 116001, position="b", gridlines=False, cross_between="midCat"
    ) + val_axis(116001, 116000, position="l", cross_between="midCat")
    return chart_space(
        title("Share against growth") + plot_area(groups, axes) + tail(legend("b"))
    )


def pie() -> bytes:
    """Slide 7 -- ``pieChart``, one series, ``c:varyColors`` on, legend at the right."""
    groups = (
        "<c:pieChart><c:varyColors val='1'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Spend")
        + cats(SEGMENTS)
        + nums("val", (34, 21, 18, 15, 12))
        + "</c:ser>"
        "<c:firstSliceAng val='0'/></c:pieChart>"
    )
    return chart_space(
        title("Spend by channel") + plot_area(groups, "") + tail(legend("r"))
    )


def doughnut() -> bytes:
    """Slide 8 -- ``doughnutChart`` with an explicit hole, legend on the **left**.

    ``c:holeSize`` is stated rather than left out on purpose: PowerPoint draws a
    ``doughnutChart`` that names no hole size as a **solid pie**, which is a separate
    measured case and not the one this slide is for.
    """
    groups = (
        "<c:doughnutChart><c:varyColors val='1'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Headcount")
        + cats(SEGMENTS)
        + nums("val", (28, 24, 19, 17, 12))
        + "</c:ser>"
        "<c:firstSliceAng val='0'/><c:holeSize val='55'/></c:doughnutChart>"
    )
    return chart_space(NO_TITLE + plot_area(groups, "") + tail(legend("l")))


def of_pie() -> bytes:
    """Slide 9 -- ``ofPieChart`` in its **bar** form, split by position.

    The bar form packs the two plots differently from the pie form and has its own
    measured divisor (:data:`resolve.chart.OF_PIE_BAR_GAP_DIVISOR`), fitted to a single
    ``gapWidth=100`` observation -- which is the value stated here, so the slide sits on
    the measurement rather than beside it.
    """
    names = ("Enterprise", "Mid market", "Small business", "Education", "Public", "Other")
    groups = (
        "<c:ofPieChart><c:ofPieType val='bar'/><c:varyColors val='1'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Accounts")
        + cats(names)
        + nums("val", (48, 22, 14, 7, 5, 4))
        + "</c:ser>"
        "<c:gapWidth val='100'/><c:splitType val='pos'/><c:splitPos val='3'/>"
        "<c:secondPieSize val='75'/></c:ofPieChart>"
    )
    return chart_space(
        title("Accounts by tier") + plot_area(groups, "") + tail(legend("b"))
    )


def radar() -> bytes:
    """Slide 10 -- ``radarChart``, six spokes, two series, legend below.

    Latin category labels at 10 pt: the radial-axis ring count was fitted over two probe
    decks at exactly this size and face, so this slide sits inside the fit rather than on
    the one corpus reading that contradicts it (which is Japanese at 12 pt Arial).
    """
    spokes = ("Speed", "Cost", "Quality", "Support", "Reach", "Risk")
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + "<c:marker><c:symbol val='circle'/><c:size val='5'/></c:marker>"
        + cats(spokes)
        + nums("val", row)
        + "</c:ser>"
        for i, (name, row) in enumerate(
            zip(("This year", "Last year"), ((80, 55, 90, 65, 45, 70), (60, 70, 75, 50, 60, 55)))
        )
    )
    groups = (
        "<c:radarChart><c:radarStyle val='marker'/><c:varyColors val='0'/>"
        + sers
        + "<c:axId val='117000'/><c:axId val='117001'/></c:radarChart>"
    )
    axes = cat_axis(117000, 117001) + val_axis(117001, 117000)
    return chart_space(NO_TITLE + plot_area(groups, axes) + tail(legend("b")))


def stock() -> bytes:
    """Slide 11 -- ``stockChart``, high-low-close, with ``c:hiLowLines``.

    Every series states ``<a:ln><a:noFill/></a:ln>``: a stock chart's lines are hidden by
    the **file**, not by the renderer, which is how a ``stockChart`` with neither
    ``c:hiLowLines`` nor ``c:upDownBars`` came back from PowerPoint drawn exactly like a
    line chart.  Only the close series carries a marker.
    """
    days = ("Mon", "Tue", "Wed", "Thu", "Fri")
    rows = (
        ("High", (118, 124, 121, 131, 129), "<c:marker><c:symbol val='none'/></c:marker>"),
        ("Low", (104, 109, 112, 116, 119), "<c:marker><c:symbol val='none'/></c:marker>"),
        ("Close", (112, 121, 115, 128, 123), "<c:marker><c:symbol val='dot'/><c:size val='3'/></c:marker>"),
    )
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + "<c:spPr><a:ln w='19050'><a:noFill/></a:ln></c:spPr>"
        + marker
        + cats(days)
        + nums("val", row)
        + "</c:ser>"
        for i, (name, row, marker) in enumerate(rows)
    )
    groups = (
        "<c:stockChart>"
        + sers
        + "<c:hiLowLines/>"
        "<c:axId val='118000'/><c:axId val='118001'/></c:stockChart>"
    )
    axes = cat_axis(118000, 118001) + val_axis(118001, 118000)
    return chart_space(
        title("Session range") + plot_area(groups, axes) + tail(legend("b"))
    )


def surface() -> bytes:
    """Slide 12 -- ``surfaceChart``: a lit mesh coloured by **value band**.

    The one chart type in the deck whose marks are not its series.  Three series over four
    quarters make a lattice of twelve points, drawn as a sheet through the scene's depth
    and cut into nine bands by the value axis' own intervals, each band a step of the
    accent ramp; the legend is of those bands rather than of the series, which is a legend
    model no other slide here has.  See ROADMAP.md 3.4.

    A surface group needs three ``c:axId`` children, so it carries a real ``c:serAx``.
    """
    rows = (("Low", (10, 16, 22, 19)), ("Mid", (18, 26, 31, 27)), ("High", (24, 33, 41, 36)))
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + cats(QUARTERS)
        + nums("val", row)
        + "</c:ser>"
        for i, (name, row) in enumerate(rows)
    )
    groups = (
        "<c:surfaceChart><c:wireframe val='0'/>"
        + sers
        + "<c:axId val='119000'/><c:axId val='119001'/><c:axId val='119002'/>"
        "</c:surfaceChart>"
    )
    axes = (
        cat_axis(119000, 119001)
        + val_axis(119001, 119000)
        + ser_axis(119002, 119001)
    )
    return chart_space(
        VIEW_3D + title("Surface by quarter") + plot_area(groups, axes) + tail(legend("r"))
    )


def bar_3d() -> bytes:
    """Slide 13 -- ``bar3DChart``.  Parses as ``barChart`` and draws flat.

    Never compared against real output before this deck (ROADMAP.md 3.4).
    """
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + cats(QUARTERS)
        + nums("val", row)
        + "</c:ser>"
        for i, (name, row) in enumerate(zip(("Plan", "Actual"), ((36, 48, 41, 55), (31, 52, 46, 49))))
    )
    groups = (
        "<c:bar3DChart><c:barDir val='col'/><c:grouping val='clustered'/>"
        "<c:varyColors val='0'/>"
        + sers
        + "<c:gapWidth val='150'/><c:gapDepth val='150'/><c:shape val='box'/>"
        "<c:axId val='120000'/><c:axId val='120001'/><c:axId val='120002'/>"
        "</c:bar3DChart>"
    )
    axes = cat_axis(120000, 120001) + val_axis(120001, 120000) + ser_axis(120002, 120001)
    return chart_space(
        VIEW_3D + title("Three-D columns") + plot_area(groups, axes) + tail(legend("b"))
    )


def line_3d() -> bytes:
    """Slide 14 -- ``line3DChart``.  Three ``c:axId`` children, exactly, per the schema."""
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + cats(QUARTERS)
        + nums("val", row)
        + "<c:smooth val='0'/></c:ser>"
        for i, (name, row) in enumerate(zip(("North", "South"), ((14, 21, 19, 27), (9, 13, 22, 18))))
    )
    groups = (
        "<c:line3DChart><c:grouping val='standard'/><c:varyColors val='0'/>"
        + sers
        + "<c:gapDepth val='150'/>"
        "<c:axId val='121000'/><c:axId val='121001'/><c:axId val='121002'/>"
        "</c:line3DChart>"
    )
    axes = cat_axis(121000, 121001) + val_axis(121001, 121000) + ser_axis(121002, 121001)
    return chart_space(VIEW_3D + NO_TITLE + plot_area(groups, axes) + tail(legend("b")))


def pie_3d() -> bytes:
    """Slide 15 -- ``pie3DChart``.  No axes at all, like the 2-D pie it degrades to."""
    groups = (
        "<c:pie3DChart><c:varyColors val='1'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Spend")
        + cats(SEGMENTS)
        + nums("val", (31, 24, 20, 14, 11))
        + "</c:ser></c:pie3DChart>"
    )
    return chart_space(
        VIEW_3D + title("Three-D pie") + plot_area(groups, "") + tail(legend("r"))
    )


def area_3d() -> bytes:
    """Slide 16 -- ``area3DChart``, stacked."""
    sers = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>"
        + series_name(name)
        + cats(QUARTERS)
        + nums("val", row)
        + "</c:ser>"
        for i, (name, row) in enumerate(zip(("Base", "Uplift"), ((22, 26, 30, 35), (8, 12, 9, 15))))
    )
    groups = (
        "<c:area3DChart><c:grouping val='stacked'/><c:varyColors val='0'/>"
        + sers
        + "<c:gapDepth val='150'/>"
        "<c:axId val='122000'/><c:axId val='122001'/><c:axId val='122002'/>"
        "</c:area3DChart>"
    )
    axes = cat_axis(122000, 122001) + val_axis(122001, 122000) + ser_axis(122002, 122001)
    return chart_space(VIEW_3D + NO_TITLE + plot_area(groups, axes) + tail(legend("b")))


def combo() -> bytes:
    """Slide 17 -- a combo: ``barChart`` and ``lineChart`` over a **secondary** axis.

    Four axes: the primary pair, plus a deleted secondary category axis and a secondary
    value axis that ``crosses='max'`` puts on the right.  This slide pinned the defect
    ROADMAP.md 3.3 was written around -- the renderer drew the first group it could and
    dropped the rest -- and now pins the fix; it is the one chart in the corpus that
    exercises two value axes at once.
    """
    bars = (
        "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
        "<c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        + series_name("Revenue")
        + cats(QUARTERS)
        + nums("val", (120, 148, 131, 166))
        + "</c:ser>"
        "<c:gapWidth val='150'/>"
        "<c:axId val='123000'/><c:axId val='123001'/></c:barChart>"
    )
    lines = (
        "<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>"
        "<c:ser><c:idx val='1'/><c:order val='1'/>"
        + series_name("Margin %")
        + "<c:marker><c:symbol val='circle'/><c:size val='5'/></c:marker>"
        + cats(QUARTERS)
        + nums("val", (12.5, 14.0, 11.8, 15.2))
        + "<c:smooth val='0'/></c:ser>"
        "<c:marker val='1'/>"
        "<c:axId val='123002'/><c:axId val='123003'/></c:lineChart>"
    )
    axes = (
        cat_axis(123000, 123001)
        + val_axis(123001, 123000)
        + val_axis(123003, 123002, position="r", gridlines=False, crosses="max")
        + cat_axis(123002, 123003, delete=True)
    )
    return chart_space(
        title("Revenue and margin") + plot_area(bars + lines, axes) + tail(legend("b"))
    )


# --------------------------------------------------------------------------------------
# The deck
# --------------------------------------------------------------------------------------

WIDE = (648 * EMU_PER_POINT, 336 * EMU_PER_POINT)
MEDIUM = (520 * EMU_PER_POINT, 336 * EMU_PER_POINT)
NARROW = (288 * EMU_PER_POINT, 288 * EMU_PER_POINT)
SQUARE = (400 * EMU_PER_POINT, 336 * EMU_PER_POINT)

#: One row per slide: the chart part's builder, the frame, and the caption that names the
#: group element so a reader of the rendered SVG can tell the slides apart.
SLIDES: list[dict] = [
    {"key": "barChart", "build": bar_column, "frame": NARROW,
     "caption": "1  barChart - clustered columns, rotated category labels, legend below"},
    {"key": "barChart-bar", "build": bar_horizontal, "frame": WIDE,
     "caption": "2  barChart - horizontal bars, value axis along the bottom, data labels"},
    {"key": "lineChart", "build": line, "frame": WIDE,
     "caption": "3  lineChart - three series with markers, legend above"},
    {"key": "areaChart", "build": area, "frame": WIDE,
     "caption": "4  areaChart - stacked, no crossBetween stated, legend at the right"},
    {"key": "scatterChart", "build": scatter, "frame": MEDIUM,
     "caption": "5  scatterChart - two value axes, no category axis"},
    {"key": "bubbleChart", "build": bubble, "frame": MEDIUM,
     "caption": "6  bubbleChart - two series, no bubble3D (it hangs PowerPoint)"},
    {"key": "pieChart", "build": pie, "frame": SQUARE,
     "caption": "7  pieChart - one series, varyColors, legend at the right"},
    {"key": "doughnutChart", "build": doughnut, "frame": SQUARE,
     "caption": "8  doughnutChart - holeSize 55, legend at the left"},
    {"key": "ofPieChart", "build": of_pie, "frame": MEDIUM,
     "caption": "9  ofPieChart - bar form, split by position"},
    {"key": "radarChart", "build": radar, "frame": SQUARE,
     "caption": "10  radarChart - six spokes, two series"},
    {"key": "stockChart", "build": stock, "frame": WIDE,
     "caption": "11  stockChart - high/low/close with hiLowLines"},
    {"key": "surfaceChart", "build": surface, "frame": MEDIUM,
     "caption": "12  surfaceChart - a lit mesh banded by value, with a legend of bands"},
    {"key": "bar3DChart", "build": bar_3d, "frame": MEDIUM,
     "caption": "13  bar3DChart - parses as barChart and draws flat"},
    {"key": "line3DChart", "build": line_3d, "frame": MEDIUM,
     "caption": "14  line3DChart - parses as lineChart and draws flat"},
    {"key": "pie3DChart", "build": pie_3d, "frame": SQUARE,
     "caption": "15  pie3DChart - parses as pieChart and draws flat"},
    {"key": "area3DChart", "build": area_3d, "frame": MEDIUM,
     "caption": "16  area3DChart - parses as areaChart and draws flat"},
    {"key": "combo", "build": combo, "frame": WIDE,
     "caption": "17  combo - barChart plus lineChart on a secondary value axis"},
]


def caption_shape(index: int, text: str) -> str:
    """A plain text box naming the slide.  Theme minor font, stated size, no ``a:latin``.

    Nothing here names a typeface, so ``fidelity.requested_faces`` sees only what the
    theme names -- Aptos and Aptos Display -- and the deck stays scorable.
    """
    return (
        "<p:sp><p:nvSpPr>"
        f"<p:cNvPr id='{10 + index}' name='Caption'/><p:cNvSpPr txBox='1'/><p:nvPr/>"
        "</p:nvSpPr><p:spPr>"
        f"<a:xfrm><a:off x='{CAPTION_OFF[0]}' y='{CAPTION_OFF[1]}'/>"
        f"<a:ext cx='{CAPTION_EXT[0]}' cy='{CAPTION_EXT[1]}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>"
        "<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>"
        "<p:txBody><a:bodyPr wrap='square' lIns='0' rIns='0' tIns='0' bIns='0'/>"
        "<a:lstStyle/><a:p>"
        "<a:r><a:rPr lang='en-US' sz='1000' b='0'>"
        "<a:solidFill><a:srgbClr val='334155'/></a:solidFill></a:rPr>"
        f"<a:t>{escape(text)}</a:t></a:r><a:endParaRPr lang='en-US' sz='1000'/>"
        "</a:p></p:txBody></p:sp>"
    )


def slide_xml(index: int, spec: dict) -> str:
    cx, cy = spec["frame"]
    uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    graphic_data = f"<c:chart {C} {R} r:id='rIdChart'/>"
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + caption_shape(index, spec["caption"])
        + f"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='{100 + index}' "
        f"name='{escape(spec['key'])}'/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        f"<p:xfrm><a:off x='{FRAME_OFF[0]}' y='{FRAME_OFF[1]}'/>"
        f"<a:ext cx='{cx}' cy='{cy}'/></p:xfrm>"
        f"<a:graphic><a:graphicData uri='{uri}'>{graphic_data}"
        "</a:graphicData></a:graphic></p:graphicFrame>"
        "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


def slide_rels(index: int) -> str:
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        f"<Relationship Id='rId1' Type='{LAYOUT_REL}' Target='../slideLayouts/slideLayout1.xml'/>"
        f"<Relationship Id='rIdChart' Type='{CHART_REL}' "
        f"Target='../charts/chart{index + 1}.xml'/>"
        "</Relationships>"
    )


def blank_template(xml: str) -> str:
    """Empty a layout or master so nothing but the chart and its caption is drawn."""
    xml = re.sub(r"<p:spTree>.*?</p:spTree>", EMPTY_TREE, xml, flags=re.S)
    return re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)


def write_deck(target: Path) -> None:
    source = zipfile.ZipFile(SOURCE)
    presentation = source.read("ppt/presentation.xml").decode()
    pres_rels = source.read("ppt/_rels/presentation.xml.rels").decode()
    content_types = source.read("[Content_Types].xml").decode()
    numbers = range(1, len(SLIDES) + 1)

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
        r'<Override PartName="/ppt/(slides/slide|charts/chart)\d+\.xml"[^>]*/>',
        "",
        content_types,
    )
    # The source's embedded workbook and its picture are not carried over, so neither is
    # the content-type Override that named the workbook -- an Override for a part that is
    # not in the package is not valid OPC.
    content_types = re.sub(
        r'<Override PartName="/ppt/embeddings/[^>]*/>', "", content_types
    )
    content_types = content_types.replace(
        "</Types>",
        "".join(
            f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{SLIDE_TYPE}"/>'
            f'<Override PartName="/ppt/charts/chart{n}.xml" ContentType="{CHART_TYPE}"/>'
            for n in numbers
        )
        + "</Types>",
    )

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            name = item.filename
            # The source's own slide, chart, embedded workbook and picture go; this deck
            # keeps only its package skeleton.
            if name.startswith(
                ("ppt/slides/", "ppt/charts/", "ppt/embeddings/", "ppt/media/")
            ):
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
        for index, spec in enumerate(SLIDES):
            out.writestr(f"ppt/slides/slide{index + 1}.xml", slide_xml(index, spec))
            out.writestr(f"ppt/slides/_rels/slide{index + 1}.xml.rels", slide_rels(index))
            out.writestr(f"ppt/charts/chart{index + 1}.xml", spec["build"]())


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else TARGET
    write_deck(target)
    for index, spec in enumerate(SLIDES, start=1):
        width, height = spec["frame"]
        print(
            f"{index:3d}  {spec['key']:16s} "
            f"{width / EMU_PER_POINT:.0f}x{height / EMU_PER_POINT:.0f} pt"
        )
    print(f"\n{len(SLIDES)} slides -> {target} ({target.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

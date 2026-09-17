#!/usr/bin/env python3
"""Build decks that measure a ``bubbleChart``'s **value axis domain**.

``chart-gallery`` slide 6 runs both axes 0..12 by 2 where this library runs them 0..10 by
1 over the same data, and ROADMAP.md 3.2a's diagnosis is that PowerPoint pads for the
drawn **radii** where we pad for the centres.  That reading makes the domain a function of
the bubble sizes, which is a circle worth naming before writing any code: a bubble's
diameter is fixed by the *region* -- the frame less its insets, title and legend band, and
emphatically not the plot rectangle (:data:`~pptx2svg.resolve.chart.BUBBLE_REGION_INSET_PT`)
-- so the radii are known before the axis is, and the circle closes after one step rather
than never.  What is *not* known is what the axis then does with them, and that is what
these decks are for.

The discriminating experiment is ``c:bubbleScale``.  It changes every radius and **no data
value at all**, so a domain that moves with it is padding for ink and a domain that does
not is padding for numbers -- and no reading of "pad the data by a percentage" survives
either way.

| family | holds | varies | settles |
| --- | --- | --- | --- |
| ``z`` | data, frame | ``c:bubbleScale`` 10..300 | that the axis follows the *ink* |
| ``w`` | data, scale | frame width and height | which length the pad is measured in |
| ``d`` | frame, scale | the extreme point's own size | per-point or whole-chart |
| ``m`` | frame, scale | a non-zero minimum, and negatives | what the low end does |
| ``e`` | frame, scale | data maxima across a nice-number step | where the rounding lands |
| ``v`` | -- | fresh combinations | validation, fitted on none of it |

Every slide states both axes explicitly with ``c:majorGridlines`` on the y axis, so the
drawn domain comes back as tick **labels** -- text, which the export carries exactly --
rather than having to be inferred from where a bubble landed.

Usage -- the deck's file name picks its probe table::

    python3 tools/make_bubble_probe.py ~/pptx2svg-oracle/bubble-axis.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/bubble-axis.pptx ~/pptx2svg-oracle/bubble-axis.pdf
    python3 tools/read_bubble_probe.py ~/pptx2svg-oracle/bubble-axis.pdf

**No ``c:bubble3D`` anywhere**: ``<c:bubble3D val="1"/>`` hangs PowerPoint 16.x on macOS
reproducibly (ROADMAP.md 0.1).  The decks and their exports are throwaway and are not
committed; ``~/pptx2svg-oracle`` is left holding its eighteen corpus files.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SOURCE = ROOT / "tests/fixtures/authoring-integration.pptx"

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

EMU = 12700
FRAME_OFF = (18 * EMU, 18 * EMU)

EMPTY_TREE = (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
)

X_AXIS = 120000
Y_AXIS = 120001

#: The gallery's own two series, which is the reading being chased.
GALLERY = [
    {"name": "Region", "x": (2.0, 4.0, 6.0, 8.0, 10.0), "y": (5.0, 8.0, 4.0, 9.0, 6.0),
     "size": (14.0, 32.0, 9.0, 40.0, 21.0)},
    {"name": "Channel", "x": (3.0, 5.5, 7.0, 9.5), "y": (2.0, 6.5, 7.5, 3.5),
     "size": (25.0, 12.0, 30.0, 18.0)},
]

#: One series, four points, whose largest bubble is **not** at either extreme.  A rule that
#: pads by the biggest radius in the chart and a rule that pads by the extreme point's own
#: radius differ here and nowhere else.
INNER = [
    {"name": "Inner", "x": (2.0, 5.0, 8.0, 9.0), "y": (2.0, 9.0, 5.0, 3.0),
     "size": (5.0, 5.0, 100.0, 5.0)},
]
OUTER = [
    {"name": "Outer", "x": (2.0, 5.0, 8.0, 9.0), "y": (2.0, 9.0, 5.0, 3.0),
     "size": (5.0, 100.0, 5.0, 5.0)},
]


def _nums(tag: str, values) -> str:
    inner = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
    return (
        f"<c:{tag}><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{inner}</c:numCache></c:numRef></c:{tag}>"
    )


def _series(index: int, spec: dict) -> str:
    """One ``c:ser``.  The schema's order is ``idx, order, tx, xVal, yVal, bubbleSize``."""
    return (
        f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        f"<c:pt idx='0'><c:v>{spec['name']}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        + _nums("xVal", spec["x"])
        + _nums("yVal", spec["y"])
        + _nums("bubbleSize", spec["size"])
        + "</c:ser>"
    )


def _value_axis(ax_id: int, cross_id: int, position: str, gridlines: bool) -> str:
    return (
        f"<c:valAx><c:axId val='{ax_id}'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='0'/><c:axPos val='{position}'/>"
        + ("<c:majorGridlines/>" if gridlines else "")
        + "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        f"<c:crossAx val='{cross_id}'/><c:crosses val='autoZero'/>"
        "<c:crossBetween val='midCat'/></c:valAx>"
    )


def chart_xml(probe: dict) -> bytes:
    series = "".join(_series(i, spec) for i, spec in enumerate(probe["series"]))
    groups = (
        "<c:bubbleChart><c:varyColors val='0'/>"
        + series
        + f"<c:bubbleScale val='{probe.get('scale', 100)}'/>"
        "<c:showNegBubbles val='0'/>"
        + (f"<c:sizeRepresents val='{probe['represents']}'/>" if probe.get("represents") else "")
        + f"<c:axId val='{X_AXIS}'/><c:axId val='{Y_AXIS}'/></c:bubbleChart>"
    )
    axes = _value_axis(X_AXIS, Y_AXIS, "b", False) + _value_axis(Y_AXIS, X_AXIS, "l", True)
    legend = probe.get("legend")
    legend_xml = (
        f"<c:legend><c:legendPos val='{legend}'/><c:overlay val='0'/></c:legend>"
        if legend
        else ""
    )
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        + groups
        + axes
        + "</c:plotArea>"
        + legend_xml
        + "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        "<a:defRPr sz='1000'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
        "</c:chartSpace>"
    ).encode()


# --------------------------------------------------------------------------------------
# The probe tables
# --------------------------------------------------------------------------------------


def _frame(width_pt: float, height_pt: float = 260.0) -> tuple[int, int]:
    return (int(round(width_pt * EMU)), int(round(height_pt * EMU)))


WIDE = _frame(480.0, 260.0)


def probe(key: str, series=None, **extra) -> dict:
    return {"key": key, "frame": WIDE, "series": series or GALLERY, **extra}


def scaled(name: str, high_y: float, high_x: float, size: float) -> list[dict]:
    """One series of four points whose maxima are stated, for the rounding sweep."""
    return [
        {
            "name": name,
            "x": (high_x * 0.2, high_x * 0.55, high_x, high_x * 0.8),
            "y": (high_y * 0.3, high_y, high_y * 0.5, high_y * 0.75),
            "size": (size * 0.25, size, size * 0.5, size * 0.75),
        }
    ]


AXIS_PROBES: list[dict] = [
    # **The discriminating experiment**: one chart, nine radii, no data value touched.
    *[probe(f"z-{scale:03d}", scale=scale) for scale in (10, 25, 50, 75, 100, 150, 200, 300)],
    # **Which length the pad is in.**  The plot is what a value maps onto; the region is
    # what the radius comes from.  Moving the frame moves both, but not in step.
    *[
        probe(f"w-{width:.0f}x{height:.0f}", frame=_frame(width, height))
        for width, height in (
            (480.0, 160.0), (480.0, 200.0), (480.0, 340.0),
            (300.0, 260.0), (660.0, 260.0), (300.0, 160.0),
        )
    ],
    # **Per point or per chart.**  ``d-inner`` puts the 100-size bubble in the middle of
    # both ranges and ``d-outer`` puts it on the y maximum; the data is otherwise identical.
    probe("d-inner", INNER),
    probe("d-outer", OUTER),
    probe("d-inner-l", INNER, scale=200),
    probe("d-outer-l", OUTER, scale=200),
    # **The low end.**  A bubble at the minimum hangs below it by its own radius, so if the
    # axis clears the ink it has to move down as well as up -- unless it is pinned to zero.
    probe("m-lift", [{"name": "Lift", "x": (4.0, 6.0, 8.0, 10.0),
                      "y": (22.0, 25.0, 28.0, 30.0), "size": (10.0, 40.0, 20.0, 30.0)}]),
    probe("m-lift-l", [{"name": "Lift", "x": (4.0, 6.0, 8.0, 10.0),
                        "y": (22.0, 25.0, 28.0, 30.0), "size": (10.0, 40.0, 20.0, 30.0)}],
          scale=250),
    probe("m-neg", [{"name": "Neg", "x": (-4.0, -1.0, 2.0, 5.0),
                     "y": (-6.0, 3.0, 8.0, -2.0), "size": (10.0, 40.0, 20.0, 30.0)}]),
    probe("m-zero", [{"name": "Zero", "x": (0.0, 3.0, 6.0, 9.0),
                      "y": (0.0, 4.0, 9.0, 2.0), "size": (40.0, 10.0, 20.0, 30.0)}]),
    # **Where the rounding lands.**  Same shape of data at maxima that walk across a nice
    # step, so the padded value and the drawn domain can be compared digit by digit.
    *[probe(f"e-{high:g}", scaled("E", high, 10.0, 40.0)) for high in (9.0, 9.5, 10.0, 11.0, 12.0)],
    *[probe(f"eb-{high:g}", scaled("E", high, 10.0, 40.0), scale=300) for high in (9.0, 10.0, 12.0)],
    *[probe(f"ec-{high:g}", scaled("E", high, 10.0, 40.0), scale=20) for high in (9.0, 10.0, 12.0)],
    # A decade away in each direction, to check the rule is scale-free in the data.
    probe("e-small", scaled("E", 0.9, 1.0, 40.0)),
    probe("e-big", scaled("E", 900.0, 1000.0, 40.0)),
    # **Validation**, on combinations nothing above is fitted to.
    probe("v-legend", legend="b"),
    probe("v-legend-r", legend="r"),
    probe("v-width", represents="w"),
    probe("v-mixed", GALLERY, scale=175, frame=_frame(360.0, 300.0)),
    probe("v-tall", INNER, scale=60, frame=_frame(240.0, 340.0)),
]

DECKS = {"bubble-axis": AXIS_PROBES}


# --------------------------------------------------------------------------------------
# Deck assembly -- the same shape as tools/make_legend_probe.py
# --------------------------------------------------------------------------------------


def slide_xml(index: int, probe_: dict) -> str:
    cx, cy = probe_["frame"]
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='{100 + index}' "
        f"name='{probe_['key']}'/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
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
    xml = re.sub(r"<p:spTree>.*?</p:spTree>", EMPTY_TREE, xml, flags=re.S)
    return re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)


def probes_for(target: Path) -> list[dict]:
    for name, table in DECKS.items():
        if name in target.stem:
            return table
    raise SystemExit(f"name the deck after one of {sorted(DECKS)}, not {target.stem!r}")


def write_deck(target: Path, probes: list[dict]) -> None:
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
        for index, item in enumerate(probes):
            out.writestr(f"ppt/slides/slide{index + 1}.xml", slide_xml(index, item))
            out.writestr(f"ppt/slides/_rels/slide{index + 1}.xml.rels", slide_rels(index))
            out.writestr(f"ppt/charts/probe{index}.xml", chart_xml(item))


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    probes = probes_for(target)
    write_deck(target, probes)
    print(f"wrote {target} -- {len(probes)} slides")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

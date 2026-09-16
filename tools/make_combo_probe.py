#!/usr/bin/env python3
"""Build decks of **combo** charts -- several ``c:*Chart`` groups in one plot area.

ROADMAP.md 3.3 lists six questions a combo chart raises and says to measure rather than
infer every one of them.  This file is the deck side of that; ``tools/read_combo_probe.py``
is the reading side.

The questions, and the family of probes that answers each:

``o``  **Draw order.**  Which group paints on top when two overlap.  Document order and a
       type precedence are indistinguishable on one deck, so every pair is drawn both
       ways round: ``o-col-line`` and ``o-line-col`` are the same two groups with their
       positions in ``c:plotArea`` swapped.
``d``  **Which series feed which axis.**  A value axis' domain should come from the series
       attached to *it* by ``c:axId``.  Each probe holds the primary group at 0..9 and
       moves the secondary group's range by decades; a shared domain moves the left axis
       with it and a per-axis domain does not.  ``d-both`` is the control that must move.
``s``  **The interval rule on the secondary axis.**  The N-meter from ``axis-rung``: five
       datasets whose 1-2-5 boundaries fall at different interval counts, so the drawn
       unit reads the count back exactly.  Swept over five frame heights, with the primary
       group held at dataset B -- which reads the *primary* count off the same slide, so
       every slide is two readings and the two axes can be compared directly.
``p``  **What the secondary axis does to the plot rectangle.**  Read from the gridlines'
       own x extent, against the same chart with no secondary axis at all.
``g``  **Bar geometry across groups.**  Whether a bar group that shares a plot with a line
       group keeps the whole category band, and what two bar groups do to each other.
``l``  **Legend composition.**  Entry order and count across groups.

Usage -- the deck's file name picks its probe table::

    python3 tools/make_combo_probe.py ~/pptx2svg-oracle/combo-order.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/combo-order.pptx ~/pptx2svg-oracle/combo-order.pdf
    python3 tools/read_combo_probe.py ~/pptx2svg-oracle/combo-order.pdf

The decks and their exports are throwaway and are **not** committed; ``~/pptx2svg-oracle``
is the directory PowerPoint is allowed to write to and it is left holding its eighteen
corpus files.  See ROADMAP.md section 0.1 before blaming a failed export on the path.
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

#: The four axis ids.  The secondary pair is what ``c:axId`` on a second group names, and
#: the secondary *category* axis is deleted, which is what PowerPoint itself writes.
PRIMARY_CAT = 90000
PRIMARY_VAL = 90001
SECOND_CAT = 90002
SECOND_VAL = 90003

CATEGORIES = ("C1", "C2", "C3", "C4", "C5")

#: The N-meter from ``axis-rung``: five datasets whose 1-2-5 boundaries fall at different
#: interval counts, so one drawn unit names one count.  See ROADMAP.md 0.4.
METER = (("A", 5.0), ("B", 9.0), ("C", 3.8), ("D", 6.6), ("E", 8.5))


def values_for(high: float, count: int = 5) -> list[float]:
    """Five values reaching *high* exactly once, in the shape every axis probe uses."""
    factors = (0.3, 1.0, 0.55, 0.8, 0.45)
    out = [high * factors[i % len(factors)] for i in range(count)]
    out[1] = high
    return out


# --------------------------------------------------------------------------------------
# Chart XML
# --------------------------------------------------------------------------------------


def _points(values) -> str:
    return "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))


def _cats(names) -> str:
    inner = "".join(f"<c:pt idx='{i}'><c:v>{n}</c:v></c:pt>" for i, n in enumerate(names))
    return (
        "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='{len(names)}'/>{inner}</c:strCache></c:strRef></c:cat>"
    )


def _vals(values) -> str:
    return (
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{_points(values)}</c:numCache></c:numRef></c:val>"
    )


def _series(index: int, name: str, values, *, smooth: bool) -> str:
    """One ``c:ser``.  **``c:idx`` must be unique across every group in the plot area** --
    two series sharing ``idx=0`` hangs PowerPoint 16.x on macOS (ROADMAP.md 0.1)."""
    return (
        f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        f"<c:pt idx='0'><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        + _cats(CATEGORIES)
        + _vals(values)
        + ("<c:smooth val='0'/>" if smooth else "")
        + "</c:ser>"
    )


def group_xml(group: dict, first_index: int) -> tuple[str, int]:
    """One ``c:*Chart`` element, and the next free series index.

    ``group`` keys: ``kind`` (``col``/``bar``/``line``/``area``), ``highs`` (one value per
    series), ``secondary``, ``gap``, ``overlap``, ``names``.
    """
    kind = group["kind"]
    highs = group["highs"]
    names = group.get("names") or [f"S{first_index + i}" for i in range(len(highs))]
    smooth = kind == "line"
    body = "".join(
        _series(first_index + i, names[i], values_for(high), smooth=smooth)
        for i, high in enumerate(highs)
    )
    cat_id = SECOND_CAT if group.get("secondary") else PRIMARY_CAT
    val_id = SECOND_VAL if group.get("secondary") else PRIMARY_VAL
    ids = f"<c:axId val='{cat_id}'/><c:axId val='{val_id}'/>"
    if kind in ("col", "bar"):
        tail = ""
        if group.get("gap") is not None:
            tail += f"<c:gapWidth val='{group['gap']}'/>"
        if group.get("overlap") is not None:
            tail += f"<c:overlap val='{group['overlap']}'/>"
        xml = (
            f"<c:barChart><c:barDir val='{kind}'/>"
            f"<c:grouping val='{group.get('grouping', 'clustered')}'/>"
            f"<c:varyColors val='0'/>{body}{tail}{ids}</c:barChart>"
        )
    elif kind == "line":
        xml = (
            f"<c:lineChart><c:grouping val='{group.get('grouping', 'standard')}'/>"
            f"<c:varyColors val='0'/>{body}<c:marker val='1'/>{ids}</c:lineChart>"
        )
    elif kind == "area":
        xml = (
            f"<c:areaChart><c:grouping val='{group.get('grouping', 'standard')}'/>"
            f"<c:varyColors val='0'/>{body}{ids}</c:areaChart>"
        )
    else:
        raise SystemExit(f"unknown group kind {kind!r}")
    return xml, first_index + len(highs)


def cat_axis(ax_id: int, cross_id: int, *, delete: bool = False, position: str = "b") -> str:
    return (
        f"<c:catAx><c:axId val='{ax_id}'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if delete else 0}'/><c:axPos val='{position}'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
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
    crosses: str = "autoZero",
) -> str:
    return (
        f"<c:valAx><c:axId val='{ax_id}'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if delete else 0}'/><c:axPos val='{position}'/>"
        + ("<c:majorGridlines/>" if gridlines else "")
        + "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        f"<c:crossAx val='{cross_id}'/><c:crosses val='{crosses}'/>"
        "<c:crossBetween val='between'/></c:valAx>"
    )


def chart_xml(probe: dict) -> bytes:
    groups_xml = ""
    index = 0
    for group in probe["groups"]:
        xml, index = group_xml(group, index)
        groups_xml += xml
    secondary = any(group.get("secondary") for group in probe["groups"])

    axes = cat_axis(PRIMARY_CAT, PRIMARY_VAL) + val_axis(
        PRIMARY_VAL, PRIMARY_CAT, gridlines=probe.get("primary_gridlines", True)
    )
    if secondary:
        axes += val_axis(
            SECOND_VAL,
            SECOND_CAT,
            position=probe.get("second_pos", "r"),
            delete=probe.get("second_delete", False),
            gridlines=probe.get("second_gridlines", False),
            crosses=probe.get("second_crosses", "max"),
        )
        axes += cat_axis(SECOND_CAT, SECOND_VAL, delete=True, position="b")

    head = (
        (
            "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>"
            f"<a:t>{probe['title']}</a:t></a:r></a:p></c:rich></c:tx>"
            "<c:overlay val='0'/></c:title><c:autoTitleDeleted val='0'/>"
        )
        if probe.get("title")
        else "<c:autoTitleDeleted val='1'/>"
    )
    legend = (
        f"<c:legend><c:legendPos val='{probe['legend']}'/><c:overlay val='0'/></c:legend>"
        if probe.get("legend")
        else ""
    )
    size = probe.get("size", 1000)
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
        "<c:chart>" + head + "<c:plotArea><c:layout/>"
        + groups_xml
        + axes
        + "</c:plotArea>"
        + legend
        + "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
        "</c:chartSpace>"
    ).encode()


# --------------------------------------------------------------------------------------
# The probe tables
# --------------------------------------------------------------------------------------

WIDE = (480 * EMU, 260 * EMU)


def _frame(width_pt: float, height_pt: float) -> tuple[int, int]:
    return (int(round(width_pt * EMU)), int(round(height_pt * EMU)))


def _pair(first: dict, second: dict, key: str, **extra) -> dict:
    return {"key": key, "frame": WIDE, "groups": [first, second], **extra}


def col(highs, **extra) -> dict:
    return {"kind": "col", "highs": highs, **extra}


def line(highs, **extra) -> dict:
    return {"kind": "line", "highs": highs, **extra}


def area(highs, **extra) -> dict:
    return {"kind": "area", "highs": highs, **extra}


#: **Draw order.**  Every pair twice, with the groups swapped, which is what separates
#: document order from a precedence by type.  The two groups share the primary axis and
#: the same data range, so whatever is drawn second covers the first.
ORDER_PROBES: list[dict] = [
    _pair(col([9.0]), line([9.0]), "o-col-line"),
    _pair(line([9.0]), col([9.0]), "o-line-col"),
    _pair(col([9.0]), area([9.0]), "o-col-area"),
    _pair(area([9.0]), col([9.0]), "o-area-col"),
    _pair(line([9.0]), area([9.0]), "o-line-area"),
    _pair(area([9.0]), line([9.0]), "o-area-line"),
    # The same two, with the second group on the secondary axis -- the real combo shape.
    _pair(col([9.0]), line([9.0], secondary=True), "o-col-line2"),
    _pair(line([9.0]), col([9.0], secondary=True), "o-line-col2"),
    # Two bar groups, which is the case where "which is on top" is also "how wide".
    _pair(col([9.0]), col([9.0], secondary=True), "o-col-col2"),
]

#: **Which series feed which axis.**  The primary group is held at 0..9 on every slide.
#: A left axis that moves with the secondary group's range is a shared domain.
DOMAIN_PROBES: list[dict] = [
    {"key": "d-alone", "frame": WIDE, "groups": [col([9.0])]},
    *[
        _pair(col([9.0]), line([high], secondary=True), f"d-sec{high:g}")
        for high in (5.0, 50.0, 500.0, 0.5)
    ],
    # The control: both groups on the *primary* axis, where one domain is correct.
    *[
        _pair(col([9.0]), line([high]), f"d-shared{high:g}")
        for high in (50.0, 500.0)
    ],
    # And the mirror -- a big primary against a small secondary -- which says whether the
    # *right* axis reads the left group.
    _pair(col([500.0]), line([9.0], secondary=True), "d-big-small"),
]

#: **The interval rule on the secondary axis.**  Five meter datasets on the secondary
#: group against a held dataset B on the primary, swept over five frame heights.  Each
#: slide reads both counts, so "are the two axes independent" and "does the secondary obey
#: the frame-derived rule" come off the same 25 readings.
RUNG_HEIGHTS = (60.0, 90.0, 120.0, 150.0, 180.0)

SIDE_PROBES: list[dict] = [
    _pair(
        col([9.0]),
        line([high], secondary=True),
        f"s{height:.0f}-{tag}",
        frame=_frame(480.0, height),
    )
    for height in RUNG_HEIGHTS
    for tag, high in METER
]

#: The same sweep with the **primary** group carrying the meter and the secondary held,
#: which is the check that the two axes are read the same way round.
SIDE_PROBES += [
    _pair(
        col([high]),
        line([9.0], secondary=True),
        f"sp{height:.0f}-{tag}",
        frame=_frame(480.0, height),
    )
    for height in (90.0, 150.0)
    for tag, high in METER
]

#: **What the secondary axis does to the plot rectangle**, and where it lands.
PLOT_PROBES: list[dict] = [
    {"key": "p-none", "frame": WIDE, "groups": [col([9.0])]},
    _pair(col([9.0]), line([9.0], secondary=True), "p-sec"),
    # Labels a million times wider: a right inset that follows the label width moves.
    _pair(col([9.0]), line([9.0e6], secondary=True), "p-secwide"),
    _pair(col([9.0]), line([9.0e3], secondary=True), "p-secmid"),
    # A deleted secondary axis still has series on it; does it still take the band?
    _pair(col([9.0]), line([9.0], secondary=True), "p-secdel", second_delete=True),
    # `crosses=autoZero` rather than `max` -- the schema's other answer for where it goes.
    _pair(col([9.0]), line([9.0], secondary=True), "p-seczero", second_crosses="autoZero"),
    # Gridlines on the secondary axis as well as the primary: do they both draw?
    _pair(col([9.0]), line([9.0], secondary=True), "p-secgrid", second_gridlines=True),
    # A secondary axis that PowerPoint would put at the left instead.
    _pair(col([9.0]), line([9.0], secondary=True), "p-secleft", second_pos="l"),
    # Two groups on the primary axis only: the control for a plot rect with no second axis.
    _pair(col([9.0]), line([9.0]), "p-twoprimary"),
]

#: **Bar geometry across groups.**  The bar's drawn width against the category band is
#: what says how many slots the band was divided into.
BAR_PROBES: list[dict] = [
    {"key": "g-one", "frame": WIDE, "groups": [col([9.0], gap=150)]},
    {"key": "g-two", "frame": WIDE, "groups": [col([9.0, 7.0], gap=150)]},
    _pair(col([9.0], gap=150), line([9.0], secondary=True), "g-bar-line"),
    _pair(col([9.0], gap=150), line([9.0]), "g-bar-line-primary"),
    _pair(col([9.0], gap=150), area([9.0], secondary=True), "g-bar-area"),
    # Two bar groups: clustered into one band, or drawn over each other?
    _pair(col([9.0], gap=150), col([7.0], gap=150), "g-bar-bar"),
    _pair(col([9.0], gap=150), col([7.0], gap=150, secondary=True), "g-bar-bar2"),
    # Disagreeing gap widths -- which group's wins.
    _pair(col([9.0], gap=50), col([7.0], gap=300, secondary=True), "g-gap-clash"),
    _pair(col([9.0], gap=300), line([9.0], secondary=True), "g-gap300"),
    _pair(col([9.0], gap=50), line([9.0], secondary=True), "g-gap50"),
    # Overlap stated on the bar group of a combo.
    _pair(col([9.0, 7.0], gap=150, overlap=-27), line([9.0], secondary=True), "g-overlap"),
]

#: **Legend composition.**  The entries are read by their own x positions, so the order is
#: measured rather than assumed.
LEGEND_PROBES: list[dict] = [
    _pair(col([9.0, 7.0], names=["Bar one", "Bar two"]),
          line([9.0], secondary=True, names=["Line one"]),
          "l-bottom", legend="b"),
    _pair(line([9.0], names=["Line one"]),
          col([9.0, 7.0], secondary=True, names=["Bar one", "Bar two"]),
          "l-reversed", legend="b"),
    _pair(col([9.0, 7.0], names=["Bar one", "Bar two"]),
          line([9.0], secondary=True, names=["Line one"]),
          "l-right", legend="r"),
    _pair(col([9.0], names=["Bar one"]),
          line([9.0], secondary=True, names=["Line one"]),
          "l-title", legend="b", title="Combo probe"),
]

DECKS = {
    "combo-order": ORDER_PROBES,
    "combo-domain": DOMAIN_PROBES,
    "combo-side": SIDE_PROBES,
    "combo-plot": PLOT_PROBES,
    "combo-bar": BAR_PROBES,
    "combo-legend": LEGEND_PROBES,
    "combo-all": (
        ORDER_PROBES + DOMAIN_PROBES + PLOT_PROBES + BAR_PROBES + LEGEND_PROBES
    ),
}


# --------------------------------------------------------------------------------------
# Deck assembly -- the same shape as tools/make_axis_probe.py
# --------------------------------------------------------------------------------------


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
        for index, probe in enumerate(probes):
            out.writestr(f"ppt/slides/slide{index + 1}.xml", slide_xml(index, probe))
            out.writestr(f"ppt/slides/_rels/slide{index + 1}.xml.rels", slide_rels(index))
            out.writestr(f"ppt/charts/probe{index}.xml", chart_xml(probe))


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

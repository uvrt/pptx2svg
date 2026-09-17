#!/usr/bin/env python3
"""Build decks that measure the **horizontal legend's inter-entry gap**.

``LEGEND_ENTRY_GAP_EM = 0.5`` in ``resolve/chart.py`` was refuted by ROADMAP.md 3.3: four
charts solved for four gaps -- 0.77, 1.03, 1.12 and 1.15 em -- and the gap is constant
*within* each chart, so it is not a per-entry quantity and not a measurement error.  What
it is a function of is the question these decks were built to answer, and **they answer
it**: see ROADMAP.md 3.5 and :data:`~pptx2svg.resolve.chart.LEGEND_ENTRY_SLACK`.  The run
is padded by a fifth of its own natural width and that slack is cut into ``n + 1`` equal
gaps, so the gap grows with the entries and the frame does not enter at all.

The reading is geometric rather than typographic.  Every entry's **key** is a path -- a
filled swatch for a bar, a horizontal rule for a line -- so the pitch between consecutive
keys comes off the export with no font metric in it at all.  One pitch is

    pitch(i) = key + key_gap + advance(name(i)) + G

and since ``key + key_gap`` is one number per chart, the only unknown a slide leaves is
``key + key_gap + G``, which is what ``read_legend_probe.py`` calls **K**.  Whether the
``G`` inside it is a constant is exactly the open question.

Two readings were in play when these decks were written, and one experiment separates them:

* **packed** -- entries are laid out with a fixed gap and the run is then centred;
* **distributed** -- entries are spread over an available width, so the "gap" is the
  residue ``(available - sum of widths) / (n - 1)`` and differs per chart by construction.

Hold a chart fixed and **lengthen one entry's name by delta**.  A packed layout leaves the
pitch between two *unchanged* entries alone and moves the first key left by delta/2; a
distributed one shrinks that pitch by ``delta / (n - 1)`` and leaves the first key where it
was.  The ``x-`` family is that experiment; the rest of the sweep says what the answer is
a function of.  **Neither reading was right**: the pitch between two untouched entries
*grew*, which is the packed sign and the distributed one's opposite, and the ``f`` family
then showed the frame does not enter at all.

| family | holds | varies |
| --- | --- | --- |
| ``x`` | frame, count, key | one name's length -- the discriminating experiment |
| ``n`` | frame, names, key | entry count, 2 through 7 |
| ``w`` | frame, count, key | total name width, including one long entry among short |
| ``k`` | frame, count, names | key type: bar, line, marker-only, area, pie, scatter, combo |
| ``f`` | count, names, key | frame width, 240 through 720 pt |
| ``z`` | frame, count, names | font size, 8 through 18 pt -- the answer is stated in ems |
| ``p`` | everything | ``legendPos``: ``b``, ``t``, ``tr`` |

Usage -- the deck's file name picks its probe table::

    python3 tools/make_legend_probe.py ~/pptx2svg-oracle/legend-pack.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/legend-pack.pptx ~/pptx2svg-oracle/legend-pack.pdf
    python3 tools/read_legend_probe.py ~/pptx2svg-oracle/legend-pack.pdf

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

PRIMARY_CAT = 90000
PRIMARY_VAL = 90001
SECOND_CAT = 90002
SECOND_VAL = 90003

CATEGORIES = ("C1", "C2", "C3", "C4", "C5")

#: Entry names.  **Every one begins with the same letter**, so a pitch read off the text's
#: ink box carries no left-side-bearing difference and can be compared with the pitch read
#: off the keys.  The keys are the primary reading; this only keeps the cross-check honest.
NAMES = [f"W{'i' * n}" for n in range(1, 9)]

#: A long name and a short one, for the width sweep and the discriminating experiment.
SHORT = "Wm"
LONG = "Wmmmmmmmmmmmm"


def values_for(high: float, count: int = 5) -> list[float]:
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


def _vals(values, tag: str = "val") -> str:
    return (
        f"<c:{tag}><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{_points(values)}</c:numCache></c:numRef></c:{tag}>"
    )


def _name(text: str) -> str:
    return (
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        f"<c:pt idx='0'><c:v>{text}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
    )


def _series(index: int, name: str, values, cats, *, kind: str, marker: bool) -> str:
    """One ``c:ser``.  ``c:idx`` must be unique across every group (ROADMAP.md 0.1).

    The child order is the schema's and is not negotiable: ``idx``, ``order``, ``tx``,
    ``spPr``, ``marker``, ``cat``, ``val``, ``smooth``.  Out-of-order children cost an
    earlier probe a -9074.
    """
    head = f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>" + _name(name)
    if kind in ("line", "scatter"):
        head += (
            "<c:marker><c:symbol val='circle'/><c:size val='5'/></c:marker>"
            if marker
            else "<c:marker><c:symbol val='none'/></c:marker>"
        )
    if kind == "scatter":
        xs = [float(i + 1) for i in range(len(values))]
        return head + _vals(xs, "xVal") + _vals(values, "yVal") + "<c:smooth val='0'/></c:ser>"
    body = _cats(cats) + _vals(values)
    tail = "<c:smooth val='0'/>" if kind == "line" else ""
    return head + body + tail + "</c:ser>"


def group_xml(group: dict, first_index: int, cats) -> tuple[str, int]:
    """One ``c:*Chart`` element, and the next free series index."""
    kind = group["kind"]
    names = group["names"]
    highs = group.get("highs") or [9.0] * len(names)
    marker = group.get("marker", True)
    body = "".join(
        _series(first_index + i, names[i], values_for(highs[i]), cats,
                kind="scatter" if kind == "scatter" else ("line" if kind == "line" else kind),
                marker=marker)
        for i in range(len(names))
    )
    cat_id = SECOND_CAT if group.get("secondary") else PRIMARY_CAT
    val_id = SECOND_VAL if group.get("secondary") else PRIMARY_VAL
    ids = f"<c:axId val='{cat_id}'/><c:axId val='{val_id}'/>"
    if kind == "col":
        xml = (
            "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
            f"<c:varyColors val='0'/>{body}<c:gapWidth val='150'/>{ids}</c:barChart>"
        )
    elif kind == "line":
        xml = (
            "<c:lineChart><c:grouping val='standard'/>"
            f"<c:varyColors val='0'/>{body}<c:marker val='{1 if marker else 0}'/>"
            f"{ids}</c:lineChart>"
        )
    elif kind == "area":
        xml = (
            "<c:areaChart><c:grouping val='standard'/>"
            f"<c:varyColors val='0'/>{body}{ids}</c:areaChart>"
        )
    elif kind == "pie":
        # A pie legends its **categories**, so its one series carries every entry name in
        # ``c:cat`` and the series name is irrelevant.  No axes.
        series = _series(first_index, "Pie", values_for(9.0, len(names)), names,
                         kind="pie", marker=False)
        xml = f"<c:pieChart><c:varyColors val='1'/>{series}<c:firstSliceAng val='0'/></c:pieChart>"
        return xml, first_index + 1
    elif kind == "scatter":
        xml = (
            f"<c:scatterChart><c:scatterStyle val='{group.get('style', 'lineMarker')}'/>"
            f"<c:varyColors val='0'/>{body}{ids}</c:scatterChart>"
        )
    else:
        raise SystemExit(f"unknown group kind {kind!r}")
    return xml, first_index + len(names)


def cat_axis(ax_id: int, cross_id: int, *, delete: bool = False) -> str:
    return (
        f"<c:catAx><c:axId val='{ax_id}'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if delete else 0}'/><c:axPos val='b'/>"
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
    groups = probe["groups"]
    cats = probe.get("cats", CATEGORIES)
    groups_xml = ""
    index = 0
    for group in groups:
        xml, index = group_xml(group, index, cats)
        groups_xml += xml
    kinds = {group["kind"] for group in groups}
    secondary = any(group.get("secondary") for group in groups)

    axes = ""
    if kinds != {"pie"}:
        if kinds == {"scatter"}:
            axes = val_axis(PRIMARY_CAT, PRIMARY_VAL, position="b", gridlines=False)
            axes += val_axis(PRIMARY_VAL, PRIMARY_CAT)
        else:
            axes = cat_axis(PRIMARY_CAT, PRIMARY_VAL) + val_axis(PRIMARY_VAL, PRIMARY_CAT)
        if secondary:
            axes += val_axis(SECOND_VAL, SECOND_CAT, position="r", gridlines=False,
                             crosses="max")
            axes += cat_axis(SECOND_CAT, SECOND_VAL, delete=True)

    size = probe.get("size", 1000)
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        + groups_xml
        + axes
        + "</c:plotArea>"
        + f"<c:legend><c:legendPos val='{probe.get('legend', 'b')}'/>"
        "<c:overlay val='0'/></c:legend>"
        "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
        "</c:chartSpace>"
    ).encode()


# --------------------------------------------------------------------------------------
# The probe tables
# --------------------------------------------------------------------------------------


def _frame(width_pt: float, height_pt: float = 260.0) -> tuple[int, int]:
    return (int(round(width_pt * EMU)), int(round(height_pt * EMU)))


WIDE = _frame(480.0)


def col(names, **extra) -> dict:
    return {"kind": "col", "names": list(names), **extra}


def line(names, **extra) -> dict:
    return {"kind": "line", "names": list(names), **extra}


def probe(key: str, groups: list[dict], **extra) -> dict:
    return {"key": key, "frame": WIDE, "groups": groups, **extra}


#: **The discriminating experiment.**  Four bar entries on one frame; ``x-long1`` and
#: ``x-long4`` lengthen exactly one name.  The pitch between entries 2 and 3 -- neither of
#: which moved -- is what separates a packed layout from a distributed one, and the first
#: key's x says whether the run is centred on its own width or pinned to a band.
FOUR = [SHORT, SHORT + "m", SHORT + "mm", SHORT + "mmm"]

PACK_PROBES: list[dict] = [
    probe("x-base", [col(FOUR)]),
    probe("x-long1", [col([LONG] + FOUR[1:])]),
    probe("x-long4", [col(FOUR[:3] + [LONG])]),
    # The same three with a line key, whose entry width is dominated by the 19.2 pt rule.
    probe("x-lbase", [line(FOUR)]),
    probe("x-llong1", [line([LONG] + FOUR[1:])]),
    probe("x-llong4", [line(FOUR[:3] + [LONG])]),
    # **Entry count** at fixed names.
    *[probe(f"n-{n}", [col(NAMES[:n])]) for n in range(2, 8)],
    # **Name width** at fixed count: three entries whose total text grows, and the case
    # the brief singles out -- one very long entry among short ones.
    probe("w-sss", [col([SHORT, SHORT, SHORT])]),
    probe("w-ssl", [col([SHORT, SHORT, LONG])]),
    probe("w-sls", [col([SHORT, LONG, SHORT])]),
    probe("w-lll", [col([LONG, LONG, LONG])]),
    probe("w-mmm", [col([SHORT + "mmmm"] * 3)]),
]

#: **Key type**, at one frame and one name set, plus the three legend positions.  A line
#: group is known to widen *every* key to 19.2 pt; a marker-only and an area key have
#: never been measured.
THREE = [SHORT, SHORT + "m", SHORT + "mm"]

KEY_PROBES: list[dict] = [
    probe("k-bar", [col(THREE)]),
    probe("k-line", [line(THREE)]),
    probe("k-linenm", [line(THREE, marker=False)]),
    probe("k-area", [{"kind": "area", "names": THREE}]),
    probe("k-pie", [{"kind": "pie", "names": THREE}], cats=THREE),
    probe("k-scatter", [{"kind": "scatter", "names": THREE}]),
    probe("k-scatmark", [{"kind": "scatter", "names": THREE, "style": "marker"}]),
    # The mixed case a combo produces: two bar entries and one line entry.
    probe("k-combo", [col(THREE[:2]), line(THREE[2:], secondary=True)]),
    probe("k-combo2", [col(THREE[:1]), line(THREE[1:], secondary=True)]),
    # **Position.**  ``tr`` may not follow the same rule at all.
    probe("p-b", [col(THREE)], legend="b"),
    probe("p-t", [col(THREE)], legend="t"),
    probe("p-tr", [col(THREE)], legend="tr"),
    probe("p-lb", [line(THREE)], legend="b"),
    probe("p-lt", [line(THREE)], legend="t"),
    probe("p-ltr", [line(THREE)], legend="tr"),
]

#: **Frame width** at a fixed entry set, and **font size** at a fixed frame.  The lesson
#: from the tick sweep: a quantity that looks like a function of the drawn plot is often a
#: function of the frame, so the frame is swept on its own with everything else held.
FRAME_PROBES: list[dict] = [
    *[
        probe(f"f-{width:.0f}", [col(THREE)], frame=_frame(width))
        for width in (240.0, 300.0, 360.0, 420.0, 480.0, 600.0, 720.0)
    ],
    # The same sweep with a line key, whose entries are 19.2 pt wider apiece.
    *[
        probe(f"fl-{width:.0f}", [line(THREE)], frame=_frame(width))
        for width in (240.0, 360.0, 480.0, 720.0)
    ],
    # A frame sweep at **six** entries, where a distributed residue would be six times
    # more sensitive to the frame than a two-entry one.
    *[
        probe(f"f6-{width:.0f}", [col(NAMES[:6])], frame=_frame(width))
        for width in (300.0, 480.0, 720.0)
    ],
    # **Font size.**  The answer is quoted in ems, so a gap that is a constant number of
    # points instead shows up here and nowhere else.
    *[
        probe(f"z-{size // 100}", [col(THREE)], size=size)
        for size in (800, 1000, 1200, 1400, 1800)
    ],
]

def wide(count: int) -> str:
    """A name of a chosen width, always beginning with the same letter."""
    return "W" + "m" * count


#: **The overflow cap, and the slope at the entry counts the first deck only sampled
#: once.**  The first round said the gap is ``0.2 * sum(W) / (n + 1)`` -- it grows with the
#: entries rather than shrinking, so the legend is packed and not distributed -- with one
#: exception, ``w-lll``, whose legend box came out at exactly 0.9 of the frame.  The ``c``
#: family walks content across that threshold at three frame widths and three entry
#: counts, which is what says whether 0.9 is a cap or a coincidence; the ``s`` families
#: sweep the name width at n = 2, 5 and 6, where the first deck had a single point each
#: and so could not see the slope at all.
FIT_PROBES: list[dict] = [
    # n = 3 at 480 pt: the natural box passes 0.9 * frame between k = 11 and k = 12.
    *[probe(f"c3-{k:02d}", [col([wide(k)] * 3)]) for k in (8, 10, 11, 12, 13, 15, 18, 22)],
    # The same threshold at two other frames.  A cap on the frame moves with it.
    *[
        probe(f"c3s-{k:02d}", [col([wide(k)] * 3)], frame=_frame(300.0))
        for k in (4, 6, 8, 12, 18)
    ],
    *[
        probe(f"c3w-{k:02d}", [col([wide(k)] * 3)], frame=_frame(720.0))
        for k in (12, 18, 20, 24)
    ],
    # And at two other entry counts, where the threshold sits at a different name width.
    *[probe(f"c2-{k:02d}", [col([wide(k)] * 2)]) for k in (14, 18, 20, 24)],
    *[probe(f"c6-{k:02d}", [col([wide(k)] * 6)]) for k in (3, 4, 5, 6, 8)],
    # **The slope** at the counts the first deck sampled once.
    *[probe(f"s2-{k:02d}", [col([wide(k)] * 2)]) for k in (1, 4, 8, 12)],
    *[probe(f"s5-{k:02d}", [col([wide(k)] * 5)]) for k in (1, 3, 5)],
    *[probe(f"s6-{k:02d}", [col([wide(k)] * 6)]) for k in (1, 2, 3)],
    # **Validation.**  Fresh combinations of frame, count, key and size, none of which the
    # model was fitted on.
    probe("v-line4", [line([wide(2), wide(5), wide(1), wide(7)])], frame=_frame(360.0)),
    probe("v-bar5", [col([wide(6), wide(1), wide(9), wide(2), wide(4)])],
          frame=_frame(600.0)),
    probe("v-line7", [line([wide(k) for k in (1, 2, 3, 4, 5, 6, 7)])], frame=_frame(720.0)),
    probe("v-small", [col([wide(1), wide(3)])], frame=_frame(200.0)),
    probe("v-size14", [col([wide(2), wide(6), wide(1), wide(4), wide(8)])], size=1400),
    probe("v-top", [col([wide(9), wide(2), wide(5)])], legend="t", frame=_frame(420.0)),
    probe("v-combo", [col([wide(3), wide(6)]), line([wide(2)], secondary=True)],
          frame=_frame(540.0)),
]

DECKS = {
    "legend-pack": PACK_PROBES,
    "legend-key": KEY_PROBES,
    "legend-frame": FRAME_PROBES,
    "legend-fit": FIT_PROBES,
}


# --------------------------------------------------------------------------------------
# Deck assembly -- the same shape as tools/make_combo_probe.py
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

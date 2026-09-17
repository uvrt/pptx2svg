#!/usr/bin/env python3
"""Build decks that measure how an ``ofPieChart`` **packs its two plots**.

``resolve/chart.OF_PIE_BAR_GAP_DIVISOR`` says the bar form's radius is
``r = W / (2 + s + g/200)`` over the polar region's width ``W``, with ``s`` the
``secondPieSize`` fraction and ``g`` the ``gapWidth`` percent.  It was fitted to a
**single** ``gapWidth=100`` reading and says so in its own docstring; ``chart-gallery``
slide 9 is a second reading of the same configuration and ROADMAP.md 3.2a records it as
disagreeing.  These decks sweep both parameters and settle it.

The bar form is the easy one to read: the second plot is an exact **rectangle**, ``s*r``
wide and ``2*s*r`` tall, so one slide yields ``r`` twice with no curve fitting at all.  The
pie form's radius comes off a slice's straight edges, whose endpoints are the only path
points that lie *on* the circle -- an arc's bezier control points sit outside it, which is
why a bounding box is not the diameter here the way it is for a bubble.

Four families, and what each is for:

| family | holds | varies | settles |
| --- | --- | --- | --- |
| ``g`` | frame, ``s`` | ``gapWidth`` 0..300 | the divisor's slope in ``g`` |
| ``s`` | frame, ``g`` | ``secondPieSize`` 25..125 | its slope in ``s`` |
| ``x`` | frame | both at once | that the two are one plane and not two lines |
| ``p`` | frame | the **pie** form over the same sweep | the region ``W``, independently |
| ``r`` | ``g``, ``s`` | the title and legend furniture | where the region's edges are |
| ``c`` | ``g``, ``s`` | frames short enough to **clamp** | what a short region does |

The ``g``, ``s``, ``x`` and ``p`` families sit on a 300 x 300 pt frame with no title and no
legend, which is tall enough that the height clamp cannot bind anywhere in the sweep --
the smallest divisor reached is 2.25, and ``W/2.25`` is 123.6 pt against the 139 pt half
the region's height allows.  The ``c`` family exists precisely to cross that line, because
``chart-gallery`` slide 9 is on the wrong side of it and the clamp has never been measured.

Usage -- the deck's file name picks its probe table::

    python3 tools/make_ofpie_probe.py ~/pptx2svg-oracle/ofpie-pack.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/ofpie-pack.pptx ~/pptx2svg-oracle/ofpie-pack.pdf
    python3 tools/read_ofpie_probe.py ~/pptx2svg-oracle/ofpie-pack.pdf

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

#: Six categories, the last three of which move to the second plot.  The values matter
#: only in that no slice may be so thin that its straight edges fall inside a rounding of
#: each other; these are the gallery's, which are already known to export.
CATEGORIES = ("Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot")
VALUES = (48.0, 22.0, 14.0, 7.0, 5.0, 4.0)
SPLIT_POS = 3


def _cats(names) -> str:
    inner = "".join(f"<c:pt idx='{i}'><c:v>{n}</c:v></c:pt>" for i, n in enumerate(names))
    return (
        "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='{len(names)}'/>{inner}</c:strCache></c:strRef></c:cat>"
    )


def _vals(values) -> str:
    inner = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
    return (
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{inner}</c:numCache></c:numRef></c:val>"
    )


def chart_xml(probe: dict) -> bytes:
    """One ``c:ofPieChart``, in whichever form the probe names."""
    series = (
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Accounts</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        + _cats(CATEGORIES)
        + _vals(VALUES)
        + "</c:ser>"
    )
    groups = (
        f"<c:ofPieChart><c:ofPieType val='{probe.get('form', 'bar')}'/>"
        "<c:varyColors val='1'/>"
        + series
        + f"<c:gapWidth val='{probe.get('gap', 100)}'/>"
        f"<c:splitType val='pos'/><c:splitPos val='{SPLIT_POS}'/>"
        f"<c:secondPieSize val='{probe.get('second', 75)}'/></c:ofPieChart>"
    )
    title = probe.get("title")
    title_xml = (
        "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>"
        f"<a:t>{title}</a:t></a:r></a:p></c:rich></c:tx>"
        "<c:overlay val='0'/></c:title><c:autoTitleDeleted val='0'/>"
        if title
        else "<c:autoTitleDeleted val='1'/>"
    )
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
        "<c:chart>"
        + title_xml
        + "<c:plotArea><c:layout/>"
        + groups
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


def _frame(width_pt: float, height_pt: float) -> tuple[int, int]:
    return (int(round(width_pt * EMU)), int(round(height_pt * EMU)))


#: Tall enough that the height clamp cannot bind anywhere in the sweep -- see the module
#: docstring.  Bare of a title and a legend so the region is the frame less one edge inset
#: on each side and nothing else.
SQUARE = _frame(300.0, 300.0)


def probe(key: str, **extra) -> dict:
    return {"key": key, "frame": SQUARE, **extra}


PACK_PROBES: list[dict] = [
    # **gapWidth**, at the default second size.  0 and 300 are the schema's ends.
    *[probe(f"g-{gap:03d}", gap=gap) for gap in (0, 25, 50, 100, 150, 200, 300)],
    # **secondPieSize**, at the default gap.
    *[probe(f"s-{size:03d}", second=size) for size in (25, 50, 75, 100, 125)],
    # **Both at once.**  A divisor linear in each separately still has to be linear in the
    # pair, and these are the slides that say so.
    probe("x-0-25", gap=0, second=25),
    probe("x-300-100", gap=300, second=100),
    probe("x-50-125", gap=50, second=125),
    probe("x-200-050", gap=200, second=50),
    probe("x-150-035", gap=150, second=35),
    # **The pie form** over the same sweep, whose divisor is already measured to the last
    # decimal.  Its job here is to fix the region's width independently, so that a residual
    # in the bar form cannot hide in a mis-read ``W``.
    *[probe(f"p-{gap:03d}", form="pie", gap=gap) for gap in (0, 100, 300)],
    *[probe(f"ps-{size:03d}", form="pie", second=size) for size in (25, 50, 100)],
]

#: **Where the region's edges are.**  ``chart-gallery`` slide 9 carries a title and a
#: bottom legend and its drawn radius is 1.33 pt under what this library computes, which is
#: either a divisor error or a region error and cannot be told apart on that slide alone.
#: Every slide here holds ``gapWidth`` and ``secondPieSize`` at the gallery's values and
#: moves only the furniture, in both forms -- so the pie form, whose divisor is not in
#: question, reads the region out directly and the bar form is then over-determined.
REGION_PROBES: list[dict] = [
    *[
        probe(f"r-{name}", form=form, **extra)
        for form in ("bar", "pie")
        for name, extra in (
            ("bare", {}),
            ("title", {"title": "Accounts by tier"}),
            ("legb", {"legend": "b"}),
            ("legt", {"legend": "t"}),
            ("legr", {"legend": "r"}),
            ("both", {"title": "Accounts by tier", "legend": "b"}),
        )
    ],
]

#: **The clamp.**  ``_of_pie_geometry`` caps the radius at half the region's height, and
#: at ``region.height / (2 * s)`` for the bar, and says in its own docstring that neither
#: is measured -- no probe frame was ever short enough to reach them.  These frames are:
#: every one is wide enough that the width law would ask for a radius taller than the
#: region, so the drawn radius is the cap and the drawn *position* says what the layout
#: does with the width it did not use.  ``c-gallery`` is slide 9's own frame and furniture.
CLAMP_PROBES: list[dict] = [
    *[
        {"key": f"c-{width:.0f}x{height:.0f}", "frame": _frame(width, height), "gap": gap,
         "second": second, "form": form}
        for form in ("bar", "pie")
        for width, height, gap, second in (
            (520.0, 200.0, 100, 75),
            (600.0, 180.0, 100, 75),
            (480.0, 160.0, 0, 75),
            (520.0, 220.0, 100, 25),
            (520.0, 220.0, 300, 125),
        )
    ],
    {"key": "c-gallery", "frame": _frame(520.0, 336.0), "gap": 100, "second": 75,
     "form": "bar", "title": "Accounts by tier", "legend": "b"},
    {"key": "c-gallery-pie", "frame": _frame(520.0, 336.0), "gap": 100, "second": 75,
     "form": "pie", "title": "Accounts by tier", "legend": "b"},
]

DECKS = {
    "ofpie-pack": PACK_PROBES,
    "ofpie-region": REGION_PROBES,
    "ofpie-clamp": CLAMP_PROBES,
}


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

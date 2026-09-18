#!/usr/bin/env python3
"""Build decks of **3-D** charts, to measure what ``c:view3D`` does to the axis.

ROADMAP.md 3.4 records that PowerPoint rasterises 3-D chart geometry but leaves the text
vector, so the axis, its tick labels and therefore the plot rectangle are exactly
measurable even though the scene itself is a picture.  These decks are that measurement.

The defect these exist for is gallery slide 16, whose ``area3DChart`` PowerPoint draws
**0..50 by 5** where the flat fallback draws 0..60 by 10.  The stated cause was a depth
reservation shrinking the plot, which would lower the interval count through
:func:`~pptx2svg.resolve.chart.side_axis_intervals` -- so each probe reads back two
things at once:

* the **axis** PowerPoint chose, from the tick labels' values, and
* the **plot rectangle**, from the extreme tick labels' own centres, which is where the
  depth reservation would show.

Both decks draw one chart per slide, no title and no legend, so the frame the axis
divides is the whole frame and nothing else has to be subtracted from it.

``view3d-meter``
    The N-meter (see ``tools/make_axis_probe.py``) on a ``bar3DChart`` over seven frame
    heights, plus ``F`` (0..50) and ``G`` (0..96), which are the two datasets that
    separate "the range is padded 5% at each end" from "it is not".  A ``col`` control
    and a ``line3DChart``/``area3DChart``/``pie3DChart`` row say whether the answer is a
    property of 3-D or of one group element.
``view3d-view``
    One dataset, one frame, and ``c:view3D`` swept one element at a time --
    ``depthPercent``, ``hPercent``, ``rotX``, ``rotY``, ``rAngAx``, ``perspective``, and
    the element absent altogether.  The axis is held at the interval cap so that anything
    that moves is geometry rather than the count; a short-frame block at the end puts the
    count back in play at the extremes of the depth.

Usage -- the deck's file name picks its probe table::

    python3 tools/make_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/view3d-meter.pptx ~/pptx2svg-oracle/view3d-meter.pdf
    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pdf

The decks and their exports are throwaway and are **not** committed; ``~/pptx2svg-oracle``
is the only directory PowerPoint is allowed to write to, and it is left holding its
eighteen corpus files.  See ROADMAP.md section 0.1 before blaming a failed export.
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from make_axis_probe import A, C, R, write_deck  # noqa: E402

#: The five N-meter datasets, whose drawn units together name the interval count, plus
#: the two that separate the padded rule from the bare one.  Under the shipped 2-D rule
#: ``F`` draws 0..60 by 10 and ``G`` 0..120 by 20 at ten intervals; with no padding at all
#: they draw 0..50 by 5 and 0..100 by 10, which is what gallery slide 16 shows.
METER = (("A", 5.0), ("B", 9.0), ("C", 3.8), ("D", 6.6), ("E", 8.5))
PADDING = (("F", 50.0), ("G", 96.0))

#: ``c:view3D`` as PowerPoint writes it for a chart inserted from the 3-D gallery.
DEFAULT_VIEW = {"rotX": 15, "rotY": 20, "depthPercent": 100, "rAngAx": 1}

FRAME_WIDTH = 8686800  # 684 pt, the width every axis deck has used
HEIGHTS = (60, 90, 120, 150, 195, 250, 330)

METER_PROBES: list[dict] = [
    *[
        {
            "key": f"m{height}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, height * 12700),
            "kind": "bar3D",
            "view": DEFAULT_VIEW,
        }
        for height in HEIGHTS
        for tag, high in METER + PADDING
    ],
    # The same two padding probes drawn **flat**, on the same deck and the same frames:
    # the control that says the export machine still draws the 2-D rule this is compared
    # against, rather than that the rule has moved under everyone's feet.
    *[
        {
            "key": f"flat{height}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, height * 12700),
            "kind": "col",
            "view": None,
        }
        for height in (120, 195)
        for tag, high in PADDING
    ],
    # Is the answer a property of "3-D" or of ``bar3DChart``?  Same data, same frame.
    *[
        {
            "key": f"type-{kind}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, 195 * 12700),
            "kind": kind,
            "view": DEFAULT_VIEW,
        }
        for kind in ("line3D", "area3D", "area3Dstack")
        for tag, high in PADDING
    ],
]

#: The view sweep.  One dataset (``B``, 0..9) on a frame tall enough that the interval
#: count is at its cap, so the drawn axis is the same on every slide and every difference
#: read back is geometry.
VIEW_FRAME = (FRAME_WIDTH, 195 * 12700)


def _view(**overrides) -> dict:
    view = dict(DEFAULT_VIEW)
    view.update(overrides)
    return {key: value for key, value in view.items() if value is not None}


VIEW_CASES: list[tuple[str, dict | None]] = [
    ("absent", None),
    ("default", DEFAULT_VIEW),
    ("empty", {}),
    *[(f"d{value}", _view(depthPercent=value)) for value in (20, 50, 200, 500, 1000, 2000)],
    *[(f"h{value}", _view(hPercent=value)) for value in (20, 50, 100, 200, 500)],
    *[(f"rx{value}", _view(rotX=value)) for value in (0, 5, 30, 45, 60, 90, -15, -45)],
    *[(f"ry{value}", _view(rotY=value)) for value in (0, 45, 90, 135, 180, 270, 340)],
    ("ra0", _view(rAngAx=0)),
    ("ra0-rx30", _view(rAngAx=0, rotX=30)),
    *[(f"ra0-p{value}", _view(rAngAx=0, perspective=value)) for value in (0, 30, 60, 120)],
    # `rAngAx=1` with a perspective stated: the schema says the perspective is ignored
    # when the axes are at right angles, and this is the probe that says whether it is.
    ("ra1-p120", _view(perspective=120)),
]

VIEW_PROBES: list[dict] = [
    *[
        {
            "key": f"v-{name}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": "bar3D",
            "view": view,
        }
        for name, view in VIEW_CASES
    ],
    # The count back in play: a 120 pt frame takes six intervals flat, so if the depth
    # reserve comes off the band the count is read from, these differ from each other.
    *[
        {
            "key": f"n{name}-{tag}",
            "high": high,
            "frame": (FRAME_WIDTH, 120 * 12700),
            "kind": "bar3D",
            "view": view,
        }
        for name, view in (
            ("d20", _view(depthPercent=20)),
            ("d100", DEFAULT_VIEW),
            ("d2000", _view(depthPercent=2000)),
            ("rx0", _view(rotX=0)),
            ("rx60", _view(rotX=60)),
            ("h500", _view(hPercent=500)),
        )
        for tag, high in (("A", 5.0), ("B", 9.0), ("F", 50.0))
    ],
]

#: The group element as a variable.  ``view3d-view`` swept ``c:view3D`` on a
#: ``bar3DChart`` alone, and the three readings beside it on ``view3d-meter`` say the
#: other 3-D group elements do *not* share its scene: on one frame and one view a
#: ``bar3DChart`` drew a 127.68 pt axis where ``line3DChart`` drew 88.56 and
#: ``area3DChart`` 59.04.  This repeats the sweep per element, and adds the series count,
#: which is the depth's own row divisor and the obvious candidate for the difference.
TYPE_CASES: list[tuple[str, dict | None, dict]] = [
    ("base", DEFAULT_VIEW, {}),
    ("d20", _view(depthPercent=20), {}),
    ("d500", _view(depthPercent=500), {}),
    ("rx0", _view(rotX=0), {}),
    ("rx60", _view(rotX=60), {}),
    ("ry0", _view(rotY=0), {}),
    ("ry90", _view(rotY=90), {}),
    ("h50", _view(hPercent=50), {}),
    ("h200", _view(hPercent=200), {}),
    ("s2", DEFAULT_VIEW, {"series": 2}),
    ("s3", DEFAULT_VIEW, {"series": 3}),
    ("f120", DEFAULT_VIEW, {"frame": (FRAME_WIDTH, 120 * 12700)}),
    ("f330", DEFAULT_VIEW, {"frame": (FRAME_WIDTH, 330 * 12700)}),
]

TYPE_PROBES: list[dict] = [
    {
        "key": f"t-{kind}-{name}",
        "high": 9.0,
        "frame": VIEW_FRAME,
        "kind": kind,
        "view": view,
        **extra,
    }
    for kind in ("bar3D", "line3D", "area3D")
    for name, view, extra in TYPE_CASES
]

#: The **series count** as a variable, which ``view3d-type`` found and could not settle.
#: A ``bar3DChart``'s depth reservation *shrinks* as series are added -- one, two and
#: three series reserved 0.0600, 0.0459 and 0.0373 of the scene's width on one frame --
#: and `2.5 / (1.5 + n)` and `1 / sqrt(n)` both reproduce those three to 3%.  They part
#: company at six (0.333 against 0.408), which is what this deck is for.  ``gapDepth`` is
#: swept beside it because the first of those two laws is written in terms of it.
SERIES_PROBES: list[dict] = [
    *[
        {
            "key": f"n{count}-f{height}",
            "high": 9.0,
            "frame": (FRAME_WIDTH, height * 12700),
            "kind": "bar3D",
            "view": DEFAULT_VIEW,
            "series": count,
        }
        for height in (195, 120)
        for count in (1, 2, 3, 4, 5, 6)
    ],
    *[
        {
            "key": f"g{gap}-s{count}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": "bar3D",
            "view": DEFAULT_VIEW,
            "series": count,
            "gapDepth": gap,
        }
        for count in (1, 2, 4)
        for gap in (0, 50, 300, 500)
    ],
    *[
        {
            "key": f"{kind}-n{count}",
            "high": 9.0,
            "frame": VIEW_FRAME,
            "kind": kind,
            "view": DEFAULT_VIEW,
            "series": count,
        }
        for kind in ("line3D", "area3D")
        for count in (1, 2, 3, 4)
    ],
]

DECKS = {
    "view3d-meter": METER_PROBES,
    "view3d-view": VIEW_PROBES,
    "view3d-type": TYPE_PROBES,
    "view3d-series": SERIES_PROBES,
}

CATEGORIES = ("C1", "C2", "C3", "C4", "C5")
#: The same shape every axis deck has used, so a value is never the maximum twice.
FACTORS = (0.3, 1.0, 0.55, 0.8, 0.45)

VIEW_ORDER = ("rotX", "hPercent", "rotY", "depthPercent", "rAngAx", "perspective")


def view_xml(view: dict | None) -> str:
    """``c:view3D``, in schema order.  ``None`` writes no element at all."""
    if view is None:
        return ""
    body = "".join(
        f"<c:{name} val='{view[name]}'/>" for name in VIEW_ORDER if name in view
    )
    return f"<c:view3D>{body}</c:view3D>"


def series_xml(high: float, index: int = 0, scale: float = 1.0) -> str:
    values = [high * factor * scale for factor in FACTORS]
    values[1] = high * scale
    points = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
    cats = "".join(
        f"<c:pt idx='{i}'><c:v>{name}</c:v></c:pt>" for i, name in enumerate(CATEGORIES)
    )
    return (
        f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        f"<c:pt idx='0'><c:v>S{index + 1}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='{len(CATEGORIES)}'/>{cats}</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef></c:val>"
        "</c:ser>"
    )


def axes_xml(kind: str) -> str:
    """The category, value and (for a 3-D group) series axes this chart needs.

    The value axis carries gridlines and ten-point labels exactly as every earlier axis
    deck's does, so the reading is comparable to them.  The ``c:serAx`` is deleted --
    PowerPoint still draws the depth, it just prints nothing along it.
    """
    if kind == "pie3D":
        return ""
    depth = (
        "<c:serAx><c:axId val='100004'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='1'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/>"
        "<c:crossAx val='100003'/><c:crosses val='autoZero'/></c:serAx>"
        if "3D" in kind
        else ""
    )
    cross_between = "midCat" if kind.startswith("area") else "between"
    return (
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
        f"<c:crosses val='autoZero'/><c:crossBetween val='{cross_between}'/></c:valAx>"
        + depth
    )


def group_xml(kind: str, high: float, series: int = 1, gap_depth: int = 150) -> str:
    """One ``c:*Chart`` group.  A 3-D group states three ``c:axId`` children, exactly.

    *series* is the number of ``c:ser`` children, which is what a 3-D chart lays out
    along its **depth** -- one row per series -- and therefore the lever that separates
    "the scene's depth is ``depthPercent`` of its width" from "it is that per row".
    """
    body = "".join(
        series_xml(high, index, 1.0 if series == 1 else 0.6 - 0.15 * index)
        for index in range(series)
    )
    depth_gap = f"<c:gapDepth val='{gap_depth}'/>"
    ids3 = "<c:axId val='100002'/><c:axId val='100003'/><c:axId val='100004'/>"
    ids2 = "<c:axId val='100002'/><c:axId val='100003'/>"
    if kind == "bar3D":
        return (
            "<c:bar3DChart><c:barDir val='col'/><c:grouping val='clustered'/>"
            "<c:varyColors val='0'/>" + body + "<c:gapWidth val='150'/>"
            + depth_gap + "<c:shape val='box'/>" + ids3 + "</c:bar3DChart>"
        )
    if kind == "col":
        return (
            "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
            "<c:varyColors val='0'/>" + body + "<c:gapWidth val='150'/>"
            + ids2 + "</c:barChart>"
        )
    if kind == "line3D":
        return (
            "<c:line3DChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            + body
            + depth_gap + ids3 + "</c:line3DChart>"
        )
    if kind == "area3D":
        return (
            "<c:area3DChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            + body
            + depth_gap + ids3 + "</c:area3DChart>"
        )
    if kind == "area3Dstack":
        # Gallery slide 16 is stacked and two-series, which is the one shape the single
        # series above cannot rule out on its own.
        return (
            "<c:area3DChart><c:grouping val='stacked'/><c:varyColors val='0'/>"
            + series_xml(high, 0, 0.6)
            + series_xml(high, 1, 0.4)
            + "<c:gapDepth val='150'/>" + ids3 + "</c:area3DChart>"
        )
    raise SystemExit(f"unknown probe kind {kind!r}")


def chart_part(probe: dict) -> bytes:
    size = probe.get("size", 1000)
    kind = probe["kind"]
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:date1904 val='0'/><c:lang val='en-US'/><c:roundedCorners val='0'/>"
        "<c:chart>"
        + view_xml(probe["view"])
        + "<c:autoTitleDeleted val='1'/>"
        "<c:plotArea><c:layout/>"
        + group_xml(
            kind, probe["high"], probe.get("series", 1), probe.get("gapDepth", 150)
        )
        + axes_xml(kind)
        + "</c:plotArea>"
        "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{size}'/></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
        "</c:chartSpace>"
    ).encode()


def probes_for(target: Path) -> list[dict]:
    for name, table in DECKS.items():
        if name in target.stem:
            return table
    raise SystemExit(f"name the deck after one of {sorted(DECKS)}, not {target.stem!r}")


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else Path("view3d-meter.pptx")
    probes = probes_for(target)
    write_deck(target, probes, part_for=chart_part)
    for index, probe in enumerate(probes):
        view = probe["view"]
        shown = "absent" if view is None else ",".join(f"{k}={v}" for k, v in view.items())
        print(
            f"{index + 1:3d} {probe['key']:16s} {probe['kind']:12s} 0..{probe['high']:<6g} "
            f"frame={probe['frame'][1] / 12700:.0f}pt  {shown}"
        )
    print(f"\n{len(probes)} probes -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

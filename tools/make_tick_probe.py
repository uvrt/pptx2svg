#!/usr/bin/env python3
"""Build a deck of charts that differ in their axis **tick marks** and nothing else.

The question this exists to settle is what PowerPoint draws for ``c:majorTickMark`` and
``c:minorTickMark``: how long a tick is, which side of the axis it stands on for
``out``, ``in`` and ``cross``, where along the axis the ticks go on a category axis and
on a value axis, what stroke they take, and what an absent element means.

Every probe is one chart in the same 400 x 250 pt frame with **no gridlines**, so every
stroke on the page is an axis line or a tick and the reader can take them all.  The bars
and markers are filled, not stroked, which keeps them out of the stroke list.

The deck's file name picks its probe table, and each table's comment says what the one
before it raised: ``tick-marks`` (every setting, every 2-D kind), ``tick-length`` (the
length against the label face and size, the far side, secondary and radar axes),
``tick-check``, ``tick-radar``, ``tick-default`` (the missing element),
``tick-verbatim`` and ``tick-workbook`` (a real 2007 chart part), and ``tick-app-12``,
``-12mac``, ``-14``, ``-14mac``, ``-15``, ``-16`` (the same questions under each
``docProps/app.xml`` producer).  ``resolve/chart.py`` and ROADMAP.md, *Axis tick marks,
measured*, have what they found.

Usage::

    python3 tools/make_tick_probe.py ~/zz-tick-marks.pptx
    osascript tools/powerpoint_export_pdf.applescript ~/zz-tick-marks.pptx \\
        ~/zz-tick-marks.pdf
    python3 tools/read_tick_probe.py ~/zz-tick-marks.pdf

The deck and its export are throwaway and are **not** committed, and they do **not** go
in ``~/pptx2svg-oracle``: the fidelity harness scores every deck in that directory.
PowerPoint has to be able to write where they go (ROADMAP.md 0.1); the home directory
itself is one it can.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from make_axis_probe import A, C, R, write_deck  # noqa: E402

PT = 12700
FRAME = (int(400 * PT), int(250 * PT))

#: ``None`` leaves the element out; ``""`` writes it with no ``val``.
ABSENT = None
NO_VAL = ""


def _tick(name: str, value: str | None) -> str:
    if value is None:
        return ""
    if value == "":
        return f"<c:{name}/>"
    return f"<c:{name} val='{value}'/>"


def _line(line: dict | None) -> str:
    """An axis ``c:spPr``: ``{"w": emu, "color": "RRGGBB", "dash": "dash"}`` or noFill."""
    if line is None:
        return ""
    if line.get("none"):
        return "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>"
    width = f" w='{line['w']}'" if "w" in line else ""
    fill = f"<a:solidFill><a:srgbClr val='{line['color']}'/></a:solidFill>" if "color" in line else ""
    dash = f"<a:prstDash val='{line['dash']}'/>" if "dash" in line else ""
    cap = f" cap='{line['cap']}'" if "cap" in line else ""
    return f"<c:spPr><a:ln{width}{cap}>{fill}{dash}</a:ln></c:spPr>"


def _text(text: dict | None) -> str:
    """An axis' own ``c:txPr``: ``{"size": 1400, "face": "Arial"}``."""
    if text is None:
        return ""
    size = f" sz='{text['size']}'" if "size" in text else ""
    face = f"<a:latin typeface='{text['face']}'/>" if "face" in text else ""
    return (
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr{size}>{face}</a:defRPr></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
    )


def _axis(
    tag: str,
    axis_id: str,
    cross_id: str,
    position: str,
    options: dict,
    *,
    category: bool,
    cross_between: str | None = None,
) -> str:
    major = options.get("major", "out")
    minor = options.get("minor", "none")
    scaling = "<c:scaling><c:orientation val='minMax'/>"
    if "max" in options:
        scaling += f"<c:max val='{options['max']}'/>"
    if "min" in options:
        scaling += f"<c:min val='{options['min']}'/>"
    scaling += "</c:scaling>"
    body = (
        f"<c:{tag}><c:axId val='{axis_id}'/>{scaling}"
        + ("" if options.get("no_delete") else f"<c:delete val='{1 if options.get('delete') else 0}'/>")
        + f"<c:axPos val='{position}'/>"
        + ("<c:majorGridlines/>" if options.get("grid") else "")
        + "<c:numFmt formatCode='General' sourceLinked='1'/>"
        + _tick("majorTickMark", major)
        + _tick("minorTickMark", minor)
        + f"<c:tickLblPos val='{options.get('labels', 'nextTo')}'/>"
        + _line(options.get("line"))
        + _text(options.get("text"))
        + f"<c:crossAx val='{cross_id}'/>"
        + (
            f"<c:crossesAt val='{options['crosses_at']}'/>"
            if "crosses_at" in options
            else f"<c:crosses val='{options.get('crosses', 'autoZero')}'/>"
        )
    )
    if category:
        body += "<c:auto val='1'/><c:lblAlgn val='ctr'/><c:lblOffset val='100'/>"
        if "skip" in options:
            body += f"<c:tickMarkSkip val='{options['skip']}'/>"
        body += "<c:noMultiLvlLbl val='0'/>"
    else:
        body += f"<c:crossBetween val='{cross_between or 'between'}'/>"
        if "major_unit" in options:
            body += f"<c:majorUnit val='{options['major_unit']}'/>"
        if "minor_unit" in options:
            body += f"<c:minorUnit val='{options['minor_unit']}'/>"
    return body + f"</c:{tag}>"


def _cache(values) -> str:
    points = "".join(f"<c:pt idx='{i}'><c:v>{v!r}</c:v></c:pt>" for i, v in enumerate(values))
    return (
        "<c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef>"
    )


def chart_part(probe: dict) -> bytes:
    kind = probe["kind"]
    values = probe.get("values", (3.0, 9.0, 5.0, 7.0, 4.0))
    names = [f"C{i + 1}" for i in range(len(values))]
    cats = "".join(f"<c:pt idx='{i}'><c:v>{name}</c:v></c:pt>" for i, name in enumerate(names))
    head = (
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Probe</c:v></c:pt></c:strCache></c:strRef></c:tx>"
    )
    # Series drawn filled and never stroked, so the page's strokes are the axes' own.
    no_line = "<c:spPr><a:solidFill><a:srgbClr val='4472C4'/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>"
    marker = (
        "<c:marker><c:symbol val='square'/><c:size val='5'/>"
        "<c:spPr><a:solidFill><a:srgbClr val='4472C4'/></a:solidFill>"
        "<a:ln><a:noFill/></a:ln></c:spPr></c:marker>"
    )
    category_series = (
        head
        + (no_line if kind in ("col", "bar", "area") else "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>")
        + (marker if kind in ("line", "radar") else "")
        + "<c:cat><c:strRef><c:strCache>"
        f"<c:ptCount val='{len(names)}'/>{cats}</c:strCache></c:strRef></c:cat>"
        f"<c:val>{_cache(values)}</c:val>"
        + ("<c:smooth val='0'/>" if kind == "line" else "")
        + "</c:ser>"
    )
    cat = probe.get("cat", {})
    val = probe.get("val", {})
    ids = "<c:axId val='100002'/><c:axId val='100003'/>"
    if kind == "combo":
        # Columns on the primary axes and a line on a secondary pair: the secondary
        # category axis deleted, the secondary value axis up the right edge.
        line_series = (
            category_series.replace("<c:idx val='0'/><c:order val='0'/>", "<c:idx val='1'/><c:order val='1'/>")
            .replace(no_line, "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>" + marker)
            .replace("</c:val>", "</c:val><c:smooth val='0'/>")
        )
        plot = (
            "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
            f"<c:varyColors val='0'/>{category_series}<c:gapWidth val='150'/>"
            f"{ids}</c:barChart>"
            "<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            f"{line_series}<c:marker val='1'/>"
            "<c:axId val='100004'/><c:axId val='100005'/></c:lineChart>"
        )
    elif kind in ("col", "bar"):
        plot = (
            f"<c:barChart><c:barDir val='{kind}'/><c:grouping val='clustered'/>"
            f"<c:varyColors val='0'/>{category_series}<c:gapWidth val='150'/>"
            f"{ids}</c:barChart>"
        )
    elif kind == "line":
        plot = (
            "<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            f"{category_series}<c:marker val='1'/>{ids}</c:lineChart>"
        )
    elif kind == "area":
        plot = (
            "<c:areaChart><c:grouping val='standard'/><c:varyColors val='0'/>"
            f"{category_series}{ids}</c:areaChart>"
        )
    elif kind == "radar":
        plot = (
            "<c:radarChart><c:radarStyle val='marker'/><c:varyColors val='0'/>"
            f"{category_series}{ids}</c:radarChart>"
        )
    elif kind == "scatter":
        xs = probe.get("xs", (1.0, 2.0, 3.0, 4.0, 5.0))
        plot = (
            "<c:scatterChart><c:scatterStyle val='lineMarker'/><c:varyColors val='0'/>"
            + head
            + "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>"
            + marker
            + f"<c:xVal>{_cache(xs)}</c:xVal><c:yVal>{_cache(values)}</c:yVal>"
            "<c:smooth val='0'/></c:ser>" + ids + "</c:scatterChart>"
        )
    else:
        raise SystemExit(f"unknown chart kind {kind!r}")

    if kind == "scatter":
        axes = _axis("valAx", "100002", "100003", "b", cat, category=False, cross_between="midCat")
        axes += _axis("valAx", "100003", "100002", "l", val, category=False, cross_between="midCat")
    else:
        cat_pos, val_pos = ("l", "b") if kind == "bar" else ("b", "l")
        axes = _axis("catAx", "100002", "100003", cat_pos, cat, category=True)
        axes += _axis(
            "valAx",
            "100003",
            "100002",
            val_pos,
            val,
            category=False,
            cross_between=probe.get("cross_between", "midCat" if kind == "area" else "between"),
        )
    if kind == "combo":
        axes += _axis("catAx", "100004", "100005", "b", {"delete": True, **both("out")}, category=True)
        axes += _axis(
            "valAx",
            "100005",
            "100004",
            "r",
            {"crosses": "max", **probe.get("second", {})},
            category=False,
        )
    size = probe.get("size")
    face = probe.get("face")
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        + plot
        + axes
        + "</c:plotArea><c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
        + (
            "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
            f"<a:defRPr sz='{size or 1000}'>"
            + (f"<a:latin typeface='{face}'/>" if face else "")
            + "</a:defRPr></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
            if size or face
            else ""
        )
        + "</c:chartSpace>"
    ).encode()


def verbatim_part(probe: dict) -> bytes:
    """A chart part read from another deck, edited by the probe's ``edits``.

    ``source`` is ``(deck, part name)``.  The part's ``c:externalData`` and
    ``c:userShapes`` are dropped, since their relationships are not carried over; each
    edit is a ``(pattern, replacement)`` regular-expression substitution.
    """
    import re
    import zipfile

    deck, name = probe["source"]
    xml = zipfile.ZipFile(Path(deck).expanduser()).read(name).decode()
    xml = re.sub(r"<c:userShapes[^>]*/>", "", xml)
    if not probe.get("workbook"):
        xml = re.sub(r"<c:externalData[^>]*/>", "", xml)
    for pattern, replacement in probe.get("edits", ()):
        xml = re.sub(pattern, replacement, xml)
    return xml.encode()


def any_part(probe: dict) -> bytes:
    if "source" in probe:
        return verbatim_part(probe)
    xml = chart_part(probe)
    if probe.get("workbook"):
        xml = xml.replace(
            b"</c:chartSpace>",
            b"<c:externalData r:id='rId1'><c:autoUpdate val='0'/></c:externalData></c:chartSpace>",
        )
    return xml


XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PACKAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"


def embed_workbooks(target: Path, probes: list[dict], workbook: bytes) -> None:
    """Give every probe marked ``workbook`` an embedded workbook, as ``rId1``.

    Office writes one beside every chart it saves; a hand-written chart part has none.
    The workbook is only ever *named* -- nothing here asks PowerPoint to open it -- so any
    valid one does.
    """
    import zipfile

    if not any(item.get("workbook") for item in probes):
        return
    with zipfile.ZipFile(target) as source:
        entries = [(info, source.read(info.filename)) for info in source.infolist()]
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            if info.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    f'<Default Extension="xlsx" ContentType="{XLSX_TYPE}"/></Types>'.encode(),
                )
            out.writestr(info.filename, data)
        for index, item in enumerate(probes):
            if not item.get("workbook"):
                continue
            out.writestr(f"ppt/embeddings/probe{index}.xlsx", workbook)
            out.writestr(
                f"ppt/charts/_rels/probe{index}.xml.rels",
                "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
                "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
                f"<Relationship Id='rId1' Type='{PACKAGE_REL}' "
                f"Target='../embeddings/probe{index}.xlsx'/></Relationships>",
            )


def probe(key: str, kind: str = "col", **extra) -> dict:
    row = {"key": key, "kind": kind, "frame": extra.pop("frame", FRAME)}
    row.update(extra)
    return row


def both(major, minor="none") -> dict:
    return {"major": major, "minor": minor}


RED3 = {"w": 38100, "color": "FF0000"}
BLUE2 = {"w": 25400, "color": "0000FF"}

#: The first deck: every setting on a column chart, then the other chart kinds.
MARKS_PROBES: list[dict] = [
    probe("col-absent", cat=both(ABSENT, ABSENT), val=both(ABSENT, ABSENT)),
    probe("col-noval", cat=both(NO_VAL, NO_VAL), val=both(NO_VAL, NO_VAL)),
    probe("col-out", cat=both("out"), val=both("out")),
    probe("col-in", cat=both("in"), val=both("in")),
    probe("col-cross", cat=both("cross"), val=both("cross")),
    probe("col-none", cat=both("none"), val=both("none")),
    probe("col-min-out", cat=both("none", "out"), val=both("none", "out")),
    probe("col-min-in", cat=both("none", "in"), val=both("none", "in")),
    probe("col-min-cross", cat=both("none", "cross"), val=both("none", "cross")),
    probe("col-both-out", cat=both("out", "out"), val=both("out", "out")),
    probe("col-both-x", cat=both("cross", "in"), val=both("in", "cross")),
    probe(
        "col-color",
        cat={**both("out", "out"), "line": BLUE2},
        val={**both("out", "out"), "line": RED3},
    ),
    probe(
        "col-color-x",
        cat={**both("cross", "in"), "line": BLUE2},
        val={**both("cross", "in"), "line": RED3},
    ),
    probe(
        "col-dash",
        cat={**both("out"), "line": {"w": 19050, "color": "00AA00", "dash": "dash"}},
        val={**both("out"), "line": {"w": 19050, "color": "00AA00", "dash": "sysDot"}},
    ),
    probe(
        "col-noline",
        cat={**both("out", "out"), "line": {"none": True}},
        val={**both("out", "out"), "line": {"none": True}},
    ),
    probe("col-neg", values=(3.0, -4.0, 5.0, -2.0, 6.0), cat=both("out", "out"), val=both("out", "out")),
    probe(
        "col-neg-cross",
        values=(3.0, -4.0, 5.0, -2.0, 6.0),
        cat=both("cross", "in"),
        val=both("cross"),
    ),
    probe(
        "col-neg-low",
        values=(3.0, -4.0, 5.0, -2.0, 6.0),
        cat={**both("out"), "labels": "low"},
        val=both("out"),
    ),
    probe("col-skip2", cat={**both("out"), "skip": 2}, val=both("out")),
    probe("col-lblnone", cat={**both("out"), "labels": "none"}, val={**both("out"), "labels": "none"}),
    probe("col-lblhigh", cat={**both("out"), "labels": "high"}, val={**both("out"), "labels": "high"}),
    probe("col-delval", cat=both("out"), val={**both("out"), "delete": True}),
    probe("col-units", cat=both("out"), val={**both("out", "out"), "major_unit": 2, "minor_unit": 0.5}),
    probe("col-units-default", cat=both("out"), val={**both("out", "out"), "major_unit": 3}),
    probe("col-size20", size=2000, cat=both("out", "out"), val=both("out", "out")),
    probe(
        "col-small", frame=(int(160 * PT), int(110 * PT)), cat=both("out", "out"), val=both("out", "out")
    ),
    probe("col-many", values=tuple(float(1 + i % 7) for i in range(14)), cat=both("out"), val=both("out")),
    probe("bar-out", "bar", cat=both("out", "out"), val=both("out", "out")),
    probe("bar-cross", "bar", cat=both("cross", "in"), val=both("cross", "in")),
    probe("bar-neg", "bar", values=(3.0, -4.0, 5.0, -2.0, 6.0), cat=both("out"), val=both("out")),
    probe("line-out", "line", cat=both("out", "out"), val=both("out", "out")),
    probe("line-midcat", "line", cross_between="midCat", cat=both("out", "out"), val=both("out")),
    probe("line-cross", "line", cat=both("cross"), val=both("cross")),
    probe("area-out", "area", cat=both("out", "out"), val=both("out")),
    probe("scat-out", "scatter", cat=both("out", "out"), val=both("out", "out")),
    probe("scat-cross", "scatter", cat=both("cross", "in"), val=both("in", "cross")),
    probe(
        "scat-neg",
        "scatter",
        xs=(-3.0, -1.0, 1.0, 2.0, 4.0),
        values=(3.0, -4.0, 5.0, -2.0, 6.0),
        cat=both("out", "out"),
        val=both("out", "out"),
    ),
    probe("scat-absent", "scatter", cat=both(ABSENT, ABSENT), val=both(ABSENT, ABSENT)),
    probe("radar-out", "radar", cat=both("out", "out"), val=both("out", "out")),
    probe("radar-cross", "radar", cat=both("cross"), val=both("cross", "in")),
    probe("radar-none", "radar", cat=both("none"), val=both("none")),
    probe("radar-absent", "radar", cat=both(ABSENT, ABSENT), val=both(ABSENT, ABSENT)),
]

#: The second deck, on what the first raised.  The tick at 20 pt text was twice the tick
#: at 10 pt, so the length follows the text: ``z`` sweeps the size, ``f`` the face (an
#: em rule and an ascent rule part company across faces), and ``ax`` gives each axis its
#: own size to say whose text it is.  ``pos`` moves each axis to the far side of the
#: plot, and ``sec`` is a real secondary value axis, to say which way ``out`` points on
#: an axis that is not at the left or the bottom.  ``rv``/``rc`` split the radar's ticks
#: between its two axes.
OUT_OUT = both("out", "out")
LENGTH_PROBES: list[dict] = [
    *[probe(f"z{size // 100}", size=size, cat=OUT_OUT, val=OUT_OUT) for size in (600, 800, 1400, 2800)],
    *[
        probe(f"f-{face.split()[0]}{size // 100}", size=size, face=face, cat=OUT_OUT, val=OUT_OUT)
        for face in ("Arial", "Courier New", "Times New Roman", "Aptos")
        for size in (1000, 2000)
    ],
    probe(
        "ax-cat20",
        cat={**OUT_OUT, "text": {"size": 2000}},
        val={**OUT_OUT, "text": {"size": 1000}},
    ),
    probe(
        "ax-val20",
        cat={**OUT_OUT, "text": {"size": 1000}},
        val={**OUT_OUT, "text": {"size": 2000}},
    ),
    probe(
        "ax-lblnone20",
        size=2000,
        cat={**OUT_OUT, "labels": "none"},
        val={**OUT_OUT, "labels": "none"},
    ),
    probe("pos-catmax", cat={**OUT_OUT, "crosses": "max"}, val=OUT_OUT),
    probe("pos-valmax", cat=OUT_OUT, val={**OUT_OUT, "crosses": "max"}),
    probe("pos-catmax-in", cat={**both("in", "in"), "crosses": "max"}, val={**both("in"), "crosses": "max"}),
    probe("pos-bar-max", "bar", cat={**OUT_OUT, "crosses": "max"}, val={**OUT_OUT, "crosses": "max"}),
    probe("sec-out", "combo", cat=OUT_OUT, val=OUT_OUT, second=both("out", "out")),
    probe("sec-in", "combo", cat=OUT_OUT, val=OUT_OUT, second=both("in", "none")),
    probe("sec-absent", "combo", cat=OUT_OUT, val=OUT_OUT, second=both(ABSENT, ABSENT)),
    probe("sec-red", "combo", cat=OUT_OUT, val=OUT_OUT, second={**both("cross"), "line": RED3}),
    probe("rv-out", "radar", cat=both("none"), val=both("out")),
    probe("rv-in", "radar", cat=both("none"), val=both("in")),
    probe("rv-cross", "radar", cat=both("none"), val=both("cross")),
    probe("rv-minor", "radar", cat=both("none"), val=both("none", "out")),
    probe("rc-out", "radar", cat=both("out", "out"), val=both("none")),
    probe("rv-red", "radar", cat=both("none"), val={**both("out"), "line": RED3}),
    probe("rv-del", "radar", cat=both("none"), val={**both("out"), "delete": True}),
]

#: The third deck, a check on the rules the first two fitted.  ``m`` reads a lone minor
#: tick in three more faces, where the second deck's minors had a major beside them;
#: ``skip-minor`` says whether ``c:tickMarkSkip`` thins the minor ticks as well;
#: ``col-midcat`` is a column chart told ``crossBetween="midCat"``, ``scat-ymax`` a
#: scatter whose y axis crosses at the maximum, and ``rv-nofill`` a radar whose value axis
#: says ``a:noFill``.
CHECK_PROBES: list[dict] = [
    *[
        probe(f"m-{face.split()[0]}{size // 100}", size=size, face=face, cat=both("none", "out"), val=both("none", "out"))
        for face in ("Arial", "Courier New", "Times New Roman")
        for size in (1000, 2000)
    ],
    probe("skip-minor", cat={**OUT_OUT, "skip": 2}, val=both("none")),
    probe("col-midcat", cross_between="midCat", cat=OUT_OUT, val=both("none")),
    probe("col-neg-low2", values=(3.0, -4.0, 5.0, -2.0, 6.0), cat={**OUT_OUT, "labels": "low"}, val=both("none")),
    probe("scat-ymax", "scatter", cat=OUT_OUT, val={**OUT_OUT, "crosses": "max"}),
    probe("rv-nofill", "radar", cat=both("none"), val={**both("out"), "line": {"none": True}}),
]

#: The fourth deck, on the radar alone: ``chart-gallery``'s radar says ``out`` and drew
#: no tick, where every radar in ``tick-length`` drew them.  Each probe here takes one
#: of that chart's differences from the probes -- its gridlines, its six categories, its
#: second series, its bottom legend, its category axis' own ticks -- to find which.
RADAR_PROBES: list[dict] = [
    probe("r-plain", "radar", cat=both("none"), val=both("out")),
    probe("r-grid", "radar", cat=both("none"), val={**both("out"), "grid": True}),
    probe("r-six", "radar", values=(3.0, 9.0, 5.0, 7.0, 4.0, 6.0), cat=both("none"), val=both("out")),
    probe("r-catout", "radar", cat=both("out"), val=both("out")),
    probe("r-catout-grid", "radar", cat=both("out"), val={**both("out"), "grid": True}),
    probe("r-grid-line", "radar", cat=both("none"), val={**both("out"), "grid": True, "line": RED3}),
    probe("r-grid-minor", "radar", cat=both("none"), val={**both("out", "out"), "grid": True}),
]

#: The fifth deck, on the default.  ``real-college-template``'s chart states
#: ``majorTickMark="none"`` on its category axis and no ``c:minorTickMark`` at all, and
#: PowerPoint drew no minor tick there -- where ``col-absent``, missing both elements,
#: drew both sets across the axis.  Each probe here states one of the two and leaves the
#: other out, and ``d-rct`` copies that axis whole.
DEFAULT_PROBES: list[dict] = [
    probe("d-both-absent", cat=both(ABSENT, ABSENT), val=both(ABSENT, ABSENT)),
    probe("d-maj-none", cat=both("none", ABSENT), val=both("none", ABSENT)),
    probe("d-maj-out", cat=both("out", ABSENT), val=both("out", ABSENT)),
    probe("d-maj-in", cat=both("in", ABSENT), val=both("in", ABSENT)),
    probe("d-min-none", cat=both(ABSENT, "none"), val=both(ABSENT, "none")),
    probe("d-min-out", cat=both(ABSENT, "out"), val=both(ABSENT, "out")),
    probe(
        "d-nodelete",
        cat={**both(ABSENT, ABSENT), "no_delete": True},
        val={**both(ABSENT, ABSENT), "no_delete": True},
    ),
    probe(
        "d-rct",
        cat={
            **both("none", ABSENT),
            "no_delete": True,
            "labels": "low",
            "line": {"w": 17638, "color": "000000"},
        },
        val={**both("none", ABSENT), "no_delete": True, "grid": True, "line": {"none": True}},
    ),
]

#: The sixth deck: ``real-college-template``'s chart itself, and the same part with one
#: thing changed at a time, because ``d-rct`` copied its category axis and did not
#: reproduce it -- that probe drew no axis line at all, where the deck draws one.
COLLEGE = ("~/pptx2svg-oracle/real-college-template.pptx", "ppt/charts/chart1.xml")
MINOR_CROSS = (r'(<c:majorTickMark val="none"/>)', r'\1<c:minorTickMark val="cross"/>')
VERBATIM_PROBES: list[dict] = [
    {**probe("v-asis"), "source": COLLEGE},
    {**probe("v-minor-cross"), "source": COLLEGE, "edits": (MINOR_CROSS,)},
    {
        **probe("v-no-1904"),
        "source": COLLEGE,
        "edits": ((r'<c:date1904 val="1"/><c:lang val="en-US"/>', ""),),
    },
    {
        **probe("v-delete0"),
        "source": COLLEGE,
        "edits": ((r'(</c:scaling>)', r'\1<c:delete val="0"/>'),),
    },
    {
        **probe("v-delete0-minor"),
        "source": COLLEGE,
        "edits": ((r'(</c:scaling>)', r'\1<c:delete val="0"/>'), MINOR_CROSS),
    },
]

#: The seventh deck: the verbatim part drew no axis line until it was told
#: ``<c:delete val="0"/>``, yet the deck it came from draws one without being told.  The
#: one thing the deck has that the probe dropped is the embedded workbook, so these
#: carry one, and ``w-*`` repeat ``tick-default``'s questions with a workbook beside them.
WORKBOOK_PROBES: list[dict] = [
    {**probe("wv-asis"), "source": COLLEGE, "workbook": True},
    {**probe("wv-minor-cross"), "source": COLLEGE, "edits": (MINOR_CROSS,), "workbook": True},
    {**probe("w-both-absent", cat=both(ABSENT, ABSENT), val=both(ABSENT, ABSENT)), "workbook": True},
    {**probe("w-maj-none", cat=both("none", ABSENT), val=both("none", ABSENT)), "workbook": True},
    {**probe("w-maj-out", cat=both("out", ABSENT), val=both("out", ABSENT)), "workbook": True},
    {**probe("w-min-out", cat=both(ABSENT, "out"), val=both(ABSENT, "out")), "workbook": True},
    {
        **probe(
            "w-nodelete",
            cat={**both("out", "out"), "no_delete": True},
            val={**both("out", "out"), "no_delete": True},
        ),
        "workbook": True,
    },
]

#: The eighth set of decks: the workbook changed nothing.  What the source deck has that
#: no probe had is a ``docProps/app.xml`` naming PowerPoint 2007 -- ``AppVersion``
#: 12.0000, where every probe deck names ``pptx-glimpse`` -- so these say that, or another
#: producer, and repeat the questions the others asked of the defaults.  The deck's name
#: picks the producer: ``tick-app-12.pptx``, ``tick-app-14mac.pptx`` and so on.
APP_PROBES: list[dict] = [
    {**probe("a-asis"), "source": COLLEGE},
    *[{**item, "key": "a" + item["key"][1:]} for item in DEFAULT_PROBES],
]

#: ``docProps/app.xml``'s ``Application`` and ``AppVersion``, by deck-name suffix.
APP_VERSIONS = {
    "12": ("Microsoft Office PowerPoint", "12.0000"),
    "12mac": ("Microsoft Macintosh PowerPoint", "12.0000"),
    "14": ("Microsoft Office PowerPoint", "14.0000"),
    "14mac": ("Microsoft Macintosh PowerPoint", "14.0000"),
    "15": ("Microsoft Office PowerPoint", "15.0000"),
    "16": ("Microsoft Office PowerPoint", "16.0000"),
}


def set_app_version(target: Path, application: str, version: str) -> None:
    """Rewrite the deck's ``docProps/app.xml`` to name ``application`` at ``version``."""
    import re
    import zipfile

    with zipfile.ZipFile(target) as source:
        entries = [(info, source.read(info.filename)) for info in source.infolist()]
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            if info.filename == "docProps/app.xml":
                text = re.sub(r"<AppVersion>.*?</AppVersion>", "", data.decode())
                text = re.sub(
                    r"<Application>.*?</Application>",
                    f"<Application>{application}</Application><AppVersion>{version}</AppVersion>",
                    text,
                )
                data = text.encode()
            out.writestr(info.filename, data)


DECKS = {
    "tick-marks": MARKS_PROBES,
    "tick-length": LENGTH_PROBES,
    "tick-check": CHECK_PROBES,
    "tick-radar": RADAR_PROBES,
    "tick-default": DEFAULT_PROBES,
    "tick-verbatim": VERBATIM_PROBES,
    "tick-workbook": WORKBOOK_PROBES,
    "tick-app-": APP_PROBES,
}


def probes_for(target: Path) -> list[dict]:
    for name, table in DECKS.items():
        if name in target.stem:
            return table
    raise SystemExit(f"name the deck after one of {sorted(DECKS)}, not {target.stem!r}")


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else Path("tick-marks.pptx")
    probes = probes_for(target)
    write_deck(target, probes, part_for=any_part)
    if "tick-app-" in target.stem:
        suffix = target.stem.split("tick-app-", 1)[1]
        if suffix not in APP_VERSIONS:
            raise SystemExit(f"name a producer from {sorted(APP_VERSIONS)}, not {suffix!r}")
        set_app_version(target, *APP_VERSIONS[suffix])
    if any(item.get("workbook") for item in probes):
        import zipfile

        embed_workbooks(
            target,
            probes,
            zipfile.ZipFile(Path(COLLEGE[0]).expanduser()).read(
                "ppt/embeddings/Microsoft_Office_Excel_Worksheet1.xlsx"
            ),
        )
    for index, item in enumerate(probes):
        print(f"{index + 1:3d} {item['key']}")
    print(f"\n{len(probes)} probes -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

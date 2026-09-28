#!/usr/bin/env python3
"""Build the probe deck for what the one-rasteriser instrument exposed (ROADMAP.md 0.5).

Five findings came out of recalibrating the fidelity instrument, and each needs PowerPoint
itself to say what the right drawing is before anything is changed.  One deck asks all of
them, so one export answers them, and ``tools/read_exposed_probe.py`` reads it back:

* **text** -- where a shape's text sits in a preset's *text rectangle* (``a:rect``).
  Every shape holds one short run of 14 pt Arial whose string names the probe; a plain
  ``rect`` with no outline is the control, so a reading is the difference between a
  shape's first glyph and the control's and no font metric enters it.  Each geometry is
  set three ways -- top-left, bottom-right and centred -- which reads all four sides of
  its rectangle.  Then the **outline**: a 1, 4 and 8 pt line, an 8 pt ``a:noFill`` one
  that still states its width, a width that comes from the theme's line style
  (``p:style/a:lnRef``), and an 8 pt line under zero insets -- on a ``rect``, a
  ``roundRect`` and an ``ellipse``.  docx2svg measured Word laying text inside half the
  outline's width; this asks PowerPoint.  Last, three flipped shapes whose rectangle is
  not symmetric.
* **icc** -- pictures carrying an ICC profile: a PNG with an ``iCCP`` chunk and a JPEG
  with an ``APP2`` one, under two synthetic matrix/TRC profiles (Adobe RGB's primaries at
  gamma 563/256; Display P3's at gamma 1.8), plus the same patches untagged.  Flat
  patches, so what PowerPoint wrote -- the samples, the colour space, the profile -- is
  read per patch.
* **radar** -- the order a radar's parts are painted in: every part in its own colour (the
  spokes red, the rings blue, the value axis' line green) and the series wide, in the
  three ``radarStyle`` s.
* **legend3d** -- a ``line3DChart``'s legend key at three positions and two sizes, with a
  flat ``lineChart`` as the control.
* **scene** -- ``line3DChart`` and ``bar3DChart`` frames from 150 pt to the whole slide,
  to see whether the raster PowerPoint draws a 3-D scene into is held at 300 dpi or
  capped at some pixel size.

A second deck, ``two``, asks what the first export raised: the JPEGs again, each distinct
in a pixel (the first deck's three identical ones came back as one image object); where a
radar's spokes take their line from; and a 3-D scene's category tick marks.  A third,
``three``, sets text taller than its text rectangle at each anchor, to see which way it
spills.

Usage::

    python3 tools/make_exposed_probe.py ~/pptx2svg-oracle/exposed-probe.pptx [two]
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/exposed-probe.pptx ~/pptx2svg-oracle/exposed-probe.pdf
    mv ~/pptx2svg-oracle/exposed-probe.* somewhere-else/
    python3 tools/read_exposed_probe.py somewhere-else/exposed-probe.pdf [text|icc|...]

(name the second deck ``exposed-probe-2``: the reader picks the layout by that suffix).
PowerPoint is sandboxed: both paths must be in a directory it has already been granted
(``~/pptx2svg-oracle``).  **Move the deck and its PDF out as soon as the export is
written**: ``tools/fidelity.py`` scores every deck in that directory, and caches its
conversion and rasters there.  The JPEGs are written with Pillow, so this is a
development tool.
"""

from __future__ import annotations

import io
import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_chart_gallery import (  # noqa: E402
    A,
    C,
    NO_TITLE,
    QUARTERS,
    R,
    VIEW_3D,
    cat_axis,
    cats,
    chart_space,
    legend,
    nums,
    plot_area,
    ser_axis,
    series_name,
    tail,
    val_axis,
)

HERE = Path(__file__).resolve().parent
SOURCE = HERE.parent / "tests/fixtures/authoring-integration.pptx"

NS = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
)
SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"

EMU_PER_PT = 12700
SLIDE_W, SLIDE_H = 9144000, 5143500

# --------------------------------------------------------------------------------------
# text: the text rectangle
# --------------------------------------------------------------------------------------

#: Every shape's box.  Wide enough that no probe string wraps in any rectangle here.
BOX_W, BOX_H = 2400000, 1400000
#: Three columns and two rows a slide.
COLUMNS = (300000, 3372000, 6444000)
ROWS = (700000, 3000000)
TEXT_SIZE = 1400

#: ``(key, geometry xml)``.  ``custom`` states its own ``a:rect`` a quarter in from the
#: top and left; ``rect`` is the control every reading is taken against.
GEOMETRIES = [
    ("rect", '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'),
    ("roundRect", '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'),
    ("roundRect40", '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 40000"/></a:avLst></a:prstGeom>'),
    ("ellipse", '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'),
    ("octagon", '<a:prstGeom prst="octagon"><a:avLst/></a:prstGeom>'),
    ("triangle", '<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>'),
    ("rtTriangle", '<a:prstGeom prst="rtTriangle"><a:avLst/></a:prstGeom>'),
    ("diamond", '<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>'),
    ("hexagon", '<a:prstGeom prst="hexagon"><a:avLst/></a:prstGeom>'),
    ("rightArrow", '<a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>'),
    ("wedgeRectCallout", '<a:prstGeom prst="wedgeRectCallout"><a:avLst/></a:prstGeom>'),
    ("cloud", '<a:prstGeom prst="cloud"><a:avLst/></a:prstGeom>'),
    ("custom",
     '<a:custGeom><a:avLst/><a:gdLst><a:gd name="q" fmla="*/ w 1 4"/><a:gd name="v" fmla="*/ h 1 4"/>'
     '</a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="q" t="v" r="r" b="b"/><a:pathLst><a:path w="100" h="100">'
     '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="0"/></a:lnTo>'
     '<a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/>'
     "</a:path></a:pathLst></a:custGeom>"),
]

#: The three placements: which corner or centre the run is anchored to.
PLACEMENTS = [("tl", "t", "l"), ("br", "b", "r"), ("cc", "ctr", "ctr")]

NO_LINE = "<a:ln><a:noFill/></a:ln>"


def _line(width: int | None, fill: bool = True) -> str:
    paint = '<a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill>' if fill else "<a:noFill/>"
    return f'<a:ln{f" w={chr(34)}{width}{chr(34)}" if width is not None else ""}>{paint}</a:ln>'


#: ``(key, spPr line xml, p:style xml, insets)``.  ``style3`` states no ``a:ln`` at all
#: and takes the theme's third line (1.5 pt) through ``a:lnRef idx="3"``.
OUTLINES = [
    ("ln1", _line(12700), "", None),
    ("ln4", _line(50800), "", None),
    ("ln8", _line(101600), "", None),
    ("ln8nofill", _line(101600, fill=False), "", None),
    ("style3", "",
     '<p:style><a:lnRef idx="3"><a:srgbClr val="7F7F7F"/></a:lnRef><a:fillRef idx="0">'
     '<a:srgbClr val="FFFFFF"/></a:fillRef><a:effectRef idx="0"><a:srgbClr val="000000"/>'
     '</a:effectRef><a:fontRef idx="minor"><a:srgbClr val="000000"/></a:fontRef></p:style>', None),
    ("ln8zero", _line(101600), "", (0, 0, 0, 0)),
]
OUTLINED = ["rect", "roundRect", "ellipse"]

#: ``(key, geometry, flipH, flipV)``.
FLIPS = [
    ("triangle-flipV", "triangle", False, True),
    ("rtTriangle-flipH", "rtTriangle", True, False),
    ("rightArrow-flipH", "rightArrow", True, False),
]


def text_probes() -> list[dict]:
    """Every text probe, in slide order: its key, the run it holds and how it is set."""
    geometry = dict(GEOMETRIES)
    out: list[dict] = []
    for key, _ in GEOMETRIES:
        for place, anchor, algn in PLACEMENTS:
            out.append(dict(key=f"{key}/{place}", geometry=geometry[key], anchor=anchor, algn=algn,
                            line=NO_LINE, style="", insets=None, flip=(False, False)))
    for name in OUTLINED:
        for key, line, style, insets in OUTLINES:
            for place, anchor, algn in PLACEMENTS[:2]:
                out.append(dict(key=f"{name}+{key}/{place}", geometry=geometry[name], anchor=anchor,
                                algn=algn, line=line, style=style, insets=insets, flip=(False, False)))
    # The controls the outline family needs: zero insets, no outline, both corners.
    for place, anchor, algn in PLACEMENTS[:2]:
        out.append(dict(key=f"rect+zero/{place}", geometry=geometry["rect"], anchor=anchor, algn=algn,
                        line=NO_LINE, style="", insets=(0, 0, 0, 0), flip=(False, False)))
    for key, name, flip_h, flip_v in FLIPS:
        for place, anchor, algn in PLACEMENTS[:2]:
            out.append(dict(key=f"{key}/{place}", geometry=geometry[name], anchor=anchor, algn=algn,
                            line=NO_LINE, style="", insets=None, flip=(flip_h, flip_v)))
    for index, probe in enumerate(out):
        # The run: a tag the reader finds the probe by.  Upper-case letters and digits
        # only, so no glyph has a descender or an accent to move its box.
        probe["text"] = f"H{index:03d}"
    return out


def text_shape(shape_id: int, x: int, y: int, probe: dict) -> str:
    flip_h, flip_v = probe["flip"]
    flips = (' flipH="1"' if flip_h else "") + (' flipV="1"' if flip_v else "")
    insets = ""
    if probe["insets"] is not None:
        left, top, right, bottom = probe["insets"]
        insets = f' lIns="{left}" tIns="{top}" rIns="{right}" bIns="{bottom}"'
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(probe["key"])}"/><p:cNvSpPr/><p:nvPr/>'
        f"</p:nvSpPr><p:spPr><a:xfrm{flips}><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{BOX_W}\" cy=\"{BOX_H}\"/>"
        f"</a:xfrm>{probe['geometry']}"
        '<a:solidFill><a:srgbClr val="E8EEF8"/></a:solidFill>'
        f"{probe['line']}</p:spPr>{probe['style']}"
        f'<p:txBody><a:bodyPr wrap="square" anchor="{probe["anchor"]}"{insets}/><a:lstStyle/>'
        f'<a:p><a:pPr algn="{probe["algn"]}"><a:buNone/></a:pPr>'
        f'<a:r><a:rPr lang="en-US" sz="{TEXT_SIZE}" dirty="0"><a:solidFill><a:srgbClr val="000000"/>'
        f'</a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>{probe["text"]}</a:t></a:r></a:p>'
        "</p:txBody></p:sp>"
    )


def text_slides() -> list[dict]:
    probes = text_probes()
    slides = []
    per = len(COLUMNS) * len(ROWS)
    for start in range(0, len(probes), per):
        shapes = []
        for offset, probe in enumerate(probes[start:start + per]):
            x = COLUMNS[offset % len(COLUMNS)]
            y = ROWS[offset // len(COLUMNS)]
            probe["box"] = (x, y, BOX_W, BOX_H)
            shapes.append(text_shape(10 + offset, x, y, probe))
        slides.append(dict(body="".join(shapes), parts=[]))
    return slides


# --------------------------------------------------------------------------------------
# icc: pictures with embedded profiles
# --------------------------------------------------------------------------------------

#: Flat patches, 32 px square, two rows of four.
PATCHES = [
    (255, 0, 0), (0, 255, 0), (0, 0, 255), (128, 128, 128),
    (200, 100, 50), (30, 60, 90), (240, 230, 10), (12, 12, 12),
]
PATCH = 32

#: D50-adapted primaries as an ICC profile states them.  Adobe RGB (1998)'s, and Display
#: P3's (Bradford-adapted, as Apple's profile states them).
ADOBE_PRIMARIES = ((0.60974, 0.31111, 0.01947), (0.20528, 0.62567, 0.06087), (0.14919, 0.06322, 0.74457))
P3_PRIMARIES = ((0.51512, 0.24120, -0.00105), (0.29198, 0.69225, 0.04189), (0.15710, 0.06657, 0.78407))
D50 = (0.9642, 1.0, 0.8249)


def _s15(value: float) -> bytes:
    return struct.pack(">i", round(value * 65536))


def matrix_profile(description: str, primaries, gamma: float) -> bytes:
    """A version 2 RGB display profile: three primaries, one gamma, D50 white."""

    def xyz(values) -> bytes:
        return b"XYZ \x00\x00\x00\x00" + b"".join(_s15(v) for v in values)

    def curve(value: float) -> bytes:
        return b"curv\x00\x00\x00\x00" + struct.pack(">IH", 1, round(value * 256)) + b"\x00\x00"

    text = description.encode("ascii")
    desc = (b"desc\x00\x00\x00\x00" + struct.pack(">I", len(text) + 1) + text + b"\x00"
            + b"\x00" * 8 + b"\x00" * 3 + b"\x00" * 67)
    copyright_ = b"text\x00\x00\x00\x00" + b"No copyright, synthetic.\x00"
    tags = [
        (b"desc", desc), (b"cprt", copyright_), (b"wtpt", xyz(D50)),
        (b"rXYZ", xyz(primaries[0])), (b"gXYZ", xyz(primaries[1])), (b"bXYZ", xyz(primaries[2])),
        (b"rTRC", curve(gamma)), (b"gTRC", curve(gamma)), (b"bTRC", curve(gamma)),
    ]
    offset = 128 + 4 + 12 * len(tags)
    table, data = b"", b""
    for signature, body in tags:
        while (offset + len(data)) % 4:
            data += b"\x00"
        table += signature + struct.pack(">II", offset + len(data), len(body))
        data += body
    size = offset + len(data)
    header = (
        struct.pack(">I", size) + b"none" + bytes([2, 0x10, 0, 0]) + b"mntrRGB XYZ "
        + struct.pack(">6H", 2026, 1, 1, 0, 0, 0) + b"acsp" + b"APPL" + b"\x00" * 4
        + b"none" + b"none" + b"\x00" * 8 + struct.pack(">I", 0)
        + _s15(D50[0]) + _s15(D50[1]) + _s15(D50[2]) + b"none" + b"\x00" * 44
    )
    assert len(header) == 128, len(header)
    return header + struct.pack(">I", len(tags)) + table + data


def patch_pixels(variant: int = 0) -> tuple[int, int, bytes]:
    """The patches as RGB rows.  ``variant`` moves the last patch by that many levels, so
    two pictures that must stay two pictures in PowerPoint's export differ in a pixel and
    not only in their metadata: the first deck's three JPEGs, identical but for their
    profiles, came back as **one** image object drawn three times."""
    width, height = PATCH * 4, PATCH * 2
    colours = list(PATCHES)
    colours[-1] = tuple(min(255, value + variant) for value in colours[-1])
    rows = []
    for y in range(height):
        row = bytearray()
        for x in range(width):
            row += bytes(colours[(y // PATCH) * 4 + x // PATCH])
        rows.append(bytes(row))
    return width, height, b"".join(rows)


def _chunk(tag: bytes, body: bytes) -> bytes:
    return struct.pack(">I", len(body)) + tag + body + struct.pack(">I", zlib.crc32(tag + body))


def png(profile: bytes | None, name: bytes = b"probe", variant: int = 0) -> bytes:
    width, height, pixels = patch_pixels(variant)
    stride = width * 3
    raw = b"".join(b"\x00" + pixels[y * stride:(y + 1) * stride] for y in range(height))
    chunks = [_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))]
    if profile is not None:
        chunks.append(_chunk(b"iCCP", name + b"\x00\x00" + zlib.compress(profile)))
    chunks.append(_chunk(b"IDAT", zlib.compress(raw, 9)))
    chunks.append(_chunk(b"IEND", b""))
    return b"\x89PNG\r\n\x1a\n" + b"".join(chunks)


def jpeg(profile: bytes | None, variant: int = 0) -> bytes:
    from PIL import Image

    width, height, pixels = patch_pixels(variant)
    image = Image.frombytes("RGB", (width, height), pixels)
    buffer = io.BytesIO()
    options = dict(quality=100, subsampling=0)
    if profile is not None:
        options["icc_profile"] = profile
    image.save(buffer, "JPEG", **options)
    return buffer.getvalue()


#: ``(key, extension, bytes-builder)``: what the reader looks for, in picture order.
PICTURES = [
    ("png-plain", "png", lambda: png(None)),
    ("png-adobe", "png", lambda: png(matrix_profile("Probe Adobe primaries 2.2", ADOBE_PRIMARIES, 563 / 256))),
    ("png-p3-18", "png", lambda: png(matrix_profile("Probe P3 primaries 1.8", P3_PRIMARIES, 1.8))),
    ("jpeg-plain", "jpeg", lambda: jpeg(None)),
    ("jpeg-adobe", "jpeg", lambda: jpeg(matrix_profile("Probe Adobe primaries 2.2", ADOBE_PRIMARIES, 563 / 256))),
    ("jpeg-p3-18", "jpeg", lambda: jpeg(matrix_profile("Probe P3 primaries 1.8", P3_PRIMARIES, 1.8))),
]
#: The second deck's: every JPEG distinct in its pixels, and one a slide.
PICTURES_TWO = [
    ("jpeg-plain", "jpeg", lambda: jpeg(None, 0)),
    ("jpeg-adobe", "jpeg", lambda: jpeg(matrix_profile("Probe Adobe primaries 2.2", ADOBE_PRIMARIES, 563 / 256), 2)),
    ("jpeg-p3-18", "jpeg", lambda: jpeg(matrix_profile("Probe P3 primaries 1.8", P3_PRIMARIES, 1.8), 4)),
]
#: Each picture is drawn 2 pt a pixel, so a patch is 64 pt square.
PICTURE_SCALE = 2 * EMU_PER_PT


def picture_slides(pictures=None, per: int = 3) -> list[dict]:
    pictures = PICTURES if pictures is None else pictures
    slides = []
    for start in range(0, len(pictures), per):
        body, parts = [], []
        for offset, (key, extension, build) in enumerate(pictures[start:start + per]):
            rid = f"rIdPic{offset}"
            target = f"../media/{key}.{extension}"
            x = 300000 + offset * 2900000
            y = 700000
            cx, cy = PATCH * 4 * PICTURE_SCALE // 2, PATCH * 2 * PICTURE_SCALE // 2
            body.append(
                f'<p:pic><p:nvPicPr><p:cNvPr id="{20 + offset}" name="{key}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
                f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
            )
            parts.append((rid, REL + "image", target, f"ppt/media/{key}.{extension}", build(), None))
        slides.append(dict(body="".join(body), parts=parts, keys=[key for key, _, _ in pictures[start:start + per]]))
    return slides


# --------------------------------------------------------------------------------------
# charts: radar order, the 3-D line legend, the 3-D scene's raster
# --------------------------------------------------------------------------------------


def _ln(width: int, colour: str) -> str:
    return f'<c:spPr><a:ln w="{width}"><a:solidFill><a:srgbClr val="{colour}"/></a:solidFill></a:ln></c:spPr>'


def radar(style: str, *, cat_line: bool = True, val_line: str | None = "solid", val_delete: bool = False) -> bytes:
    """A six-spoke radar with every part in its own colour: the category axis' line red,
    the rings blue, the value axis' line green (``val_line``: ``"solid"``, ``"none"`` for
    an explicit ``a:noFill``, ``None`` for no ``c:spPr`` at all)."""
    spokes = ("Speed", "Cost", "Quality", "Support", "Reach", "Risk")
    series = []
    for index, (name, row, colour) in enumerate(
        zip(("A", "B"), ((80, 55, 90, 65, 45, 70), (60, 70, 75, 50, 60, 55)), ("F2A33A", "7030A0"))
    ):
        paint = (f'<c:spPr><a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>'
                 f'<a:ln w="50800"><a:solidFill><a:srgbClr val="{colour}"/></a:solidFill></a:ln></c:spPr>')
        series.append(
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>" + series_name(name) + paint
            + ("<c:marker><c:symbol val='circle'/><c:size val='9'/></c:marker>" if style == "marker" else "")
            + cats(spokes) + nums("val", row) + "</c:ser>"
        )
    groups = (f"<c:radarChart><c:radarStyle val='{style}'/><c:varyColors val='0'/>" + "".join(series)
              + "<c:axId val='301'/><c:axId val='302'/></c:radarChart>")
    axes = (
        "<c:catAx><c:axId val='301'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/><c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/><c:tickLblPos val='nextTo'/>"
        + (_ln(38100, "FF0000") if cat_line else "")
        + "<c:crossAx val='302'/><c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/><c:noMultiLvlLbl val='0'/></c:catAx>"
        "<c:valAx><c:axId val='302'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if val_delete else 0}'/><c:axPos val='l'/><c:majorGridlines>" + _ln(25400, "0000FF")
        + "</c:majorGridlines>"
        "<c:numFmt formatCode='General' sourceLinked='1'/><c:majorTickMark val='cross'/>"
        "<c:minorTickMark val='none'/><c:tickLblPos val='nextTo'/>"
        + {"solid": _ln(19050, "00B050"), "none": "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>", None: ""}[val_line]
        + "<c:crossAx val='301'/><c:crosses val='autoZero'/><c:crossBetween val='between'/></c:valAx>"
    )
    return chart_space(NO_TITLE + plot_area(groups, axes) + tail(legend("r")))


def line3d(position: str = "b", *, marker: bool = False, flat: bool = False, size: int = 1000) -> bytes:
    series = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>" + series_name(name)
        + ("<c:marker><c:symbol val='circle'/><c:size val='7'/></c:marker>" if marker else "")
        + cats(QUARTERS) + nums("val", row) + "<c:smooth val='0'/></c:ser>"
        for i, (name, row) in enumerate(zip(("North", "South", "East"), ((14, 21, 19, 27), (9, 13, 22, 18), (5, 8, 12, 9))))
    )
    if flat:
        groups = ("<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>" + series
                  + "<c:marker val='1'/><c:axId val='401'/><c:axId val='402'/></c:lineChart>")
        return chart_space(NO_TITLE + plot_area(groups, cat_axis(401, 402) + val_axis(402, 401))
                           + tail(legend(position)), size=size)
    groups = ("<c:line3DChart><c:grouping val='standard'/><c:varyColors val='0'/>" + series
              + "<c:gapDepth val='150'/><c:axId val='401'/><c:axId val='402'/><c:axId val='403'/></c:line3DChart>")
    axes = cat_axis(401, 402) + val_axis(402, 401) + ser_axis(403, 402)
    return chart_space(VIEW_3D + NO_TITLE + plot_area(groups, axes) + tail(legend(position)), size=size)


def bar3d() -> bytes:
    series = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>" + series_name(name) + cats(QUARTERS)
        + nums("val", row) + "</c:ser>"
        for i, (name, row) in enumerate(zip(("Plan", "Actual"), ((12, 18, 15, 22), (10, 20, 14, 25))))
    )
    groups = ("<c:bar3DChart><c:barDir val='col'/><c:grouping val='clustered'/><c:varyColors val='0'/>"
              + series + "<c:gapWidth val='150'/><c:shape val='box'/>"
              "<c:axId val='501'/><c:axId val='502'/><c:axId val='503'/></c:bar3DChart>")
    axes = cat_axis(501, 502) + val_axis(502, 501) + ser_axis(503, 502)
    return chart_space(VIEW_3D + NO_TITLE + plot_area(groups, axes) + tail(legend("b")))


def _pt(value: float) -> int:
    return round(value * EMU_PER_PT)


#: ``(key, builder, frame (x, y, w, h) in pt)``.  One chart a slide.
CHARTS = [
    ("radar-standard", lambda: radar("standard"), (36, 30, 500, 350)),
    ("radar-marker", lambda: radar("marker"), (36, 30, 500, 350)),
    ("radar-filled", lambda: radar("filled"), (36, 30, 500, 350)),
    ("legend3d-b", lambda: line3d("b"), (36, 30, 520, 336)),
    ("legend3d-r", lambda: line3d("r"), (36, 30, 520, 336)),
    ("legend3d-t14", lambda: line3d("t", size=1400), (36, 30, 520, 336)),
    ("legend3d-marker", lambda: line3d("b", marker=True), (36, 30, 520, 336)),
    ("legend-flat-b", lambda: line3d("b", flat=True), (36, 30, 520, 336)),
    ("scene-line-150", lambda: line3d("b"), (36, 30, 150, 110)),
    ("scene-line-full", lambda: line3d("b"), (4, 4, 712, 397)),
    ("scene-bar-full", bar3d, (4, 4, 712, 397)),
]


def bar3d_axis(tick: str = "out", line: bool = False) -> bytes:
    """A ``bar3DChart`` whose category axis states its tick marks, and optionally a red line."""
    axis = cat_axis(501, 502).replace("<c:majorTickMark val='out'/>", f"<c:majorTickMark val='{tick}'/>")
    if line:
        axis = axis.replace("<c:crossAx", _ln(28575, "FF0000") + "<c:crossAx", 1)
    series = "".join(
        f"<c:ser><c:idx val='{i}'/><c:order val='{i}'/>" + series_name(name) + cats(QUARTERS)
        + nums("val", row) + "</c:ser>"
        for i, (name, row) in enumerate(zip(("Plan", "Actual"), ((12, 18, 15, 22), (10, 20, 14, 25))))
    )
    groups = ("<c:bar3DChart><c:barDir val='col'/><c:grouping val='clustered'/><c:varyColors val='0'/>"
              + series + "<c:gapWidth val='150'/><c:shape val='box'/>"
              "<c:axId val='501'/><c:axId val='502'/><c:axId val='503'/></c:bar3DChart>")
    axes = axis + val_axis(502, 501) + ser_axis(503, 502)
    return chart_space(VIEW_3D + NO_TITLE + plot_area(groups, axes) + tail(legend("b")))


#: The second deck's charts: where a radar's spokes take their line from, and a 3-D
#: scene's category axis line and ticks.
CHARTS_TWO = [
    ("radar-filled-cat-only", lambda: radar("filled", val_line=None), (36, 30, 500, 350)),
    ("radar-filled-val-only", lambda: radar("filled", cat_line=False), (36, 30, 500, 350)),
    ("radar-filled-neither", lambda: radar("filled", cat_line=False, val_line=None), (36, 30, 500, 350)),
    ("radar-filled-val-nofill", lambda: radar("filled", val_line="none"), (36, 30, 500, 350)),
    ("radar-filled-val-deleted", lambda: radar("filled", val_delete=True), (36, 30, 500, 350)),
    ("radar-marker-cat-only", lambda: radar("marker", val_line=None), (36, 30, 500, 350)),
    ("radar-marker-val-only", lambda: radar("marker", cat_line=False), (36, 30, 500, 350)),
    ("bar3d-tick-out", lambda: bar3d_axis("out"), (36, 30, 520, 336)),
    ("bar3d-tick-none", lambda: bar3d_axis("none"), (36, 30, 520, 336)),
    ("bar3d-tick-out-red", lambda: bar3d_axis("out", line=True), (36, 30, 520, 336)),
]


def chart_slides(start: int, charts=None) -> list[dict]:
    charts = CHARTS if charts is None else charts
    slides = []
    for offset, (key, build, (x, y, w, h)) in enumerate(charts):
        number = start + offset
        body = (
            f"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='30' name='{key}'/><p:cNvGraphicFramePr/><p:nvPr/>"
            f"</p:nvGraphicFramePr><p:xfrm><a:off x='{_pt(x)}' y='{_pt(y)}'/><a:ext cx='{_pt(w)}' cy='{_pt(h)}'/>"
            "</p:xfrm><a:graphic><a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/chart'>"
            f"<c:chart {C} {R} r:id='rIdChart'/></a:graphicData></a:graphic></p:graphicFrame>"
        )
        parts = [("rIdChart", REL + "chart", f"../charts/chart{number}.xml", f"ppt/charts/chart{number}.xml",
                  build(), CHART_TYPE)]
        slides.append(dict(body=body, parts=parts, key=key))
    return slides


# --------------------------------------------------------------------------------------
# The package
# --------------------------------------------------------------------------------------

EMPTY_TREE = (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
)


def slide_xml(body: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f"<p:sld {NS}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + body + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


#: The third deck: text taller than its text rectangle, at each anchor.  Three lines of
#: 14 pt in a box 60 pt high, so every one overflows; which way it spills is the reading.
OVERFLOW_BOX = (2400000, 762000)
OVERFLOW_GEOMETRIES = ["rect", "ellipse", "roundRect"]


def overflow_probes() -> list[dict]:
    geometry = dict(GEOMETRIES)
    out = []
    for name in OVERFLOW_GEOMETRIES:
        for anchor in ("t", "ctr", "b"):
            out.append(dict(key=f"{name}/{anchor}", geometry=geometry[name], anchor=anchor))
    for index, probe in enumerate(out):
        probe["lines"] = [f"V{index:02d}{tag}" for tag in "ABC"]
    return out


def overflow_slides() -> list[dict]:
    probes = overflow_probes()
    slides = []
    per = len(COLUMNS) * len(ROWS)
    width, height = OVERFLOW_BOX
    for start in range(0, len(probes), per):
        shapes = []
        for offset, probe in enumerate(probes[start:start + per]):
            x = COLUMNS[offset % len(COLUMNS)]
            y = ROWS[offset // len(COLUMNS)] + 600000
            paragraphs = "".join(
                '<a:p><a:pPr algn="l"><a:buNone/></a:pPr>'
                f'<a:r><a:rPr lang="en-US" sz="{TEXT_SIZE}" dirty="0"><a:solidFill><a:srgbClr val="000000"/>'
                f'</a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>{line}</a:t></a:r></a:p>'
                for line in probe["lines"]
            )
            shapes.append(
                f'<p:sp><p:nvSpPr><p:cNvPr id="{10 + offset}" name="{probe["key"]}"/><p:cNvSpPr/><p:nvPr/>'
                f'</p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/>'
                f"</a:xfrm>{probe['geometry']}"
                '<a:solidFill><a:srgbClr val="E8EEF8"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
                f'<p:txBody><a:bodyPr wrap="square" anchor="{probe["anchor"]}"/><a:lstStyle/>{paragraphs}'
                "</p:txBody></p:sp>"
            )
        slides.append(dict(body="".join(shapes), parts=[], overflow=[p["key"] for p in probes[start:start + per]]))
    return slides


def all_slides(deck: str = "one") -> list[dict]:
    """``one``: text, pictures and the first charts.  ``two``: the follow-ups the first
    export asked for -- distinct JPEGs, and the radar's and the 3-D axis' line sources."""
    if deck == "two":
        return picture_slides(PICTURES_TWO, per=1) + chart_slides(1, CHARTS_TWO)
    if deck == "three":
        return overflow_slides()
    slides = text_slides()
    slides += picture_slides()
    slides += chart_slides(1)
    return slides


def write_deck(target: Path, deck: str = "one") -> list[dict]:
    slides = all_slides(deck)
    source = zipfile.ZipFile(SOURCE)
    presentation = source.read("ppt/presentation.xml").decode()
    pres_rels = source.read("ppt/_rels/presentation.xml.rels").decode()
    types = source.read("[Content_Types].xml").decode()
    numbers = range(1, len(slides) + 1)
    presentation = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>",
                          "<p:sldIdLst>" + "".join(f'<p:sldId id="{255 + n}" r:id="rIdS{n}"/>' for n in numbers)
                          + "</p:sldIdLst>", presentation, flags=re.S)
    pres_rels = re.sub(r'<Relationship[^>]*Type="[^"]*/slide"[^>]*/>', "", pres_rels)
    pres_rels = pres_rels.replace("</Relationships>", "".join(
        f'<Relationship Id="rIdS{n}" Type="{REL}slide" Target="slides/slide{n}.xml"/>' for n in numbers
    ) + "</Relationships>")
    types = re.sub(r'<Override PartName="/ppt/(slides/slide|charts/chart|embeddings/)[^>]*/>', "", types)
    extra = "".join(f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{SLIDE_TYPE}"/>' for n in numbers)
    for slide in slides:
        for _, _, _, part, _, content_type in slide["parts"]:
            if content_type:
                extra += f'<Override PartName="/{part}" ContentType="{content_type}"/>'
    if 'Extension="jpeg"' not in types:
        extra += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
    types = types.replace("</Types>", extra + "</Types>")

    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            name = item.filename
            if name.startswith(("ppt/slides/", "ppt/charts/", "ppt/embeddings/", "ppt/media/")):
                continue
            data = source.read(name)
            if name == "ppt/presentation.xml":
                data = presentation.encode()
            elif name == "ppt/_rels/presentation.xml.rels":
                data = pres_rels.encode()
            elif name == "[Content_Types].xml":
                data = types.encode()
            elif re.search(r"slide(Layouts|Masters)/slide\w+\d+\.xml$", name):
                xml = re.sub(r"<p:spTree>.*?</p:spTree>", EMPTY_TREE, data.decode(), flags=re.S)
                data = re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S).encode()
            out.writestr(name, data)
        for number, slide in enumerate(slides, start=1):
            rels = [f'<Relationship Id="rId1" Type="{REL}slideLayout" Target="../slideLayouts/slideLayout1.xml"/>']
            for rid, rel_type, target_path, part, data, _ in slide["parts"]:
                rels.append(f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target_path}"/>')
                out.writestr(part, data)
            out.writestr(f"ppt/slides/slide{number}.xml", slide_xml(slide["body"]))
            out.writestr(f"ppt/slides/_rels/slide{number}.xml.rels",
                         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                         + "".join(rels) + "</Relationships>")
    return slides


def main() -> int:
    if len(sys.argv) not in (2, 3):
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    slides = write_deck(target, sys.argv[2] if len(sys.argv) == 3 else "one")
    print(f"{len(slides)} slides, {len(text_probes())} text probes -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""Build ``tests/fixtures/feature-sweep.pptx`` -- one unexercised feature per slide.

Why this deck exists
--------------------

This project has been bitten three times by the same defect: markup the parser reads and
the renderer throws away.  ``flat_chart_kind`` erased the 3-D spelling of a chart type;
a ChartEx frame warned about the wrong thing; ``a:clrChange`` was parsed, resolved onto
``BlipEffects.clr_change`` and read by nothing at all.  The last of those was found by a
coverage test rather than by anyone looking at a picture, because **nothing in the corpus
uses it** -- there was no snapshot to move and no fidelity score to fall.

That is the gap this deck closes.  A sweep of the three surfaces (model fields the
renderer ignores, parsed fields that never reach the model, and XML read nowhere) turned
up a list of features with no committed deck behind them; the ones worth having a picture
of are here, one per slide, so a score localises to a cause the way ``chart-gallery``'s
per-slide table did for three separate chart defects.

**Not every slide is an assertion of correctness.**  Six of the twelve are features this
sweep fixed; the other six are features it deliberately did not, pinned so that their
current behaviour is recorded honestly rather than left invisible.  Which is which is
stated per slide in ``SLIDES`` below and in ``tests/fixtures/FIXTURES-README.md``; read
the baseline for a pinned slide as "this is what we do today", never as "this is right".

Authored to be *scorable*
-------------------------

A deck ``tools/fidelity.py`` skips is worth far less than one it scores -- four of the
corpus decks are skipped because PowerPoint itself substitutes a face, which turns the
comparison into a measurement of font availability.  So, exactly as ``make_chart_gallery``
does:

* **No CJK anywhere**, and no character outside Latin-1 on any slide.
* **Nothing names a typeface.**  No ``a:latin``, no ``a:ea``, no ``a:cs``, and in
  particular no ``a:buFont`` -- ``fidelity.requested_faces`` counts that one, because a
  bullet's face is a face PowerPoint really does embed.  Every face is the theme's, and
  the theme is the Aptos one carried by ``authoring-integration.pptx``, which the oracle
  already scores.  ``requested_faces`` therefore sees exactly ``Aptos`` and
  ``Aptos Display``.

Where the deck comes from
-------------------------

Entirely ours.  The package skeleton -- theme, master, layout, ``presentation.xml`` -- is
lifted from ``tests/fixtures/authoring-integration.pptx``, which this repository generates
from its own authoring API, and the master and layout are then emptied.  Every slide, and
every image in ``ppt/media/``, is written by the functions below: the rasters are encoded
here with :mod:`zlib` a few dozen bytes at a time, and the one vector is a circle and a
rectangle.  No third-party deck, no third-party artwork.

The pictures are flat blocks of *exact* colours on purpose.  ``a:clrChange`` matches
exactly, so a photograph would have nothing to key on; a PNG is lossless, so the colour
PowerPoint decodes and the colour resvg decodes are the same one this file wrote.

Authoring hazards this file works around
----------------------------------------

* ``tests/deckbuilder.py`` and this file both write content-type ``Override`` elements
  with the leading slash **already prepended**.  Passing ``f"/{part}"`` yields
  ``PartName="//ppt/..."``, which PowerPoint opens as ``[Repaired]`` -- a deck it repairs
  is not a fixture.
* OOXML child order is enforced.  ``a:blip``'s children go before ``a:srcRect``/
  ``a:stretch``; ``a:extLst`` goes last within ``a:blip``; ``a:ln``'s ``a:solidFill``
  precedes ``a:prstDash``.
* ``mc:AlternateContent`` must declare ``mc:`` on the element itself, and the
  ``Requires`` value has to be a prefix that is actually in scope, or PowerPoint repairs
  the slide.

Usage::

    python3 tools/make_feature_sweep.py                        # rewrite the fixture
    python3 tools/make_feature_sweep.py ~/pptx2svg-oracle/feature-sweep.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/feature-sweep.pptx ~/pptx2svg-oracle/feature-sweep.pdf

Then rebaseline the snapshots and the oracle::

    python3 -m pytest tests/test_vrt.py --update-snapshots
    python3 tools/fidelity.py --update
"""

from __future__ import annotations

import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SOURCE = ROOT / "tests/fixtures/authoring-integration.pptx"
TARGET = ROOT / "tests/fixtures/feature-sweep.pptx"

NAMESPACES = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG"'
)

SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
LAYOUT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

#: The GUID Office writes on the extension that carries a picture's vector original.
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"

EMU_PER_POINT = 12700
PT = EMU_PER_POINT

#: The slide is 720 x 405 pt.  A caption on the first band, the demonstration below it.
CAPTION_OFF = (36 * PT, 14 * PT)
CAPTION_EXT = (648 * PT, 22 * PT)
BODY_TOP = 52 * PT
BODY_LEFT = 36 * PT

EMPTY_TREE = (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
)


# --------------------------------------------------------------------------------------
# Artwork.  Written here so the deck carries nothing whose provenance cannot be read.
# --------------------------------------------------------------------------------------


def png(rows: list[list[tuple[int, int, int, int]]]) -> bytes:
    """A minimal RGBA PNG.  No filtering, one IDAT -- the encoder is the point, not speed."""
    height, width = len(rows), len(rows[0])

    def chunk(tag: bytes, payload: bytes) -> bytes:
        return (
            struct.pack(">I", len(payload))
            + tag
            + payload
            + struct.pack(">I", zlib.crc32(tag + payload) & 0xFFFFFFFF)
        )

    raw = b"".join(
        b"\x00" + b"".join(struct.pack("BBBB", *pixel) for pixel in row) for row in rows
    )
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", _stored_deflate(raw))
        + chunk(b"IEND", b"")
    )


def _stored_deflate(raw: bytes) -> bytes:
    """A zlib stream of *raw* in uncompressed blocks, byte-identical on every platform.

    ``zlib.compress`` is **not** a deterministic function of its input: the output depends
    on the zlib build CPython was linked against, and Windows ships a different one.  The
    committed fixture is checked byte-for-byte against what this generator writes -- that
    is the deck's provenance claim -- so an encoder that varies by platform failed all four
    Windows legs of CI while passing everywhere else.

    Deflate's stored block is fully specified by RFC 1951 and has no encoder freedom at
    all, so writing it by hand removes the dependency rather than pinning a version of it.
    These images are a few hundred bytes of flat colour; the ~0.1% the compression bought
    is not worth a fixture that only reproduces on the machine that wrote it.
    """
    blocks = bytearray(b"\x78\x01")  # zlib header: deflate, 32K window, no preset dict
    for start in range(0, max(len(raw), 1), 0xFFFF):
        piece = raw[start : start + 0xFFFF]
        final = 1 if start + 0xFFFF >= len(raw) else 0
        blocks += bytes([final])
        blocks += struct.pack("<HH", len(piece), len(piece) ^ 0xFFFF)
        blocks += piece
    blocks += struct.pack(">I", zlib.adler32(raw) & 0xFFFFFFFF)
    return bytes(blocks)


#: Four exact, saturated blocks.  ``a:clrChange`` keys on the first of them, so they have
#: to be colours a lossless round trip preserves bit for bit -- which is why this is a
#: grid of flat fills and not a gradient or a photograph.
SWATCH_COLORS = ((0xE0, 0x3C, 0x31), (0x2E, 0x7D, 0x32), (0x15, 0x65, 0xC0), (0xF9, 0xA8, 0x25))
CLR_CHANGE_FROM = "E03C31"
CLR_CHANGE_TO = "6A1B9A"


def swatch_png() -> bytes:
    """Four 48 px blocks side by side, plus a white keyline so the edges are visible."""
    size, count = 48, len(SWATCH_COLORS)
    rows: list[list[tuple[int, int, int, int]]] = []
    for y in range(size):
        row: list[tuple[int, int, int, int]] = []
        for x in range(size * count):
            border = y < 2 or y >= size - 2 or x % size < 2 or x % size >= size - 2
            red, green, blue = SWATCH_COLORS[x // size]
            row.append((255, 255, 255, 255) if border else (red, green, blue, 255))
        rows.append(row)
    return png(rows)


def bullet_png() -> bytes:
    """A 32 px diamond on transparency -- a picture bullet with real alpha to preserve."""
    size = 32
    half = size // 2
    rows = []
    for y in range(size):
        row = []
        for x in range(size):
            inside = abs(x - half) + abs(y - half) < half - 2
            row.append((0x1D, 0x4E, 0xD8, 255) if inside else (0, 0, 0, 0))
        rows.append(row)
    return png(rows)


def tile_png() -> bytes:
    """A 32 px L, strongly asymmetric so a mirrored tile is unmistakable."""
    size = 32
    rows = []
    for y in range(size):
        row = []
        for x in range(size):
            ink = x < 10 or y >= size - 10
            row.append((0x0F, 0x76, 0x6E, 255) if ink else (0xE2, 0xF5, 0xF3, 255))
        rows.append(row)
    return png(rows)


def svg_mark() -> bytes:
    """A *faithful* vector of :func:`swatch_png` -- the same four blocks and keyline.

    It is tempting to draw something visibly different here so that "which branch did we
    take" can be read off the picture.  That would make a bad instrument: this slide would
    then score badly **because** the renderer did the right thing, and a fidelity number
    that punishes correctness measures nothing.

    PowerPoint settles it.  Its own PDF export of this deck draws the *raster* on both
    sides -- it treats ``asvg:svgBlip`` as provenance, not as what to rasterise -- so a
    deliberately different vector would show up as a whole-picture mismatch against the
    reference on every future run, for ever, with no defect behind it.

    Which branch was taken is pinned where it belongs instead: the committed VRT snapshot
    contains ``data:image/svg+xml`` rather than ``data:image/png``, and
    ``tests/test_unrendered_sweep.py`` asserts the embedded bytes outright.
    """
    size, count = 48, len(SWATCH_COLORS)
    blocks = "".join(
        f'<rect x="{index * size + 2}" y="2" width="{size - 4}" height="{size - 4}" '
        f'fill="#{red:02x}{green:02x}{blue:02x}"/>'
        for index, (red, green, blue) in enumerate(SWATCH_COLORS)
    )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {size * count} {size}" '
        f'width="{size * count}" height="{size}" preserveAspectRatio="none">'
        f'<rect width="{size * count}" height="{size}" fill="#ffffff"/>{blocks}</svg>'
    ).encode()

MEDIA = {
    "ppt/media/sweep-swatch.png": swatch_png,
    "ppt/media/sweep-bullet.png": bullet_png,
    "ppt/media/sweep-tile.png": tile_png,
    "ppt/media/sweep-mark.svg": svg_mark,
}

#: Every slide gets the same relationship set, so a slide can use any image without this
#: file tracking which.  An unused relationship is valid OPC and PowerPoint keeps it.
SLIDE_RELATIONSHIPS = (
    ("rIdSwatch", IMAGE_REL, "../media/sweep-swatch.png"),
    ("rIdBullet", IMAGE_REL, "../media/sweep-bullet.png"),
    ("rIdTile", IMAGE_REL, "../media/sweep-tile.png"),
    ("rIdMark", IMAGE_REL, "../media/sweep-mark.svg"),
)


# --------------------------------------------------------------------------------------
# Shape helpers
# --------------------------------------------------------------------------------------

_ids = iter(range(200, 100000))


def _next_id() -> int:
    return next(_ids)


def frame(x: int, y: int, cx: int, cy: int) -> str:
    return f"<a:xfrm><a:off x='{x}' y='{y}'/><a:ext cx='{cx}' cy='{cy}'/></a:xfrm>"


def label(text: str, x: int, y: int, cx: int, *, size: int = 900) -> str:
    """A caption under a demonstration.  Names no face; the theme's minor font draws it."""
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{_next_id()}' name='Label'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        + frame(x, y, cx, 16 * PT)
        + "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/>"
        "<a:ln><a:noFill/></a:ln></p:spPr><p:txBody>"
        "<a:bodyPr wrap='square' lIns='0' rIns='0' tIns='0' bIns='0'/><a:lstStyle/>"
        f"<a:p><a:r><a:rPr lang='en-US' sz='{size}'>"
        "<a:solidFill><a:srgbClr val='475569'/></a:solidFill></a:rPr>"
        f"<a:t>{escape(text)}</a:t></a:r>"
        f"<a:endParaRPr lang='en-US' sz='{size}'/></a:p></p:txBody></p:sp>"
    )


def picture(
    rel: str,
    x: int,
    y: int,
    cx: int,
    cy: int,
    *,
    name: str = "Picture",
    blip_children: str = "",
    blip_extension: str = "",
) -> str:
    """``p:pic``.  ``a:blip``'s children come first and its ``a:extLst`` last."""
    return (
        f"<p:pic><p:nvPicPr><p:cNvPr id='{_next_id()}' name='{escape(name)}'/>"
        "<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>"
        f"<a:blip r:embed='{rel}'>{blip_children}{blip_extension}</a:blip>"
        "<a:stretch><a:fillRect/></a:stretch></p:blipFill>"
        f"<p:spPr>{frame(x, y, cx, cy)}"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom></p:spPr></p:pic>"
    )


def svg_extension(rel: str) -> str:
    return (
        f"<a:extLst><a:ext uri='{SVG_EXT_URI}'>"
        f"<asvg:svgBlip r:embed='{rel}'/></a:ext></a:extLst>"
    )


def shape(x: int, y: int, cx: int, cy: int, *, fill: str, line: str = "", geom: str = "rect",
          name: str = "Shape", body: str = "") -> str:
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{_next_id()}' name='{escape(name)}'/>"
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        + frame(x, y, cx, cy)
        + f"<a:prstGeom prst='{geom}'><a:avLst/></a:prstGeom>{fill}{line}"
        + "</p:spPr>"
        + (body or "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>")
        + "</p:sp>"
    )


def solid(hex_value: str) -> str:
    return f"<a:solidFill><a:srgbClr val='{hex_value}'/></a:solidFill>"


def outline(hex_value: str, width_pt: float, *, compound: str | None = None) -> str:
    attrs = f"w='{int(width_pt * PT)}'" + (f" cmpd='{compound}'" if compound else "")
    return f"<a:ln {attrs}>{solid(hex_value)}</a:ln>"


def text_box(x: int, y: int, cx: int, cy: int, paragraphs: str, *, body_pr: str = "<a:bodyPr/>",
             name: str = "Text") -> str:
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{_next_id()}' name='{escape(name)}'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        + frame(x, y, cx, cy)
        + "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        f"<p:txBody>{body_pr}<a:lstStyle/>{paragraphs}</p:txBody></p:sp>"
    )


def run(text: str, size: int) -> str:
    return f"<a:r><a:rPr lang='en-US' sz='{size}'/><a:t>{escape(text)}</a:t></a:r>"


# --------------------------------------------------------------------------------------
# Slides.  One feature each.
# --------------------------------------------------------------------------------------


def slide_clr_change() -> str:
    """FIXED.  ``a:clrChange`` -- one exact colour replaced throughout a picture."""
    change = (
        "<a:clrChange>"
        f"<a:clrFrom><a:srgbClr val='{CLR_CHANGE_FROM}'/></a:clrFrom>"
        f"<a:clrTo><a:srgbClr val='{CLR_CHANGE_TO}'/></a:clrTo>"
        "</a:clrChange>"
    )
    return (
        picture("rIdSwatch", BODY_LEFT, BODY_TOP, 256 * PT, 64 * PT, name="Untouched")
        + label("as authored", BODY_LEFT, BODY_TOP + 70 * PT, 256 * PT)
        + picture("rIdSwatch", BODY_LEFT + 320 * PT, BODY_TOP, 256 * PT, 64 * PT,
                  name="Recoloured", blip_children=change)
        + label(f"a:clrChange  #{CLR_CHANGE_FROM} -> #{CLR_CHANGE_TO}",
                BODY_LEFT + 320 * PT, BODY_TOP + 70 * PT, 256 * PT)
        + label("Only the first block changes: a:clrChange matches exactly, so the other "
                "three and the white keyline are untouched.",
                BODY_LEFT, BODY_TOP + 120 * PT, 620 * PT)
    )


#: The swatch artwork's own aspect, so neither the raster nor the vector is stretched.
ART_W, ART_H = 240 * PT, 60 * PT


def slide_svg_blip() -> str:
    """FIXED.  ``asvg:svgBlip`` -- the vector a picture was rasterised from."""
    return (
        picture("rIdSwatch", BODY_LEFT, BODY_TOP, ART_W, ART_H, name="Raster only")
        + label("a:blip alone -- the PNG is drawn", BODY_LEFT, BODY_TOP + 68 * PT, 300 * PT)
        + picture("rIdSwatch", BODY_LEFT, BODY_TOP + 100 * PT, ART_W, ART_H,
                  name="Vector preferred", blip_extension=svg_extension("rIdMark"))
        + label("same a:blip plus an asvg:svgBlip extension -- the SVG is drawn",
                BODY_LEFT, BODY_TOP + 168 * PT, 460 * PT)
        + label("Both draw the same artwork, deliberately: PowerPoint's own PDF export "
                "rasterises from the PNG on both, so a vector drawn differently would "
                "score as a defect for ever with none behind it. Which source we used is "
                "pinned by the VRT snapshot, which holds data:image/svg+xml here.",
                BODY_LEFT, BODY_TOP + 210 * PT, 620 * PT)
    )


def slide_alternate_content() -> str:
    """FIXED.  ``mc:AlternateContent`` -- the Choice we can draw beats the Fallback."""
    choice = picture(
        "rIdSwatch", BODY_LEFT, BODY_TOP + 100 * PT, ART_W, ART_H,
        name="Choice vector", blip_extension=svg_extension("rIdMark"),
    )
    fallback = picture(
        "rIdSwatch", BODY_LEFT, BODY_TOP + 100 * PT, ART_W, ART_H, name="Fallback raster",
    )
    return (
        picture("rIdSwatch", BODY_LEFT, BODY_TOP, ART_W, ART_H, name="Reference")
        + label("a plain p:pic, in no branch at all", BODY_LEFT, BODY_TOP + 68 * PT, 320 * PT)
        + "<mc:AlternateContent>"
        + f"<mc:Choice Requires='asvg'>{choice}</mc:Choice>"
        + f"<mc:Fallback>{fallback}</mc:Fallback>"
        + "</mc:AlternateContent>"
        + label("mc:Choice (carries the vector) beside mc:Fallback (raster only)",
                BODY_LEFT, BODY_TOP + 168 * PT, 460 * PT)
        + label("Same artwork on both, for the reason slide 2 gives. The Choice is taken: "
                "the snapshot's second picture is data:image/svg+xml, where taking the "
                "Fallback unconditionally -- which this used to do -- gave a PNG.",
                BODY_LEFT, BODY_TOP + 210 * PT, 620 * PT)
    )


def slide_bu_blip() -> str:
    """FIXED.  ``a:buBlip`` -- a picture used as the bullet glyph."""
    bullet = "<a:buBlip><a:blip r:embed='rIdBullet'/></a:buBlip>"
    paragraphs = "".join(
        f"<a:p><a:pPr marL='457200' indent='-457200'>{bullet}</a:pPr>"
        f"{run(text, 1600)}</a:p>"
        for text in (
            "Picture bullets come from a:buBlip",
            "They are not a character in any face",
            "So they cannot be a tspan the way buChar is",
        )
    )
    return (
        text_box(BODY_LEFT, BODY_TOP, 400 * PT, 120 * PT, paragraphs, name="Picture bullets")
        + label("a:buBlip -- a 32 px diamond with real alpha, drawn as <image> beside the "
                "<text> because <image> is not a text content element.",
                BODY_LEFT, BODY_TOP + 130 * PT, 620 * PT)
    )


def slide_bu_sz_pts() -> str:
    """FIXED.  ``a:buSzPts`` -- the absolute spelling of a bullet's size."""
    def bulleted(pr: str, text: str, size: int) -> str:
        return f"<a:p><a:pPr marL='457200' indent='-457200'>{pr}</a:pPr>{run(text, size)}</a:p>"

    dot = "<a:buChar char='&#8226;'/>"
    paragraphs = (
        bulleted(dot, "no size given: the bullet takes the run's 32 pt", 3200)
        + bulleted(f"<a:buSzPct val='50000'/>{dot}", "a:buSzPct 50%: half the run", 3200)
        + bulleted(f"<a:buSzPts val='1000'/>{dot}", "a:buSzPts 10 pt: absolute", 3200)
    )
    return (
        text_box(BODY_LEFT, BODY_TOP, 620 * PT, 180 * PT, paragraphs, name="Bullet sizes")
        + label("Only a:buSzPct was read, so every a:buSzPts bullet silently took the "
                "run's size. Decks usually author the two equal, which is why the gap "
                "never showed.",
                BODY_LEFT, BODY_TOP + 200 * PT, 620 * PT)
    )


def slide_alpha_mod_fix() -> str:
    """FIXED.  ``a:alphaModFix`` -- picture opacity."""
    backdrop = shape(BODY_LEFT - 8 * PT, BODY_TOP - 8 * PT, 600 * PT, 96 * PT,
                     fill=solid("0F172A"), name="Backdrop")
    return (
        backdrop
        + picture("rIdSwatch", BODY_LEFT, BODY_TOP, 176 * PT, 44 * PT, name="Opaque")
        + picture("rIdSwatch", BODY_LEFT + 200 * PT, BODY_TOP, 176 * PT, 44 * PT,
                  name="Half", blip_children="<a:alphaModFix amt='50000'/>")
        + picture("rIdSwatch", BODY_LEFT + 400 * PT, BODY_TOP, 176 * PT, 44 * PT,
                  name="Faint", blip_children="<a:alphaModFix amt='20000'/>")
        + label("100%", BODY_LEFT, BODY_TOP + 52 * PT, 176 * PT)
        + label("a:alphaModFix 50%", BODY_LEFT + 200 * PT, BODY_TOP + 52 * PT, 176 * PT)
        + label("a:alphaModFix 20%", BODY_LEFT + 400 * PT, BODY_TOP + 52 * PT, 176 * PT)
        + label("A bare <a:alphaModFix/> means 100% and is dropped rather than emitting a "
                "no-op filter -- Google Slides writes one on every exported picture.",
                BODY_LEFT, BODY_TOP + 120 * PT, 620 * PT)
    )


def slide_compound_lines() -> str:
    """PINNED, NOT FIXED.  ``a:ln@cmpd`` -- two or three strokes across one width.

    SVG gives a path exactly one stroke, centred on it, and there is no way to offset a
    stroke outward from an arbitrary path.  All five boxes below are drawn identically:
    one 6 pt stroke.  The resolver emits ``line-compound-flattened`` to say so.
    """
    spellings = ("sng", "dbl", "thickThin", "thinThick", "tri")
    parts = []
    for index, value in enumerate(spellings):
        x = BODY_LEFT + index * 124 * PT
        parts.append(
            shape(x, BODY_TOP, 108 * PT, 72 * PT, fill=solid("F8FAFC"),
                  line=outline("1F4E79", 6, compound=value), name=f"cmpd {value}")
        )
        parts.append(label(f"cmpd='{value}'", x, BODY_TOP + 80 * PT, 108 * PT))
    parts.append(
        label("NOT RENDERED. All five are drawn as one 6 pt stroke: right colour, right "
              "weight, right place, missing split. Warns line-compound-flattened.",
              BODY_LEFT, BODY_TOP + 120 * PT, 620 * PT)
    )
    return "".join(parts)


def slide_gradient_paths() -> str:
    """PINNED, NOT FIXED.  ``a:path@path`` -- circle, rect and shape all become radial.

    ``parse_gradient_fill`` reads ``a:path``'s ``a:fillToRect`` and never its ``path``
    attribute, so a rectangular or shape-following gradient is drawn as an ellipse.
    """
    stops = (
        "<a:gsLst><a:gs pos='0'><a:srgbClr val='F9A825'/></a:gs>"
        "<a:gs pos='100000'><a:srgbClr val='0F766E'/></a:gs></a:gsLst>"
    )
    rect = "<a:fillToRect l='50000' t='50000' r='50000' b='50000'/>"
    variants = (
        ("lin ang=0", f"<a:gradFill>{stops}<a:lin ang='0' scaled='0'/></a:gradFill>"),
        ("path=circle", f"<a:gradFill>{stops}<a:path path='circle'>{rect}</a:path></a:gradFill>"),
        ("path=rect", f"<a:gradFill>{stops}<a:path path='rect'>{rect}</a:path></a:gradFill>"),
        ("path=shape", f"<a:gradFill>{stops}<a:path path='shape'>{rect}</a:path></a:gradFill>"),
    )
    parts = []
    for index, (name, fill) in enumerate(variants):
        x = BODY_LEFT + index * 156 * PT
        parts.append(shape(x, BODY_TOP, 136 * PT, 96 * PT, fill=fill, name=name))
        parts.append(label(name, x, BODY_TOP + 104 * PT, 136 * PT))
    parts.append(
        label("PARTLY RENDERED. a:path@path is read nowhere, so circle, rect and shape "
              "all come out as the same radial gradient.",
              BODY_LEFT, BODY_TOP + 140 * PT, 620 * PT)
    )
    return "".join(parts)


def slide_pattern_fills() -> str:
    """RENDERED.  ``a:pattFill`` -- a regression pin for a feature with no fixture.

    Nothing in the corpus uses a pattern fill at all, so this slide is the only thing that
    would catch the cell geometry drifting -- and it earned its place immediately: at
    SSIM 0.0563 it is what showed that the cell was a hardcoded 8 *pixels* (6 pt) where
    PowerPoint's is 8 **points**, and that ``dkDnDiag`` was drawn as two hairlines where
    PowerPoint draws a 2 px diagonal.  Both are fixed from measurements taken with
    ``tools/make_fill_probe.py``; every preset here is now the bitmap PowerPoint's own PDF
    export carries for it.

    Read the score with care.  Both rasterisers quantise a pattern's period to whole
    device pixels and round it opposite ways: at the fidelity harness's 1280 px the 8 pt
    cell is 14.22 px, pdfium draws it at 15 and resvg at 14, which drifts a pixel a cell
    and holds SSIM near zero however right the vector output is.  Rendered at 5760 px the
    two agree rule for rule.  The histogram is the number that moved: 0.8798 to 0.9357.
    """
    presets = ("ltUpDiag", "dkDnDiag", "cross", "diagCross", "horz", "pct25")
    parts = []
    for index, preset in enumerate(presets):
        x = BODY_LEFT + (index % 3) * 208 * PT
        y = BODY_TOP + (index // 3) * 116 * PT
        fill = (
            f"<a:pattFill prst='{preset}'>"
            "<a:fgClr><a:srgbClr val='0F766E'/></a:fgClr>"
            "<a:bgClr><a:srgbClr val='FFFFFF'/></a:bgClr></a:pattFill>"
        )
        parts.append(shape(x, y, 180 * PT, 84 * PT, fill=fill,
                           line=outline("94A3B8", 0.75), name=preset))
        parts.append(label(f"a:pattFill prst='{preset}'", x, y + 90 * PT, 180 * PT))
    return "".join(parts)


def slide_tile_fill() -> str:
    """``a:tile@flip`` and ``@algn`` -- **the on-slide caption and label are now stale.**

    They say NOT RENDERED, and both are drawn: a mirrored axis doubles the SVG pattern
    cell and holds the mirror inside it, which is exactly what PowerPoint's own export
    does.  The deeper defect this slide pinned is fixed too -- the tile was sized at
    ``sx`` of the *shape's box* and came out 8.3x too big, where ``sx`` scales the
    picture's own natural size (``tools/make_fill_probe.py``, deck ``fill-tile``).

    The text is left wrong on purpose, for now.  Correcting it rewrites
    ``tests/fixtures/feature-sweep.pptx``, which forces a fresh PowerPoint export of
    ``feature-sweep.pdf`` -- the oracle every recorded number for this deck was measured
    against.  Doing that in the same change as a renderer fix would mix two causes in one
    before-and-after table.  Regenerate the deck and re-export it as its own change, and
    take ``state`` here from ``pinned`` to ``rendered`` with it.
    """
    parts = []
    for index, (flip, align) in enumerate(
        (("none", "tl"), ("x", "tl"), ("xy", "tl"), ("none", "ctr"))
    ):
        x = BODY_LEFT + index * 156 * PT
        fill = (
            "<a:blipFill><a:blip r:embed='rIdTile'/>"
            f"<a:tile tx='0' ty='0' sx='60000' sy='60000' flip='{flip}' algn='{align}'/>"
            "</a:blipFill>"
        )
        parts.append(shape(x, BODY_TOP, 136 * PT, 96 * PT, fill=fill,
                           line=outline("94A3B8", 0.75), name=f"tile {flip} {align}"))
        parts.append(label(f"flip='{flip}' algn='{align}'", x, BODY_TOP + 104 * PT, 136 * PT))
    parts.append(
        label("NOT RENDERED. An SVG <pattern> repeats one tile unchanged, so flip and "
              "algn are both dropped and all four tile identically.",
              BODY_LEFT, BODY_TOP + 140 * PT, 620 * PT)
    )
    return "".join(parts)


def slide_anchor_ctr() -> str:
    """PINNED, NOT FIXED.  ``a:bodyPr@anchorCtr`` -- centre the text block, not the lines.

    It centres the *bounding box of the lines* horizontally in the shape while leaving
    each line's own alignment alone.  ``anchorCtr`` appears nowhere in ``src/``.

    The corpus carries 654 of them and **every one is ``anchorCtr="0"``**, the default --
    so the gap costs the corpus nothing and a count of occurrences would have been a
    misleading way to rank it.  This slide is the only place the attribute is ever set,
    which is exactly why it had to be authored rather than found.
    """
    paragraphs = (
        f"<a:p><a:pPr algn='l'/>{run('short', 1600)}</a:p>"
        f"<a:p><a:pPr algn='l'/>{run('a much longer line than the first', 1600)}</a:p>"
        f"<a:p><a:pPr algn='l'/>{run('mid length line', 1600)}</a:p>"
    )
    return (
        shape(BODY_LEFT, BODY_TOP, 280 * PT, 110 * PT, fill=solid("F1F5F9"),
              line=outline("94A3B8", 0.75), name="anchorCtr off",
              body=f"<p:txBody><a:bodyPr wrap='square' anchorCtr='0'/>"
                   f"<a:lstStyle/>{paragraphs}</p:txBody>")
        + label("anchorCtr='0'", BODY_LEFT, BODY_TOP + 118 * PT, 280 * PT)
        + shape(BODY_LEFT + 320 * PT, BODY_TOP, 280 * PT, 110 * PT, fill=solid("F1F5F9"),
                line=outline("94A3B8", 0.75), name="anchorCtr on",
                body=f"<p:txBody><a:bodyPr wrap='square' anchorCtr='1'/>"
                     f"<a:lstStyle/>{paragraphs}</p:txBody>")
        + label("anchorCtr='1'", BODY_LEFT + 320 * PT, BODY_TOP + 118 * PT, 280 * PT)
        + label("NOT RENDERED. anchorCtr appears nowhere in src/, so both boxes are drawn "
                "left-aligned; PowerPoint centres the right-hand block as a whole.",
                BODY_LEFT, BODY_TOP + 150 * PT, 620 * PT)
    )


def slide_shadow_and_blur() -> str:
    """PINNED, NOT FIXED.  Two ``UNRENDERED_FIELDS`` excuses, restated by this sweep.

    ``OuterShadow.rotate_with_shape``: ``rotWithShape='0'`` asks the shadow to keep
    pointing the same way while the shape turns.  The filter is emitted inside the
    shape's own transform, so it always turns with it.

    ``BlurEffect.grow``: the excuse said "we always let the blur grow past the shape's
    bounds".  The blip filter region is fixed at 0%/100%, so it is the *grow=1* case that
    loses its halo -- the opposite of what was claimed.
    """
    def shadow(rot: str) -> str:
        return (
            "<a:effectLst><a:outerShdw blurRad='76200' dist='114300' dir='2700000' "
            f"rotWithShape='{rot}'><a:srgbClr val='0F172A'><a:alpha val='55000'/>"
            "</a:srgbClr></a:outerShdw></a:effectLst>"
        )

    def box(x: int, rotation: int, rot_with_shape: str, name: str) -> str:
        """The same blue box twice, differing only in rotation and in rotWithShape."""
        turn = f" rot='{rotation}'" if rotation else ""
        return (
            f"<p:sp><p:nvSpPr><p:cNvPr id='{_next_id()}' name='{name}'/>"
            "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
            f"<a:xfrm{turn}><a:off x='{x}' y='{BODY_TOP}'/>"
            f"<a:ext cx='{120 * PT}' cy='{80 * PT}'/></a:xfrm>"
            "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>"
            + solid("1D4ED8")
            + shadow(rot_with_shape)
            + "</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
        )

    parts = [
        box(BODY_LEFT, 0, "1", "Shadow upright"),
        label("upright, rotWithShape='1'", BODY_LEFT, BODY_TOP + 116 * PT, 180 * PT),
        box(BODY_LEFT + 200 * PT, 1800000, "0", "Shadow turned"),
        label("rot=30deg, rotWithShape='0'", BODY_LEFT + 200 * PT,
              BODY_TOP + 116 * PT, 200 * PT),
    ]
    parts.append(
        picture("rIdSwatch", BODY_LEFT + 430 * PT, BODY_TOP, 160 * PT, 40 * PT,
                name="Blur grow", blip_children="<a:blur rad='76200' grow='1'/>")
    )
    parts.append(label("a:blur grow='1'", BODY_LEFT + 430 * PT, BODY_TOP + 116 * PT, 180 * PT))
    parts.append(
        label("NOT RENDERED. The shadow turns with the shape whatever rotWithShape says, "
              "and the blur is clipped at the picture's edge whatever grow says.",
              BODY_LEFT, BODY_TOP + 150 * PT, 620 * PT)
    )
    return "".join(parts)


def slide_table_compound() -> str:
    """PINNED, NOT FIXED.  ``a:lnL/lnR/lnT/lnB@cmpd`` on table cell borders.

    Why this is a slide of its own: the corpus carries 716 compound spellings on cell
    borders and 139 on shape outlines, and the warning is emitted once per spelling per
    slide precisely so a table like this one does not produce hundreds of lines.
    """
    def cell(text: str) -> str:
        border = (
            "<a:lnL w='38100' cmpd='dbl'>" + solid("1F4E79") + "</a:lnL>"
            "<a:lnR w='38100' cmpd='dbl'>" + solid("1F4E79") + "</a:lnR>"
            "<a:lnT w='38100' cmpd='dbl'>" + solid("1F4E79") + "</a:lnT>"
            "<a:lnB w='38100' cmpd='dbl'>" + solid("1F4E79") + "</a:lnB>"
        )
        return (
            "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>"
            f"<a:p><a:pPr algn='ctr'/>{run(text, 1400)}</a:p></a:txBody>"
            f"<a:tcPr marL='45720' marR='45720' marT='45720' marB='45720'>{border}"
            + solid("FFFFFF")
            + "</a:tcPr></a:tc>"
        )

    rows = "".join(
        "<a:tr h='" + str(36 * PT) + "'>"
        + "".join(cell(text) for text in row)
        + "</a:tr>"
        for row in (("Region", "Q1", "Q2"), ("North", "184", "211"), ("South", "142", "168"))
    )
    grid = "".join(f"<a:gridCol w='{160 * PT}'/>" for _ in range(3))
    return (
        f"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='{_next_id()}' name='Compound table'/>"
        "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        f"<p:xfrm><a:off x='{BODY_LEFT}' y='{BODY_TOP}'/>"
        f"<a:ext cx='{480 * PT}' cy='{108 * PT}'/></p:xfrm>"
        "<a:graphic><a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/table'>"
        "<a:tbl><a:tblPr firstRow='0' bandRow='0'/>"
        f"<a:tblGrid>{grid}</a:tblGrid>{rows}</a:tbl>"
        "</a:graphicData></a:graphic></p:graphicFrame>"
        + label("NOT RENDERED. Every cell border is cmpd='dbl'; all twelve are drawn as "
                "one 3 pt stroke, and line-compound-flattened fires once for the slide, "
                "not once per border.",
                BODY_LEFT, BODY_TOP + 130 * PT, 620 * PT)
    )


SLIDES = [
    {"key": "clr-change", "build": slide_clr_change, "state": "fixed",
     "caption": "1  a:clrChange - one exact colour replaced throughout a picture"},
    {"key": "svg-blip", "build": slide_svg_blip, "state": "fixed",
     "caption": "2  asvg:svgBlip - the vector a picture was rasterised from"},
    {"key": "alternate-content", "build": slide_alternate_content, "state": "fixed",
     "caption": "3  mc:AlternateContent - a Choice we can draw beats the Fallback"},
    {"key": "bu-blip", "build": slide_bu_blip, "state": "fixed",
     "caption": "4  a:buBlip - picture bullets"},
    {"key": "bu-sz-pts", "build": slide_bu_sz_pts, "state": "fixed",
     "caption": "5  a:buSzPts - a bullet sized in points rather than in percent"},
    {"key": "alpha-mod-fix", "build": slide_alpha_mod_fix, "state": "fixed",
     "caption": "6  a:alphaModFix - picture opacity"},
    {"key": "compound-lines", "build": slide_compound_lines, "state": "pinned",
     "caption": "7  a:ln@cmpd - compound lines (NOT RENDERED; warns)"},
    {"key": "table-compound", "build": slide_table_compound, "state": "pinned",
     "caption": "8  a:lnL/R/T/B@cmpd - compound cell borders (NOT RENDERED; warns once)"},
    {"key": "gradient-paths", "build": slide_gradient_paths, "state": "pinned",
     "caption": "9  a:path@path - circle, rect and shape (PARTLY RENDERED)"},
    {"key": "tile-fill", "build": slide_tile_fill, "state": "pinned",
     "caption": "10  a:tile@flip and @algn - mirrored tiles (NOT RENDERED)"},
    {"key": "anchor-ctr", "build": slide_anchor_ctr, "state": "pinned",
     "caption": "11  a:bodyPr@anchorCtr - centre the text block (NOT RENDERED)"},
    {"key": "shadow-blur", "build": slide_shadow_and_blur, "state": "pinned",
     "caption": "12  a:outerShdw@rotWithShape and a:blur@grow (NOT RENDERED)"},
    {"key": "pattern-fills", "build": slide_pattern_fills, "state": "rendered",
     "caption": "13  a:pattFill - pattern presets (RENDERED; regression pin)"},
]


# --------------------------------------------------------------------------------------
# Packaging
# --------------------------------------------------------------------------------------


def caption_shape(index: int, text: str) -> str:
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{10 + index}' name='Caption'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        f"<a:xfrm><a:off x='{CAPTION_OFF[0]}' y='{CAPTION_OFF[1]}'/>"
        f"<a:ext cx='{CAPTION_EXT[0]}' cy='{CAPTION_EXT[1]}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>"
        "<a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>"
        "<p:txBody><a:bodyPr wrap='square' lIns='0' rIns='0' tIns='0' bIns='0'/>"
        "<a:lstStyle/><a:p>"
        "<a:r><a:rPr lang='en-US' sz='1100' b='1'>"
        "<a:solidFill><a:srgbClr val='0F172A'/></a:solidFill></a:rPr>"
        f"<a:t>{escape(text)}</a:t></a:r><a:endParaRPr lang='en-US' sz='1100'/>"
        "</a:p></p:txBody></p:sp>"
    )


def slide_xml(index: int, spec: dict) -> str:
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + caption_shape(index, spec["caption"])
        + spec["build"]()
        + "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


def slide_rels(index: int) -> str:
    extra = "".join(
        f"<Relationship Id='{rel_id}' Type='{rel_type}' Target='{target}'/>"
        for rel_id, rel_type, target in SLIDE_RELATIONSHIPS
    )
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        f"<Relationship Id='rId1' Type='{LAYOUT_REL}' Target='../slideLayouts/slideLayout1.xml'/>"
        f"{extra}</Relationships>"
    )


def blank_template(xml: str) -> str:
    """Empty a layout or master so nothing but each slide's own content is drawn."""
    xml = re.sub(r"<p:spTree>.*?</p:spTree>", EMPTY_TREE, xml, flags=re.S)
    return re.sub(r"<p:bg>.*?</p:bg>", "", xml, flags=re.S)


#: A fixed timestamp for every entry, so two runs of this file produce the same bytes.
#: Without it ``zipfile`` stamps the current clock and the deck is different every time --
#: which would make `test_the_fixture_is_exactly_what_its_generator_writes` impossible and
#: would put a spurious whole-file diff in front of anyone who regenerates the fixture.
ZIP_TIMESTAMP = (2026, 1, 1, 0, 0, 0)


def _write(out: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(name, date_time=ZIP_TIMESTAMP)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o600 << 16
    out.writestr(info, data)


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
        r'<Override PartName="/ppt/(slides/slide|charts/chart)\d+\.xml"[^>]*/>', "", content_types
    )
    content_types = re.sub(r'<Override PartName="/ppt/embeddings/[^>]*/>', "", content_types)
    # `_with_content_types` in tests/deckbuilder.py and this block both prepend the slash.
    # `PartName="//ppt/..."` opens as [Repaired], and a deck PowerPoint repairs is not a
    # fixture -- so the part paths below must be passed without one.
    additions = "".join(
        f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{SLIDE_TYPE}"/>'
        for n in numbers
    )
    for extension, mime in (("png", "image/png"), ("svg", "image/svg+xml")):
        if not re.search(rf'Extension="{extension}"', content_types, re.IGNORECASE):
            additions += f'<Default Extension="{extension}" ContentType="{mime}"/>'
    content_types = content_types.replace("</Types>", additions + "</Types>")

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
                data = content_types.encode()
            elif re.search(r"slide(Layouts|Masters)/slide\w+\d+\.xml$", name):
                data = blank_template(data.decode()).encode()
            _write(out, name, data)
        for path, build in MEDIA.items():
            _write(out, path, build())
        for index, spec in enumerate(SLIDES):
            _write(out, f"ppt/slides/slide{index + 1}.xml", slide_xml(index, spec).encode())
            _write(
                out,
                f"ppt/slides/_rels/slide{index + 1}.xml.rels",
                slide_rels(index).encode(),
            )


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else TARGET
    write_deck(target)
    for index, spec in enumerate(SLIDES, start=1):
        print(f"{index:3d}  {spec['state']:9s} {spec['key']}")
    print(f"\n{len(SLIDES)} slides -> {target} ({target.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

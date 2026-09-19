#!/usr/bin/env python3
"""Build decks that measure what PowerPoint draws for ``a:pattFill`` and ``a:tile``.

Both fills were guesses.  ``render/fill.py`` hardcoded an 8-unit hatch cell with nothing
in the file saying where 8 came from, implemented 23 of the 54 presets in
``ST_PresetPatternVal`` and fell through to a **solid fill** for the other 31; and it
sized a tile at ``sx`` of the *shape's bounding box*, which is not what ``a:tile@sx``
means.  ``tests/fixtures/feature-sweep.pptx`` slides 13 and 10 pin both, at SSIM 0.0563
and 0.1857.

The instrument here is better than a raster reading, and that is the whole reason these
decks are cheap enough to sweep all 54 presets at once.  **PowerPoint's PDF export writes
a fill of either kind as a PDF tiling pattern** (``/PatternType 1``) whose dictionary
states the cell outright::

    << /Type /Pattern /PatternType 1 /BBox [0 0 64 64] /XStep 64 /YStep 64
       /Matrix [0.125 0 0 0.125 32 349] /Resources ... >>

so the cell is ``XStep * Matrix[0]`` points -- exactly, with no pixel quantisation and no
antialiasing to threshold -- and ``Matrix[4], Matrix[5]`` give the lattice's registration
point in page coordinates.  The pattern's *content* is a single image XObject holding one
cell, which for ``a:pattFill`` is the preset's own 8x8 bitmap upsampled 8x.  Decoding that
image **is** reading PowerPoint's drawing of the preset, rather than inferring it from the
preset's name -- and the name is not reliable: ``dkDnDiag`` is a 2 px wide diagonal at a
4 px pitch, not the pair of hairlines the old code drew.

``fill-patt``
    All 54 presets of ``ST_PresetPatternVal``, one shape each, flat black on white so the
    decoded cell is two-valued.  24 to a slide.
``fill-pitch``
    Whether the cell is a fixed length.  One preset over eight shape sizes from 24x18 to
    640x320 pt, then the same preset at eight shape *origins* chosen off the 8 pt lattice
    (x = 36.5, 41, 100.3 pt ...), which is what separates "registered to the shape" from
    "registered to the page".
``fill-tile``
    ``a:tile``.  Sweeps, one axis at a time: the picture's pixel dimensions (8 to 96 px,
    and one non-square); its ``pHYs`` density (absent, 72, 96, 144, 300 dpi), which is the
    other sense in which a picture has a native size; ``sx``/``sy`` from 25% to 200%
    including ``sx != sy``; the shape's own box at three sizes with everything else held;
    all nine ``@algn`` values; and ``@tx``/``@ty``.

Usage -- the deck's file name picks its probe table::

    python3 tools/make_fill_probe.py ~/pptx2svg-oracle/fill-patt.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/pptx2svg-oracle/fill-patt.pptx ~/pptx2svg-oracle/fill-patt.pdf
    python3 tools/read_fill_probe.py ~/pptx2svg-oracle/fill-patt.pdf

``~/pptx2svg-oracle/`` holds the eleven fixtures and their PDFs and nothing else that
lasts; a probe deck and its export are throwaway and must be deleted again.
"""

from __future__ import annotations

import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from make_feature_sweep import (  # noqa: E402
    IMAGE_REL,
    LAYOUT_REL,
    NAMESPACES,
    PT,
    SLIDE_REL,
    SLIDE_TYPE,
    ZIP_TIMESTAMP,
    blank_template,
    label,
    outline,
    shape,
)

#: The slide is 720 x 405 pt.  Probes are packed into it left to right, top to bottom, and
#: a probe that will not fit starts a new slide -- a shape hanging off the slide still gets
#: a pattern in the export, but a clipped one is a reading nobody should have to trust.
SLIDE_WIDTH, SLIDE_HEIGHT = 720, 405
MARGIN, GAP, CAPTION = 18, 16, 16


def pack(rows: list[dict]) -> list[dict]:
    """Place each probe's box, flowing onto a new slide when the row will not fit."""
    slide, x, y, row_height = 0, MARGIN, 34, 0
    out = []
    for row in rows:
        width, height = row["cx"], row["cy"]
        if "at" in row:
            # A fixed origin: its whole point is to sit where the packer would not put it.
            left, top, fixed_slide = row["at"]
            out.append({**row, "slide": fixed_slide, "x": int(round(left * PT)),
                        "y": int(round(top * PT)), "cx": int(width * PT),
                        "cy": int(height * PT)})
            continue
        if x + width > SLIDE_WIDTH - MARGIN and x > MARGIN:
            x, y, row_height = MARGIN, y + row_height + CAPTION + GAP, 0
        if y + height + CAPTION > SLIDE_HEIGHT - MARGIN and (x > MARGIN or y > 34):
            slide, x, y, row_height = slide + 1, MARGIN, 34, 0
        out.append({**row, "slide": slide, "x": int(x * PT), "y": int(y * PT),
                    "cx": int(width * PT), "cy": int(height * PT)})
        x += width + GAP
        row_height = max(row_height, height)
    return out

ROOT = HERE.parent
SOURCE = ROOT / "tests/fixtures/authoring-integration.pptx"

# --------------------------------------------------------------------------------------
# ST_PresetPatternVal, ECMA-376 Part 1 20.1.10.51 -- all 54 of them, in the schema's order.
# --------------------------------------------------------------------------------------

PRESETS = (
    "pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70",
    "pct75", "pct80", "pct90",
    "horz", "vert", "ltHorz", "ltVert", "dkHorz", "dkVert", "narHorz", "narVert",
    "dashHorz", "dashVert",
    "cross", "dnDiag", "upDiag", "ltDnDiag", "ltUpDiag", "dkDnDiag", "dkUpDiag",
    "wdDnDiag", "wdUpDiag", "dashDnDiag", "dashUpDiag", "diagCross",
    "smGrid", "lgGrid", "dotGrid", "smCheck", "lgCheck",
    "openDmnd", "solidDmnd", "dotDmnd",
    "plaid", "sphere", "weave", "divot", "shingle", "wave", "trellis", "zigZag",
    "smConfetti", "lgConfetti", "horzBrick", "diagBrick",
)


# --------------------------------------------------------------------------------------
# Artwork
# --------------------------------------------------------------------------------------


def png(rows: list[list[tuple[int, int, int, int]]], dpi: int | None = None) -> bytes:
    """A minimal RGBA PNG, optionally carrying a ``pHYs`` density.

    ``pHYs`` is the second sense in which a picture has a native size, and the whole
    question ``fill-tile`` exists to settle is which sense ``a:tile@sx`` scales.
    """
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
    out = b"\x89PNG\r\n\x1a\n" + chunk(
        b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    )
    if dpi is not None:
        # pHYs is pixels per *metre*; 1 unit = 1 metre is the only value PNG defines.
        per_metre = int(round(dpi / 0.0254))
        out += chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    return out + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


def corner_jpeg(width: int, height: int, dpi: int | None = None) -> bytes:
    """The same mark as a JPEG, to say whether the density law is PNG's or the format's.

    A JPEG states its density in a JFIF ``APP0`` rather than a ``pHYs``, and ``units=0``
    there means "aspect ratio only, no physical size" -- which is the JPEG spelling of the
    case a ``pHYs``-less PNG puts.  Needs Pillow; the deck simply omits these rows without.
    """
    from io import BytesIO

    from PIL import Image

    image = Image.new("RGB", (width, height), (0xE2, 0xF5, 0xF3))
    pixels = image.load()
    for y in range(height):
        for x in range(width):
            if x < max(1, width // 3) or y >= height - max(1, height // 3):
                pixels[x, y] = (0x0F, 0x76, 0x6E)
    buffer = BytesIO()
    if dpi is None:
        image.save(buffer, "JPEG", quality=95)
    else:
        image.save(buffer, "JPEG", quality=95, dpi=(dpi, dpi))
    return buffer.getvalue()


def corner_png(width: int, height: int, dpi: int | None = None) -> bytes:
    """An asymmetric two-colour mark: a bar down the left and one across the bottom.

    Asymmetric in both axes so a mirrored copy is unmistakable, and two-valued so the
    export's own upsample of it decodes cleanly.
    """
    thick_x = max(1, width // 3)
    thick_y = max(1, height // 3)
    rows = []
    for y in range(height):
        row = []
        for x in range(width):
            ink = x < thick_x or y >= height - thick_y
            row.append((0x0F, 0x76, 0x6E, 255) if ink else (0xE2, 0xF5, 0xF3, 255))
        rows.append(row)
    return png(rows, dpi)


# --------------------------------------------------------------------------------------
# Probe tables.  Each entry is one shape; the reader matches it by the pattern's own
# registration point, so the shapes are spaced well apart.
# --------------------------------------------------------------------------------------


def patt_shape(x: int, y: int, cx: int, cy: int, preset: str, name: str) -> str:
    fill = (
        f"<a:pattFill prst='{preset}'>"
        "<a:fgClr><a:srgbClr val='000000'/></a:fgClr>"
        "<a:bgClr><a:srgbClr val='FFFFFF'/></a:bgClr></a:pattFill>"
    )
    return shape(x, y, cx, cy, fill=fill, line=outline("94A3B8", 0.5), name=name)


def patt_probes() -> list[dict]:
    """All 54 presets, one shape each."""
    return pack(
        [
            {"kind": "patt", "key": preset, "preset": preset, "cx": 100, "cy": 62}
            for preset in PRESETS
        ]
    )


#: Shape sizes, in points.  If the cell is a fixed length none of these moves it.
PITCH_SIZES = (
    (24, 18), (48, 36), (100, 62), (160, 90),
    (240, 120), (320, 160), (480, 240), (640, 320),
)

#: Shape origins deliberately off the 8 pt lattice.  A cell registered to the *shape*
#: puts the pattern's origin at each of these; one registered to the *page* snaps.
PITCH_ORIGINS = (36.0, 36.5, 38.0, 41.0, 100.3, 173.75, 260.125, 411.0)


def pitch_probes() -> list[dict]:
    sizes = pack(
        [
            {"kind": "patt", "key": f"size-{cx}x{cy}", "preset": "cross", "cx": cx, "cy": cy}
            for cx, cy in PITCH_SIZES
        ]
    )
    after = max(probe["slide"] for probe in sizes) + 1
    origins = []
    for index, left in enumerate(PITCH_ORIGINS):
        slide, cell = divmod(index, 4)
        origins.append(
            {
                "slide": after + slide,
                "kind": "patt",
                "key": f"origin-{left}",
                "preset": "horz",
                "x": int(round(left * PT)),
                "y": int(round((34 + cell * 88) * PT)),
                "cx": int(200 * PT),
                "cy": int(62 * PT),
            }
        )
    return sizes + origins


#: ``(pixels_wide, pixels_high, dpi_or_None)`` -- the picture's two native sizes.
TILE_IMAGES = (
    (8, 8, None), (16, 16, None), (32, 32, None), (64, 64, None), (96, 96, None),
    (48, 24, None),
    (32, 32, 72), (32, 32, 96), (32, 32, 144), (32, 32, 300),
)


def tile_fill(rel: str, sx: int, sy: int, *, flip: str = "none", algn: str = "tl",
              tx: int = 0, ty: int = 0) -> str:
    return (
        f"<a:blipFill><a:blip r:embed='{rel}'/>"
        f"<a:tile tx='{tx}' ty='{ty}' sx='{sx}' sy='{sy}' flip='{flip}' algn='{algn}'/>"
        "</a:blipFill>"
    )


def tile_probes() -> list[dict]:
    """Vary one thing at a time; ``base`` is the row every sweep returns to."""
    base = {"image": (32, 32, None), "sx": 100000, "sy": 100000, "flip": "none",
            "algn": "tl", "tx": 0, "ty": 0, "cx": 136, "cy": 96}
    rows: list[dict] = []

    for pixels in TILE_IMAGES:
        rows.append({**base, "image": pixels,
                     "key": f"px-{pixels[0]}x{pixels[1]}-dpi{pixels[2] or 'none'}"})
    for sx, sy in ((25000, 25000), (50000, 50000), (60000, 60000), (150000, 150000),
                   (200000, 200000), (50000, 150000), (150000, 50000)):
        rows.append({**base, "sx": sx, "sy": sy, "key": f"s-{sx}x{sy}"})
    for cx, cy in ((68, 48), (136, 96), (272, 192), (400, 96)):
        rows.append({**base, "cx": cx, "cy": cy, "key": f"box-{cx}x{cy}"})
    for algn in ("tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"):
        rows.append({**base, "algn": algn, "sx": 50000, "sy": 50000,
                     "key": f"algn-{algn}"})
    for tx, ty in ((0, 0), (4 * PT, 0), (0, 4 * PT), (-4 * PT, 6 * PT), (12 * PT, 12 * PT)):
        rows.append({**base, "tx": tx, "ty": ty, "sx": 50000, "sy": 50000,
                     "key": f"t-{tx}-{ty}"})
    for flip in ("none", "x", "y", "xy"):
        rows.append({**base, "flip": flip, "sx": 50000, "sy": 50000,
                     "key": f"flip-{flip}"})

    return pack([{**row, "kind": "tile"} for row in rows])


ALIGNMENTS = ("tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br")


def algn_probes() -> list[dict]:
    """``@algn`` alone, on boxes that are **not** whole multiples of the tile.

    The first ``fill-tile`` sweep put all nine alignments on a 136 x 96 pt box with an
    8 x 8 pt tile.  136 and 96 are 17 and 12 whole tiles, so top-, centre- and
    bottom-registration all land on the same lattice and the sweep said nothing.  Every
    box here is deliberately indivisible by its tile in both axes, and two tile sizes are
    used so that an accidental divisor cannot hide a second time.
    """
    rows = []
    for sx, box in ((50000, (130, 90)), (75000, (137, 83))):
        for algn in ALIGNMENTS:
            rows.append(
                {
                    "kind": "tile",
                    "key": f"algn-{algn}-{box[0]}x{box[1]}-s{sx // 1000}",
                    "image": (32, 32, None),
                    "sx": sx,
                    "sy": sx,
                    "flip": "none",
                    "algn": algn,
                    "tx": 0,
                    "ty": 0,
                    "cx": box[0],
                    "cy": box[1],
                }
            )
    return pack(rows)


def jpeg_probes() -> list[dict]:
    """Is the density law the *format's* or the picture's?

    ``fill-tile`` measured it on PNGs only, and a PNG with no ``pHYs`` came out at
    144 dpi.  A JPEG says its density in a JFIF ``APP0`` instead, and a JPEG written with
    no density at all still carries ``units=0`` there -- so "no physical size" is spelled
    differently and could well be answered differently.
    """
    rows = []
    for dpi in (None, 72, 96, 144, 300):
        for pixels in ((32, 32), (64, 64)):
            rows.append(
                {
                    "kind": "tile",
                    "format": "jpeg",
                    "key": f"jpeg-{pixels[0]}px-dpi{dpi or 'none'}",
                    "image": (pixels[0], pixels[1], dpi),
                    "sx": 100000, "sy": 100000, "flip": "none", "algn": "tl",
                    "tx": 0, "ty": 0, "cx": 136, "cy": 96,
                }
            )
    return pack(rows)


DECKS = {
    "fill-patt": patt_probes,
    "fill-pitch": pitch_probes,
    "fill-algn": algn_probes,
    "fill-jpeg": jpeg_probes,
    "fill-tile": tile_probes,
}


def probes_for(name: str) -> list[dict]:
    for key, build in DECKS.items():
        if key in name:
            return build()
    raise SystemExit(f"no probe table matches {name!r}; expected one of {sorted(DECKS)}")


# --------------------------------------------------------------------------------------
# Deck assembly
# --------------------------------------------------------------------------------------


def image_name(probe: dict) -> str:
    width, height, dpi = probe["image"]
    return f"probe-{width}x{height}x{dpi or 0}.{probe.get('format', 'png')}"


def image_rel(probe: dict) -> str:
    return "rId" + image_name(probe).replace("-", "").replace(".", "")


def media_for(probes: list[dict]) -> dict[str, bytes]:
    media = {}
    for probe in probes:
        if probe["kind"] != "tile":
            continue
        width, height, dpi = probe["image"]
        build = corner_jpeg if probe.get("format") == "jpeg" else corner_png
        media[f"ppt/media/{image_name(probe)}"] = build(width, height, dpi)
    return media


def relationships_for(probes: list[dict]) -> list[tuple[str, str, str]]:
    seen, rels = set(), []
    for probe in probes:
        if probe["kind"] != "tile":
            continue
        rel = image_rel(probe)
        if rel in seen:
            continue
        seen.add(rel)
        rels.append((rel, IMAGE_REL, f"../media/{image_name(probe)}"))
    return rels


def shape_for(probe: dict) -> str:
    if probe["kind"] == "patt":
        body = patt_shape(
            probe["x"], probe["y"], probe["cx"], probe["cy"], probe["preset"], probe["key"]
        )
    else:
        fill = tile_fill(
            image_rel(probe), probe["sx"], probe["sy"],
            flip=probe["flip"], algn=probe["algn"], tx=probe["tx"], ty=probe["ty"],
        )
        body = shape(probe["x"], probe["y"], probe["cx"], probe["cy"], fill=fill,
                     line=outline("94A3B8", 0.5), name=probe["key"])
    return body + label(probe["key"], probe["x"], probe["y"] + probe["cy"] + 4 * PT,
                        max(probe["cx"], 120 * PT), size=700)


def slide_xml(probes: list[dict], index: int) -> str:
    shapes = "".join(shape_for(probe) for probe in probes if probe["slide"] == index)
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="0" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + shapes
        + "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    )


def slide_rels(rels: list[tuple[str, str, str]]) -> str:
    extra = "".join(
        f"<Relationship Id='{rel_id}' Type='{rel_type}' Target='{target}'/>"
        for rel_id, rel_type, target in rels
    )
    return (
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
        f"<Relationship Id='rId1' Type='{LAYOUT_REL}' Target='../slideLayouts/slideLayout1.xml'/>"
        f"{extra}</Relationships>"
    )


def _write(out: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(name, date_time=ZIP_TIMESTAMP)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o600 << 16
    out.writestr(info, data)


def write_deck(target: Path, probes: list[dict]) -> int:
    count = max(probe["slide"] for probe in probes) + 1
    numbers = range(1, count + 1)
    rels = relationships_for(probes)

    source = zipfile.ZipFile(SOURCE)
    presentation = source.read("ppt/presentation.xml").decode()
    pres_rels = source.read("ppt/_rels/presentation.xml.rels").decode()
    content_types = source.read("[Content_Types].xml").decode()

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
    additions = "".join(
        f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{SLIDE_TYPE}"/>'
        for n in numbers
    )
    for extension, mime in (("png", "image/png"), ("jpeg", "image/jpeg")):
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
        for path, payload in media_for(probes).items():
            _write(out, path, payload)
        for index in range(count):
            _write(out, f"ppt/slides/slide{index + 1}.xml", slide_xml(probes, index).encode())
            _write(out, f"ppt/slides/_rels/slide{index + 1}.xml.rels", slide_rels(rels).encode())
    return count


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    probes = probes_for(target.name)
    count = write_deck(target, probes)
    print(f"{len(probes)} probes over {count} slides -> {target} "
          f"({target.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

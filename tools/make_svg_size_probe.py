#!/usr/bin/env python3
"""Build a deck that measures what size PowerPoint draws an SVG picture at.

An SVG picture (``asvg:svgBlip``) states its own size on its root -- ``width="72pt"``,
``height="50%"``, ``6em``, nothing -- and PowerPoint draws the SVG at that size and then
stretches the result to the picture's frame.  So the size is visible in the frame: each
probe draws a blue square of known side in user space over a red ground, in a frame of
1.4 x 0.7 inches, and the fraction of the frame the square covers says what size the root
was taken to be.  ``tools/read_svg_size_probe.py`` reads the export and compares it with
:func:`ooxml_common.svgimage.natural_size`.

* ``vb_*`` have ``viewBox="0 0 20 20"`` and a 10-unit square at the origin: the size only
  matters through its proportions, which with the default ``preserveAspectRatio``
  letterbox a square ``viewBox`` into a 2:1 root.
* ``nvb_*`` have no ``viewBox``: the square is drawn at its own size, in a viewport of the
  root's, and its side is half the size predicted, so the reading is most precise where
  the prediction is right.

Every picture's ``a:blip`` names a one-pixel **green** PNG as the fallback, so a probe
PowerPoint would not draw from its SVG shows up as green rather than as a wrong size.

Usage::

    python3 tools/make_svg_size_probe.py out/svg-size.pptx

Then export it to PDF with ``tools/powerpoint_export_pdf.applescript``.
"""

from __future__ import annotations

import io
import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

from deckbuilder import derive_deck  # noqa: E402

from ooxml_common.svgimage import natural_size  # noqa: E402

SVG = 'xmlns="http://www.w3.org/2000/svg"'
EMU_PER_INCH = 914400
FRAME = (1.4, 0.7)  # inches
COLUMNS = 6

#: (name, the root's size attributes, whether it has the 20x20 viewBox)
CASES: list[tuple[str, str, bool]] = [
    ("vb_pt", 'width="72pt" height="72pt"', True),
    ("vb_pc", 'width="6pc" height="6pc"', True),
    ("vb_in", 'width="1in" height="1in"', True),
    ("vb_cm", 'width="2.54cm" height="2.54cm"', True),
    ("vb_mm", 'width="25.4mm" height="25.4mm"', True),
    ("vb_em", 'width="6em" height="6em"', True),
    ("vb_pct", 'width="100%" height="100%"', True),
    ("vb_pct_mixed", 'width="100%" height="50%"', True),
    ("vb_none", "", True),
    ("vb_width_only", 'width="2in"', True),
    ("vb_2to1", 'width="2in" height="1in"', True),
    ("vb_em_2to1", 'width="12em" height="6em"', True),
    ("vb_upper", 'width="2IN" height="1IN"', True),
    ("vb_exponent", 'width="1.44e2pt" height="0.72e2pt"', True),
    ("nvb_px", 'width="96" height="96"', False),
    ("nvb_pt", 'width="72pt" height="72pt"', False),
    ("nvb_in", 'width="1in" height="1in"', False),
    ("nvb_cm", 'width="2.54cm" height="2.54cm"', False),
    ("nvb_mm", 'width="25.4mm" height="25.4mm"', False),
    ("nvb_em", 'width="6em" height="6em"', False),
    ("nvb_ex", 'width="12ex" height="12ex"', False),
    ("nvb_em_font_size", 'width="6em" height="6em" font-size="20"', False),
    ("nvb_em_style", 'width="6em" height="6em" style="font-size:20px"', False),
    ("nvb_em_font_size_pt", 'width="6em" height="6em" font-size="12pt"', False),
    ("nvb_pct", 'width="50%" height="50%"', False),
    ("nvb_none", "", False),
    ("nvb_width_only", 'width="1in"', False),
    ("nvb_upper", 'width="6PC" height="6PC"', False),
    ("nvb_spaced", 'width=" 72pt " height="72pt"', False),
    ("nvb_exponent", 'width="0.72e2pt" height="0.72e2pt"', False),
]


def attributes(name: str) -> dict[str, str]:
    """Probe ``name``'s root attributes that bear on its size."""
    size, has_box = next((size, box) for case, size, box in CASES if case == name)
    found = dict(re.findall(r'([\w-]+)="([^"]*)"', size))
    return {**found, "viewBox": "0 0 20 20"} if has_box else found


def blue_side(name: str) -> float:
    """The blue square's side in user units: 10 of the 20x20 viewBox, or half the
    predicted size without one."""
    size, has_box = next((size, box) for case, size, box in CASES if case == name)
    if has_box:
        return 10.0
    predicted = natural_size(attributes(name))
    assert predicted is not None, name
    return round(min(predicted) / 2, 3)


def svg(name: str) -> bytes:
    size, has_box = next((size, box) for case, size, box in CASES if case == name)
    side = blue_side(name)
    box = ' viewBox="0 0 20 20"' if has_box else ""
    ground = 20 if has_box else 10000
    return (f'<svg {SVG} {size}{box}>'
            f'<rect width="{ground}" height="{ground}" fill="#ff0000"/>'
            f'<rect width="{side:g}" height="{side:g}" fill="#0000ff"/></svg>').encode()


def frame(index: int) -> tuple[int, int, int, int]:
    """The frame of probe ``index``: x, y, cx, cy in EMU."""
    column, row = index % COLUMNS, index // COLUMNS
    x, y = 0.2 + 1.6 * column, 0.2 + 1.05 * row
    return (round(x * EMU_PER_INCH), round(y * EMU_PER_INCH),
            round(FRAME[0] * EMU_PER_INCH), round(FRAME[1] * EMU_PER_INCH))


def green_png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data)))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\x00\xc8\x00")) + chunk(b"IEND", b""))


def build() -> bytes:
    image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    parts = {"ppt/media/svgsize-fallback.png": green_png()}
    relationships = [("rIdSvgSizeFallback", image, "../media/svgsize-fallback.png")]
    shapes = []
    for index, (name, _size, _box) in enumerate(CASES):
        parts[f"ppt/media/svgsize{index}.svg"] = svg(name)
        relationships.append((f"rIdSvgSize{index}", image, f"../media/svgsize{index}.svg"))
        x, y, cx, cy = frame(index)
        shapes.append(
            '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<p:nvPicPr><p:cNvPr id="{900 + index}" name="{name}"/><p:cNvPicPr/><p:nvPr/>'
            '</p:nvPicPr><p:blipFill><a:blip r:embed="rIdSvgSizeFallback"><a:extLst>'
            '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
            '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" '
            f'r:embed="rIdSvgSize{index}"/></a:ext></a:extLst></a:blip>'
            '<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>'
            f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    # A 10 x 5.625 inch deck; what its first slide draws is dropped before the probes go on.
    base = zipfile.ZipFile(ROOT / "tests/fixtures/authoring-integration.pptx")
    stripped = io.BytesIO()
    with zipfile.ZipFile(stripped, "w", zipfile.ZIP_DEFLATED) as target:
        for part in base.namelist():
            data = base.read(part)
            if part == "ppt/slides/slide1.xml":
                for tag in (b"sp", b"pic", b"graphicFrame", b"cxnSp"):
                    data = re.sub(rb"<p:" + tag + rb"\b.*?</p:" + tag + rb">", b"", data,
                                  flags=re.S)
            target.writestr(part, data)
    # derive_deck opens its source with zipfile, which reads a file object as well as a path.
    return derive_deck(stripped, parts=parts, shapes_xml="".join(shapes),  # type: ignore[arg-type]
                       slide_relationships=relationships, defaults={"svg": "image/svg+xml"})


def main(argv: list[str]) -> int:
    if len(argv) != 1:
        print(__doc__.split("Usage::")[1].strip().splitlines()[0], file=sys.stderr)
        return 2
    path = Path(argv[0])
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(build())
    print(f"wrote {path}: {len(CASES)} probes on slide 1")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

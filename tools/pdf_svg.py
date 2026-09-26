#!/usr/bin/env python3
"""PowerPoint's PDF pages as SVG (PyMuPDF), so both sides of the fidelity score go through resvg.

A copy of the sibling ``docx2svg``'s ``tools/pdf_svg.py`` (its commit ``02822f8``; its
ROADMAP.md 5.13 has the measurements), as ``tools/fidelity.py`` is itself a copy per
harness: the converter is instrument, not runtime, so it is not shared code.  What
differs here is only what the two harnesses differ in -- the oracle directory is an
argument (``--oracle``, default ``~/pptx2svg-oracle``), the pages are slides scored at
``fidelity.WIDTH`` pixels wide rather than at 300 dpi, and the PDFs validated are every
export in the oracle directory.

**Why.**  ``tools/fidelity.py`` used to rasterise PowerPoint's PDF with pdfium and our SVG
with resvg, so part of every score was pdfium against resvg.  pdfium is the odd engine
out: it grid-fits glyph outlines (a stem or a bar moves up to a pixel) and widens an
axis-aligned fill to whole pixels, where resvg and MuPDF draw the outline as given.
Measured by the sibling on Word's own page: pdfium against MuPDF on the *same PDF*
scores SSIM 0.890, and MuPDF against resvg of this module's SVG 0.996.  Converting
PowerPoint's page to SVG and rasterising both sides with resvg leaves no engine
difference to score.

**The converter is part of the instrument**, so :func:`validate` holds it to the PDF
before its output is trusted, page by page:

* **same engine, two routes** -- MuPDF's raster of the PDF against MuPDF's raster of the
  SVG: whatever differs is the conversion's, since the engine is the same;
* **no pixel beyond anti-aliasing** -- :func:`beyond_antialiasing` counts the pixels of
  one raster outside the 3x3 neighbourhood range of the other (plus a tolerance): a
  moved glyph, a missing element or a changed colour lands there, an edge rendered a
  fraction of a pixel differently does not;
* **no colour changed** -- :func:`flat_colour_difference`, where both rasters are flat;
* **the glyphs are the embedded outlines** -- every font is embedded, every glyph the SVG
  defines is addressed by the embedded *subset's* glyph id, and is redrawn from that
  program, unhinted (:func:`unhinted_outlines`: MuPDF writes outlines with the font's
  hinting applied, which moves PowerPoint's Aptos 30-35 units); and the SVG draws with no
  text at all, which :func:`rasterise` proves by giving resvg no font to find.

**What PowerPoint's pages needed beyond the sibling's.**  Validating every export here
found three more places where PyMuPDF's SVG is not the PDF, each corrected or accounted
for and each measured:

* **stroked text** -- PowerPoint's synthetic bold (fill and stroke) came out with a
  stroke four times too wide (:func:`stroked_text_widths`);
* **ICC-tagged images** -- a photograph in Adobe RGB drawn as though sRGB, 9-11 levels
  off (:func:`srgb_images`);
* **MuPDF's own SVG reader does not tile a ``<pattern>``**, so on those pages the
  same-engine route is skipped as it is for ``<mask>``, and resvg's route decides.

And the gate at the scoring resolution is taken on rasters drawn at 4x and averaged down
(at 96-128 dpi MuPDF's glyph cache and its bitmap filter leave 28-101 isolated pixels a
page at 1x that are not the conversion's); at 300 dpi, the sibling's 1x gate, or 2x where
stroked text fails it (:func:`validate`).  A page that still fails is **not scored**:
``tools/fidelity.py`` reads each page's verdict (:func:`verdicts`) and reports it instead.

**What no rasteriser can un-rasterise.**  PowerPoint draws some things into its PDF as
bitmaps: a 3-D chart's scene (a 300 dpi image under a soft mask), and a shape or picture
with an effect it does not export as vectors.  The conversion carries those bitmaps over
as they are, and :func:`bitmap_regions` lists them per page, so that a region where our
vector drawing meets PowerPoint's raster is known for what it is.

**Font-derived artefacts.**  A converted page carries the glyph outlines of the fonts
PowerPoint embedded -- Microsoft's, subset.  They are treated like font files: never
written into a repository (:func:`_refuse_repository`), cached only next to the oracle's
PDFs (``ORACLE/svg/``, outside the tree), and never published.

**Development tooling only**, behind the ``fidelity`` extra (PyMuPDF, AGPL-3.0 -- fine
for a local tool that is not distributed with the library; nothing under ``src/``
imports it), installed into the project's own ``.venv`` (README.md, "Checking fidelity
against PowerPoint").  Where PyMuPDF is missing, :func:`available` is false and the tests
that need it skip.

Usage::

    python tools/pdf_svg.py --validate                    # every export in the oracle directory
    python tools/pdf_svg.py --validate PDF...             # these PDFs
    python tools/pdf_svg.py PDF...                        # convert (cached); print the cache paths
    python tools/pdf_svg.py --bitmaps                     # list the bitmaps PowerPoint embedded
"""

from __future__ import annotations

import argparse
import hashlib
import io
import os
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

#: Where PowerPoint's exports live (``tools/fidelity.py --oracle``'s default).
DEFAULT_ORACLE = Path(os.path.expanduser("~/pptx2svg-oracle"))
#: The sibling's validation resolution (Word's export grid).  The scoring resolution here
#: is lower -- ``fidelity.WIDTH`` px across a slide, 96 dpi on a 16:9 deck -- and the
#: converter is validated at both (:func:`validate`).
DPI = 300
#: A pixel differs beyond anti-aliasing when a channel is further than this outside the
#: range of the other raster's 3x3 neighbourhood.  An edge pixel two rasterisers cover
#: differently stays inside its neighbours' range; a glyph moved by more than a pixel, a
#: missing stroke or a changed colour does not.
BEYOND_TOLERANCE = 48
#: The largest difference of flat colour (:func:`flat_colour_difference`) a faithful
#: conversion may show: a level of rounding in a colour conversion either way.
MAX_FLAT_COLOUR = 2
#: Isolated pixels beyond anti-aliasing allowed on a page.  MuPDF draws a PDF's text
#: through a glyph cache that places glyphs on a sub-pixel grid, and the SVG's glyphs as
#: paths at their exact position, so a thin stem can land a third of a pixel from the
#: other route's, even under one engine: measured by the sibling, at most 18 pixels of a
#: page.  A moved glyph, a missing element or a changed colour is thousands.
MAX_BEYOND = 20
#: The conversion's own version, part of the cache key: 2 redraws the glyphs unhinted (the
#: sibling's); 3 also corrects stroked text's width and converts ICC-tagged images to sRGB
#: (:func:`stroked_text_widths`, :func:`srgb_images`: PowerPoint's pages need both).
FORMAT = 3
#: The supersampling factor of the scoring-resolution gate (:func:`validate`).
SUPERSAMPLE = 4
#: The least SSIM a supersampled gate (:func:`faithful`) may show: two rasters of one
#: drawing, drawn at 4x and averaged down, score 0.985-0.9999 on every page that passes;
#: a page whose tiles land in another phase (feature-sweep's tile fill, 0.89) stays under
#: the pixel count and is not faithful.
MIN_GATE_SSIM = 0.98


def available() -> bool:
    """Whether PyMuPDF (the converter) imports."""
    try:
        import pymupdf  # noqa: F401
    except Exception:
        return False
    return True


def converter_version() -> str:
    import pymupdf

    return f"pymupdf-{pymupdf.VersionBind}-f{FORMAT}"


def cache_dir(pdf: Path, oracle: Path | None = None) -> Path:
    """Where ``pdf``'s converted pages are cached: under the oracle directory (outside any
    repository), keyed by the PDF's content and the converter's version."""
    oracle = Path(os.path.expanduser(str(oracle))) if oracle is not None else DEFAULT_ORACLE
    digest = hashlib.sha256(Path(pdf).read_bytes()).hexdigest()[:16]
    return oracle / "svg" / f"{Path(pdf).stem}-{digest}-{converter_version()}"


def _refuse_repository(path: Path) -> None:
    """Raise unless ``path`` is outside this repository and outside any git checkout.

    A converted page is font-derived (Microsoft's glyph outlines), so it may not be
    written where it could be committed -- this checkout, a worktree of it, or any other
    repository an oracle directory might have been pointed into."""
    resolved = Path(path).resolve()
    repository = HERE.parent.resolve()
    if resolved.is_relative_to(repository):
        raise RuntimeError(f"refusing to write a converted page (font-derived) inside the repository: {path}")
    for parent in (resolved, *resolved.parents):
        if (parent / ".git").exists():
            raise RuntimeError(f"refusing to write a converted page (font-derived) inside a git checkout ({parent}): "
                               f"{path}")


def convert(pdf: Path, cache: bool = True, oracle: Path | None = None) -> list[tuple[str, dict]]:
    """Every page of ``pdf`` as SVG -- glyphs as paths, redrawn from the embedded programs
    (:func:`unhinted_outlines`) -- with that function's report.  Cached under ``oracle``
    (outside the repository) when ``cache``."""
    import json

    import pymupdf

    pdf = Path(pdf)
    directory = cache_dir(pdf, oracle) if cache else None
    if directory is not None and (directory / "done").exists():
        count = int((directory / "done").read_text())
        return [((directory / f"p{i + 1}.svg").read_text(), json.loads((directory / f"p{i + 1}.json").read_text()))
                for i in range(count)]
    document = pymupdf.open(str(pdf))
    out = []
    for index, page in enumerate(document):
        svg, report = unhinted_outlines(document, index, page.get_svg_image(text_as_path=True))
        svg, report["stroked_text"] = stroked_text_widths(page, svg)
        svg, report["icc_images"] = srgb_images(document, index, svg)
        report["ok"] = report["ok"] and not report["stroked_text"]["unmatched"] and not any(
            image.get("error") for image in report["icc_images"])
        out.append((svg, report))
    if directory is not None:
        _refuse_repository(directory)
        directory.mkdir(parents=True, exist_ok=True)
        for index, (svg, report) in enumerate(out):
            (directory / f"p{index + 1}.svg").write_text(svg)
            (directory / f"p{index + 1}.json").write_text(json.dumps(report))
        (directory / "done").write_text(str(len(out)))
    return out


def page_svgs(pdf: Path, cache: bool = True, oracle: Path | None = None) -> list[str]:
    """:func:`convert`'s pages."""
    return [svg for svg, _report in convert(pdf, cache, oracle)]


_ROOT = re.compile(r'<svg\b[^>]*?\swidth="([\d.]+)"\s+height="([\d.]+)"\s+viewBox="0 0 ([\d.]+) ([\d.]+)"')


def page_size(svg: str) -> tuple[float, float]:
    """A converted page's size in points, from its viewBox."""
    found = _ROOT.search(svg)
    if found is None:
        raise ValueError("a converted page's root has no width, height and viewBox from the origin")
    return float(found.group(3)), float(found.group(4))


def dpi_for_width(svg: str, width: int) -> float:
    """The resolution at which a converted page is ``width`` device pixels across -- the
    scale ``tools/fidelity.py`` renders both sides at."""
    return width * 72 / page_size(svg)[0]


def device_grid(svg: str, dpi: float = DPI) -> str:
    """``svg`` with its root sized in whole device pixels and its viewBox widened to
    match, so that one point is exactly ``dpi / 72`` px.

    PyMuPDF writes the page's size in points, and resvg rounds a root's size to whole
    units *before* zooming: the sibling's A4 came out 595 by 842, drawn 2,479 px wide
    and stretched 1e-4 vertically -- 0.3 px at the foot of the page.  Sizing the root in
    device pixels (the page's size rounded, as pdfium and our SVG round it) and widening
    the viewBox by the fraction of a pixel that adds leaves the drawing unscaled."""
    found = _ROOT.search(svg)
    if found is None:
        raise ValueError("a converted page's root has no width, height and viewBox from the origin")
    width_pt, height_pt = float(found.group(3)), float(found.group(4))
    width_px, height_px = round(width_pt * dpi / 72), round(height_pt * dpi / 72)
    root = found.group(0)
    sized = (root.replace(f'width="{found.group(1)}"', f'width="{width_px}"', 1)
             .replace(f'height="{found.group(2)}"', f'height="{height_px}"', 1)
             .replace(f'viewBox="0 0 {found.group(3)} {found.group(4)}"',
                      f'viewBox="0 0 {width_px * 72 / dpi!r} {height_px * 72 / dpi!r}"', 1))
    return svg.replace(root, sized, 1)


def rasterise(svg: str, dpi: float = DPI, scale: int = 1):
    """resvg's raster of a converted page as an RGB array, on white, with **no fonts at all**:
    a converted page draws its glyphs as paths, and a ``<text>`` would draw nothing.  The
    page is drawn on the device grid (:func:`device_grid`), then zoomed by ``scale``."""
    import numpy as np
    import resvg_py
    from PIL import Image

    Image.MAX_IMAGE_PIXELS = None
    png = resvg_py.svg_to_bytes(svg_string=device_grid(svg, dpi), zoom=float(scale), background="white",
                                skip_system_fonts=True)
    return np.asarray(Image.open(io.BytesIO(bytes(png))).convert("RGB"))


def truth_pages(pdf: Path, width: int, oracle: Path | None = None) -> list:
    """PowerPoint's pages ``width`` px across, through the converter and resvg."""
    return [rasterise(svg, dpi_for_width(svg, width)) for svg in page_svgs(pdf, oracle=oracle)]


# -- validation -------------------------------------------------------------------------


def beyond_antialiasing(a, b, tolerance: int = BEYOND_TOLERANCE):
    """The mask of pixels where either raster leaves the 3x3 neighbourhood range of the
    other by more than ``tolerance`` in some channel."""
    import numpy as np

    def neighbourhood(image):
        padded = np.pad(image, ((1, 1), (1, 1), (0, 0)), mode="edge")
        rows, cols = image.shape[:2]
        low, high = image.copy(), image.copy()
        for dy in range(3):
            for dx in range(3):
                window = padded[dy:dy + rows, dx:dx + cols]
                np.minimum(low, window, out=low)
                np.maximum(high, window, out=high)
        return low, high

    bad = np.zeros(a.shape[:2], bool)
    for x, y in ((a, b), (b, a)):
        x, y = x.astype(np.int16), y.astype(np.int16)
        low, high = neighbourhood(x)
        bad |= ((y < low - tolerance) | (y > high + tolerance)).any(axis=2)
    return bad


def _fit(a, b):
    rows, cols = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
    return a[:rows, :cols], b[:rows, :cols]


def flat_colour_difference(a, b) -> int:
    """The largest channel difference between two rasters where both are flat -- a pixel
    whose 3x3 neighbourhood is one colour in each: a fill, a line's core, a glyph's
    interior.  A changed colour shows here, whatever anti-aliasing does at edges.  (The
    fidelity score's 64-bin histogram cannot say it: a fill of 215 against 216 crosses a
    bin boundary, and ``table-test`` scored 0.85 on one level of one channel.)"""
    import numpy as np

    def flat(image):
        image = image.astype(np.int16)
        padded = np.pad(image, ((1, 1), (1, 1), (0, 0)), mode="edge")
        rows, cols = image.shape[:2]
        same = np.ones((rows, cols), bool)
        for dy in range(3):
            for dx in range(3):
                same &= (padded[dy:dy + rows, dx:dx + cols] == image).all(axis=2)
        return same

    both = flat(a) & flat(b)
    if not both.any():
        return 0
    return int(np.abs(a.astype(np.int16) - b.astype(np.int16))[both].max())


def compare(a, b, images=(), dpi: float = DPI) -> dict:
    """SSIM and histogram (fidelity.score), mean and max |difference|, the pixels beyond
    anti-aliasing and the largest difference of flat colour, of two rasters of one page
    at ``dpi``.  Inside ``images`` (rectangles in pt, :func:`bitmap_regions`) the pixels
    beyond anti-aliasing are counted apart, and ``flat_colour`` leaves them out
    (``flat_colour_all`` does not): see :func:`faithful`."""
    import numpy as np

    import fidelity

    a, b = _fit(a, b)
    row = fidelity.score(b, a)
    difference = np.abs(a.astype(np.int16) - b.astype(np.int16)).max(axis=2)
    beyond = beyond_antialiasing(a, b)
    inside = np.zeros(beyond.shape, bool)
    k = dpi / 72
    for image in images:
        x0, y0, x1, y1 = image["bbox"]
        inside[max(int(y0 * k) - 2, 0):int(y1 * k) + 3, max(int(x0 * k) - 2, 0):int(x1 * k) + 3] = True
    outside_a, outside_b = a.copy(), b.copy()
    outside_b[inside] = outside_a[inside]
    return {"ssim": row["ssim"], "histogram": row["histogram"], "mean_abs": round(float(difference.mean()), 3),
            "max_abs": int(difference.max()), "beyond": int((beyond & ~inside).sum()),
            "beyond_in_images": int((beyond & inside).sum()),
            "flat_colour": flat_colour_difference(outside_a, outside_b), "flat_colour_all": flat_colour_difference(a, b)}


def _mupdf(document, index: int, dpi: float = DPI):
    import numpy as np

    pixmap = document[index].get_pixmap(matrix=_matrix(dpi), alpha=False)
    return np.frombuffer(pixmap.samples, np.uint8).reshape(pixmap.height, pixmap.width, 3)


def _matrix(dpi: float):
    import pymupdf

    return pymupdf.Matrix(dpi / 72, dpi / 72)


def _pdfium(pdf: Path, index: int, dpi: float = DPI):
    import numpy as np
    import pypdfium2 as pdfium

    return np.asarray(pdfium.PdfDocument(str(pdf))[index].render(scale=dpi / 72).to_pil().convert("RGB"))


def _downsample(image, factor: int):
    import numpy as np

    rows, cols = image.shape[0] // factor * factor, image.shape[1] // factor * factor
    out = np.empty((rows // factor, cols // factor, 3), np.uint8)
    strip = 256  # output rows at a time: the whole page in float would be gigabytes
    for top in range(0, rows // factor, strip):
        bottom = min(top + strip, rows // factor)
        blocks = image[top * factor:bottom * factor, :cols].reshape(bottom - top, factor, cols // factor, factor, 3)
        out[top:bottom] = blocks.astype(np.float32).mean(axis=(1, 3)).round().astype(np.uint8)
    return out


_GLYPH = re.compile(r'<path id="font_(\d+)_(\d+)" d="([^"]*)"')
_USE = re.compile(r'<use data-text="[^"]*" xlink:href="#font_(\d+)_(\d+)" transform="matrix\(([^)]*)\)"')


def _programs(document, index: int) -> tuple[dict, list[str]]:
    """The embedded font programs of page ``index`` by ``/BaseFont`` (``None`` where
    fontTools cannot read one), and the fonts not embedded."""
    from fontTools.ttLib import TTFont

    programs, not_embedded = {}, []
    for font in document[index].get_fonts(full=True):
        buffer = document.extract_font(font[0])[3]
        if not buffer:
            not_embedded.append(font[3])
            continue
        try:
            programs[font[3]] = TTFont(io.BytesIO(buffer), fontNumber=0)
        except Exception:
            programs[font[3]] = None
    return programs, sorted(set(not_embedded))


def _traced_names(page) -> dict:
    """MuPDF's text trace: the font name of every glyph drawn, by glyph id and origin."""
    out = {}
    for span in page.get_texttrace():
        for _unicode, gid, origin, _bbox in span["chars"]:
            out[(gid, round(origin[0], 1), round(origin[1], 1))] = span["font"]
    return out


def _glyph_bounds(font, gid: int):
    from fontTools.pens.boundsPen import BoundsPen

    glyph_set = font.getGlyphSet()
    pen = BoundsPen(glyph_set)
    glyph_set[font.getGlyphOrder()[gid]].draw(pen)
    return pen.bounds


def _glyph_path(font, gid: int) -> str:
    """The glyph's outline as SVG path data, in ems, y up (as MuPDF writes its glyphs)."""
    from fontTools.pens.svgPathPen import SVGPathPen
    from fontTools.pens.transformPen import TransformPen

    glyph_set = font.getGlyphSet()
    scale = 1 / font["head"].unitsPerEm
    pen = SVGPathPen(glyph_set, ntos=lambda v: f"{v:.9g}")
    glyph_set[font.getGlyphOrder()[gid]].draw(TransformPen(pen, (scale, 0, 0, scale, 0, 0)))
    return pen.getCommands()


def unhinted_outlines(document, index: int, svg: str) -> tuple[str, dict]:
    """``svg`` (MuPDF's page ``index``) with every glyph redrawn from the embedded font
    program it was drawn with, unhinted; and the report :func:`validate` holds.

    **Why redraw.**  MuPDF's SVG writes a glyph's outline as FreeType loads it *with the
    font's hinting* (at one pixel a font unit): PowerPoint's Aptos subsets come out with
    their x-height pulled up 30 units (1.5% of the em, 2.7 px at 44 pt on 300 dpi) --
    measured by the sibling on ``table-test.pdf``.  Both PDF rasterisers draw the outline
    unhinted at that scale, so the converter's own reading of a glyph is replaced by the
    program's.

    **Which program.**  MuPDF names each SVG font ``font_<n>`` and each glyph by its id in
    the *embedded subset* (an id only that subset gives meaning to -- a font lookup would
    carry the installed face's ids).  Font ``n`` is the embedded program named by MuPDF's
    text trace for its glyphs' origins, and among subsets of one face the one whose
    outlines are nearest MuPDF's, in font units.  The report says how far MuPDF's hinted
    outline was from the one drawn (``hinting_units``), and fails a font no embedded
    program accounts for."""
    from fontTools.pens.boundsPen import BoundsPen
    from fontTools.svgLib.path import parse_path

    programs, not_embedded = _programs(document, index)
    traced = _traced_names(document[index])
    names: dict[str, set] = {}
    for n, gid, matrix in _USE.findall(svg):
        values = [float(v) for v in matrix.split(",")]
        name = traced.get((int(gid), round(values[4], 1), round(values[5], 1)))
        if name is not None:
            names.setdefault(n, set()).add(name)
    glyphs: dict[str, list] = {}
    for n, gid, d in _GLYPH.findall(svg):
        pen = BoundsPen(None)
        parse_path(d, pen)
        glyphs.setdefault(n, []).append((int(gid), pen.bounds))

    choice, unmatched, hinting = {}, [], 0.0
    for n, rows in sorted(glyphs.items(), key=lambda item: int(item[0])):
        wanted = names.get(n, set())
        best = None
        for base, font in programs.items():
            if font is None or (wanted and base.split("+", 1)[-1] not in wanted):
                continue
            order_size, units = len(font.getGlyphOrder()), font["head"].unitsPerEm
            if any(gid >= order_size for gid, _ in rows):
                continue
            total, worst, fits = 0.0, 0.0, True
            for gid, bounds in rows:
                theirs = _glyph_bounds(font, gid)
                if (bounds is None) != (theirs is None):
                    fits = False
                    break
                if bounds is not None:
                    off = max(abs(p * units - q) for p, q in zip(bounds, theirs))
                    total, worst = total + off, max(worst, off)
            if fits and (best is None or total < best[1]):
                best = (base, total, worst)
        if best is None:
            unmatched.append(n)
        else:
            choice[n] = best[0]
            hinting = max(hinting, best[2])

    def redraw(match) -> str:
        n, gid = match.group(1), int(match.group(2))
        if n not in choice:
            return match.group(0)
        return f'<path id="font_{n}_{gid}" d="{_glyph_path(programs[choice[n]], gid)}"'

    out = _GLYPH.sub(redraw, svg)
    text_elements = len(re.findall(r"<text\b", out))
    report = {"fonts": len(programs) + len(not_embedded), "not_embedded": not_embedded,
              "unreadable": sorted(b for b, f in programs.items() if f is None), "svg_fonts": len(glyphs),
              "matched": {n: [choice[n], len(glyphs[n])] for n in sorted(choice, key=int)},
              "unmatched": unmatched, "named": all(n in names for n in choice), "text_elements": text_elements,
              "hinting_units": round(hinting, 2)}
    report["ok"] = not not_embedded and not unmatched and not text_elements and report["named"]
    return out, report


def _stroked_glyphs(page) -> dict:
    """Every glyph ``page`` strokes, by glyph id and origin (pt, rounded to 0.1), with
    the stroke's width in page space: the line width of the graphics state times the
    expansion of the transformation current when the text is painted -- read from MuPDF's
    own device calls, since its text trace reports the line width untransformed."""
    import math

    from pymupdf import mupdf

    class Device(mupdf.FzDevice2):
        def __init__(self):
            super().__init__()
            self.use_virtual_stroke_text()
            self.found = {}

        def stroke_text(self, _ctx, text, stroke, ctm, *_rest):
            width = stroke.linewidth * math.sqrt(abs(ctm.a * ctm.d - ctm.b * ctm.c))
            span = text.head
            while span:
                for i in range(span.len):
                    item = mupdf.FzTextSpan(span).items(i)
                    x = ctm.a * item.x + ctm.c * item.y + ctm.e
                    y = ctm.b * item.x + ctm.d * item.y + ctm.f
                    self.found[(item.gid, round(x, 1), round(y, 1))] = width
                span = span.next

    device = Device()
    mupdf.fz_run_page(page.this, device, mupdf.FzMatrix(), mupdf.FzCookie())
    mupdf.fz_close_device(device)
    return device.found


_STROKED_USE = re.compile(r'<use data-text="[^"]*" xlink:href="#font_\d+_(\d+)" stroke-width="([^"]+)"'
                          r'([^>]*?) transform="matrix\(([^)]*)\)"')


def stroked_text_widths(page, svg: str) -> tuple[str, dict]:
    """``svg`` with every stroked glyph's ``stroke-width`` corrected, and a report.

    **Why.**  PowerPoint draws a run it cannot find a bold face for with the regular
    face, filled *and* stroked (text render mode 2), and sets the line width before the
    ``cm`` that scales the text object: ``1.166667 w 2 Tr q 0.24 0 0 0.24 ... cm BT``,
    a stroke 0.28 pt wide.  Both PDF rasterisers draw it so.  MuPDF's SVG writer gives
    the stroke the line width as though it were already in page space -- 1.17 pt, four
    times too wide -- so ``real-financial-report``'s bold Japanese came out as blots
    (17,000 pixels beyond anti-aliasing a page).  The width is recomputed from MuPDF's
    own device calls (:func:`_stroked_glyphs`) and written in the glyph's space."""
    import math

    widths = _stroked_glyphs(page)
    fixed, unmatched = 0, []

    def repair(match) -> str:
        nonlocal fixed
        gid, matrix = int(match.group(1)), [float(v) for v in match.group(4).split(",")]
        key = (gid, round(matrix[4], 1), round(matrix[5], 1))
        width = widths.get(key)
        if width is None:
            near = [w for (g, x, y), w in widths.items()
                    if g == gid and abs(x - key[1]) <= 0.15 and abs(y - key[2]) <= 0.15]
            width = near[0] if near else None
        if width is None:
            unmatched.append(key)
            return match.group(0)
        fixed += 1
        scale = math.sqrt(abs(matrix[0] * matrix[3] - matrix[1] * matrix[2]))
        return match.group(0).replace(f'stroke-width="{match.group(2)}"', f'stroke-width="{width / scale:.9g}"', 1)

    out = _STROKED_USE.sub(repair, svg)
    return out, {"glyphs": fixed, "unmatched": unmatched[:5]}


_DATA_URI = re.compile(r'xlink:href="data:image/(png|jpeg);base64,([^"]*)"')
_INTENTS = {"Perceptual": 0, "RelativeColorimetric": 1, "Saturation": 2, "AbsoluteColorimetric": 3}


def srgb_images(document, index: int, svg: str) -> tuple[str, list[dict]]:
    """``svg`` with every image the PDF tags with an ICC profile other than sRGB converted
    to sRGB, and a report.

    **Why.**  The PDF says what colours an image's samples mean (``/ICCBased``), and
    both PDF rasterisers honour it; the SVG carries the samples and nothing else, and
    resvg reads them as sRGB.  ``real-college-template``'s photograph is tagged Adobe RGB
    (1998): drawn untransformed, its flat colours came out 9-11 levels off MuPDF's and
    pdfium's.  Each such image is found in the SVG by its bytes (a JPEG is carried over
    as the PDF's own stream), transformed with the embedded profile and the image's
    rendering intent (little CMS, through Pillow, as MuPDF uses little CMS), and written
    back as a PNG."""
    import base64

    from PIL import Image, ImageCms

    out, report = svg, []
    for entry in document[index].get_images(full=True):
        xref, colourspace = entry[0], entry[5]
        if colourspace != "ICCBased":
            continue
        described = document.xref_get_key(xref, "ColorSpace")
        found = re.search(r"(\d+) 0 R", described[1]) if described[0] == "xref" else None
        array = document.xref_object(int(found.group(1))) if found else described[1]
        profile_ref = re.search(r"/ICCBased\s+(\d+) 0 R", array)
        if profile_ref is None:
            continue
        profile_bytes = document.xref_stream(int(profile_ref.group(1)))
        profile = ImageCms.ImageCmsProfile(io.BytesIO(profile_bytes))
        name = ImageCms.getProfileDescription(profile).strip()
        row = {"xref": xref, "profile": name}
        if "srgb" in name.casefold().replace(" ", ""):
            continue
        raw = document.xref_stream_raw(xref)
        target = None
        for match in _DATA_URI.finditer(out):
            data = base64.b64decode(match.group(2))
            if data == raw:
                target = match
                break
        if target is None:
            row["error"] = "not found in the SVG by its bytes"
            report.append(row)
            continue
        image = Image.open(io.BytesIO(raw))
        if image.mode != "RGB":
            row["error"] = f"a {image.mode} image; only RGB is converted"
            report.append(row)
            continue
        intent = _INTENTS.get(document.xref_get_key(xref, "Intent")[1].lstrip("/"), 1)
        converted = ImageCms.profileToProfile(image, profile, ImageCms.createProfile("sRGB"),
                                              renderingIntent=intent, outputMode="RGB")
        buffer = io.BytesIO()
        converted.save(buffer, "PNG")
        encoded = base64.b64encode(buffer.getvalue()).decode()
        out = out[:target.start()] + f'xlink:href="data:image/png;base64,{encoded}"' + out[target.end():]
        row["converted"] = image.size
        report.append(row)
    return out, report


def bitmap_regions(pdf: Path, index: int) -> list[dict]:
    """Every raster image page ``index`` draws: its rectangle (pt), its pixel size, its
    resolution over that rectangle (dpi, the larger axis), and whether it is drawn under a
    soft mask (``smask``) -- PowerPoint's 3-D chart scenes and exported effects are; a
    picture the deck placed usually is not."""
    import pymupdf

    document = pymupdf.open(str(pdf))
    page = document[index]
    masks = {entry[0]: entry[1] for entry in page.get_images(full=True)}
    out = []
    for info in page.get_image_info(xrefs=True):
        x0, y0, x1, y1 = info["bbox"]
        width_pt, height_pt = max(x1 - x0, 1e-6), max(y1 - y0, 1e-6)
        dpi = max(info["width"] / width_pt, info["height"] / height_pt) * 72
        out.append({"bbox": (round(x0, 1), round(y0, 1), round(x1, 1), round(y1, 1)),
                    "pixels": (info["width"], info["height"]), "dpi": round(dpi),
                    "smask": bool(masks.get(info.get("xref"), 0)), "xref": info.get("xref")})
    return out


def validate(pdf: Path, supersample: int = 0, pages: list[int] | None = None, cache: bool = True,
             oracle: Path | None = None, dpis: tuple = (DPI,), width: int | None = None) -> list[dict]:
    """Per page and resolution: the conversion under one engine (MuPDF's PDF raster
    against MuPDF's SVG raster), resvg's raster of the SVG against MuPDF's and pdfium's
    rasters of the PDF -- at 1x and, when ``supersample``, both drawn at that factor and
    area-averaged down -- and :func:`unhinted_outlines`' report.

    ``dpis`` are the resolutions validated; ``width`` adds the one at which the page is
    that many pixels across (``tools/fidelity.py``'s scoring resolution)."""
    import pymupdf

    converted = convert(pdf, cache, oracle)
    document = pymupdf.open(str(pdf))
    out = []
    for index, (svg, report) in enumerate(converted):
        if pages is not None and index not in pages:
            continue
        images = bitmap_regions(pdf, index)
        resolutions = [(dpi, False) for dpi in dpis] + ([(dpi_for_width(svg, width), True)] if width else [])
        for dpi, scoring in resolutions:
            ours = rasterise(svg, dpi)
            truth = _mupdf(document, index, dpi)
            row = {"page": index + 1, "dpi": round(dpi, 2), "scoring": scoring, "outlines": report,
                   "images": images, "mupdf_pdf_vs_resvg_svg": compare(truth, ours, images, dpi),
                   "pdfium_vs_resvg_svg": compare(_pdfium(pdf, index, dpi), ours, images, dpi)}
            # MuPDF's own SVG reader ignores <mask> (an image's soft mask: PowerPoint's chart
            # scenes), drawing the mask's image black, and does not tile a <pattern>
            # (feature-sweep's tile fills: 0.08 against its own PDF raster, where resvg's
            # raster of the same SVG agrees with MuPDF's PDF raster to anti-aliasing);
            # there the same-engine route says nothing about the conversion, and resvg's
            # route stands alone.
            if "<mask" not in svg and "<pattern" not in svg:
                svg_document = pymupdf.open(stream=device_grid(svg, dpi).encode(), filetype="svg")
                row["mupdf_pdf_vs_mupdf_svg"] = compare(truth, _mupdf(svg_document, 0, 72), images, dpi)
            if not scoring and not faithful(row):
                # Stroked text (PowerPoint's synthetic bold) goes through MuPDF's glyph cache
                # too: real-financial-report page 2 at 300 dpi, 222 isolated pixels beyond
                # anti-aliasing at 1x, 3 with both rasters drawn at 2x and averaged down, 0 at
                # 4x.  Where 1x fails, 2x is the gate (:func:`faithful`).
                big = _downsample(_mupdf(document, index, dpi * 2), 2)
                row["gate"] = compare(big, _downsample(rasterise(svg, dpi, scale=2), 2), images, dpi)
            if scoring:
                # At the scoring resolution (96-128 dpi) MuPDF's glyph cache puts a glyph on
                # a sub-pixel grid that is a larger share of a stem, and it minifies a bitmap
                # with a filter of its own: at 1x, 28-101 isolated pixels a page on
                # chart-gallery, on text and 3-D scenes, where the same pages hold at 300 dpi.
                # Both rasters drawn at 4x and averaged down to this grid leave none of it,
                # so that is the gate here; 1x is reported beside it.
                big = _downsample(_mupdf(document, index, dpi * SUPERSAMPLE), SUPERSAMPLE)
                row["gate"] = compare(big, _downsample(rasterise(svg, dpi, scale=SUPERSAMPLE), SUPERSAMPLE),
                                      images, dpi)
            if supersample:
                big = _downsample(_pdfium(pdf, index, dpi * supersample), supersample)
                row[f"pdfium_vs_resvg_svg_{supersample}x"] = compare(
                    big, _downsample(rasterise(svg, dpi, scale=supersample), supersample), images, dpi)
            out.append(row)
    return out


def validation_pdfs(oracle: Path | None = None) -> list[Path]:
    """Every export in the oracle directory that has its deck beside it -- the pages
    ``tools/fidelity.py`` scores, and those of the decks it skips."""
    oracle = Path(os.path.expanduser(str(oracle))) if oracle is not None else DEFAULT_ORACLE
    return [pdf for pdf in sorted(oracle.glob("*.pdf")) if pdf.with_suffix(".pptx").exists()]


def faithful(row: dict) -> bool:
    """Whether :func:`validate`'s row passes: every glyph the embedded program's; no flat
    colour off by more than :data:`MAX_FLAT_COLOUR`; and at most :data:`MAX_BEYOND`
    isolated pixels beyond anti-aliasing -- between MuPDF's raster of the PDF and resvg's
    of the SVG over the whole page, embedded bitmaps included, and between MuPDF's two
    rasters outside the bitmaps.

    MuPDF's own SVG reader is not a faithful reader of everything the conversion writes,
    measured by the sibling where resvg's and pdfium's rasters agree with MuPDF's PDF
    raster exactly: it ignores ``<mask>`` (an image's soft mask: PowerPoint's chart scenes
    drawn black), and it filters a stretched bitmap its own way.  So the same-engine
    route is skipped on a page with a mask and holds only outside the bitmaps.

    At the scoring resolution (``row["scoring"]``) both routes are held on rasters drawn
    at :data:`SUPERSAMPLE` times and averaged down (``row["gate"]``): see :func:`validate`."""
    same, cross = row.get("mupdf_pdf_vs_mupdf_svg"), row["mupdf_pdf_vs_resvg_svg"]

    def holds(same, cross, floor=0.0) -> bool:
        return (row["outlines"]["ok"]
                and (same is None or (same["beyond"] <= MAX_BEYOND and same["flat_colour"] <= MAX_FLAT_COLOUR))
                and cross["beyond"] + cross["beyond_in_images"] <= MAX_BEYOND
                and cross["flat_colour_all"] <= MAX_FLAT_COLOUR
                and cross["ssim"] >= floor)

    if row.get("scoring"):
        return holds(None, row["gate"], MIN_GATE_SSIM)
    if holds(same, cross):
        return True
    # 1x failed: the supersampled route decides, and the same-engine route still may not
    # change a colour.
    return ("gate" in row and holds(None, row["gate"], MIN_GATE_SSIM)
            and (same is None or same["flat_colour"] <= MAX_FLAT_COLOUR))


def verdicts(pdf: Path, oracle: Path | None = None, width: int | None = None) -> dict:
    """Every page's validation, cached beside its conversion (``validation-<width>.json``):
    ``{page: {"faithful": bool, "dpi": [...], "why": str}}`` at :data:`DPI` and at the
    scoring resolution ``width`` px across.  Computed once per PDF and converter version
    -- about half a minute a page -- and read by ``tools/fidelity.py``, which will not
    score a slide against a page that fails."""
    import json

    import fidelity

    width = width or fidelity.WIDTH
    path = cache_dir(pdf, oracle) / f"validation-{width}.json"
    if path.exists():
        return {int(k): v for k, v in json.loads(path.read_text()).items()}
    return record_verdicts(pdf, validate(pdf, oracle=oracle, width=width), oracle, width)


def record_verdicts(pdf: Path, rows: list[dict], oracle: Path | None, width: int) -> dict:
    """:func:`verdicts` from :func:`validate`'s ``rows`` (every page, at :data:`DPI` and
    ``width``), written to the cache."""
    import json

    path = cache_dir(pdf, oracle) / f"validation-{width}.json"
    out: dict = {}
    for row in rows:
        entry = out.setdefault(row["page"], {"faithful": True, "dpi": [], "why": ""})
        entry["dpi"].append(row["dpi"])
        if not faithful(row):
            entry["faithful"] = False
            entry["why"] = (entry["why"] + "; " if entry["why"] else "") + _why(row)
    convert(pdf, oracle=oracle)  # the cache directory exists
    _refuse_repository(path)
    path.write_text(json.dumps(out, indent=1, sort_keys=True))
    return out


def _why(row: dict) -> str:
    outlines = row["outlines"]
    if not outlines["ok"]:
        return f"{row['dpi']} dpi: glyphs or images not accounted for"
    same, cross = row.get("mupdf_pdf_vs_mupdf_svg"), row.get("gate") or row["mupdf_pdf_vs_resvg_svg"]
    parts = [f"{row['dpi']} dpi" + (f" ({SUPERSAMPLE if row.get('scoring') else 2}x)" if row.get("gate") else "")]
    if same is not None and not row.get("scoring"):
        parts.append(f"MuPDF pdf/svg beyond {same['beyond']} colour {same['flat_colour']}")
    parts.append(f"MuPDF/resvg SSIM {cross['ssim']} beyond {cross['beyond'] + cross['beyond_in_images']} "
                 f"colour {cross['flat_colour_all']}")
    return ", ".join(parts)


def _print_row(name: str, row: dict, supersample: int) -> bool:
    same, cross, pdfium = row.get("mupdf_pdf_vs_mupdf_svg"), row["mupdf_pdf_vs_resvg_svg"], row["pdfium_vs_resvg_svg"]
    big = row.get(f"pdfium_vs_resvg_svg_{supersample}x")
    outlines = row["outlines"]
    ok = faithful(row)

    def route(r) -> str:
        return (f"{r['ssim']:.4f} colour {r['flat_colour_all']} beyond {r['beyond']:>4}"
                + (f" (+{r['beyond_in_images']} in bitmaps)" if r["beyond_in_images"] else ""))

    print(f"{name[:26]:26} p{row['page']:<3} {row['dpi']:>6.1f} dpi  MuPDF pdf/svg "
          + (route(same) if same else "n/a (<mask> or <pattern>)")
          + " | MuPDF/resvg " + route(cross)
          + (f" | {SUPERSAMPLE if row.get('scoring') else 2}x (gate) " + route(row["gate"]) if row.get("gate") else "")
          + " | pdfium/resvg " + route(pdfium)
          + (f" | {supersample}x " + route(big) if big else "")
          + f" | glyphs {sum(n for _, n in outlines['matched'].values())} in {outlines['svg_fonts']} fonts"
          + f" (hinting moved <= {outlines['hinting_units']} units)"
          + ("" if outlines["named"] else " UNNAMED")
          + (f" UNMATCHED {outlines['unmatched']}" if outlines["unmatched"] else "")
          + (f" NOT EMBEDDED {outlines['not_embedded']}" if outlines["not_embedded"] else "")
          + (f" stroked {outlines['stroked_text']['glyphs']}" if outlines.get("stroked_text", {}).get("glyphs") else "")
          + (f" UNMATCHED STROKES {outlines['stroked_text']['unmatched']}"
             if outlines.get("stroked_text", {}).get("unmatched") else "")
          + "".join(f" ICC {i['profile']}->sRGB" + (f" ERROR {i['error']}" if i.get("error") else "")
                    for i in outlines.get("icc_images", []))
          + (f" bitmaps {len(row['images'])}" if row["images"] else "")
          + ("" if ok else "  FAIL"))
    return ok


def _print_bitmaps(pdf: Path) -> None:
    import pymupdf

    for index in range(pymupdf.open(str(pdf)).page_count):
        for image in bitmap_regions(pdf, index):
            print(f"{pdf.stem[:26]:26} p{index + 1:<3} bbox {image['bbox']} pt  {image['pixels'][0]} x "
                  f"{image['pixels'][1]} px  {image['dpi']} dpi" + ("  soft mask" if image["smask"] else ""))


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pdf", nargs="*", type=Path)
    parser.add_argument("--oracle", type=Path, default=DEFAULT_ORACLE,
                        help="directory of deck.pptx / deck.pdf pairs; converted pages are cached in its svg/")
    parser.add_argument("--validate", action="store_true",
                        help="hold the converter to the PDFs (default: every export in the oracle directory)")
    parser.add_argument("--bitmaps", action="store_true", help="list the bitmaps each page draws")
    parser.add_argument("--supersample", type=int, default=0,
                        help="also compare pdfium and resvg at this factor (default 0: no)")
    parser.add_argument("--dpi", type=float, action="append",
                        help=f"validate at this resolution (repeatable; default {DPI} and the scoring width)")
    parser.add_argument("--max-pages", type=int, help="validate at most this many pages of each PDF")
    parser.add_argument("--no-cache", action="store_true", help="convert in memory only (nothing is written)")
    args = parser.parse_args(argv[1:])
    if not available():
        print("PyMuPDF is not installed: see README.md, 'Checking fidelity against PowerPoint'")
        return 2
    oracle = Path(os.path.expanduser(str(args.oracle)))
    pdfs = args.pdf or validation_pdfs(oracle)
    if args.bitmaps:
        for pdf in pdfs:
            _print_bitmaps(pdf)
        return 0
    if not args.validate:
        for pdf in pdfs:
            page_svgs(pdf, oracle=oracle)
            print(cache_dir(pdf, oracle))
        return 0
    import fidelity

    dpis = tuple(args.dpi) if args.dpi else (DPI,)
    width = None if args.dpi else fidelity.WIDTH
    failures = 0
    for pdf in pdfs:
        pages = list(range(args.max_pages)) if args.max_pages else None
        rows = validate(pdf, args.supersample, pages, cache=not args.no_cache, oracle=oracle, dpis=dpis,
                        width=width)
        for row in rows:
            failures += not _print_row(pdf.stem, row, args.supersample)
        if pages is None and width and dpis == (DPI,) and not args.no_cache:
            record_verdicts(pdf, rows, oracle, width)  # what tools/fidelity.py reads
    print(f"{failures} page(s) failed" if failures else "every page converted within anti-aliasing")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

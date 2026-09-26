"""The fidelity instrument's converter (``tools/pdf_svg.py``): PowerPoint's PDF page to SVG.

The instrument's own arithmetic runs everywhere numpy does.  The converter's faithfulness
to PowerPoint's PDF -- same engine, two routes; no pixel beyond anti-aliasing; no colour
changed; glyphs as the embedded outlines -- needs PyMuPDF (the ``fidelity`` extra, in the
project's ``.venv``) and PowerPoint's exports in ``~/pptx2svg-oracle``, so it skips where
either is missing, CI among them.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "tools"))

import pdf_svg  # noqa: E402  (imports nothing outside the standard library at module level)

SLIDE = ('<svg xmlns="http://www.w3.org/2000/svg" version="1.1" width="720" height="405" '
         'viewBox="0 0 720 405">\n<path d="M0 0H10V10Z"/>\n</svg>\n')
A4 = ('<svg xmlns="http://www.w3.org/2000/svg" version="1.1" width="595.2" height="841.92" '
      'viewBox="0 0 595.2 841.92">\n<path d="M0 0H10V10Z"/>\n</svg>\n')


def test_a_slide_is_drawn_at_the_scoring_width():
    """``tools/fidelity.py`` scores at 1280 px across: 128 dpi on a 720 pt slide, 96 on a
    960 pt one, and the root is sized in those device pixels exactly."""
    assert pdf_svg.dpi_for_width(SLIDE, 1280) == pytest.approx(128)
    sized = pdf_svg.device_grid(SLIDE, pdf_svg.dpi_for_width(SLIDE, 1280))
    assert 'width="1280" height="720"' in sized
    assert sized.endswith('<path d="M0 0H10V10Z"/>\n</svg>\n')


def test_the_page_is_drawn_on_the_device_grid():
    """resvg rounds a root's size in points before zooming (the sibling's A4 came out
    2,479 px wide and stretched 1e-4 tall); the root is sized in device pixels and the
    viewBox widened to match, so one point is exactly dpi/72 px."""
    sized = pdf_svg.device_grid(A4)
    assert 'width="2480" height="3508"' in sized
    width, height = (float(v) for v in sized.split('viewBox="0 0 ')[1].split('"')[0].split())
    assert abs(2480 / width - 300 / 72) < 1e-9 and abs(3508 / height - 300 / 72) < 1e-9


def test_an_edge_is_anti_aliasing_and_a_moved_stroke_is_not():
    np = pytest.importorskip("numpy")
    page = np.full((40, 40, 3), 255, np.uint8)
    page[10:30, 10:14] = 0
    softer = page.copy()
    softer[10:30, 14] = 128  # the same stroke's edge covered differently
    assert pdf_svg.beyond_antialiasing(page, softer).sum() == 0
    moved = np.full_like(page, 255)
    moved[10:30, 13:17] = 0  # three pixels right
    assert pdf_svg.beyond_antialiasing(page, moved).sum() > 0
    recoloured = page.copy()
    recoloured[10:30, 10:14] = (200, 0, 0)
    assert pdf_svg.beyond_antialiasing(page, recoloured).sum() > 0


def test_a_changed_colour_is_seen_where_the_histogram_would_not_say():
    np = pytest.importorskip("numpy")
    page = np.full((20, 20, 3), 255, np.uint8)
    page[5:15, 5:15] = (204, 210, 215)
    rounded = page.copy()
    rounded[5:15, 5:15] = (204, 210, 216)  # a level of rounding: faithful
    assert pdf_svg.flat_colour_difference(page, rounded) == 1
    shifted = page.copy()
    shifted[5:15, 5:15] = (204, 200, 215)
    assert pdf_svg.flat_colour_difference(page, shifted) == 10


def test_a_converted_page_is_never_written_into_a_repository(tmp_path):
    """A converted page holds Microsoft's glyph outlines: not in this checkout, and not
    in any other git checkout an oracle directory might be pointed into."""
    with pytest.raises(RuntimeError):
        pdf_svg._refuse_repository(REPO / "tests" / "anything")
    checkout = tmp_path / "elsewhere"
    (checkout / ".git").mkdir(parents=True)
    with pytest.raises(RuntimeError):
        pdf_svg._refuse_repository(checkout / "svg" / "deck")
    pdf_svg._refuse_repository(tmp_path / "oracle" / "svg")


#: A sample of PowerPoint's pages (``python tools/pdf_svg.py --validate`` runs every one):
#: Aptos Display and Aptos, whose hinted outlines MuPDF moves 30-35 units (``table-test``);
#: a 3-D chart scene PowerPoint embedded as a bitmap under a soft mask (``chart-gallery``
#: 13); synthetic bold, whose stroke MuPDF writes four times too wide
#: (``real-financial-report`` 2).  (The Adobe RGB photograph of the local-only
#: ``real-college-template`` 8 is held by the full validation, not here: 20 s a page.)
SAMPLE = (("table-test", 0), ("chart-gallery", 12), ("real-financial-report", 1))


def _needs_the_converter():
    if not pdf_svg.available():
        pytest.skip("PyMuPDF (the fidelity extra) is not installed here")
    for module in ("numpy", "PIL", "resvg_py", "fontTools", "pypdfium2"):
        pytest.importorskip(module)


def test_the_converter_is_faithful_to_powerpoints_pdf():
    """At the scoring resolution: MuPDF's raster of the PDF and resvg's raster of the SVG,
    both drawn at 4x and averaged down, differ only at anti-aliasing scale, in no colour;
    every glyph is the embedded program's, and the SVG has no text."""
    _needs_the_converter()
    import fidelity

    checked = 0
    for stem, page in SAMPLE:
        pdf = pdf_svg.DEFAULT_ORACLE / f"{stem}.pdf"
        if not pdf.exists():
            continue
        for row in pdf_svg.validate(pdf, pages=[page], dpis=(), width=fidelity.WIDTH, cache=False):
            where = f"{stem} page {row['page']}"
            assert pdf_svg.faithful(row), (where, row)
            assert row["outlines"]["text_elements"] == 0, where
            checked += 1
    if not checked:
        pytest.skip("PowerPoint's exports are not in ~/pptx2svg-oracle here")


def test_stroked_text_is_as_wide_as_powerpoint_drew_it():
    """``1.166667 w`` before a ``0.24`` ``cm``: a stroke 0.28 pt wide, which MuPDF's SVG
    writes as 1.17 pt.  Every stroked glyph is found, and none is left as written."""
    _needs_the_converter()
    pdf = pdf_svg.DEFAULT_ORACLE / "real-financial-report.pdf"
    if not pdf.exists():
        pytest.skip("PowerPoint's exports are not in ~/pptx2svg-oracle here")
    import pymupdf

    page = pymupdf.open(str(pdf))[1]
    raw = page.get_svg_image(text_as_path=True)
    fixed, report = pdf_svg.stroked_text_widths(page, raw)
    assert report == {"glyphs": 37, "unmatched": []}
    widths = {float(w) for w in pdf_svg.re.findall(r'stroke-width="([^"]+)"', fixed)
              if w not in pdf_svg.re.findall(r'stroke-width="([^"]+)"', raw)}
    assert widths and all(abs(w * 8 - 0.28) < 0.01 for w in widths), widths  # 8 pt glyphs


def test_a_page_the_converter_cannot_hold_is_refused():
    """``feature-sweep`` slide 13: pattern fills whose tiles each rasteriser samples in its
    own phase.  The gate says so, and ``tools/fidelity.py`` then reports the slide rather
    than averaging it."""
    _needs_the_converter()
    import fidelity

    pdf = pdf_svg.DEFAULT_ORACLE / "feature-sweep.pdf"
    if not pdf.exists():
        pytest.skip("PowerPoint's exports are not in ~/pptx2svg-oracle here")
    rows = pdf_svg.validate(pdf, pages=[12], dpis=(), width=fidelity.WIDTH, cache=False)
    assert rows and not any(pdf_svg.faithful(row) for row in rows)

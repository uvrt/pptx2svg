"""``ConvertOptions(text="outlines")``: the slide's text as the glyph paths its PNG is drawn
with (:func:`pptx2svg.png.outline_svg`, :mod:`ooxml_common.outlines`).

Reproducible renders only (``host_fonts=False``: the bundle's faces), so every machine
draws the same.  The outlined SVG, rasterised, must draw the pixels the text draws:
resvg draws both, shaping and placing the text itself in one and drawing the paths in
the other.
"""

from __future__ import annotations

import io
import re

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_png, convert_pptx_to_svg, rasterise_svgs, render_svgs
from pptx2svg.png import available_backends

pytestmark = pytest.mark.skipif(not available_backends(), reason="no SVG rasterizer installed")


def _options(**kwargs) -> ConvertOptions:
    return ConvertOptions(host_fonts=False, **kwargs)


def _pixels(png: bytes):
    np = pytest.importorskip("numpy")
    image = pytest.importorskip("PIL.Image")
    return np.asarray(image.open(io.BytesIO(png)).convert("RGBA")).astype(int)


def _compare(path):
    text = render_svgs(path, _options())[0]
    outlines_options = _options(text="outlines")
    outlines = render_svgs(path, outlines_options)[0]
    return text, outlines, outlines_options


@pytest.mark.parametrize("name", ["sample.pptx", "sample-cjk.pptx", "chart-gallery.pptx"])
def test_outlined_slides_draw_the_pixels_the_text_draws(name):
    from pathlib import Path

    path = Path(__file__).parent / "fixtures" / name
    text, outlines, options = _compare(path)
    assert not [w for w in options.warnings if w.code == "text-not-outlined"]
    assert all("<text" not in svg for svg in outlines)
    for a, b in zip(rasterise_svgs(text, _options()), rasterise_svgs(outlines, _options())):
        diff = abs(_pixels(a) - _pixels(b)).max(axis=2)
        # An anti-aliased edge pixel here and there: a glyph's path is placed through
        # another chain of float32 transforms than resvg's own.
        assert diff.max() <= 64
        assert (diff > 16).sum() <= 200


def test_text_left_as_text_is_reported_once(product_page):
    # The emoji's face (where macOS has one) draws pictures, not outlines; without it
    # every text is outlined.
    options = _options(text="outlines")
    outlines = convert_pptx_to_svg(product_page, options)
    kept = [w for w in options.warnings if w.code == "text-not-outlined"]
    assert len(kept) == len({w.message for w in kept})
    assert sum(svg.count("<text") for svg in outlines) >= len(kept)


def test_the_default_is_text_and_unchanged(product_page):
    assert convert_pptx_to_svg(product_page, _options()) == convert_pptx_to_svg(product_page, _options(text="text"))
    assert "<text" in convert_pptx_to_svg(product_page, _options())[0]


def test_an_unknown_text_mode_is_refused(product_page):
    with pytest.raises(ValueError, match="text must be one of text, outlines"):
        convert_pptx_to_svg(product_page, _options(text="paths"))


def test_selectable_text_is_kept_transparent_over_the_outlines(product_page):
    svg = convert_pptx_to_svg(product_page, _options(text="outlines", selectable_text=True))[0]
    layers = re.findall(r"<text\b[^>]*data-text-layer[^>]*>", svg)
    assert layers and all('fill-opacity="0"' in layer for layer in layers)
    assert 'aria-hidden="true"' in svg


def test_png_output_is_the_same_whatever_the_text_mode(product_page):
    assert convert_pptx_to_png(product_page, _options(text="outlines")) == convert_pptx_to_png(
        product_page, _options())

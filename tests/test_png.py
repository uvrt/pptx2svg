"""PNG rasterisation through the external backend."""

from __future__ import annotations

import struct

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_png
from pptx2svg.png import available_backends, svg_to_png

pytestmark = pytest.mark.skipif(
    not available_backends(),
    reason="no SVG rasterizer installed (pip install pptx2svg[png])",
)

PNG_MAGIC = b"\x89PNG\r\n\x1a\n"

SIMPLE_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 20" width="40" height="20">'
    '<rect width="40" height="20" fill="#3366cc"/></svg>'
)


def png_size(data: bytes) -> tuple[int, int]:
    """Read width/height out of the IHDR chunk."""
    assert data[:8] == PNG_MAGIC
    width, height = struct.unpack(">II", data[16:24])
    return width, height


def test_rasterizes_to_png():
    assert svg_to_png(SIMPLE_SVG)[:8] == PNG_MAGIC


def test_output_size_follows_the_svg_by_default():
    assert png_size(svg_to_png(SIMPLE_SVG)) == (40, 20)


def test_explicit_width_scales_the_output():
    width, height = png_size(svg_to_png(SIMPLE_SVG, width=400))
    assert width == 400
    assert height == 200  # aspect ratio preserved by the backend


def test_unknown_backend_is_rejected():
    with pytest.raises(ValueError):
        svg_to_png(SIMPLE_SVG, backend="nonexistent")


def test_convert_pptx_to_png_produces_one_image_per_slide(product_page):
    images = convert_pptx_to_png(product_page, ConvertOptions(width=640))
    assert len(images) == 1
    assert images[0][:8] == PNG_MAGIC
    assert png_size(images[0])[0] == 640


def test_rendered_slide_is_not_blank(product_page):
    """A slide with text must produce more than a flat background."""
    image = convert_pptx_to_png(product_page, ConvertOptions(width=320))[0]
    # A blank fill compresses far smaller than a slide with glyphs on it.
    assert len(image) > 2000

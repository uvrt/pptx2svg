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


# --------------------------------------------------------------------------------------
# Fonts: reproducible by default, and never silently degraded
# --------------------------------------------------------------------------------------

TEXT_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="400" height="60">'
    '<rect width="400" height="60" fill="#fff"/>'
    '<text x="10" y="40" font-family="Calibri, Carlito, sans-serif" font-size="28"'
    ' fill="#000">Handgloves</text></svg>'
)


def test_bundled_fonts_are_used_and_the_host_is_ignored_by_default():
    """The default render must not depend on what this machine happens to have.

    Compared against an explicit system-fonts render rather than against a stored hash:
    a hash would pin this test to one resvg build, while the property that matters is
    that the two configurations are *different*, which is only true if the default is
    really reading the bundle.
    """
    from pptx2svg.fonts import bundle_dir

    if bundle_dir() is None:
        pytest.skip("pptx2svg-fonts is not importable")
    default = svg_to_png(TEXT_SVG, backend="resvg")
    system = svg_to_png(
        TEXT_SVG, backend="resvg", use_bundled_fonts=False, skip_system_fonts=False
    )
    assert default[:8] == PNG_MAGIC
    # This machine has neither Calibri nor Carlito outside the bundle, so a default
    # render that matched the system one would mean the bundle was never consulted.
    assert default != system


def test_without_a_bundle_the_host_fonts_are_used_rather_than_none(monkeypatch):
    """Skipping system fonts with nothing to replace them renders blank slides."""
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
    image = svg_to_png(TEXT_SVG, backend="resvg")
    assert image[:8] == PNG_MAGIC
    # Text drawn with *something* compresses larger than an empty white rectangle.
    blank = svg_to_png(
        '<svg xmlns="http://www.w3.org/2000/svg" width="400" height="60">'
        '<rect width="400" height="60" fill="#fff"/></svg>',
        backend="resvg",
    )
    assert len(image) > len(blank)


def test_caller_font_dirs_take_precedence_over_the_bundle(tmp_path):
    """An empty directory must not knock out the bundle; order is dirs, then bundle."""
    (tmp_path / "not-a-font.txt").write_text("x")
    image = svg_to_png(TEXT_SVG, backend="resvg", font_dirs=[str(tmp_path)])
    assert image[:8] == PNG_MAGIC
    assert image == svg_to_png(TEXT_SVG, backend="resvg")

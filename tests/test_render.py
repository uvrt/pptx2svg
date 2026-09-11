"""SVG generation, text layout and PNG rasterisation."""

from __future__ import annotations

import re
from xml.etree.ElementTree import fromstring

import pytest

from pptx2svg import convert_pptx_to_svg, model as m
from pptx2svg.render.context import RenderContext
from pptx2svg.render.geometry import PRESET_GEOMETRIES, preset_geometry_svg, render_geometry
from pptx2svg.render.shape import build_transform_attr
from pptx2svg.render.svg import render_slide_to_svg
from pptx2svg.text.measure import DefaultTextMeasurer
from pptx2svg.text.wrap import wrap_paragraph


# -- Geometry --------------------------------------------------------------------------


@pytest.mark.parametrize("preset", sorted(PRESET_GEOMETRIES))
def test_every_preset_emits_one_well_formed_element(preset):
    svg = preset_geometry_svg(preset, 200, 100, {})
    fromstring(svg)  # raises on malformed markup


def test_unknown_preset_falls_back_to_a_rectangle():
    assert preset_geometry_svg("notARealShape", 40, 20, {}) == '<rect width="40" height="20"/>'


def test_custom_geometry_is_scaled_onto_the_shape_box():
    geometry = m.CustomGeometry(
        paths=[m.CustomGeometryPath(width=100, height=100, commands="M 0 0 L 100 100")]
    )
    svg = render_geometry(geometry, 200, 50)
    assert 'transform="scale(2, 0.5)"' in svg


# -- Transforms ------------------------------------------------------------------------


def test_transform_translates_rotates_about_the_centre_then_flips():
    transform = m.Transform(
        offset_x=914400,  # 1 inch -> 96 px
        offset_y=0,
        extent_width=914400,
        extent_height=914400,
        rotation=45,
        flip_h=True,
    )
    attr = build_transform_attr(transform)
    assert attr == "translate(96, 0) rotate(45, 48, 48) translate(96, 0) scale(-1, 1)"


def test_transform_without_rotation_or_flip_is_just_a_translate():
    assert build_transform_attr(m.Transform(offset_x=0, offset_y=0)) == "translate(0, 0)"


# -- Text layout -----------------------------------------------------------------------


def make_paragraph(text: str, **properties) -> m.Paragraph:
    return m.Paragraph(runs=[m.TextRun(text, m.RunProperties(**properties))])


def test_wrapping_breaks_latin_text_at_spaces():
    paragraph = make_paragraph(
        "The quick brown fox jumps over the lazy dog", font_size=18, font_family="Calibri"
    )
    lines = wrap_paragraph(paragraph, 200)
    assert len(lines) > 1
    for line in lines:
        assert not "".join(s.text for s in line.segments).startswith(" ")


def test_wrapping_breaks_cjk_between_any_characters():
    paragraph = make_paragraph("日本語のテキストです" * 3, font_size=18, font_family="Meiryo")
    lines = wrap_paragraph(paragraph, 100)
    assert len(lines) > 1


def test_a_forced_break_starts_a_new_line():
    paragraph = m.Paragraph(runs=[m.TextRun("one\ntwo", m.RunProperties(font_size=12))])
    lines = wrap_paragraph(paragraph, 1000)
    assert [" ".join(s.text for s in line.segments) for line in lines] == ["one", "two"]


def test_an_overlong_word_is_split_by_character():
    paragraph = make_paragraph("supercalifragilistic", font_size=24, font_family="Calibri")
    lines = wrap_paragraph(paragraph, 30)
    assert len(lines) > 1


def test_measurement_scales_linearly_with_font_size():
    measurer = DefaultTextMeasurer()
    small = measurer.measure_text_width("Hello", 10, False, "Calibri")
    large = measurer.measure_text_width("Hello", 20, False, "Calibri")
    assert large == pytest.approx(small * 2)


def test_bold_text_measures_wider():
    measurer = DefaultTextMeasurer()
    plain = measurer.measure_text_width("Hello", 18, False, "Calibri")
    bold = measurer.measure_text_width("Hello", 18, True, "Calibri")
    assert bold > plain


def test_cjk_characters_measure_full_width():
    measurer = DefaultTextMeasurer()
    width = measurer.measure_text_width("日", 18, False, None, "Meiryo")
    assert width == pytest.approx(18 * 96 / 72, rel=0.01)


# -- Slide SVG -------------------------------------------------------------------------


def minimal_slide() -> tuple[m.Slide, m.SlideSize]:
    shape = m.ShapeElement(
        transform=m.Transform(offset_x=0, offset_y=0, extent_width=914400, extent_height=914400),
        geometry=m.PresetGeometry(preset="rect"),
        fill=m.SolidFill(color=m.ResolvedColor(hex="#ff0000")),
    )
    return m.Slide(slide_number=1, elements=[shape]), m.SlideSize(9144000, 5143500)


def test_slide_svg_is_well_formed_and_sized_from_the_slide():
    slide, size = minimal_slide()
    svg = render_slide_to_svg(slide, size)
    root = fromstring(svg)
    assert root.get("viewBox") == "0 0 960 540"
    assert root.get("width") == "960"
    assert root.get("height") == "540"


def test_width_only_keeps_the_aspect_ratio():
    slide, size = minimal_slide()
    root = fromstring(render_slide_to_svg(slide, size, width=1920))
    assert root.get("width") == "1920"
    assert root.get("height") == "1080"


def test_missing_background_paints_white_not_transparent():
    slide, size = minimal_slide()
    assert 'fill="#FFFFFF"' in render_slide_to_svg(slide, size)


def test_ids_are_deterministic_across_runs():
    """Gradient/pattern ids come from a counter, so output is byte-stable."""
    slide, size = minimal_slide()
    slide.elements[0].fill = m.GradientFill(
        stops=[
            m.GradientStop(0.0, m.ResolvedColor("#000000")),
            m.GradientStop(1.0, m.ResolvedColor("#ffffff")),
        ]
    )
    first = render_slide_to_svg(slide, size, RenderContext())
    second = render_slide_to_svg(slide, size, RenderContext())
    assert first == second
    assert "grad-1" in first


def test_group_maps_the_child_coordinate_space():
    group = m.GroupElement(
        transform=m.Transform(offset_x=0, offset_y=0, extent_width=914400, extent_height=914400),
        child_transform=m.Transform(
            offset_x=914400, offset_y=0, extent_width=1828800, extent_height=1828800
        ),
    )
    svg = render_slide_to_svg(
        m.Slide(slide_number=1, elements=[group]), m.SlideSize(9144000, 5143500)
    )
    # ext/chExt = 0.5, and chOff is translated away.
    assert "scale(0.5, 0.5)" in svg
    assert "translate(-96, 0)" in svg


def test_zero_child_extent_does_not_divide_by_zero():
    group = m.GroupElement(
        transform=m.Transform(offset_x=0, offset_y=0, extent_width=914400, extent_height=914400),
        child_transform=m.Transform(offset_x=0, offset_y=0, extent_width=0, extent_height=0),
    )
    svg = render_slide_to_svg(
        m.Slide(slide_number=1, elements=[group]), m.SlideSize(9144000, 5143500)
    )
    assert "scale(1, 1)" in svg


# -- End to end ------------------------------------------------------------------------


def test_fixtures_render_to_well_formed_svg(pptx_path):
    for document in convert_pptx_to_svg(pptx_path):
        fromstring(document)


def test_font_size_is_written_without_a_unit_suffix(pptx_path):
    """resvg rejects `pt` on font-size presentation attributes, so we emit user units."""
    for document in convert_pptx_to_svg(pptx_path):
        assert not re.search(r'font-size="[\d.]+(pt|em|%)"', document)


def test_no_css_classes_or_style_elements(pptx_path):
    """librsvg and resvg do not apply CSS selectors; everything must be an attribute."""
    for document in convert_pptx_to_svg(pptx_path):
        assert "<style" not in document
        assert 'class="' not in document


def test_slide_text_reaches_the_svg(product_page):
    document = convert_pptx_to_svg(product_page)[0]
    assert "The Next Generation of Workflow Automation" in document


def test_rendering_is_reproducible(product_page):
    assert convert_pptx_to_svg(product_page) == convert_pptx_to_svg(product_page)

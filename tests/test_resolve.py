"""Theme resolution and the placeholder / text inheritance cascades."""

from __future__ import annotations


from pptx2svg import ConvertOptions, convert_pptx_to_model, model as m
from pptx2svg.opc import OpcPackage
from pptx2svg.parse.parts import read_presentation
from pptx2svg.resolve import resolve_presentation


def resolve(path):
    package = OpcPackage.open(str(path))
    return package, resolve_presentation(package, read_presentation(package))


def walk(elements):
    for element in elements:
        yield element
        if isinstance(element, m.GroupElement):
            yield from walk(element.children)


def test_every_fixture_resolves(pptx_path):
    _, resolved = resolve(pptx_path)
    assert resolved.slides
    assert resolved.slide_size.width > 0


def test_resolved_colours_are_concrete_hex(pptx_path):
    _, resolved = resolve(pptx_path)
    for slide in resolved.slides:
        for element in walk(slide.elements):
            fill = getattr(element, "fill", None)
            if isinstance(fill, m.SolidFill):
                assert fill.color.hex.startswith("#") and len(fill.color.hex) == 7
                assert 0.0 <= fill.color.alpha <= 1.0


def test_theme_fonts_are_expanded(basic_theme):
    """`+mj-lt` / `+mn-lt` must become real typeface names, never leak through."""
    _, resolved = resolve(basic_theme)
    typefaces = [
        run.properties.font_family
        for slide in resolved.slides
        for element in walk(slide.elements)
        if getattr(element, "text_body", None)
        for paragraph in element.text_body.paragraphs
        for run in paragraph.runs
    ]
    assert typefaces, "expected some text runs"
    assert not any(name and name.startswith("+") for name in typefaces)


def test_background_falls_back_through_layout_and_master(product_page):
    _, resolved = resolve(product_page)
    background = resolved.slides[0].background
    assert background is not None and background.fill is not None


def test_empty_placeholders_are_dropped(authoring):
    """An unfilled "Click to add title" placeholder must not render its prompt."""
    _, resolved = resolve(authoring)
    for element in walk(resolved.slides[0].elements):
        body = getattr(element, "text_body", None)
        if body is None or getattr(element, "placeholder_type", None) is None:
            continue
        assert any(run.text for para in body.paragraphs for run in para.runs)


def test_slide_selection_filters_by_number(pptx_path):
    package = OpcPackage.open(str(pptx_path))
    presentation = read_presentation(package)
    resolved = resolve_presentation(package, presentation, slide_numbers=[1])
    assert [slide.slide_number for slide in resolved.slides] == [1]


def test_images_carry_base64_payloads(authoring):
    _, resolved = resolve(authoring)
    images = [e for e in walk(resolved.slides[0].elements) if isinstance(e, m.ImageElement)]
    assert images, "fixture should contain a picture"
    for image in images:
        assert image.image_data and image.mime_type.startswith("image/")


def test_unsupported_graphic_frames_warn_rather_than_vanish(authoring):
    options = ConvertOptions()
    convert_pptx_to_model(authoring, options)
    codes = {warning.code for warning in options.warnings}
    assert "unsupported-graphic-frame" in codes


def test_text_inherits_size_from_the_master_text_styles(basic_theme):
    """No run should end up sizeless: the cascade always reaches a defRPr."""
    _, resolved = resolve(basic_theme)
    runs = [
        run
        for slide in resolved.slides
        for element in walk(slide.elements)
        if getattr(element, "text_body", None)
        for paragraph in element.text_body.paragraphs
        for run in paragraph.runs
        if run.text.strip()
    ]
    assert runs
    assert all(run.properties.font_size for run in runs)

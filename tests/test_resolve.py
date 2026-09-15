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
    """A frame holding something we cannot draw says so instead of vanishing.

    `authoring-integration.pptx` used to be the fixture for this because its chart was
    the unsupported thing; now the chart renders, so the frame has to be a synthetic one.
    """
    from tests.deckbuilder import derive_deck

    frame = (
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='98' name='Embedded movie'/>"
        "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        "<p:xfrm><a:off x='0' y='0'/><a:ext cx='1000000' cy='1000000'/></p:xfrm>"
        "<a:graphic><a:graphicData "
        "uri='http://schemas.openxmlformats.org/presentationml/2006/media'>"
        "<p:media/></a:graphicData></a:graphic></p:graphicFrame>"
    )
    options = ConvertOptions()
    convert_pptx_to_model(derive_deck(authoring, shapes_xml=frame), options)
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


def _titles(resolved, slide_index: int) -> list[m.TextRun]:
    return [
        run
        for element in walk(resolved.slides[slide_index].elements)
        if getattr(element, "text_body", None)
        for paragraph in element.text_body.paragraphs
        for run in paragraph.runs
        if run.text.strip()
    ]


def test_bold_is_inherited_from_the_layout_and_the_master_text_styles(college_template):
    """PowerPoint bolds a title whose only ``b="1"`` is a layout's or the master's.

    Both routes are in this one deck.  Slide 6's title inherits from the master's
    ``p:titleStyle``; slide 2's from ``slideLayout10``'s title placeholder ``a:lstStyle``,
    and that slide's own run says nothing.  PowerPoint's export of both draws them bold
    and embeds ``Arial-BoldMT``.  These used to come out regular, on pptx-glimpse's rule
    that decorations do not cross the placeholder boundary.
    """
    _, resolved = resolve(college_template)
    # Slide 2: bold from the layout's lstStyle.
    assert any(
        run.text.startswith("Presentation Title") and run.properties.bold
        for run in _titles(resolved, 1)
    )
    # Slide 6: bold from the master's titleStyle -- its layout says nothing about weight.
    assert any(
        run.text == "List Title" and run.properties.bold for run in _titles(resolved, 5)
    )
    # Body text under the same master is *not* bold: the master bolds titles only, so a
    # blanket "inherit everything" would show up here.
    assert not any(
        run.properties.bold
        for run in _titles(resolved, 5)
        if run.text.startswith("Lorem")
    )


# -- Table styles ----------------------------------------------------------------------


def table_of(resolved, slide_index=0):
    for element in walk(resolved.slides[slide_index].elements):
        if isinstance(element, m.TableElement):
            return element
    raise AssertionError("no table on that slide")


def test_builtin_table_style_paints_banding_header_and_gridlines(authoring):
    """The deck names "Medium Style 2 - Accent 1" and defines it nowhere.

    PowerPoint never writes a built-in style's definition into the file, so without the
    catalogue this table would render as a bare grid.  The expected colours are the ones
    PowerPoint itself produces for this deck's theme (accent1 = #4472c4).
    """
    _, resolved = resolve(authoring)
    table = table_of(resolved)

    header, body = table.table.rows
    # The header cells carry their own fill, and explicit formatting always wins.
    assert header.cells[0].fill.color.hex == "#dbeafe"
    # The body row has no fill of its own, so the style's first banded row supplies one.
    assert body.cells[0].fill.color.hex == "#cfd5ea"
    # Gridlines come from the style: 1 pt white between cells, 3 pt under the header.
    assert header.cells[0].borders.bottom.fill.color.hex == "#ffffff"
    assert header.cells[0].borders.bottom.width == 38100
    assert body.cells[0].borders.right.width == 12700
    # The header row's text style is part of the region, not just its fill.
    header_run = header.cells[0].text_body.paragraphs[0].runs[0]
    assert header_run.properties.bold is True
    assert header_run.properties.color.hex == "#ffffff"
    # Body text keeps the whole-table style: dark and not bold.
    body_run = body.cells[0].text_body.paragraphs[0].runs[0]
    assert body_run.properties.bold is False
    assert body_run.properties.color.hex == "#000000"


def test_a_table_style_outranks_the_masters_other_text_style(college_template):
    """A cell is not a placeholder, so `p:otherStyle` must not reach it first.

    `real-college-template` defines "Medium Style 2 - Accent 1" in its own
    `tableStyles.xml` with `<a:schemeClr val="lt1"/>` on `a:firstRow`, and its master's
    `p:otherStyle` sets `tx1` -- #151515 in that theme.  PowerPoint inks the header white
    on the #C00000 band; we inked it #151515, because the table style was appended to the
    cascade *below* the master's.  ``authoring-integration`` could not see this: its
    master states no colour there, so the two orders agree.
    """
    _, resolved = resolve(college_template)
    table = table_of(resolved, slide_index=4)
    header = table.table.rows[0].cells[0]
    run = header.text_body.paragraphs[0].runs[0]
    assert run.text == "Column A"
    assert header.fill.color.hex == "#c00000"
    assert run.properties.color.hex == "#ffffff"
    assert run.properties.bold is True
    # The body rows keep `a:wholeTbl`'s dk1, which is this theme's near-black.
    body_run = table.table.rows[1].cells[0].text_body.paragraphs[0].runs[0]
    assert body_run.properties.color.hex == "#151515"


def test_custom_table_style_from_the_deck_is_applied(basic_theme):
    """A Google Slides export defines its own style in `tableStyles.xml`."""
    _, resolved = resolve(basic_theme)
    table = table_of(resolved, slide_index=1)
    cell = table.table.rows[0].cells[0]
    assert cell.borders.bottom.fill.color.hex == "#9e9e9e"
    assert cell.borders.right.fill.color.hex == "#9e9e9e"
    # That style sets no fills at all, so the cells stay unpainted.
    assert cell.fill is None

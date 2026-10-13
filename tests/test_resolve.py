"""Theme resolution and the placeholder / text inheritance cascades."""

from __future__ import annotations

from pathlib import Path

from pptx2svg import ConvertOptions, convert_pptx_to_model, model as m
from pptx2svg.opc import OpcPackage
from pptx2svg.parse.parts import read_presentation
from pptx2svg.resolve import resolve_presentation
from pptx2svg.text.measure import is_cjk


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


# -- The last two built-in styles, and the id nobody knows ------------------------------

#: A table naming a style id, spliced into a fixture so the id under test is the only
#: thing that varies.  ``firstRow``/``bandRow`` are on so the header and the first
#: banded row -- the two regions a style is most visible in -- are both exercised.
TABLE_NAMING_STYLE = (
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="900" name="Probe table"/>'
    '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
    '<p:xfrm><a:off x="300000" y="300000"/><a:ext cx="2400000" cy="900000"/></p:xfrm>'
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
    '<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{style_id}</a:tableStyleId>'
    "</a:tblPr>"
    '<a:tblGrid><a:gridCol w="800000"/><a:gridCol w="800000"/><a:gridCol w="800000"/></a:tblGrid>'
    + (
        '<a:tr h="300000">'
        + (
            "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>"
            '<a:r><a:rPr lang="en-US" sz="800"/><a:t>x</a:t></a:r>'
            "</a:p></a:txBody><a:tcPr/></a:tc>"
        )
        * 3
        + "</a:tr>"
    )
    * 3
    + "</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"
)


def table_naming(authoring, style_id):
    """Resolve ``authoring-integration`` with one extra table naming ``style_id``."""
    from deckbuilder import derive_deck

    options = ConvertOptions()
    deck = derive_deck(
        authoring, shapes_xml=TABLE_NAMING_STYLE.format(style_id=style_id)
    )
    resolved = convert_pptx_to_model(deck, options)
    tables = [
        element
        for element in walk(resolved.slides[0].elements)
        if isinstance(element, m.TableElement)
    ]
    # The spliced frame lands last in the shape tree, after the fixture's own table.
    return tables[-1], options.warnings


def test_light_style_1_accent_4_resolves(authoring):
    """The GUID was the only thing missing; the measurement was never the problem.

    ``{D27102A9-...}`` came from a published list and was confirmed by rendering it
    through PowerPoint: it does not measure like the unrecognised-id fallback, and it
    measures like the rest of the Light Style 1 family with accent 4 in the accent slot
    -- hairlines above and below, 20 %-alpha banding, bold edges, no fills.  This
    theme's accent 4 is ``#ffc000``.
    """
    table, _ = table_naming(authoring, "{D27102A9-8310-4765-A935-A1911B00CA55}")

    header, band = table.table.rows[0].cells[0], table.table.rows[1].cells[0]
    assert header.fill is None  # the family paints no fills at all
    assert band.fill.color.hex == "#ffc000"
    assert band.fill.color.alpha == 0.2
    assert header.borders.bottom.fill.color.hex == "#ffc000"
    assert header.borders.bottom.width == 12700  # 1 pt
    assert header.text_body.paragraphs[0].runs[0].properties.bold is True


def test_medium_style_1_resolves(authoring):
    """``{793D81CF-...}`` was read out of PowerPoint's own executable.

    The binary carries fifteen brace-delimited GUIDs; fourteen are table styles already
    catalogued, and this was the fifteenth.  Measured, it is the Medium Style 1 family's
    shape with ``dk1`` where the accents sit, which is what an accent-less base variant
    looks like -- a black header with white bold text, a light banded row, and a 3 pt
    rule above the footer.
    """
    table, _ = table_naming(authoring, "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}")

    header, band = table.table.rows[0].cells[0], table.table.rows[1].cells[0]
    assert header.fill.color.hex == "#000000"
    header_run = header.text_body.paragraphs[0].runs[0]
    assert header_run.properties.color.hex == "#ffffff"
    assert header_run.properties.bold is True
    # `dk1` tinted 20 %, which this theme's black resolves to as a light grey.
    assert band.fill.color.hex == "#e7e7e7"
    # `wholeTbl` fills white and rules every inside edge in dk1 at 1 pt.
    assert band.borders.top.fill.color.hex == "#000000"
    assert band.borders.top.width == 12700


def test_an_unresolvable_table_style_id_warns(authoring):
    """A style id in neither the deck nor the catalogue must not fail silently.

    PowerPoint draws an unrecognised id unstyled too, so the render is defensible -- but
    an unstyled table is pixel-identical to a table whose style genuinely carries
    nothing, so without a warning there is no way to tell a correct render from one that
    lost every band, rule and header colour.  Same failure mode as `font-substituted`.
    """
    table, warnings = table_naming(authoring, "{DEADBEEF-0000-0000-0000-000000000000}")

    assert table.table.rows[0].cells[0].fill is None
    warning = next(w for w in warnings if w.code == "table-style-unknown")
    assert "Probe table" in warning.message
    assert "{DEADBEEF-0000-0000-0000-000000000000}" in warning.message
    assert warning.slide_number == 1


def test_a_known_table_style_id_raises_no_warning(authoring):
    """The guard against a warning that fires on every deck in the corpus."""
    _, warnings = table_naming(authoring, "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}")
    assert not [w for w in warnings if w.code == "table-style-unknown"]


# -- The East Asian typeface cascade -----------------------------------------------------


def test_a_run_naming_no_east_asian_face_is_drawn_in_ms_gothic(financial):
    """`<a:ea typeface=""/>` in the theme means "no face named", not "the empty face" --
    and the theme's `<a:font script="Jpan"/>` does not stand in for it.

    `real-financial-report.pptx`'s table cells state no typeface anywhere and no Japanese
    `lang`, so they take the theme: Calibri for Latin, and for the rest the `<a:ea>` slot,
    which is empty.  PowerPoint's export draws their Japanese in **MS-Gothic** -- not the
    游ゴシック of the theme's `Jpan` entry, which the same deck's chart labels are drawn in
    -- because Calibri is a sans serif face (pptx2svg.resolve.east_asian; measured with
    `tools/make_font_resolution_probe.py`).  Before, a run with no `<a:ea>` of its own was
    measured through the Latin table, then through the `Jpan` entry: neither is what
    PowerPoint draws.
    """
    _, resolved = resolve(financial)
    runs = [
        run
        for element in walk(resolved.slides[3].elements)
        if isinstance(element, m.TableElement)
        for row in element.table.rows
        for cell in row.cells
        if cell.text_body is not None
        for paragraph in cell.text_body.paragraphs
        for run in paragraph.runs
        if any(is_cjk(ord(char)) for char in run.text)
    ]
    assert runs
    assert {run.properties.font_family for run in runs} == {"Calibri"}
    assert {run.properties.font_family_ea for run in runs} == {"MS Gothic"}


def test_a_latin_face_named_as_the_east_asian_one_decides_gothic_or_mincho(basic_theme):
    """`real-basic-theme.pptx` writes `<a:ea typeface="Raleway"/>` on its master's title
    style and `Lato` on its body.

    Neither has kana, and PowerPoint's own export of the deck drew the Japanese in MS
    Mincho (the titles) and MS Gothic (the body): the face named decides between the two
    by its PANOSE -- Lato's says sans serif, Raleway's says nothing at all -- rather than
    giving way to the theme's `Jpan` entry, ＭＳ Ｐゴシック, which it never reaches.
    """
    _, resolved = resolve(basic_theme)
    faces = {}
    for slide in resolved.slides:
        for element in walk(slide.elements):
            if getattr(element, "text_body", None) is None:
                continue
            for paragraph in element.text_body.paragraphs:
                for run in paragraph.runs:
                    if any(is_cjk(ord(char)) for char in run.text):
                        faces.setdefault(run.properties.font_family, set()).add(run.properties.font_family_ea)
    assert faces["Raleway"] == {"MS Mincho"}
    assert faces["Lato"] == {"MS Gothic"}


# -- Round 3: what PowerPoint draws that the catalogue and the empty-placeholder rule missed --


def test_light_style_2_draws_its_outline_and_band_edges(authoring):
    """Measured in PowerPoint 16: every Light Style 2 style rules the table's outline and the
    top and bottom of each odd banded row in its colour, 1 pt (the generated catalogue had
    only the header fill and the total-row rule).  This theme's accent 3 is ``#a5a5a5``."""
    table, _ = table_naming(authoring, "{F2DE63D5-997A-4646-A377-4702673A728D}")
    rows = table.table.rows
    header, band, plain = rows[0].cells[0], rows[1].cells[0], rows[2].cells[1]
    assert header.fill.color.hex == "#a5a5a5"
    for edge in (header.borders.left, header.borders.top, band.borders.top,
                 band.borders.bottom, plain.borders.bottom):
        assert edge.width == 12700 and edge.fill.color.hex == "#a5a5a5"
    # Between two cells of a row there is no line: the style rules no gridlines.
    assert plain.borders.left is None or plain.borders.left.fill is None


#: A deck's own table styles that leave edges unstated, as a template's can.
OWN_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">'
    '<a:tblStyle styleId="{8F44A2D9-6A0B-4E0C-9C51-2C4A0F2A7B00}" styleName="Rows only">'
    '<a:wholeTbl><a:tcTxStyle><a:schemeClr val="tx1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>'
    '<a:insideH><a:ln w="6350"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'
    '</a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV></a:tcBdr>'
    '<a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle></a:tblStyleLst>')


def test_a_decks_own_style_draws_its_unstated_edges_black(basic_theme):
    """Measured in PowerPoint 16: an edge a deck's own style does not state is a 1 pt black
    line (black whatever the theme's dk1); one it states as no line is not drawn."""
    from deckbuilder import derive_deck

    options = ConvertOptions()
    deck = derive_deck(basic_theme, parts={"ppt/tableStyles.xml": OWN_STYLES.encode()},
                       shapes_xml=TABLE_NAMING_STYLE.format(
                           style_id="{8F44A2D9-6A0B-4E0C-9C51-2C4A0F2A7B00}"))
    resolved = convert_pptx_to_model(deck, options)
    table = [e for e in walk(resolved.slides[0].elements) if isinstance(e, m.TableElement)][-1]
    middle = table.table.rows[1].cells[1]
    corner = table.table.rows[0].cells[0]
    assert middle.borders.top.fill.color.hex == "#ff0000"           # stated: insideH
    assert middle.borders.left is None or middle.borders.left.fill is None   # stated: none
    for edge in (corner.borders.top, corner.borders.left):           # unstated: the outline
        assert edge.width == 12700 and edge.fill.color.hex == "#000000"


def test_an_empty_placeholder_draws_its_box_but_no_prompt(authoring):
    """Measured in PowerPoint 16: an empty placeholder's fill is drawn in the export (its
    prompt is not); one with neither fill nor outline draws nothing, as before."""
    from deckbuilder import derive_deck

    def placeholder(number, fill):
        return (f'<p:sp><p:nvSpPr><p:cNvPr id="{number}" name="Empty {number}"/><p:cNvSpPr/>'
                f'<p:nvPr><p:ph type="body" idx="{number}"/></p:nvPr></p:nvSpPr><p:spPr>'
                '<a:xfrm><a:off x="300000" y="300000"/><a:ext cx="900000" cy="600000"/>'
                f'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}</p:spPr>'
                '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p>'
                '</p:txBody></p:sp>')

    deck = derive_deck(authoring, shapes_xml=placeholder(
        901, '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>') + placeholder(902, ""))
    resolved = convert_pptx_to_model(deck, ConvertOptions())
    shapes = {e.element_id: e for e in walk(resolved.slides[0].elements)
              if isinstance(e, m.ShapeElement)}
    box = next(e for key, e in shapes.items() if key and key.endswith(".901"))
    assert box.fill.color.hex == "#ff0000" and box.text_body is None
    assert not any(key and key.endswith(".902") for key in shapes)


# -- Gaps: "Themed Style 1" and "2" draw on the theme's format scheme ---------------------


def themed_table(source, style_id, flags):
    """``source`` with one 3 x 3 table in ``style_id`` and the ``tblPr`` ``flags`` given."""
    from deckbuilder import derive_deck

    attributes = "".join(f' {flag}="1"' for flag in flags)
    frame = TABLE_NAMING_STYLE.format(style_id=style_id).replace(
        '<a:tblPr firstRow="1" bandRow="1">', f"<a:tblPr{attributes}>")
    resolved = convert_pptx_to_model(derive_deck(source, shapes_xml=frame), ConvertOptions())
    return [e for e in walk(resolved.slides[0].elements) if isinstance(e, m.TableElement)][-1]


THEMED_1_ACCENT_1 = "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}"
THEMED_2_ACCENT_1 = "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}"
SAMPLE = Path(__file__).parent / "fixtures" / "sample.pptx"


def test_themed_style_2_draws_on_a_theme_fill_and_effect_style(financial):
    """PowerPoint's definition: the table on fill style 3 and effect style 3 in accent 1,
    text ``lt1``, the header ruled with line style 3 and the first column, the total row
    and the corners with line style 2 in ``lt1``.  Measured on PowerPoint 16's export under
    this (2013 Office) theme: a gradient across the table, a shadow under it, 1.5 pt and
    1 pt white rules, and no rule into or along the total row's corner cells."""
    table = themed_table(financial, THEMED_2_ACCENT_1, ("firstRow", "lastRow", "firstCol",
                                                        "bandRow"))
    data, rows = table.table, table.table.rows
    assert isinstance(data.background, m.GradientFill)
    assert data.effects is not None and data.effects.outer_shadow is not None
    header, band = rows[0].cells[1], rows[1].cells[1]
    corner, total = rows[2].cells[0], rows[2].cells[1]
    for cell in (header, total):                               # a:noFill over the gradient
        assert cell.fill is None or isinstance(cell.fill, m.NoFill)
    assert band.fill.color.hex == "#ffffff" and band.fill.color.alpha == 0.2
    assert header.text_body.paragraphs[0].runs[0].properties.color.hex == "#ffffff"
    assert (header.borders.bottom.width, header.borders.bottom.fill.color.hex) == (
        19050, "#ffffff")
    assert band.borders.top == header.borders.bottom           # one line on the shared edge
    assert total.borders.top.width == 12700
    assert rows[1].cells[0].borders.right.width == 12700       # firstCol's right rule
    assert corner.borders.top is None and corner.borders.right is None   # swCell: none


def test_themed_style_1_rules_every_edge_with_a_theme_line_style(financial):
    """Fill style 2 under the table, every gridline line style 1 in the accent, the header
    solid accent with a line style 2 ``lt1`` rule under it.  PowerPoint draws that white
    rule over the body's 0.5 pt accent ``insideH`` on the edge the two share (it draws the
    region of higher precedence there), and no shadow: this theme's effect style 1 is
    empty."""
    table = themed_table(financial, THEMED_1_ACCENT_1, ("firstRow", "bandRow"))
    data, rows = table.table, table.table.rows
    assert isinstance(data.background, m.GradientFill)
    assert data.effects is None or data.effects.is_empty()
    header, body = rows[0].cells[0], rows[1].cells[0]
    assert header.fill.color.hex == "#4472c4"
    assert (body.borders.top.width, body.borders.top.fill.color.hex) == (12700, "#ffffff")
    assert (body.borders.right.width, body.borders.right.fill.color.hex) == (6350, "#4472c4")
    assert body.fill.color.alpha == 0.4                         # band1H over the gradient


def test_themed_styles_take_a_2007_themes_shadow_and_no_bevel():
    """The 2007 Office theme's effect styles are all shadows, the third with a bevel: both
    families draw a shadow under the table there, and PowerPoint draws no bevel."""
    for style_id in (THEMED_1_ACCENT_1, THEMED_2_ACCENT_1):
        table = themed_table(SAMPLE, style_id, ("firstRow", "bandRow"))
        assert isinstance(table.table.background, m.GradientFill)
        assert table.table.effects.outer_shadow is not None

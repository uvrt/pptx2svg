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


def test_shapes_carry_their_source_identity(pptx_path):
    """Every drawn element is traceable back to the shape it came from.

    Downstream editors need to map a rendered group to a shape in the deck; without an id on
    the output the only correspondence is document order, which breaks as soon as template
    shapes are flattened in or empty placeholders are dropped.
    """
    for document in convert_pptx_to_svg(pptx_path):
        identifiers = re.findall(r'data-pptx-id="([^"]*)"', document)
        assert identifiers, "no element carried an id"
        for identifier in identifiers:
            # "<sldId>.<cNvPr id>" for slide shapes, "lay:"/"mst:" for inherited ones.
            assert re.fullmatch(r"(lay:|mst:|\d+\.)[^\s\"]+", identifier), identifier
        # An id alone is not unique in real decks, so a path always accompanies it.
        assert len(re.findall(r'data-pptx-path="[^"]*"', document)) == len(identifiers)


def test_identity_is_unique_per_slide(pptx_path):
    """The (id, path) pair addresses exactly one element."""
    for document in convert_pptx_to_svg(pptx_path):
        pairs = re.findall(
            r'data-pptx-id="([^"]*)"\s+data-pptx-path="([^"]*)"', document
        )
        assert len(pairs) == len(set(pairs))

# -- Run and body properties the renderer used to drop -----------------------------------


def text_svg(body: m.TextBody, width: float = 6000000, height: float = 800000) -> str:
    from pptx2svg.render.text import render_text_body

    context = RenderContext(
        measurer=DefaultTextMeasurer(), font_mapping={}, jpan_fallback_font=None
    )
    return render_text_body(
        body, m.Transform(extent_width=width, extent_height=height), context
    )


def one_run(text: str, **properties) -> m.TextBody:
    return m.TextBody(
        paragraphs=[m.Paragraph(runs=[m.TextRun(text, m.RunProperties(**properties))])]
    )


def test_highlight_paints_a_rectangle_behind_the_run():
    body = m.TextBody(paragraphs=[m.Paragraph(runs=[
        m.TextRun("plain ", m.RunProperties(font_size=18)),
        m.TextRun("lit", m.RunProperties(font_size=18, highlight=m.ResolvedColor(hex="#ffff00"))),
    ])])
    svg = text_svg(body)
    # SVG has no text background, so the highlight is a rect -- and it has to come first
    # so it cannot paint over the glyphs.
    assert svg.index("<rect") < svg.index("<text")
    rect = re.search(r"<rect [^>]*/>", svg).group()
    assert 'fill="#ffff00"' in rect
    # It starts where the unhighlighted run ends, not at the left inset.
    assert float(re.search(r'x="([\d.]+)"', rect).group(1)) > 60


def test_text_without_a_highlight_emits_no_rectangle():
    assert "<rect" not in text_svg(one_run("plain", font_size=18))


def test_underline_style_is_carried_through_to_the_attribute():
    svg = text_svg(one_run("x", font_size=18, underline=True, underline_style="dbl"))
    assert 'text-decoration="underline"' in svg
    assert 'text-decoration-style="double"' in svg
    # A plain single underline stays plain: no redundant attribute.
    plain = text_svg(one_run("x", font_size=18, underline=True))
    assert "text-decoration-style" not in plain


def test_complex_script_font_joins_the_family_stack():
    svg = text_svg(one_run("x", font_size=18, font_family="Calibri", font_family_cs="Arial"))
    family = re.search(r'font-family="([^"]+)"', svg).group(1)
    assert "Calibri" in family and "Arial" in family
    assert family.index("Calibri") < family.index("Arial")


def test_body_rotation_turns_the_text_about_the_box_centre():
    body = one_run("x", font_size=18)
    body.body_properties.rotation = 45
    svg = text_svg(body, width=2000000, height=1000000)
    assert svg.startswith('<g transform="rotate(45, 104.987, 52.493)">')


def test_tabs_jump_to_the_default_one_inch_grid():
    svg = text_svg(one_run("a\tb\tc", font_size=18))
    # 1 inch is 96 px, measured from the 0.1 inch left inset.
    assert 'x="105.6"' in svg and 'x="201.6"' in svg


def test_an_explicit_tab_stop_wins_and_carries_its_alignment():
    paragraph = m.Paragraph(
        runs=[m.TextRun("name\tvalue", m.RunProperties(font_size=18))],
        properties=m.ParagraphProperties(
            tab_stops=[m.TabStop(position=2743200, alignment="r")]
        ),
    )
    svg = text_svg(m.TextBody(paragraphs=[paragraph]))
    assert 'x="297.6" text-anchor="end"' in svg


def test_tabs_are_left_alone_in_a_centred_paragraph():
    paragraph = m.Paragraph(
        runs=[m.TextRun("a\tb", m.RunProperties(font_size=18))],
        properties=m.ParagraphProperties(alignment="ctr"),
    )
    svg = text_svg(m.TextBody(paragraphs=[paragraph]))
    # One chunk, anchored once: the paragraph's own alignment decides the position.
    assert svg.count("text-anchor=") == 1


# -- Images ------------------------------------------------------------------------------


def image_svg(**kwargs) -> str:
    from pptx2svg.render.shape import render_image

    context = RenderContext(
        measurer=DefaultTextMeasurer(), font_mapping={}, jpan_fallback_font=None
    )
    element = m.ImageElement(
        transform=m.Transform(extent_width=1000000, extent_height=500000),
        image_data="AAAA",
        mime_type="image/png",
        **kwargs,
    )
    svg = render_image(element, context)
    return "".join(context.defs) + svg if hasattr(context, "defs") else svg


def test_stretch_fill_rect_insets_the_bitmap():
    svg = image_svg(stretch=m.StretchFillRect(left=0.1, top=0.2, right=0.1, bottom=0.0))
    image = re.search(r"<image [^>]*/>", svg).group()
    assert 'x="10.499"' in image and 'y="10.499"' in image
    assert 'width="83.99"' in image and 'height="41.995"' in image


def test_an_all_zero_stretch_rect_still_fills_the_frame():
    svg = image_svg(stretch=m.StretchFillRect())
    image = re.search(r"<image [^>]*/>", svg).group()
    assert 'x="0"' in image and 'width="104.987"' in image


def test_tile_repeats_the_bitmap_through_a_pattern():
    svg = image_svg(tile=m.TileInfo(sx=0.5, sy=0.25, tx=91440, ty=0))
    assert "<rect" in svg and 'fill="url(#' in svg


# -- Multi-column text -------------------------------------------------------------------


def many_paragraphs(count: int) -> list[m.Paragraph]:
    return [
        m.Paragraph(runs=[m.TextRun(
            f"Paragraph number {i} with enough words in it to wrap at least once.",
            m.RunProperties(font_size=12),
        )])
        for i in range(1, count + 1)
    ]


def test_columns_flow_into_one_text_element_each():
    body = m.TextBody(
        paragraphs=many_paragraphs(6),
        body_properties=m.BodyProperties(num_col=2),
    )
    svg = text_svg(body, width=5000000, height=1400000)
    assert svg.count("<text") == 2

    first, second = svg.split("</text>")[:2]
    # The second column sits a column-width to the right of the first...
    left = sorted(set(re.findall(r'<tspan x="([\d.]+)"', first)))
    right = sorted(set(re.findall(r'<tspan x="([\d.]+)"', second)))
    assert float(right[0]) > float(left[0]) + 200
    # ...and holds later paragraphs, rather than the same text twice.
    assert "number 1" in first and "number 1" not in second


def test_a_single_column_body_is_unchanged():
    body = m.TextBody(
        paragraphs=many_paragraphs(6),
        body_properties=m.BodyProperties(num_col=1),
    )
    assert text_svg(body, width=5000000, height=1400000).count("<text") == 1


def test_numbering_carries_on_across_a_column_break():
    paragraphs = [
        m.Paragraph(
            runs=[m.TextRun(f"item {i} with enough words to take a whole line", m.RunProperties(font_size=12))],
            properties=m.ParagraphProperties(bullet=m.AutoNumBullet(scheme="arabicPeriod")),
        )
        for i in range(1, 7)
    ]
    body = m.TextBody(
        paragraphs=paragraphs, body_properties=m.BodyProperties(num_col=2)
    )
    svg = text_svg(body, width=5000000, height=700000)
    numbers = [n for n in ("1.", "2.", "3.", "4.", "5.", "6.") if f">{n}<" in svg]
    assert numbers == ["1.", "2.", "3.", "4.", "5.", "6."]


# -- Model coverage ----------------------------------------------------------------------

#: Render-model fields the renderer deliberately does not read, and why.  Everything
#: else must be read somewhere under `render/` or `text/`: this test exists because the
#: "parsed but not rendered" category grew silently once already, and an audit only
#: catches it if the audit runs every time.
#:
#: Adding an entry here is a decision, not a formality -- it says "this field is not
#: meant for the renderer", not "this is not implemented yet".  A genuinely missing
#: feature belongs in the roadmap, not in this list.
UNRENDERED_FIELDS = {
    # Theme tables.  The resolver consumes these and hands the renderer concrete
    # colours and typefaces; they are on the model for callers of
    # `convert_pptx_to_model`, not for drawing.
    **{f"ColorScheme.{name}": "resolved before the renderer sees it" for name in (
        "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4",
        "accent5", "accent6", "hlink", "folHlink",
    )},
    **{f"FontScheme.{name}": "resolved before the renderer sees it" for name in (
        "major_font", "minor_font", "major_font_ea", "minor_font_ea",
        "major_font_cs", "minor_font_cs", "major_font_jpan", "minor_font_jpan",
    )},

    # Metadata that describes a shape rather than drawing it.
    "ShapeElement.alt_text": "accessibility metadata, not geometry",
    "ImageElement.alt_text": "accessibility metadata, not geometry",
    "ConnectorElement.alt_text": "accessibility metadata, not geometry",
    "GroupElement.alt_text": "accessibility metadata, not geometry",
    "TableElement.alt_text": "accessibility metadata, not geometry",
    "Hyperlink.tooltip": "no SVG equivalent short of a <title> child",
    "ShapeElement.placeholder_type": "used by the resolver's inheritance, not drawn",
    "ShapeElement.placeholder_idx": "used by the resolver's inheritance, not drawn",
    "Slide.slide_number": "identifies the slide; the caller decides what to do with it",
    "Slide.show_master_sp": "the resolver has already applied it to the element list",

    # Genuinely not implemented.  These are roadmap items, listed so the gap is
    # visible rather than merely absent.
    "BlipEffects.clr_change": "a:clrChange has no clean SVG filter equivalent",
    "ClrChangeEffect.clr_from": "a:clrChange is not rendered",
    "ClrChangeEffect.clr_to": "a:clrChange is not rendered",
    "BlurEffect.grow": "we always let the blur grow past the shape's bounds",
    "OuterShadow.rotate_with_shape": "shadows are emitted inside the shape's own transform",
    "ImageFillTile.flip": "SVG patterns cannot mirror alternate tiles",
    "ImageFillTile.align": "tile origin comes from tx/ty alone",
    "TileInfo.flip": "SVG patterns cannot mirror alternate tiles",
    "TileInfo.align": "tile origin comes from tx/ty alone",
}


def test_the_renderer_reads_every_field_the_model_carries():
    import dataclasses
    from pathlib import Path

    from pptx2svg import model

    root = Path(__file__).resolve().parent.parent / "src/pptx2svg"
    sources = "".join(
        path.read_text()
        for directory in ("render", "text")
        for path in sorted((root / directory).glob("*.py"))
    )

    unread = []
    for name in sorted(dir(model)):
        obj = getattr(model, name)
        if not dataclasses.is_dataclass(obj):
            continue
        for field in dataclasses.fields(obj):
            key = f"{name}.{field.name}"
            if field.name == "type" or key in UNRENDERED_FIELDS:
                continue
            # Attribute access, or a dynamic read: `render_element` reaches fields common
            # to the element union with getattr, since the union has no shared base.
            attribute = re.escape(field.name)
            if not re.search(
                rf"\.{attribute}\b|getattr\([^)]*[\"']{attribute}[\"']", sources
            ):
                unread.append(key)

    assert not unread, (
        "these render-model fields are never read by the renderer:\n  "
        + "\n  ".join(unread)
        + "\n\nEither render them, or add them to UNRENDERED_FIELDS with the reason."
    )


def test_the_unrendered_list_has_no_stale_entries():
    """An entry that no longer names a real field hides a regression behind itself."""
    import dataclasses

    from pptx2svg import model

    known = {
        f"{name}.{field.name}"
        for name in dir(model)
        if dataclasses.is_dataclass(getattr(model, name))
        for field in dataclasses.fields(getattr(model, name))
    }
    assert not (set(UNRENDERED_FIELDS) - known)


# -- Tables --------------------------------------------------------------------------


def cell_text(text: str) -> m.TextBody:
    return m.TextBody(
        paragraphs=[m.Paragraph(runs=[m.TextRun(text, m.RunProperties(font_size=18))])]
    )


def test_a_row_grows_to_fit_text_that_does_not_fit_its_stated_height():
    """`a:tr@h` is a minimum, not an exact height."""
    from pptx2svg.render.shape import render_table

    marker = m.CellBorders(
        top=m.Outline(fill=m.SolidFill(color=m.ResolvedColor(hex="#ff0000")))
    )
    rows = [
        m.TableRow(height=200000, cells=[m.TableCell(text_body=cell_text("short"))]),
        m.TableRow(height=200000, cells=[m.TableCell(text_body=cell_text(
            "a much longer run of text that has to wrap over several lines here"
        ))]),
        m.TableRow(height=200000, cells=[
            m.TableCell(text_body=cell_text("last"), borders=marker)
        ]),
    ]
    table = m.TableElement(
        transform=m.Transform(extent_width=1500000, extent_height=600000),
        table=m.TableData(rows=rows, columns=[m.TableColumn(width=1500000)]),
    )
    context = RenderContext(
        measurer=DefaultTextMeasurer(), font_mapping={}, jpan_fallback_font=None
    )
    svg = render_table(table, context)

    # The marked border sits below two rows.  Taking `h` literally would put it at
    # 2 x 21 px; the middle row has to have grown well past that.
    y = float(re.search(r'<line x1="0" y1="([\d.]+)"', svg).group(1))
    assert y > 100


def test_rows_that_already_fit_keep_their_stated_height():
    from pptx2svg.render.shape import render_table

    marker = m.CellBorders(
        top=m.Outline(fill=m.SolidFill(color=m.ResolvedColor(hex="#ff0000")))
    )
    rows = [
        m.TableRow(height=900000, cells=[m.TableCell(text_body=cell_text("a"))]),
        m.TableRow(height=900000, cells=[m.TableCell(text_body=cell_text("b"), borders=marker)]),
    ]
    table = m.TableElement(
        transform=m.Transform(extent_width=3000000, extent_height=1800000),
        table=m.TableData(rows=rows, columns=[m.TableColumn(width=3000000)]),
    )
    context = RenderContext(
        measurer=DefaultTextMeasurer(), font_mapping={}, jpan_fallback_font=None
    )
    svg = render_table(table, context)
    y = float(re.search(r'<line x1="0" y1="([\d.]+)"', svg).group(1))
    assert abs(y - 94.49) < 0.5

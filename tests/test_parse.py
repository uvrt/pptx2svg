"""Reading OOXML into the source model."""

from __future__ import annotations

from xml.etree.ElementTree import fromstring

import pytest

from pptx2svg.opc import OpcPackage, normalize_part_path, rels_path_for, resolve_target
from pptx2svg.parse import source as s
from pptx2svg.parse.drawing import parse_custom_geometry, parse_fill, parse_geometry, parse_transform
from pptx2svg.parse.parts import read_presentation
from pptx2svg.parse.text import parse_text_body

A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def xml(markup: str):
    return fromstring(markup)


# -- OPC path handling -----------------------------------------------------------------


def test_rels_path_for_a_slide():
    assert rels_path_for("ppt/slides/slide1.xml") == "ppt/slides/_rels/slide1.xml.rels"


def test_relationship_targets_resolve_relative_to_the_part():
    assert resolve_target("ppt/slides", "../media/image1.png") == "ppt/media/image1.png"


def test_absolute_relationship_targets_drop_the_leading_slash():
    assert resolve_target("ppt/slides", "/ppt/media/image1.png") == "ppt/media/image1.png"


def test_normalize_part_path_handles_backslashes():
    assert normalize_part_path("ppt\\slides\\slide1.xml") == "ppt/slides/slide1.xml"


# -- DrawingML -------------------------------------------------------------------------


def test_parse_transform_reads_offset_extent_rotation_and_flip():
    node = xml(
        f'<a:spPr {A}><a:xfrm rot="2700000" flipH="1">'
        '<a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></a:spPr>'
    )
    transform = parse_transform(node)
    assert (transform.offset_x, transform.offset_y) == (100, 200)
    assert (transform.width, transform.height) == (300, 400)
    assert transform.rotation == 2700000
    assert transform.flip_horizontal is True
    assert transform.flip_vertical is False


def test_transform_without_offset_or_extent_is_not_a_transform():
    assert parse_transform(xml(f'<a:spPr {A}><a:xfrm/></a:spPr>')) is None


def test_parse_solid_fill_keeps_the_scheme_reference_unresolved():
    node = xml(f'<a:spPr {A}><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:spPr>')
    fill = parse_fill(node)
    assert isinstance(fill, s.SourceSolidFill)
    assert isinstance(fill.color, s.SchemeColor)
    assert fill.color.scheme == "accent1"


def test_parse_no_fill():
    assert isinstance(parse_fill(xml(f'<a:spPr {A}><a:noFill/></a:spPr>')), s.SourceNoFill)


def test_parse_gradient_fill_reads_stops_and_angle():
    node = xml(
        f'<a:spPr {A}><a:gradFill><a:gsLst>'
        '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'
        '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>'
        '</a:gsLst><a:lin ang="5400000"/></a:gradFill></a:spPr>'
    )
    fill = parse_fill(node)
    assert isinstance(fill, s.SourceGradientFill)
    assert [stop.position for stop in fill.stops] == [0.0, 1.0]
    assert fill.angle == 5400000
    assert fill.gradient_type == "linear"


def test_gradient_with_a_path_element_is_radial():
    node = xml(
        f'<a:spPr {A}><a:gradFill><a:gsLst>'
        '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst>'
        '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'
        "</a:gradFill></a:spPr>"
    )
    fill = parse_fill(node)
    assert fill.gradient_type == "radial"
    assert fill.center_x == pytest.approx(0.5)


def test_parse_preset_geometry_with_adjustments():
    node = xml(
        f'<a:spPr {A}><a:prstGeom prst="roundRect">'
        '<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom></a:spPr>'
    )
    geometry = parse_geometry(node)
    assert geometry.preset == "roundRect"
    assert geometry.adjust_values == {"adj": 25000.0}


def test_custom_geometry_becomes_svg_path_data():
    node = xml(
        f'<a:custGeom {A}><a:pathLst><a:path w="100" h="100">'
        '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>'
        '<a:lnTo><a:pt x="100" y="0"/></a:lnTo>'
        '<a:cubicBezTo><a:pt x="100" y="50"/><a:pt x="50" y="100"/><a:pt x="0" y="100"/></a:cubicBezTo>'
        "<a:close/></a:path></a:pathLst></a:custGeom>"
    )
    paths = parse_custom_geometry(node)
    assert len(paths) == 1
    assert paths[0].commands == "M 0 0 L 100 0 C 100 50, 50 100, 0 100 Z"


def test_custom_geometry_evaluates_guide_formulas():
    node = xml(
        f'<a:custGeom {A}>'
        '<a:gdLst><a:gd name="half" fmla="*/ w 1 2"/></a:gdLst>'
        '<a:pathLst><a:path w="200" h="100">'
        '<a:moveTo><a:pt x="half" y="0"/></a:moveTo>'
        '<a:lnTo><a:pt x="w" y="h"/></a:lnTo>'
        "</a:path></a:pathLst></a:custGeom>"
    )
    paths = parse_custom_geometry(node)
    assert paths[0].commands == "M 100 0 L 200 100"


# -- Text ------------------------------------------------------------------------------


def test_runs_and_breaks_keep_document_order():
    node = xml(
        f'<p:txBody {P} {A}><a:p>'
        "<a:r><a:t>one</a:t></a:r><a:br/><a:r><a:t>two</a:t></a:r>"
        "</a:p></p:txBody>"
    )
    body = parse_text_body(node)
    assert [run.text for run in body.paragraphs[0].runs] == ["one", "\n", "two"]


def test_run_properties_convert_hundredth_points_to_points():
    node = xml(
        f'<p:txBody {P} {A}><a:p><a:r>'
        '<a:rPr sz="1800" b="1" i="0"><a:latin typeface="Calibri"/></a:rPr>'
        "<a:t>x</a:t></a:r></a:p></p:txBody>"
    )
    properties = parse_text_body(node).paragraphs[0].runs[0].properties
    assert properties.font_size == 18.0
    assert properties.bold is True
    # `i="0"` is an explicit "not italic", which must not read as "inherit".
    assert properties.italic is False


def test_absent_run_properties_stay_none_so_they_can_inherit():
    node = xml(f'<p:txBody {P} {A}><a:p><a:r><a:rPr sz="1200"/><a:t>x</a:t></a:r></a:p></p:txBody>')
    properties = parse_text_body(node).paragraphs[0].runs[0].properties
    assert properties.font_size == 12.0
    assert properties.bold is None
    assert properties.typeface is None


def test_bullet_char_reference_is_decoded():
    node = xml(
        f'<p:txBody {P} {A}><a:p><a:pPr><a:buChar char="&#38;#x2022;"/></a:pPr>'
        "<a:r><a:t>x</a:t></a:r></a:p></p:txBody>"
    )
    bullet = parse_text_body(node).paragraphs[0].properties.bullet
    assert bullet.char == "•"


def test_body_properties_read_autofit_and_margins():
    node = xml(
        f'<p:txBody {P} {A}><a:bodyPr lIns="0" anchor="ctr" wrap="none">'
        '<a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>'
        "<a:p/></p:txBody>"
    )
    properties = parse_text_body(node).properties
    assert properties.margin_left == 0
    assert properties.anchor == "ctr"
    assert properties.wrap == "none"
    assert properties.auto_fit == "normAutofit"
    assert properties.font_scale == 0.625
    assert properties.ln_spc_reduction == 0.2


# -- Table styles ----------------------------------------------------------------------


TABLE_STYLE_XML = f"""
<a:tblStyleLst {A} def="{{AAAA0000-0000-0000-0000-000000000001}}">
  <a:tblStyle styleId="{{AAAA0000-0000-0000-0000-000000000001}}" styleName="Example">
    <a:wholeTbl>
      <a:tcTxStyle b="on"><a:fontRef idx="minor"/><a:srgbClr val="112233"/></a:tcTxStyle>
      <a:tcStyle>
        <a:tcBdr>
          <a:insideH><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></a:insideH>
          <a:left><a:ln w="25400"><a:noFill/></a:ln></a:left>
        </a:tcBdr>
        <a:fill><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:fill>
      </a:tcStyle>
    </a:wholeTbl>
    <a:firstRow>
      <a:tcTxStyle b="def" i="on"/>
      <a:tcStyle><a:tcBdr/></a:tcStyle>
    </a:firstRow>
  </a:tblStyle>
</a:tblStyleLst>
"""


def test_table_style_reads_fills_borders_and_text():
    from pptx2svg.parse.parts import parse_table_style

    style = parse_table_style(xml(TABLE_STYLE_XML)[0])
    assert style.name == "Example"

    whole = style.whole_table
    assert isinstance(whole.fill, s.SourceSolidFill)
    assert whole.fill.color.scheme == "accent2"
    assert whole.border_inside_h.width == 12700
    # An explicit no-fill line is not the same as an absent one: it clears the border.
    assert isinstance(whole.border_left.fill, s.SourceNoFill)
    assert whole.border_top is None
    assert whole.text.bold is True
    assert whole.text.color.hex == "112233"
    # `a:fontRef idx="minor"` is handed on as the placeholder the text resolver expands.
    assert whole.text.typeface == "+mn-lt"

    # `b="def"` means "inherit", which has to stay None rather than becoming False.
    assert style.first_row.text.bold is None
    assert style.first_row.text.italic is True


def test_table_style_regions_absent_from_the_xml_stay_none():
    from pptx2svg.parse.parts import parse_table_style

    style = parse_table_style(xml(TABLE_STYLE_XML)[0])
    assert style.band1_h is None and style.last_row is None and style.nw_cell is None


def test_deck_table_styles_are_read_from_the_presentation_part(basic_theme):
    package = OpcPackage.open(str(basic_theme))
    styles = read_presentation(package).table_styles
    assert styles is not None
    # This deck is a Google Slides export, which writes a real custom style out.
    style = styles.styles[styles.default_style_id]
    assert style.name == "Table_0"
    assert style.whole_table.border_inside_v.fill.color.hex == "9E9E9E"


def test_table_reads_its_style_id_and_region_flags(authoring):
    package = OpcPackage.open(str(authoring))
    presentation = read_presentation(package)
    tables = [
        shape
        for shape in presentation.slides[0].shapes
        if isinstance(shape, s.SourceTable)
    ]
    assert len(tables) == 1
    assert tables[0].style_id == "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    assert tables[0].first_row and tables[0].band_row
    assert not tables[0].last_row and not tables[0].band_col


# -- Built-in table style catalogue ----------------------------------------------------


def test_builtin_catalogue_expands_into_the_same_shape_the_reader_produces():
    from pptx2svg.parse.table_styles_builtin import builtin_table_style

    style = builtin_table_style("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    assert style.name == "Medium Style 2 - Accent 1"
    # Measured from PowerPoint's own render: the body is a 20% tint of the accent, the
    # banded rows a 40% tint, the header solid accent with white bold text, and every
    # gridline a 1 pt white rule with a 3 pt one under the header.
    assert style.whole_table.fill.color.scheme == "accent1"
    assert style.whole_table.fill.color.transforms[0].value == 20000
    assert style.band1_h.fill.color.transforms[0].value == 40000
    assert style.band2_h is None
    assert style.whole_table.border_inside_h.width == 12700
    assert style.first_row.fill.color.transforms == []
    assert style.first_row.border_bottom.width == 38100
    assert style.first_row.text.bold is True
    assert style.first_row.text.color.scheme == "lt1"


def test_builtin_catalogue_is_keyed_by_upper_case_guid():
    from pptx2svg.parse.table_styles_builtin import BUILTIN_TABLE_STYLES, builtin_table_style

    assert BUILTIN_TABLE_STYLES
    for guid, spec in BUILTIN_TABLE_STYLES.items():
        assert guid == guid.upper(), guid
        assert guid.startswith("{") and guid.endswith("}")
        assert spec["name"]
        # Every entry must expand without raising.
        assert builtin_table_style(guid) is not None
    # Lookup is case-insensitive, since decks do not agree on the case of a GUID.
    lower = "{5c22544a-7ee6-4342-b048-85bdc9fd1c3a}"
    assert builtin_table_style(lower).name == "Medium Style 2 - Accent 1"
    assert builtin_table_style("{not-a-style}") is None


def test_the_builtin_catalogue_is_complete():
    """All 74 of PowerPoint's built-in table styles, measured out of PowerPoint itself.

    The gallery is ten families of six accents, plus eight accent-less variants and the
    two "No Style" entries.  "Dark Style 2" is the exception the count has to allow for:
    PowerPoint pairs its accents as "Accent 1/Accent 2", "Accent 3/Accent 4" and "Accent
    5/Accent 6", so that family has three accent variants rather than six.

    A missing entry is invisible in output -- the table just renders unstyled -- which is
    why it is asserted here rather than left to be noticed.
    """
    from pptx2svg.parse.table_styles_builtin import BUILTIN_TABLE_STYLES

    names = {spec["name"] for spec in BUILTIN_TABLE_STYLES.values()}
    assert len(BUILTIN_TABLE_STYLES) == 74
    assert len(names) == 74, "two entries share a name"

    expected = {"No Style, No Grid", "No Style, Table Grid"}
    for family in ("Themed Style 1", "Themed Style 2", "Light Style 1", "Light Style 2",
                   "Light Style 3", "Medium Style 1", "Medium Style 2", "Medium Style 3",
                   "Medium Style 4", "Dark Style 1"):
        expected |= {f"{family} - Accent {n}" for n in range(1, 7)}
    # The accent-less base variants PowerPoint offers; Themed Style 1 and 2 have none.
    expected |= {"Light Style 1", "Light Style 2", "Light Style 3", "Medium Style 1",
                 "Medium Style 2", "Medium Style 3", "Medium Style 4", "Dark Style 1",
                 "Dark Style 2"}
    expected |= {"Dark Style 2 - Accent 1/Accent 2", "Dark Style 2 - Accent 3/Accent 4",
                 "Dark Style 2 - Accent 5/Accent 6"}
    assert names == expected


# -- Whole packages --------------------------------------------------------------------


def test_every_fixture_parses(pptx_path):
    package = OpcPackage.open(str(pptx_path))
    presentation = read_presentation(package)
    assert presentation.slides, "expected at least one slide"
    assert presentation.slide_width > 0 and presentation.slide_height > 0
    # Every slide should reach a layout, and every layout a master.
    for slide in presentation.slides:
        assert slide.layout_part_path in presentation.layouts

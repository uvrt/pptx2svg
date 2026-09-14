"""SmartArt, rendered from the DrawingML drawing PowerPoint caches beside the data model.

**Provenance warning.** None of the fixtures contain SmartArt, ``python-pptx`` cannot
author it, and no permissively-licensed sample was available on this machine, so the
diagram parts below are hand-written from ECMA-376 Part 1 §21.4 and the
``diagramDrawing`` relationship Microsoft defines for the cached rendering.  They are
modelled closely on what PowerPoint emits -- ``dsp:``-namespaced shapes, ``id="0"`` on
every ``cNvPr``, ``modelId`` GUIDs carrying the real identity, a ``grpSpPr/a:xfrm``
matching the frame -- but they are **not** a genuine PowerPoint file.

What that means in practice: these tests prove the lookup chain, the coordinate mapping
and the fallbacks behave as specified.  They cannot prove the specification was read
correctly.  A real SmartArt deck still needs to be run through
``convert_pptx_to_model`` and compared against PowerPoint's own export.
"""

from __future__ import annotations

import pytest

import pptx2svg.model as m
from deckbuilder import derive_deck
from pptx2svg import ConvertOptions, convert_pptx_to_model, convert_pptx_to_svg

DIAGRAM_DATA_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
)
MS_DRAWING_REL = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
OCLC_DRAWING_REL = "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramDrawing"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

DATA_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
)
DRAWING_CONTENT_TYPE = "application/vnd.ms-office.drawingml.diagramDrawing+xml"

#: The frame on the slide: 2 in x 1 in at (1 in, 1 in).
FRAME_X, FRAME_Y, FRAME_CX, FRAME_CY = 914400, 914400, 1828800, 914400

GRAPHIC_FRAME_XML = """<p:graphicFrame
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <p:nvGraphicFramePr><p:cNvPr id="7001" name="Diagram 1"/>
  <p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
 <p:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
 <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
   r:dm="rIdDm" r:lo="rIdDm" r:qs="rIdDm" r:cs="rIdDm"/>
 </a:graphicData></a:graphic></p:graphicFrame>"""

DATA_MODEL_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <dgm:ptLst>
  <dgm:pt modelId="{11111111-1111-1111-1111-111111111111}" type="doc"><dgm:prSet/></dgm:pt>
 </dgm:ptLst>
 <dgm:cxnLst/><dgm:bg/><dgm:whole/>
</dgm:dataModel>"""


def drawing_xml(*, group_xfrm: str, shapes: str) -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <dsp:spTree>
  <dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr>
  <dsp:grpSpPr>{group_xfrm}</dsp:grpSpPr>
  {shapes}
 </dsp:spTree>
</dsp:drawing>""".encode()


GROUP_XFRM = (
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm>'
)


def node_shape(model_id: str, x: int, cx: int, colour: str, text: str) -> str:
    """One laid-out SmartArt node, exactly as PowerPoint writes it into the cache."""
    return f"""<dsp:sp modelId="{{{model_id}}}">
  <dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>
  <dsp:spPr>
   <a:xfrm><a:off x="{x}" y="0"/><a:ext cx="{cx}" cy="914400"/></a:xfrm>
   <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
   <a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>
  </dsp:spPr>
  <dsp:txBody><a:bodyPr/><a:lstStyle/>
   <a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>{text}</a:t></a:r></a:p>
  </dsp:txBody>
 </dsp:sp>"""


TWO_NODES = node_shape(
    "22222222-2222-2222-2222-222222222222", 0, 838200, "4472C4", "Plan"
) + node_shape("33333333-3333-3333-3333-333333333333", 990600, 838200, "ED7D31", "Ship")


def deck_with_diagram(
    basic_theme,
    *,
    drawing: bytes | None = drawing_xml(group_xfrm=GROUP_XFRM, shapes=TWO_NODES),
    drawing_rel_type: str = MS_DRAWING_REL,
    extra_parts: dict[str, bytes] | None = None,
    drawing_rels: bytes | None = None,
) -> bytes:
    """A deck whose slide 1 gains a SmartArt frame, with the parts it points at.

    ``drawing=None`` models the case PowerPoint produces when the cached rendering has
    been stripped: the data model is still there, the drawing relationship is not.
    """
    parts: dict[str, bytes] = {"ppt/diagrams/data1.xml": DATA_MODEL_XML}
    overrides = {"ppt/diagrams/data1.xml": DATA_CONTENT_TYPE}
    data_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    )
    if drawing is not None:
        parts["ppt/diagrams/drawing1.xml"] = drawing
        overrides["ppt/diagrams/drawing1.xml"] = DRAWING_CONTENT_TYPE
        data_rels += (
            f'<Relationship Id="rIdDrw" Type="{drawing_rel_type}" Target="drawing1.xml"/>'
        )
    data_rels += "</Relationships>"
    parts["ppt/diagrams/_rels/data1.xml.rels"] = data_rels.encode()
    if drawing_rels is not None:
        parts["ppt/diagrams/_rels/drawing1.xml.rels"] = drawing_rels
    parts.update(extra_parts or {})

    return derive_deck(
        basic_theme,
        parts=parts,
        shapes_xml=GRAPHIC_FRAME_XML,
        slide_relationships=[("rIdDm", DIAGRAM_DATA_REL, "../diagrams/data1.xml")],
        overrides=overrides,
    )


def diagram_group(slide) -> m.GroupElement | None:
    for element in slide.elements:
        if isinstance(element, m.GroupElement) and (element.element_id or "").endswith(".7001"):
            return element
    return None


# --------------------------------------------------------------------------------------


def test_a_cached_diagram_drawing_becomes_a_group_of_shapes(basic_theme):
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck_with_diagram(basic_theme), options)

    group = diagram_group(resolved.slides[0])
    assert group is not None, "the graphic frame should resolve to a group, not a placeholder"
    assert len(group.children) == 2
    assert [c.fill.color.hex for c in group.children] == ["#4472c4", "#ed7d31"]
    assert not [w for w in options.warnings if w.code == "unsupported-graphic-frame"]


def test_the_frame_transform_and_the_diagram_coordinate_space_both_survive(basic_theme):
    resolved = convert_pptx_to_model(
        deck_with_diagram(basic_theme), ConvertOptions(slide_numbers=[1])
    )
    group = diagram_group(resolved.slides[0])
    assert group is not None
    # Outer: where the frame sits on the slide.
    assert (group.transform.offset_x, group.transform.offset_y) == (FRAME_X, FRAME_Y)
    assert (group.transform.extent_width, group.transform.extent_height) == (FRAME_CX, FRAME_CY)
    # Inner: the space the cached shapes were laid out in, from grpSpPr/a:xfrm chOff/chExt.
    assert (group.child_transform.offset_x, group.child_transform.offset_y) == (0, 0)
    assert group.child_transform.extent_width == FRAME_CX


def test_diagram_text_is_resolved_like_any_other_drawingml(basic_theme):
    resolved = convert_pptx_to_model(
        deck_with_diagram(basic_theme), ConvertOptions(slide_numbers=[1])
    )
    group = diagram_group(resolved.slides[0])
    assert group is not None
    runs = [
        run.text
        for shape in group.children
        for paragraph in (shape.text_body.paragraphs if shape.text_body else [])
        for run in paragraph.runs
    ]
    assert runs == ["Plan", "Ship"]


def test_the_purl_spelling_of_the_drawing_relationship_is_accepted(basic_theme):
    """ISO/IEC transitional files use purl.oclc.org for the same relationship."""
    deck = deck_with_diagram(basic_theme, drawing_rel_type=OCLC_DRAWING_REL)
    resolved = convert_pptx_to_model(deck, ConvertOptions(slide_numbers=[1]))
    group = diagram_group(resolved.slides[0])
    assert group is not None and len(group.children) == 2


def test_a_missing_group_xfrm_falls_back_to_the_frame_extent(basic_theme):
    """Not `replace(transform)`: diagram coordinates start at the frame's origin, not the slide's."""
    deck = deck_with_diagram(
        basic_theme, drawing=drawing_xml(group_xfrm="", shapes=TWO_NODES)
    )
    resolved = convert_pptx_to_model(deck, ConvertOptions(slide_numbers=[1]))
    group = diagram_group(resolved.slides[0])
    assert group is not None
    assert (group.child_transform.offset_x, group.child_transform.offset_y) == (0, 0)
    assert group.child_transform.extent_width == FRAME_CX
    assert group.child_transform.extent_height == FRAME_CY


def test_a_diagram_scales_when_its_layout_space_differs_from_the_frame(basic_theme):
    """A cache laid out at a different size must be scaled, not clipped."""
    half = (
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/></a:xfrm>'
    )
    deck = deck_with_diagram(basic_theme, drawing=drawing_xml(group_xfrm=half, shapes=TWO_NODES))
    resolved = convert_pptx_to_model(deck, ConvertOptions(slide_numbers=[1]))
    group = diagram_group(resolved.slides[0])
    assert group is not None
    assert group.child_transform.extent_width == 914400
    # The renderer turns extent/chExt into the scale factor; 2x here.
    svg = convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[1]))[0]
    assert "scale(2, 2)" in svg


def test_a_deleted_drawing_part_still_renders_and_still_warns(basic_theme):
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck_with_diagram(basic_theme, drawing=None), options)

    assert diagram_group(resolved.slides[0]) is None
    warnings = [w for w in options.warnings if w.code == "unsupported-graphic-frame"]
    assert warnings and "diagram" in warnings[0].message


def test_an_empty_cached_drawing_warns_rather_than_emitting_an_empty_group(basic_theme):
    deck = deck_with_diagram(basic_theme, drawing=drawing_xml(group_xfrm=GROUP_XFRM, shapes=""))
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    assert diagram_group(resolved.slides[0]) is None
    assert [w for w in options.warnings if w.code == "unsupported-graphic-frame"]


def test_a_corrupt_drawing_part_does_not_abort_the_deck(basic_theme):
    deck = deck_with_diagram(basic_theme, drawing=b"<dsp:drawing not xml at all")
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    assert diagram_group(resolved.slides[0]) is None
    assert [w for w in options.warnings if w.code == "unsupported-graphic-frame"]
    assert resolved.slides[0].elements, "the rest of the slide must still resolve"


def test_pictures_inside_a_diagram_resolve_against_the_drawing_parts_relationships(basic_theme):
    """A diagram's images are related to the *drawing* part; the slide has no such rel."""
    png = bytes.fromhex(
        "89504e470d0a1a0a0000000d49484452000000010000000108020000009077"
        "3d780000000c4944415408d763f8cfc00000030101007a5dc0620000000049454e44ae426082"
    )
    picture = """<dsp:pic xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <dsp:nvPicPr><dsp:cNvPr id="0" name=""/><dsp:cNvPicPr/></dsp:nvPicPr>
 <dsp:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></dsp:blipFill>
 <dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
 <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr></dsp:pic>"""

    deck = deck_with_diagram(
        basic_theme,
        drawing=drawing_xml(group_xfrm=GROUP_XFRM, shapes=picture),
        extra_parts={"ppt/diagrams/media/dgmimage1.png": png},
        drawing_rels=(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rIdImg" Type="{IMAGE_REL}" Target="media/dgmimage1.png"/>'
            "</Relationships>"
        ).encode(),
    )
    resolved = convert_pptx_to_model(deck, ConvertOptions(slide_numbers=[1]))
    group = diagram_group(resolved.slides[0])
    assert group is not None and len(group.children) == 1
    assert isinstance(group.children[0], m.ImageElement)
    assert group.children[0].mime_type == "image/png"


def test_diagram_children_get_unique_addressable_ids(basic_theme):
    """PowerPoint writes id="0" on every cached shape, so ours must be rebuilt."""
    resolved = convert_pptx_to_model(
        deck_with_diagram(basic_theme), ConvertOptions(slide_numbers=[1])
    )
    group = diagram_group(resolved.slides[0])
    assert group is not None
    ids = [child.element_id for child in group.children]
    assert len(set(ids)) == len(ids)
    assert all(identifier.startswith("256.7001/") for identifier in ids)
    assert [child.element_path for child in group.children] == ["2-0", "2-1"]


def test_the_rendered_svg_carries_the_diagram_and_its_identity(basic_theme):
    svg = convert_pptx_to_svg(deck_with_diagram(basic_theme), ConvertOptions(slide_numbers=[1]))[0]
    assert 'data-pptx-id="256.7001"' in svg
    assert "Plan" in svg and "Ship" in svg
    assert "#4472c4" in svg


def test_diagram_output_is_reproducible(basic_theme):
    deck = deck_with_diagram(basic_theme)
    first = convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[1]))[0]
    second = convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[1]))[0]
    assert first == second


@pytest.mark.skip(
    reason="activates when a genuine PowerPoint-authored SmartArt fixture is added to "
    "tests/fixtures/; everything above is hand-written from the spec"
)
def test_a_real_powerpoint_smartart_deck_renders_its_nodes():  # pragma: no cover
    from pathlib import Path

    deck = Path(__file__).parent / "fixtures" / "real-smartart.pptx"
    options = ConvertOptions()
    resolved = convert_pptx_to_model(deck, options)
    groups = [
        element
        for slide in resolved.slides
        for element in slide.elements
        if isinstance(element, m.GroupElement) and element.children
    ]
    assert groups, "the diagram frame should resolve to a populated group"
    assert not [w for w in options.warnings if w.code == "unsupported-graphic-frame"]

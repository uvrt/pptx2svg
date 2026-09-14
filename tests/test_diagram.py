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


def test_a_deleted_drawing_part_says_the_cache_is_missing(basic_theme):
    """Twenty of the forty-six real decks measured are in exactly this state -- Office
    2007 did not always write the cache.  "no cached drawing" is a different fact from
    "failed to load", and someone looking at a blank rectangle needs to know which."""
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck_with_diagram(basic_theme, drawing=None), options)

    assert diagram_group(resolved.slides[0]) is None
    warnings = [w for w in options.warnings if w.code == "diagram-no-cached-drawing"]
    assert warnings, [str(w) for w in options.warnings]
    assert "not implemented" in warnings[0].message
    # The vaguer warning must not be piled on top of the specific one.
    assert not [w for w in options.warnings if w.code == "unsupported-graphic-frame"]


def test_an_empty_cached_drawing_warns_rather_than_emitting_an_empty_group(basic_theme):
    """A complete dsp:spTree holding nvGrpSpPr and grpSpPr and no shapes.  PowerPoint
    really writes this -- thirteen of the forty-six real decks measured are like it --
    and it needs the same explanation as having no drawing at all."""
    deck = deck_with_diagram(basic_theme, drawing=drawing_xml(group_xfrm=GROUP_XFRM, shapes=""))
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    assert diagram_group(resolved.slides[0]) is None
    warnings = [w for w in options.warnings if w.code == "diagram-no-cached-drawing"]
    assert warnings and "empty" in warnings[0].message


def test_a_corrupt_drawing_part_does_not_abort_the_deck(basic_theme):
    deck = deck_with_diagram(basic_theme, drawing=b"<dsp:drawing not xml at all")
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    assert diagram_group(resolved.slides[0]) is None
    warnings = [w for w in options.warnings if w.code == "diagram-unreadable"]
    assert warnings and "well-formed" in warnings[0].message
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


# --------------------------------------------------------------------------------------
# The layout PowerPoint actually writes
#
# The decks above use the data part's own relationships, which is what the schema reads
# like and what an independent producer might write.  Real PowerPoint does something
# else, and the fixtures could not have shown it: `dgm:relIds` names the data model,
# layout, quick style and colours but *not* the drawing, because the cached drawing was
# added after `relIds` was specified.  The drawing's relationship id lives in an
# extension inside the data part and resolves against the *slide's* relationships.
#
# Measured across the 46 SmartArt decks in LibreOffice's test corpus: 27 carry a cached
# drawing, and all 27 key it this way.  Exactly one has a
# `ppt/diagrams/_rels/data1.xml.rels` at all, and not for the drawing.
# --------------------------------------------------------------------------------------

DATA_MODEL_EXT_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram"


def data_model_with_ext(relationship_id: str | None) -> bytes:
    """A data-model part carrying ``dsp:dataModelExt@relId``, as PowerPoint writes it."""
    extension = ""
    if relationship_id is not None:
        extension = (
            "<dgm:extLst>"
            f'<a:ext uri="{DATA_MODEL_EXT_NS}">'
            f'<dsp:dataModelExt xmlns:dsp="{DATA_MODEL_EXT_NS}" '
            f'relId="{relationship_id}" minVer="http://schemas.openxmlformats.org/'
            'drawingml/2006/diagram"/>'
            "</a:ext></dgm:extLst>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        "<dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/>"
        f"{extension}</dgm:dataModel>"
    ).encode()


def frame_xml(shape_id: int, data_rel_id: str, x: int) -> str:
    return f"""<p:graphicFrame
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="Diagram {shape_id}"/>
  <p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
 <p:xfrm><a:off x="{x}" y="914400"/><a:ext cx="1828800" cy="914400"/></p:xfrm>
 <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
   r:dm="{data_rel_id}" r:lo="{data_rel_id}" r:qs="{data_rel_id}" r:cs="{data_rel_id}"/>
 </a:graphicData></a:graphic></p:graphicFrame>"""


def powerpoint_layout_deck(basic_theme, frames, *, extra_slide_rels=(), decoys=0):
    """Build a deck the way PowerPoint lays SmartArt out.

    ``frames`` is a list of ``(shape_id, label, ext_rel_id)``: ``ext_rel_id`` is what the
    data part's ``dataModelExt`` names, or ``None`` to omit the extension entirely.  No
    ``ppt/diagrams/_rels/*`` parts are written at all, because real decks do not have
    them -- so a lookup that goes through the data part's own relationships finds
    nothing, which is the bug this guards.
    """
    parts: dict[str, bytes] = {}
    overrides: dict[str, str] = {}
    slide_rels = list(extra_slide_rels)
    shapes = ""

    # `decoys` adds cached drawings that no frame points at.  Their only job is to make
    # the count of drawing relationships greater than one, which disables the
    # lone-relationship fallback -- so a test using them can only pass through the
    # `dataModelExt` id.
    for decoy in range(decoys):
        path = f"ppt/diagrams/drawingDecoy{decoy}.xml"
        parts[path] = drawing_xml(
            group_xfrm=GROUP_XFRM,
            shapes=node_shape(
                "9999999-9999-9999-9999-99999999999" + str(decoy),
                0,
                838200,
                "FF0000",
                f"DECOY{decoy}",
            ),
        )
        overrides[path] = DRAWING_CONTENT_TYPE
        slide_rels.append(
            (f"rIdDecoy{decoy}", MS_DRAWING_REL, f"../diagrams/drawingDecoy{decoy}.xml")
        )

    for index, (shape_id, label, ext_rel_id) in enumerate(frames, start=1):
        data_path = f"ppt/diagrams/data{index}.xml"
        drawing_path = f"ppt/diagrams/drawing{index}.xml"
        data_rel, drawing_rel = f"rIdDm{index}", f"rIdDrw{index}"

        parts[data_path] = data_model_with_ext(ext_rel_id)
        overrides[data_path] = DATA_CONTENT_TYPE
        parts[drawing_path] = drawing_xml(
            group_xfrm=GROUP_XFRM,
            shapes=node_shape(
                f"{index}1111111-1111-1111-1111-111111111111", 0, 838200, "4472C4", label
            ),
        )
        overrides[drawing_path] = DRAWING_CONTENT_TYPE

        slide_rels.append((data_rel, DIAGRAM_DATA_REL, f"../diagrams/data{index}.xml"))
        slide_rels.append(
            (drawing_rel, MS_DRAWING_REL, f"../diagrams/drawing{index}.xml")
        )
        shapes += frame_xml(shape_id, data_rel, 914400 + (index - 1) * 2000000)

    return derive_deck(
        basic_theme,
        parts=parts,
        shapes_xml=shapes,
        slide_relationships=slide_rels,
        overrides=overrides,
    )


def group_for(slide, shape_id: int) -> m.GroupElement | None:
    for element in slide.elements:
        if isinstance(element, m.GroupElement) and (element.element_id or "").endswith(
            f".{shape_id}"
        ):
            return element
    return None


def labels_of(group) -> list[str]:
    return [
        run.text
        for shape in group.children
        for paragraph in (shape.text_body.paragraphs if shape.text_body else [])
        for run in paragraph.runs
    ]


def test_the_drawing_is_found_through_dataModelExt_against_the_slides_rels(basic_theme):
    """The whole of the real-world case: no diagram _rels part exists, and the id that
    names the drawing is written in the data part but resolved on the slide.

    A decoy drawing is added so that the slide carries two drawing relationships.  That
    disables the lone-relationship fallback, leaving the `dataModelExt` id as the only
    route -- without it this test would pass on the fallback alone and prove nothing.
    """
    deck = powerpoint_layout_deck(basic_theme, [(8001, "Plan", "rIdDrw1")], decoys=1)
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    group = group_for(resolved.slides[0], 8001)
    assert group is not None, [str(w) for w in options.warnings]
    assert labels_of(group) == ["Plan"], "picked up the decoy instead of its own drawing"
    assert not options.warnings


def test_two_frames_on_one_slide_each_get_their_own_drawing(basic_theme):
    """`relIds` cannot disambiguate -- it does not name the drawing at all -- so the
    `dataModelExt` id is the only thing keeping two diagrams on one slide apart."""
    deck = powerpoint_layout_deck(
        basic_theme, [(8001, "First", "rIdDrw1"), (8002, "Second", "rIdDrw2")]
    )
    resolved = convert_pptx_to_model(deck, ConvertOptions(slide_numbers=[1]))

    first = group_for(resolved.slides[0], 8001)
    second = group_for(resolved.slides[0], 8002)
    assert first is not None and second is not None
    assert labels_of(first) == ["First"]
    assert labels_of(second) == ["Second"]


def test_the_frames_are_not_crossed_when_the_ids_are_swapped(basic_theme):
    """A lookup that ignored the id and simply took the first drawing it found would
    pass the test above and fail this one."""
    deck = powerpoint_layout_deck(
        basic_theme, [(8001, "First", "rIdDrw2"), (8002, "Second", "rIdDrw1")]
    )
    resolved = convert_pptx_to_model(deck, ConvertOptions(slide_numbers=[1]))

    assert labels_of(group_for(resolved.slides[0], 8001)) == ["Second"]
    assert labels_of(group_for(resolved.slides[0], 8002)) == ["First"]


def test_a_lone_drawing_relationship_is_used_when_there_is_no_extension(basic_theme):
    """Last-resort fallback: one diagram, one drawing relationship on the slide, no
    extension to key it with.  Unambiguous, so it is used."""
    deck = powerpoint_layout_deck(basic_theme, [(8001, "Plan", None)])
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    group = group_for(resolved.slides[0], 8001)
    assert group is not None and labels_of(group) == ["Plan"]


def test_two_candidate_drawings_without_an_extension_are_refused(basic_theme):
    """With two drawings on the slide and nothing to key them by, guessing would put the
    wrong diagram in the wrong frame half the time.  Refusing says so instead."""
    deck = powerpoint_layout_deck(
        basic_theme, [(8001, "First", None), (8002, "Second", None)]
    )
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    assert group_for(resolved.slides[0], 8001) is None
    assert group_for(resolved.slides[0], 8002) is None
    assert len([w for w in options.warnings if w.code == "diagram-no-cached-drawing"]) == 2


def test_an_extension_pointing_at_a_missing_part_falls_back_rather_than_failing(basic_theme):
    """A dangling relId must not shadow the fallback -- one of the 46 real decks has a
    diagramDrawing relationship whose target is not in the package."""
    deck = powerpoint_layout_deck(basic_theme, [(8001, "Plan", "rIdNoSuchThing")])
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck, options)

    group = group_for(resolved.slides[0], 8001)
    assert group is not None and labels_of(group) == ["Plan"]


# --------------------------------------------------------------------------------------
# dsp:txXfrm -- the label's own box
# --------------------------------------------------------------------------------------

TX_SHAPE = """<dsp:sp xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" modelId="{{A}}">
  <dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>
  <dsp:spPr>
   <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>
   <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
  </dsp:spPr>
  <dsp:txBody><a:bodyPr/><a:lstStyle/>
   <a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>Ring</a:t></a:r></a:p>
  </dsp:txBody>
  {txXfrm}
 </dsp:sp>"""

#: 1/4 inch right and 1/2 inch down from the shape's own origin, and wide enough that
#: the label still fits on one line -- the offset is what this fixture is about.
TX_XFRM = (
    '<dsp:txXfrm xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<a:off x="228600" y="457200"/><a:ext cx="1828800" cy="228600"/></dsp:txXfrm>'
)


def deck_with_tx_shape(basic_theme, tx_xfrm: str) -> bytes:
    return deck_with_diagram(
        basic_theme,
        drawing=drawing_xml(
            group_xfrm=GROUP_XFRM, shapes=TX_SHAPE.format(txXfrm=tx_xfrm)
        ),
    )


def test_a_shape_without_txXfrm_puts_its_text_in_itself(basic_theme):
    """The control: every shape outside SmartArt has no txXfrm and must be untouched."""
    svg = convert_pptx_to_svg(
        deck_with_tx_shape(basic_theme, ""), ConvertOptions(slide_numbers=[1])
    )[0]
    assert "Ring" in svg
    assert 'transform="translate(24, 48)"' not in svg


def test_txXfrm_moves_the_label_to_its_own_box(basic_theme):
    """SmartArt places a label in the part of the shape it is meant to annotate -- the
    sliver a Venn ring does not share, the space beside a cycle arrow.  Without this the
    label lands at the shape's origin, which for a circle is the corner of its bounding
    box."""
    svg = convert_pptx_to_svg(
        deck_with_tx_shape(basic_theme, TX_XFRM), ConvertOptions(slide_numbers=[1])
    )[0]
    assert "Ring" in svg
    # 228600 EMU = 24 px, 457200 EMU = 48 px, relative to the shape's own origin.
    assert 'transform="translate(24, 48)"' in svg


def test_the_label_is_wrapped_to_its_own_width_not_the_shapes(basic_theme):
    """The box supplies the extent as well as the offset, so a narrow label box wraps
    text the shape's own width would have fitted on one line."""
    narrow = (
        '<dsp:txXfrm xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:off x="0" y="0"/><a:ext cx="228600" cy="914400"/></dsp:txXfrm>'
    )
    wide = convert_pptx_to_svg(
        deck_with_tx_shape(basic_theme, ""), ConvertOptions(slide_numbers=[1])
    )[0]
    thin = convert_pptx_to_svg(
        deck_with_tx_shape(basic_theme, narrow), ConvertOptions(slide_numbers=[1])
    )[0]
    assert wide != thin


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

"""Reader for ``p:spTree`` -- the shape tree of a slide, layout, master or group.

The tree mixes six child element types (``p:sp``, ``p:pic``, ``p:cxnSp``, ``p:grpSp``,
``p:graphicFrame``, ``mc:AlternateContent``) and their order is the z-order, so children
are walked in document order and each is classified by local name.

A ``p:graphicFrame`` wraps whatever a slide embeds through the DrawingML graphic
extension point: tables are read fully; charts, SmartArt, OLE objects and media are
recorded as :class:`SourceUnsupported` with their transform, so the resolver can still
draw a positioned fallback rather than silently dropping the frame.
"""

from __future__ import annotations

from xml.etree.ElementTree import Element

from ..xmlutil import (
    attr,
    child,
    child_text,
    children,
    int_attr,
    is_true,
    local_name,
    ns_attr,
    num_attr,
)
from .drawing import (
    parse_blip_effects,
    parse_effect_list,
    parse_fill,
    parse_geometry,
    parse_group_transforms,
    parse_image_fill_tile,
    parse_line,
    parse_outline,
    parse_relative_rect,
    parse_shape_style,
    parse_transform,
)
from .source import (
    SourceConnector,
    SourceGroup,
    SourceImage,
    SourcePlaceholder,
    SourceShape,
    SourceShapeNode,
    SourceTable,
    SourceTableCell,
    SourceTableRow,
    SourceUnsupported,
)
from .text import parse_text_body

#: ``a:graphicData@uri`` values that tell us what a graphic frame actually holds.
GRAPHIC_DATA_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"
GRAPHIC_DATA_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
GRAPHIC_DATA_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
GRAPHIC_DATA_OLE = "http://schemas.openxmlformats.org/presentationml/2006/ole"


def parse_shape_tree(sp_tree: Element | None) -> list[SourceShapeNode]:
    """Parse the children of ``p:spTree`` (or ``p:grpSp``) in z-order."""
    if sp_tree is None:
        return []

    nodes: list[SourceShapeNode] = []
    for node in sp_tree:
        parsed = parse_shape_node(node)
        if parsed is not None:
            nodes.append(parsed)
    return nodes


def parse_shape_node(node: Element) -> SourceShapeNode | None:
    name = local_name(node.tag)
    if name == "sp":
        return parse_shape(node)
    if name == "pic":
        return parse_picture(node)
    if name == "cxnSp":
        return parse_connector(node)
    if name == "grpSp":
        return parse_group(node)
    if name == "graphicFrame":
        return parse_graphic_frame(node)
    if name == "AlternateContent":
        return parse_alternate_content(node)
    # nvGrpSpPr / grpSpPr on a spTree, extLst, and anything else is not a drawable child.
    return None


def parse_alternate_content(node: Element) -> SourceShapeNode | None:
    """``mc:AlternateContent`` -- prefer the ``mc:Fallback`` branch, which is plain DrawingML."""
    for branch_name in ("Fallback", "Choice"):
        branch = child(node, branch_name)
        if branch is None:
            continue
        for candidate in branch:
            parsed = parse_shape_node(candidate)
            if parsed is not None:
                return parsed
    return None


# --------------------------------------------------------------------------------------
# Shapes and connectors
# --------------------------------------------------------------------------------------


def parse_shape(sp: Element) -> SourceShape:
    nv_sp_pr = child(sp, "nvSpPr")
    c_nv_pr = child(nv_sp_pr, "cNvPr")
    sp_pr = child(sp, "spPr")

    return SourceShape(
        name=attr(c_nv_pr, "name"),
        shape_id=attr(c_nv_pr, "id"),
        alt_text=_alt_text(c_nv_pr),
        placeholder=parse_placeholder(nv_sp_pr),
        transform=parse_transform(sp_pr),
        geometry=parse_geometry(sp_pr),
        fill=parse_fill(sp_pr),
        outline=parse_outline(sp_pr),
        effects=parse_effect_list(child(sp_pr, "effectLst")),
        style=parse_shape_style(child(sp, "style")),
        text_body=parse_text_body(child(sp, "txBody")),
        hyperlink_rel_id=_hyperlink_rel_id(c_nv_pr),
        hidden=_hidden(c_nv_pr),
    )


def parse_connector(cxn_sp: Element) -> SourceConnector:
    nv_cxn_sp_pr = child(cxn_sp, "nvCxnSpPr")
    c_nv_pr = child(nv_cxn_sp_pr, "cNvPr")
    sp_pr = child(cxn_sp, "spPr")

    return SourceConnector(
        name=attr(c_nv_pr, "name"),
        shape_id=attr(c_nv_pr, "id"),
        alt_text=_alt_text(c_nv_pr),
        transform=parse_transform(sp_pr),
        geometry=parse_geometry(sp_pr),
        outline=parse_outline(sp_pr),
        effects=parse_effect_list(child(sp_pr, "effectLst")),
        style=parse_shape_style(child(cxn_sp, "style")),
        hidden=_hidden(c_nv_pr),
    )


def parse_picture(pic: Element) -> SourceImage:
    nv_pic_pr = child(pic, "nvPicPr")
    c_nv_pr = child(nv_pic_pr, "cNvPr")
    sp_pr = child(pic, "spPr")
    blip_fill = child(pic, "blipFill")
    blip = child(blip_fill, "blip")

    return SourceImage(
        blip_relationship_id=ns_attr(blip, "embed"),
        name=attr(c_nv_pr, "name"),
        shape_id=attr(c_nv_pr, "id"),
        alt_text=_alt_text(c_nv_pr),
        placeholder=parse_placeholder(nv_pic_pr),
        transform=parse_transform(sp_pr),
        geometry=parse_geometry(sp_pr),
        outline=parse_outline(sp_pr),
        effects=parse_effect_list(child(sp_pr, "effectLst")),
        blip_effects=parse_blip_effects(blip),
        src_rect=parse_relative_rect(child(blip_fill, "srcRect")),
        stretch=parse_relative_rect(child(child(blip_fill, "stretch"), "fillRect")),
        tile=parse_image_fill_tile(child(blip_fill, "tile")),
        hyperlink_rel_id=_hyperlink_rel_id(c_nv_pr),
        hidden=_hidden(c_nv_pr),
    )


def parse_group(grp_sp: Element) -> SourceGroup:
    nv_grp_sp_pr = child(grp_sp, "nvGrpSpPr")
    c_nv_pr = child(nv_grp_sp_pr, "cNvPr")
    grp_sp_pr = child(grp_sp, "grpSpPr")
    transform, child_transform = parse_group_transforms(grp_sp_pr)

    return SourceGroup(
        name=attr(c_nv_pr, "name"),
        shape_id=attr(c_nv_pr, "id"),
        alt_text=_alt_text(c_nv_pr),
        transform=transform,
        child_transform=child_transform,
        fill=parse_fill(grp_sp_pr),
        effects=parse_effect_list(child(grp_sp_pr, "effectLst")),
        children=parse_shape_tree(grp_sp),
        hidden=_hidden(c_nv_pr),
    )


def parse_placeholder(nv_pr_parent: Element | None) -> SourcePlaceholder | None:
    ph = child(child(nv_pr_parent, "nvPr"), "ph")
    if ph is None:
        return None
    return SourcePlaceholder(type=attr(ph, "type"), idx=int_attr(ph, "idx"))


def _hidden(c_nv_pr: Element | None) -> bool:
    """``p:cNvPr@hidden`` -- PowerPoint's "hide" in the selection pane."""
    return is_true(attr(c_nv_pr, "hidden"))


def _alt_text(c_nv_pr: Element | None) -> str | None:
    if c_nv_pr is None:
        return None
    return attr(c_nv_pr, "descr") or None


def _hyperlink_rel_id(c_nv_pr: Element | None) -> str | None:
    return ns_attr(child(c_nv_pr, "hlinkClick"), "id")


# --------------------------------------------------------------------------------------
# Graphic frames
# --------------------------------------------------------------------------------------


def parse_graphic_frame(frame: Element) -> SourceShapeNode | None:
    nv_pr = child(frame, "nvGraphicFramePr")
    c_nv_pr = child(nv_pr, "cNvPr")
    transform = parse_transform(frame)  # graphicFrame uses p:xfrm, a direct child
    graphic_data = child(child(frame, "graphic"), "graphicData")
    uri = attr(graphic_data, "uri") or ""

    table = child(graphic_data, "tbl")
    if table is not None:
        return parse_table(
            table,
            name=attr(c_nv_pr, "name"),
            shape_id=attr(c_nv_pr, "id"),
            alt_text=_alt_text(c_nv_pr),
            transform=transform,
            hidden=_hidden(c_nv_pr),
        )

    if uri == GRAPHIC_DATA_CHART:
        what = "chart"
        fallback_rel_id = ns_attr(child(graphic_data, "chart"), "id")
    elif uri == GRAPHIC_DATA_DIAGRAM:
        what = "diagram"
        fallback_rel_id = _diagram_drawing_rel_id(graphic_data)
    elif uri == GRAPHIC_DATA_OLE:
        what = "ole"
        fallback_rel_id = _ole_preview_rel_id(graphic_data)
    else:
        what = _local_graphic_kind(graphic_data) or "graphicFrame"
        fallback_rel_id = None

    return SourceUnsupported(
        what=what,
        name=attr(c_nv_pr, "name"),
        alt_text=_alt_text(c_nv_pr),
        transform=transform,
        fallback_rel_id=fallback_rel_id,
        hidden=_hidden(c_nv_pr),
    )


def _local_graphic_kind(graphic_data: Element | None) -> str | None:
    if graphic_data is None:
        return None
    first = next(iter(graphic_data), None)
    return local_name(first.tag) if first is not None else None


def _diagram_drawing_rel_id(graphic_data: Element | None) -> str | None:
    """SmartArt ships a pre-rendered DrawingML fallback under ``dsp:dataModelExt``."""
    rel_ids = child(graphic_data, "relIds")
    if rel_ids is None:
        return None
    # `r:dm`/`r:lo`/`r:qs`/`r:cs` point at the data model parts; the drawing part is
    # reached from the data model part's own relationships, handled by the resolver.
    return ns_attr(rel_ids, "dm")


def _ole_preview_rel_id(graphic_data: Element | None) -> str | None:
    ole = child(graphic_data, "oleObj")
    if ole is None:
        return None
    pic = child(ole, "pic")
    blip = child(child(pic, "blipFill"), "blip")
    return ns_attr(blip, "embed")


# --------------------------------------------------------------------------------------
# Tables
# --------------------------------------------------------------------------------------


def parse_table(
    tbl: Element,
    *,
    name: str | None,
    shape_id: str | None,
    alt_text: str | None,
    transform,
    hidden: bool = False,
) -> SourceTable:
    tbl_pr = child(tbl, "tblPr")
    columns = [num_attr(col, "w") or 0 for col in children(child(tbl, "tblGrid"), "gridCol")]
    rows = [parse_table_row(tr) for tr in children(tbl, "tr")]

    return SourceTable(
        name=name,
        shape_id=shape_id,
        alt_text=alt_text,
        transform=transform,
        columns=columns,
        rows=rows,
        first_row=is_true(attr(tbl_pr, "firstRow")),
        last_row=is_true(attr(tbl_pr, "lastRow")),
        first_col=is_true(attr(tbl_pr, "firstCol")),
        last_col=is_true(attr(tbl_pr, "lastCol")),
        band_row=is_true(attr(tbl_pr, "bandRow")),
        band_col=is_true(attr(tbl_pr, "bandCol")),
        style_id=(child_text(tbl_pr, "tableStyleId") or "").strip() or None,
        hidden=hidden,
    )


def parse_table_row(tr: Element) -> SourceTableRow:
    return SourceTableRow(
        height=num_attr(tr, "h") or 0,
        cells=[parse_table_cell(tc) for tc in children(tr, "tc")],
    )


def parse_table_cell(tc: Element) -> SourceTableCell:
    tc_pr = child(tc, "tcPr")
    anchor = attr(tc_pr, "anchor")
    return SourceTableCell(
        text_body=parse_text_body(child(tc, "txBody")),
        fill=parse_fill(tc_pr),
        border_top=parse_line(child(tc_pr, "lnT")),
        border_bottom=parse_line(child(tc_pr, "lnB")),
        border_left=parse_line(child(tc_pr, "lnL")),
        border_right=parse_line(child(tc_pr, "lnR")),
        grid_span=int_attr(tc, "gridSpan") or 1,
        row_span=int_attr(tc, "rowSpan") or 1,
        h_merge=is_true(attr(tc, "hMerge")),
        v_merge=is_true(attr(tc, "vMerge")),
        margin_left=num_attr(tc_pr, "marL"),
        margin_right=num_attr(tc_pr, "marR"),
        margin_top=num_attr(tc_pr, "marT"),
        margin_bottom=num_attr(tc_pr, "marB"),
        anchor={"t": "t", "ctr": "ctr", "b": "b"}.get(anchor or "t"),  # type: ignore[arg-type]
    )

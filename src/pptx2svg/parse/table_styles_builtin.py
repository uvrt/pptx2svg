"""PowerPoint's built-in table styles as this renderer's model, keyed by GUID.

PowerPoint does not write a built-in style's definition into ``ppt/tableStyles.xml``,
not even for the style a table in the deck uses, so a table naming one would draw as a
bare grid without a catalogue.  The catalogue -- all 74 styles, measured from PowerPoint
16 by ``tools/derive_table_styles.py`` and ``tools/derive_themed_table_styles.py``, and
where each number came from -- is :mod:`ooxml_common.drawingml.table_styles`, which
pptx-agent reads too (review, S3); this module expands an entry into the same
:class:`~pptx2svg.parse.source.SourceTableStyle` the XML reader produces, so the resolver
cannot tell the two apart.  "Themed Style 1" and "Themed Style 2", which draw on the
theme's format scheme, go through that reader itself.

A table naming a GUID that is in neither the deck nor the catalogue still falls back to
no style, which is what PowerPoint does too -- but the resolver says so with a
``table-style-unknown`` warning, because the render being defensible does not make the
silence acceptable.
"""

from __future__ import annotations

from ooxml_common.drawingml.table_styles import (
    BUILTIN_TABLE_STYLES,
    builtin_regions,
    color_parts,
    themed_style_xml,
)

from ..xmlutil import parse_xml
from .source import (
    ColorTransform,
    SchemeColor,
    SourceOutline,
    SourceRunProperties,
    SourceSolidFill,
    SourceTableCellStyle,
    SourceTableStyle,
    TABLE_STYLE_REGIONS,
)

EMU_PER_POINT = 12700

_BORDER_SIDES = {
    "left": "border_left",
    "right": "border_right",
    "top": "border_top",
    "bottom": "border_bottom",
    "insideH": "border_inside_h",
    "insideV": "border_inside_v",
}


def _color(spec: str) -> SchemeColor:
    """``"accent1/tint40"`` -> the scheme colour with that transform applied."""
    slot, transforms = color_parts(spec)
    return SchemeColor(scheme=slot, transforms=[ColorTransform(kind=kind, value=value)
                                                for kind, value in transforms])


def _region(spec: dict) -> SourceTableCellStyle:
    style = SourceTableCellStyle()
    if "fill" in spec:
        style.fill = SourceSolidFill(color=_color(spec["fill"]))
    if "text" in spec or spec.get("bold"):
        style.text = SourceRunProperties(
            bold=True if spec.get("bold") else None,
            color=_color(spec["text"]) if "text" in spec else None,
        )
    for side, (width_pt, color) in spec.get("borders", {}).items():
        setattr(
            style,
            _BORDER_SIDES[side],
            SourceOutline(
                width=width_pt * EMU_PER_POINT,
                fill=SourceSolidFill(color=_color(color)),
            ),
        )
    return style


_CACHE: dict[str, SourceTableStyle] = {}


def builtin_table_style(style_id: str) -> SourceTableStyle | None:
    """The built-in style with this GUID, or ``None`` if it is not one we know."""
    if not style_id:
        return None
    key = style_id.upper()
    if key in _CACHE:
        return _CACHE[key]
    xml = themed_style_xml(style_id)
    if xml is not None:
        _CACHE[key] = _themed_style(xml)
        return _CACHE[key]
    regions = builtin_regions(key)
    if regions is None:
        return None

    style = SourceTableStyle(style_id=style_id, name=BUILTIN_TABLE_STYLES[key]["name"])
    for element_name, attribute in TABLE_STYLE_REGIONS:
        region = regions.get(element_name)
        if region is not None:
            setattr(style, attribute, _region(region))
    _CACHE[key] = style
    return style


def _themed_style(xml: str) -> SourceTableStyle:
    """A themed style's ``a:tblStyle``, read by the reader a deck's ``tableStyles.xml``
    goes through."""
    from .parts import parse_table_style  # parts reads tableStyles.xml; imported late

    style = parse_table_style(parse_xml(xml))
    assert style is not None
    return style


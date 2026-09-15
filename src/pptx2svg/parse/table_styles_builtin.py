"""PowerPoint's built-in table styles, keyed by the GUID a deck refers to them with.

**Why this file has to exist.**  A table's ``a:tableStyleId`` is a GUID, and the obvious
place to look it up is ``ppt/tableStyles.xml``.  For a built-in style that lookup always
fails: PowerPoint does not write built-in definitions into the file, not even for the
style a table in the deck is actually using.  A PowerPoint-authored deck typically ships
a ``tableStyles.xml`` containing nothing but ``def="{5C22544A-...}"`` -- the id of
"Medium Style 2 - Accent 1", the default applied to any table whose ``a:tblPr`` names no
style at all -- and the definition lives inside the application.  Without a catalogue
here, every table in such a deck renders as a bare grid.

**Where the numbers come from.**  They were measured from PowerPoint 16.x itself, not
transcribed from a specification.  ``tools/derive_table_styles.py`` builds a deck with
one slide per candidate GUID, each slide carrying three tables that between them expose
every conditional region, exports it through PowerPoint twice -- once over a white slide
background and once over black, so a colour's alpha can be solved for rather than
guessed -- and reads the fills, stroke colours, stroke widths, text colours and bold
back out of the render.  Colours are named by matching them against a sheet of 546
swatches whose OOXML expressions are known, which is what turns a sampled ``#cfd5ea``
into ``accent1 tint 40%``.

A GUID PowerPoint does not recognise renders exactly like "No Style, Table Grid", which
is how three candidates that turned out to be wrong were detected and dropped rather
than shipped as plausible-looking guesses.

**All 74 built-in styles are now here.**  The last two to arrive were "Light Style 1 -
Accent 4" (``{D27102A9-...}``) and "Medium Style 1" (``{793D81CF-...}``), which were
missing only because their GUIDs were unknown, not because measuring them failed.
``{793D81CF-...}`` was read out of PowerPoint's own executable: the binary carries
fifteen brace-delimited GUIDs, fourteen of which are table styles already catalogued
here, and that was the fifteenth.  ``{D27102A9-...}`` came from a published list
(``aiden0z/pptx-renderer``) and so was only ever a candidate.  Both were then confirmed
the same way any other entry is -- by measurement.  Neither rendered like the
unrecognised-GUID fallback, and each came out with its family's exact shape carrying the
colour a reader would predict but the tool was never told to expect: accent 4 in the
Light Style 1 shape, and ``dk1`` where the accents sit in the Medium Style 1 shape,
which is what the accent-less base of a family looks like elsewhere (compare "Dark Style
2" against "Dark Style 2 - Accent 1/Accent 2").  Re-deriving the whole catalogue in the
same run reproduced the other 72 entries byte for byte.

*Refuted:* the fourth accent of "Dark Style 2" was once thought missing too.  It is not
-- PowerPoint genuinely pairs that family's accents as "Accent 1/Accent 2", "Accent
3/Accent 4" and "Accent 5/Accent 6", and all three pairs are present.

A table naming a GUID that is in neither the deck nor this table still falls back to no
style, which is what PowerPoint does too -- but the resolver now says so with a
``table-style-unknown`` warning, because the render being defensible does not make the
silence acceptable.

Colour expressions are written compactly as ``slot`` or ``slot/transform`` --
``accent1``, ``accent1/tint40``, ``dk1/alpha20`` -- and borders as
``(width_in_points, colour)``.  :func:`builtin_table_style` expands them into the same
:class:`~pptx2svg.parse.source.SourceTableStyle` the XML reader produces, so the
resolver cannot tell the two apart.
"""

from __future__ import annotations

import re

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

_TRANSFORM = re.compile(r"^([a-zA-Z]+)(\d+)$")

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
    slot, *rest = spec.split("/")
    transforms = []
    for part in rest:
        match = _TRANSFORM.match(part)
        if match is None:
            continue
        kind, value = match.group(1), int(match.group(2))
        transforms.append(ColorTransform(kind=kind, value=value * 1000))
    return SchemeColor(scheme=slot, transforms=transforms)


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
    spec = BUILTIN_TABLE_STYLES.get(key)
    if spec is None:
        return None

    style = SourceTableStyle(style_id=style_id, name=spec["name"])
    for element_name, attribute in TABLE_STYLE_REGIONS:
        if element_name in spec:
            setattr(style, attribute, _region(spec[element_name]))
    _CACHE[key] = style
    return style


#: GUID -> style definition.  Generated by ``tools/derive_table_styles.py``; edit that
#: and re-run it rather than editing the table by hand.
BUILTIN_TABLE_STYLES: dict[str, dict] = {
    "{2D5ABB26-0587-4C30-8999-92F81FD0307C}": {
        "name": "No Style, No Grid",
        "wholeTbl": {"text": "dk1"},
    },
    "{5940675A-B579-460E-94D1-54222C63F5DA}": {
        "name": "No Style, Table Grid",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "dk1"), "right": (1.0, "dk1"), "top": (1.0, "dk1"), "bottom": (1.0, "dk1"), "insideH": (1.0, "dk1"), "insideV": (1.0, "dk1")}},
    },
    "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}": {
        "name": "Themed Style 1 - Accent 1",
        "wholeTbl": {"fill": "accent1", "text": "dk1"},
        "firstRow": {"text": "lt1", "bold": True, "borders": {"bottom": (1.0, "lt1")}},
        "lastRow": {"bold": True},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{284E427A-3D55-4303-BF80-6455036E1DE7}": {
        "name": "Themed Style 1 - Accent 2",
        "wholeTbl": {"fill": "accent2", "text": "dk1"},
        "firstRow": {"text": "lt1", "bold": True, "borders": {"bottom": (1.0, "lt1")}},
        "lastRow": {"bold": True},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{69C7853C-536D-4A76-A0AE-DD22124D55A5}": {
        "name": "Themed Style 1 - Accent 3",
        "wholeTbl": {"fill": "accent3", "text": "dk1"},
        "firstRow": {"text": "lt1", "bold": True, "borders": {"bottom": (1.0, "lt1")}},
        "lastRow": {"bold": True},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{775DCB02-9BB8-47FD-8907-85C794F793BA}": {
        "name": "Themed Style 1 - Accent 4",
        "wholeTbl": {"fill": "accent4", "text": "dk1"},
        "firstRow": {"text": "lt1", "bold": True, "borders": {"bottom": (1.0, "lt1")}},
        "lastRow": {"bold": True},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{35758FB7-9AC5-4552-8A53-C91805E547FA}": {
        "name": "Themed Style 1 - Accent 5",
        "wholeTbl": {"fill": "accent5", "text": "dk1"},
        "firstRow": {"text": "lt1", "bold": True, "borders": {"bottom": (1.0, "lt1")}},
        "lastRow": {"bold": True},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{08FB837D-C827-4EFA-A057-4D05807E0F7C}": {
        "name": "Themed Style 1 - Accent 6",
        "wholeTbl": {"fill": "accent6", "text": "dk1"},
        "firstRow": {"text": "lt1", "bold": True, "borders": {"bottom": (1.0, "lt1")}},
        "lastRow": {"bold": True},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}": {
        "name": "Themed Style 2 - Accent 1",
        "wholeTbl": {"fill": "accent1", "text": "lt1"},
        "band1H": {"fill": "accent1/tint90"},
        "band1V": {"fill": "accent1/tint90"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.5, "lt1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "lt1")}},
        "firstCol": {"bold": True, "borders": {"right": (1.0, "lt1")}},
        "lastCol": {"bold": True, "borders": {"left": (1.0, "lt1")}},
    },
    "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}": {
        "name": "Themed Style 2 - Accent 2",
        "wholeTbl": {"fill": "accent2", "text": "lt1"},
        "band1H": {"fill": "accent2/tint90"},
        "band1V": {"fill": "accent2/tint90"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.5, "lt1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "lt1")}},
        "firstCol": {"bold": True, "borders": {"right": (1.0, "lt1")}},
        "lastCol": {"bold": True, "borders": {"left": (1.0, "lt1")}},
    },
    "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}": {
        "name": "Themed Style 2 - Accent 3",
        "wholeTbl": {"fill": "accent3", "text": "lt1"},
        "band1H": {"fill": "accent3/tint90"},
        "band1V": {"fill": "accent3/tint90"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.5, "lt1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "lt1")}},
        "firstCol": {"bold": True, "borders": {"right": (1.0, "lt1")}},
        "lastCol": {"bold": True, "borders": {"left": (1.0, "lt1")}},
    },
    "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}": {
        "name": "Themed Style 2 - Accent 4",
        "wholeTbl": {"fill": "accent4", "text": "lt1"},
        "band1H": {"fill": "accent4/tint90"},
        "band1V": {"fill": "accent4/tint90"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.5, "lt1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "lt1")}},
        "firstCol": {"bold": True, "borders": {"right": (1.0, "lt1")}},
        "lastCol": {"bold": True, "borders": {"left": (1.0, "lt1")}},
    },
    "{327F97BB-C833-4FB7-BDE5-3F7075034690}": {
        "name": "Themed Style 2 - Accent 5",
        "wholeTbl": {"fill": "accent5", "text": "lt1"},
        "band1H": {"fill": "accent5/tint90"},
        "band1V": {"fill": "accent5/tint90"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.5, "lt1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "lt1")}},
        "firstCol": {"bold": True, "borders": {"right": (1.0, "lt1")}},
        "lastCol": {"bold": True, "borders": {"left": (1.0, "lt1")}},
    },
    "{638B1855-1B75-4FBE-930C-398BA8C253C6}": {
        "name": "Themed Style 2 - Accent 6",
        "wholeTbl": {"fill": "accent6", "text": "lt1"},
        "band1H": {"fill": "accent6/tint90"},
        "band1V": {"fill": "accent6/tint90"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.5, "lt1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "lt1")}},
        "firstCol": {"bold": True, "borders": {"right": (1.0, "lt1")}},
        "lastCol": {"bold": True, "borders": {"left": (1.0, "lt1")}},
    },
    "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}": {
        "name": "Light Style 1",
        "wholeTbl": {"text": "dk1", "borders": {"top": (1.0, "dk1"), "bottom": (1.0, "dk1")}},
        "band1H": {"fill": "dk1/alpha20"},
        "band1V": {"fill": "dk1/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}": {
        "name": "Light Style 1 - Accent 1",
        "wholeTbl": {"text": "dk1", "borders": {"top": (1.0, "accent1"), "bottom": (1.0, "accent1")}},
        "band1H": {"fill": "accent1/alpha20"},
        "band1V": {"fill": "accent1/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.0, "accent1")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "accent1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}": {
        "name": "Light Style 1 - Accent 2",
        "wholeTbl": {"text": "dk1", "borders": {"top": (1.0, "accent2"), "bottom": (1.0, "accent2")}},
        "band1H": {"fill": "accent2/alpha20"},
        "band1V": {"fill": "accent2/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.0, "accent2")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "accent2")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}": {
        "name": "Light Style 1 - Accent 3",
        "wholeTbl": {"text": "dk1", "borders": {"top": (1.0, "accent3"), "bottom": (1.0, "accent3")}},
        "band1H": {"fill": "accent3/alpha20"},
        "band1V": {"fill": "accent3/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.0, "accent3")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "accent3")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{D27102A9-8310-4765-A935-A1911B00CA55}": {
        "name": "Light Style 1 - Accent 4",
        "wholeTbl": {"text": "dk1", "borders": {"top": (1.0, "accent4"), "bottom": (1.0, "accent4")}},
        "band1H": {"fill": "accent4/alpha20"},
        "band1V": {"fill": "accent4/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.0, "accent4")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "accent4")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}": {
        "name": "Light Style 1 - Accent 5",
        "wholeTbl": {"text": "dk1", "borders": {"top": (1.0, "accent5"), "bottom": (1.0, "accent5")}},
        "band1H": {"fill": "accent5/alpha20"},
        "band1V": {"fill": "accent5/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.0, "accent5")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "accent5")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{68D230F3-CF80-4859-8CE7-A43EE81993B5}": {
        "name": "Light Style 1 - Accent 6",
        "wholeTbl": {"text": "dk1", "borders": {"top": (1.0, "accent6"), "bottom": (1.0, "accent6")}},
        "band1H": {"fill": "accent6/alpha20"},
        "band1V": {"fill": "accent6/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (1.0, "accent6")}},
        "lastRow": {"bold": True, "borders": {"top": (1.0, "accent6")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}": {
        "name": "Light Style 2",
        "wholeTbl": {"text": "dk1"},
        "firstRow": {"fill": "dk1", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}": {
        "name": "Light Style 2 - Accent 1",
        "wholeTbl": {"text": "dk1"},
        "firstRow": {"fill": "accent1", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}": {
        "name": "Light Style 2 - Accent 2",
        "wholeTbl": {"text": "dk1"},
        "firstRow": {"fill": "accent2", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent2")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{F2DE63D5-997A-4646-A377-4702673A728D}": {
        "name": "Light Style 2 - Accent 3",
        "wholeTbl": {"text": "dk1"},
        "firstRow": {"fill": "accent3", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent3")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}": {
        "name": "Light Style 2 - Accent 4",
        "wholeTbl": {"text": "dk1"},
        "firstRow": {"fill": "accent4", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent4")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{5A111915-BE36-4E01-A7E5-04B1672EAD32}": {
        "name": "Light Style 2 - Accent 5",
        "wholeTbl": {"text": "dk1"},
        "firstRow": {"fill": "accent5", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent5")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{912C8C85-51F0-491E-9774-3900AFEF0FD7}": {
        "name": "Light Style 2 - Accent 6",
        "wholeTbl": {"text": "dk1"},
        "firstRow": {"fill": "accent6", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent6")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{616DA210-FB5B-4158-B5E0-FEB733F419BA}": {
        "name": "Light Style 3",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "dk1"), "right": (1.0, "dk1"), "top": (1.0, "dk1"), "bottom": (1.0, "dk1"), "insideH": (1.0, "dk1"), "insideV": (1.0, "dk1")}},
        "band1H": {"fill": "dk1/alpha20"},
        "band1V": {"fill": "dk1/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (2.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}": {
        "name": "Light Style 3 - Accent 1",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "accent1"), "right": (1.0, "accent1"), "top": (1.0, "accent1"), "bottom": (1.0, "accent1"), "insideH": (1.0, "accent1"), "insideV": (1.0, "accent1")}},
        "band1H": {"fill": "accent1/alpha20"},
        "band1V": {"fill": "accent1/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (2.0, "accent1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{5DA37D80-6434-44D0-A028-1B22A696006F}": {
        "name": "Light Style 3 - Accent 2",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "accent2"), "right": (1.0, "accent2"), "top": (1.0, "accent2"), "bottom": (1.0, "accent2"), "insideH": (1.0, "accent2"), "insideV": (1.0, "accent2")}},
        "band1H": {"fill": "accent2/alpha20"},
        "band1V": {"fill": "accent2/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (2.0, "accent2")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent2")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{8799B23B-EC83-4686-B30A-512413B5E67A}": {
        "name": "Light Style 3 - Accent 3",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "accent3"), "right": (1.0, "accent3"), "top": (1.0, "accent3"), "bottom": (1.0, "accent3"), "insideH": (1.0, "accent3"), "insideV": (1.0, "accent3")}},
        "band1H": {"fill": "accent3/alpha20"},
        "band1V": {"fill": "accent3/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (2.0, "accent3")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent3")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}": {
        "name": "Light Style 3 - Accent 4",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "accent4"), "right": (1.0, "accent4"), "top": (1.0, "accent4"), "bottom": (1.0, "accent4"), "insideH": (1.0, "accent4"), "insideV": (1.0, "accent4")}},
        "band1H": {"fill": "accent4/alpha20"},
        "band1V": {"fill": "accent4/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (2.0, "accent4")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent4")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}": {
        "name": "Light Style 3 - Accent 5",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "accent5"), "right": (1.0, "accent5"), "top": (1.0, "accent5"), "bottom": (1.0, "accent5"), "insideH": (1.0, "accent5"), "insideV": (1.0, "accent5")}},
        "band1H": {"fill": "accent5/alpha20"},
        "band1V": {"fill": "accent5/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (2.0, "accent5")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent5")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}": {
        "name": "Light Style 3 - Accent 6",
        "wholeTbl": {"text": "dk1", "borders": {"left": (1.0, "accent6"), "right": (1.0, "accent6"), "top": (1.0, "accent6"), "bottom": (1.0, "accent6"), "insideH": (1.0, "accent6"), "insideV": (1.0, "accent6")}},
        "band1H": {"fill": "accent6/alpha20"},
        "band1V": {"fill": "accent6/alpha20"},
        "firstRow": {"bold": True, "borders": {"bottom": (2.0, "accent6")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent6")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}": {
        "name": "Medium Style 1",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"left": (1.0, "dk1"), "right": (1.0, "dk1"), "top": (1.0, "dk1"), "bottom": (1.0, "dk1"), "insideH": (1.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "dk1", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{B301B821-A1FF-4177-AEE7-76D212191A09}": {
        "name": "Medium Style 1 - Accent 1",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"left": (1.0, "accent1"), "right": (1.0, "accent1"), "top": (1.0, "accent1"), "bottom": (1.0, "accent1"), "insideH": (1.0, "accent1")}},
        "band1H": {"fill": "accent1/tint20"},
        "band1V": {"fill": "accent1/tint20"},
        "firstRow": {"fill": "accent1", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}": {
        "name": "Medium Style 1 - Accent 2",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"left": (1.0, "accent2"), "right": (1.0, "accent2"), "top": (1.0, "accent2"), "bottom": (1.0, "accent2"), "insideH": (1.0, "accent2")}},
        "band1H": {"fill": "accent2/tint20"},
        "band1V": {"fill": "accent2/tint20"},
        "firstRow": {"fill": "accent2", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent2")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}": {
        "name": "Medium Style 1 - Accent 3",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"left": (1.0, "accent3"), "right": (1.0, "accent3"), "top": (1.0, "accent3"), "bottom": (1.0, "accent3"), "insideH": (1.0, "accent3")}},
        "band1H": {"fill": "accent3/tint20"},
        "band1V": {"fill": "accent3/tint20"},
        "firstRow": {"fill": "accent3", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent3")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{1E171933-4619-4E11-9A3F-F7608DF75F80}": {
        "name": "Medium Style 1 - Accent 4",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"left": (1.0, "accent4"), "right": (1.0, "accent4"), "top": (1.0, "accent4"), "bottom": (1.0, "accent4"), "insideH": (1.0, "accent4")}},
        "band1H": {"fill": "accent4/tint20"},
        "band1V": {"fill": "accent4/tint20"},
        "firstRow": {"fill": "accent4", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent4")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}": {
        "name": "Medium Style 1 - Accent 5",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"left": (1.0, "accent5"), "right": (1.0, "accent5"), "top": (1.0, "accent5"), "bottom": (1.0, "accent5"), "insideH": (1.0, "accent5")}},
        "band1H": {"fill": "accent5/tint20"},
        "band1V": {"fill": "accent5/tint20"},
        "firstRow": {"fill": "accent5", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent5")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{10A1B5D5-9B99-4C35-A422-299274C87663}": {
        "name": "Medium Style 1 - Accent 6",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"left": (1.0, "accent6"), "right": (1.0, "accent6"), "top": (1.0, "accent6"), "bottom": (1.0, "accent6"), "insideH": (1.0, "accent6")}},
        "band1H": {"fill": "accent6/tint20"},
        "band1V": {"fill": "accent6/tint20"},
        "firstRow": {"fill": "accent6", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "accent6")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}": {
        "name": "Medium Style 2",
        "wholeTbl": {"fill": "dk1/tint20", "text": "dk1", "borders": {"left": (1.0, "lt1"), "right": (1.0, "lt1"), "top": (1.0, "lt1"), "bottom": (1.0, "lt1"), "insideH": (1.0, "lt1"), "insideV": (1.0, "lt1")}},
        "band1H": {"fill": "dk1/tint40"},
        "band1V": {"fill": "dk1/tint40"},
        "firstRow": {"fill": "dk1", "text": "lt1", "bold": True, "borders": {"bottom": (3.0, "lt1")}},
        "lastRow": {"fill": "dk1", "text": "lt1", "bold": True, "borders": {"top": (3.0, "lt1")}},
        "firstCol": {"fill": "dk1", "text": "lt1", "bold": True},
        "lastCol": {"fill": "dk1", "text": "lt1", "bold": True},
    },
    "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}": {
        "name": "Medium Style 2 - Accent 1",
        "wholeTbl": {"fill": "accent1/tint20", "text": "dk1", "borders": {"left": (1.0, "lt1"), "right": (1.0, "lt1"), "top": (1.0, "lt1"), "bottom": (1.0, "lt1"), "insideH": (1.0, "lt1"), "insideV": (1.0, "lt1")}},
        "band1H": {"fill": "accent1/tint40"},
        "band1V": {"fill": "accent1/tint40"},
        "firstRow": {"fill": "accent1", "text": "lt1", "bold": True, "borders": {"bottom": (3.0, "lt1")}},
        "lastRow": {"fill": "accent1", "text": "lt1", "bold": True, "borders": {"top": (3.0, "lt1")}},
        "firstCol": {"fill": "accent1", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent1", "text": "lt1", "bold": True},
    },
    "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}": {
        "name": "Medium Style 2 - Accent 2",
        "wholeTbl": {"fill": "accent2/tint20", "text": "dk1", "borders": {"left": (1.0, "lt1"), "right": (1.0, "lt1"), "top": (1.0, "lt1"), "bottom": (1.0, "lt1"), "insideH": (1.0, "lt1"), "insideV": (1.0, "lt1")}},
        "band1H": {"fill": "accent2/tint40"},
        "band1V": {"fill": "accent2/tint40"},
        "firstRow": {"fill": "accent2", "text": "lt1", "bold": True, "borders": {"bottom": (3.0, "lt1")}},
        "lastRow": {"fill": "accent2", "text": "lt1", "bold": True, "borders": {"top": (3.0, "lt1")}},
        "firstCol": {"fill": "accent2", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent2", "text": "lt1", "bold": True},
    },
    "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}": {
        "name": "Medium Style 2 - Accent 3",
        "wholeTbl": {"fill": "accent3/tint20", "text": "dk1", "borders": {"left": (1.0, "lt1"), "right": (1.0, "lt1"), "top": (1.0, "lt1"), "bottom": (1.0, "lt1"), "insideH": (1.0, "lt1"), "insideV": (1.0, "lt1")}},
        "band1H": {"fill": "accent3/tint40"},
        "band1V": {"fill": "accent3/tint40"},
        "firstRow": {"fill": "accent3", "text": "lt1", "bold": True, "borders": {"bottom": (3.0, "lt1")}},
        "lastRow": {"fill": "accent3", "text": "lt1", "bold": True, "borders": {"top": (3.0, "lt1")}},
        "firstCol": {"fill": "accent3", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent3", "text": "lt1", "bold": True},
    },
    "{00A15C55-8517-42AA-B614-E9B94910E393}": {
        "name": "Medium Style 2 - Accent 4",
        "wholeTbl": {"fill": "accent4/tint20", "text": "dk1", "borders": {"left": (1.0, "lt1"), "right": (1.0, "lt1"), "top": (1.0, "lt1"), "bottom": (1.0, "lt1"), "insideH": (1.0, "lt1"), "insideV": (1.0, "lt1")}},
        "band1H": {"fill": "accent4/tint40"},
        "band1V": {"fill": "accent4/tint40"},
        "firstRow": {"fill": "accent4", "text": "lt1", "bold": True, "borders": {"bottom": (3.0, "lt1")}},
        "lastRow": {"fill": "accent4", "text": "lt1", "bold": True, "borders": {"top": (3.0, "lt1")}},
        "firstCol": {"fill": "accent4", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent4", "text": "lt1", "bold": True},
    },
    "{7DF18680-E054-41AD-8BC1-D1AEF772440D}": {
        "name": "Medium Style 2 - Accent 5",
        "wholeTbl": {"fill": "accent5/tint20", "text": "dk1", "borders": {"left": (1.0, "lt1"), "right": (1.0, "lt1"), "top": (1.0, "lt1"), "bottom": (1.0, "lt1"), "insideH": (1.0, "lt1"), "insideV": (1.0, "lt1")}},
        "band1H": {"fill": "accent5/tint40"},
        "band1V": {"fill": "accent5/tint40"},
        "firstRow": {"fill": "accent5", "text": "lt1", "bold": True, "borders": {"bottom": (3.0, "lt1")}},
        "lastRow": {"fill": "accent5", "text": "lt1", "bold": True, "borders": {"top": (3.0, "lt1")}},
        "firstCol": {"fill": "accent5", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent5", "text": "lt1", "bold": True},
    },
    "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}": {
        "name": "Medium Style 2 - Accent 6",
        "wholeTbl": {"fill": "accent6/tint20", "text": "dk1", "borders": {"left": (1.0, "lt1"), "right": (1.0, "lt1"), "top": (1.0, "lt1"), "bottom": (1.0, "lt1"), "insideH": (1.0, "lt1"), "insideV": (1.0, "lt1")}},
        "band1H": {"fill": "accent6/tint40"},
        "band1V": {"fill": "accent6/tint40"},
        "firstRow": {"fill": "accent6", "text": "lt1", "bold": True, "borders": {"bottom": (3.0, "lt1")}},
        "lastRow": {"fill": "accent6", "text": "lt1", "bold": True, "borders": {"top": (3.0, "lt1")}},
        "firstCol": {"fill": "accent6", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent6", "text": "lt1", "bold": True},
    },
    "{8EC20E35-A176-4012-BC5E-935CFFF8708E}": {
        "name": "Medium Style 3",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"top": (2.0, "dk1"), "bottom": (2.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "dk1", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"fill": "dk1", "text": "lt1", "bold": True},
        "lastCol": {"fill": "dk1", "text": "lt1", "bold": True},
    },
    "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}": {
        "name": "Medium Style 3 - Accent 1",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"top": (2.0, "dk1"), "bottom": (2.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "accent1", "text": "lt1", "bold": True, "borders": {"bottom": (2.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"fill": "accent1", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent1", "text": "lt1", "bold": True},
    },
    "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}": {
        "name": "Medium Style 3 - Accent 2",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"top": (2.0, "dk1"), "bottom": (2.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "accent2", "text": "lt1", "bold": True, "borders": {"bottom": (2.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"fill": "accent2", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent2", "text": "lt1", "bold": True},
    },
    "{EB344D84-9AFB-497E-A393-DC336BA19D2E}": {
        "name": "Medium Style 3 - Accent 3",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"top": (2.0, "dk1"), "bottom": (2.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "accent3", "text": "lt1", "bold": True, "borders": {"bottom": (2.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"fill": "accent3", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent3", "text": "lt1", "bold": True},
    },
    "{EB9631B5-78F2-41C9-869B-9F39066F8104}": {
        "name": "Medium Style 3 - Accent 4",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"top": (2.0, "dk1"), "bottom": (2.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "accent4", "text": "lt1", "bold": True, "borders": {"bottom": (2.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"fill": "accent4", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent4", "text": "lt1", "bold": True},
    },
    "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}": {
        "name": "Medium Style 3 - Accent 5",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"top": (2.0, "dk1"), "bottom": (2.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "accent5", "text": "lt1", "bold": True, "borders": {"bottom": (2.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"fill": "accent5", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent5", "text": "lt1", "bold": True},
    },
    "{2A488322-F2BA-4B5B-9748-0D474271808F}": {
        "name": "Medium Style 3 - Accent 6",
        "wholeTbl": {"fill": "lt1", "text": "dk1", "borders": {"top": (2.0, "dk1"), "bottom": (2.0, "dk1")}},
        "band1H": {"fill": "dk1/tint20"},
        "band1V": {"fill": "dk1/tint20"},
        "firstRow": {"fill": "accent6", "text": "lt1", "bold": True, "borders": {"bottom": (2.0, "dk1")}},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"fill": "accent6", "text": "lt1", "bold": True},
        "lastCol": {"fill": "accent6", "text": "lt1", "bold": True},
    },
    "{D7AC3CCA-C797-4891-BE02-D94E43425B78}": {
        "name": "Medium Style 4",
        "wholeTbl": {"fill": "dk1/tint20", "text": "dk1", "borders": {"left": (1.0, "dk1/alpha80"), "right": (1.0, "dk1"), "top": (1.0, "dk1"), "bottom": (1.0, "dk1"), "insideH": (1.0, "dk1"), "insideV": (1.0, "dk1")}},
        "band1H": {"fill": "dk1/tint40"},
        "band1V": {"fill": "dk1/tint40"},
        "firstRow": {"bold": True},
        "lastRow": {"bold": True, "borders": {"top": (2.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{69CF1AB2-1976-4502-BF36-3FF5EA218861}": {
        "name": "Medium Style 4 - Accent 1",
        "wholeTbl": {"fill": "accent1/tint20", "text": "dk1", "borders": {"left": (1.0, "accent1"), "right": (1.0, "accent1"), "top": (1.0, "accent1"), "bottom": (1.0, "accent1"), "insideH": (1.0, "accent1"), "insideV": (1.0, "accent1")}},
        "band1H": {"fill": "accent1/tint40"},
        "band1V": {"fill": "accent1/tint40"},
        "firstRow": {"bold": True},
        "lastRow": {"bold": True, "borders": {"top": (2.0, "accent1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{8A107856-5554-42FB-B03E-39F5DBC370BA}": {
        "name": "Medium Style 4 - Accent 2",
        "wholeTbl": {"fill": "accent2/tint20", "text": "dk1", "borders": {"left": (1.0, "accent2"), "right": (1.0, "accent2"), "top": (1.0, "accent2"), "bottom": (1.0, "accent2"), "insideH": (1.0, "accent2"), "insideV": (1.0, "accent2")}},
        "band1H": {"fill": "accent2/tint40"},
        "band1V": {"fill": "accent2/tint40"},
        "firstRow": {"bold": True},
        "lastRow": {"bold": True, "borders": {"top": (2.0, "accent2")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{0505E3EF-67EA-436B-97B2-0124C06EBD24}": {
        "name": "Medium Style 4 - Accent 3",
        "wholeTbl": {"fill": "accent3/tint20", "text": "dk1", "borders": {"left": (1.0, "accent3"), "right": (1.0, "accent3"), "top": (1.0, "accent3"), "bottom": (1.0, "accent3"), "insideH": (1.0, "accent3"), "insideV": (1.0, "accent3")}},
        "band1H": {"fill": "accent3/tint40"},
        "band1V": {"fill": "accent3/tint40"},
        "firstRow": {"bold": True},
        "lastRow": {"bold": True, "borders": {"top": (2.0, "accent3")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}": {
        "name": "Medium Style 4 - Accent 4",
        "wholeTbl": {"fill": "accent4/tint20", "text": "dk1", "borders": {"left": (1.0, "accent4"), "right": (1.0, "accent4"), "top": (1.0, "accent4"), "bottom": (1.0, "accent4"), "insideH": (1.0, "accent4"), "insideV": (1.0, "accent4")}},
        "band1H": {"fill": "accent4/tint40"},
        "band1V": {"fill": "accent4/tint40"},
        "firstRow": {"bold": True},
        "lastRow": {"bold": True, "borders": {"top": (2.0, "accent4")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{22838BEF-8BB2-4498-84A7-C5851F593DF1}": {
        "name": "Medium Style 4 - Accent 5",
        "wholeTbl": {"fill": "accent5/tint20", "text": "dk1", "borders": {"left": (1.0, "accent5"), "right": (1.0, "accent5"), "top": (1.0, "accent5"), "bottom": (1.0, "accent5"), "insideH": (1.0, "accent5"), "insideV": (1.0, "accent5")}},
        "band1H": {"fill": "accent5/tint40"},
        "band1V": {"fill": "accent5/tint40"},
        "firstRow": {"bold": True},
        "lastRow": {"bold": True, "borders": {"top": (2.0, "accent5")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}": {
        "name": "Medium Style 4 - Accent 6",
        "wholeTbl": {"fill": "accent6/tint20", "text": "dk1", "borders": {"left": (1.0, "accent6"), "right": (1.0, "accent6"), "top": (1.0, "accent6"), "bottom": (1.0, "accent6"), "insideH": (1.0, "accent6"), "insideV": (1.0, "accent6")}},
        "band1H": {"fill": "accent6/tint40"},
        "band1V": {"fill": "accent6/tint40"},
        "firstRow": {"bold": True},
        "lastRow": {"bold": True, "borders": {"top": (2.0, "accent6")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}": {
        "name": "Dark Style 1",
        "wholeTbl": {"fill": "dk1/tint20", "text": "dk1/tint20"},
        "band1H": {"fill": "dk1/tint40"},
        "band1V": {"fill": "dk1/tint40"},
        "firstRow": {"fill": "dk1", "text": "lt1", "bold": True, "borders": {"bottom": (2.0, "lt1")}},
        "lastRow": {"fill": "accent3/tint95", "text": "lt1", "bold": True, "borders": {"top": (2.0, "lt1")}},
        "firstCol": {"fill": "accent3/tint95", "text": "lt1", "bold": True, "borders": {"right": (2.0, "lt1")}},
        "lastCol": {"fill": "accent3/tint95", "text": "lt1", "bold": True, "borders": {"left": (2.0, "lt1")}},
    },
    "{125E5076-3810-47DD-B79F-674D7AD40C01}": {
        "name": "Dark Style 1 - Accent 1",
        "wholeTbl": {"fill": "accent1", "text": "lt1"},
        "band1H": {"fill": "accent1/shade60"},
        "band1V": {"fill": "accent1/shade60"},
        "firstRow": {"fill": "dk1", "bold": True, "borders": {"bottom": (2.0, "lt1")}},
        "lastRow": {"fill": "accent1/shade40", "bold": True, "borders": {"top": (2.0, "lt1")}},
        "firstCol": {"fill": "accent1/shade60", "bold": True, "borders": {"right": (2.0, "lt1")}},
        "lastCol": {"fill": "accent1/shade60", "bold": True, "borders": {"left": (2.0, "lt1")}},
    },
    "{37CE84F3-28C3-443E-9E96-99CF82512B78}": {
        "name": "Dark Style 1 - Accent 2",
        "wholeTbl": {"fill": "accent2", "text": "lt1"},
        "band1H": {"fill": "accent2/shade60"},
        "band1V": {"fill": "accent2/shade60"},
        "firstRow": {"fill": "dk1", "bold": True, "borders": {"bottom": (2.0, "lt1")}},
        "lastRow": {"fill": "accent2/shade40", "bold": True, "borders": {"top": (2.0, "lt1")}},
        "firstCol": {"fill": "accent2/shade60", "bold": True, "borders": {"right": (2.0, "lt1")}},
        "lastCol": {"fill": "accent2/shade60", "bold": True, "borders": {"left": (2.0, "lt1")}},
    },
    "{D03447BB-5D67-496B-8E87-E561075AD55C}": {
        "name": "Dark Style 1 - Accent 3",
        "wholeTbl": {"fill": "accent3", "text": "lt1"},
        "band1H": {"fill": "accent3/shade60"},
        "band1V": {"fill": "accent3/shade60"},
        "firstRow": {"fill": "dk1", "bold": True, "borders": {"bottom": (2.0, "lt1")}},
        "lastRow": {"fill": "accent3/shade40", "bold": True, "borders": {"top": (2.0, "lt1")}},
        "firstCol": {"fill": "accent3/shade60", "bold": True, "borders": {"right": (2.0, "lt1")}},
        "lastCol": {"fill": "accent3/shade60", "bold": True, "borders": {"left": (2.0, "lt1")}},
    },
    "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}": {
        "name": "Dark Style 1 - Accent 4",
        "wholeTbl": {"fill": "accent4", "text": "lt1"},
        "band1H": {"fill": "accent4/shade60"},
        "band1V": {"fill": "accent4/shade60"},
        "firstRow": {"fill": "dk1", "bold": True, "borders": {"bottom": (2.0, "lt1")}},
        "lastRow": {"fill": "accent4/shade40", "bold": True, "borders": {"top": (2.0, "lt1")}},
        "firstCol": {"fill": "accent4/shade60", "bold": True, "borders": {"right": (2.0, "lt1")}},
        "lastCol": {"fill": "accent4/shade60", "bold": True, "borders": {"left": (2.0, "lt1")}},
    },
    "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}": {
        "name": "Dark Style 1 - Accent 5",
        "wholeTbl": {"fill": "accent5", "text": "lt1"},
        "band1H": {"fill": "accent5/shade60"},
        "band1V": {"fill": "accent5/shade60"},
        "firstRow": {"fill": "dk1", "bold": True, "borders": {"bottom": (2.0, "lt1")}},
        "lastRow": {"fill": "accent5/shade40", "bold": True, "borders": {"top": (2.0, "lt1")}},
        "firstCol": {"fill": "accent5/shade60", "bold": True, "borders": {"right": (2.0, "lt1")}},
        "lastCol": {"fill": "accent5/shade60", "bold": True, "borders": {"left": (2.0, "lt1")}},
    },
    "{AF606853-7671-496A-8E4F-DF71F8EC918B}": {
        "name": "Dark Style 1 - Accent 6",
        "wholeTbl": {"fill": "accent6", "text": "lt1"},
        "band1H": {"fill": "accent6/shade60"},
        "band1V": {"fill": "accent6/shade60"},
        "firstRow": {"fill": "dk1", "bold": True, "borders": {"bottom": (2.0, "lt1")}},
        "lastRow": {"fill": "accent6/shade40", "bold": True, "borders": {"top": (2.0, "lt1")}},
        "firstCol": {"fill": "accent6/shade60", "bold": True, "borders": {"right": (2.0, "lt1")}},
        "lastCol": {"fill": "accent6/shade60", "bold": True, "borders": {"left": (2.0, "lt1")}},
    },
    "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}": {
        "name": "Dark Style 2",
        "wholeTbl": {"fill": "dk1/tint20", "text": "dk1"},
        "band1H": {"fill": "dk1/tint40"},
        "band1V": {"fill": "dk1/tint40"},
        "firstRow": {"fill": "dk1", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}": {
        "name": "Dark Style 2 - Accent 1/Accent 2",
        "wholeTbl": {"fill": "accent1/tint20", "text": "dk1"},
        "band1H": {"fill": "accent1/tint40"},
        "band1V": {"fill": "accent1/tint40"},
        "firstRow": {"fill": "accent2", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}": {
        "name": "Dark Style 2 - Accent 3/Accent 4",
        "wholeTbl": {"fill": "accent3/tint20", "text": "dk1"},
        "band1H": {"fill": "accent3/tint40"},
        "band1V": {"fill": "accent3/tint40"},
        "firstRow": {"fill": "accent4", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
    "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}": {
        "name": "Dark Style 2 - Accent 5/Accent 6",
        "wholeTbl": {"fill": "accent5/tint20", "text": "dk1"},
        "band1H": {"fill": "accent5/tint40"},
        "band1V": {"fill": "accent5/tint40"},
        "firstRow": {"fill": "accent6", "text": "lt1", "bold": True},
        "lastRow": {"bold": True, "borders": {"top": (3.0, "dk1")}},
        "firstCol": {"bold": True},
        "lastCol": {"bold": True},
    },
}

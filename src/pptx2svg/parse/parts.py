"""Part readers: presentation.xml, slides, layouts, masters and themes.

Walks the package graph from ``presentation.xml``: the slide id list gives slide parts in
presentation order, each slide's relationships point at its layout, each layout at its
master, and each master at its theme.  Layouts and masters are read once and shared, so a
50-slide deck built on one master parses that master once.
"""

from __future__ import annotations

from xml.etree.ElementTree import Element

from ..opc import (
    OpcPackage,
    REL_SLIDE,
    REL_SLIDE_LAYOUT,
    REL_SLIDE_MASTER,
    REL_THEME,
)
from ..xmlutil import attr, child, children, is_true, ns_attr, num_attr
from .drawing import parse_color, parse_effect_list, parse_fill, parse_line, parse_style_reference
from .shapes import parse_shape_tree
from .source import (
    SourceBackground,
    SourceColorMap,
    SourceFontScheme,
    SourceFormatScheme,
    SourcePresentation,
    SourceSlide,
    SourceSlideLayout,
    SourceSlideMaster,
    SourceTheme,
)
from .text import parse_text_style

#: ``p:clrMap`` slots, in the order the schema declares them.
COLOR_MAP_SLOTS = (
    "bg1",
    "tx1",
    "bg2",
    "tx2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)

#: ``a:clrScheme`` children, in schema order.
COLOR_SCHEME_KEYS = (
    "dk1",
    "lt1",
    "dk2",
    "lt2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
)


def read_presentation(package: OpcPackage) -> SourcePresentation:
    """Read a whole deck into the source model."""
    presentation_path = package.presentation_part()
    if presentation_path is None:
        raise ValueError("package has no presentation part; is this a .pptx?")

    root = package.read_xml(presentation_path)
    if root is None:
        raise ValueError(f"presentation part {presentation_path!r} is missing")

    sld_sz = child(root, "sldSz")
    presentation = SourcePresentation(
        part_path=presentation_path,
        slide_width=num_attr(sld_sz, "cx") or 9144000,
        slide_height=num_attr(sld_sz, "cy") or 6858000,
        default_text_style=parse_text_style(child(root, "defaultTextStyle")),
    )

    for number, slide_path in enumerate(_slide_paths(package, presentation_path, root), start=1):
        slide = read_slide(package, slide_path, number)
        if slide is None:
            continue
        presentation.slides.append(slide)
        _ensure_ancestry(package, presentation, slide)

    return presentation


def _slide_paths(package: OpcPackage, presentation_path: str, root: Element) -> list[str]:
    """Slide parts in presentation order, from ``p:sldIdLst``."""
    relationships = package.relationships(presentation_path)
    paths: list[str] = []
    for sld_id in children(child(root, "sldIdLst"), "sldId"):
        rel_id = ns_attr(sld_id, "id")
        relationship = relationships.get(rel_id) if rel_id else None
        if relationship is not None and relationship.target_part is not None:
            paths.append(relationship.target_part)
    if paths:
        return paths
    # Malformed decks without a slide id list: fall back to relationship order.
    return package.related_parts_of_type(presentation_path, REL_SLIDE)


def _ensure_ancestry(
    package: OpcPackage, presentation: SourcePresentation, slide: SourceSlide
) -> None:
    """Read the slide's layout, master and theme once each, caching by part path."""
    layout_path = slide.layout_part_path
    if layout_path is None or layout_path in presentation.layouts:
        layout = presentation.layouts.get(layout_path) if layout_path else None
    else:
        layout = read_slide_layout(package, layout_path)
        if layout is not None:
            presentation.layouts[layout_path] = layout

    if layout is None:
        return

    master_path = layout.master_part_path
    if master_path is None:
        return
    master = presentation.masters.get(master_path)
    if master is None:
        master = read_slide_master(package, master_path)
        if master is None:
            return
        presentation.masters[master_path] = master

    theme_path = master.theme_part_path
    if theme_path is not None and theme_path not in presentation.themes:
        theme = read_theme(package, theme_path)
        if theme is not None:
            presentation.themes[theme_path] = theme


# --------------------------------------------------------------------------------------
# Slide / layout / master
# --------------------------------------------------------------------------------------


def read_slide(package: OpcPackage, part_path: str, slide_number: int) -> SourceSlide | None:
    root = package.read_xml(part_path)
    if root is None:
        return None
    common = child(root, "cSld")
    show_master = attr(root, "showMasterSp")

    return SourceSlide(
        part_path=part_path,
        shapes=parse_shape_tree(child(common, "spTree")),
        background=parse_background(child(common, "bg")),
        color_map_override=parse_color_map_override(child(root, "clrMapOvr")),
        layout_part_path=package.first_related_part(part_path, REL_SLIDE_LAYOUT),
        show_master_shapes=True if show_master is None else is_true(show_master),
        slide_number=slide_number,
    )


def read_slide_layout(package: OpcPackage, part_path: str) -> SourceSlideLayout | None:
    root = package.read_xml(part_path)
    if root is None:
        return None
    common = child(root, "cSld")
    show_master = attr(root, "showMasterSp")

    return SourceSlideLayout(
        part_path=part_path,
        shapes=parse_shape_tree(child(common, "spTree")),
        background=parse_background(child(common, "bg")),
        color_map_override=parse_color_map_override(child(root, "clrMapOvr")),
        master_part_path=package.first_related_part(part_path, REL_SLIDE_MASTER),
        show_master_shapes=True if show_master is None else is_true(show_master),
        layout_type=attr(root, "type"),
    )


def read_slide_master(package: OpcPackage, part_path: str) -> SourceSlideMaster | None:
    root = package.read_xml(part_path)
    if root is None:
        return None
    common = child(root, "cSld")
    tx_styles = child(root, "txStyles")

    return SourceSlideMaster(
        part_path=part_path,
        shapes=parse_shape_tree(child(common, "spTree")),
        background=parse_background(child(common, "bg")),
        theme_part_path=package.first_related_part(part_path, REL_THEME),
        color_map=parse_color_map(child(root, "clrMap")),
        title_style=parse_text_style(child(tx_styles, "titleStyle")),
        body_style=parse_text_style(child(tx_styles, "bodyStyle")),
        other_style=parse_text_style(child(tx_styles, "otherStyle")),
    )


def parse_background(bg: Element | None) -> SourceBackground | None:
    """Read ``p:bg`` -- either an explicit ``p:bgPr`` fill or a ``p:bgRef`` theme index."""
    if bg is None:
        return None

    bg_pr = child(bg, "bgPr")
    if bg_pr is not None:
        fill = parse_fill(bg_pr)
        if fill is not None:
            return SourceBackground(fill=fill)

    bg_ref = child(bg, "bgRef")
    if bg_ref is not None:
        return SourceBackground(bg_ref=parse_style_reference(bg_ref))

    return None


def parse_color_map(clr_map: Element | None) -> SourceColorMap | None:
    if clr_map is None:
        return None
    mapping = {slot: attr(clr_map, slot) for slot in COLOR_MAP_SLOTS}
    resolved = {slot: value for slot, value in mapping.items() if value is not None}
    return SourceColorMap(mapping=resolved) if resolved else None


def parse_color_map_override(clr_map_ovr: Element | None) -> SourceColorMap | None:
    """``p:clrMapOvr`` either says "use the master's map" or supplies a full override."""
    if clr_map_ovr is None:
        return None
    return parse_color_map(child(clr_map_ovr, "overrideClrMapping"))


# --------------------------------------------------------------------------------------
# Theme
# --------------------------------------------------------------------------------------


def read_theme(package: OpcPackage, part_path: str) -> SourceTheme | None:
    root = package.read_xml(part_path)
    if root is None:
        return None
    elements = child(root, "themeElements")

    return SourceTheme(
        part_path=part_path,
        color_scheme=parse_color_scheme(child(elements, "clrScheme")),
        font_scheme=parse_font_scheme(child(elements, "fontScheme")),
        format_scheme=parse_format_scheme(child(elements, "fmtScheme")),
    )


def parse_color_scheme(clr_scheme: Element | None) -> dict:
    if clr_scheme is None:
        return {}
    scheme = {}
    for key in COLOR_SCHEME_KEYS:
        color = parse_color(child(clr_scheme, key))
        if color is not None:
            scheme[key] = color
    return scheme


def parse_font_scheme(font_scheme: Element | None) -> SourceFontScheme:
    if font_scheme is None:
        return SourceFontScheme()
    major = child(font_scheme, "majorFont")
    minor = child(font_scheme, "minorFont")
    return SourceFontScheme(
        major_latin=attr(child(major, "latin"), "typeface"),
        minor_latin=attr(child(minor, "latin"), "typeface"),
        major_east_asian=attr(child(major, "ea"), "typeface"),
        minor_east_asian=attr(child(minor, "ea"), "typeface"),
        major_complex_script=attr(child(major, "cs"), "typeface"),
        minor_complex_script=attr(child(minor, "cs"), "typeface"),
        major_japanese=_script_typeface(major, "Jpan"),
        minor_japanese=_script_typeface(minor, "Jpan"),
    )


def _script_typeface(font: Element | None, script: str) -> str | None:
    """``a:font script="Jpan"`` -- the per-script override list inside a font collection."""
    for node in children(font, "font"):
        if attr(node, "script") == script:
            return attr(node, "typeface")
    return None


def parse_format_scheme(fmt_scheme: Element | None) -> SourceFormatScheme:
    if fmt_scheme is None:
        return SourceFormatScheme()
    return SourceFormatScheme(
        fill_styles=[
            fill
            for fill in (parse_fill_style(node) for node in children(child(fmt_scheme, "fillStyleLst")))
            if fill is not None
        ],
        line_styles=[
            line
            for line in (parse_line(node) for node in children(child(fmt_scheme, "lnStyleLst"), "ln"))
            if line is not None
        ],
        effect_styles=[
            parse_effect_list(child(node, "effectLst"))
            for node in children(child(fmt_scheme, "effectStyleLst"), "effectStyle")
        ],
        bg_fill_styles=[
            fill
            for fill in (
                parse_fill_style(node) for node in children(child(fmt_scheme, "bgFillStyleLst"))
            )
            if fill is not None
        ],
    )


def parse_fill_style(node: Element | None):
    """A fill style list holds bare fill elements; wrap each so ``parse_fill`` can read it."""
    if node is None:
        return None
    wrapper = Element("wrapper")
    wrapper.append(node)
    return parse_fill(wrapper)

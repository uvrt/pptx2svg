"""A template view: a slide layout or master drawn on its own, without a slide.

What a slide built on the layout starts as -- the background, the master's shapes (unless
the layout's ``showMasterSp`` turns them off), the layout's own shapes -- with every
placeholder of the layout drawn as a **wireframe**: its box outlined in a grey dash, its
prompt text in the style a slide's text there inherits, and a small tag naming its type
and ``idx`` ("title", "content 1", "dt 10") on its top edge.  A master's view is the same with the master's
own placeholders.  A banner in the top right corner says which layout or master it is.

The wireframe is this library's own drawing, chosen so a designer can see where each
placeholder sits and how its text is styled; it is not how PowerPoint draws a layout.
PowerPoint shows no empty placeholder and no prompt on a slide outside editing (the
renderer's slide view leaves them out), and how Slide Master view draws its placeholder
outlines and prompts has not been measured.  What is drawn as on a slide -- background,
shapes, the placeholders' geometry and text styles -- goes through the slide pipeline.

A slide-number field shows its own text -- the ``‹#›`` a layout's usually holds -- since
the view is of no slide; a date field shows the date, as on a slide.
"""

from __future__ import annotations

from dataclasses import replace

from .. import model as m
from ..parse import source as s
from ..parse.parts import read_template
from .color import ColorContext, build_effective_color_map
from .naming import NamingColorContext
from .view import (
    ResolveContext,
    _resolve_background,
    _resolve_template_elements,
    resolve_element,
)

#: The wireframe's outline (EMU width, colour) and its tags' and banner's colours.
WIRE_WIDTH = 12700
WIRE_COLOR = "#7F7F7F"
TAG_FILL = "#404040"
TAG_TEXT = "#FFFFFF"
BANNER_FILL = "#1F1F1F"

#: A tag names a content placeholder (``p:ph`` with no type, or ``obj``) as describe does.
_TYPE_NAMES = {"obj": "content"}


def resolve_template(
    package,
    presentation: s.SourcePresentation,
    part_path: str,
    *,
    metafile_converter=None,
    east_asian=None,
    color_names: dict | None = None,
    now=None,
    locale=None,
    warnings: list | None = None,
) -> m.Slide:
    """The template view of the layout or master ``part_path`` (see the module); a
    :class:`~pptx2svg.model.Slide` with ``slide_number`` 0 and ``template`` set.  What
    resolution warns about is appended to ``warnings``, with ``part_path`` saying where."""
    stand_in = read_template(package, presentation, part_path)
    if stand_in.template == "layout":
        layout = presentation.layouts[stand_in.part_path]
        master = presentation.masters.get(layout.master_part_path or "")
        owner: s.SourceSlideBase = layout
    else:
        layout = None
        master = presentation.masters[stand_in.part_path]
        owner = master
    theme = presentation.themes.get(master.theme_part_path or "") if master else None
    color_map = build_effective_color_map(
        master.color_map if master else None,
        layout.color_map_override if layout else None,
        None,
    )
    context = ResolveContext(
        package=package,
        presentation=presentation,
        slide=stand_in,
        layout=layout,
        master=master,
        theme=theme,
        colors=(ColorContext(theme, color_map) if color_names is None
                else NamingColorContext(theme, color_map, color_names)),
        part_path=stand_in.part_path,
        metafile_converter=metafile_converter,
    )
    context.east_asian = east_asian
    if now is not None:
        context.now = now
    if locale is not None:
        context.locale = locale

    elements: list[m.SlideElement] = []
    show_master = layout.show_master_shapes if layout is not None else True
    if master is not None and show_master:
        context.id_prefix = "mst:"
        elements.extend(_resolve_template_elements(context, master.shapes, master.part_path))
    if layout is not None:
        context.id_prefix = "lay:"
        elements.extend(_resolve_template_elements(context, layout.shapes, layout.part_path))

    context.part_path = owner.part_path
    context.id_prefix = "lay:" if layout is not None else "mst:"
    size = _tag_size(presentation)
    tags: list[m.SlideElement] = []
    for index, shape in enumerate(owner.shapes):
        if getattr(shape, "placeholder", None) is None:
            continue
        element = resolve_element(context, shape, (index,))
        if element is None:
            continue
        if isinstance(element, m.ShapeElement):
            element = replace(element, outline=m.Outline(
                width=WIRE_WIDTH, fill=m.SolidFill(m.ResolvedColor(WIRE_COLOR)),
                dash_style="dash"))
            # The tag sits on the box's top edge, above it where the slide has room, so
            # it covers none of the prompt.
            top = element.transform.offset_y - _tag_height(size)
            tags.append(_tag(_label(element), element.transform.offset_x,
                             top if top >= 0 else element.transform.offset_y, size,
                             TAG_FILL, theme))
        elements.append(element)
    elements.extend(tags)

    kind = "Layout" if layout is not None else "Master"
    name = owner.name or (_theme_name(package, master) if layout is None else None) or ""
    banner = f"{kind} “{name}” · placeholders outlined" if name else \
        f"{kind} · placeholders outlined"
    width = _text_width(banner, size)
    elements.append(_tag(banner, presentation.slide_width - width, 0, size, BANNER_FILL, theme,
                         width=width))
    if warnings is not None:
        warnings.extend(context.warnings)
    return m.Slide(
        slide_number=0,
        background=_resolve_background(context),
        elements=elements,
        show_master_sp=show_master,
        template=stand_in.template,
        template_part=stand_in.part_path,
        template_name=owner.name,
    )


def _label(element: m.ShapeElement) -> str:
    kind = element.placeholder_type or "obj"
    kind = _TYPE_NAMES.get(kind, kind)
    index = element.placeholder_idx
    return f"{kind} {index}" if index else kind


def _tag_size(presentation: s.SourcePresentation) -> float:
    """The tags' text size in points: a hundredth of the slide's width, 8 pt at least --
    10 pt on a 16:9 slide."""
    return max(8.0, round(presentation.slide_width / 12700 / 96, 1))


def _text_width(text: str, size: float) -> float:
    """A tag's width in EMU: its text at an estimated 0.6 em a character, plus padding.
    An estimate is enough -- the text is drawn unwrapped and left-aligned in it."""
    return (len(text) * 0.6 * size + size) * 12700


def _tag_height(size: float) -> float:
    return size * 1.6 * 12700


def _tag(text: str, x: float, y: float, size: float, fill: str, theme,
         *, width: float | None = None) -> m.ShapeElement:
    """A small filled box with ``text`` in it, at ``(x, y)`` (EMU)."""
    font = (theme.font_scheme.minor_latin if theme is not None else None) or "Calibri"
    run = m.TextRun(text=text, properties=m.RunProperties(
        font_size=size, font_family=font, color=m.ResolvedColor(TAG_TEXT)))
    pad = size * 0.5 * 12700
    return m.ShapeElement(
        transform=m.Transform(offset_x=x, offset_y=y,
                              extent_width=width if width is not None
                              else _text_width(text, size),
                              extent_height=_tag_height(size)),
        geometry=m.PresetGeometry(preset="rect"),
        fill=m.SolidFill(m.ResolvedColor(fill, alpha=0.85)),
        text_body=m.TextBody(
            paragraphs=[m.Paragraph(runs=[run],
                                    properties=m.ParagraphProperties(alignment="l"))],
            body_properties=m.BodyProperties(anchor="ctr", margin_left=pad, margin_right=pad,
                                             margin_top=0, margin_bottom=0, wrap="none")),
        alt_text="template view tag",
    )


def _theme_name(package, master: s.SourceSlideMaster | None) -> str | None:
    """A master with no name of its own is shown by its theme's (as pptx-agent names it)."""
    if master is None or master.theme_part_path is None:
        return None
    root = package.read_xml(master.theme_part_path)
    return root.get("name") if root is not None else None

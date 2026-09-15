"""Shared state for one SVG render pass.

Three things need to be threaded through every renderer:

* **The ``<defs>`` accumulator.** Gradients, patterns, arrow markers and filters must be
  declared once at the top of the document and referenced by id from wherever they are
  used, so renderers push definitions here and return only the referencing attributes.
* **An id counter.** pptx-glimpse mints these with ``crypto.randomUUID()``; a per-render
  counter is used instead so the same deck always produces byte-identical SVG, which
  makes the output diffable and snapshot-testable.
* **The enclosing groups' coordinate scale.** Geometry inside a group is scaled by the
  group's ``ext``/``chExt`` ratio and text is *not*; see :attr:`RenderContext.group_scale`.
  This carries the *uniform* case only -- a non-uniform ratio is folded into the
  children's own boxes instead, for the reason in
  :func:`~pptx2svg.render.svg.swaps_group_axes`.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..text.fontmap import DEFAULT_FONT_MAPPING
from ..text.measure import DefaultTextMeasurer, TextMeasurer


@dataclass
class RenderContext:
    measurer: TextMeasurer = field(default_factory=DefaultTextMeasurer)
    font_mapping: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_FONT_MAPPING))
    #: Last-resort typeface for East Asian text, from the theme's ``Jpan`` script font.
    #:
    #: A backstop rather than the rule: the cascade that picks an East Asian face now runs
    #: at resolution time (``pptx2svg.resolve.text._theme_east_asian``) and reaches the
    #: measurer as ``RunProperties.font_family_ea``, so measurement and drawing cannot
    #: disagree about it.  This stays for a caller who builds a ``RenderContext`` by hand
    #: and for a model that predates that.
    jpan_fallback_font: str | None = None
    #: Emit text as ``<text>``/``<tspan>``.  (Text-to-path would need embedded fonts.)
    defs: list[str] = field(default_factory=list)
    #: The product of every enclosing group's ``ext``/``chExt`` ratio, as ``(x, y)``.
    #:
    #: A group maps its children's authored coordinate space onto its on-slide box, and
    #: **PowerPoint applies that mapping to geometry only.  Text keeps the point size it
    #: was authored at.**  So the text frame is the shape's *scaled* rectangle while the
    #: type inside it is absolute, and a renderer that draws text inside the group's
    #: ``<g transform="scale(...)">`` gets the font size wrong by exactly this factor.
    #:
    #: Measured, not assumed.  A probe deck of 21 groups was exported through PowerPoint
    #: 16.x and the drawn text read back out of the PDF with ``pypdfium2``: every 18 pt
    #: Arial run inside a group drew with the same 12.89 pt cap height as the ungrouped
    #: control, at scale 4.0 (uniform), 6.0 (nested 2x inside 3x), 0.25 (a group that
    #: shrinks), and at 4.0 on one axis only.  The non-uniform probe settles the tempting
    #: alternative reading -- that text is scaled and the template is simply mis-authored
    #: -- since a stretched run would have drawn four times as wide as the control, and
    #: drew 58.67 pt against the control's 58.67 pt.
    #:
    #: Everything else in the text frame is absolute too, each measured on that deck:
    #: ``bodyPr@lIns`` of 91440 EMU drew a 7.20 pt indent at scale 1 and 7.20 pt at scale
    #: 4; ``lnSpc`` ``spcPts`` 3000 gave a 30.00 pt baseline pitch at both; ``spcBef``
    #: ``spcPts`` 1200 gave a 32.88 pt paragraph pitch at both; ``normAutofit``
    #: ``fontScale`` 50% gave a 6.44 pt cap height at both; and the first baseline sat
    #: 18.1 pt below the frame top at both.  Only the frame itself scales -- a long
    #: paragraph broke into identical lines inside a 4x group and in an ungrouped box of
    #: the same on-slide width, so wrapping happens at the scaled width and the authored
    #: size.
    #:
    #: Kept per-axis because the non-uniform probe proves the two are independent -- but
    #: in practice the two entries are now always equal.  A *non-uniform* group no longer
    #: emits an SVG ``scale()`` at all: one around a rotated child would compose to a
    #: shear, and PowerPoint draws a rotated rectangle instead, so the ratio is folded
    #: into each child's own box by :func:`~pptx2svg.render.svg.render_group`.  A folded
    #: child's frame is already in the enclosing space, so it needs no counter-transform
    #: and this attribute is left alone for it.  Only a uniform scale -- which commutes
    #: with rotation and is therefore exact as a ``scale()`` -- still lands here.
    group_scale: tuple[float, float] = (1.0, 1.0)
    _next_id: int = 0

    def new_id(self, prefix: str) -> str:
        self._next_id += 1
        return f"{prefix}-{self._next_id}"

    def add_def(self, definition: str) -> None:
        if definition:
            self.defs.append(definition)


def escape_xml_text(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def escape_xml_attr(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace('"', "&quot;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def num(value: float) -> str:
    """Compact number formatting for SVG attribute values."""
    rounded = round(value, 3)
    if rounded == int(rounded):
        return str(int(rounded))
    return f"{rounded:g}"

"""Shared state for one SVG render pass.

Two things need to be threaded through every renderer:

* **The ``<defs>`` accumulator.** Gradients, patterns, arrow markers and filters must be
  declared once at the top of the document and referenced by id from wherever they are
  used, so renderers push definitions here and return only the referencing attributes.
* **An id counter.** pptx-glimpse mints these with ``crypto.randomUUID()``; a per-render
  counter is used instead so the same deck always produces byte-identical SVG, which
  makes the output diffable and snapshot-testable.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..text.fontmap import DEFAULT_FONT_MAPPING
from ..text.measure import DefaultTextMeasurer, TextMeasurer


@dataclass
class RenderContext:
    measurer: TextMeasurer = field(default_factory=DefaultTextMeasurer)
    font_mapping: dict[str, str] = field(default_factory=lambda: dict(DEFAULT_FONT_MAPPING))
    #: Final fallback typeface for CJK text, from the theme's ``Jpan`` script font.
    jpan_fallback_font: str | None = None
    #: Emit text as ``<text>``/``<tspan>``.  (Text-to-path would need embedded fonts.)
    defs: list[str] = field(default_factory=list)
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

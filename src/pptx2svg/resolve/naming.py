"""Theme-colour names for resolved colours: what the agent view writes beside each hex.

Resolution turns ``<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>``
into ``#0B5259`` and forgets where it came from, which is right for drawing.  A reader that
edits the deck wants the name back (``accent1 lumMod=75%``), so a conversion that asks for
names resolves through a :class:`NamingColorContext`, and every scheme colour resolved
through it is remembered against the resolved object.  A plain :class:`ColorContext` has no
``names`` and resolves exactly as before: the normal render is untouched.
"""

from __future__ import annotations

from typing import Any

from .color import POWERPOINT, ColorContext
from .color import resolve_color as _resolve_color

#: Transforms written as a percentage (OOXML stores 1/1000 percent).
_PERCENT = {"lumMod", "lumOff", "tint", "shade", "alpha", "alphaMod", "alphaOff", "satMod",
            "satOff", "sat", "lum", "hueMod"}


class NamingColorContext(ColorContext):
    """A :class:`ColorContext` that remembers the theme name of each colour it resolves."""

    __slots__ = ("names",)

    def __init__(self, theme, color_map, names: dict[int, tuple[Any, str]], rules=POWERPOINT):
        super().__init__(theme, color_map, rules)
        #: ``id(resolved) -> (resolved, name)``; the object is kept so its id stays unique.
        self.names = names


def describe(scheme: str, transforms) -> str:
    """``accent1``, ``accent1 lumMod=75%``, ``tx1 alpha=50%``."""
    parts = [scheme]
    for transform in transforms or ():
        kind, value = getattr(transform, "kind", None), getattr(transform, "value", None)
        if kind is None:
            continue
        if value is None:
            parts.append(kind)
        elif kind in _PERCENT:
            parts.append(f"{kind}={value / 1000:g}%")
        else:
            parts.append(f"{kind}={value:g}")
    return " ".join(parts)


def resolve_color(context, color, visited: frozenset = frozenset()):
    """:func:`ooxml_common.drawingml.color.resolve_color`, noting the name when asked to."""
    resolved = _resolve_color(context, color, visited)
    names = getattr(context, "names", None)
    if names is not None and resolved is not None and getattr(color, "kind", None) == "scheme" \
            and color.scheme != "phClr":
        names[id(resolved)] = (resolved, describe(color.scheme, color.transforms))
    return resolved


def note_derived(context, resolved, base, transforms) -> None:
    """``resolved`` is ``base`` (a resolved colour) with ``transforms``: name it after
    ``base`` when ``base`` has a name (a style reference's ``phClr``)."""
    names = getattr(context, "names", None)
    if names is None or resolved is None or base is None:
        return
    known = names.get(id(base))
    if known is not None and known[0] is base:
        extra = describe("", transforms).strip()
        names[id(resolved)] = (resolved, f"{known[1]} {extra}".strip())


def name_of(names: dict[int, tuple[Any, str]] | None, resolved) -> str | None:
    """The theme name a resolved colour was given, if any."""
    if not names or resolved is None:
        return None
    known = names.get(id(resolved))
    return known[1] if known is not None and known[0] is resolved else None

"""Namespace-agnostic helpers over ``xml.etree.ElementTree``.

OOXML mixes a dozen namespaces (``a:``, ``p:``, ``r:``, ``dgm:`` ...) and the same
local name means the same thing in each.  pptx-glimpse matches on local names
throughout; these helpers do the same so callers never spell out a namespace URI.

Element order is significant in DrawingML (a paragraph's runs and breaks interleave),
and ElementTree preserves document order, so ``children()`` is the ordered view that
the TypeScript port had to reconstruct by hand from fast-xml-parser output.
"""

from __future__ import annotations

import re
from typing import Iterable, Iterator
from xml.etree.ElementTree import Element, fromstring

# Relationship namespaces we need to spell out, because the attribute name alone
# (``id``, ``embed``, ``link``) is ambiguous without them.
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CHAR_REF = re.compile(r"&#x([0-9a-fA-F]+);|&#([0-9]+);")


def parse_xml(data: bytes | str) -> Element:
    if isinstance(data, str):
        data = data.encode("utf-8")
    return fromstring(data)


def local_name(tag: str) -> str:
    """``{http://...}sp`` -> ``sp``; ``a:sp`` -> ``sp``."""
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    if ":" in tag:
        return tag.rsplit(":", 1)[-1]
    return tag


def children(parent: Element | None, name: str | None = None) -> list[Element]:
    """Direct children in document order, optionally filtered by local name."""
    if parent is None:
        return []
    if name is None:
        return list(parent)
    return [child for child in parent if local_name(child.tag) == name]


def child(parent: Element | None, name: str) -> Element | None:
    """First direct child with the given local name."""
    if parent is None:
        return None
    for node in parent:
        if local_name(node.tag) == name:
            return node
    return None


def child_path(parent: Element | None, *names: str) -> Element | None:
    """Walk a chain of local names, e.g. ``child_path(sp, "spPr", "xfrm", "off")``."""
    node = parent
    for name in names:
        node = child(node, name)
        if node is None:
            return None
    return node


def has_child(parent: Element | None, name: str) -> bool:
    return child(parent, name) is not None


def descendants(parent: Element | None, name: str) -> Iterator[Element]:
    if parent is None:
        return
    for node in parent.iter():
        if node is not parent and local_name(node.tag) == name:
            yield node


def find_descendant(parent: Element | None, name: str) -> Element | None:
    return next(descendants(parent, name), None)


def attr(node: Element | None, name: str, default: str | None = None) -> str | None:
    """Attribute lookup that ignores namespaces on the attribute name."""
    if node is None:
        return default
    value = node.get(name)
    if value is not None:
        return value
    for key, val in node.attrib.items():
        if local_name(key) == name:
            return val
    return default


def ns_attr(node: Element | None, name: str, namespace: str = NS_R) -> str | None:
    """Namespace-qualified attribute (``r:embed``, ``r:id``)."""
    if node is None:
        return None
    return node.get(f"{{{namespace}}}{name}")


def num_attr(node: Element | None, name: str) -> float | None:
    """Numeric attribute; ``None`` when absent or not a finite number."""
    raw = attr(node, name)
    if raw is None:
        return None
    try:
        value = float(raw)
    except (TypeError, ValueError):
        return None
    if value != value or value in (float("inf"), float("-inf")):
        return None
    return value


def int_attr(node: Element | None, name: str) -> int | None:
    value = num_attr(node, name)
    return None if value is None else int(value)


def is_true(value: str | None) -> bool:
    """OOXML boolean attribute (``1`` / ``0`` / ``true`` / ``false``)."""
    return value in ("1", "true")


def bool_attr(node: Element | None, name: str) -> bool:
    return is_true(attr(node, name))


def child_text(parent: Element | None, name: str) -> str | None:
    node = child(parent, name)
    if node is None:
        return None
    return "".join(node.itertext())


def decode_char_refs(value: str) -> str:
    """Decode numeric character references left in attribute values (bullet chars)."""

    def replace(match: re.Match[str]) -> str:
        hex_digits, decimal = match.group(1), match.group(2)
        if hex_digits is not None:
            return chr(int(hex_digits, 16))
        return chr(int(decimal or "0", 10))

    return _CHAR_REF.sub(replace, value)


def enum_value(value: str | None, allowed: Iterable[str], default: str | None = None) -> str | None:
    return value if value is not None and value in set(allowed) else default

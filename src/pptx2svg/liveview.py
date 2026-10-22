"""What a live, editable view of a slide needs from its SVG (round 13).

A page that shows each slide as SVG, re-renders a slide after an edit and sends only the
shapes that changed needs four things the plain render does not give it:

* **pictures by reference** (``ConvertOptions(images={})``): each picture the SVG would
  carry as a ``data:`` URI written once, under a name made of its content's hash and its
  format (``img/3fa1c9….png``), and referenced by ``href``; the bytes go to the caller's
  dict.  :func:`externalise_images` does it, :func:`embed_images` undoes it -- what
  :func:`pptx2svg.rasterise_svgs` does with its ``images`` before drawing, so the PNG is
  the embedded form's, byte for byte.
* **edit marks** (``ConvertOptions(edit_marks=True)``): ``data-pptx-source`` on every
  ``data-pptx-id``, ``data-pptx-size`` on the root, table cells and text bodies marked
  (:attr:`ooxml_common.drawingml.context.RenderContext.edit_marks`), and every shape's
  definitions under ids of its own, so a shape that did not change draws the same bytes.
* **the slide by shape** (:func:`slide_groups`): each top-level shape's markup by its id,
  in drawing order, with the definitions it uses.
* **no hyperlinks** (``ConvertOptions(hyperlinks=False)``).

Standard library only, string work on the renderer's own output: its attribute values
escape ``<`` and ``>``, and it writes no comments, CDATA or processing instructions.
"""

from __future__ import annotations

import base64
import hashlib
import re
from typing import Mapping, MutableMapping

__all__ = ["embed_images", "externalise_images", "mark_sources", "slide_groups"]

#: A picture as the renderer embeds it: ``href`` or ``xlink:href``, a media type, base64.
_DATA_HREF = re.compile(r'((?:xlink:)?href)="data:([^";,]+);base64,([A-Za-z0-9+/=]*)"')

#: The extension a picture's name takes, by what its bytes say it is.
_SNIFF = (
    (b"\x89PNG\r\n\x1a\n", "png"),
    (b"\xff\xd8\xff", "jpg"),
    (b"GIF87a", "gif"),
    (b"GIF89a", "gif"),
    (b"BM", "bmp"),
    (b"II*\x00", "tif"),
    (b"MM\x00*", "tif"),
)

#: The media type a name's extension is embedded under again.
_MIME = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "webp": "image/webp",
    "svg": "image/svg+xml",
}
_EXTENSION = {mime: extension for extension, mime in _MIME.items()}


def _format(data: bytes, mime: str) -> str:
    """The extension of a picture: what its bytes are, else what its media type says."""
    for magic, extension in _SNIFF:
        if data.startswith(magic):
            return extension
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    head = data[:512].lstrip()
    if head.startswith(b"<svg") or (head.startswith(b"<?xml") and b"<svg" in data[:4096]):
        return "svg"
    if mime in _EXTENSION:
        return _EXTENSION[mime]
    # A type with no common extension keeps its subtype, so it is embedded under it again.
    subtype = mime.partition("/")[2] or "bin"
    return re.sub(r"[^A-Za-z0-9.+-]", "_", subtype)


def externalise_images(svg: str, store: MutableMapping[str, bytes], prefix: str = "img/") -> str:
    """``svg`` with every embedded picture (``href="data:…;base64,…"``) referenced by name
    instead, its bytes put in ``store`` under that name.

    The name is ``prefix`` + the first 32 hex digits of the SHA-256 of the bytes + ``.`` +
    the format the bytes are (PNG, JPEG, GIF, BMP, TIFF, WebP, SVG), else the media type's
    subtype: the same picture -- on another slide, or twice on one -- is one entry.  Every
    picture is written this way: a picture, a picture fill, a slide background, a picture
    bullet, a metafile's drawing or preview (an SVG or PNG).  What a referenced SVG itself
    embeds stays inside it.
    """
    if "base64," not in svg:
        return svg

    def replace(match: re.Match) -> str:
        attribute, mime, encoded = match.groups()
        data = base64.b64decode(encoded)
        name = f"{prefix}{hashlib.sha256(data).hexdigest()[:32]}.{_format(data, mime)}"
        store.setdefault(name, data)
        return f'{attribute}="{name}"'

    return _DATA_HREF.sub(replace, svg)


def embed_images(svg: str, images: Mapping[str, bytes]) -> str:
    """``svg`` with every picture :func:`externalise_images` named embedded again, as a
    ``data:`` URI of the bytes ``images`` holds under its name.  A name ``images`` lacks
    raises ``KeyError``, naming it: drawn, the picture would be left out without a word."""
    def replace(match: re.Match) -> str:
        attribute, name = match.groups()
        data = images.get(name)
        if data is None:
            return match.group(0)
        extension = name.rpartition(".")[2]
        mime = _MIME.get(extension, f"image/{extension}")
        return f'{attribute}="data:{mime};base64,{base64.b64encode(data).decode("ascii")}"'

    embedded = _NAMED_HREF.sub(replace, svg) if images else svg
    left = _IMAGE_BY_NAME.search(embedded)
    if left is not None:
        raise KeyError(f"the SVG references the picture {left.group(1)!r} and `images` has no "
                       "such entry; pass the images the render stored (ConvertOptions.images)")
    return embedded


_NAMED_HREF = re.compile(r'((?:xlink:)?href)="(?!data:|#)([^"]*)"')
_IMAGE_BY_NAME = re.compile(r'<image\b[^>]*?\bhref="(?!data:|#)([^"]+)"')
_ID = re.compile(r'data-pptx-id="([^"]*)"')


def mark_sources(svg: str) -> str:
    """``data-pptx-source`` after every ``data-pptx-id``: ``layout`` for a layout's shape
    (``lay:``), ``master`` for a master's (``mst:``), ``slide`` for the rest."""

    def replace(match: re.Match) -> str:
        value = match.group(1)
        source = "layout" if value.startswith("lay:") else "master" if value.startswith("mst:") else "slide"
        return f'{match.group(0)} data-pptx-source="{source}"'

    return _ID.sub(replace, svg)


def id_scope(element_id: str | None, path: str | None, used: set[str]) -> str:
    """The id scope (:meth:`~ooxml_common.drawingml.context.RenderContext.scope_ids`) of a
    top-level shape: ``s`` + its id, any character an id may not hold made ``_``, else
    ``p`` + its path; a scope already in ``used`` gets ``_2``, ``_3``... (``cNvPr@id`` is
    not unique in every deck)."""
    base = f"s{element_id}" if element_id else f"p{path or 'x'}"
    base = re.sub(r"[^A-Za-z0-9_]", "_", base)
    scope, count = base, 1
    while scope in used:
        count += 1
        scope = f"{base}_{count}"
    used.add(scope)
    return f"{scope}-"


_TAG = re.compile(r"<(/?)([A-Za-z][\w:.-]*)([^>]*)>")
_REFERENCE = re.compile(r'url\(#([^)]+)\)|href="#([^"]+)"')
_ATTRIBUTE_ID = re.compile(r'\bid="([^"]+)"')


def _children(markup: str, start: int, end: int) -> list[tuple[int, int, str]]:
    """The top-level elements of ``markup[start:end]``, as ``(start, end, tag name)``."""
    out: list[tuple[int, int, str]] = []
    depth = 0
    opened = 0
    name = ""
    for match in _TAG.finditer(markup, start, end):
        closing, tag, rest = match.groups()
        empty = rest.endswith("/")
        if closing:
            depth -= 1
            if depth == 0:
                out.append((opened, match.end(), name))
        elif empty:
            if depth == 0:
                out.append((match.start(), match.end(), tag))
        else:
            if depth == 0:
                opened, name = match.start(), tag
            depth += 1
    return out


def slide_groups(svg: str) -> dict:
    """One slide's SVG by shape: ``{"root", "order", "groups", "defs"}``.

    ``groups`` maps each top-level shape's key to its markup, exactly as it stands in
    ``svg`` (a hyperlink's ``<a>`` around it included); ``order`` is the keys in drawing
    order, the slide's background first, as ``"background"``.  A shape's key is its
    ``data-pptx-id``; an id drawn again (``cNvPr@id`` is not unique in every deck;
    pptx-agent's ids are) takes ``#2``, ``#3``...; a top-level element with no id, ``#``
    and its position.  ``defs`` maps each key to the definitions its markup uses
    (gradients, patterns, clip paths, masks, filters, markers -- and those they use), in
    document order, as one string, ``""`` for none.  ``root`` is the ``<svg …>`` tag.

    The definitions stay in the document's one ``<defs>``: a page that swaps a group
    swaps its definitions there too.  Rendered with ``ConvertOptions(edit_marks=True)``
    each shape's definitions carry ids of their own (``grad-s256_4-1``), so a shape that
    did not change has byte-identical markup and definitions from render to render, and
    no two shapes share one; without it they are numbered across the slide, and every
    shape after a changed one may differ.  Works on any SVG this library draws, and on
    pptx-agent's (``Slide.render_svg``), whose ids are its own.
    """
    root = re.search(r"<svg\b[^>]*>", svg)
    if root is None:
        raise ValueError("not an SVG document: no <svg> root")
    end = svg.rindex("</svg>")
    children = _children(svg, root.end(), end)

    definitions: dict[str, str] = {}
    def_order: list[str] = []
    groups: dict[str, str] = {}
    order: list[str] = []
    for start, stop, tag in children:
        markup = svg[start:stop]
        if tag == "defs":
            open_end = svg.index(">", start) + 1
            for inner_start, inner_end, _ in _children(svg, open_end, stop - len("</defs>")):
                definition = svg[inner_start:inner_end]
                found = _ATTRIBUTE_ID.search(definition)
                if found is not None:
                    definitions[found.group(1)] = definition
                    def_order.append(found.group(1))
            continue
        if not order:
            key = "background"
        else:
            found = _ID.search(markup[:markup.index(">") + 1]) or (
                _ID.search(markup) if markup.startswith("<a ") else None)
            key = found.group(1) if found is not None else f"#{len(order)}"
            if key in groups:
                count = 2
                while f"{key}#{count}" in groups:
                    count += 1
                key = f"{key}#{count}"
        groups[key] = markup
        order.append(key)

    position = {name: index for index, name in enumerate(def_order)}
    used_by: dict[str, str] = {}
    for key in order:
        seen: set[str] = set()
        pending = [name for pair in _REFERENCE.findall(groups[key]) for name in pair if name]
        while pending:
            name = pending.pop()
            if name in seen or name not in definitions:
                continue
            seen.add(name)
            pending.extend(n for pair in _REFERENCE.findall(definitions[name]) for n in pair if n)
        used_by[key] = "".join(definitions[name] for name in sorted(seen, key=position.__getitem__))
    return {"root": root.group(0), "order": order, "groups": groups, "defs": used_by}

"""Text a rasteriser will not draw: no font it loads answers to the text or has its glyphs.

resvg draws a run of text with the first family in its ``font-family`` stack that a loaded
face answers to (a generic keyword through the generic family it is pointed at), and a
character that face lacks with any other loaded face that has it.  Two things go wrong
without a word (resvg-py 0.5, ``tests/test_glyphs.py``):

* **no family in the stack is loaded** -- the run is not drawn at all, Latin and all;
* **one is, but no loaded face has a character** -- the character is drawn as the face's
  empty box, or not at all.

Without the ``pptx2svg-fonts`` bundle, on a host with no CJK face (a Linux server, a slim
container), every Japanese, Chinese and Korean character of a deck went one of those two
ways, and the PNG looked finished.  So before an SVG is rasterised, each run's stack is
resolved against the family names of the faces the rasteriser will load -- the caller's
font files and directories, the bundle's, and this machine's own fonts unless they are
skipped -- and its characters checked against their character maps.  What will not be
drawn is reported **once per face and script** (``Noto Sans JP`` / CJK, ``Noto Sans JP`` /
Hangul), never once per glyph: :func:`missing_glyphs`.  :func:`pptx2svg.svg_to_png` raises
a :class:`MissingGlyphsWarning` for each, and :func:`pptx2svg.convert_pptx_to_png` also puts
a ``glyphs-missing`` warning in ``ConvertOptions.warnings``.

Only the fonts' ``name`` and ``cmap`` tables are read, where the files are installed, and
only until every run is resolved and every character found: a deck the bundle covers never
reads a system font.  What a file holds is remembered for the life of the process.
"""

from __future__ import annotations

import functools
import os
import struct
import sys
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Mapping, Sequence

__all__ = ["MissingGlyphs", "MissingGlyphsWarning", "missing_glyphs", "script_of",
           "system_font_dirs", "text_by_stack"]

_FONT_SUFFIXES = {".ttf", ".otf", ".ttc", ".otc"}
#: How many of a face's missing characters a report quotes.
_SAMPLE = 8
#: What resvg's font database calls the generic families when nobody points them elsewhere.
RESVG_GENERIC_FAMILIES = {
    "serif": "Times New Roman", "sans-serif": "Arial", "cursive": "Comic Sans MS",
    "fantasy": "Impact", "monospace": "Courier New",
}
_GENERIC_OPTIONS = {"serif": "serif_family", "sans-serif": "sans_serif_family",
                    "cursive": "cursive_family", "fantasy": "fantasy_family",
                    "monospace": "monospace_family"}


class MissingGlyphsWarning(UserWarning):
    """Some text will not be drawn: no font the rasteriser loads can draw it."""


@dataclass(frozen=True)
class MissingGlyphs:
    """Characters of one script, asked for in one face, that the rasteriser will not draw.

    ``face`` is the first family of the text's ``font-family`` (the face the deck asks
    for), ``script`` its script (:func:`script_of`: ``"CJK"``, ``"Hangul"``, ``"Latin"``...)
    and ``sample`` up to eight of the characters, in the order the text has them.
    ``unloaded`` says why: ``True`` when no family of the stack (``stack``) is loaded, so
    the whole run is dropped; ``False`` when one is but no loaded face has the glyphs.
    """

    face: str
    script: str
    sample: str
    unloaded: bool = False
    stack: str = ""

    def message(self) -> str:
        if self.unloaded:
            why = (f"no font this render loads answers to its font-family ({self.stack}), "
                   f"so its {self.script} text ({self.sample}) is not drawn")
        else:
            why = (f"no font this render loads has its {self.script} glyphs ({self.sample}), "
                   "so that text is not drawn")
        return (f"{self.face}: {why}. Install pptx2svg-fonts "
                "(pip install 'pptx2svg[fonts]') or pass font_dirs= a face that has them")


def text_by_stack(svg: str) -> dict[str, str]:
    """Every character an SVG's text draws, by its ``font-family`` stack, each once, in
    order: ``{"'Noto Sans JP', Calibri, sans-serif": "サンプル拡張..."}``.  Whitespace and
    format characters, which need no glyph, are left out."""
    try:
        root = ET.fromstring(svg)
    except ET.ParseError:
        return {}
    out: dict[str, dict[str, None]] = {}

    def add(stack: str | None, text: str | None) -> None:
        if not text:
            return
        seen = out.setdefault(stack or "serif", {})
        for char in text:
            if not char.isspace() and unicodedata.category(char) not in ("Cf", "Cc"):
                seen.setdefault(char, None)

    def walk(element: ET.Element, stack: str | None, in_text: bool) -> None:
        stack = element.get("font-family") or stack
        tag = element.tag.rpartition("}")[2]
        texty = tag in ("text", "tspan", "textPath") or (in_text and tag == "a")
        if texty:
            add(stack, element.text)
        for child in element:
            walk(child, stack, texty)
            if texty:
                add(stack, child.tail)

    walk(root, None, False)
    return {stack: "".join(chars) for stack, chars in out.items() if chars}


def _families(stack: str) -> list[str]:
    names = (name.strip().strip("'\"").strip() for name in stack.split(","))
    return [name for name in names if name]


def missing_glyphs(
    svg: "str | Mapping[str, str]",
    *,
    font_files: Sequence[str] = (),
    font_dirs: Sequence[str] = (),
    system_fonts: bool = False,
    generic_families: Mapping[str, str] | None = None,
) -> list[MissingGlyphs]:
    """What of ``svg``'s text (or of a ``{font-family stack: characters}`` mapping, as
    :func:`text_by_stack` gives) a rasteriser loading ``font_files``, ``font_dirs`` and,
    with ``system_fonts``, this machine's font folders (:func:`system_font_dirs`) will not
    draw: one :class:`MissingGlyphs` per face and script.  ``generic_families`` are the
    families the generic keywords are pointed at, as :func:`pptx2svg.svg_to_png` passes
    them to resvg (``sans_serif_family``...); resvg's own defaults where ``None``."""
    texts = text_by_stack(svg) if isinstance(svg, str) else dict(svg)
    if not texts:
        return []
    generic = dict(RESVG_GENERIC_FAMILIES)
    for keyword, option in _GENERIC_OPTIONS.items():
        if generic_families and generic_families.get(option):
            generic[keyword] = generic_families[option]
    wanted = {stack: [generic.get(name.lower(), name).lower() for name in _families(stack)]
              for stack in texts}
    unresolved = set(texts)
    uncovered = {ord(char) for chars in texts.values() for char in chars}
    for path in _font_paths(font_files, font_dirs, system_fonts):
        if uncovered:
            uncovered -= _font_info(path, _coverage_of)
        if unresolved:
            families = _font_info(path, _families_of)
            unresolved = {stack for stack in unresolved
                          if not any(name in families for name in wanted[stack])}
        if not uncovered and not unresolved:
            return []
    out: list[MissingGlyphs] = []
    reported: set[tuple[str, str]] = set()
    for stack, chars in texts.items():
        names = _families(stack)
        face = names[0] if names else stack
        lost = stack in unresolved
        by_script: dict[str, list[str]] = {}
        for char in chars:
            if lost or ord(char) in uncovered:
                by_script.setdefault(script_of(char), []).append(char)
        for script, missing in by_script.items():
            if (face, script) not in reported:
                reported.add((face, script))
                out.append(MissingGlyphs(face, script, "".join(missing[:_SAMPLE]), lost,
                                         ", ".join(names)))
    return out


def script_of(char: str) -> str:
    """The script a character is written in, by its Unicode name: ``"CJK"`` for Chinese
    and Japanese -- ideographs, kana, bopomofo, their punctuation and full-width forms,
    which one face draws together -- ``"Hangul"`` for Korean, ``"Latin"``, ``"Arabic"``...;
    digits and punctuation with no script of their own count as Latin, a symbol or emoji
    as ``"Symbols"``, what has no name as ``"Other"``."""
    name = unicodedata.name(char, "")
    if not name:
        return "Other"
    if "HANGUL" in name:
        return "Hangul"
    word = name.split()[0].split("-")[0]
    if word in _CJK or "IDEOGRAPH" in name or "CJK" in name:
        return "CJK"
    if word in _SCRIPTS:
        return word.title()
    category = unicodedata.category(char)
    if category.startswith("S"):
        return "Symbols"
    if category.startswith(("N", "P")):
        return "Latin"
    return word.title()


_CJK = {"CJK", "KANGXI", "IDEOGRAPHIC", "HIRAGANA", "KATAKANA", "BOPOMOFO", "FULLWIDTH",
        "HALFWIDTH", "KANBUN"}
_SCRIPTS = {
    "LATIN", "GREEK", "CYRILLIC", "ARMENIAN", "HEBREW", "ARABIC", "SYRIAC", "THAANA",
    "DEVANAGARI", "BENGALI", "GURMUKHI", "GUJARATI", "ORIYA", "TAMIL", "TELUGU", "KANNADA",
    "MALAYALAM", "SINHALA", "THAI", "LAO", "TIBETAN", "MYANMAR", "GEORGIAN", "ETHIOPIC",
    "CHEROKEE", "KHMER", "MONGOLIAN", "YI",
}


def system_font_dirs() -> list[str]:
    """The folders resvg loads this machine's fonts from (its font database's)."""
    if sys.platform == "darwin":
        from .fonts import office

        return office.resvg_system_dirs()
    from ooxml_common.fonts import office as shared

    return [str(path) for path in shared.system_font_dirs()]


def _font_paths(font_files: Sequence[str], font_dirs: Sequence[str],
                system_fonts: bool) -> Iterable[str]:
    """Every font file the rasteriser would load, the caller's own first, lazily."""
    seen: set[str] = set()
    for path in font_files:
        if str(path) not in seen:
            seen.add(str(path))
            yield str(path)
    for directory in list(font_dirs) + (system_font_dirs() if system_fonts else []):
        for path in _files_in(str(directory)):
            if path not in seen:
                seen.add(path)
                yield path


@functools.lru_cache(maxsize=256)
def _files_in(directory: str) -> tuple[str, ...]:
    out: list[str] = []
    for root, _, files in os.walk(directory):
        for name in sorted(files):
            if Path(name).suffix.lower() in _FONT_SUFFIXES:
                out.append(os.path.join(root, name))
    return tuple(out)


def _font_info(path: str, read):
    """``read(path, mtime, size)`` -- cached on all three -- or nothing when the file is gone."""
    try:
        stat = os.stat(path)
    except OSError:
        return frozenset()
    return read(path, stat.st_mtime_ns, stat.st_size)


@functools.lru_cache(maxsize=None)
def _families_of(path: str, mtime: int, size: int) -> frozenset[str]:
    """The families resvg files a font's faces under, lowercased (its ``name`` table)."""
    from ooxml_common.fonts.office import faces_in

    families: set[str] = set()
    for face in faces_in(Path(path), "glyphs"):
        families |= face.rasteriser_families
    return frozenset(families)


@functools.lru_cache(maxsize=None)
def _coverage_of(path: str, mtime: int, size: int) -> frozenset[int]:
    """Every code point a font's faces map (their ``cmap`` tables)."""
    from ooxml_common.fonts.office import read_cmap

    covered: set[int] = set()
    try:
        with open(path, "rb") as handle:
            head = handle.read(12)
            if head[:4] == b"ttcf":
                count = struct.unpack_from(">I", head, 8)[0]
                offsets = struct.unpack(f">{count}I", handle.read(4 * count))
            else:
                offsets = (0,)
            for offset in offsets:
                handle.seek(offset)
                tables = struct.unpack_from(">H", handle.read(12), 4)[0]
                directory = handle.read(16 * tables)
                for index in range(tables):
                    entry = directory[16 * index:16 * index + 16]
                    if entry[:4] == b"cmap":
                        table_offset, length = struct.unpack_from(">II", entry, 8)
                        handle.seek(table_offset)
                        covered.update(read_cmap(handle.read(length)))
                        break
    except Exception:  # noqa: BLE001 -- a damaged font must not stop a render
        pass
    return frozenset(covered)

"""Where PowerPoint finds a face on this machine, read in place.

PowerPoint for Mac draws from three places, and only one of them is a place other
programs look:

1. **PowerPoint's bundle** -- ``Microsoft PowerPoint.app/Contents/Resources/DFonts``, where
   Office keeps Aptos, Calibri, Cambria, Consolas and some two hundred other faces that
   the operating system does not list;
2. **macOS's own fonts** -- ``/System/Library/Fonts`` (and ``Supplemental``),
   ``/Library/Fonts``, ``~/Library/Fonts``;
3. **Office's cloud-font cache** -- ``~/Library/Group Containers/UBF8T346G9.Office/
   FontCache/4/CloudFonts``, one folder per family, where Office downloads faces on
   demand: Aptos Display, the heading face of every deck PowerPoint 365 makes, lives only
   here.

A rasteriser reading "the system's fonts" sees the second and nothing else, so a deck in
Office's default theme rendered on the very machine PowerPoint drew it on came out in a
substitute.  This module finds a family the way PowerPoint does -- **in that order, the
first place that has the family winning** -- and is what :func:`pptx2svg.png.svg_to_png`
hands the rasteriser and what :func:`pptx2svg.convert_pptx_to_svg` measures with, so
that the two agree.

**The order is measured, not assumed** (``tools/make_face_source_probe.py``,
``tools/read_face_source_probe.py``).  29 family-and-style pairs are installed both in
macOS and in PowerPoint's bundle on the machine this was measured on, and two of them
can be told apart in PowerPoint's PDF export:

* **Rockwell.**  macOS ships Rockwell 13.0 and the bundle Rockwell 1.65; they disagree on
  all 95 printable ASCII advances and on every outline.  In all four styles PowerPoint's
  glyphs sit where the *bundle's* advances put them (largest error 0.5--1.5 pt over a
  20-character line, against 21--26 pt for macOS's copy) and all 19 distinct glyph
  outlines it embeds are the bundle's, none macOS's.
* **Symbol.**  macOS's maps Unicode Greek only; the bundle's is symbol-encoded.
  PowerPoint draws ``abgdpqw`` in Symbol as SymbolMT -- the bundle's -- at its advances
  (0.09 pt), against 75.6 pt for macOS's.

The other 27 (Arial, Times New Roman, Verdana, Tahoma, Wingdings and the like) have equal
advances and, where checked, identical outlines in both copies, so either answer draws
the same.  So the bundle comes first -- Word, measured the same way by docx2svg, lays out
with macOS's copy and draws with its bundle's; PowerPoint does both with the bundle's.  No
family is in both the bundle and the cloud cache, or in macOS and the cache, so the cache's
place is moot today; it comes last, as Word's does.

**Nothing is copied.**  Files are read where they are installed: a few header tables to
learn a face's names and style, and its advance widths when a layout needs them.  Only
paths and numbers leave this module.  No font file enters the SVG, the repository or any
file this library writes.

Where the folders are absent -- another operating system, a Mac without Office, CI --
every function here answers "nothing found", and rendering is exactly what it was
without this module.  Set ``PPTX2SVG_OFFICE_FONTS=0`` to make that so on a machine that
has them.
"""

from __future__ import annotations

import functools
import os
import re
import struct
import sys
from dataclasses import dataclass
from pathlib import Path

__all__ = [
    "OFFICE_CLOUD_FONTS",
    "POWERPOINT_FONTS",
    "DrawingPlan",
    "HostFace",
    "cloud_font_dirs",
    "drawing_plan",
    "enabled",
    "find",
    "layout_metrics",
    "metrics",
    "pass_svg",
    "search_dirs",
    "system_font_dirs",
]

#: PowerPoint's own fonts, inside the application bundle.
POWERPOINT_FONTS = Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts")

#: Where Office for Mac keeps the faces it downloads on demand, one folder per family.
OFFICE_CLOUD_FONTS = (
    Path.home() / "Library" / "Group Containers" / "UBF8T346G9.Office" / "FontCache" / "4"
    / "CloudFonts"
)

#: macOS's font folders, in the order its own font list reads them.
MACOS_FONT_DIRS = (
    Path("/System/Library/Fonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("/Library/Fonts"),
    Path.home() / "Library" / "Fonts",
)

_SUFFIXES = (".ttf", ".otf", ".ttc")
#: The locations in the order PowerPoint searches them (see the module docstring).
LOCATIONS = ("powerpoint", "system", "cloud")


def enabled() -> bool:
    """Whether host faces may be used at all: ``PPTX2SVG_OFFICE_FONTS=0`` says not."""
    return os.environ.get("PPTX2SVG_OFFICE_FONTS", "1").strip().lower() not in ("0", "false", "no", "off")


def system_font_dirs() -> tuple[Path, ...]:
    """The operating system's font folders that exist here, searched after PowerPoint's.

    macOS's are the ones measured against PowerPoint.  Elsewhere they are the usual
    places, so a host face is measured from the same file the rasteriser draws it with;
    no PowerPoint runs there to say otherwise.
    """
    if sys.platform == "darwin":
        return tuple(path for path in MACOS_FONT_DIRS if path.is_dir())
    home = Path.home()
    if os.name == "nt":
        windir = Path(os.environ.get("WINDIR", "C:/Windows"))
        roots = [windir / "Fonts", home / "AppData/Local/Microsoft/Windows/Fonts"]
    else:
        roots = [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
                 home / ".local/share/fonts", home / ".fonts"]
    out: list[Path] = []
    for root in roots:
        if root.is_dir():
            out.append(root)
            try:
                out.extend(sorted(path for path in root.rglob("*") if path.is_dir()))
            except OSError:
                pass
    return tuple(out)


def cloud_font_dirs(root: Path | None = None) -> tuple[Path, ...]:
    """The family folders of Office's cloud-font cache, sorted; none where it is absent."""
    try:
        return tuple(sorted(path for path in (root or OFFICE_CLOUD_FONTS).iterdir() if path.is_dir()))
    except OSError:
        return ()


def search_dirs() -> tuple[tuple[str, Path], ...]:
    """``(location, folder)`` for every folder that exists, in PowerPoint's order."""
    out = [("powerpoint", POWERPOINT_FONTS)] if POWERPOINT_FONTS.is_dir() else []
    out.extend(("system", path) for path in system_font_dirs())
    out.extend(("cloud", path) for path in cloud_font_dirs())
    return tuple(out)


# --------------------------------------------------------------------------------------
# Reading a face's names and style, a few tables at a time
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class HostFace:
    """One installed face: where it is, what a deck calls it, what a rasteriser does."""

    path: str
    #: Its index in a collection (``.ttc``), 0 for a single font.
    number: int
    location: str
    #: Name ID 1 in every language, normalised (:func:`~pptx2svg.text.fontmap.family_key`):
    #: the four-style family a deck spells.
    families: frozenset
    #: What resvg's font database files the face under: name ID 16 where the face has
    #: one, else name ID 1 -- every language's, lowercased.
    rasteriser_families: frozenset
    bold: bool
    italic: bool
    #: ``(usWeightClass, usWidthClass, "normal" | "italic" | "oblique")``, as resvg matches.
    css: tuple


def _decode(platform: int, encoding: int, raw: bytes) -> str | None:
    try:
        if platform in (0, 3):
            return raw.decode("utf-16-be")
        if platform == 1 and encoding == 0:
            return raw.decode("mac_roman")
    except UnicodeDecodeError:
        return None
    return None


def _names(table: bytes) -> dict[int, list[str]]:
    out: dict[int, list[str]] = {}
    if len(table) < 6:
        return out
    count, storage = struct.unpack_from(">HH", table, 2)
    for i in range(count):
        entry = 6 + 12 * i
        if entry + 12 > len(table):
            break
        platform, encoding, _language, name_id, length, offset = struct.unpack_from(">6H", table, entry)
        if name_id not in (1, 16):
            continue
        start = storage + offset
        text = _decode(platform, encoding, table[start:start + length])
        if text and text.strip():
            names = out.setdefault(name_id, [])
            if text.strip() not in names:
                names.append(text.strip())
    return out


def _read_at(handle, offset: int, length: int) -> bytes:
    handle.seek(offset)
    return handle.read(length)


def _faces_in(path: Path, location: str) -> list[HostFace]:
    """Every face in ``path``, from its header tables alone; none if it is unreadable."""
    from ..text.fontmap import family_key

    out: list[HostFace] = []
    try:
        with open(path, "rb") as handle:
            head = handle.read(12)
            if head[:4] == b"ttcf":
                count = struct.unpack_from(">I", head, 8)[0]
                offsets = list(struct.unpack(f">{count}I", handle.read(4 * count)))
            else:
                offsets = [0]
            for number, offset in enumerate(offsets):
                header = _read_at(handle, offset, 12)
                tables_count = struct.unpack_from(">H", header, 4)[0]
                directory = _read_at(handle, offset + 12, 16 * tables_count)
                tables = {}
                for i in range(tables_count):
                    tag = directory[16 * i:16 * i + 4]
                    table_offset, length = struct.unpack_from(">II", directory, 16 * i + 8)
                    tables[tag] = (table_offset, length)
                if b"name" not in tables:
                    continue
                names = _names(_read_at(handle, *tables[b"name"]))
                os2 = _read_at(handle, *tables[b"OS/2"]) if b"OS/2" in tables else b""
                head_table = _read_at(handle, *tables[b"head"]) if b"head" in tables else b""
                weight, width, selection = 400, 5, 0
                if len(os2) >= 64:
                    weight, width = struct.unpack_from(">HH", os2, 4)
                    selection = struct.unpack_from(">H", os2, 62)[0]
                mac_style = struct.unpack_from(">H", head_table, 44)[0] if len(head_table) >= 46 else 0
                family = names.get(1, [])
                if not family:
                    continue
                rasteriser = names.get(16) or family
                style = "italic" if selection & 0x01 else "oblique" if selection & 0x200 else "normal"
                out.append(HostFace(
                    path=str(path),
                    number=number,
                    location=location,
                    families=frozenset(family_key(name) for name in family),
                    rasteriser_families=frozenset(name.lower() for name in rasteriser),
                    bold=bool(selection & 0x20 or mac_style & 1),
                    italic=bool(selection & 0x01 or mac_style & 2),
                    css=(weight, width if 1 <= width <= 9 else 5, style),
                ))
    except (OSError, struct.error, ValueError):
        return out
    return out


@functools.lru_cache(maxsize=None)
def _index(dirs: tuple[tuple[str, Path], ...]) -> dict[str, dict[str, list[HostFace]]]:
    """``family key -> location -> faces``, every folder in ``dirs`` read once."""
    index: dict[str, dict[str, list[HostFace]]] = {}
    for location, directory in dirs:
        try:
            paths = sorted(directory.iterdir())
        except OSError:
            continue
        for path in paths:
            if path.suffix.lower() not in _SUFFIXES:
                continue
            for face in _faces_in(path, location):
                for family in face.families:
                    index.setdefault(family, {}).setdefault(location, []).append(face)
    return index


def find(family: str | None) -> tuple[HostFace, ...]:
    """The faces of ``family`` PowerPoint would use here: every style of it in the first
    location that has the family (:data:`LOCATIONS`), the first copy of each style
    winning.  Empty where nothing installed answers to the name."""
    if not family or family.startswith("+") or not enabled():
        return ()
    from ..text.fontmap import family_key

    key = family_key(family)
    found = _index(search_dirs()).get(key)
    if not found:
        return ()
    for location in LOCATIONS:
        faces = found.get(location)
        if faces:
            chosen: dict[tuple[bool, bool], HostFace] = {}
            for face in faces:
                chosen.setdefault((face.bold, face.italic), face)
            return tuple(chosen.values())
    return ()


# --------------------------------------------------------------------------------------
# Advance widths, for the layout
# --------------------------------------------------------------------------------------


#: The tables a layout needs from a face; the outlines are never read.
_LAYOUT_TABLES = (b"head", b"hhea", b"maxp", b"hmtx", b"cmap", b"OS/2", b"name")


def _face_bytes(face: HostFace) -> bytes:
    """The tables of ``face`` that measurement reads, as a font in memory only -- the
    outlines left behind, a collection member's tables gathered the way a single font
    holds them.  Never written anywhere."""
    from ooxml_common.fonts.sfnt import write_sfnt

    with open(face.path, "rb") as handle:
        head = handle.read(12)
        offset = 0
        if head[:4] == b"ttcf":
            offset = struct.unpack(">I", _read_at(handle, 12 + 4 * face.number, 4))[0]
        count = struct.unpack_from(">H", _read_at(handle, offset + 4, 2))[0]
        directory = _read_at(handle, offset + 12, 16 * count)
        tables = []
        for i in range(count):
            tag = directory[16 * i:16 * i + 4]
            if tag in _LAYOUT_TABLES:
                table_offset, length = struct.unpack_from(">II", directory, 16 * i + 8)
                tables.append((tag, _read_at(handle, table_offset, length)))
    return write_sfnt(tables)


@functools.lru_cache(maxsize=64)
def _metrics_of(faces: tuple[HostFace, ...]):
    from ooxml_common.fonts.embedded import _metrics_for

    return _metrics_for({(face.bold, face.italic): _face_bytes(face) for face in faces})


def metrics(family: str | None):
    """A :class:`~pptx2svg.text.metrics.FontMetrics` built from the installed faces of
    ``family`` (:func:`find`), or ``None`` where there are none or they cannot be read."""
    faces = find(family)
    if not faces:
        return None
    try:
        return _metrics_of(faces)
    except Exception:  # noqa: BLE001 -- an unreadable face is one we do not have
        return None


def has_own_table(family: str) -> bool:
    """Whether the static tables measure ``family`` with its own advance widths: a table
    of the very face (Aptos, Cambria), or a clone verified to the same widths (Calibri
    -> Carlito).  Those stay as they are -- they were measured against PowerPoint, kern
    pairs included.  A family reached only by stripping a weight word ("Aptos Light" ->
    Aptos), or measured from another face (Aptos Narrow, Yu Gothic), or not at all, is
    better measured from the installed face where there is one."""
    from ..text.fontmap import family_key, substitution_for

    row = substitution_for(family)
    if row is None or family_key(row.office) != family_key(family):
        return False
    return row.metric_compatible or family_key(row.metrics) == family_key(family)


def layout_metrics(families) -> dict:
    """``family key -> FontMetrics`` from the installed faces, for every family in
    ``families`` that this machine has and the static tables do not measure as itself
    (:func:`has_own_table`).  What :class:`~pptx2svg.text.measure.DefaultTextMeasurer`
    takes as ``extra_metrics``."""
    from ..text.fontmap import family_key

    out = {}
    for family in families:
        if not family or has_own_table(family):
            continue
        found = metrics(family)
        if found is not None:
            out[family_key(family)] = found
    return out


# --------------------------------------------------------------------------------------
# Drawing: which files the rasteriser needs, and in how many passes
# --------------------------------------------------------------------------------------

_FAMILY_ATTRIBUTE = re.compile(r'font-family="([^"]*)"')
_ENTITIES = {"&quot;": '"', "&apos;": "'", "&lt;": "<", "&gt;": ">", "&amp;": "&"}


def stack_names(value: str) -> list[str]:
    """The names in a ``font-family`` attribute's value, unquoted, in order."""
    for entity, char in _ENTITIES.items():
        value = value.replace(entity, char)
    out = []
    for part in value.split(","):
        name = part.strip()
        if len(name) >= 2 and name[0] == name[-1] and name[0] in "'\"":
            name = name[1:-1]
        if name:
            out.append(name)
    return out


@dataclass
class DrawingPlan:
    """What a rasteriser needs to draw an SVG with the faces PowerPoint uses.

    ``passes[0]`` is drawn first, over the whole SVG; each later pass draws only the text
    whose face it holds, over the one before.  More than one pass is needed only where
    two faces a slide uses answer to the same name in the rasteriser -- Aptos Display
    and Aptos are both filed under "Aptos", at the same weight, width and style, so
    loaded together one of them draws both -- and nothing in the SVG can tell them apart
    without copying a Microsoft font to rename it, which this library does not do.
    """

    #: Per pass, the files to load (in order: they win over anything loaded after).
    passes: list[list[str]]
    #: ``font-family`` value -> the pass whose files draw it.  A value absent here is
    #: drawn in pass 0 by whatever else the rasteriser has.
    stacks: dict[str, int]
    #: Whether the files must be loaded *ahead of* the system's fonts, because a system
    #: face answers to the same name: the rasteriser keeps the first it loads.
    ahead_of_system: bool


def drawing_plan(svg: str, *, supplied: frozenset = frozenset()) -> DrawingPlan | None:
    """The installed faces ``svg`` asks for (:func:`find`), grouped into passes; ``None``
    where it asks for none.

    Each ``font-family`` stack is resolved to the first name in it that something
    installed answers to.  A name in ``supplied`` -- family keys of faces the caller
    hands the rasteriser itself, a deck's embedded fonts among them -- stops the search:
    those win, as they do for measurement.
    """
    if not enabled():
        return None
    from ..text.fontmap import family_key

    chosen: dict[str, tuple[str, tuple[HostFace, ...]]] = {}
    for value in dict.fromkeys(_FAMILY_ATTRIBUTE.findall(svg)):
        for name in stack_names(value):
            if family_key(name) in supplied:
                break
            faces = find(name)
            if faces:
                chosen[value] = (family_key(name), faces)
                break
    if not chosen:
        return None

    passes: list[dict[str, tuple[HostFace, ...]]] = []
    family_pass: dict[str, int] = {}
    for key, faces in chosen.values():
        if key in family_pass:
            continue
        names = set().union(*(face.rasteriser_families for face in faces))
        for number, members in enumerate(passes):
            taken = set().union(*(f.rasteriser_families for fs in members.values() for f in fs))
            if not names & taken:
                members[key] = faces
                family_pass[key] = number
                break
        else:
            passes.append({key: faces})
            family_pass[key] = len(passes) - 1

    system_names = _system_rasteriser_names()
    ahead = any(
        face.location != "system" and face.rasteriser_families & system_names
        for members in passes for faces in members.values() for face in faces
    )
    return DrawingPlan(
        passes=[list(dict.fromkeys(f.path for fs in members.values() for f in fs)) for members in passes],
        stacks={value: family_pass[key] for value, (key, _) in chosen.items()},
        ahead_of_system=ahead,
    )


def _system_rasteriser_names() -> frozenset:
    index = _index(search_dirs())
    return frozenset(
        name for locations in index.values() for face in locations.get("system", ())
        for name in face.rasteriser_families
    )


def supplied_families(font_files=(), font_dirs=()) -> frozenset:
    """Family keys of the faces in ``font_files`` and ``font_dirs`` (a caller's own)."""
    out: set[str] = set()
    paths = [Path(path) for path in font_files or ()]
    for directory in font_dirs or ():
        try:
            paths.extend(p for p in Path(directory).rglob("*") if p.suffix.lower() in _SUFFIXES)
        except OSError:
            continue
    for path in paths:
        for face in _faces_in(path, "caller"):
            out.update(face.families)
    return frozenset(out)


#: macOS's font folders as resvg's font database reads them when it loads "the system's
#: fonts" itself (fontdb's ``load_system_fonts``): handed over explicitly instead when
#: Office's files have to be loaded ahead of them (:attr:`DrawingPlan.ahead_of_system`).
def resvg_system_dirs() -> list[str]:
    if sys.platform != "darwin":
        return []
    out = ["/Library/Fonts", "/System/Library/Fonts"]
    assets = Path("/System/Library/AssetsV2")
    try:
        out.extend(str(p) for p in sorted(assets.iterdir())
                   if p.name.startswith("com_apple_MobileAsset_Font") and p.is_dir())
    except OSError:
        pass
    out.append("/Network/Library/Fonts")
    out.append(str(Path.home() / "Library" / "Fonts"))
    return [path for path in out if Path(path).is_dir()]


_ROOT = re.compile(r"<svg\b[^>]*>")
_VIEWBOX = re.compile(r'viewBox="\s*([-\d.eE]+)[\s,]+([-\d.eE]+)[\s,]+([-\d.eE]+)[\s,]+([-\d.eE]+)\s*"')


def pass_svg(svg: str, plan: DrawingPlan, number: int, underlay: bytes | None) -> str:
    """``svg`` for pass ``number`` of ``plan``.

    Pass 0 is the SVG with the text of every later pass hidden.  A later pass is the SVG
    with *everything* hidden but its own text, over ``underlay`` -- the passes before it,
    as PNG, drawn pixel for pixel beneath.  ``visibility`` hides without moving: hidden
    glyphs keep their advance, so nothing after them shifts.
    """
    def element(match: re.Match) -> str:
        value = match.group(1)
        own = plan.stacks.get(value, 0)
        if number == 0:
            return match.group(0) + (' visibility="hidden"' if own != 0 else "")
        return match.group(0) + (' visibility="visible"' if own == number else ' visibility="hidden"')

    out = _FAMILY_ATTRIBUTE.sub(element, svg)
    if number == 0 or underlay is None:
        return out
    import base64

    root = _ROOT.search(out)
    if root is None:
        return out
    box = _VIEWBOX.search(root.group(0))
    if box is not None:
        x, y, width, height = box.groups()
    else:
        x = y = "0"
        width = re.search(r'\bwidth="([\d.]+)', root.group(0)).group(1)  # type: ignore[union-attr]
        height = re.search(r'\bheight="([\d.]+)', root.group(0)).group(1)  # type: ignore[union-attr]
    image = (
        f'<image x="{x}" y="{y}" width="{width}" height="{height}" preserveAspectRatio="none" '
        f'image-rendering="optimizeSpeed" visibility="visible" '
        f'href="data:image/png;base64,{base64.b64encode(underlay).decode("ascii")}"/>'
    )
    # Definitions inherit `visibility` from where they are written, not from where they
    # are used, and a hidden shape in a clip path clips everything away.
    for tag in ("clipPath", "mask", "pattern", "marker", "symbol"):
        out = out.replace(f"<{tag} ", f'<{tag} visibility="visible" ').replace(
            f"<{tag}>", f'<{tag} visibility="visible">')
    head = out[: root.end()]
    head = head[:-1] + ' visibility="hidden">'
    return head + image + out[root.end():]

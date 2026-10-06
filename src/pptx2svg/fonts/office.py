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

**The lookup itself is shared** with docx2svg, which reads the same places in Word's
order: :mod:`ooxml_common.fonts.office` has the locations, each application's measured
order, the header-only index and the advance tables built from installed faces -- with
their legacy ``kern`` pairs, which PowerPoint charges (:mod:`ooxml_common.text.kerning`).
This module is PowerPoint's side of it: the ``PPTX2SVG_OFFICE_FONTS`` switch, and the
rasteriser's plan.

**Nothing is copied.**  Files are read where they are installed: a few header tables to
learn a face's names and style, and its advance widths and kern pairs when a layout needs
them.  Only paths and numbers leave this module.  No font file enters the SVG, the
repository or any file this library writes.

**On by default where Office is installed** (:func:`available`): a conversion measures
and draws with these faces whenever PowerPoint's bundle or Office's cloud cache is on the
machine, with or without the ``pptx2svg-fonts`` bundle, because that is what most closely
resembles PowerPoint -- and the output then depends on the machine.  Where the folders are
absent -- another operating system, a Mac without Office, CI -- every function here
answers "nothing found", and rendering is exactly what it was without this module.  Set
``PPTX2SVG_OFFICE_FONTS=0`` (or ``host_fonts=False``) to make that so on a machine that
has them: the reproducible render.
"""

from __future__ import annotations

import functools
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from ooxml_common.fonts import office as _shared
from ooxml_common.fonts.office import (  # noqa: F401  (re-exported)
    MACOS_FONT_DIRS,
    OFFICE_CLOUD_FONTS,
    POWERPOINT,
    POWERPOINT_FONTS,
    HostFace,
    cloud_font_dirs,
    has_own_table,
    system_font_dirs,
)

__all__ = [
    "OFFICE_CLOUD_FONTS",
    "POWERPOINT_FONTS",
    "DrawingPlan",
    "HostFace",
    "APPLE_COLOR_EMOJI",
    "available",
    "cloud_font_dirs",
    "drawing_plan",
    "emoji_face",
    "enabled",
    "find",
    "layout_metrics",
    "metrics",
    "pass_svg",
    "search_dirs",
    "supplied_faces",
    "system_font_dirs",
    "with_emoji",
]

_SUFFIXES = _shared.FONT_SUFFIXES
#: The locations in the order PowerPoint searches them (see the module docstring).
LOCATIONS = POWERPOINT.layout


def enabled() -> bool:
    """Whether host faces may be used at all: ``PPTX2SVG_OFFICE_FONTS=0`` says not."""
    return os.environ.get("PPTX2SVG_OFFICE_FONTS", "1").strip().lower() not in ("0", "false", "no", "off")


def available() -> bool:
    """Whether Office's fonts are here to be used: PowerPoint's bundle or Office's
    cloud-font cache exists, and :func:`enabled`.  What ``host_fonts=None`` resolves to
    when the machine has them (``ConvertOptions.host_fonts``)."""
    return enabled() and (POWERPOINT_FONTS.is_dir() or bool(cloud_font_dirs(OFFICE_CLOUD_FONTS)))


def search_dirs() -> tuple[tuple[str, Path], ...]:
    """``(location, folder)`` for every folder that exists, in PowerPoint's order."""
    out = [("powerpoint", POWERPOINT_FONTS)] if POWERPOINT_FONTS.is_dir() else []
    out.extend(("system", path) for path in system_font_dirs())
    out.extend(("cloud", path) for path in cloud_font_dirs(OFFICE_CLOUD_FONTS))
    return tuple(out)


def _faces_in(path: Path, location: str) -> list[HostFace]:
    """Every face in ``path``, from its header tables alone (:func:`ooxml_common.fonts.office.faces_in`)."""
    return _shared.faces_in(path, location)


def _index(dirs):
    return _shared.index(dirs)


def find(family: str | None) -> tuple[HostFace, ...]:
    """The faces of ``family`` PowerPoint would use here: every style of it in the first
    location that has the family (:data:`LOCATIONS`), the first copy of each style
    winning.  Empty where nothing installed answers to the name, or where
    ``PPTX2SVG_OFFICE_FONTS=0``."""
    if not enabled():
        return ()
    return _shared.find(family, POWERPOINT, dirs=search_dirs())


def metrics(family: str | None):
    """A :class:`~pptx2svg.text.metrics.FontMetrics` built from the installed faces of
    ``family`` (:func:`find`) -- advances and legacy kern pairs -- or ``None`` where there
    are none or they cannot be read."""
    faces = find(family)
    if not faces:
        return None
    try:
        return _shared.metrics_of(faces)
    except Exception:  # noqa: BLE001 -- an unreadable face is one we do not have
        return None


def layout_metrics(families) -> dict:
    """``family key -> FontMetrics`` from the installed faces, for every family in
    ``families`` that this machine has and the static tables do not measure as itself
    (:func:`has_own_table`).  What :class:`~pptx2svg.text.measure.DefaultTextMeasurer`
    takes as ``extra_metrics``."""
    out = _shared.layout_metrics(families, POWERPOINT, measure=metrics)
    # A bundled family is measured from the bundle's own release, which need not be the
    # one installed here: the bundle's Raleway is later than the 4.026 Office's cloud
    # cache holds, 338 of 340 shared advances apart.  Installed faces now draw a deck's
    # family even where the deck embeds it (as PowerPoint draws it), so the installed
    # release is the one to measure -- a static one: a variable face's advances are its
    # default instance's (Noto Sans JP's is Thin), which no run draws.
    from ooxml_common.fonts import BUNDLED_FAMILIES
    from ooxml_common.text.fontmap import family_key

    bundled = {family_key(name) for name in BUNDLED_FAMILIES}
    for family in families:
        key = family_key(family) if family else None
        if key in bundled and key not in out:
            faces = find(family)
            if faces and not any(face.variable for face in faces):
                table = metrics(family)
                if table is not None:
                    out[key] = table
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


# --------------------------------------------------------------------------------------
# Emoji: the face PowerPoint draws them in when the run's face has none
# --------------------------------------------------------------------------------------

#: macOS's colour emoji face, read in place.
APPLE_COLOR_EMOJI = Path("/System/Library/Fonts/Apple Color Emoji.ttc")
_EMOJI_FAMILY = "Apple Color Emoji"
_TSPAN = re.compile(r'(<tspan\b[^>]*\bfont-family="([^"]*)"[^>]*>)([^<]*)(</tspan>)')


def emoji_face() -> HostFace | None:
    """macOS's Apple Color Emoji, where it is installed and host faces are in use."""
    if not enabled() or sys.platform != "darwin" or not APPLE_COLOR_EMOJI.is_file():
        return None
    faces = _faces_in(APPLE_COLOR_EMOJI, "system")
    return faces[0] if faces else None


@functools.lru_cache(maxsize=64)
def _cmap(face: HostFace) -> frozenset:
    """The code points ``face`` maps, read from its tables alone (no outlines)."""
    try:
        return frozenset(_shared.Face(_shared.face_bytes(face)).cmap)
    except Exception:  # noqa: BLE001 -- an unreadable face covers nothing
        return frozenset()


def with_emoji(svg: str, *, supplied: dict | None = None) -> tuple[str, list[str]]:
    """``svg`` with every emoji its face cannot draw named in Apple Color Emoji, and the
    file to hand the rasteriser for it; ``svg`` unchanged and no file where there is
    nothing to do or the face is not installed.

    **Measured** (``tools/make_font_resolution_probe.py``): PowerPoint for Mac draws ⚡ 📱
    🔒 ✅ ❤ 😀 ✔ in AppleColorEmoji, in Calibri, Aptos, Lato and Noto Sans JP alike --
    none of them has the glyphs -- while a character the run's own face has stays in it
    (Noto Sans JP's ★, Calibri's →).  So a character is named in the emoji face only where
    Apple Color Emoji maps it and the face drawing its run does not: the first face its
    ``font-family`` finds installed (:func:`find`), or one the caller supplies
    (``supplied``, family key -> faces), or the bundle's.

    By name, rather than by handing resvg the file and leaving the rest to its fallback:
    resvg falls back to the first face it loaded that has a glyph, and the emoji face,
    loaded among the caller's files, then drew ↔ ♥ ⚠ that the bundle's Noto Sans JP has.
    """
    emoji = emoji_face()
    if emoji is None:
        return svg, []
    emoji_map = _cmap(emoji)
    if not any(ord(char) > 0x7F and ord(char) in emoji_map for char in svg):
        return svg, []
    from ..text.fontmap import family_key

    supplied = supplied or {}
    bundle = _bundle_faces()

    @functools.lru_cache(maxsize=None)
    def drawing(value: str) -> frozenset:
        for name in stack_names(value):
            key = family_key(name)
            faces = supplied.get(key) or find(name) or bundle.get(key)
            if faces:
                regular = [f for f in faces if not f.bold and not f.italic] or list(faces)
                return _cmap(regular[0])
        return frozenset()

    changed = False

    def tspan(match: re.Match) -> str:
        nonlocal changed
        head, value, text, tail = match.groups()
        if not any(ord(c) > 0x7F and ord(c) in emoji_map for c in text):
            return match.group(0)
        covered = drawing(value)
        out, run = [], []
        for char in text:
            code = ord(char)
            if code > 0x7F and code in emoji_map and code not in covered or (run and code in (0xFE0F, 0x200D)):
                run.append(char)
                continue
            if run:
                out.append(f'<tspan font-family="\'{_EMOJI_FAMILY}\'">{"".join(run)}</tspan>')
                run = []
            out.append(char)
        if run:
            out.append(f'<tspan font-family="\'{_EMOJI_FAMILY}\'">{"".join(run)}</tspan>')
        new = "".join(out)
        if new != text:
            changed = True
        return head + new + tail

    out = _TSPAN.sub(tspan, svg)
    return (out, [emoji.path]) if changed else (svg, [])


@functools.lru_cache(maxsize=1)
def _bundle_faces() -> dict:
    from ooxml_common.fonts import bundle_dir

    directory = bundle_dir()
    if directory is None:
        return {}
    return {key: tuple(places.get("bundle", ())) for key, places in _index((("bundle", Path(directory)),)).items()}


def supplied_faces(font_files=(), font_dirs=()) -> dict:
    """``family key -> faces`` of the faces in ``font_files`` and ``font_dirs`` (a caller's own)."""
    out: dict[str, list[HostFace]] = {}
    paths = [Path(path) for path in font_files or ()]
    for directory in font_dirs or ():
        try:
            paths.extend(p for p in Path(directory).rglob("*") if p.suffix.lower() in _SUFFIXES)
        except OSError:
            continue
    for path in paths:
        for face in _faces_in(path, "caller"):
            for key in face.families:
                out.setdefault(key, []).append(face)
    return {key: tuple(faces) for key, faces in out.items()}


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

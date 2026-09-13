"""Where the fonts are, whether we have them, and what it means when we do not.

Why this exists
---------------

Layout is computed from a table of advance widths (:mod:`pptx2svg.text.metrics`) and then
the SVG names a ``font-family`` for a rasteriser to draw.  If the host does not have that
face, the rasteriser quietly picks another one.  resvg does not warn; on a bare Debian box
it has nothing Office-like to pick.  The text still appears, at widths nothing computed,
and the output is wrong in a way no error message points at.  Rendering the same string in
Calibri, Carlito, Aptos, Noto Sans JP and Lato on the development Mac produced
*byte-identical* PNGs: five different faces, one silent fallback.

The fix is to draw with fonts we ship, measured from the files we ship, so the invariant

    every family in :data:`pptx2svg.text.metrics.METRICS` is a family we can draw, and
    every family we draw has a metrics table generated from that exact file

holds by construction.  ``tools/extract_font_metrics.py`` generates the table and
``tests/test_fonts.py`` fails if the two drift apart.  (Aptos and Cambria are deliberate
exceptions -- measured, not shipped; see :mod:`pptx2svg.text.fontmap`.)

Two modes, and you can always tell which one you are in
-------------------------------------------------------

The font files live in a **separate distribution**, ``pptx2svg-fonts``, so the base wheel
stays small and standard-library-only.  That gives two possible states, and the whole
point of this module is that they are never confused:

``bundled``
    ``pptx2svg-fonts`` is installed.  Rendering uses its directory with
    ``skip_system_fonts=True``: the host's font inventory is ignored entirely, so the same
    deck rasterises to the same bytes on a laptop and on a build server.  This is what
    ``pip install 'pptx2svg[fonts]'`` gets you, and what the documentation recommends.

``system``
    It is not installed.  Rendering falls back to whatever the machine happens to have,
    which is by definition not reproducible -- and a warning saying exactly that goes into
    ``ConvertOptions.warnings``, because a fallback nobody is told about is the bug this
    subsystem was built to eliminate.  Use this mode deliberately: with your own
    ``font_dirs``, or on a host whose fonts you control.

:func:`bundle_mode` is the single source of truth for which one is active, and
``pptx2svg fonts`` prints it at the top of its report.

What the bundle contains
------------------------

Office's default faces are proprietary and cannot be redistributed.  Each one is drawn
with an open font built to the same advance widths, so substituting it changes the glyph
outlines but not where anything lands:

==================  ============  ===========================  ==============
Office face         Substitute    Debian package               Licence
==================  ============  ===========================  ==============
Calibri             Carlito       ``fonts-crosextra-carlito``  SIL OFL 1.1
Cambria             Caladea       ``fonts-crosextra-caladea``  SIL OFL 1.1
Arial, Helvetica    Arimo         ``fonts-croscore``           SIL OFL 1.1
Times New Roman     Tinos         ``fonts-croscore``           SIL OFL 1.1
Courier New         Cousine       ``fonts-croscore``           SIL OFL 1.1
Japanese Gothic     Noto Sans JP  ``fonts-noto-cjk``           SIL OFL 1.1
==================  ============  ===========================  ==============

plus Lato and Raleway, which are not substitutes for anything -- a deck that names Lato
wants Lato.  Arimo, Tinos and Cousine are the Chrome OS core fonts; Liberation
Sans/Serif/Mono are derived from them and measure identically (checked character by
character: 0 of 191 differ), so a host with ``fonts-liberation2`` is equally correct.

Aptos, Microsoft's Office default since 2023, has **no** metric-compatible open clone, and
neither, despite the universal claim, does Cambria: Caladea's advance widths run 4.5%
narrow.  See :mod:`pptx2svg.text.fontmap` for what happens to those two.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

__all__ = [
    "BUNDLED_FAMILIES",
    "GENERIC_FAMILY_DEFAULTS",
    "INSTALL_HINT",
    "available_families",
    "bundle_dir",
    "bundle_mode",
    "font_dirs",
    "missing_families",
]

#: Families the ``pptx2svg-fonts`` distribution supplies, spelled exactly as the font
#: files name themselves -- that is what a rasteriser matches ``font-family`` against.
BUNDLED_FAMILIES = frozenset(
    {"Arimo", "Caladea", "Carlito", "Cousine", "Lato", "Noto Sans JP", "Raleway", "Tinos"}
)

#: What to tell someone who has ended up in ``system`` mode by accident.
INSTALL_HINT = "pip install 'pptx2svg[fonts]'"

#: What the CSS generic families resolve to when the bundle is in use.  Every
#: ``font-family`` this library emits ends in one of these (see
#: :func:`pptx2svg.text.fontmap.font_family_value`), so they are the last thing standing
#: between an unknown face and text that does not draw at all: with ``skip_system_fonts``
#: set, resvg renders *nothing* for a family it cannot resolve, and its own defaults are
#: "Arial"/"Times New Roman"/"Courier New", none of which the bundle contains.  Pointing
#: them at the bundle turns an invisible paragraph into one at roughly the right size.
GENERIC_FAMILY_DEFAULTS = {
    "font_family": "Arimo",
    "sans_serif_family": "Arimo",
    "serif_family": "Tinos",
    "monospace_family": "Cousine",
    "cursive_family": "Tinos",
    "fantasy_family": "Arimo",
}


def bundle_dir() -> Path | None:
    """Font directory of the ``pptx2svg-fonts`` distribution, or ``None`` if absent.

    Looked up through :mod:`importlib.util` rather than by importing it, so the core
    library keeps no import-time dependency on the extra: it is data we hand to a
    rasteriser, not code we run.

    A real filesystem path, because resvg and cairo both open files by name.  An install
    that keeps the package zipped (zipimport, some PyInstaller modes) therefore has no
    usable bundle, and this returns ``None`` rather than a path that does not exist.
    """
    spec = importlib.util.find_spec("pptx2svg_fonts")
    if spec is None or not spec.origin:
        return None
    directory = Path(spec.origin).resolve().parent / "files"
    return directory if directory.is_dir() else None


def bundle_mode() -> str:
    """``"bundled"`` when rendering is reproducible, ``"system"`` when it is not."""
    return "bundled" if bundle_dir() is not None else "system"


def font_dirs() -> list[str]:
    """Directories to hand a rasteriser.  Empty in ``system`` mode."""
    directory = bundle_dir()
    return [str(directory)] if directory is not None else []


def available_families() -> frozenset[str]:
    """Families we can draw ourselves.  Empty in ``system`` mode.

    Deliberately empty rather than "whatever fontconfig reports": in ``system`` mode we
    genuinely do not know what the rasteriser will pick, and claiming otherwise would
    reintroduce the confident-but-wrong answer this module exists to remove.
    """
    return BUNDLED_FAMILIES if bundle_dir() is not None else frozenset()


def missing_families() -> frozenset[str]:
    """Bundled families this installation cannot supply."""
    return BUNDLED_FAMILIES - available_families()

"""Every font pptx2svg draws with.

This distribution carries no code beyond this module -- it is font *data* that
:mod:`pptx2svg.fonts` discovers by name.  pptx2svg never imports it; it locates the
directory with :func:`importlib.util.find_spec` and hands the path to the rasteriser, so
the core library keeps its promise of importing nothing outside the standard library.

It is separate from ``pptx2svg`` because the two have genuinely different audiences.  SVG
output embeds no fonts at all, and callers who only want SVG -- or who point
``font_dirs`` at their own corporate faces -- should not download 20 MB of typefaces to
get a 100 kB pure-Python library.  Install it, as the documentation recommends, with::

    pip install 'pptx2svg[fonts]'

rather than by name: the extra pins a compatible version.

===============  ============================================  ==============
Family           Stands in for                                 Licence
===============  ============================================  ==============
Carlito          Calibri, Calibri Light                        SIL OFL 1.1
Caladea          Cambria (approximately -- 4.5% narrow)        SIL OFL 1.1
Arimo            Arial, Helvetica                              SIL OFL 1.1
Tinos            Times New Roman                               SIL OFL 1.1
Cousine          Courier New                                   SIL OFL 1.1
Noto Sans JP     MS Gothic, Meiryo, Yu Gothic, MS Mincho, ...  SIL OFL 1.1
Lato             itself                                        SIL OFL 1.1
Raleway          itself                                        SIL OFL 1.1
===============  ============================================  ==============

Every file is redistributed byte-for-byte as published on Google Fonts -- none has been
subsetted, renamed or otherwise modified -- so the Reserved Font Name clauses that
Carlito, Lato, Raleway and Noto Sans JP carry are satisfied.  The full licence text for
each family is in ``licenses/`` next to the fonts, as the OFL requires.

On Debian the same faces are ``fonts-crosextra-carlito``, ``fonts-crosextra-caladea``,
``fonts-croscore``, ``fonts-noto-cjk``, ``fonts-lato`` and ``fonts-raleway``.  Those work
when pointed at with ``--font-dir``, but they are whatever version the distribution
shipped; only this distribution pins the exact files, and only pinned files give
byte-identical PNGs across machines.
"""

from __future__ import annotations

from pathlib import Path

__all__ = ["FILES_DIR", "LICENSE_DIR", "FAMILIES", "__version__"]

__version__ = "0.1.0"

#: Directory pptx2svg hands to the rasteriser.
FILES_DIR = Path(__file__).resolve().parent / "files"

#: Full SIL OFL 1.1 text for every family above, one file each.
LICENSE_DIR = Path(__file__).resolve().parent / "licenses"

#: Families supplied here, spelled as the fonts name themselves.  Kept in step with
#: :data:`pptx2svg.fonts.BUNDLED_FAMILIES` by tests/test_fonts.py.
FAMILIES = frozenset(
    {"Arimo", "Caladea", "Carlito", "Cousine", "Lato", "Noto Sans JP", "Raleway", "Tinos"}
)

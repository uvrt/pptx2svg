#!/usr/bin/env python3
"""Build ``tests/fixtures/sample-cjk.pptx`` -- ``sample.pptx`` with an East Asian face.

Why this deck exists
--------------------

``sample.pptx`` is six slides of real Japanese text and it has never been scored.
``tools/fidelity.py`` skips it, and every attempt to explain the skip before this file
was written got the reason wrong, so it is worth stating what was actually measured.

The deck's Japanese runs resolve their face through ``+mj-ea`` / ``+mn-ea``, and both of
its font schemes write ``<a:ea typeface=""/>``.  **The deck therefore names no East Asian
face of its own.**  The ``<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>`` further down
the same scheme is the fallback our renderer resolves through; ＭＳ Ｐゴシック is a face
this PowerPoint cannot use, so its export draws MS Gothic and MS Mincho instead and the
two sides are comparing different outlines at different widths.

Which of the three OOXML slots can be *made* to change PowerPoint's mind was measured
rather than guessed.  Seven exports, each a copy of a corpus deck with one theme string
changed, each read back through its PDF's ``/BaseFont`` entries:

=========================  =============================  =============================
deck                       theme edit                     what PowerPoint drew
=========================  =============================  =============================
``real-basic-theme``       -- (as committed)              ``MS-Gothic``, ``MS-Mincho``
``real-basic-theme``       ``Jpan`` -> ``MS PGothic``     ``MS-Gothic``, ``MS-Mincho``
``real-basic-theme``       ``Jpan`` -> ``MS Gothic``      ``MS-Gothic``, ``MS-Mincho``
``real-basic-theme``       ``a:ea`` -> ＭＳ Ｐゴシック         ``MS-Gothic``, ``MS-Mincho``
``real-basic-theme``       ``a:ea`` -> ``Noto Sans JP``   ``MS-Gothic``, ``MS-Mincho``
``sample``                 ``Jpan`` -> ``Noto Sans JP``   ``MS-Gothic``, ``MS-Mincho``
``sample``                 ``a:ea`` -> ``Noto Sans JP``   ``NotoSansJP-Thin_Regular``
                                                          and ``_Bold``
=========================  =============================  =============================

Two things fall out of that.  **``real-basic-theme`` cannot be fixed from its theme at
all** -- not even with a face PowerPoint draws happily in three other decks -- because
its Japanese runs name a *Latin* face as their own ``a:ea`` (Lato, Arial, Raleway) and
PowerPoint discards that and supplies a Japanese default of its own without consulting
the theme.  And for ``sample`` the working lever is the scheme's ``<a:ea>``, not the
``Jpan`` list, which this PowerPoint does not reach for that deck.

The ``Jpan`` half of that does not generalise and is left unsettled here:
``real-financial-report``'s ``Jpan`` entries *are* consulted -- rewriting its 游ゴシック
pair to ``MS Mincho`` takes YuGothic-Regular out of its export and puts MS-Mincho in.
Knowing *why* would be interesting; it is not needed to act on what was measured.

What this file changes, and what it deliberately does not
---------------------------------------------------------

One string, twice per theme: ``<a:ea typeface=""/>`` becomes
``<a:ea typeface="Noto Sans JP"/>`` in both font schemes of both themes, and the
``script="Jpan"`` entry is pointed at the same face so the theme says one thing rather
than two.  Nothing else -- not a slide, not a layout, not a master, not one character of
text.

**The authored face is drawn by neither side, and that is the point of saying so here.**
``sample.pptx`` asks for ＭＳ Ｐゴシック and this deck does not; it asks for Noto Sans JP,
which both renderers then ink.  The score belongs to the derived deck and to nothing
else.  What it measures is this library's Japanese *layout* -- line breaking, run
splitting, line advance -- with the glyph-shape term taken out of it, which is the only
thing a comparison against PowerPoint can honestly be about.

``sample.pptx`` itself is left exactly as md-pptx generated it, which is the point of
deriving rather than editing.  Its empty ``a:ea`` is a real-world case this project
handles deliberately (see ``src/pptx2svg/text/fontmap.py``), it is the input behind
several committed measurements, and a fixture that quietly stopped being the file it
claims to be would cost more than it bought.  ``tests/deckbuilder.py`` takes the same
line for the same reason.

Noto Sans JP because it is the one Japanese face **both** renderers can draw here: we
ship it in ``pptx2svg-fonts`` and measure from its own tables, and PowerPoint draws it
once it is installed for the user -- see ``FONTS.md``, "Making the oracle draw Japanese".
That is what makes this deck scorable where the other two Japanese decks are not: the
face the deck names is the face both sides ink.

Usage::

    python3 tools/make_cjk_deck.py                    # writes the fixture
    python3 tools/make_cjk_deck.py /tmp/probe.pptx    # somewhere else
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "tests/fixtures/sample.pptx"
TARGET = ROOT / "tests/fixtures/sample-cjk.pptx"

#: The face both renderers draw.  Spelled as the font names itself (nameID 16), which is
#: what a deck asks for and what ``fidelity.requested_faces`` reads back.
FACE = "Noto Sans JP"

#: The empty East Asian slot ``sample.pptx`` carries in both schemes of both themes.
EMPTY_EA = '<a:ea typeface=""/>'

#: The script-fallback entry.  Repointed for consistency rather than for effect: this
#: PowerPoint ignores it, and ours would otherwise resolve Japanese to a second face the
#: theme also names, which leaves ``fidelity.font_profile`` unable to tell whether the
#: export used the face this deck is about.
JPAN_ENTRY = '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>'


def rewrite(theme: str) -> tuple[str, int]:
    """``(theme with an East Asian face, number of substitutions made)``."""
    replaced = theme.count(EMPTY_EA) + theme.count(JPAN_ENTRY)
    theme = theme.replace(EMPTY_EA, f'<a:ea typeface="{FACE}"/>')
    theme = theme.replace(JPAN_ENTRY, f'<a:font script="Jpan" typeface="{FACE}"/>')
    return theme, replaced


def write_deck(source: Path, target: Path) -> int:
    """Copy ``source`` to ``target`` with its themes' East Asian face filled in.

    Every other part is copied byte for byte.  Each entry is written back through its own
    ``ZipInfo`` rather than by name, which carries the source's modification times and
    compression across: ``writestr`` given a plain name stamps *now* instead, and two
    runs of this script would then differ in every local header while differing in no
    byte that means anything.  A fixture regenerated from an unchanged source should be
    the same file.
    """
    total = 0
    with zipfile.ZipFile(source) as original:
        target.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename.startswith("ppt/theme/"):
                    text, count = rewrite(data.decode("utf-8"))
                    data = text.encode("utf-8")
                    total += count
                out.writestr(item, data)
    return total


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else TARGET
    if not SOURCE.exists():
        print(f"no source deck at {SOURCE}", file=sys.stderr)
        return 2
    count = write_deck(SOURCE, target)
    if not count:
        print(
            f"{SOURCE.name} no longer carries the theme entries this script rewrites; "
            "the derivation is stale, not the deck",
            file=sys.stderr,
        )
        return 1
    print(f"{SOURCE.name} -> {target} ({count} theme entries -> {FACE}, "
          f"{target.stat().st_size:,} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

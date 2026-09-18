#!/usr/bin/env python3
"""Regenerate ``text/metrics.py`` and ``text/kerning.py`` from the fonts we actually ship.

The whole point of the metrics table is that layout is computed from the *same* advance
widths the rasteriser will draw with.  Keeping that true by hand does not work -- the
table this replaces was copied out of pptx-glimpse, and by the time it was checked its
Noto Sans JP column disagreed with the shipped Noto Sans JP for 186 of 191 characters.
So the table is generated, from the exact font files the ``pptx2svg-fonts``
distribution ships, and a test re-runs this and fails if the checked-in file drifts.

Aptos is the one entry with no font behind it.  Microsoft's new Office default is
proprietary and has no metric-compatible clone, so there is nothing we could ship that
draws it correctly.  Its advance widths are *measured* -- from the copy Office installs
on this machine, cross-checked against PowerPoint's own PDF export (see
``tools/measure_aptos.py``) -- and measurements are facts, not font software: no Aptos
outline, table or file is redistributed.  This is the same footing on which Carlito and
Arimo exist at all.

Two later additions stand on exactly that footing and are worth naming here, because both
are places where "measure with what we draw with" is the *wrong* rule:

* **``line_gap`` is read from the Office face, not from the clone.**  A line gap never
  reaches a glyph -- it moves baselines -- so the number that matters is the one
  PowerPoint paced the deck with.  Three of the four clones agree with their original and
  Tinos does not (87 against ``times.ttf``'s 0), which is why this is a table and not a
  field read.  See :data:`LINE_GAP_SOURCES`.
* **Thirteen fixed-pitch faces have tables and no files.**  SimSun, MingLiU, BatangChe,
  Lucida Console and their siblings have two advances between them -- half an em and a
  full em -- so their entire advance table is two integers, verified against the installed
  face by :func:`verify_fixed_pitch` and pruned to the two dozen characters that break the
  rule.  Nothing is downloaded, licensed or shipped for them.  See :data:`FIXED_PITCH`.

**Kern pairs are the third thing on that footing, and the largest.**  A kern value is a
measurement of a design, exactly as an advance width is, and the same faces are read for
both.  They go to ``text/kerning.py`` rather than into the advance table because they are
twice its size; :func:`classify_kern` is where the size question was settled and
:func:`effective_kern` is where the three sources -- ``PairPos`` format 1, format 2 and
the legacy ``kern`` table -- are reduced to one function of two characters.

Usage::

    python3 tools/extract_font_metrics.py --check     # exit 1 if either file is stale
    python3 tools/extract_font_metrics.py --write     # rewrite both

Needs fontTools (``pip install pptx2svg[measure]``).  Dev-only: the library itself never
reads a font file at runtime.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "src"))
# So a source checkout works without `pip install -e packages/pptx2svg-fonts`.
sys.path.insert(0, str(ROOT / "packages" / "pptx2svg-fonts" / "src"))

TARGET = ROOT / "src" / "pptx2svg" / "text" / "metrics.py"
KERN_TARGET = ROOT / "src" / "pptx2svg" / "text" / "kerning.py"

BEGIN = "# --- BEGIN GENERATED METRICS (tools/extract_font_metrics.py) ---"
END = "# --- END GENERATED METRICS ---"
KERN_BEGIN = "# --- BEGIN GENERATED KERNING (tools/extract_font_metrics.py) ---"
KERN_END = "# --- END GENERATED KERNING ---"

#: Faces we measure but never draw, resolved through the same local font profile the
#: fidelity harness uses (``tools/fidelity.py --write-profile``).  Going through the
#: profile rather than hard-coding paths matters twice over: Office keeps Aptos Display
#: only in an on-demand cloud-font cache under an opaque numeric filename, and the two
#: tools then agree by construction about which file a given face means.
#:
#: Read only.  Nothing from those directories is ever copied into the repository or a
#: wheel; see this module's docstring for why publishing the *measurements* is a
#: different question from redistributing the fonts.


# --------------------------------------------------------------------------------------
# Which characters get an entry
# --------------------------------------------------------------------------------------

def _sample_characters() -> list[str]:
    """The characters worth storing a real advance width for.

    Everything outside this set falls back to ``default_width`` (or ``cjk_width``), which
    is close enough for wrapping and wrong enough that widening the set is the cheapest
    fidelity win available.  Real decks are full of curly quotes, en dashes, ellipses and
    bullets, and those used to measure at the 0.6 em guess.
    """
    chars: list[str] = []
    chars += [chr(c) for c in range(0x20, 0x7F)]      # printable ASCII
    chars += [chr(c) for c in range(0xA0, 0x100)]     # Latin-1 supplement
    chars += [chr(c) for c in range(0x100, 0x180)]    # Latin Extended-A
    chars += [
        "ˆ", "˜",                            # circumflex, small tilde
        "–", "—", "―",                  # en/em dash, horizontal bar
        "‘", "’", "‚", "‛",        # single quotes
        "“", "”", "„",                  # double quotes
        "†", "‡", "•", "…",        # dagger, bullet, ellipsis
        "‰", "‹", "›", "⁄",        # per mille, guillemets, fraction
        "€", "™", "−",                  # euro, trademark, minus
        "■", "▪", "○", "●",        # square/circle bullets
        "◦", "⁃", "→",                  # hollow bullet, hyphen bullet, arrow
    ]
    seen: set[str] = set()
    ordered: list[str] = []
    for char in chars:
        if char not in seen:
            seen.add(char)
            ordered.append(char)
    return ordered


SAMPLE = _sample_characters()

#: The advance ``cjk_width`` stands for: every CJK character with no row of its own.
#:
#: A kanji, deliberately, not the hiragana this used to probe with.  Kana are enumerated
#: below and ideographs are not, so what is left for this fallback to cover is almost
#: entirely ideographs -- and in a proportional face the two disagree.  ＭＳ Ｐゴシック
#: draws 編 at a full em and あ at 0.941, so probing with あ quietly measured every kanji
#: in the corpus 6% narrow.  For a monospaced face the choice makes no difference.
CJK_PROBE = "編"  # CJK UNIFIED IDEOGRAPH-7DE8


def _cjk_sample_characters() -> list[str]:
    """Japanese characters whose advance is worth storing individually.

    ``cjk_width`` assumes every CJK character is one em wide, which is true of the
    monospaced faces (MS Gothic, Noto Sans JP) and false of the proportional ones.  The
    "P" in ``ＭＳ Ｐゴシック`` *means* proportional: measured from the file Office ships,
    its katakana run from 0.648 em (ト) to 1.0, and its ideographic comma and full stop
    are 0.664.  Measuring those at 1.0 overstates a line of katakana by up to a third,
    which wraps it early and then draws the wrapped text with the correct outlines -- the
    layout is wrong while every glyph is right, which is the hardest kind of error to see.

    Kanji are deliberately not enumerated: they are full-width in every Japanese face in
    practice, and there are tens of thousands of them.  ``cjk_width`` remains their rule.
    """
    ranges = (
        (0x3000, 0x303F),  # CJK symbols and punctuation: 、。「」〜
        (0x3041, 0x309F),  # hiragana
        (0x30A0, 0x30FF),  # katakana, including the long-vowel mark ー
        (0xFF01, 0xFF60),  # fullwidth ASCII forms
        (0xFF61, 0xFF9F),  # halfwidth katakana
    )
    return [chr(c) for start, end in ranges for c in range(start, end + 1)]


CJK_SAMPLE = _cjk_sample_characters()


# --------------------------------------------------------------------------------------
# Reading one face
# --------------------------------------------------------------------------------------

def _open(path: Path, weight: int | None, index: int = 0):
    """Open a face, instancing a variable font at ``weight`` when one is asked for.

    Arimo, Raleway and Noto Sans JP ship as single variable files.  resvg reads the
    weight axis correctly -- Arimo at ``wght=700`` renders pixel-for-pixel like static
    Liberation Sans Bold -- so the bold table has to be taken from the same instance the
    rasteriser will produce, not from the file's default instance.

    ``index`` selects a face inside a TrueType collection.  It matters for exactly the
    faces this was extended for: ``msgothic.ttc`` holds ＭＳ ゴシック, MS UI Gothic and
    ＭＳ Ｐゴシック at indices 0, 1 and 2, and only the last of the three is proportional.
    Taking index 0 for all of them would have measured the proportional face as
    monospaced and hidden the bug this is here to fix.
    """
    from fontTools.ttLib import TTFont

    font = TTFont(os.fspath(path), fontNumber=index, lazy=True)
    if weight is None or "fvar" not in font:
        return font
    from fontTools.varLib.instancer import instantiateVariableFont

    return instantiateVariableFont(
        TTFont(os.fspath(path), fontNumber=index), {"wght": weight}, inplace=False
    )


def collection_index(path: Path, family: str) -> int:
    """Which face inside a ``.ttc`` carries ``family``; 0 for a plain font file."""
    if path.suffix.lower() != ".ttc":
        return 0
    from fontTools.ttLib import TTCollection

    for index, font in enumerate(TTCollection(os.fspath(path), lazy=True).fonts):
        for record in font["name"].names:
            if record.nameID == 1 and record.toUnicode() == family:
                return index
    return 0


def _widths(font) -> tuple[int, dict[str, int], int]:
    units_per_em = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]

    widths: dict[str, int] = {}
    for char in SAMPLE:
        glyph = cmap.get(ord(char))
        if glyph is not None:
            widths[char] = hmtx[glyph][0]

    cjk_glyph = cmap.get(ord(CJK_PROBE))
    cjk_width = hmtx[cjk_glyph][0] if cjk_glyph is not None else units_per_em

    # Only the characters that disagree with cjk_width earn a row.  A monospaced face
    # adds nothing here; a proportional one adds the hundred-odd entries that make it
    # proportional, and the table stays readable either way.
    for char in CJK_SAMPLE:
        glyph = cmap.get(ord(char))
        if glyph is not None and hmtx[glyph][0] != cjk_width:
            widths[char] = hmtx[glyph][0]

    return units_per_em, widths, cjk_width


def _default_width(widths: dict[str, int], units_per_em: int) -> int:
    """Stand-in for characters with no entry.

    The mean over the lower-case Latin alphabet, which beats the font's own
    ``advanceWidthMax`` or the width of a space -- both of which earlier tables used and
    both of which are far from typical.
    """
    alphabet = [widths[c] for c in "abcdefghijklmnopqrstuvwxyz" if c in widths]
    return round(sum(alphabet) / len(alphabet)) if alphabet else units_per_em // 2


def read_face(
    regular: Path, bold: Path, bold_weight: int | None, index: int = 0
) -> dict:
    """Everything one entry of the table needs, from the regular and bold files."""
    font = _open(regular, None, index)
    units_per_em, widths, cjk_width = _widths(font)
    hhea = font["hhea"]

    bold_font = _open(bold, bold_weight, index)
    bold_upm, bold_widths, bold_cjk = _widths(bold_font)
    if bold_upm != units_per_em:  # pragma: no cover - would mean a mismatched pair
        raise SystemExit(f"{regular.name} and {bold.name} disagree on unitsPerEm")

    return {
        "units_per_em": units_per_em,
        # hhea, not OS/2 sTypo: the first-baseline rule in text/measure.py was calibrated
        # against PowerPoint using hhea's descent, and the two differ (Arial: -434 vs
        # -431, Carlito: -550 vs -512).
        "ascender": hhea.ascender,
        "descender": hhea.descender,
        # hhea again, and for the same reason the two above are hhea: PowerPoint's own
        # line pitch is ascent + descent + *this*.  For a face we only measure, this file
        # is the Office face and the number is already the right one; for a clone it is
        # overwritten from the Office original -- see LINE_GAP_SOURCES.
        "line_gap": hhea.lineGap,
        "default_width": _default_width(widths, units_per_em),
        "cjk_width": cjk_width,
        "widths": widths,
        "bold_default_width": _default_width(bold_widths, units_per_em),
        "bold_cjk_width": bold_cjk,
        "bold_widths": bold_widths,
    }


def prune_fixed_pitch(face: dict) -> dict:
    """Drop every row the two fallback constants already answer.

    A fixed-pitch East Asian face has exactly two advances -- half an em for Latin and
    half-width kana, a full em for everything ideographic -- and ``default_width`` and
    ``cjk_width`` are those two numbers.  Writing out three hundred and fifty rows that
    all repeat one of them costs a hundred and fifty lines an entry and says nothing;
    what is worth reading is the handful of characters that are *not* where the rule puts
    them, which for these faces is the typographic punctuation an East Asian design draws
    full-width (``—``, ``“``, ``…``, ``■``) although Unicode files it under Latin.

    Behaviour is unchanged by construction: :func:`pptx2svg.text.measure` reaches the same
    number for a pruned row through the fallback it was equal to.
    """
    pruned = dict(face)
    for widths_key, default_key, cjk_key in (
        ("widths", "default_width", "cjk_width"),
        ("bold_widths", "bold_default_width", "bold_cjk_width"),
    ):
        default, cjk = face[default_key], face[cjk_key]
        pruned[widths_key] = {
            char: width
            for char, width in face[widths_key].items()
            if width != (cjk if _is_cjk(char) else default)
        }
    return pruned


def _is_cjk(char: str) -> bool:
    """The ranges :func:`pptx2svg.text.measure.is_cjk` answers for, imported from it.

    Pruning has to agree with the measurer character for character: a row dropped because
    it equalled ``cjk_width`` is only harmless if the measurer will ask ``cjk_width`` for
    that character too.
    """
    from pptx2svg.text.measure import is_cjk

    return is_cjk(ord(char))


# --------------------------------------------------------------------------------------
# Reading one face's kern pairs
# --------------------------------------------------------------------------------------

def _kern_lookups(font) -> list:
    """The GPOS lookups the ``kern`` feature reaches, across every script it is under.

    Taking the union over scripts rather than one script's list is deliberate: a deck
    mixes Latin and Japanese in one text box and PowerPoint kerns both, so the table has
    to answer for both.  Where a face registers ``kern`` under several scripts they share
    the same lookups anyway -- checked over all eight bundled faces.
    """
    if "GPOS" not in font:
        return []
    gpos = font["GPOS"].table
    if gpos is None or gpos.FeatureList is None:
        return []
    wanted = set()
    for record in gpos.FeatureList.FeatureRecord:
        if record.FeatureTag == "kern":
            wanted.update(record.Feature.LookupListIndex)
    return [gpos.LookupList.Lookup[index] for index in sorted(wanted)]


def _pair_subtables(lookups) -> list[list]:
    """Pair-adjustment subtables, grouped by lookup; see :func:`effective_kern`."""
    groups = []
    for lookup in lookups:
        group = []
        for subtable in lookup.SubTable:
            if lookup.LookupType == 9:  # extension positioning: unwrap it
                subtable = subtable.ExtSubTable
                kind = subtable.LookupType
            else:
                kind = lookup.LookupType
            if kind == 2:
                group.append(subtable)
        if group:
            groups.append(group)
    return groups


def _x_advance(value) -> int:
    return 0 if value is None else (getattr(value, "XAdvance", 0) or 0)


def _pair_adjustment(subtable, first: str, second: str) -> int | None:
    """What one ``PairPos`` subtable does to ``first``, or ``None`` if it covers neither.

    ``None`` and ``0`` are different answers and the difference is load-bearing: a
    subtable that covers the pair and adjusts it by nothing *stops* the lookup, so a
    later subtable in the same lookup must not be consulted.
    """
    coverage = subtable.Coverage.glyphs
    if subtable.Format == 1:
        try:
            index = coverage.index(first)
        except ValueError:
            return None
        for record in subtable.PairSet[index].PairValueRecord:
            if record.SecondGlyph == second:
                return _x_advance(record.Value1)
        return None
    if first not in coverage:
        return None
    first_class = subtable.ClassDef1.classDefs.get(first, 0)
    second_class = subtable.ClassDef2.classDefs.get(second, 0)
    if first_class >= subtable.Class1Count or second_class >= subtable.Class2Count:
        return None
    return _x_advance(subtable.Class1Record[first_class].Class2Record[second_class].Value1)


def effective_kern(font, characters: list[str]) -> dict[tuple[str, str], int]:
    """``{(first, second): units}`` for every pair among ``characters`` that kerns.

    The OpenType rule this implements, and both halves matter: *within* one lookup the
    first subtable that covers the pair wins and the rest are skipped, while each lookup
    is its own pass over the run, so several lookups' adjustments **add**.  Reading only
    the first subtable of the first lookup -- the obvious shortcut -- measures Lato and
    Cambria wrong, which both split their kerning across a format-1 lookup and a
    format-2 one.
    """
    cmap = font.getBestCmap()
    groups = _pair_subtables(_kern_lookups(font))
    glyphs = {char: cmap[ord(char)] for char in characters if ord(char) in cmap}
    pairs: dict[tuple[str, str], int] = {}
    for first, first_glyph in glyphs.items():
        for second, second_glyph in glyphs.items():
            total = 0
            for group in groups:
                for subtable in group:
                    value = _pair_adjustment(subtable, first_glyph, second_glyph)
                    if value is not None:
                        total += value
                        break
            if total:
                pairs[(first, second)] = total
    return pairs


def legacy_kern(font, characters: list[str]) -> dict[tuple[str, str], int]:
    """The same, from the pre-OpenType ``kern`` table.

    Only reached for a face that has one and no GPOS ``kern`` feature.  Every face here
    that carries a ``kern`` table carries the feature as well and the feature is the one
    a shaper reads, so this is the fallback for a face none of ours turns out to be --
    written because the roadmap asked for it and because the next Office face to arrive
    may well be one.
    """
    if "kern" not in font:
        return {}
    cmap = font.getBestCmap()
    by_glyph: dict[str, str] = {}
    for char in characters:
        glyph = cmap.get(ord(char))
        if glyph is not None:
            by_glyph.setdefault(glyph, char)
    pairs: dict[tuple[str, str], int] = {}
    try:
        subtables = font["kern"].kernTables
    except Exception:  # pragma: no cover - a malformed kern table is not worth a crash
        return {}
    for subtable in subtables:
        for (first, second), value in subtable.kernTable.items():
            if value and first in by_glyph and second in by_glyph:
                key = (by_glyph[first], by_glyph[second])
                pairs[key] = pairs.get(key, 0) + value
    return {key: value for key, value in pairs.items() if value}


def classify_kern(pairs: dict[tuple[str, str], int]):
    """``{(a, b): v}`` -> ``(left classes, right classes, matrix)``.

    Two characters share a left class when their whole row of adjustments is identical,
    and a right class when their column is.  Derived from the function rather than from
    the font's own ``ClassDef`` tables, so an explicit-pair face, a class-matrix face and
    a legacy ``kern`` table all compress into one shape -- and so the result is canonical,
    which is what lets ``--check`` compare two runs byte for byte.

    The decomposition is exact whenever one exists, and one always does: in the worst
    case every character gets its own class and the matrix is the pair list again.  See
    :func:`verify_kern` for the assertion that it round-trips.
    """
    rows: dict[str, dict[str, int]] = {}
    columns: dict[str, dict[str, int]] = {}
    for (first, second), value in pairs.items():
        rows.setdefault(first, {})[second] = value
        columns.setdefault(second, {})[first] = value

    def grouped(vectors: dict[str, dict[str, int]]) -> list[str]:
        by_signature: dict[tuple, list[str]] = {}
        for char, vector in vectors.items():
            by_signature.setdefault(tuple(sorted(vector.items())), []).append(char)
        order = sorted(by_signature, key=lambda sig: min(by_signature[sig]))
        return ["".join(sorted(by_signature[sig])) for sig in order]

    left = grouped(rows)
    right = grouped(columns)
    left_index = {char: i for i, group in enumerate(left, 1) for char in group}
    right_index = {char: i for i, group in enumerate(right, 1) for char in group}

    matrix: dict[int, dict[int, int]] = {}
    for (first, second), value in pairs.items():
        matrix.setdefault(left_index[first], {})[right_index[second]] = value
    return tuple(left), tuple(right), matrix


def verify_kern(key: str, pairs: dict[tuple[str, str], int], classified) -> None:
    """Assert the class matrix reproduces the pair list exactly, in both directions."""
    from pptx2svg.text.kerning import KernTable

    left, right, matrix = classified
    table = KernTable(left=left, right=right, matrix=matrix)
    for (first, second), value in pairs.items():
        if table.adjustment(first, second) != value:
            raise SystemExit(f"{key}: class matrix loses the pair {first!r}{second!r}")
    for group in left:
        for first in group:
            for other in right:
                for second in other:
                    if table.adjustment(first, second) != pairs.get((first, second), 0):
                        raise SystemExit(
                            f"{key}: class matrix invents the pair {first!r}{second!r}"
                        )


def read_kern(regular: Path, bold: Path, bold_weight: int | None, index: int = 0) -> dict:
    """One face's upright and bold kern tables, or ``{}`` for a face that does not kern."""
    characters = SAMPLE + CJK_SAMPLE
    upright = _open(regular, None, index)
    pairs = effective_kern(upright, characters) or legacy_kern(upright, characters)
    bold_font = _open(bold, bold_weight, index)
    bold_pairs = effective_kern(bold_font, characters) or legacy_kern(bold_font, characters)
    if not pairs and not bold_pairs:
        return {}
    return {"pairs": pairs, "bold_pairs": bold_pairs}


# --------------------------------------------------------------------------------------
# Which faces make up the table
# --------------------------------------------------------------------------------------

def source_faces() -> dict[str, tuple[Path, Path, int | None]]:
    """Metrics key -> (regular file, bold file, weight to instance the bold at)."""
    from pptx2svg.fonts import bundle_dir

    bundle = bundle_dir()
    if bundle is None:
        raise SystemExit(
            "the pptx2svg-fonts distribution is not importable, so there are no font\n"
            "files to measure.  Install it from this checkout with\n"
            "    pip install -e packages/pptx2svg-fonts"
        )
    return {
        "Carlito": (bundle / "Carlito-Regular.ttf", bundle / "Carlito-Bold.ttf", None),
        "Arimo": (bundle / "Arimo[wght].ttf", bundle / "Arimo[wght].ttf", 700),
        "Tinos": (bundle / "Tinos-Regular.ttf", bundle / "Tinos-Bold.ttf", None),
        "Cousine": (bundle / "Cousine-Regular.ttf", bundle / "Cousine-Bold.ttf", None),
        "Caladea": (bundle / "Caladea-Regular.ttf", bundle / "Caladea-Bold.ttf", None),
        "Noto Sans JP": (
            bundle / "NotoSansJP[wght].ttf", bundle / "NotoSansJP[wght].ttf", 700
        ),
        "Lato": (bundle / "Lato-Regular.ttf", bundle / "Lato-Bold.ttf", None),
        "Raleway": (bundle / "Raleway[wght].ttf", bundle / "Raleway[wght].ttf", 700),
    }


#: Metrics key -> the Office face whose ``hhea`` lineGap the entry must carry.
#:
#: Everything else in a clone's entry is read off the clone, because metric compatibility
#: means the clone *is* the measurement.  The line gap is the exception: it is the one
#: vertical metric the clones are free to differ on, and two of the four do.  Measured
#: here, in units of 2048:
#:
#:     Calibri 0 / Carlito 0            agree
#:     Arial 67 / Arimo 67              agree
#:     Times New Roman 0 / Tinos 87     **disagree**
#:     Courier New 0 / Cousine 0        agree
#:
#: Tinos is the whole reason this table exists rather than a call to ``hhea.lineGap`` on
#: the bundled file.  Its 87 is a real property of Tinos and the wrong number for a deck
#: that asked for Times New Roman, which is what PowerPoint laid out with; taking it would
#: buy a 0.33 pt fix on Arial at the cost of a 0.42 pt regression on Times New Roman.
#:
#: The four keys not listed here take their own file's gap, and that is right rather than
#: a gap in the list: Lato, Raleway and Noto Sans JP are drawn under the name the deck
#: asked for, and Caladea's entry is only ever reached by a deck that names Caladea --
#: Cambria has a measured table of its own.  A deck that names Tinos or Liberation Serif
#: directly therefore lays out with a zero gap where the real Tinos has 87; that is
#: exactly what it did before this field existed, and it is the smaller of the two errors.
LINE_GAP_SOURCES = {
    "Carlito": "Calibri",
    "Arimo": "Arial",
    "Tinos": "Times New Roman",
    "Cousine": "Courier New",
}

#: Measured-only faces whose entire advance table is two constants.
#:
#: Every one of these has exactly two advances -- half an em and a full em -- across the
#: whole sample, so ``default_width`` and ``cjk_width`` answer all but a couple of dozen
#: characters and :func:`prune_fixed_pitch` drops the rest.  :func:`verify_fixed_pitch`
#: asserts the property against the file rather than trusting the name: a face that had
#: quietly gone proportional would fail the build instead of shipping a flat table.
#:
#: Measured on the copies Office installs here, ``(unitsPerEm, half, full)``:
#:
#:     SimSun, NSimSun, SimHei, KaiTi, FangSong        (256, 128, 256)
#:     MingLiU, MingLiU_HKSCS                          (1024, 512, 1024)
#:     BatangChe, GulimChe, DotumChe, GungsuhChe       (1024, 512, 1024)
#:     ＭＳ ゴシック, ＭＳ 明朝                         (256, 128, 256)
#:     Lucida Console, Lucida Sans Typewriter          (2048, 1234, --)
#:
#: The two Lucidas are Latin-only -- no kana, no ideographs -- so their ``cjk_width`` is
#: the ``units_per_em`` non-answer :func:`_widths` writes for a face with no glyph for the
#: probe kanji, and only their half-width column means anything.
#:
#: The two ＭＳ faces are listed because they *are* fixed-pitch, and pruning them shortens
#: two entries that were already checked in at full length; nothing about them changes.
FIXED_PITCH = frozenset(
    {
        "ＭＳ ゴシック", "ＭＳ 明朝",
        "SimSun", "NSimSun", "SimHei", "KaiTi", "FangSong",
        "MingLiU", "MingLiU_HKSCS",
        "BatangChe", "GulimChe", "DotumChe", "GungsuhChe",
        "Lucida Console", "Lucida Sans Typewriter",
    }
)


def verify_fixed_pitch(key: str, face: dict) -> None:
    """Refuse to write a flat table for a face that is not flat.

    Two claims, both checked against the file and not against the family name:

    * every printable ASCII character the face covers advances exactly
      ``default_width`` -- the "fixed pitch at the width you record" claim; and
    * every kana, ideograph and full-width form it covers advances exactly ``cjk_width``,
      which is what it means for :func:`prune_fixed_pitch` to have left no CJK row behind.

    ＭＳ Ｐゴシック is the face this is guarding against: it lives in the same file as
    ＭＳ ゴシック, differs only in being proportional, and a generator that took the wrong
    index would produce a table that looked perfectly healthy.
    """
    default, cjk = face["default_width"], face["cjk_width"]
    odd = {
        char: width
        for char, width in face["widths"].items()
        if 0x20 <= ord(char) < 0x7F and width != default
    }
    if odd:
        raise SystemExit(
            f"{key} is not fixed-pitch: {len(odd)} ASCII advances differ from "
            f"{default} ({odd})"
        )
    east_asian = {
        char: width for char, width in face["widths"].items() if _is_cjk(char)
    }
    if east_asian:
        raise SystemExit(
            f"{key} is not full-width in the East Asian ranges: {len(east_asian)} "
            f"advances differ from {cjk} ({east_asian})"
        )


#: Faces we measure but never draw: Office-only, and no open font reproduces their
#: advance widths.  See the module docstring for why storing measurements is legitimate.
#:
#: Cambria is here for a reason worth recording.  Caladea is universally described as
#: "metric-compatible with Cambria", and for the line box it is -- but its advance widths
#: are not Cambria's: measured over four representative strings it runs 4.5% narrow
#: (ratios 0.9555, 0.9384, 0.9525, 0.9542).  Taking that claim on trust is how this
#: project ended up with a Noto Sans JP table that disagreed with Noto Sans JP.
#:
#: Aptos Display is a *cloud* font: Office does not install it, it downloads it on first
#: use into ~/Library/Group Containers/UBF8T346G9.Office/FontCache.  It took reading
#: /BaseFont out of PowerPoint's own PDF export to notice it was there at all -- the
#: export embeds "AptosDisplay", so PowerPoint had it even while every font directory
#: said otherwise.  Two of the corpus's seven decks use it, and they are the two that
#: scored worst.
#: The Japanese faces are here because a deck can need one without ever naming it.
#: ``sample.pptx`` declares ``<a:ea typeface=""/>`` -- no East-Asian face at all -- and
#: then sets Japanese body text, so both renderers fall back: PowerPoint to MS Gothic,
#: MS PGothic and MS Mincho (they are embedded in its PDF export), and we to the same
#: family, which Office installs.  We drew the right outlines and measured them with
#: Noto Sans JP's widths, because the table had no entry for the face we were drawing.
#: The thirteen after the Japanese four are the cheapest entries on this page: their
#: advance table is two constants, so they add thirteen family names to what we can
#: measure without adding a byte to any wheel.  Eleven are the fixed-pitch CJK faces --
#: the ``Che`` suffix on the Korean ones and the ``N`` on ``NSimSun`` *mean* fixed-pitch,
#: and MingLiU and SimSun are fixed-pitch outright -- and two are Latin monospace.  See
#: :data:`FIXED_PITCH` for the measurements and ``text/fontmap.py`` for what each is drawn
#: with, which is a separate and less happy question.
MEASURED_ONLY = (
    "Cambria",
    "Aptos",
    "Aptos Display",
    "ＭＳ Ｐゴシック",
    "ＭＳ ゴシック",
    "ＭＳ Ｐ明朝",
    "ＭＳ 明朝",
    "SimSun",
    "NSimSun",
    "SimHei",
    "KaiTi",
    "FangSong",
    "MingLiU",
    "MingLiU_HKSCS",
    "BatangChe",
    "GulimChe",
    "DotumChe",
    "GungsuhChe",
    "Lucida Console",
    "Lucida Sans Typewriter",
)


def measured_only_faces() -> dict[str, tuple[Path, Path, int | None, int]]:
    """Resolve :data:`MEASURED_ONLY` against the local profile, skipping what is absent."""
    sys.path.insert(0, str(HERE))
    import fidelity

    profile = fidelity.load_profile()
    if profile is None:
        return {}
    faces = profile.get("faces", {})
    resolved: dict[str, tuple[Path, Path, int | None, int]] = {}
    for family in MEASURED_ONLY:
        styles = faces.get(family)
        if not styles or "regular" not in styles:
            continue
        regular = Path(styles["regular"]["path"])
        # A face with no bold cut measures its bold from the upright, which the generated
        # table then reports as a 1.0 ratio.  That is a real property of the font, not a
        # gap: it is what a rasteriser will draw too.
        bold = Path(styles.get("bold", styles["regular"])["path"])
        # The profile keys a collection's families to one shared path, so the face has to
        # be picked out by name here rather than trusted to be first.
        resolved[family] = (regular, bold, None, collection_index(regular, family))
    return resolved


# --------------------------------------------------------------------------------------
# Emitting the table
# --------------------------------------------------------------------------------------

def _literal(char: str) -> str:
    """Source spelling for a dict key: readable for ASCII, escaped above it."""
    if char == '"':
        return '"\\""'
    if char == "\\":
        return '"\\\\"'
    if 0x20 <= ord(char) < 0x7F:
        return f'"{char}"'
    return f'"\\u{ord(char):04x}"'


def _width_rows(widths: dict[str, int]) -> list[str]:
    rows: list[str] = []
    row: list[str] = []
    for char, width in widths.items():
        row.append(f"{_literal(char)}: {width},")
        if len(row) == 6:
            rows.append("            " + " ".join(row))
            row = []
    if row:
        rows.append("            " + " ".join(row))
    return rows


def _width_block(name: str, widths: dict[str, int]) -> list[str]:
    """One ``widths={...}`` field.  Empty on one line, because a fixed-pitch face's is."""
    if not widths:
        return [f"        {name}={{}},"]
    return [f"        {name}={{"] + _width_rows(widths) + ["        },"]


def _string_literal(text: str) -> str:
    """One class's characters as a Python string literal: readable ASCII, escaped above."""
    out = []
    for char in text:
        if char == '"':
            out.append('\\"')
        elif char == "\\":
            out.append("\\\\")
        elif 0x20 <= ord(char) < 0x7F:
            out.append(char)
        else:
            out.append(f"\\u{ord(char):04x}")
    return '"' + "".join(out) + '"'


def _wrapped(pieces: list[str], indent: str, opener: str, closer: str, width: int = 96):
    """``opener`` + comma-separated ``pieces`` + ``closer``, folded to ``width`` columns."""
    lines = [indent + opener]
    row = ""
    for piece in pieces:
        candidate = f"{row} {piece}" if row else piece
        if row and len(indent) + 4 + len(candidate) > width:
            lines.append(indent + "    " + row)
            row = piece
        else:
            row = candidate
    if row:
        lines.append(indent + "    " + row)
    lines.append(indent + closer)
    return lines


def _class_block(name: str, classes: tuple[str, ...]) -> list[str]:
    if not classes:
        return [f"        {name}=(),"]
    pieces = [_string_literal(group) + "," for group in classes]
    return _wrapped(pieces, "        ", f"{name}=(", "),")


def _matrix_block(name: str, matrix: dict[int, dict[int, int]]) -> list[str]:
    if not matrix:
        return [f"        {name}={{}},"]
    lines = [f"        {name}={{"]
    for first in sorted(matrix):
        row = matrix[first]
        cells = [f"{second}: {row[second]}," for second in sorted(row)]
        one_line = f"            {first}: {{{' '.join(cells)[:-1]}}},"
        # A row that fits on one line goes on one line.  Folding every row costs two
        # lines of braces and twelve columns of indent each, which over four thousand
        # rows is a third of this file.
        if len(one_line) <= 96:
            lines.append(one_line)
        else:
            lines += _wrapped(cells, "            ", f"{first}: {{", "},")
    lines.append("        },")
    return lines


def render_kern_entry(key: str, kern: dict) -> str:
    left, right, matrix = classify_kern(kern["pairs"])
    verify_kern(key, kern["pairs"], (left, right, matrix))
    bold_left, bold_right, bold_matrix = classify_kern(kern["bold_pairs"])
    verify_kern(f"{key} bold", kern["bold_pairs"], (bold_left, bold_right, bold_matrix))
    pairs = len(kern["pairs"]) + len(kern["bold_pairs"])
    cells = sum(len(row) for row in matrix.values())
    cells += sum(len(row) for row in bold_matrix.values())
    lines = [f'    "{key}": KernTable(']
    lines.append(f"        # {_kern_note(pairs, cells)}")
    lines += _class_block("left", left)
    lines += _class_block("right", right)
    lines += _matrix_block("matrix", matrix)
    lines += _class_block("bold_left", bold_left)
    lines += _class_block("bold_right", bold_right)
    lines += _matrix_block("bold_matrix", bold_matrix)
    lines.append("    ),")
    return "\n".join(lines)


def _kern_note(pairs: int, cells: int) -> str:
    return f"{pairs} kern pairs in {cells} matrix cells; see the module docstring"


def render_entry(key: str, face: dict, note: str) -> str:
    lines = [f'    "{key}": FontMetrics(']
    lines.append(f"        # {note}")
    lines.append(f'        units_per_em={face["units_per_em"]},')
    lines.append(f'        ascender={face["ascender"]},')
    lines.append(f'        descender={face["descender"]},')
    lines.append(f'        line_gap={face["line_gap"]},')
    lines.append(f'        default_width={face["default_width"]},')
    lines.append(f'        cjk_width={face["cjk_width"]},')
    lines += _width_block("widths", face["widths"])
    lines.append(f'        bold_default_width={face["bold_default_width"]},')
    lines.append(f'        bold_cjk_width={face["bold_cjk_width"]},')
    lines += _width_block("bold_widths", face["bold_widths"])
    # The pairs themselves live in text/kerning.py: they are twice the size of every
    # advance width in this file put together, and a reader looking up how wide "W" is
    # should not have to scroll past them.  ``.get`` rather than ``[...]`` because a face
    # that does not kern has no entry there at all.
    lines.append(f'        kerning=_KERN.get("{key}"),')
    lines.append("    ),")
    return "\n".join(lines)


NOTES = {
    "Carlito": "metric-compatible with Calibri; shipped in pptx2svg-fonts",
    "Arimo": "metric-compatible with Arial and Helvetica; shipped in pptx2svg-fonts",
    "Tinos": "metric-compatible with Times New Roman; shipped in pptx2svg-fonts",
    "Cousine": "metric-compatible with Courier New; shipped in pptx2svg-fonts",
    "Caladea": "the closest serif we ship to Cambria; shipped in pptx2svg-fonts",
    "Noto Sans JP": "stands in for Japanese faces; shipped in pptx2svg-fonts",
    "Lato": "itself; shipped in pptx2svg-fonts",
    "Raleway": "itself; shipped in pptx2svg-fonts",
    "Cambria": "MEASURED ONLY -- proprietary; Caladea is 4.5% narrower, so it is not it",
    "Aptos": "MEASURED ONLY -- proprietary, no clone exists, drawn with a substitute",
    "Aptos Display": "MEASURED ONLY -- Office cloud font, no clone exists",
    "ＭＳ Ｐゴシック": "MEASURED ONLY -- proportional: kana run 0.648-1.0 em, not full-width",
    "ＭＳ ゴシック": "MEASURED ONLY -- fixed pitch 128/256 of 256, the non-proportional cut",
    "ＭＳ Ｐ明朝": "MEASURED ONLY -- proportional serif; PowerPoint falls back to it too",
    "ＭＳ 明朝": "MEASURED ONLY -- fixed pitch 128/256 of 256, the serif's non-proportional cut",
    "SimSun": "MEASURED ONLY -- fixed pitch 128/256 of 256; rows are full-width Latin",
    "NSimSun": "MEASURED ONLY -- fixed pitch 128/256 of 256; the SimSun file's face 1",
    "SimHei": "MEASURED ONLY -- fixed pitch 128/256 of 256; rows are full-width Latin",
    "KaiTi": "MEASURED ONLY -- fixed pitch 128/256 of 256; rows are full-width Latin",
    "FangSong": "MEASURED ONLY -- fixed pitch 128/256 of 256; rows are full-width Latin",
    "MingLiU": "MEASURED ONLY -- fixed pitch 512/1024 of 1024; rows are full-width Latin",
    "MingLiU_HKSCS": "MEASURED ONLY -- fixed pitch 512/1024; the MingLiU file's face 2",
    "BatangChe": "MEASURED ONLY -- fixed pitch 512/1024 of 1024; 65 full-width Latin rows",
    "GulimChe": "MEASURED ONLY -- fixed pitch 512/1024 of 1024; 65 full-width Latin rows",
    "DotumChe": "MEASURED ONLY -- fixed pitch 512/1024 of 1024; 65 full-width Latin rows",
    "GungsuhChe": "MEASURED ONLY -- fixed pitch 512/1024 of 1024; 65 full-width Latin rows",
    "Lucida Console": "MEASURED ONLY -- monospace 1234/2048 = 0.6025 em; no CJK at all",
    "Lucida Sans Typewriter": "MEASURED ONLY -- monospace 1234/2048, the same pitch",
}


def office_line_gaps() -> dict[str, int]:
    """:data:`LINE_GAP_SOURCES` resolved against the local profile, skipping what is absent.

    Read-only, and only ``hhea.lineGap`` comes out -- one integer per face.  Nothing else
    from an Office file reaches the clones' entries, which stay generated from the
    bundled files exactly as before.
    """
    sys.path.insert(0, str(HERE))
    import fidelity

    profile = fidelity.load_profile()
    if profile is None:
        return {}
    faces = profile.get("faces", {})
    gaps: dict[str, int] = {}
    for key, office in LINE_GAP_SOURCES.items():
        styles = faces.get(office)
        if not styles or "regular" not in styles:
            continue
        path = Path(styles["regular"]["path"])
        if not path.exists():
            continue
        gaps[key] = _open(path, None, collection_index(path, office))["hhea"].lineGap
    return gaps


def _line_gap_for(key: str, measured: dict[str, int]) -> int | None:
    """The gap a bundled face's entry carries: Office's, or the checked-in one, or none.

    The fallback order matters for ``--check`` on CI, which has no Office and must not see
    drift.  Re-emitting what is checked in keeps it green; falling through to ``None``
    rather than to the bundled file's own gap is what stops a machine with no Office
    quietly writing Tinos' 87 into a column that is supposed to hold Times New Roman's 0.
    """
    if key in measured:
        return measured[key]
    if key not in LINE_GAP_SOURCES:
        return None  # caller keeps the face's own gap; see LINE_GAP_SOURCES.
    from pptx2svg.text.metrics import METRICS as CURRENT

    existing = CURRENT.get(key)
    return existing.line_gap if existing is not None else None


def build_block() -> str:
    parts = ["METRICS: dict[str, FontMetrics] = {"]
    gaps = office_line_gaps()
    for key, (regular, bold, weight) in source_faces().items():
        for path in (regular, bold):
            if not path.exists():
                raise SystemExit(f"missing bundled font: {path}")
        face = read_face(regular, bold, weight)
        if key in LINE_GAP_SOURCES:
            face["line_gap"] = _line_gap_for(key, gaps)
        parts.append(render_entry(key, face, NOTES[key]))
    local = measured_only_faces()
    for key in MEASURED_ONLY:
        entry = local.get(key)
        if entry is not None and entry[0].exists() and entry[1].exists():
            face = read_face(*entry)
            if key in FIXED_PITCH:
                face = prune_fixed_pitch(face)
                verify_fixed_pitch(key, face)
            parts.append(render_entry(key, face, NOTES[key]))
        else:
            # No Office on this machine, or no font profile written yet.  Re-emit what is
            # already checked in rather than dropping the entry: ``--check`` has to stay
            # usable on CI, and its job is to catch drift between the *bundled* fonts and
            # their tables.  Measured-only faces can only be regenerated where the
            # licensed originals are installed.
            parts.append(render_entry(key, _checked_in(key), NOTES[key]))
    parts.append("}")
    return "\n".join(parts)


def build_kern_block() -> str:
    """The ``KERNING`` table for ``src/pptx2svg/text/kerning.py``.

    Faces that do not kern are simply absent, which is what ``_KERN.get`` in the metrics
    table expects.  The emitted source is executed and compared against the data it came
    from before it is returned: the class strings fold across lines with implicit
    concatenation, and that is exactly the kind of thing a generator gets subtly wrong.
    """
    parts = ["KERNING: dict[str, KernTable] = {"]
    tables: dict[str, dict] = {}
    for key, (regular, bold, weight) in source_faces().items():
        kern = read_kern(regular, bold, weight)
        if kern:
            tables[key] = kern
            parts.append(render_kern_entry(key, kern))
    local = measured_only_faces()
    for key in MEASURED_ONLY:
        entry = local.get(key)
        if entry is not None and entry[0].exists() and entry[1].exists():
            kern = read_kern(*entry)
            if kern:
                tables[key] = kern
                parts.append(render_kern_entry(key, kern))
        else:
            # Same rule as the advance widths: without the licensed original installed,
            # re-emit what is checked in rather than dropping the face.
            checked = _checked_in_kern(key)
            if checked is not None:
                parts.append(checked)
    parts.append("}")
    block = "\n".join(parts)
    _verify_kern_block(block, tables)
    return block


def _verify_kern_block(block: str, tables: dict[str, dict]) -> None:
    from pptx2svg.text.kerning import KernTable

    namespace: dict = {"KernTable": KernTable}
    exec(compile(block, "<kerning>", "exec"), namespace)
    emitted = namespace["KERNING"]
    for key, kern in tables.items():
        table = emitted[key]
        for bold, pairs in ((False, kern["pairs"]), (True, kern["bold_pairs"])):
            for (first, second), value in pairs.items():
                if table.adjustment(first, second, bold) != value:
                    raise SystemExit(f"{key}: emitted source loses {first!r}{second!r}")


def _checked_in_kern(key: str) -> str | None:
    """Re-render the checked-in entry for a face we cannot measure here.

    Byte-identical to what the measured path would write, which is the whole point: the
    pair count in the note is recovered from the matrix (a cell stands for
    ``len(left class) * len(right class)`` pairs) rather than guessed at, so ``--check``
    stays green on a machine with no Office.
    """
    from pptx2svg.text.kerning import KERNING as CURRENT

    table = CURRENT.get(key)
    if table is None:
        return None
    pairs = _kern_pair_count(table.left, table.right, table.matrix)
    pairs += _kern_pair_count(table.bold_left, table.bold_right, table.bold_matrix)
    cells = sum(len(row) for row in table.matrix.values())
    cells += sum(len(row) for row in table.bold_matrix.values())
    lines = [f'    "{key}": KernTable(']
    lines.append(f"        # {_kern_note(pairs, cells)}")
    lines += _class_block("left", table.left)
    lines += _class_block("right", table.right)
    lines += _matrix_block("matrix", table.matrix)
    lines += _class_block("bold_left", table.bold_left)
    lines += _class_block("bold_right", table.bold_right)
    lines += _matrix_block("bold_matrix", table.bold_matrix)
    lines.append("    ),")
    return "\n".join(lines)


def _kern_pair_count(left, right, matrix) -> int:
    return sum(
        len(left[first - 1]) * len(right[second - 1])
        for first, row in matrix.items()
        for second in row
    )


def _checked_in(key: str) -> dict:
    from pptx2svg.text.metrics import METRICS as CURRENT

    existing = CURRENT.get(key)
    if existing is None:
        raise SystemExit(
            f"{key} is neither installed locally nor already present in metrics.py; "
            "regenerate on a machine with Microsoft Office"
        )
    return {
        "units_per_em": existing.units_per_em,
        "ascender": existing.ascender,
        "descender": existing.descender,
        "line_gap": existing.line_gap,
        "default_width": existing.default_width,
        "cjk_width": existing.cjk_width,
        "widths": existing.widths,
        "bold_default_width": existing.bold_default_width,
        "bold_cjk_width": existing.bold_cjk_width,
        "bold_widths": existing.bold_widths,
    }


def splice(text: str, block: str, begin: str = BEGIN, end: str = END) -> str:
    start = text.index(begin) + len(begin)
    stop = text.index(end)
    return text[:start] + "\n\n" + block + "\n\n" + text[stop:]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true", help="rewrite the tables in place")
    group.add_argument("--check", action="store_true", help="exit 1 if a table is stale")
    args = parser.parse_args()

    # The kern block is built first: it is what ``_KERN.get`` in the metrics block reads,
    # so writing metrics.py against a stale kerning.py would key entries to tables that
    # are about to change.
    jobs = [
        (KERN_TARGET, KERN_BEGIN, KERN_END, build_kern_block()),
        (TARGET, BEGIN, END, build_block()),
    ]

    stale = []
    for target, begin, end, block in jobs:
        current = target.read_text(encoding="utf-8")
        updated = splice(current, block, begin, end)
        if updated == current:
            continue
        if args.write:
            target.write_text(updated, encoding="utf-8")
            print(f"wrote {target.relative_to(ROOT)}")
        else:
            stale.append(target.relative_to(ROOT))

    if args.write:
        return 0

    if stale:
        print(
            f"{', '.join(str(path) for path in stale)} does not match the bundled fonts; "
            "run tools/extract_font_metrics.py --write",
            file=sys.stderr,
        )
        return 1
    print("metrics.py and kerning.py match the bundled fonts")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

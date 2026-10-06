#!/usr/bin/env python3
"""Build a deck that measures *which kern pairs* PowerPoint applies: the face's legacy
``kern`` table, its OpenType ``GPOS`` ``kern`` feature, or neither.

A face can carry its kerning twice.  Aptos, Calibri, Arial, Times New Roman, Cambria and
Meiryo hold a legacy ``kern`` table (format 0) that is a strict subset of their ``GPOS``
feature, agreeing on every pair the two share; Segoe UI holds both and they disagree;
Verdana, Tahoma, Trebuchet MS and Garamond hold only the legacy table; Lato, Raleway, Noto
Sans JP and Yu Gothic only ``GPOS``.  Which one PowerPoint reads is invisible on a pair
both tables agree on, so every probe line is built from pairs of one *class* in one face:

``legacy``  in the legacy table only (Segoe UI and the legacy-only faces);
``gpos``    in ``GPOS`` only -- Aptos's ``ss`` (-33/2048 em) is the trial's case;
``agree``   in both, with the same value;
``differ``  in both, with different values (Segoe UI);

with a separator between the pairs that kerns with nothing in either table, so the gap
after every glyph is either a probed pair or zero.  The Japanese faces get a line of kana
pairs too, set through ``a:ea`` as East Asian text is.

Every probe is one line in its own text box -- no wrap, no insets, no autofit, left
aligned, 24 pt, ``kern="100"`` so the master's 12 pt kerning threshold is not what is
measured -- and the deck's master, layouts and theme are ``real-financial-report.pptx``'s,
emptied, as ``make_face_source_probe.py`` does.  ``tools/read_kern_source_probe.py``
reads PowerPoint's PDF export.

Usage::

    python3 tools/make_kern_source_probe.py ~/kern-source-probe.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/kern-source-probe.pptx ~/kern-source-probe.pdf
    python3 tools/read_kern_source_probe.py ~/kern-source-probe.pdf

Needs fontTools to choose the pairs, from the faces where Office and macOS installed them
(:mod:`pptx2svg.fonts.office`); nothing is copied.  A probe deck and its export are
throwaway and must be deleted again.
"""

from __future__ import annotations

import sys
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "src"))

from make_feature_sweep import NAMESPACES, PT  # noqa: E402
from make_style_probe import write_deck  # noqa: E402

SOURCE = ROOT / "tests/fixtures/real-financial-report.pptx"

LEFT, TOP, ROW, BOX_W, BOX_H = 20, 14, 38, 920, 32
SIZE = 2400
ROWS_PER_SLIDE = 13
PAIRS_PER_LINE = 8

#: The faces probed and the file each is read from, by family as a deck names it.
FACES = (
    "Aptos", "Aptos Display", "Calibri", "Arial", "Times New Roman", "Cambria", "Segoe UI",
    "Verdana", "Tahoma", "Lato", "Raleway", "Noto Sans JP", "Yu Gothic", "Meiryo",
    # Faces only macOS installs: static, unlike Noto Sans JP, and kerned in GPOS alone
    # (Minion Pro, Hiragino Sans) or in both tables (PT Sans, Avenir Next).
    "Minion Pro", "PT Sans", "Avenir Next", "Hiragino Sans",
    # A second variable face, after Noto Sans JP: macOS's STIX Two Text, GPOS only.
    "STIX Two Text",
)
#: Where a family's regular face is not the first one its name finds: Hiragino Sans is
#: ten files of one family, W0 to W9, and PowerPoint draws a regular run in W3.
FILE_HINTS = {"Hiragino Sans": "W3"}
JAPANESE = ("Noto Sans JP", "Yu Gothic", "Meiryo", "Hiragino Sans")
LATIN = ([chr(c) for c in range(0x41, 0x5B)] + [chr(c) for c in range(0x61, 0x7B)] + list(".,-\u2018\u2019\u201c\u201d")
         + [chr(c) for c in range(0xC0, 0x100) if c not in (0xD7, 0xF7)])
KANA = [chr(c) for c in range(0x3041, 0x3097)] + [chr(c) for c in range(0x30A1, 0x30FB)] + ["ー"]
SEPARATORS = "|lIHi!:0158#"
CLASSES = ("legacy", "gpos", "agree", "differ")


def face_file(family: str):
    """``(path, number in its collection, wght or None)`` of the regular face PowerPoint
    uses for ``family`` here (its bundle, then macOS, then the cloud cache)."""
    from pptx2svg.fonts import office

    hint = FILE_HINTS.get(family)
    if hint:
        for _, directory in office.search_dirs():
            for path in sorted(Path(directory).glob(f"*{hint}*")):
                return str(path), 0
    for face in office.find(family):
        if not face.bold and not face.italic:
            return face.path, face.number
    # A variable face names its default instance in name ID 1 ("Noto Sans JP Thin") and
    # the family only in name ID 16, which PowerPoint matches and the index above does not.
    from fontTools.ttLib import TTFont

    for _, directory in office.search_dirs():
        for path in sorted(Path(directory).glob("*.[ot]tf")):
            if "Italic" in path.name:
                continue
            try:
                names = TTFont(str(path), lazy=True)["name"]
            except Exception:
                continue
            if (names.getDebugName(16) or "").lower() == family.lower():
                return str(path), 0
    raise SystemExit(f"{family} is not installed here")


def open_face(family: str):
    """The regular face of ``family`` as fontTools reads it -- a variable face (Noto Sans
    JP) instanced at 400, the instance PowerPoint draws for a regular run."""
    from fontTools.ttLib import TTFont

    path, number = face_file(family)
    font = TTFont(path, fontNumber=number)
    if "fvar" in font:
        from fontTools.varLib import instancer

        font = instancer.instantiateVariableFont(font, {"wght": 400})
    return font


def kern_tables(font):
    """``(legacy, gpos)``: ``(glyph, glyph) -> units`` from the legacy table, and a
    function of two glyphs for the ``GPOS`` ``kern`` feature."""
    sys.path.insert(0, str(HERE))
    import extract_font_metrics as extract

    legacy: dict = {}
    if "kern" in font:
        for subtable in font["kern"].kernTables:
            for pair, value in getattr(subtable, "kernTable", {}).items():
                legacy.setdefault(pair, value)
    groups = extract._pair_subtables(extract._kern_lookups(font))

    def gpos(first: str, second: str) -> int:
        total = 0
        for group in groups:
            for subtable in group:
                value = extract._pair_adjustment(subtable, first, second)
                if value is not None:
                    total += value
                    break
        return total

    return legacy, gpos


def classify(font, chars) -> dict[str, list[tuple[str, str, int, int]]]:
    """Every kerning pair among ``chars``, by class: ``(first, second, legacy, gpos)``."""
    cmap = font.getBestCmap()
    legacy, gpos = kern_tables(font)
    glyphs = {char: cmap[ord(char)] for char in chars if ord(char) in cmap}
    out: dict[str, list] = {name: [] for name in CLASSES}
    for first, a in glyphs.items():
        for second, b in glyphs.items():
            old, new = legacy.get((a, b), 0), gpos(a, b)
            if not old and not new:
                continue
            kind = "gpos" if not old else "legacy" if not new else "agree" if old == new else "differ"
            out[kind].append((first, second, old, new))
    return out


def separator(font, pairs) -> str | None:
    """A character that kerns with nothing in ``pairs`` on either side, in either table."""
    cmap = font.getBestCmap()
    legacy, gpos = kern_tables(font)
    used = {c for first, second, _, _ in pairs for c in (first, second)}
    for candidate in SEPARATORS:
        glyph = cmap.get(ord(candidate))
        if glyph is None:
            continue
        clean = True
        for char in used:
            other = cmap[ord(char)]
            if legacy.get((glyph, other)) or legacy.get((other, glyph)) or gpos(glyph, other) or gpos(other, glyph):
                clean = False
                break
        if clean:
            return candidate
    return None


def spread(pairs, count: int):
    """``count`` pairs: the largest adjustments and the smallest, half each -- a small pair
    (Aptos's ``ss``) is the kind a wrap boundary turns on."""
    ranked = sorted(pairs, key=lambda p: (-max(abs(p[2]), abs(p[3])), p[0], p[1]))
    if len(ranked) <= count:
        return ranked
    small = [p for p in ranked if p[0] == "s" and p[1] == "s"]
    rest = [p for p in ranked if p not in small]
    return small + rest[:count // 2] + rest[-(count - count // 2 - len(small)):]


def probes() -> list[tuple[str, str, str, str]]:
    """``(key, family, script, text)``: ``script`` is ``latin`` or ``ea``."""
    out = []
    for family in FACES:
        font = open_face(family)
        sets = [("latin", LATIN)] + ([("ea", KANA)] if family in JAPANESE else [])
        for script, chars in sets:
            classes = classify(font, chars)
            for kind in CLASSES:
                pool = classes[kind]
                chosen = spread(pool, PAIRS_PER_LINE)
                gap = separator(font, chosen) if chosen else None
                while chosen and gap is None:
                    # Drop the pairs whose glyphs kern with "|" and try again.
                    pool = [p for p in pool if separator(font, [p]) == SEPARATORS[0]]
                    chosen = spread(pool, PAIRS_PER_LINE)
                    gap = separator(font, chosen) if chosen else None
                if gap is None:
                    continue
                text = gap + gap.join(first + second for first, second, _, _ in chosen) + gap
                out.append((f"{family}/{script}/{kind}", family, script, text))
    return out


def box(index: int, key: str, family: str, script: str, text: str) -> str:
    face = escape(family)
    fonts = f"<a:latin typeface='{face}'/><a:ea typeface='{face}'/><a:cs typeface='{face}'/>"
    row = index % ROWS_PER_SLIDE
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{100 + index}' name='{escape(key)}'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        f"<a:xfrm><a:off x='{LEFT * PT}' y='{(TOP + row * ROW) * PT}'/>"
        f"<a:ext cx='{BOX_W * PT}' cy='{BOX_H * PT}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        "<p:txBody><a:bodyPr wrap='none' lIns='0' tIns='0' rIns='0' bIns='0'><a:noAutofit/></a:bodyPr>"
        "<a:lstStyle/>"
        f"<a:p><a:r><a:rPr lang='{'ja-JP' if script == 'ea' else 'en-US'}' sz='{SIZE}' kern='100' dirty='0'>"
        f"{fonts}</a:rPr><a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>"
    )


def slides(found) -> list[str]:
    out = []
    for start in range(0, len(found), ROWS_PER_SLIDE):
        shapes = "".join(box(start + i, *probe) for i, probe in enumerate(found[start:start + ROWS_PER_SLIDE]))
        out.append(
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
            f"<p:sld {NAMESPACES}><p:cSld>"
            '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + shapes + "</p:spTree></p:cSld>"
            "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        )
    return out


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    found = probes()
    count = write_deck(target, SOURCE, slides(found))
    print(f"{len(found)} probes on {count} slide(s) -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

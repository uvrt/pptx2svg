#!/usr/bin/env python3
"""Build the decks that measure *which face* PowerPoint resolves a run to.

Three questions, one probe family (``tools/read_font_resolution_probe.py`` reads the
exports, ``tests/fixtures/font-resolution-probe.json`` records what they said):

**Which face draws a run's Japanese.**  Every box sets the same kana and ideographs,
``かなカナ漢字``, with one thing changed: the run's ``<a:latin>``, the run's ``<a:ea>``
(spelled in English, in Japanese, full-width, as a theme slot), the run's ``lang``, a
table cell instead of a text box -- over themes whose ``<a:ea>`` slot and
``<a:font script="Jpan"/>`` entry are themselves varied, one deck per theme:

``jpan-script``   ``<a:ea typeface=""/>``, ``Jpan`` ＭＳ Ｐゴシック (``sample.pptx``'s own)
``jpan-yu``       ``<a:ea typeface=""/>``, ``Jpan`` 游ゴシック
``ea-full-width`` ``<a:ea typeface="ＭＳ Ｐゴシック"/>``
``ea-english``    ``<a:ea typeface="MS PGothic"/>``
``ea-yu``         ``<a:ea typeface="游ゴシック"/>``
``latin-ea``      ``Jpan`` 游ゴシック, ``<a:ea>`` empty: a Japanese face named only as the
                  Latin one, and ``altLang="ja-JP"``
``bold-sizes``    bold MS Gothic and MS Mincho from 6 to 28 pt: the synthetic bold's advance
``names``         ``sample.pptx``'s theme; the run's ``<a:ea>`` names an installed face by each
                  of its name records in turn (name ID 1 and 16, English and Japanese)

**What PowerPoint draws where no named face has the glyph** -- emoji and symbols in
Calibri, Aptos, Noto Sans JP and Lato (on ``jpan-script``).

**Whether a face the deck embeds beats the same family installed.**  Built over
``real-basic-theme.pptx``, which embeds Lato 1.104 -- a 272-glyph subset, no ``●`` --
while Office's cloud cache holds Lato 2.015, whose advances differ:

``embed-older``   the deck's own Lato 1.104, unchanged
``embed-newer``   the same face claiming version 9.000 (``head.fontRevision``, name 5)
``embed-absent``  the same face renamed ``Latoprobe``, a family nothing installs

The embedded faces are re-wrapped as uncompressed EOT, the shape a ``.fntdata`` part may
take (ECMA-376 leaves the payload to EOT, and EOT's flags allow it); nothing else about
them changes.  They are the deck's own OFL faces, and the decks are throwaway.

Usage::

    python3 tools/make_font_resolution_probe.py ~/          # writes ~/frp-*.pptx
    osascript tools/powerpoint_export_pdf.applescript ~/frp-jpan-script.pptx ~/frp-jpan-script.pdf
    ...
    python3 tools/read_font_resolution_probe.py ~/          # reads ~/frp-*.pdf

A probe deck and its export are throwaway and must be deleted again.
"""

from __future__ import annotations

import io
import re
import struct
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "src"))

from make_feature_sweep import NAMESPACES, PT  # noqa: E402
from make_style_probe import write_deck  # noqa: E402

JAPANESE_SOURCE = ROOT / "tests/fixtures/sample.pptx"
EMBEDDED_SOURCE = ROOT / "tests/fixtures/real-basic-theme.pptx"

KANA = "かなカナ漢字"
LATIN = "Hamburgefonstiv 0123"

#: The grid, in points: three columns of boxes, one line each.
LEFT, TOP, COLUMN, ROW, BOX_W, BOX_H, PER_COLUMN, SIZE = 10, 10, 315, 24, 300, 22, 21, 1400
EMBED_ROW, EMBED_PER_COLUMN = 30, 15

#: The Latin faces a run's Japanese falls back by, from every corner of PANOSE and the
#: OS/2 family class: sans and serif, monospaced, script, display, Office's and macOS's.
LATIN_FACES = [
    "Calibri", "Calibri Light", "Arial", "Aptos", "Aptos Display", "Helvetica Neue", "Verdana",
    "Tahoma", "Segoe UI", "Century Gothic", "Gill Sans MT", "Franklin Gothic Book", "Trebuchet MS",
    "Futura", "Avenir Next", "Optima", "Lato", "Raleway", "Courier New", "Consolas", "Menlo",
    "Times New Roman", "Georgia", "Cambria", "Garamond", "Palatino Linotype", "Book Antiqua",
    "Baskerville", "Didot", "Rockwell", "American Typewriter", "Comic Sans MS", "Impact",
    "Brush Script MT", "NoSuchLatinFace",
]

#: Names a run's ``<a:ea>`` may spell a Japanese face by.
EA_NAMES = [
    "ＭＳ Ｐゴシック", "MS PGothic", "ＭＳ ゴシック", "MS Gothic", "ＭＳ Ｐ明朝", "MS PMincho",
    "ＭＳ 明朝", "MS Mincho", "MS UI Gothic", "游ゴシック", "Yu Gothic", "游ゴシック Light",
    "游明朝", "Yu Mincho", "メイリオ", "Meiryo", "ヒラギノ角ゴシック", "Hiragino Sans",
    "ヒラギノ角ゴ ProN", "Hiragino Kaku Gothic ProN", "ヒラギノ明朝 ProN", "Hiragino Mincho ProN",
    "Noto Sans JP", "Osaka", "Arial Unicode MS", "+mn-ea", "+mj-ea", "NoSuchJapaneseFace",
]

#: Which name records PowerPoint matches a family by: each face's name ID 1 and 16, in
#: English and Japanese, for faces in PowerPoint's bundle and in macOS's folders.
NAME_PROBES = [
    "ヒラギノ角ゴシック W3", "Hiragino Sans W3", "ヒラギノ角ゴシック", "Hiragino Sans",
    "ヒラギノ明朝 ProN W3", "Hiragino Mincho ProN W3", "ヒラギノ明朝 ProN", "Hiragino Mincho ProN",
    "ヒラギノ角ゴ ProN W3", "Hiragino Kaku Gothic ProN W3", "ヒラギノ丸ゴ ProN", "Hiragino Maru Gothic ProN",
    "黒体-繁", "Heiti TC", "游明朝 Demibold", "Yu Mincho Demibold", "游ゴシック Medium",
    "Yu Gothic Medium", "Noto Sans JP Thin", "Noto Sans JP", "HG丸ｺﾞｼｯｸM-PRO", "HGMaruGothicMPRO",
]

SYMBOLS = ["⚡📱🔒", "✅❤😀", "★☆✓✔", "→⇒①②", "●○■□◆", "€™…§", "∑√∞≈", "αβγΩ", "한글", "简体字"]


def probe(key, latin=None, ea=None, lang="en-US", text=KANA, kind="box", bold=False, size=None,
          alt_lang=None):
    return {"key": key, "latin": latin, "ea": ea, "lang": lang, "text": text, "kind": kind,
            "bold": bold, "size": size or SIZE, "alt_lang": alt_lang}


def japanese_probes(full: bool) -> list[dict]:
    out = [probe(f"latin:{face}", latin=face) for face in (LATIN_FACES if full else
           ["Calibri", "Arial", "Aptos", "Courier New", "Times New Roman", "Raleway", "Lato"])]
    out += [probe("theme:none"), probe("theme:+mn-ea", ea="+mn-ea"), probe("theme:+mj-ea", ea="+mj-ea")]
    if full:
        out += [probe(f"ea:{name}", latin="Calibri", ea=name) for name in EA_NAMES]
        out += [probe(f"courier+ea:{name}", latin="Courier New", ea=name)
                for name in ("ＭＳ Ｐゴシック", "MS PGothic", "Noto Sans JP", "+mn-ea", "NoSuchJapaneseFace")]
        out += [probe(f"cross:{latin}/{ea}", latin=latin, ea=ea) for latin, ea in (
            ("Calibri", "Courier New"), ("Courier New", "Calibri"), ("Raleway", "Lato"),
            ("Lato", "Raleway"), ("Arial", "Times New Roman"), ("Times New Roman", "Arial"))]
        out += [probe(f"bold:{face}", latin=face, bold=True) for face in ("Calibri", "Times New Roman")]
    out += [probe(f"lang-ja:{face}", latin=face, lang="ja-JP") for face in ("Calibri", "Courier New")]
    out += [probe("lang-ja:ea:MS PGothic", latin="Calibri", ea="MS PGothic", lang="ja-JP")]
    out += [probe(f"table:{face or 'theme'}", latin=face, kind="table")
            for face in (None, "Calibri", "Courier New", "Raleway")]
    out += [probe(f"table:ea:{name}", ea=name, kind="table") for name in ("ＭＳ Ｐゴシック", "游ゴシック", "+mn-ea")]
    if full:
        out += [probe(f"sym:{face}:{text}", latin=face, ea=face, text=text)
                for face in ("Calibri", "Aptos", "Noto Sans JP", "Lato") for text in SYMBOLS]
    return out


#: Point sizes, in hundredths, at which ``bold-sizes`` sets bold MS Gothic and MS Mincho.
BOLD_SIZES = (600, 750, 800, 900, 1000, 1050, 1100, 1200, 1400, 1450, 1500, 1550, 1600, 1800, 2000, 2400, 2800)


def bold_probes() -> list[dict]:
    """Ten kana in MS Gothic and MS Mincho, bold, at each of :data:`BOLD_SIZES`: what
    PowerPoint's synthetic emboldening adds to an advance."""
    return [probe(f"{face}:{size}:{'bold' if bold else 'regular'}", latin="Calibri", ea=face,
                  text="かなカナかなカナかな", bold=bold, size=size)
            for face in ("MS Gothic", "MS Mincho") for size in BOLD_SIZES for bold in (True, False)
            if bold or size in (750, 2400)]


def latin_probes() -> list[dict]:
    """A Japanese face named as the run's *Latin* face, with no East Asian face named
    (the theme's ``<a:ea>`` is empty), and ``altLang`` in place of ``lang``."""
    out = [probe(f"latin-only:{face}", latin=face) for face in
           ("Noto Sans JP", "ＭＳ Ｐゴシック", "MS Mincho", "游ゴシック", "Yu Mincho", "Meiryo")]
    out += [probe(f"altLang-ja:{face}", latin=face, alt_lang="ja-JP") for face in ("Calibri", "Courier New")]
    return out


def name_probes() -> list[dict]:
    return [probe(f"ea:{name}", latin="Calibri", ea=name) for name in NAME_PROBES]


def embedded_probes(family: str) -> list[dict]:
    out = []
    for bold, italic in ((False, False), (True, False)):
        out.append(probe(f"{family}:{'bold' if bold else 'regular'}", latin=family, ea=family,
                         text=LATIN, bold=bold))
    out.append(probe(f"{family}:bullet", latin=family, ea=family, text="Hamburg ● ○ ■ fonstiv"))
    out.append(probe(f"{family}:kana", latin=family, ea=family, text="Lato かなカナ"))
    return out


def run_xml(spec: dict) -> str:
    attrs = (f"lang='{spec['lang']}'" + (f" altLang='{spec['alt_lang']}'" if spec.get("alt_lang") else "")
             + f" sz='{spec['size']}'" + (" b='1'" if spec["bold"] else ""))
    fonts = ""
    if spec["latin"]:
        fonts += f"<a:latin typeface='{escape(spec['latin'])}'/>"
    if spec["ea"]:
        fonts += f"<a:ea typeface='{escape(spec['ea'])}'/>"
    if spec["latin"]:
        fonts += f"<a:cs typeface='{escape(spec['latin'])}'/>"
    return (f"<a:r><a:rPr {attrs} dirty='0'><a:solidFill><a:srgbClr val='000000'/></a:solidFill>"
            f"{fonts}</a:rPr><a:t>{escape(spec['text'])}</a:t></a:r>")


def place(index: int, per_column: int, row: float) -> tuple[int, float, float]:
    """``(slide, x, y)`` in points of probe ``index``."""
    slide, within = divmod(index, 3 * per_column)
    column, line = divmod(within, per_column)
    return slide, LEFT + column * COLUMN, TOP + line * row


def shape(shape_id: int, spec: dict, x: float, y: float) -> str:
    frame = (f"<a:off x='{int(x * PT)}' y='{int(y * PT)}'/>"
             f"<a:ext cx='{int(BOX_W * PT)}' cy='{int(BOX_H * PT)}'/>")
    body = ("<a:bodyPr wrap='none' lIns='0' tIns='0' rIns='0' bIns='0'/><a:lstStyle/>"
            f"<a:p>{run_xml(spec)}</a:p>")
    if spec["kind"] == "table":
        return (
            f"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='{shape_id}' name='{escape(spec['key'])}'/>"
            "<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp='1'/></p:cNvGraphicFramePr><p:nvPr/>"
            f"</p:nvGraphicFramePr><p:xfrm>{frame}</p:xfrm><a:graphic>"
            "<a:graphicData uri='http://schemas.openxmlformats.org/drawingml/2006/table'>"
            f"<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w='{int(BOX_W * PT)}'/></a:tblGrid>"
            f"<a:tr h='{int(BOX_H * PT)}'><a:tc><a:txBody>{body}</a:txBody>"
            "<a:tcPr marL='0' marR='0' marT='0' marB='0'/></a:tc></a:tr></a:tbl>"
            "</a:graphicData></a:graphic></p:graphicFrame>"
        )
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{shape_id}' name='{escape(spec['key'])}'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        f"<a:xfrm>{frame}</a:xfrm><a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        f"<p:txBody>{body}</p:txBody></p:sp>"
    )


def slides(specs: list[dict], per_column: int, row: float) -> list[str]:
    trees: dict[int, list[str]] = {}
    for index, spec in enumerate(specs):
        slide, x, y = place(index, per_column, row)
        trees.setdefault(slide, []).append(shape(100 + index, spec, x, y))
    return [
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + "".join(trees[slide])
        + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        for slide in sorted(trees)
    ]


# --------------------------------------------------------------------------------------
# Theme and font edits
# --------------------------------------------------------------------------------------


def theme_with(xml: str, ea: str | None = None, jpan: str | None = None) -> str:
    """The theme with both collections' ``<a:ea>`` and/or ``Jpan`` entry set."""
    if ea is not None:
        xml = re.sub(r'(<a:(?:major|minor)Font>\s*<a:latin [^>]*/>\s*)<a:ea typeface="[^"]*"',
                     lambda m: f'{m.group(1)}<a:ea typeface="{escape(ea)}"', xml)
    if jpan is not None:
        xml = re.sub(r'<a:font script="Jpan" typeface="[^"]*"', f'<a:font script="Jpan" typeface="{escape(jpan)}"', xml)
    return xml


def _eot_names(data: bytes) -> tuple[list[str], int]:
    """The four header names of an EOT and the offset just past them."""
    names, offset = [], 80
    for _ in range(4):
        size = struct.unpack_from("<H", data, offset + 2)[0]
        names.append(data[offset + 4:offset + 4 + size].decode("utf-16-le"))
        offset += 4 + size
    return names, offset


def rewrap(fntdata: bytes, font: bytes, family: str | None) -> bytes:
    """``fntdata``'s header around ``font``, uncompressed, its family renamed to ``family``."""
    from ooxml_common.fonts.eot import read_eot_header

    header = read_eot_header(fntdata)
    names, after = _eot_names(fntdata)
    if family is not None:
        names = [names[0].replace("Lato", family), names[1], names[2], names[3].replace("Lato", family)]
    tail = fntdata[after:header.font_data_offset]
    fixed = bytearray(fntdata[:80])
    struct.pack_into("<I", fixed, 12, header.flags & ~0x4)  # not MTX-compressed
    strings = b"".join(struct.pack("<HH", 0, len(n.encode("utf-16-le"))) + n.encode("utf-16-le") for n in names)
    body = bytes(fixed) + strings + tail
    total = len(body) + len(font)
    out = bytearray(body + font)
    struct.pack_into("<II", out, 0, total, len(font))
    return bytes(out)


def edited_face(font: bytes, *, version: float | None = None, family: str | None = None) -> bytes:
    from fontTools.ttLib import TTFont

    face = TTFont(io.BytesIO(font))
    name = face["name"]
    if version is not None:
        face["head"].fontRevision = version
        for record in name.names:
            if record.nameID == 5:
                record.string = f"Version {version:.3f}"
    if family is not None:
        for record in name.names:
            if record.nameID in (1, 3, 4, 6, 16, 18):
                text = record.toUnicode()
                record.string = text.replace("Lato", family)
    out = io.BytesIO()
    face.save(out)
    return out.getvalue()


def rewrite(path: Path, edit) -> None:
    """Rewrite ``path`` in place, each part through ``edit(name, data) -> data``."""
    source = zipfile.ZipFile(io.BytesIO(path.read_bytes()))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for item in source.infolist():
            out.writestr(item, edit(item.filename, source.read(item.filename)))


# --------------------------------------------------------------------------------------
# The decks
# --------------------------------------------------------------------------------------

#: ``deck name -> (source, theme edit or None, embedded-font edit or None, full probe set)``
JAPANESE_DECKS = {
    "jpan-script": ({}, True),
    "jpan-yu": ({"jpan": "游ゴシック"}, False),
    "ea-full-width": ({"ea": "ＭＳ Ｐゴシック"}, False),
    "ea-english": ({"ea": "MS PGothic"}, False),
    "ea-yu": ({"ea": "游ゴシック"}, False),
    "names": ({}, None),
    "bold-sizes": ({}, None),
    "latin-ea": ({"jpan": "游ゴシック"}, None),
}
EMBEDDED_DECKS = {
    "embed-older": {},
    "embed-newer": {"version": 9.0},
    "embed-absent": {"family": "Latoprobe"},
}


def deck_probes(deck: str) -> list[dict]:
    if deck == "names":
        return name_probes()
    if deck == "bold-sizes":
        return bold_probes()
    if deck == "latin-ea":
        return latin_probes()
    if deck in JAPANESE_DECKS:
        return japanese_probes(JAPANESE_DECKS[deck][1])
    family = EMBEDDED_DECKS[deck].get("family", "Lato")
    return embedded_probes(family) + embedded_probes("Raleway")


def geometry(deck: str) -> tuple[int, float]:
    tall = deck in EMBEDDED_DECKS or deck == "bold-sizes"
    return (EMBED_PER_COLUMN, EMBED_ROW) if tall else (PER_COLUMN, ROW)


def build(deck: str, target: Path) -> int:
    specs = deck_probes(deck)
    per_column, row = geometry(deck)
    if deck in JAPANESE_DECKS:
        count = write_deck(target, JAPANESE_SOURCE, slides(specs, per_column, row))
        theme = JAPANESE_DECKS[deck][0]
        if theme:
            rewrite(target, lambda name, data: theme_with(data.decode(), **theme).encode()
                    if re.match(r"ppt/theme/theme\d+\.xml$", name) else data)
        return count

    # The source deck whole, its slides' shapes replaced: Google Slides' masters do not
    # survive being emptied the way make_style_probe.write_deck empties them (PowerPoint
    # hangs opening the result), and nothing on them draws over the probes.
    shapes = "".join(shape(100 + index, spec, *place(index, per_column, row)[1:])
                     for index, spec in enumerate(specs))
    target.write_bytes(EMBEDDED_SOURCE.read_bytes())
    tree = re.compile(r"(</p:grpSpPr>).*(</p:spTree>)", re.S)
    rewrite(target, lambda name, data: tree.sub(
        lambda m: m.group(1) + (shapes if name.endswith("slide1.xml") else "") + m.group(2),
        data.decode()).encode() if re.match(r"ppt/slides/slide\d+\.xml$", name) else data)
    count = 2
    edit = EMBEDDED_DECKS[deck]
    if edit:
        from ooxml_common.fonts.eot import decode_eot

        def change(name: str, data: bytes) -> bytes:
            if re.match(r"ppt/fonts/Lato-\w+\.fntdata$", name):
                return rewrap(data, edited_face(decode_eot(data), **edit), edit.get("family"))
            if name == "ppt/presentation.xml" and "family" in edit:
                return data.replace(b'<p:font typeface="Lato"/>',
                                    f'<p:font typeface="{edit["family"]}"/>'.encode())
            return data

        rewrite(target, change)
    return count


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    directory = Path(sys.argv[1]).expanduser()
    for deck in [*JAPANESE_DECKS, *EMBEDDED_DECKS]:
        target = directory / f"frp-{deck}.pptx"
        count = build(deck, target)
        print(f"{deck:14} {len(deck_probes(deck)):3} probes on {count} slide(s) -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

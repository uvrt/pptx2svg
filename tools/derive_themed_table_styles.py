#!/usr/bin/env python3
"""Measure what PowerPoint's "Themed Style 1" and "Themed Style 2" table styles draw.

Every other built-in style is a fixed set of colours, and ``tools/derive_table_styles.py``
reads those off PowerPoint's export by matching them against a swatch sheet.  These two
families are not: they draw on the theme's *format scheme* -- a fill style, a line style,
an effect style, chosen by index, with the placeholder colour ``phClr`` replaced by a
colour of the style's own -- so a colour measured under one theme is wrong under the next.
What has to be measured is which index each part of the table uses and what colour it
puts in ``phClr``.  This tool does that by drawing every style of both families under
five themes built to make the answer visible:

``index``
    Each fill, line and effect style is a fixed colour no scheme colour comes near:
    fill styles 1 to 3 olive, magenta and spring green; line styles 1 to 3 brown 2 pt,
    pink 4 pt and navy 6 pt; effect styles 1 to 3 a hard 6 pt shadow, red to the right,
    green downwards and blue to the left, each over a bevel.  A colour or a width seen in
    the render names the index outright.
``phclr``
    Every entry is ``phClr`` (line widths still 2, 4 and 6 pt), so what is drawn is the
    colour the reference supplies, named against the swatch sheet on the deck's last slide.
``white`` / ``black``
    Every fill style white, then black, lines as ``index``, no effects: a cell's colour
    over the two gives its own colour and its alpha (``observed = colour * alpha + backdrop
    * (1 - alpha)``), and a cell that tracks the backdrop has no fill of its own.
``gradient``
    Every fill style a black-to-white vertical gradient: whether the background runs once
    across the table or once per cell.

The theme's fonts are Georgia (major) and Courier New (minor) and the master's other-text
style names Arial, so a cell drawn in Courier New takes its face from the style's
``fontRef``.

Each style gets a slide of six 4 x 5 tables, one per flag set, that between them expose
every region: ``A`` header, total row, first and last column, banded rows; ``D`` the same
with banded columns; ``E`` header, total row and banded rows; ``B`` banded columns; ``C``
no flags; ``F`` first and last column and banded columns.

Usage::

    python3 tools/derive_themed_table_styles.py make DIR      # five decks into DIR
    # export DIR/themed-<theme>.pptx to DIR/themed-<theme>.pdf with PowerPoint, e.g.
    #   osascript tools/powerpoint_export_pdf.applescript IN.pptx OUT.pdf
    python3 tools/derive_themed_table_styles.py read DIR      # observations, as text
    python3 tools/derive_themed_table_styles.py check DIR     # pptx2svg against them
    # (check renders into DIR/ours once; delete that folder to render again)

``read`` prints, per style and table, every cell's fill, text and four edges in the terms
above (``L2`` is line style 2, ``F3`` fill style 3; a colour is its swatch name), the
shadow (which effect style, in which colour), the background down the rows of the
``gradient`` theme and how much the rim differs from the inside (a bevel).  ``check``
renders the same decks with pptx2svg and reports every cell fill, edge, shadow and text
colour, weight and face that differs from PowerPoint's.

What it found, which ``src/pptx2svg/parse/table_styles_builtin.py`` carries as
``_THEMED_STYLE_XML``: see the comment there.  The background is one gradient down the
whole table, not one per cell; no bevel is drawn; all six accents of each family measure
alike but for the accent.

Needs ``pypdfium2`` and ``pillow``; neither is a dependency of the library.  The decks and
exports are throwaway and must not be committed.
"""

from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from derive_table_styles import swatch_expressions  # noqa: E402
from make_style_probe import write_deck  # noqa: E402

SOURCE = HERE.parent / "tests/fixtures/real-financial-report.pptx"
PT = 12700
SCALE = 8  # pixels per point when reading an export

STYLES = [
    ("{3C2FFA5D-87B4-456A-9821-1D502468CF0F}", "Themed Style 1 - Accent 1"),
    ("{284E427A-3D55-4303-BF80-6455036E1DE7}", "Themed Style 1 - Accent 2"),
    ("{69C7853C-536D-4A76-A0AE-DD22124D55A5}", "Themed Style 1 - Accent 3"),
    ("{775DCB02-9BB8-47FD-8907-85C794F793BA}", "Themed Style 1 - Accent 4"),
    ("{35758FB7-9AC5-4552-8A53-C91805E547FA}", "Themed Style 1 - Accent 5"),
    ("{08FB837D-C827-4EFA-A057-4D05807E0F7C}", "Themed Style 1 - Accent 6"),
    ("{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}", "Themed Style 2 - Accent 1"),
    ("{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}", "Themed Style 2 - Accent 2"),
    ("{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}", "Themed Style 2 - Accent 3"),
    ("{E269D01E-BC32-4049-B463-5C60D7B0CCD2}", "Themed Style 2 - Accent 4"),
    ("{327F97BB-C833-4FB7-BDE5-3F7075034690}", "Themed Style 2 - Accent 5"),
    ("{638B1855-1B75-4FBE-930C-398BA8C253C6}", "Themed Style 2 - Accent 6"),
]

# -- The themes ----------------------------------------------------------------------------

SCHEME = {"dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "EEECE1",
          "accent1": "1F77B4", "accent2": "D62728", "accent3": "2CA02C",
          "accent4": "9467BD", "accent5": "FF7F0E", "accent6": "17BECF",
          "hlink": "0563C1", "folHlink": "954F72"}
FILL_COLOURS = ("C8B400", "FF00FF", "00FF80")
LINE_COLOURS = ("804000", "FF66B2", "000080")
LINE_WIDTHS_PT = (2, 4, 6)
SHADOW_COLOURS = ("FF0000", "00C000", "0000FF")
SHADOW_DIRECTIONS = (0, 5400000, 10800000)  # right, down, left
SHADOW_PT = 6
THEMES = ("index", "phclr", "white", "black", "gradient")


def _colour(theme: str, fixed: str) -> str:
    return ('<a:schemeClr val="phClr"/>' if theme == "phclr"
            else f'<a:srgbClr val="{fixed}"/>')


def _fill_style(theme: str, k: int) -> str:
    if theme == "gradient":
        return ('<a:gradFill rotWithShape="1"><a:gsLst>'
                '<a:gs pos="0"><a:srgbClr val="000000"/></a:gs>'
                '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst>'
                '<a:lin ang="5400000" scaled="0"/></a:gradFill>')
    fixed = {"white": "FFFFFF", "black": "000000"}.get(theme, FILL_COLOURS[k])
    return f"<a:solidFill>{_colour(theme, fixed)}</a:solidFill>"


def _line_style(theme: str, k: int) -> str:
    colour = _colour(theme, LINE_COLOURS[k])
    return (f'<a:ln w="{LINE_WIDTHS_PT[k] * PT}" cap="flat" cmpd="sng" algn="ctr">'
            f'<a:solidFill>{colour}</a:solidFill><a:prstDash val="solid"/></a:ln>')


def _effect_style(theme: str, k: int) -> str:
    if theme not in ("index", "phclr"):
        return "<a:effectStyle><a:effectLst/></a:effectStyle>"
    colour = _colour(theme, SHADOW_COLOURS[k])
    bevel = ('<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" '
             'dir="t"/></a:scene3d><a:sp3d><a:bevelT w="152400" h="76200"/></a:sp3d>'
             if theme == "index" else "")
    return (f'<a:effectStyle><a:effectLst><a:outerShdw blurRad="0" dist="{SHADOW_PT * PT}" '
            f'dir="{SHADOW_DIRECTIONS[k]}" algn="ctr" rotWithShape="0">{colour}'
            f"</a:outerShdw></a:effectLst>{bevel}</a:effectStyle>")


def theme_parts(theme: str) -> tuple[str, str, str]:
    """clrScheme, fontScheme and fmtScheme for ``theme``."""
    colours = "".join(f'<a:{slot}><a:srgbClr val="{value}"/></a:{slot}>'
                      for slot, value in SCHEME.items())
    fonts = ('<a:fontScheme name="Probe"><a:majorFont><a:latin typeface="Georgia"/>'
             '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont>'
             '<a:latin typeface="Courier New"/><a:ea typeface=""/><a:cs typeface=""/>'
             "</a:minorFont></a:fontScheme>")
    fills = "".join(_fill_style(theme, k) for k in range(3))
    fmt = ('<a:fmtScheme name="Probe">'
           f"<a:fillStyleLst>{fills}</a:fillStyleLst>"
           f"<a:lnStyleLst>{''.join(_line_style(theme, k) for k in range(3))}</a:lnStyleLst>"
           "<a:effectStyleLst>" + "".join(_effect_style(theme, k) for k in range(3))
           + f"</a:effectStyleLst><a:bgFillStyleLst>{fills}</a:bgFillStyleLst></a:fmtScheme>")
    return f'<a:clrScheme name="Probe">{colours}</a:clrScheme>', fonts, fmt


# -- The slides ----------------------------------------------------------------------------

COLS, ROWS = 4, 5
CELL_W, CELL_H = 70, 30  # pt
TABLES = [  # name, x, y (pt), flags
    ("A", 30, 50, ("firstRow", "lastRow", "firstCol", "lastCol", "bandRow")),
    ("D", 340, 50, ("firstRow", "lastRow", "firstCol", "lastCol", "bandCol")),
    ("E", 650, 50, ("firstRow", "lastRow", "bandRow")),
    ("B", 30, 300, ("bandCol",)),
    ("C", 340, 300, ()),
    ("F", 650, 300, ("firstCol", "lastCol", "bandCol")),
]
NAMESPACES = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
              'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')


def _table(shape_id: int, guid: str, name: str, x: int, y: int, flags) -> str:
    attributes = "".join(f' {flag}="1"' for flag in flags)
    grid = "".join(f'<a:gridCol w="{CELL_W * PT}"/>' for _ in range(COLS))
    rows = "".join(
        f'<a:tr h="{CELL_H * PT}">' + "".join(
            '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1000"/>'
            f"<a:t>H{r}{c}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>" for c in range(COLS))
        + "</a:tr>" for r in range(ROWS))
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{shape_id}" name="{name}"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
            f'<p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x * PT}" y="{y * PT}"/>'
            f'<a:ext cx="{COLS * CELL_W * PT}" cy="{ROWS * CELL_H * PT}"/></p:xfrm><a:graphic>'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f"<a:tbl><a:tblPr{attributes}><a:tableStyleId>{guid}</a:tableStyleId></a:tblPr>"
            f"<a:tblGrid>{grid}</a:tblGrid>{rows}</a:tbl></a:graphicData></a:graphic>"
            "</p:graphicFrame>")


def _slide(body: str) -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"<p:sld {NAMESPACES}><p:cSld>"
            '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/>'
            '</p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/>'
            '<p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" '
            'cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + body + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
            "</p:sld>")


SWATCH_PT, SWATCH_COLUMNS = 16, 40


def _swatches() -> tuple[str, dict[str, tuple[float, float]]]:
    shapes, centres = [], {}
    for i, (name, slot, body) in enumerate(swatch_expressions()):
        x = 20 + (i % SWATCH_COLUMNS) * SWATCH_PT
        y = 20 + (i // SWATCH_COLUMNS) * SWATCH_PT
        centres[name] = (x + SWATCH_PT / 2, y + SWATCH_PT / 2)
        shapes.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{200 + i}" name="s{i}"/><p:cNvSpPr/><p:nvPr/>'
            f'</p:nvSpPr><p:spPr><a:xfrm><a:off x="{x * PT}" y="{y * PT}"/><a:ext '
            f'cx="{SWATCH_PT * PT}" cy="{SWATCH_PT * PT}"/></a:xfrm><a:prstGeom prst="rect">'
            f'<a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="{slot}">{body}'
            "</a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>")
    return _slide("".join(shapes)), centres


def make(out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    slides = [_slide("".join(_table(10 + i, guid, f"{name} / {t}", x, y, flags)
                             for i, (t, x, y, flags) in enumerate(TABLES)))
              for guid, name in STYLES]
    swatch_slide, centres = _swatches()
    slides.append(swatch_slide)
    (out / "swatches.json").write_text(json.dumps(centres))
    for theme in THEMES:
        target = out / f"themed-{theme}.pptx"
        write_deck(target, SOURCE, slides)
        _retheme(target, theme)
        print("wrote", target)


def _retheme(deck: Path, theme: str) -> None:
    colours, fonts, fmt = theme_parts(theme)
    original = zipfile.ZipFile(io.BytesIO(deck.read_bytes()))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for item in original.infolist():
            data = original.read(item.filename)
            if re.fullmatch(r"ppt/theme/theme\d+\.xml", item.filename):
                text = data.decode()
                text = re.sub(r"<a:clrScheme .*?</a:clrScheme>", colours, text, flags=re.S)
                text = re.sub(r"<a:fontScheme .*?</a:fontScheme>", fonts, text, flags=re.S)
                text = re.sub(r"<a:fmtScheme .*?</a:fmtScheme>", fmt, text, flags=re.S)
                data = text.encode()
            elif re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", item.filename):
                text = data.decode()
                start = text.find("<p:otherStyle>")
                end = text.find("</p:otherStyle>")
                if start >= 0:
                    other = text[start:end].replace('typeface="+mn-lt"', 'typeface="Arial"')
                    text = text[:start] + other + text[end:]
                data = text.encode()
            out.writestr(item, data)
    deck.write_bytes(buffer.getvalue())


# -- Reading an export ---------------------------------------------------------------------


def _hex(rgb) -> str:
    return "#%02x%02x%02x" % tuple(rgb[:3])


def _distance(a, b) -> int:
    return max(abs(p - q) for p, q in zip(a[:3], b[:3]))


class Page:
    """One rendered page, addressed in points."""

    def __init__(self, image, text=None):
        self.image, self.text = image, text or []

    def at(self, x: float, y: float):
        return self.image.getpixel((int(round(x * SCALE)), int(round(y * SCALE))))[:3]


def pdf_pages(pdf: Path) -> list[Page]:
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw
    import ctypes

    pages = []
    for page in pdfium.PdfDocument(str(pdf)):
        image = page.render(scale=SCALE).to_pil().convert("RGB")
        height = page.get_height()
        textpage = page.get_textpage()
        chars = []
        for index in range(raw.FPDFText_CountChars(textpage)):
            left, right, bottom, top = (ctypes.c_double() for _ in range(4))
            raw.FPDFText_GetCharBox(textpage, index, left, right, bottom, top)
            r, g, b, a = (ctypes.c_uint() for _ in range(4))
            raw.FPDFText_GetFillColor(textpage, index, r, g, b, a)
            buffer = ctypes.create_string_buffer(256)
            flags = ctypes.c_int()
            raw.FPDFText_GetFontInfo(textpage, index, buffer, 256, flags)
            chars.append({"x": (left.value + right.value) / 2,
                          "y": height - (bottom.value + top.value) / 2,
                          "colour": (r.value, g.value, b.value),
                          "font": buffer.value.decode("latin-1"),
                          "weight": raw.FPDFText_GetFontWeight(textpage, index)})
        pages.append(Page(image, chars))
    return pages


def png_pages(folder: Path, stem: str) -> list[Page]:
    from PIL import Image

    files = sorted(folder.glob(f"{stem}-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))
    return [Page(Image.open(f).convert("RGB")) for f in files]


class Namer:
    """A sampled colour's name, from the swatch slide of the same export."""

    def __init__(self, page: Page, centres: dict):
        self.table: dict[str, list[str]] = {}
        for name, (x, y) in centres.items():
            self.table.setdefault(_hex(page.at(x, y)), []).append(name)
        for names in self.table.values():
            names.sort(key=lambda n: (n.split(":")[1] != "plain", len(n)))
        self.colours = [(tuple(int(h[i:i + 2], 16) for i in (1, 3, 5)), names[0])
                        for h, names in self.table.items()]

    def __call__(self, rgb, tolerance: int = 3) -> str:
        best = min(self.colours, key=lambda item: _distance(item[0], rgb))
        if _distance(best[0], rgb) <= tolerance:
            return best[1].replace(":plain", "")
        return _hex(rgb)


def _index_of(rgb, palette, tolerance=24):
    for k, value in enumerate(palette):
        if _distance(rgb, tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))) <= tolerance:
            return k + 1
    return None


def _table_box(x: float, y: float):
    return x, y, x + COLS * CELL_W, y + ROWS * CELL_H


def _cell_sample(x0, y0, r, c):
    """A point inside the cell that no text and no line reaches."""
    return x0 + c * CELL_W + CELL_W - 14, y0 + r * CELL_H + CELL_H / 2


def _edge(pages: dict, page_no: int, x: float, y: float, horizontal: bool):
    """The line on an edge: in ``index`` its line style (or its own colour) and width; its
    colour in ``phclr``.  ``None`` when nothing is drawn.

    A line style is recognised by its colour, so a shadow or a fill beside an outer edge is
    not mistaken for it; any other line by differing from the cells on both sides within
    3.5 pt of the edge (half the widest line style and a little)."""
    page = pages["index"][page_no]

    def profile(reach):
        steps = int(2 * reach * SCALE) + 1
        points = [(x, y - reach + k / SCALE) if horizontal else (x - reach + k / SCALE, y)
                  for k in range(steps)]
        return points, [page.at(*point) for point in points]

    points, colours = profile(4.0)
    palette = [tuple(int(v[i:i + 2], 16) for i in (0, 2, 4)) for v in LINE_COLOURS]
    best = None
    for k, ink in enumerate(palette):
        run = [n for n, rgb in enumerate(colours) if _distance(rgb, ink) <= 40]
        if run and (best is None or len(run) > len(best[1])):
            best = (k + 1, run)
    if best is not None:
        style, run = best
    else:
        points, colours = profile(3.5)
        before, after = colours[0], colours[-1]
        shadows = [tuple(int(v[i:i + 2], 16) for i in (0, 2, 4)) for v in SHADOW_COLOURS]
        run = [n for n, rgb in enumerate(colours)
               if _distance(rgb, before) > 40 and _distance(rgb, after) > 40
               and all(_distance(rgb, s) > 40 for s in shadows)]
        if not run:
            return None
        style = None
    middle = run[len(run) // 2]
    width = round(len(run) / SCALE * 4) / 4
    phclr = pages["phclr"][page_no].at(*points[middle])
    return {"style": style, "width": width, "index_colour": _hex(colours[middle]),
            "phclr": phclr}


def _bevel(page: Page, x0: float, y0: float) -> int:
    """How much the background just inside the table's top edge differs from the same
    background further in: a bevel shades the rim, a flat fill does not."""
    rim = page.at(x0 + CELL_W * 1.5, y0 + 4)
    inner = page.at(x0 + CELL_W * 1.5, y0 + CELL_H * 0.5 + 6)
    return _distance(rim, inner)


def read(folder: Path, source: str = "pdf") -> dict:
    """Observations for every style, from PowerPoint's PDFs (``source="pdf"``) or from
    pptx2svg's PNGs of the same decks (``source="png"``, rendered at 8 px/pt)."""
    centres = json.loads((folder / "swatches.json").read_text())
    pages = {theme: (pdf_pages(folder / f"themed-{theme}.pdf") if source == "pdf"
                     else png_pages(folder / "ours", f"themed-{theme}"))
             for theme in THEMES}
    namer = Namer(pages["phclr"][-1], centres)
    result = {}
    for page_no, (guid, name) in enumerate(STYLES):
        style = {"guid": guid, "tables": {}}
        for t, x0, y0, flags in TABLES:
            cells, h_edges, v_edges = {}, {}, {}
            for r in range(ROWS):
                for c in range(COLS):
                    sx, sy = _cell_sample(x0, y0, r, c)
                    white = pages["white"][page_no].at(sx, sy)
                    black = pages["black"][page_no].at(sx, sy)
                    alpha = 1 - sum(w - b for w, b in zip(white, black)) / (3 * 255)
                    if alpha < 0.03:
                        fill = None
                    else:
                        colour = tuple(min(255, round(b / alpha)) for b in black)
                        fill = namer(colour) + ("" if alpha > 0.97 else f"/alpha{round(alpha * 100)}")
                    index_rgb = pages["index"][page_no].at(sx, sy)
                    cells[f"{r},{c}"] = {"fill": fill,
                                         "index": _index_of(index_rgb, FILL_COLOURS),
                                         "phclr": namer(pages["phclr"][page_no].at(sx, sy))}
            for r in range(ROWS + 1):
                for c in range(COLS):
                    h_edges[f"{r},{c}"] = _edge(pages, page_no, x0 + c * CELL_W + CELL_W * 0.7,
                                                y0 + r * CELL_H, True)
            for c in range(COLS + 1):
                for r in range(ROWS):
                    v_edges[f"{r},{c}"] = _edge(pages, page_no, x0 + c * CELL_W,
                                                y0 + r * CELL_H + CELL_H * 0.75, False)
            for edge in list(h_edges.values()) + list(v_edges.values()):
                if edge is not None:
                    edge["phclr"] = namer(edge["phclr"])
            left, top, right, bottom = _table_box(x0, y0)
            mx, my = (left + right) / 2, (top + bottom) / 2
            shadow = {}
            for side, (sx, sy) in {"right": (right + 3, my), "below": (mx, bottom + 3),
                                   "left": (left - 3, my), "above": (mx, top - 3)}.items():
                rgb = pages["index"][page_no].at(sx, sy)
                shadow[side] = {"index": _index_of(rgb, SHADOW_COLOURS, 60),
                                "phclr": namer(pages["phclr"][page_no].at(sx, sy))}
            bevel = _bevel(pages["index"][page_no], x0, y0)
            gradient = [round(sum(pages["gradient"][page_no].at(*_cell_sample(x0, y0, r, 1))) / 3)
                        for r in range(ROWS)]
            text = {}
            for char in pages["index"][page_no].text:
                c = int((char["x"] - x0) // CELL_W)
                r = int((char["y"] - y0) // CELL_H)
                if 0 <= r < ROWS and 0 <= c < COLS and f"{r},{c}" not in text:
                    text[f"{r},{c}"] = {"colour": namer(char["colour"], 1),
                                        "font": char["font"], "weight": char["weight"]}
            style["tables"][t] = {"flags": flags, "cells": cells, "h": h_edges, "v": v_edges,
                                  "shadow": shadow, "gradient": gradient, "text": text,
                                  "bevel": bevel}
        result[name] = style
    return result


# -- Printing ------------------------------------------------------------------------------


def _edge_text(edge) -> str:
    if edge is None:
        return "-"
    if edge["style"]:
        return f"L{edge['style']}:{edge['phclr']}"
    return f"{edge['index_colour']}@{edge['width']}"


def report(observations: dict) -> str:
    lines = []
    for name, style in observations.items():
        lines.append(f"== {name} {style['guid']}")
        for t, table in style["tables"].items():
            lines.append(f"  table {t} {' '.join(table['flags']) or '(no flags)'}")
            shadow = {side: (v["index"], v["phclr"]) for side, v in table["shadow"].items()}
            lines.append(f"    shadow {shadow}  gradient-by-row {table['gradient']}  "
                         f"rim-vs-inner {table['bevel']}")
            for r in range(ROWS + 1):
                lines.append("    h " + " | ".join(_edge_text(table["h"][f"{r},{c}"])
                                                 for c in range(COLS)))
                if r == ROWS:
                    break
                lines.append("    v " + " | ".join(_edge_text(table["v"][f"{r},{c}"])
                                                 for c in range(COLS + 1)))
                cells = []
                for c in range(COLS):
                    cell = table["cells"][f"{r},{c}"]
                    text = table["text"].get(f"{r},{c}", {})
                    bold = "B" if text.get("weight", 400) >= 600 or "Bold" in text.get("font", "") else ""
                    cells.append(f"{cell['fill'] or 'none'}[F{cell['index'] or '-'}/{cell['phclr']}]"
                                 f" {text.get('colour', '?')}{bold} {text.get('font', '?')[:14]}")
                lines.append("    c " + " | ".join(cells))
    return "\n".join(lines)


def check(folder: Path) -> int:
    """Render the five decks with pptx2svg at the export's scale and compare."""
    from pptx2svg.cli import main as cli

    ours = folder / "ours"
    for theme in THEMES:
        if (ours / f"themed-{theme}-1.png").exists():
            continue  # rendered by an earlier run; delete DIR/ours to render again
        cli([str(folder / f"themed-{theme}.pptx"), "-o", str(ours), "-f", "png", "-q",
             "--width", str(960 * SCALE)])

    def colour_same(left, right) -> bool:
        """Two colour names, or two unnamed colours (a blend over the background) within
        three levels of each other -- the rasterisers' rounding."""
        if left == right:
            return True
        if not (str(left).startswith("#") and str(right).startswith("#")):
            return False
        rgb = [tuple(int(v[i:i + 2], 16) for i in (1, 3, 5)) for v in (left, right)]
        return _distance(*rgb) <= 3
    theirs, mine = read(folder, "pdf"), read(folder, "png")
    differences = 0
    for name in theirs:
        for t in theirs[name]["tables"]:
            a, b = theirs[name]["tables"][t], mine[name]["tables"][t]
            for kind in ("cells", "h", "v"):
                for key in a[kind]:
                    left, right = a[kind][key], b[kind][key]
                    if kind == "cells":
                        same = ((left["fill"], left["index"]) == (right["fill"], right["index"])
                                and colour_same(left["phclr"], right["phclr"]))
                    else:
                        same = (left is None) == (right is None) and (left is None or (
                            left["style"] == right["style"]
                            and colour_same(left["phclr"], right["phclr"])
                            and abs(left["width"] - right["width"]) <= 0.5))
                    if not same:
                        differences += 1
                        print(f"{name} {t} {kind} {key}: PowerPoint {left} pptx2svg {right}")
            for side in a["shadow"]:
                if a["shadow"][side]["index"] != b["shadow"][side]["index"]:
                    differences += 1
                    print(f"{name} {t} shadow {side}: PowerPoint {a['shadow'][side]} "
                          f"pptx2svg {b['shadow'][side]}")
    differences += _check_text(folder, theirs)
    print(f"{differences} differences")
    return 1 if differences else 0


def _check_text(folder: Path, theirs: dict) -> int:
    """Each cell's text colour, weight and face in pptx2svg's model against the export."""
    from pptx2svg import ConvertOptions, convert_pptx_to_model, model as m

    centres = json.loads((folder / "swatches.json").read_text())
    namer = Namer(pdf_pages(folder / "themed-phclr.pdf")[-1], centres)
    model = convert_pptx_to_model(str(folder / "themed-index.pptx"), ConvertOptions())
    differences = 0
    for page_no, (_, name) in enumerate(STYLES):
        tables = [e for e in model.slides[page_no].elements if isinstance(e, m.TableElement)]
        for (t, *_), table in zip(TABLES, tables):
            for r, row in enumerate(table.table.rows):
                for c, cell in enumerate(row.cells):
                    run = cell.text_body.paragraphs[0].runs[0].properties
                    seen = theirs[name]["tables"][t]["text"][f"{r},{c}"]
                    rgb = tuple(int(run.color.hex[i:i + 2], 16) for i in (1, 3, 5))
                    bold = seen["weight"] >= 600 or "Bold" in seen["font"] or "-B" in seen["font"]
                    face = seen["font"].split("PS")[0].split("-")[0]
                    ours = (namer(rgb, 1), bool(run.bold), run.font_family.replace(" ", ""))
                    if ours != (seen["colour"], bold, face):
                        differences += 1
                        print(f"{name} {t} text {r},{c}: PowerPoint {(seen['colour'], bold, face)} "
                              f"pptx2svg {ours}")
    return differences


def main(argv: list[str]) -> int:
    if len(argv) != 2 or argv[0] not in ("make", "read", "check"):
        print(__doc__)
        return 2
    folder = Path(argv[1]).expanduser()
    if argv[0] == "make":
        make(folder)
        return 0
    if argv[0] == "read":
        observations = read(folder)
        (folder / "observations.json").write_text(json.dumps(observations, indent=1))
        print(report(observations))
        return 0
    return check(folder)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

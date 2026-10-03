#!/usr/bin/env python3
"""Build decks that measure what PowerPoint draws for a shape's theme style references.

A shape PowerPoint inserts with default formatting states no fill, no line and no text
colour of its own.  It carries ``p:style`` instead::

    <p:style>
      <a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>
      <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
      <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
      <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
    </p:style>

and each reference names an entry of the theme's ``a:fmtScheme`` whose colours are all
``a:schemeClr val="phClr"`` -- the placeholder the reference's own colour stands in for.
This deck puts one shape per question on the page, each named after its probe key, so
``tools/read_style_probe.py`` can find it in the export and in pptx2svg's model.

The probes are drawn twice, on two themes, because the two Office themes in this corpus
disagree on everything that matters here: ``style-2013`` takes the theme of
``real-financial-report.pptx`` (gradient fill styles of ``lumMod``/``satMod``/``tint``
stops, plain ``phClr`` lines, a shadow only on effect style 3) and ``style-2007`` the theme
of ``sample.pptx`` (gradient stops of ``tint``/``satMod``, a ``shade``/``satMod`` on line
style 1, a shadow on every effect style and a bevel on the third).

``fill-*``
    ``a:fillRef`` 0 to 3 and 1000 to 1003, so both lists and both "no fill" indices.
``mod-*``
    Colour modifiers on the reference -- ``shade``, ``lumMod``/``lumOff``, ``tint``,
    ``satMod``, ``alpha`` -- under a solid and a gradient fill style, which separates "the
    reference's colour replaces the stop" from "it replaces ``phClr`` and the stop's own
    transforms then apply".
``font-*``
    ``a:fontRef``: its colour and face, with and without a colour, against a run colour
    and a ``lstStyle`` colour, which say where it sits in the text cascade.
``line-*``
    ``a:lnRef`` 0 to 3, the regression control, and a local ``a:ln`` over it.
``effect-*``
    ``a:effectRef`` 0 to 3: which index is the first effect style.
``own-*``
    An explicit ``a:spPr`` fill or line over the reference.

Usage -- the deck's file name picks its theme::

    python3 tools/make_style_probe.py ~/Documents/style-2013.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/Documents/style-2013.pptx ~/Documents/style-2013.pdf
    python3 tools/read_style_probe.py ~/Documents/style-2013.pptx ~/Documents/style-2013.pdf

A probe deck and its export are throwaway and must be deleted again.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from make_feature_sweep import (  # noqa: E402
    NAMESPACES,
    PT,
    SLIDE_REL,
    SLIDE_TYPE,
    ZIP_TIMESTAMP,
    blank_template,
)

#: Which fixture lends each deck its master, layouts and theme.
SOURCES = {
    "style-2013": ROOT / "tests/fixtures/real-financial-report.pptx",
    "style-2007": ROOT / "tests/fixtures/sample.pptx",
}

#: Both sources are 960 x 540 pt.
SLIDE_WIDTH, SLIDE_HEIGHT = 960, 540
BOX_W, BOX_H = 150, 64
MARGIN, GAP_X, GAP_Y, CAPTION = 24, 30, 24, 14
COLUMNS = 5


def clr(value: str, **mods: int) -> str:
    """A colour element: ``accent1`` and friends are scheme colours, six hex digits sRGB."""
    children = "".join(f"<a:{name} val='{val}'/>" for name, val in mods.items())
    tag = "srgbClr" if re.fullmatch(r"[0-9A-F]{6}", value) else "schemeClr"
    return f"<a:{tag} val='{value}'>{children}</a:{tag}>"


def style(
    *,
    ln: tuple[int, str] = (0, clr("accent1")),
    fill: tuple[int, str] = (1, clr("accent1")),
    effect: tuple[int, str] = (0, clr("accent1")),
    font: tuple[str, str] = ("minor", clr("lt1")),
) -> str:
    return (
        "<p:style>"
        f"<a:lnRef idx='{ln[0]}'>{ln[1]}</a:lnRef>"
        f"<a:fillRef idx='{fill[0]}'>{fill[1]}</a:fillRef>"
        f"<a:effectRef idx='{effect[0]}'>{effect[1]}</a:effectRef>"
        f"<a:fontRef idx='{font[0]}'>{font[1]}</a:fontRef>"
        "</p:style>"
    )


#: Every probe: its key (the shape's name), the ``p:style``, and what ``a:spPr`` adds after
#: the geometry.  ``run`` is the run's own ``a:rPr`` children, ``list_style`` the body's
#: ``a:lstStyle`` level-1 run colour.
PROBES: list[dict] = []


def probe(key: str, style_xml: str, sp_pr: str = "", **extra) -> None:
    PROBES.append({"key": key, "style": style_xml, "sp_pr": sp_pr, **extra})


# a:fillRef -- each index, with a reference colour that is not the theme's default.
for idx in (0, 1, 2, 3):
    probe(f"fill-{idx}", style(fill=(idx, clr("accent1"))))
for idx in (1000, 1001, 1002, 1003):
    probe(f"fill-{idx}", style(fill=(idx, clr("accent2"))))

# Modifiers on the reference colour.
probe("mod-1-shade50", style(fill=(1, clr("accent1", shade=50000))))
probe("mod-1-lumMod75", style(fill=(1, clr("accent1", lumMod=75000))))
probe("mod-1-lumMod60-lumOff40", style(fill=(1, clr("accent1", lumMod=60000, lumOff=40000))))
probe("mod-1-tint40", style(fill=(1, clr("accent2", tint=40000))))
probe("mod-1-satMod200", style(fill=(1, clr("accent1", satMod=200000))))
probe("mod-1-srgb", style(fill=(1, clr("00B050"))))
probe("mod-1-alpha50", style(fill=(1, clr("accent1", alpha=50000))))
probe("mod-2-shade50", style(fill=(2, clr("accent2", shade=50000))))
probe("mod-3-lumMod50", style(fill=(3, clr("accent2", lumMod=50000))))
probe("mod-2-srgb", style(fill=(2, clr("00B050"))))
probe("mod-1003-shade50", style(fill=(1003, clr("accent1", shade=50000))))
probe("mod-1002-tint40", style(fill=(1002, clr("accent2", tint=40000))))

# a:fontRef.  On no fill, so the text is drawn on white whatever its colour.
probe("font-minor-lt1", style(font=("minor", clr("lt1"))))
probe("font-major-accent2", style(fill=(0, clr("accent1")), font=("major", clr("accent2"))))
probe("font-minor-none", style(fill=(0, clr("accent1")), font=("minor", "")))
probe("font-none-accent2", style(fill=(0, clr("accent1")), font=("none", clr("accent2"))))
probe("font-lumMod50", style(fill=(0, clr("accent1")), font=("minor", clr("accent1", lumMod=50000))))
probe("font-run-wins", style(fill=(0, clr("accent1")), font=("minor", clr("accent2"))),
      run="<a:solidFill><a:srgbClr val='00B050'/></a:solidFill>")
probe("font-lststyle", style(fill=(0, clr("accent1")), font=("minor", clr("accent2"))),
      list_style="<a:solidFill><a:srgbClr val='7030A0'/></a:solidFill>")

# a:lnRef -- the regression control.  No fill, so the stroke is read against white.
for idx in (0, 1, 2, 3):
    probe(f"line-{idx}", style(ln=(idx, clr("accent1", shade=50000)), fill=(0, clr("accent1"))))
probe("line-local-width", style(ln=(1, clr("accent2")), fill=(0, clr("accent1"))),
      "<a:ln w='76200'/>")
probe("default-shape", style(ln=(2, clr("accent1", shade=15000))))

# a:effectRef -- which entry the index names.
for idx in (0, 1, 2, 3):
    probe(f"effect-{idx}", style(effect=(idx, clr("accent1"))))

# The shape's own a:spPr over the reference.
probe("own-solid", style(), "<a:solidFill><a:srgbClr val='00B050'/></a:solidFill>")
probe("own-nofill", style(), "<a:noFill/>")
probe("own-line", style(ln=(2, clr("accent1"))),
      "<a:ln><a:solidFill><a:srgbClr val='FF0000'/></a:solidFill></a:ln>")
probe("own-line-none", style(ln=(3, clr("accent2"))), "<a:ln><a:noFill/></a:ln>")


def place(index: int) -> tuple[int, int, int]:
    """Slide, x and y (pt) of probe ``index``: a grid, ``COLUMNS`` across."""
    rows = (SLIDE_HEIGHT - MARGIN) // (BOX_H + CAPTION + GAP_Y)
    slide, within = divmod(index, COLUMNS * rows)
    row, column = divmod(within, COLUMNS)
    x = MARGIN + column * (BOX_W + GAP_X)
    y = MARGIN + CAPTION + row * (BOX_H + CAPTION + GAP_Y)
    return slide, x, y


def frame(x: float, y: float, cx: float, cy: float) -> str:
    return (
        f"<a:xfrm><a:off x='{int(x * PT)}' y='{int(y * PT)}'/>"
        f"<a:ext cx='{int(cx * PT)}' cy='{int(cy * PT)}'/></a:xfrm>"
    )


def caption(shape_id: int, text: str, x: float, y: float) -> str:
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{shape_id}' name='Caption'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        + frame(x, y - CAPTION, BOX_W, CAPTION)
        + "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        "<p:txBody><a:bodyPr wrap='none' lIns='0' rIns='0' tIns='0' bIns='0'/><a:lstStyle/>"
        "<a:p><a:r><a:rPr lang='en-US' sz='800'>"
        "<a:solidFill><a:srgbClr val='475569'/></a:solidFill></a:rPr>"
        f"<a:t>{escape(text)}</a:t></a:r></a:p></p:txBody></p:sp>"
    )


def probe_shape(shape_id: int, spec: dict, x: float, y: float) -> str:
    list_style = (
        f"<a:lstStyle><a:lvl1pPr><a:defRPr>{spec['list_style']}</a:defRPr></a:lvl1pPr>"
        "</a:lstStyle>"
        if "list_style" in spec
        else "<a:lstStyle/>"
    )
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{shape_id}' name='{spec['key']}'/>"
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        + frame(x, y, BOX_W, BOX_H)
        + f"<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>{spec['sp_pr']}</p:spPr>"
        + spec["style"]
        + "<p:txBody><a:bodyPr rtlCol='0' anchor='ctr'/>"
        + list_style
        + "<a:p><a:pPr algn='ctr'/>"
        f"<a:r><a:rPr lang='en-US' sz='2000' dirty='0'>{spec.get('run', '')}</a:rPr>"
        "<a:t>Styled</a:t></a:r></a:p></p:txBody></p:sp>"
    )


def slides() -> list[str]:
    trees: dict[int, list[str]] = {}
    for index, spec in enumerate(PROBES):
        slide, x, y = place(index)
        trees.setdefault(slide, []).append(
            caption(100 + 2 * index, spec["key"], x, y)
            + probe_shape(101 + 2 * index, spec, x, y)
        )
    return [
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + "".join(trees[slide])
        + "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        for slide in sorted(trees)
    ]


def _write(out: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(name, date_time=ZIP_TIMESTAMP)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o600 << 16
    out.writestr(info, data)


def write_deck(target: Path, source: Path, xml_slides: list[str] | None = None) -> int:
    """Write ``xml_slides`` (this probe's own by default) over ``source``'s master, layouts
    and theme, emptied."""
    archive = zipfile.ZipFile(source)
    presentation = archive.read("ppt/presentation.xml").decode()
    pres_rels = archive.read("ppt/_rels/presentation.xml.rels").decode()
    content_types = archive.read("[Content_Types].xml").decode()
    if xml_slides is None:
        xml_slides = slides()
    numbers = range(1, len(xml_slides) + 1)

    presentation = re.sub(
        r"<p:sldIdLst>.*?</p:sldIdLst>",
        "<p:sldIdLst>"
        + "".join(f'<p:sldId id="{255 + n}" r:id="rIdS{n}"/>' for n in numbers)
        + "</p:sldIdLst>",
        presentation,
        flags=re.S,
    )
    pres_rels = re.sub(r'<Relationship[^>]*Type="[^"]*/slide"[^>]*/>', "", pres_rels)
    pres_rels = pres_rels.replace(
        "</Relationships>",
        "".join(
            f'<Relationship Id="rIdS{n}" Type="{SLIDE_REL}" Target="slides/slide{n}.xml"/>'
            for n in numbers
        )
        + "</Relationships>",
    )
    content_types = re.sub(
        r'<Override PartName="/ppt/(slides/slide|charts/chart|notesSlides/notesSlide)\d+\.xml"'
        r"[^>]*/>",
        "",
        content_types,
    )
    content_types = content_types.replace(
        "</Types>",
        "".join(
            f'<Override PartName="/ppt/slides/slide{n}.xml" ContentType="{SLIDE_TYPE}"/>'
            for n in numbers
        )
        + "</Types>",
    )

    dropped = ("ppt/slides/", "ppt/charts/", "ppt/embeddings/", "ppt/notesSlides/")
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in archive.infolist():
            name = item.filename
            if name.startswith(dropped) or name.endswith("/"):
                continue
            data = archive.read(name)
            if name == "ppt/presentation.xml":
                data = presentation.encode()
            elif name == "ppt/_rels/presentation.xml.rels":
                data = pres_rels.encode()
            elif name == "[Content_Types].xml":
                data = content_types.encode()
            elif re.search(r"slide(Layouts|Masters)/slide\w+\d+\.xml$", name):
                data = blank_template(data.decode()).encode()
            _write(out, name, data)
        for n, xml in zip(numbers, xml_slides):
            _write(out, f"ppt/slides/slide{n}.xml", xml.encode())
            _write(
                out,
                f"ppt/slides/_rels/slide{n}.xml.rels",
                (
                    "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
                    "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/"
                    "relationships'><Relationship Id='rId1' Type='http://schemas."
                    "openxmlformats.org/officeDocument/2006/relationships/slideLayout' "
                    "Target='../slideLayouts/slideLayout1.xml'/></Relationships>"
                ).encode(),
            )
    return len(xml_slides)


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    source = SOURCES.get(target.stem)
    if source is None:
        print(f"name the deck one of {sorted(SOURCES)}, .pptx", file=sys.stderr)
        return 2
    count = write_deck(target, source)
    print(f"{len(PROBES)} probes on {count} slides -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

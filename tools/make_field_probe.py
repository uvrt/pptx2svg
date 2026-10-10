#!/usr/bin/env python3
"""Build decks that measure what PowerPoint draws for a text field (``a:fld``).

A field carries cached text -- ``‹#›`` for a slide number on a layout, the date the file
was last saved for a date -- and PowerPoint draws neither: it evaluates the field for the
slide it is drawn on.  These decks settle how, so ``resolve/fields.py`` can do the same:

* ``slidenum`` in a **non-placeholder** text box on the layout and on the master (the
  usual corporate footer), with ``p:presentation@firstSlideNum`` 1 and 5;
* ``slidenum`` and ``datetime1`` in the slide's own ``sldNum``/``dt`` placeholders, whose
  cached text is stale (``99``, ``1/1/2020``);
* ``datetime``, ``datetime1`` to ``datetime13`` and ``datetimeFigureOut``, one slide per
  language (``a:rPr@lang``), and one with no ``lang`` at all, so the format table is read
  rather than assumed;
* a field of a type PowerPoint does not evaluate, which should keep its cached text.

Every probe line is ``<key>=[<field>]``, so the reader can pick the field's text out of the
page's text by its key.  Usage::

    python3 tools/make_field_probe.py ~/field-probe.pptx            # firstSlideNum 1
    python3 tools/make_field_probe.py ~/field-probe-5.pptx --first 5
    osascript tools/powerpoint_export_pdf.applescript ~/field-probe.pptx ~/field-probe.pdf
    python3 tools/read_field_probe.py ~/field-probe.pdf

PowerPoint evaluates a date field from the clock when it draws, so note the time of the
export: ``read_field_probe.py`` prints the drawn text next to what ``resolve/fields.py``
gives for a clock you name (``--at 2026-10-10T14:05``).  The decks and exports are
throwaway and are **not** committed.
"""

from __future__ import annotations

import argparse
import io
import re
import zlib
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
SOURCE = HERE.parent / "tests/fixtures/sample.pptx"

LAYOUT = "ppt/slideLayouts/slideLayout7.xml"
MASTER = "ppt/slideMasters/slideMaster1.xml"
SLIDE_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
)

#: Languages whose date formats are read.  ``None``: the field states no ``lang``.
LANGUAGES = ("en-US", "en-GB", "nl-NL", "de-DE", "fr-FR", "ja-JP", None)
DATE_TYPES = ("datetime",) + tuple(f"datetime{n}" for n in range(1, 14)) + (
    "datetimeFigureOut",
)


def field(kind: str, cached: str, lang: str | None, size: int = 1200) -> str:
    lang_attr = f' lang="{lang}"' if lang else ""
    return (
        f'<a:fld id="{{B6F15528-21DE-4FAA-801E-{zlib.crc32(f"{kind}{lang}".encode()):012d}}}" '
        f'type="{kind}"><a:rPr{lang_attr} sz="{size}" dirty="0"/><a:t>{cached}</a:t></a:fld>'
    )


def run(text: str, size: int = 1200) -> str:
    return f'<a:r><a:rPr lang="en-US" sz="{size}" dirty="0"/><a:t>{text}</a:t></a:r>'


def text_box(shape_id: int, name: str, x: int, y: int, cx: int, cy: int, paragraphs) -> str:
    body = "".join(f"<a:p>{p}</a:p>" for p in paragraphs)
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr txBox="1"/>'
        f'<p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0">'
        f"<a:noAutofit/></a:bodyPr><a:lstStyle/>{body}</p:txBody></p:sp>"
    )


def probe_line(key: str, fld: str) -> str:
    return run(f"{key}=[") + fld + run("]")


def language_slide(lang: str | None) -> str:
    tag = lang or "none"
    lines = [probe_line(f"{tag}.{kind}", field(kind, "1/1/2020", lang)) for kind in DATE_TYPES]
    lines.append(probe_line(f"{tag}.slidenum", field("slidenum", "‹#›", lang)))
    lines.append(probe_line(f"{tag}.unknown", field("probeunknown", "CACHED", lang)))
    return text_box(10, "Probe", 457200, 300000, 11000000, 5600000, lines)


def placeholder_slide() -> str:
    def placeholder(shape_id, kind, idx, fld):
        return (
            f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{kind}"/><p:cNvSpPr>'
            f'<a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="{kind}" sz="quarter" '
            f'idx="{idx}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
            f"<a:p>{probe_line('ph.' + kind, fld)}</a:p></p:txBody></p:sp>"
        )

    return placeholder(20, "sldNum", 12, field("slidenum", "99", "en-US")) + placeholder(
        21, "dt", 10, field("datetime1", "1/1/2020", "en-US")
    )


def slide_xml(shapes: str) -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld {NS}><p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f"<p:grpSpPr/>{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/>"
        "</p:clrMapOvr></p:sld>"
    )


def build(first_slide_number: int = 1) -> bytes:
    source = zipfile.ZipFile(SOURCE)
    parts = {
        name: source.read(name)
        for name in source.namelist()
        if not name.startswith(("ppt/slides/", "ppt/notesSlides/"))
    }
    slides = [language_slide(lang) for lang in LANGUAGES] + [placeholder_slide()]

    layout_footer = text_box(
        90, "Layout footer", 457200, 6172200, 6000000, 300000,
        [probe_line("layout.slidenum", field("slidenum", "‹#›", "en-US"))
         + run(" | ") + probe_line("layout.datetime1", field("datetime1", "1/1/2020", "en-US"))],
    )
    master_footer = text_box(
        91, "Master footer", 6600000, 6172200, 5000000, 300000,
        [probe_line("master.slidenum", field("slidenum", "‹#›", "en-US"))],
    )
    for name, extra in ((LAYOUT, layout_footer), (MASTER, master_footer)):
        parts[name] = parts[name].decode().replace("</p:spTree>", extra + "</p:spTree>", 1).encode()

    presentation = parts["ppt/presentation.xml"].decode()
    rels = parts["ppt/_rels/presentation.xml.rels"].decode()
    rels = re.sub(r'<Relationship Id="[^"]+" Type="[^"]+/(slide|notesMaster)" [^>]+/>', "", rels)
    presentation = re.sub(r"<p:notesMasterIdLst>.*?</p:notesMasterIdLst>", "", presentation)
    entries = []
    for index, xml in enumerate(slides, start=1):
        parts[f"ppt/slides/slide{index}.xml"] = slide_xml(xml).encode()
        parts[f"ppt/slides/_rels/slide{index}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/slideLayout" '
            'Target="../slideLayouts/slideLayout7.xml"/></Relationships>'
        ).encode()
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="rIdS{index}" Type="{REL}/slide" '
            f'Target="slides/slide{index}.xml"/></Relationships>',
        )
        entries.append(f'<p:sldId id="{255 + index}" r:id="rIdS{index}"/>')
    presentation = re.sub(
        r"<p:sldIdLst>.*?</p:sldIdLst>", f"<p:sldIdLst>{''.join(entries)}</p:sldIdLst>",
        presentation,
    )
    presentation = presentation.replace(
        "<p:presentation ", f'<p:presentation firstSlideNum="{first_slide_number}" ', 1
    )
    parts["ppt/presentation.xml"] = presentation.encode()
    parts["ppt/_rels/presentation.xml.rels"] = rels.encode()
    # The notes master and its theme go with the notes slides.
    for name in list(parts):
        if name.startswith("ppt/notesMasters/") or name == "ppt/theme/theme2.xml":
            del parts[name]

    types = parts["[Content_Types].xml"].decode()
    types = re.sub(r'<Override PartName="/ppt/(slides|notesSlides|notesMasters)/[^>]+/>', "", types)
    types = types.replace('<Override PartName="/ppt/theme/theme2.xml" '
                          'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>', "")
    types = types.replace(
        "</Types>",
        "".join(
            f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="{SLIDE_TYPE}"/>'
            for i in range(1, len(slides) + 1)
        ) + "</Types>",
    )
    parts["[Content_Types].xml"] = types.encode()

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return out.getvalue()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("output", type=Path)
    parser.add_argument("--first", type=int, default=1, help="p:presentation@firstSlideNum")
    args = parser.parse_args()
    args.output.write_bytes(build(args.first))
    print(f"wrote {args.output}")


if __name__ == "__main__":
    main()

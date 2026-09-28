#!/usr/bin/env python3
"""Read ``tools/make_exposed_probe.py``'s deck back out of PowerPoint's PDF.

Usage::

    python3 tools/read_exposed_probe.py ~/pptx2svg-oracle/exposed-probe.pdf [text|icc|radar|legend|scene]

**text.**  Every probe's run is found by its string (``H000``...), and its first glyph's
pen position and baseline and its last glyph's advance end are read with PyMuPDF (a
character's origin).  Each reading is taken against the control -- a ``rect`` with no
outline, set the same way -- so what is printed is how much further in PowerPoint put the
text than a plain rectangle would have: on the left and top for a top-left run, on the
right and bottom for a bottom-right one, and the centre's displacement for a centred one.
Beside each is what the geometry's text rectangle alone predicts
(``ooxml_common.drawingml.geometry.text_rect``), and the difference.

**icc.**  Each picture's image object: its colour space, whether its profile is the one
the deck embedded, and the samples at each patch's centre; then MuPDF's raster of the
page at those centres, which is what a colour-managed reader of the PDF shows.

**radar.**  The page's drawing order, one line a path, with its colour and width.

**legend.**  Every small filled or stroked path near the legend's text, and the text.

**scene.**  Every image on a scene page: its pixel size, its rectangle and its resolution.

A development tool: PyMuPDF (the ``fidelity`` extra).
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

import pymupdf

sys.path.insert(0, str(Path(__file__).resolve().parent))
import make_exposed_probe as probe  # noqa: E402

EMU_PER_PT = 12700


def _spans(page) -> dict[str, dict]:
    out = {}
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                text = "".join(char["c"] for char in span["chars"])
                if text:
                    out[text] = span
    return out


def text_readings(document) -> dict[str, dict]:
    probes = probe.text_probes()
    per = len(probe.COLUMNS) * len(probe.ROWS)
    readings = {}
    for index, item in enumerate(probes):
        page = document[index // per]
        span = _spans(page).get(item["text"])
        if span is None:
            readings[item["key"]] = None
            continue
        offset = index % per
        box = (probe.COLUMNS[offset % len(probe.COLUMNS)], probe.ROWS[offset // len(probe.COLUMNS)],
               probe.BOX_W, probe.BOX_H)
        x, y, w, h = (value / EMU_PER_PT for value in box)
        first, last = span["chars"][0], span["chars"][-1]
        readings[item["key"]] = dict(
            left=first["origin"][0] - x,
            top=first["origin"][1] - y,
            right=(x + w) - last["bbox"][2],
            bottom=(y + h) - first["origin"][1],
            centre_x=(first["origin"][0] + last["bbox"][2]) / 2 - (x + w / 2),
            centre_y=first["origin"][1] - (y + h / 2),
            item=item,
        )
    return readings


def predicted(item: dict) -> tuple[float, float, float, float]:
    """The text rectangle's insets from the box, in points: left, top, right, bottom."""
    from xml.etree.ElementTree import fromstring

    from ooxml_common.drawingml import geometry, read

    sp_pr = fromstring(
        '<p:spPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + item["geometry"] + "</p:spPr>"
    )
    left, top, right, bottom = geometry.text_rect(
        read.parse_geometry_spec(sp_pr), probe.BOX_W, probe.BOX_H, rect=read.parse_text_rect(sp_pr)
    )
    return (left / EMU_PER_PT, top / EMU_PER_PT, (probe.BOX_W - right) / EMU_PER_PT,
            (probe.BOX_H - bottom) / EMU_PER_PT)


def print_text(document) -> None:
    readings = text_readings(document)
    control = {place: readings[f"rect/{place}"] for place in ("tl", "br", "cc")}
    zero = {place: readings.get(f"rect+zero/{place}") for place in ("tl", "br")}
    print(f"{'probe':34s} {'measured':>26s}   {'text rect':>26s}   {'measured - rect':>26s}")
    for key, reading in readings.items():
        if reading is None:
            print(f"{key:34s} not found")
            continue
        item = reading["item"]
        place = key.rsplit("/", 1)[1]
        base = zero[place] if item["insets"] == (0, 0, 0, 0) and place in zero and zero[place] else control[place]
        if item["insets"] == (0, 0, 0, 0) and key.startswith("rect+zero"):
            base = control[place]
        rect = predicted(item)
        if item["flip"][0]:
            rect = (rect[2], rect[1], rect[0], rect[3])
        if item["flip"][1]:
            rect = (rect[0], rect[3], rect[2], rect[1])
        if place == "tl":
            got = (reading["left"] - base["left"], reading["top"] - base["top"])
            want = (rect[0], rect[1])
        elif place == "br":
            got = (reading["right"] - base["right"], reading["bottom"] - base["bottom"])
            want = (rect[2], rect[3])
        else:
            got = (reading["centre_x"] - base["centre_x"], reading["centre_y"] - base["centre_y"])
            want = ((rect[0] - rect[2]) / 2, (rect[1] - rect[3]) / 2)
        diff = (got[0] - want[0], got[1] - want[1])
        print(f"{key:34s} {got[0]:12.3f} {got[1]:12.3f}   {want[0]:12.3f} {want[1]:12.3f}   "
              f"{diff[0]:12.3f} {diff[1]:12.3f}")


def _deck(document) -> str:
    stem = Path(document.name).stem
    return "two" if stem.endswith("-2") else "three" if stem.endswith("-3") else "one"


def _slides(document) -> list[dict]:
    return probe.all_slides(_deck(document))


def print_overflow(document) -> None:
    """Each overflow probe's three baselines, from the top of its box, and where the box's
    centre and bottom are."""
    probes = {item["key"]: item for item in probe.overflow_probes()}
    height = probe.OVERFLOW_BOX[1] / EMU_PER_PT
    for page_index, slide in enumerate(_slides(document)):
        spans = _spans(document[page_index])
        for offset, key in enumerate(slide.get("overflow", [])):
            top = (probe.ROWS[offset // len(probe.COLUMNS)] + 600000) / EMU_PER_PT
            baselines = [spans[line]["chars"][0]["origin"][1] - top if line in spans else None
                         for line in probes[key]["lines"]]
            shown = " ".join(f"{b:8.3f}" if b is not None else "  (none)" for b in baselines)
            print(f"{key:16s} baselines from the box's top: {shown}   (box {height:.1f} pt)")


def print_icc(document) -> None:
    import re

    import numpy as np
    from PIL import ImageCms

    for page_index, slide in enumerate(_slides(document)):
        if "keys" not in slide:
            continue
        page = document[page_index]
        pix = page.get_pixmap(dpi=72, colorspace=pymupdf.csRGB)
        raster = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        entries = {entry[0]: entry for entry in page.get_images(full=True)}
        for info in page.get_image_info(xrefs=True):
            xref = info["xref"]
            x0, y0, x1, y1 = info["bbox"]
            offset = int(round((x0 - 300000 / EMU_PER_PT) / (2900000 / EMU_PER_PT)))
            name = slide["keys"][offset]
            image = pymupdf.Pixmap(document, xref)
            colour_space = entries[xref][5]
            profile = None
            if colour_space == "ICCBased":
                described = document.xref_get_key(xref, "ColorSpace")
                array = document.xref_object(int(re.search(r"(\d+) 0 R", described[1]).group(1))) \
                    if described[0] == "xref" else described[1]
                found = re.search(r"/ICCBased\s+(\d+) 0 R", array)
                if found:
                    profile = document.xref_stream(int(found.group(1)))
            description = ""
            if profile:
                description = ImageCms.getProfileDescription(ImageCms.ImageCmsProfile(io.BytesIO(profile))).strip()
            samples = np.frombuffer(image.samples, dtype=np.uint8).reshape(image.height, image.width, image.n)
            filt = document.xref_get_key(xref, "Filter")[1]
            print(f"{name:12s} xref {xref} {colour_space:9s} {filt:12s} {image.width}x{image.height} "
                  f"profile={description!r} ({len(profile or b'')} bytes)")
            sx, sy = image.width / (probe.PATCH * 4), image.height / (probe.PATCH * 2)
            for patch_index, colour in enumerate(probe.PATCHES):
                px = int((patch_index % 4 + 0.5) * probe.PATCH * sx)
                py = int((patch_index // 4 + 0.5) * probe.PATCH * sy)
                rx = int(x0 + (patch_index % 4 + 0.5) * (x1 - x0) / 4)
                ry = int(y0 + (patch_index // 4 + 0.5) * (y1 - y0) / 2)
                print(f"    {str(colour):16s} samples {str(tuple(int(v) for v in samples[py, px][:3])):16s} "
                      f"MuPDF shows {tuple(int(v) for v in raster[ry, rx][:3])}")


def print_paths(page, where=None) -> None:
    for index, drawing in enumerate(page.get_drawings()):
        rect = drawing["rect"]
        if where is not None and not where(rect):
            continue
        colour = drawing.get("fill") if drawing["type"] == "f" else drawing.get("color")
        hexed = "".join(f"{round(c * 255):02X}" for c in colour) if colour else "-"
        print(f"  {index:3d} {drawing['type']:2s} {hexed} w={drawing.get('width')} items={len(drawing['items'])} "
              f"rect=({rect.x0:.3f}, {rect.y0:.3f}, {rect.x1:.3f}, {rect.y1:.3f})")


def chart_page(document, key: str):
    keys = [slide.get("key") for slide in _slides(document)]
    return document[keys.index(key)]


def _chart_keys(document, prefix: str) -> list[str]:
    return [slide["key"] for slide in _slides(document) if "key" in slide and slide["key"].startswith(prefix)]


def print_radar(document) -> None:
    for key in _chart_keys(document, "radar"):
        print(key)
        print_paths(chart_page(document, key))


def print_legend(document) -> None:
    for key in _chart_keys(document, "legend"):
        page = chart_page(document, key)
        print(key)
        print_paths(page, lambda rect: rect.width < 30 and rect.height < 20)
        for text, span in _spans(page).items():
            if text in ("North", "South", "East"):
                print(f"    text {text!r} pen x {span['chars'][0]['origin'][0]:.3f} baseline "
                      f"{span['chars'][0]['origin'][1]:.3f} size {span['size']:.2f}")


def print_scene(document) -> None:
    frames = {key: frame for key, _, frame in probe.CHARTS + probe.CHARTS_TWO}
    for key in _chart_keys(document, ""):
        frame = frames[key]
        page = chart_page(document, key)
        for info in page.get_image_info():
            x0, y0, x1, y1 = info["bbox"]
            dpi = max(info["width"] / ((x1 - x0) / 72), info["height"] / ((y1 - y0) / 72))
            print(f"{key:18s} frame {frame[2]}x{frame[3]} pt: image {info['width']}x{info['height']} px over "
                  f"({x0:.1f}, {y0:.1f}, {x1:.1f}, {y1:.1f}) = {x1 - x0:.1f}x{y1 - y0:.1f} pt, {dpi:.1f} dpi")


def print_axis3d(document) -> None:
    """Every path on a ``bar3d-`` page below the scene, and the scene image's rectangle."""
    for key in _chart_keys(document, "bar3d"):
        page = chart_page(document, key)
        print(key)
        for info in page.get_image_info():
            print(f"    image rect {tuple(round(v, 3) for v in info['bbox'])}")
        print_paths(page, lambda rect: rect.y0 > 250)


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    document = pymupdf.open(sys.argv[1])
    what = sys.argv[2] if len(sys.argv) > 2 else "all"
    shows = (("text", print_text), ("icc", print_icc), ("radar", print_radar), ("legend", print_legend),
             ("scene", print_scene), ("axis3d", print_axis3d))
    if _deck(document) == "two":
        shows = tuple(pair for pair in shows if pair[0] not in ("text", "legend"))
    if _deck(document) == "three":
        shows = (("overflow", print_overflow),)
    for name, show in shows:
        if what in (name, "all"):
            print(f"== {name}")
            show(document)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

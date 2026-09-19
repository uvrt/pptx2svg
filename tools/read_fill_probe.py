#!/usr/bin/env python3
"""Read a fill probe deck's PDF export: PowerPoint's own pattern dictionaries.

PowerPoint's PDF export writes an ``a:pattFill`` **and** a tiled ``a:blipFill`` as a PDF
tiling pattern, and the dictionary states everything the renderer needs::

    << /Type /Pattern /PatternType 1 /BBox [0 0 64 64] /XStep 64 /YStep 64
       /Matrix [0.125 0 0 0.125 32 349] /Resources 104 0 R >>

* the **cell** is ``XStep * Matrix[0]`` by ``YStep * Matrix[3]`` points, exactly;
* ``Matrix[4], Matrix[5]`` are the lattice's registration point in page coordinates
  (PDF's y runs up from the bottom of the 405 pt page), which is what ``@algn`` moves;
* the pattern's ``/Resources`` name one image XObject holding **one cell**, whose own
  pixel dimensions say whether PowerPoint composited a mirrored pair for ``@flip``.

Reading the dictionary rather than the raster is what makes a 54-preset sweep cheap: there
is no antialiasing to threshold and no pixel grid to quantise against.  For ``a:pattFill``
the cell image is the preset's 8x8 bitmap upsampled 8x, so ``--cells`` prints each preset's
bitmap as ASCII -- that is PowerPoint's drawing of the preset, read rather than inferred.

Usage::

    python3 tools/read_fill_probe.py <fill-patt.pdf> --cells
    python3 tools/read_fill_probe.py <fill-pitch.pdf>
    python3 tools/read_fill_probe.py <fill-tile.pdf>
    python3 tools/read_fill_probe.py <fill-patt.pdf> --python   # the PRESET_CELLS table

Each pattern is matched to its probe by its registration point: the probe whose box is
nearest, which the deck's layout keeps unambiguous.
"""

from __future__ import annotations

import re
import sys
import zlib
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from make_fill_probe import PRESETS, probes_for  # noqa: E402

PAGE_HEIGHT = 405.0
EMU_PER_POINT = 12700

OBJ = re.compile(rb"(\d+)\s+(\d+)\s+obj\b(.*?)\bendobj", re.S)
PATTERN = re.compile(
    rb"/BBox\s*\[([^\]]*)\]\s*/XStep\s*([\d.]+)\s*/YStep\s*([\d.]+)\s*"
    rb"/Matrix\s*\[([^\]]*)\]\s*/Resources\s*(\d+)"
)
XOBJECT = re.compile(rb"/XObject\s*<<\s*/(\w+)\s+(\d+)")
SIZE = re.compile(rb"/Width\s+(\d+)\s*/Height\s+(\d+)")
KIDS = re.compile(rb"/Kids\s*\[([^\]]*)\]")
REF = re.compile(rb"(\d+)\s+0\s+R")


def objects(data: bytes) -> dict[int, bytes]:
    return {int(m.group(1)): m.group(3) for m in OBJ.finditer(data)}


def stream_of(body: bytes) -> bytes | None:
    start = body.find(b"stream")
    if start < 0:
        return None
    cursor = start + 6
    while body[cursor : cursor + 1] in (b"\r", b"\n"):
        cursor += 1
    raw = body[cursor : body.rfind(b"endstream")]
    if b"FlateDecode" not in body[:start]:
        return raw
    try:
        return zlib.decompressobj().decompress(raw)
    except zlib.error:
        return None


def page_order(objs: dict[int, bytes]) -> list[int]:
    """Page object numbers in reading order, flattening the page tree."""
    roots = [num for num, body in objs.items() if b"/Type /Catalog" in body]
    if not roots:
        return []
    top = int(REF.search(objs[roots[0]][objs[roots[0]].find(b"/Pages") :]).group(1))

    def walk(num: int) -> list[int]:
        body = objs[num]
        if b"/Type /Page\n" in body or b"/Type /Page " in body or b"/Type /Page>" in body:
            return [num]
        kids = KIDS.search(body)
        if not kids:
            return []
        out = []
        for ref in REF.finditer(kids.group(1)):
            out.extend(walk(int(ref.group(1))))
        return out

    return walk(top)


def patterns(data: bytes) -> list[dict]:
    """Every tiling pattern in the file, with its page index and cell image."""
    objs = objects(data)
    pages = page_order(objs)
    resource_of_page = {}
    for index, num in enumerate(pages):
        res = re.search(rb"/Resources\s+(\d+)", objs[num])
        if res:
            resource_of_page[int(res.group(1))] = index

    # Which page each pattern belongs to: the page resource dict that names it.
    page_of_pattern: dict[int, int] = {}
    for res_num, page_index in resource_of_page.items():
        for ref in REF.finditer(objs[res_num]):
            page_of_pattern[int(ref.group(1))] = page_index

    out = []
    for num, body in sorted(objs.items()):
        if b"/PatternType" not in body:
            continue
        match = PATTERN.search(body)
        if not match:
            continue
        matrix = [float(v) for v in match.group(4).split()]
        resources = objs[int(match.group(5))]
        xobject = XOBJECT.search(resources)
        image = int(xobject.group(2)) if xobject else None
        entry = {
            "obj": num,
            "page": page_of_pattern.get(num, -1),
            "step": (float(match.group(2)), float(match.group(3))),
            "matrix": matrix,
            "cell": (float(match.group(2)) * matrix[0], float(match.group(3)) * matrix[3]),
            "at": (matrix[4], PAGE_HEIGHT - matrix[5]),
            "pixels": None,
            "bitmap": None,
        }
        if image is not None:
            size = SIZE.search(objs[image])
            entry["pixels"] = (int(size.group(1)), int(size.group(2)))
            entry["bitmap"] = (objs[image], stream_of(objs[image]))
        out.append(entry)
    return out


def cell_grid(pixels: tuple[int, int], raw: bytes | None, cell: int = 8) -> list[str] | None:
    """Collapse the export's upsample of a preset's bitmap back to its own grid."""
    if raw is None:
        return None
    width, height = pixels
    components = len(raw) // (width * height) if width * height else 0
    if not components or width % cell or height % cell:
        return None
    step_x, step_y = width // cell, height // cell
    grid = []
    for cy in range(cell):
        row = ""
        for cx in range(cell):
            x, y = cx * step_x + step_x // 2, cy * step_y + step_y // 2
            offset = (y * width + x) * components
            value = sum(raw[offset : offset + min(3, components)]) / min(3, components)
            row += "#" if value < 128 else "."
        grid.append(row)
    return grid


def matched(pdf: Path) -> list[tuple[dict, dict]]:
    """Pair each pattern with the probe whose box it registers against."""
    data = pdf.read_bytes()
    probes = probes_for(pdf.name)
    found = patterns(data)
    pairs = []
    for entry in found:
        candidates = [p for p in probes if p["slide"] == entry["page"]]
        if not candidates:
            continue
        x, y = entry["at"]
        best = min(
            candidates,
            key=lambda p: abs(p["x"] / EMU_PER_POINT - x) + abs(p["y"] / EMU_PER_POINT - y),
        )
        pairs.append((best, entry))
    return pairs


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    pdf = Path(sys.argv[1]).expanduser()
    flags = set(sys.argv[2:])
    pairs = matched(pdf)

    if "--python" in flags:
        by_key = {probe["key"]: entry for probe, entry in pairs}
        print("PRESET_CELLS = {")
        for preset in PRESETS:
            entry = by_key.get(preset)
            grid = cell_grid(entry["pixels"], entry["bitmap"][1]) if entry else None
            if grid is None:
                print(f"    # {preset}: no pattern -- PowerPoint drew it some other way")
                continue
            rows = ", ".join(f"0b{row.replace('#', '1').replace('.', '0')}" for row in grid)
            print(f"    {preset!r}: ({rows}),")
        print("}")
        return 0

    for probe, entry in pairs:
        box = (probe["x"] / EMU_PER_POINT, probe["y"] / EMU_PER_POINT,
               probe["cx"] / EMU_PER_POINT, probe["cy"] / EMU_PER_POINT)
        print(
            f"{probe['key']:<26} cell={entry['cell'][0]:8.4f} x {entry['cell'][1]:8.4f} pt  "
            f"image={entry['pixels']}  at=({entry['at'][0]:8.3f},{entry['at'][1]:8.3f})  "
            f"box=({box[0]:.3f},{box[1]:.3f},{box[2]:.0f}x{box[3]:.0f})"
        )
        if "--cells" in flags:
            grid = cell_grid(entry["pixels"], entry["bitmap"][1])
            for row in grid or ["  (not an 8x8 upsample)"]:
                print("      ", row)
    print(f"\n{len(pairs)} patterns")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

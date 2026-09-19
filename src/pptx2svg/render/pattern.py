"""PowerPoint's own hatch cells for ``a:pattFill``, read out of its PDF export.

Every one of the 54 presets in ``ST_PresetPatternVal`` (ECMA-376 §20.1.10.51) is an **8x8
one-bit bitmap drawn on an 8.0 pt cell** -- one point per bit.  Both halves of that are
measured, and neither was what this library did before:

* the cell used to be a hardcoded ``size = 8.0`` in *pixels*, which is 6 pt, so every
  pattern tiled a third too finely; and
* 23 presets were drawn from hand-written lines and rectangles, the other 31 fell through
  to a flat solid fill, and several of the 23 did not match the preset they named.

``tools/make_fill_probe.py`` (deck ``fill-patt``) puts all 54 on a slide; PowerPoint's PDF
export writes each as a tiling pattern whose ``/XStep`` and ``/Matrix`` give the cell
exactly -- 8.0000 x 8.0000 pt for every preset -- and whose content is that preset's own
bitmap upsampled 8x to 64x64.  ``tools/read_fill_probe.py --python`` decodes the upsample
back to the 8x8 source and prints this table, so every row below is PowerPoint's drawing
rather than a reading of the preset's name.

**The names do not predict the drawing**, which is the whole reason this is a measurement:

* ``horz`` is one rule per cell and ``ltHorz`` is two -- the old code aliased them together.
* ``lgGrid`` and ``cross`` have the **same** bitmap; the old code drew ``lgGrid`` on a cell
  twice the size.
* ``dkDnDiag`` is a **2 px wide** diagonal at a 4 px pitch, not the pair of hairlines that
  were drawn for it; ``dkHorz`` and ``dkVert`` are likewise 2 px rules, not hairlines.
* ``cross``, ``smGrid`` and ``lgGrid`` register their rules on the cell's **edge**; the old
  code centred them.
* the ``pct*`` family's coverage is nominal, not literal -- ``pct20`` is 8 bits of 64
  (12.5%) and ``pct30`` is 24 (37.5%).  They are dither masks with historical names, and
  "fix" them to their nominal density and the picture stops matching PowerPoint.

Three presets -- ``dnDiag``, ``upDiag`` and ``diagCross`` -- carry a half-intensity fringe
on both sides of the diagonal in the exported bitmap, i.e. PowerPoint antialiases those
three in the source cell.  The table keeps the solid run only; a stair-stepped 1 pt
diagonal is within half a pixel of the fringe at any size a slide is viewed at, and
carrying a second intensity through the whole pattern path to gain that is not worth it.

The cell is a **fixed length**, not a fraction of the shape: deck ``fill-pitch`` drew
``cross`` on eight boxes from 24x18 to 640x320 pt and every one came back 8.0000 pt.

Registration is to the **page**, not to the shape.  The same deck put a shape's left edge
at 36.0, 36.5, 38.0, 41.0, 100.3, 173.75, 260.125 and 411.0 pt and PowerPoint's pattern
origin snapped every one down to the 8 pt lattice measured from the slide's own top-left
corner -- 32, 32, 32, 40, 96, 168, 256, 408.  See ``render/fill.py`` for what this
renderer does about that.
"""

from __future__ import annotations

#: The cell's side, in points.  Measured at 8.0000 for all 54 presets at all shape sizes.
PATTERN_CELL_PT = 8.0

#: The cell's resolution: 8 x 8 bits, one bit per point.
PATTERN_CELL_BITS = 8

#: ``preset -> eight rows of eight bits``, most significant bit leftmost.  Set means
#: foreground.  Printed by ``tools/read_fill_probe.py --python``; do not hand-edit.
PRESET_CELLS: dict[str, tuple[int, int, int, int, int, int, int, int]] = {
    "pct5": (0b10000000, 0b00000000, 0b00000000, 0b00000000,
             0b00001000, 0b00000000, 0b00000000, 0b00000000),
    "pct10": (0b10000000, 0b00000000, 0b00001000, 0b00000000,
              0b10000000, 0b00000000, 0b00001000, 0b00000000),
    "pct20": (0b10001000, 0b00000000, 0b00100010, 0b00000000,
              0b10001000, 0b00000000, 0b00100010, 0b00000000),
    "pct25": (0b10001000, 0b00100010, 0b10001000, 0b00100010,
              0b10001000, 0b00100010, 0b10001000, 0b00100010),
    "pct30": (0b10101010, 0b01000100, 0b10101010, 0b00010001,
              0b10101010, 0b01000100, 0b10101010, 0b00010001),
    "pct40": (0b10101010, 0b01010101, 0b10101010, 0b01010001,
              0b10101010, 0b01010101, 0b10101010, 0b00010101),
    "pct50": (0b10101010, 0b01010101, 0b10101010, 0b01010101,
              0b10101010, 0b01010101, 0b10101010, 0b01010101),
    "pct60": (0b11101110, 0b01010101, 0b10111011, 0b01010101,
              0b11101110, 0b01010101, 0b10111011, 0b01010101),
    "pct70": (0b01110111, 0b11011101, 0b01110111, 0b11011101,
              0b01110111, 0b11011101, 0b01110111, 0b11011101),
    "pct75": (0b01110111, 0b11111111, 0b11011101, 0b11111111,
              0b01110111, 0b11111111, 0b11011101, 0b11111111),
    "pct80": (0b11101111, 0b11111111, 0b11111110, 0b11111111,
              0b11101111, 0b11111111, 0b11111110, 0b11111111),
    "pct90": (0b11111111, 0b11111111, 0b11111111, 0b11110111,
              0b11111111, 0b11111111, 0b11111111, 0b01111111),
    "horz": (0b11111111, 0b00000000, 0b00000000, 0b00000000,
             0b00000000, 0b00000000, 0b00000000, 0b00000000),
    "vert": (0b10000000, 0b10000000, 0b10000000, 0b10000000,
             0b10000000, 0b10000000, 0b10000000, 0b10000000),
    "ltHorz": (0b11111111, 0b00000000, 0b00000000, 0b00000000,
               0b11111111, 0b00000000, 0b00000000, 0b00000000),
    "ltVert": (0b10001000, 0b10001000, 0b10001000, 0b10001000,
               0b10001000, 0b10001000, 0b10001000, 0b10001000),
    "dkHorz": (0b11111111, 0b11111111, 0b00000000, 0b00000000,
               0b11111111, 0b11111111, 0b00000000, 0b00000000),
    "dkVert": (0b11001100, 0b11001100, 0b11001100, 0b11001100,
               0b11001100, 0b11001100, 0b11001100, 0b11001100),
    "narHorz": (0b11111111, 0b00000000, 0b11111111, 0b00000000,
                0b11111111, 0b00000000, 0b11111111, 0b00000000),
    "narVert": (0b01010101, 0b01010101, 0b01010101, 0b01010101,
                0b01010101, 0b01010101, 0b01010101, 0b01010101),
    "dashHorz": (0b11110000, 0b00000000, 0b00000000, 0b00000000,
                 0b00001111, 0b00000000, 0b00000000, 0b00000000),
    "dashVert": (0b10000000, 0b10000000, 0b10000000, 0b10000000,
                 0b00001000, 0b00001000, 0b00001000, 0b00001000),
    "cross": (0b11111111, 0b10000000, 0b10000000, 0b10000000,
              0b10000000, 0b10000000, 0b10000000, 0b10000000),
    "dnDiag": (0b10000000, 0b01000000, 0b00100000, 0b00010000,
               0b00001000, 0b00000100, 0b00000010, 0b00000001),
    "upDiag": (0b00000001, 0b00000010, 0b00000100, 0b00001000,
               0b00010000, 0b00100000, 0b01000000, 0b10000000),
    "ltDnDiag": (0b10001000, 0b01000100, 0b00100010, 0b00010001,
                 0b10001000, 0b01000100, 0b00100010, 0b00010001),
    "ltUpDiag": (0b00010001, 0b00100010, 0b01000100, 0b10001000,
                 0b00010001, 0b00100010, 0b01000100, 0b10001000),
    "dkDnDiag": (0b11001100, 0b01100110, 0b00110011, 0b10011001,
                 0b11001100, 0b01100110, 0b00110011, 0b10011001),
    "dkUpDiag": (0b00110011, 0b01100110, 0b11001100, 0b10011001,
                 0b00110011, 0b01100110, 0b11001100, 0b10011001),
    "wdDnDiag": (0b11000001, 0b11100000, 0b01110000, 0b00111000,
                 0b00011100, 0b00001110, 0b00000111, 0b10000011),
    "wdUpDiag": (0b10000011, 0b00000111, 0b00001110, 0b00011100,
                 0b00111000, 0b01110000, 0b11100000, 0b11000001),
    "dashDnDiag": (0b00000000, 0b00000000, 0b10001000, 0b01000100,
                   0b00100010, 0b00010001, 0b00000000, 0b00000000),
    "dashUpDiag": (0b00000000, 0b00000000, 0b00010001, 0b00100010,
                   0b01000100, 0b10001000, 0b00000000, 0b00000000),
    "diagCross": (0b10000001, 0b01000010, 0b00100100, 0b00011000,
                  0b00011000, 0b00100100, 0b01000010, 0b10000001),
    "smGrid": (0b11111111, 0b10001000, 0b10001000, 0b10001000,
               0b11111111, 0b10001000, 0b10001000, 0b10001000),
    "lgGrid": (0b11111111, 0b10000000, 0b10000000, 0b10000000,
               0b10000000, 0b10000000, 0b10000000, 0b10000000),
    "dotGrid": (0b10101010, 0b00000000, 0b10000000, 0b00000000,
                0b10000000, 0b00000000, 0b10000000, 0b00000000),
    "smCheck": (0b10011001, 0b01100110, 0b01100110, 0b10011001,
                0b10011001, 0b01100110, 0b01100110, 0b10011001),
    "lgCheck": (0b11110000, 0b11110000, 0b11110000, 0b11110000,
                0b00001111, 0b00001111, 0b00001111, 0b00001111),
    "openDmnd": (0b10000010, 0b01000100, 0b00101000, 0b00010000,
                 0b00101000, 0b01000100, 0b10000010, 0b00000001),
    "solidDmnd": (0b00010000, 0b00111000, 0b01111100, 0b11111110,
                  0b01111100, 0b00111000, 0b00010000, 0b00000000),
    "dotDmnd": (0b10000000, 0b00000000, 0b00100010, 0b00000000,
                0b00001000, 0b00000000, 0b00100010, 0b00000000),
    "plaid": (0b10101010, 0b01010101, 0b10101010, 0b01010101,
              0b11110000, 0b11110000, 0b11110000, 0b11110000),
    "sphere": (0b01110111, 0b10001001, 0b10001111, 0b10001111,
               0b01110111, 0b10011000, 0b11111000, 0b11111000),
    "weave": (0b10001000, 0b01010100, 0b00100010, 0b01000101,
              0b10001000, 0b00010100, 0b00100010, 0b01010001),
    "divot": (0b00000000, 0b00010000, 0b00001000, 0b00010000,
              0b00000000, 0b10000000, 0b00000001, 0b10000000),
    "shingle": (0b00000011, 0b10000100, 0b01001000, 0b00110000,
                0b00001100, 0b00000010, 0b00000001, 0b00000001),
    "wave": (0b00000000, 0b00011000, 0b00100101, 0b11000000,
             0b00000000, 0b00011000, 0b00100101, 0b11000000),
    "trellis": (0b11111111, 0b01100110, 0b11111111, 0b10011001,
                0b11111111, 0b01100110, 0b11111111, 0b10011001),
    "zigZag": (0b10000001, 0b01000010, 0b00100100, 0b00011000,
               0b10000001, 0b01000010, 0b00100100, 0b00011000),
    "smConfetti": (0b10000000, 0b00001000, 0b01000000, 0b00000010,
                   0b00010000, 0b00000001, 0b00100000, 0b00000100),
    "lgConfetti": (0b10110001, 0b00110000, 0b00000011, 0b00011011,
                   0b11011000, 0b11000000, 0b00001100, 0b10001101),
    "horzBrick": (0b11111111, 0b10000000, 0b10000000, 0b10000000,
                  0b11111111, 0b00001000, 0b00001000, 0b00001000),
    "diagBrick": (0b00000001, 0b00000010, 0b00000100, 0b00001000,
                  0b00011000, 0b00100100, 0b01000010, 0b10000001),
}


def cell_rectangles(preset: str) -> list[tuple[int, int, int, int]] | None:
    """``(x, y, width, height)`` boxes covering the preset's set bits, or ``None``.

    The bits are merged into runs along each row and then down the rows, so a solid rule
    is one rectangle rather than eight, and the whole 54-preset table costs about six
    rectangles a cell instead of sixty-four.  Emitting one ``<rect>`` per *bit* would draw
    the same picture, but it also puts a seam between every pair of touching bits in every
    renderer that antialiases, which is visible on the solid presets.
    """
    rows = PRESET_CELLS.get(preset)
    if rows is None:
        return None

    out: list[tuple[int, int, int, int]] = []
    open_runs: dict[tuple[int, int], int] = {}
    for y, bits in enumerate(rows):
        runs = set()
        x = 0
        while x < PATTERN_CELL_BITS:
            if bits & (1 << (PATTERN_CELL_BITS - 1 - x)):
                start = x
                while x < PATTERN_CELL_BITS and bits & (1 << (PATTERN_CELL_BITS - 1 - x)):
                    x += 1
                runs.add((start, x - start))
            else:
                x += 1
        for run in [run for run in open_runs if run not in runs]:
            top = open_runs.pop(run)
            out.append((run[0], top, run[1], y - top))
        for run in runs:
            open_runs.setdefault(run, y)
    for run, top in open_runs.items():
        out.append((run[0], top, run[1], PATTERN_CELL_BITS - top))
    return sorted(out)

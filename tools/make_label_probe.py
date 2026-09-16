#!/usr/bin/env python3
"""Build a deck of column charts whose category labels are too wide to fit level.

The question this exists to settle is the **cap** on
:meth:`pptx2svg.resolve.chart.ChartBuilder._bottom_label_band`'s rotated reserve.  Ours is
``ROTATED_LABEL_INSET_PT + widest * sin 45`` and runs away without limit; PowerPoint's
stops somewhere, and three observations could not say where -- see that method's docstring
for why a rule fitted to them was recorded rather than shipped.

Every probe is the same chart with two things varied: the **frame** and the **label**.  A
column chart, five categories unless the key says otherwise, one series, the value axis
carrying gridlines so the plot's own edges are readable, and labels with no space in them
so wrapping cannot rescue them and the axis must turn.

The label alphabets are chosen so the reading inverts cleanly:

``M``
    Arial's widest Latin letter, 8.33 pt at 10 pt.  A ladder of these steps the drawn
    width coarsely.
``I``
    2.78 pt at 10 pt, three times finer, which is what brackets a budget to a point or so.

Usage::

    python3 tools/make_label_probe.py ~/pptx2svg-oracle/label-cap.pptx
    osascript tools/powerpoint_export_pdf.applescript ~/pptx2svg-oracle/label-cap.pptx \
        ~/pptx2svg-oracle/label-cap.pdf
    python3 tools/read_label_probe.py ~/pptx2svg-oracle/label-cap.pdf
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from make_axis_probe import write_deck  # noqa: E402

PT = 12700

#: The frame the original rotated-label probes were measured on, and the one the two
#: anomalous readings -- 85.63 and 86.36 pt -- came off.
BASE_W, BASE_H = 220.4724, 181.1024


def probe(
    key: str,
    label: str,
    *,
    width: float = BASE_W,
    height: float = BASE_H,
    size: int = 1000,
    face: str | None = "Arial",
    count: int = 5,
    labels: tuple[str, ...] | None = None,
) -> dict:
    return {
        "key": key,
        "low": 0.0,
        "high": 9.0,
        "decade": 1.0,
        "frame": (int(round(width * PT)), int(round(height * PT))),
        "size": size,
        "kind": "col",
        "face": face,
        "labels": labels or tuple(label for _ in range(count)),
    }


#: The frame heights the sweep walks.  The slide is 405 pt tall and the frame starts 12 pt
#: down, so 380 is the tallest that fits.
HEIGHTS = (70.0, 90.0, 110.0, 130.0, 150.0, 170.0, BASE_H, 210.0, 250.0, 300.0, 340.0, 380.0)

#: The first deck.  ``hM``/``hI`` sweep the frame height under a label far too wide for any
#: plausible cap, in two alphabets whose per-character width differs by three, which
#: separates a budget counted in characters from one measured in points.  ``w`` varies the
#: frame *width* and ``n`` the category count -- both of which move the band and the plot
#: while holding the frame height, the trap the tick-rule sweep already walked into.  ``z``
#: varies the label size.  ``rep`` replicates the original ladder, five labels differing
#: only in a trailing letter, in the 10 pt Aptos it was measured in: it is the one family
#: that can reproduce the 0.73 pt inversion.  ``a`` and ``x`` read the cut itself.
CAP_PROBES: list[dict] = [
    *[probe(f"hM{height:.0f}", "M" * 40, height=height) for height in HEIGHTS],
    *[probe(f"hI{height:.0f}", "I" * 90, height=height) for height in HEIGHTS],
    *[probe(f"w{width:.0f}", "I" * 90, width=width) for width in (150.0, 180.0, 250.0, 320.0, 400.0, 500.0)],
    *[probe(f"n{count}", "I" * 90, count=count) for count in (3, 8, 12)],
    *[probe(f"z{size // 100}", "I" * 90, size=size) for size in (600, 1400, 2000)],
    *[
        probe(
            f"rep{n}",
            "",
            size=1000,
            face=None,
            labels=tuple("M" * n + letter for letter in "ABCDE"),
        )
        for n in (8, 10, 12, 14, 16, 18, 20, 22, 26, 30)
    ],
    probe("a-alpha", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"),
    probe("a-alpha-s", "ABCDEFGHIJKLMNOPQRSTUVWXYZ", height=110.0),
    probe("x-wide-first", "MMMMMMMMMMMMMMMMIIIIIIIIIIIIIIII"),
    probe("x-narrow-first", "IIIIIIIIIIIIIIIIMMMMMMMMMMMMMMMM"),
]

#: ``real-financial-report.pptx``'s chart3, which is the one reading the first deck cannot
#: place: its 135 pt frame reserves 69.538 pt where the first deck's rule allows 62.
CORPUS_CATS = ("デジタル", "プラットフォーム", "グローバル", "その他")
CJK_LONG = tuple("プラットフォーム" * 3 for _ in range(4))
CORPUS_W, CORPUS_H = 375.0, 135.0


def corpus(key: str, **kwargs) -> dict:
    return probe(
        key,
        "",
        width=kwargs.pop("width", CORPUS_W),
        height=kwargs.pop("height", CORPUS_H),
        size=kwargs.pop("size", 1200),
        labels=kwargs.pop("labels", CORPUS_CATS),
        **kwargs,
    )


#: The second deck, on the two things the first left open.
#:
#: ``f`` walks a **full** label past the point where it is cut, one narrow character at a
#: time, which says whether the budget a label must fit is the same one its ellipsis is
#: measured against.
#:
#: ``s`` sweeps the label size at three frame heights.  The first deck's four sizes agreed
#: with a cap on the *reserve* and disagreed with one on the label's width, and three
#: heights turn that into a surface rather than a line.
#:
#: ``g`` is the trap the tick sweep already fell into: a bottom legend and a title take
#: height away from the plot without touching the frame, so they separate "half the frame"
#: from "half of what is left of it".
#:
#: ``cc`` replicates chart3 and takes one thing off it at a time -- its Japanese labels,
#: its 12 pt size, its four categories, its 135 pt frame -- because that chart reserves
#: 7.5 pt more than the first deck's rule allows and only one of those can be why.
CORPUS_PROBES: list[dict] = [
    *[probe(f"f{k}", "I" * k) for k in (33, 34, 35, 36, 37, 38)],
    *[
        probe(f"s{size // 100}-{height:.0f}", "I" * 120, height=height, size=size)
        for size in (600, 800, 1200, 1400, 1800, 2400)
        for height in (110.0, BASE_H, 300.0)
    ],
    *[
        {**probe(f"g{what}", "I" * 90), **extra}
        for what, extra in (
            ("lb", {"legend": "b"}),
            ("lr", {"legend": "r"}),
            ("ti", {"title": "Probe title"}),
        )
    ],
    corpus("cc-asis"),
    corpus("cc-long", labels=CJK_LONG),
    corpus("cc-10pt", size=1000, labels=CJK_LONG),
    corpus("cc-14pt", size=1400, labels=CJK_LONG),
    corpus("cc-latin", labels=("M" * 5, "M" * 10, "M" * 6, "M" * 4)),
    corpus("cc-latinlong", labels=tuple("I" * 90 for _ in range(4))),
    corpus("cc-five", labels=tuple("プラットフォーム" * 3 for _ in range(5))),
    *[
        corpus(f"cc-h{height:.0f}", height=height, labels=CJK_LONG)
        for height in (90.0, BASE_H, 250.0)
    ],
    *[
        corpus(f"cc-lat-h{height:.0f}", height=height, labels=tuple("I" * 120 for _ in range(4)))
        for height in (90.0, BASE_H, 250.0)
    ],
]

#: The third deck, which is a check rather than a search.  The rule the first two settle
#: was fitted without any of these, and each family asks it for a prediction somewhere it
#: has not been:
#:
#: ``gb``/``gt`` put a bottom legend and a title on three frame heights each -- the first
#: two decks had one probe of each -- and ``gbt`` puts both on one chart, which is the only
#: reading that says whether the two bands come off together.
#:
#: ``j`` repeats the Japanese labels at 10 pt over four heights, because the single 10 pt
#: CJK reading in ``label-corpus`` is the one probe of 88 the rule cannot place.
#:
#: ``tnr`` and ``cal`` are a third and fourth face at four sizes: the anchor is the label's
#: **ascent** turned through 45 degrees, and Times New Roman's ascent is 0.891 em against
#: Arial's 0.905 and Aptos' 0.939, so a rule that read the line box or the em instead
#: misses it.
#:
#: ``odd`` adds two frame heights that are not round numbers, and ``q`` revisits the two
#: charts that came back with no category labels drawn at all.
CHECK_PROBES: list[dict] = [
    *[
        {**probe(f"gb{height:.0f}", "I" * 120, height=height), "legend": "b"}
        for height in (110.0, BASE_H, 300.0)
    ],
    *[
        {**probe(f"gt{height:.0f}", "I" * 120, height=height), "title": "Probe title"}
        for height in (110.0, BASE_H, 300.0)
    ],
    {**probe("gbt", "I" * 120), "legend": "b", "title": "Probe title"},
    {**probe("glt", "I" * 120), "legend": "t"},
    *[
        corpus(f"j{height:.0f}", height=height, size=1000, labels=CJK_LONG)
        for height in (90.0, 135.0, BASE_H, 250.0)
    ],
    *[
        probe(f"tnr{size // 100}", "I" * 120, size=size, face="Times New Roman")
        for size in (800, 1000, 1400, 2000)
    ],
    *[
        probe(f"cal{size // 100}", "I" * 120, size=size, face="Calibri")
        for size in (1000, 1800)
    ],
    *[probe(f"odd{height:.0f}", "I" * 120, height=height) for height in (123.0, 247.0)],
    probe("q24", "I" * 40, size=2400),
    probe("q12cats", "I" * 40, count=12),
]

DECKS = {
    "label-cap": CAP_PROBES,
    "label-corpus": CORPUS_PROBES,
    "label-check": CHECK_PROBES,
}


def probes_for(target: Path) -> list[dict]:
    for name, table in DECKS.items():
        if name in target.stem:
            return table
    raise SystemExit(f"name the deck after one of {sorted(DECKS)}, not {target.stem!r}")


def main() -> int:
    target = Path(sys.argv[1]).expanduser() if len(sys.argv) > 1 else Path("label-cap.pptx")
    probes = probes_for(target)
    write_deck(target, probes)
    for index, item in enumerate(probes):
        width, height = (dimension / PT for dimension in item["frame"])
        print(
            f"{index + 1:3d} {item['key']:14s} frame={width:.1f}x{height:.1f}pt "
            f"size={item['size'] / 100:g} face={item['face'] or 'Aptos'} "
            f"cats={len(item['labels'])} label={item['labels'][0]!r}"
        )
    print(f"\n{len(probes)} probes -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

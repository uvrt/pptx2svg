"""A data label with text of its own is drawn as PowerPoint draws it: the deck is
``tools/make_point_label_probe.py``'s, the expected texts PowerPoint 16's export of it
(the reading itself is ooxml-common's ``tests/test_chart_point_labels.py``)."""

from __future__ import annotations

import re
import sys
from pathlib import Path

from pptx2svg import convert_pptx_to_svg

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import make_point_label_probe as probe  # noqa: E402


def _texts(svg: str) -> list[str]:
    return ["".join(re.findall(r">([^<]*)<", text)) for text in re.findall(r"<text[^>]*>.*?</text>", svg, re.S)]


def test_the_probe_deck_draws_each_points_own_text(tmp_path):
    deck = tmp_path / "point-labels.pptx"
    probe.build(deck)
    rich, bubble, ranged, fields, strref = (_texts(svg) for svg in convert_pptx_to_svg(deck))
    for texts in (rich, bubble):
        assert {"R1", "Bold plain", "Big", "Hidden"} <= set(texts)
    assert {"alpha; North; 3", "beta; South; 5", "gamma; East; 4"} <= set(ranged)
    assert {"alpha | North | Sales | 3", "gamma | East | Sales | 4"} <= set(fields)
    assert "From a cell" in strref

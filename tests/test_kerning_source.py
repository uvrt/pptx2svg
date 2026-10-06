"""PowerPoint kerns a static face with its legacy ``kern`` table, not its ``GPOS`` pairs.

Measured by ``tools/make_kern_source_probe.py`` (29 lines of pairs chosen by class, in
fifteen faces) and pptx-agent's ``tools/wrap_boundary_probe.py``; the rule is
``ooxml_common.drawingml.rules.POWERPOINT.kerning`` and the tables carry both
(``ooxml_common.text.kerning``).  End to end here: trial 2's "Pass", Aptos 18 pt, in a
650,000 EMU box, which PowerPoint breaks "Pas / s" -- Aptos's ``ss`` (-33/2048 em) is in
``GPOS`` alone, so the word is 0.29 pt wider than the feature-kerned width that fit.
"""

from __future__ import annotations

import re
from pathlib import Path

from deckbuilder import derive_deck

from pptx2svg import ConvertOptions, convert_pptx_to_svg

FIXTURE = Path(__file__).parent / "fixtures" / "authoring-integration.pptx"

PASS_BOX = (
    "<p:sp><p:nvSpPr><p:cNvPr id='901' name='Pass'/><p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr>"
    "<p:spPr><a:xfrm><a:off x='914400' y='914400'/><a:ext cx='650000' cy='600000'/></a:xfrm>"
    "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom></p:spPr>"
    "<p:txBody><a:bodyPr wrap='square'><a:noAutofit/></a:bodyPr><a:lstStyle/>"
    "<a:p><a:r><a:rPr lang='en-US' sz='1800' kern='1200'><a:latin typeface='Aptos'/></a:rPr>"
    "<a:t>Pass</a:t></a:r></a:p></p:txBody></p:sp>"
)


def test_pass_breaks_where_powerpoint_breaks_it(tmp_path):
    deck = tmp_path / "pass.pptx"
    deck.write_bytes(derive_deck(FIXTURE, shapes_xml=PASS_BOX))
    svg = convert_pptx_to_svg(str(deck), ConvertOptions(slide_numbers=[1], host_fonts=False))[0]
    lines = re.findall(r">(Pas|Pass|s)</tspan>", svg)
    assert lines[-2:] == ["Pas", "s"]

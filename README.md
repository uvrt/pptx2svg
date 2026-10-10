# pptx2svg

[![CI](https://github.com/uvrt/pptx2svg/actions/workflows/ci.yml/badge.svg)](https://github.com/uvrt/pptx2svg/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13%20%7C%203.14%20%7C%203.15-blue)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

Render PowerPoint (`.pptx`) slides to **SVG**, and SVG to **PNG**, in pure Python.

Parsing, layout and SVG generation need nothing beyond the standard library (and the
sibling [ooxml-common](https://github.com/uvrt/ooxml-common)); PNG output delegates to
[resvg](https://github.com/linebender/resvg) via `resvg-py`. Shape geometry, chart layout,
table styles and font metrics are **checked against PowerPoint's own output** rather than
guessed, and the constants carry the measurement that produced them.

## Install

Neither pptx2svg nor ooxml-common is on PyPI yet; install from GitHub:

```bash
pip install "ooxml-common @ git+https://github.com/uvrt/ooxml-common@main"
pip install "pptx2svg[png] @ git+https://github.com/uvrt/pptx2svg@main"
pip install "pptx2svg-fonts @ git+https://github.com/uvrt/pptx2svg@main#subdirectory=packages/pptx2svg-fonts"
```

SVG output needs no extra. `png` adds the rasteriser. **Install `pptx2svg-fonts` for PNG
output** (the third line; about 11 MB): it is the fonts the PNG is drawn with --

- Carlito, Arimo, Tinos and Cousine, stand-ins with the advance widths of Calibri, Arial, Times
  New Roman and Courier New, so text is drawn at the widths it was laid out at; Caladea
  for Cambria;
- Noto Sans JP for Japanese and Chinese text (Meiryo, Yu Gothic, MS Gothic, SimSun...);
- Lato and Raleway, two faces template decks name.

All are under the SIL Open Font License 1.1, each with its licence text in the package
(`pptx2svg_fonts/licenses/`); the package's own code is MIT. No Microsoft font is in it.
Without it the PNG is drawn with whatever fonts the machine has, and text no installed font
can draw -- on a server with no CJK font, every Japanese, Chinese or Korean character -- is
left out: pptx2svg then warns (`glyphs-missing`, a `MissingGlyphsWarning`), once per face
and script. The other extras (`metafile`, `cairo`, `measure`) and how Office's own faces
are used on a Mac: [docs/usage.md](docs/usage.md#install-extras).

## Use

```python
from pptx2svg import convert_pptx_to_svg, convert_pptx_to_png, ConvertOptions

svgs = convert_pptx_to_svg("deck.pptx")                               # list[str], one per slide
pngs = convert_pptx_to_png("deck.pptx", ConvertOptions(width=1920))   # list[bytes]

options = ConvertOptions(slide_numbers=[3], width=1280)
svg = convert_pptx_to_svg("deck.pptx", options)[0]
for warning in options.warnings:     # anything unsupported is reported, not silently dropped
    print(warning)                   # [code] (slide N) message
```

```bash
pptx2svg deck.pptx -o out/ -f both -s 1,3-5 --width 1920
pptx2svg fonts --check deck.pptx     # exit 1 if this deck cannot be drawn faithfully
```

Also: the resolved model (`convert_pptx_to_model`) and a compact agent view for a language
model to read (`convert_pptx_to_agent_svg`) -- see [docs/usage.md](docs/usage.md).

## What is rendered

- **Shapes and text:** all 186 ECMA-376 preset geometries plus custom geometry, the full
  text inheritance cascade, bullets, wrapping (Latin and CJK), autofit, vertical text, tabs
  and columns; fills, outlines, shadows, glow, soft edges, pictures, groups, hyperlinks.
- **Tables** with merged cells and all 74 of PowerPoint's built-in table styles.
- **Charts:** every `c:*Chart` group ECMA-376 defines, with axes, gridlines, legend and data
  labels; 3-D scenes where measured, otherwise drawn flat with a warning.
- **SmartArt** from its cached drawing; **EMF/WMF** from the embedded preview.
- **Embedded fonts** (`ppt/fonts/*.fntdata`), for measurement and drawing, licence permitting.

**Not rendered:** 3-D effects, bevels and reflections (silently flat); SmartArt with no
cached drawing; vector EMF/WMF records; the ChartEx family (treemap, sunburst, waterfall,
...). Each but the first warns. Details: [docs/rendering.md](docs/rendering.md); what
closing each gap involves: [ROADMAP.md](ROADMAP.md).

## Status

Version 0.1.0, in active development, not yet released to PyPI. CI runs the suite on
Linux, macOS and Windows for Python 3.10 to 3.15. Fidelity against PowerPoint is measured
locally on a Mac with Office ([docs/fidelity.md](docs/fidelity.md)); CI checks regressions
only. Changes: [CHANGELOG.md](CHANGELOG.md).

## Documentation

- [docs/usage.md](docs/usage.md) -- install extras, API, CLI, resolved model, agent view, output format
- [docs/rendering.md](docs/rendering.md) -- the pipeline, what is and is not rendered
- [docs/fonts.md](docs/fonts.md) -- font summary, embedded fonts, escape hatches, Linux hosts
- [FONTS.md](FONTS.md) -- "the deck names font X: will it be drawn correctly?", end to end
- [docs/fidelity.md](docs/fidelity.md) -- scoring renders against PowerPoint's PDF export
- [docs/development.md](docs/development.md) -- running the tests, snapshots, generated tables
- [ROADMAP.md](ROADMAP.md) -- what is next and the measurements behind it
- [CONTRIBUTING.md](CONTRIBUTING.md), [CHANGELOG.md](CHANGELOG.md), [docs/licensing.md](docs/licensing.md)

## Family

- [pptx2svg](https://github.com/uvrt/pptx2svg) (this repo) -- renders PowerPoint (`.pptx`) slides to SVG and PNG.
- [docx2svg](https://github.com/uvrt/docx2svg) -- renders Word (`.docx`) documents to SVG, page by page.
- [ooxml-common](https://github.com/uvrt/ooxml-common) -- the format-neutral reading, DrawingML, fonts and text metrics both renderers share.
- [ooxml-edit](https://github.com/uvrt/ooxml-edit) -- lossless, undoable editing of OOXML packages, shared by both agent layers.
- [pptx-agent](https://github.com/uvrt/pptx-agent) -- an AI-editable PowerPoint layer: inspect, edit, re-render.
- [docx-agent](https://github.com/uvrt/docx-agent) -- an AI-editable Word layer: inspect, edit (optionally as tracked changes), re-render.

## Licence

MIT ([LICENSE](LICENSE)). Originally ported from
[pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse) (MIT, © Hiroki Sakabe); since
extended with geometry, charts, tables and text layout checked against PowerPoint's own PDF
export, PowerPoint's font-resolution and kerning rules, the shared
[ooxml-common](https://github.com/uvrt/ooxml-common) core, and an agent SVG view. The fonts
in `pptx2svg-fonts` are SIL OFL 1.1; the `.pptx` files in `tests/fixtures/` are third-party
documents not covered by the MIT grant; no Microsoft
font is committed anywhere. Details: [docs/licensing.md](docs/licensing.md).

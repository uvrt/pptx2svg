# pptx2svg

Render PowerPoint (`.pptx`) slides to **SVG**, and SVG to **PNG**.

Parsing, layout and SVG generation are pure Python with no dependencies beyond the
standard library. PNG output delegates to an existing rasteriser
([resvg](https://github.com/linebender/resvg) via `resvg-py`, which ships prebuilt
wheels — no compiler, no system libraries).

A Python port of [pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse).

## Install

```bash
pip install pptx2svg          # SVG only, zero dependencies
pip install pptx2svg[png]     # + PNG output via resvg-py
```

| Extra    | Pulls in    | For                                                      |
| -------- | ----------- | -------------------------------------------------------- |
| `png`    | `resvg-py`  | PNG output. Prebuilt wheels; the recommended backend.     |
| `cairo`  | `cairosvg`  | PNG via Cairo. Weaker SVG filter support (shadows, glows).|
| `fonts`  | `fonttools` | Measure with real installed fonts instead of the built-in tables. |

## Use

```python
from pptx2svg import convert_pptx_to_svg, convert_pptx_to_png, ConvertOptions

svgs = convert_pptx_to_svg("deck.pptx")                       # list[str], one per slide
pngs = convert_pptx_to_png("deck.pptx", ConvertOptions(width=1920))   # list[bytes]

# One slide, at a specific size
options = ConvertOptions(slide_numbers=[3], width=1280)
svg = convert_pptx_to_svg("deck.pptx", options)[0]

# Anything unsupported is reported rather than silently dropped
for warning in options.warnings:
    print(warning)
```

`convert_pptx_to_svg` accepts a path, `bytes`, or any binary file object.

Give only `width` or only `height` and the other follows the slide's aspect ratio.

### Command line

```bash
pptx2svg deck.pptx -o out/                    # one .svg per slide
pptx2svg deck.pptx -o out/ -f png --width 1920
pptx2svg deck.pptx -o out/ -f both -s 1,3-5   # selected slides, both formats
```

### The resolved model

To inspect slide structure instead of markup — useful for extracting text, finding
shapes, or building your own renderer:

```python
from pptx2svg import convert_pptx_to_model

resolved = convert_pptx_to_model("deck.pptx")
for slide in resolved.slides:
    for element in slide.elements:
        if getattr(element, "text_body", None):
            for paragraph in element.text_body.paragraphs:
                print("".join(run.text for run in paragraph.runs))
```

Everything in this model is already resolved: colours are concrete hex values with theme
lookups and `lumMod`/`tint`/`shade` applied, fonts are real typeface names with `+mn-lt`
expanded, and placeholder properties are merged down from layout and master.

## How it works

```
.pptx  ─▶  OPC package  ─▶  source model  ─▶  render model  ─▶  SVG  ─▶  PNG
           pptx2svg.opc    pptx2svg.parse   pptx2svg.resolve   .render  .png
```

| Module             | Responsibility                                                            |
| ------------------ | ------------------------------------------------------------------------- |
| `pptx2svg.opc`     | ZIP container, content types, relationship graph                          |
| `pptx2svg.parse`   | OOXML → source model, deliberately **unresolved** (theme refs, rel ids intact) |
| `pptx2svg.resolve` | Theme colours + the placeholder/background/text inheritance cascades       |
| `pptx2svg.render`  | Text measurement, line breaking, geometry, SVG output                     |
| `pptx2svg.text`    | Font metrics, measurement, wrapping                                       |
| `pptx2svg.png`     | Rasterisation via an external backend                                     |

Each stage is usable on its own.

The parse/resolve split is what makes inheritance work: a `<a:schemeClr val="tx1"/>`
means nothing until you know which colour map applies, and a placeholder's font size may
live on the slide, the layout, the master's `txStyles`, or the presentation's
`defaultTextStyle`. The parser records what the XML says; the resolver decides what it
means.

## Fidelity

The goal is accurate text, shapes and spatial layout — not a pixel-exact reproduction of
every PowerPoint feature.

**Supported:** shapes (134 preset geometries + full custom geometry with guide formulas),
text with the complete inheritance cascade, bullets and auto-numbering, line wrapping
(Latin and CJK), autofit (`normAutofit` shrink-to-fit and `spAutofit` grow), vertical text,
solid/gradient/pattern/image fills, outlines with dashes and arrowheads, shadows, glow and
soft edges, pictures with cropping and colour adjustments, tables with merged cells,
groups with nested coordinate spaces, hyperlinks, and alt text as `aria-label`.

**Not rendered** (reported as a warning, drawn as an empty positioned frame):

- **Charts** — the chart part holds data, not a drawing, so this needs a plotting engine
- **SmartArt** — PowerPoint caches a laid-out drawing we could read; not wired up yet
- **EMF/WMF images** — needs a metafile interpreter; a grey placeholder is drawn
- **3-D effects, bevels, reflections**
- **Table styles** — tables render with their explicit cell formatting only, so a table
  relying on a built-in style appears as an unstyled grid

Roughly 53 less-common preset shapes (action buttons, callout variants, curved arrows,
gears, scrolls) also fall back to rectangles.

See [ROADMAP.md](ROADMAP.md) for the plan to close these, with effort estimates and entry
points.

### Checking fidelity against PowerPoint

If Microsoft PowerPoint is installed, `tools/powerpoint_export_pdf.applescript` exports a
deck through PowerPoint itself, giving authoritative ground truth to compare against:

```bash
osascript tools/powerpoint_export_pdf.applescript "$PWD/deck.pptx" "$HOME/gt.pdf"
python -c "import pypdfium2 as p; d=p.PdfDocument('$HOME/gt.pdf'); \
           d[0].render(scale=2).to_pil().save('gt.png')"
```

Note that PowerPoint is sandboxed: both paths must be under your home directory, not
`/tmp`.

### Fonts

Text is measured against bundled metrics for four metric-compatible open fonts (Carlito
for Calibri, Liberation Sans for Arial, Liberation Serif for Times New Roman, Noto Sans JP
for Japanese), so layout is identical on every machine. The SVG asks for the original font
first, then the substitutes, so a host that really has Calibri uses it.

For layout that matches a specific machine's fonts, pass a measurer backed by real files:

```python
from pptx2svg import ConvertOptions, FontToolsTextMeasurer, convert_pptx_to_svg

measurer = FontToolsTextMeasurer({"Calibri": "/path/to/Calibri.ttf"})
svgs = convert_pptx_to_svg("deck.pptx", ConvertOptions(measurer=measurer))
```

For reproducible PNGs regardless of installed fonts:

```python
convert_pptx_to_png("deck.pptx", font_dirs=["./fonts"], skip_system_fonts=True)
```

Or map fonts to substitutes you do have:

```python
ConvertOptions(font_mapping={"Helvetica Neue": "Inter"})
```

## Output notes

The SVG uses **inline presentation attributes only** — no CSS classes, no `<style>`
elements. librsvg and resvg do not apply CSS selectors reliably, and this keeps output
portable across rasterisers. Element ids (`grad-1`, `patt-2`) come from a per-slide
counter, so the same input always produces byte-identical output; SVGs are diffable and
snapshot-testable.

## Development

```bash
pip install -e '.[dev]'
pytest
```

Tests run against real `.pptx` files in `tests/fixtures/` produced by PowerPoint, Google
Slides and python-pptx (shared with pptx-glimpse — see `tests/fixtures/README.md`).

## Licence

MIT. Ported from [pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse) (MIT,
© Hiroki Sakabe).

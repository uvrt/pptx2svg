# pptx2svg

[![CI](https://github.com/uvrt/pptx2svg/actions/workflows/ci.yml/badge.svg)](https://github.com/uvrt/pptx2svg/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://pypi.org/project/pptx2svg/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

Render PowerPoint (`.pptx`) slides to **SVG**, and SVG to **PNG**.

Parsing, layout and SVG generation are pure Python with no dependencies beyond the
standard library. PNG output delegates to an existing rasteriser
([resvg](https://github.com/linebender/resvg) via `resvg-py`, which ships prebuilt
wheels — no compiler, no system libraries).

Shape geometry, chart layout, table styles and font metrics are all **checked against
PowerPoint's own output** rather than guessed: `tools/fidelity.py` scores a render
against a PDF PowerPoint exported from the same deck, and the constants in this codebase
carry the measurement that produced them.

## Install

```bash
pip install pptx2svg                    # SVG only, zero dependencies
pip install 'pptx2svg[png,fonts]'       # PNG output, rendered reproducibly
```

| Extra      | Pulls in         | For                                                        |
| ---------- | ---------------- | ---------------------------------------------------------- |
| `png`      | `resvg-py`       | PNG output. Prebuilt wheels; the recommended backend.       |
| `fonts`    | `pptx2svg-fonts` | The typefaces we draw with. **Install this for PNG output** — see [Fonts](#fonts). |
| `metafile` | `pypdfium2`      | EMF/WMF pictures whose embedded preview is a PDF. Without it those draw a placeholder. |
| `cairo`    | `cairosvg`       | PNG via Cairo. Weaker SVG filter support (shadows, glows).  |
| `measure`  | `fonttools`      | Measure with real installed fonts instead of the built-in tables. |

SVG output needs nothing but the standard library — an SVG names fonts, it does not
embed them. PNG output *rasterises*, so it needs actual font files, and without them the
rasteriser quietly substitutes whatever the host has. `[fonts]` is 11 MB and is what
makes the same deck produce the same pixels on your laptop and on a build server.

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
    print(warning)          # [code] (slide N) message
```

`convert_pptx_to_svg` accepts a path, `bytes`, or any binary file object, and all three
produce byte-identical output.

Give only `width` or only `height` and the other follows the slide's aspect ratio; give
neither and the slide's own size is used.

`options.warnings` is the channel for everything the renderer could not do faithfully.
The codes are stable strings — `chart-unsupported-type`, `diagram-no-cached-drawing`,
`metafile-image`, `metafile-rasterizer-missing`, `table-style-unknown`, `font-substituted`,
`font-bundle-missing` and a handful more — so a build can fail on the ones it cares about
and ignore the rest.

### Command line

```bash
pptx2svg deck.pptx -o out/                    # one .svg per slide
pptx2svg deck.pptx -o out/ -f png --width 1920
pptx2svg deck.pptx -o out/ -f both -s 1,3-5   # selected slides, both formats
```

`--height`, `--backend {auto,resvg,cairosvg}`, `--font-dir DIR`, `--system-fonts`,
`--no-bundled-fonts` and `-q` are the rest of it; `pptx2svg --help` lists them.

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
expanded, and placeholder properties are merged down from layout and master. A chart
carries its parsed series and categories alongside the primitives it draws with, so you
can read the numbers without touching the chart XML.

## How it works

```
.pptx  ─▶  OPC package  ─▶  source model  ─▶  render model  ─▶  SVG  ─▶  PNG
           pptx2svg.opc    pptx2svg.parse   pptx2svg.resolve   .render  .png
```

| Module              | Responsibility                                                           |
| ------------------- | ------------------------------------------------------------------------ |
| `pptx2svg.opc`      | ZIP container, content types, relationship graph                         |
| `pptx2svg.parse`    | OOXML → source model, deliberately **unresolved** (theme refs, rel ids intact) |
| `pptx2svg.resolve`  | Theme colours + the placeholder/background/text inheritance cascades      |
| `pptx2svg.render`   | Text measurement, line breaking, geometry, SVG output                    |
| `pptx2svg.text`     | Font metrics, measurement, wrapping                                      |
| `pptx2svg.fonts`    | Which faces we can draw, and what happens when we cannot                 |
| `pptx2svg.metafile` | EMF/WMF preview extraction                                               |
| `pptx2svg.png`      | Rasterisation via an external backend                                    |

Each stage is usable on its own.

The parse/resolve split is what makes inheritance work: a `<a:schemeClr val="tx1"/>`
means nothing until you know which colour map applies, and a placeholder's font size may
live on the slide, the layout, the master's `txStyles`, or the presentation's
`defaultTextStyle`. The parser records what the XML says; the resolver decides what it
means.

## What gets rendered

The goal is accurate text, shapes and spatial layout — not a pixel-exact reproduction of
every PowerPoint feature.

**Shapes and text.** All 186 preset geometries in ECMA-376, plus `lineInv`, plus full
custom geometry with guide formulas — every one of them verified against PowerPoint's own
PDF export at two aspect ratios. Text carries the complete inheritance cascade, bullets
and auto-numbering, line wrapping (Latin and CJK), autofit (`normAutofit` shrink-to-fit
and `spAutofit` grow), vertical text, tabs and columns. Solid, gradient, pattern and
image fills; outlines with dashes and arrowheads; shadows, glow and soft edges; pictures
with cropping, tiling and colour adjustments; groups with nested coordinate spaces;
hyperlinks; and alt text as `aria-label`.

**Tables** render with merged cells, borders and fills, and with **PowerPoint's built-in
table styles** — all 74 of them, measured out of PowerPoint itself, because it keeps
those definitions inside the application and never writes them into the file. A table
that names one therefore renders banded and headed rather than as a bare grid.

**Charts** — `barChart`, `lineChart`, `pieChart`, `doughnutChart` and `radarChart` — are
read and drawn with their axis range, tick interval, gridlines, legend and data labels,
laid out from constants measured out of PowerPoint's PDF export. The 3-D spellings
(`bar3DChart` and friends) draw flat. No deck in the test corpus warns
`chart-unsupported-type`.

**SmartArt** renders from the DrawingML drawing PowerPoint caches beside the diagram —
shapes, text, fills and geometry, each label placed by its own `dsp:txXfrm`. Where that
cache is missing or empty there is nothing to draw, and the renderer says so rather than
leaving a blank rectangle unexplained.

**EMF/WMF pictures** render the preview Office embeds in the metafile. A DIB preview
needs nothing extra; a PDF preview needs `pptx2svg[metafile]`.

### Not rendered

Four things, and each of them is a real gap rather than a rough edge:

- **3-D effects, bevels and reflections.** `a:scene3d`, `a:sp3d` and `a:reflection` are
  ignored; the shape draws flat. This is the one item on the list that passes silently.
- **SmartArt with no cached drawing.** There is no diagram layout engine here, and
  implementing `dgm:layoutDef` is a project in its own right. Such a frame draws nothing
  and warns `diagram-no-cached-drawing`. This is common in older files: of 46 real
  SmartArt decks measured, 13 carried a drawing with shapes in it and 33 did not.
- **Vector EMF/WMF content.** Only the embedded preview is read; the metafile's own
  drawing records are not interpreted, so a metafile without a preview draws a
  placeholder and warns. `ConvertOptions(metafile_converter=...)` is the hook for
  shelling out to Inkscape or `libemf2svg` if you need the vectors.
- **Chart types beyond the five above** — `areaChart`, `scatterChart`, `bubbleChart`,
  `stockChart`, `surfaceChart`, `ofPieChart`. These warn `chart-unsupported-type` and
  draw an empty positioned frame. A combo chart draws whichever of its groups is a type
  we know and ignores the others.

See [ROADMAP.md](ROADMAP.md) for what closing each of these involves.

### Checking fidelity against PowerPoint

`tools/fidelity.py` scores our render against PowerPoint's own PDF export on SSIM and
colour-histogram correlation. It renders **our** side with the same licensed Microsoft
faces PowerPoint used, read in place from wherever Office installed them, so a difference
between the two images is attributable to this library rather than to font availability:

```bash
python3 tools/fidelity.py --write-profile               # once, on a machine with Office
python3 tools/fidelity.py --oracle ~/pptx2svg-oracle
```

The profile records paths and hashes only; no licensed font is ever copied into the
repository, and `tests/font-profile.local.json` is gitignored. Without it the harness
refuses to run, and a deck naming a face PowerPoint did not have either is skipped rather
than scored against Microsoft's fallback.

If Microsoft PowerPoint is installed, `tools/powerpoint_export_pdf.applescript` exports a
deck through PowerPoint itself, giving authoritative ground truth to compare against:

```bash
osascript tools/powerpoint_export_pdf.applescript "$PWD/deck.pptx" "$HOME/gt.pdf"
python -c "import pypdfium2 as p; d=p.PdfDocument('$HOME/gt.pdf'); \
           d[0].render(scale=2).to_pil().save('gt.png')"
```

PowerPoint is sandboxed, and being under your home directory is not enough: both paths
must be in a directory PowerPoint has already been granted access to. A brand-new one is
refused, and the failure looks like a corrupt deck rather than a permissions problem.

## Fonts

Office's typefaces are proprietary and cannot be redistributed. Layout is therefore
computed from the advance widths of open fonts built to match them, and **those same
files are what the rasteriser draws with** — measuring with one face and drawing with
another is the failure this whole subsystem exists to prevent. resvg does not warn when
it substitutes: rendering one string in Calibri, Carlito, Aptos, Noto Sans JP and Lato on
a machine with none of them produces five byte-identical PNGs.

| Office face        | Drawn with     | Debian package              | Licence     |
| ------------------ | -------------- | --------------------------- | ----------- |
| Calibri            | Carlito        | `fonts-crosextra-carlito`   | SIL OFL 1.1 |
| Arial, Helvetica   | Arimo          | `fonts-croscore`            | SIL OFL 1.1 |
| Times New Roman    | Tinos          | `fonts-croscore`            | SIL OFL 1.1 |
| Courier New        | Cousine        | `fonts-croscore`            | SIL OFL 1.1 |
| Cambria            | Caladea †      | `fonts-crosextra-caladea`   | SIL OFL 1.1 |
| Aptos †            | Carlito †      | —                           | —           |
| Japanese Gothic    | Noto Sans JP † | `fonts-noto-cjk`            | SIL OFL 1.1 |
| Lato, Raleway      | themselves     | `fonts-lato`, — ‡           | SIL OFL 1.1 |

The first four are exact: measured character by character against the copies Office
installs, the advance widths match to the unit, so substituting them changes glyph shapes
and nothing else — not one line break moves.

† **Approximate, and reported as such.** Caladea is universally described as
metric-compatible with Cambria; measured, it runs 4.5 % narrow. **Aptos**, Microsoft's
Office default since 2023, has no open clone at all — we measure it with its own widths
(so line breaks match PowerPoint) and draw it with Carlito, whose widths sit closest of
anything shippable, −3.6 % on a representative sentence. Noto Sans JP is not
metric-compatible with MS Gothic or Meiryo either; nothing is, and a Gothic standing in
for a Gothic beats a Latin fallback.

‡ Debian does not package Raleway. Use the pip bundle, or Google Fonts.

### Will my deck render faithfully?

```bash
pptx2svg fonts                       # what the bundle covers
pptx2svg fonts --check deck.pptx     # exit 1 if this deck cannot be drawn faithfully
```

```
mode:   bundled from .../site-packages/pptx2svg_fonts/files
        this machine's own fonts are ignored, so output is reproducible
        families: Arimo, Caladea, Carlito, Cousine, Lato, Noto Sans JP, Raleway, Tinos

face                       verdict      drawn with     note
Arial                      compatible   Arimo          Arimo has Arial's advance widths
Aptos                      approximate  Carlito        measured as Aptos, drawn as Carlito; no metric-compatible clone exists
Aptos Display              approximate  Carlito        measured as Aptos Display, drawn as Carlito; no metric-compatible clone exists

2 of 3 faces will not be drawn at the widths they were measured at.
```

`exact` and `compatible` are faithful; `approximate` and `missing` are not, and `--check`
exits non-zero on them so a deck that cannot be rendered faithfully fails a build instead
of shipping wrong pixels. The same information reaches library callers as
`font-substituted` warnings on `ConvertOptions.warnings`.

### Without the bundle

`pptx2svg` alone renders PNGs with the host's fonts and emits one
`font-bundle-missing` warning saying so. That is a deliberate downgrade, not a silent
one — output is then whatever the machine happens to have installed.

### Escape hatches

```python
# The host's fonts as well as the bundle, for a deck naming a face we do not carry
convert_pptx_to_png("deck.pptx", skip_system_fonts=False)

# Your own faces only
convert_pptx_to_png("deck.pptx", font_dirs=["./corporate-fonts"], use_bundled_fonts=False)

# Layout measured from specific files on this machine
from pptx2svg import ConvertOptions, FontToolsTextMeasurer
measurer = FontToolsTextMeasurer({"Calibri": "/path/to/Calibri.ttf"})
convert_pptx_to_svg("deck.pptx", ConvertOptions(measurer=measurer))

# Map a face to a substitute you do have
ConvertOptions(font_mapping={"Helvetica Neue": "Inter"})
```

On the command line: `--system-fonts`, `--no-bundled-fonts`, `--font-dir DIR`.

### Debian and other Linux hosts

Two routes, and only one of them is reproducible.

**The bundle (recommended).** `pip install 'pptx2svg[fonts]'` pins the exact font files.
Two machines running the same version produce the same pixels, which is the only way to
get that guarantee.

**System packages.** `tools/install-fonts-debian.sh` installs the same designs from apt
(`fonts-croscore`, `fonts-crosextra-carlito`, `fonts-crosextra-caladea`,
`fonts-liberation2`, `fonts-noto-cjk`, `fonts-lato`) and verifies each family with
`fc-list`. Render with `--system-fonts` to use them. This is fine for correct *layout* —
Liberation Sans/Serif/Mono are derived from Arimo/Tinos/Cousine and measure identically,
checked character by character — but apt gives you whatever version the distribution
shipped, so two machines on different releases can still differ.

The same script fetches the faces no bundle can legally contain, behind explicit flags:

```bash
tools/install-fonts-debian.sh                  # open substitutes from apt
tools/install-fonts-debian.sh --aptos          # + Aptos, from Microsoft's own download
tools/install-fonts-debian.sh --ppviewer       # + Calibri, Cambria and the ClearType set
```

`--ppviewer` extracts them from the PowerPoint Viewer installer, the route documented at
[wiki.debian.org/ppviewerFonts](https://wiki.debian.org/ppviewerFonts) (hash-pinned;
needs `cabextract`). **These are non-free Microsoft fonts — you must already hold a
licence to use them.** Nothing extracted is ever committed or shipped.

## Output notes

The SVG uses **inline presentation attributes only** — no CSS classes, no `<style>`
elements. librsvg and resvg do not apply CSS selectors reliably, and this keeps output
portable across rasterisers. Element ids (`grad-1`, `patt-2`) come from a per-slide
counter, so the same input always produces byte-identical output; SVGs are diffable and
snapshot-testable.

Every shape group also carries its source identity, so rendered output can be mapped back to
the deck it came from:

```xml
<g role="img" aria-label="Contract picture" data-pptx-id="256.3" data-pptx-path="2"
   transform="translate(346.457, 146.982)"> ... </g>
```

`data-pptx-id` is `"<sldId>.<cNvPr id>"` for slide shapes and `lay:`/`mst:` for shapes
inherited from the layout or master. `data-pptx-path` is the index path through the shape
tree, and it is **not** redundant — `cNvPr@id` is not unique in real decks, so the pair is
what addresses a shape unambiguously. Alt text is still emitted as `aria-label` alongside.

## Development

```bash
pip install -e '.[dev]'
pip install -e packages/pptx2svg-fonts
pytest
```

`pptx2svg-fonts` is a sibling distribution in this repository and is not on PyPI yet, so
it is installed from the checkout rather than named as a dependency.

Tests run against real `.pptx` files in `tests/fixtures/` produced by PowerPoint, Google
Slides and python-pptx; `tests/fixtures/README.md` records where each came from.

Decks that are not ours to redistribute are not committed at all. `tests/conftest.py`
looks for those in the gitignored `scratch/`, and the tests that need one skip where it
is absent.

The metrics table in `pptx2svg/text/metrics.py` is **generated** from the font files in
`packages/pptx2svg-fonts`. Do not edit it by hand:

```bash
python3 tools/extract_font_metrics.py --check    # fails if the table has drifted
python3 tools/extract_font_metrics.py --write    # regenerate
```

A test runs `--check`, so a font update that is not accompanied by a regenerated table
fails the suite rather than silently making layout wrong. The same idea applies to the
two other generated tables: `render/preset_specs.py` comes from
`tools/derive_preset_geometry.py`, and `parse/table_styles_builtin.py` from
`tools/derive_table_styles.py`, which measures the styles by rendering them through
PowerPoint. Edit the tool, not the table.

## Licence

MIT. Ported from [pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse) (MIT,
© Hiroki Sakabe).

The `pptx2svg-fonts` distribution is a separate matter: its Python module is MIT, but the
font files it carries are each **SIL Open Font License 1.1**, redistributed byte-for-byte
as published on [Google Fonts](https://github.com/google/fonts) with their full licence
texts alongside them in `src/pptx2svg_fonts/licenses/`. Carlito, Caladea, Arimo, Tinos and
Cousine are © their respective Project Authors; Lato is © tyPoland Łukasz Dziedzic; Noto
Sans JP is © Adobe; Raleway is © the Raleway Project Authors. None has been subsetted,
renamed or otherwise modified, so the Reserved Font Name clauses are satisfied.

No Microsoft font is bundled, vendored or committed anywhere in this repository.

Neither is ECMA-376 itself. `pptx2svg/render/preset_specs.py` is *compiled* from the
standard's `presetShapeDefinitions.xml`, but that source is not redistributed here --
`tools/derive_preset_geometry.py` requires a copy you obtained yourself and verifies its
SHA-256 before reading it.

**The `.pptx` files in `tests/fixtures/` are not covered by the MIT grant above.** They are
third-party documents included as renderer inputs, and they remain their owners'. They came
from pptx-glimpse's own `shared-fixtures/` (MIT), so that rendering differences between the
two implementations can be compared directly; their provenance is recorded in
`tests/fixtures/FIXTURES-README.md`.

Decks that are *not* ours to redistribute are not committed at all -- not even as a test
input. `tests/conftest.py` looks for those in the gitignored `scratch/`, and the tests that
need one skip where it is absent. `real-college-template.pptx`, Dickinson College's public
sample deck, is the only such deck today; `tests/fixtures/README.md` says where to get it.

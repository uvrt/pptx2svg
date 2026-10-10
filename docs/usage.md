# Using pptx2svg

The API, the command line, the resolved model, the agent view and what the SVG output looks like. Back to the [README](../README.md).

## Install extras

| Extra      | Pulls in         | For                                                        |
| ---------- | ---------------- | ---------------------------------------------------------- |
| `png`      | `resvg-py`       | PNG output. Prebuilt wheels; the recommended backend.       |
| `fonts`    | `pptx2svg-fonts` | The typefaces we draw with. **Install this for PNG output** — see [Fonts](fonts.md). |
| `metafile` | `pypdfium2`      | EMF/WMF pictures whose embedded preview is a PDF. Without it those draw a placeholder. |
| `cairo`    | `cairosvg`       | PNG via Cairo. Weaker SVG filter support (shadows, glows).  |
| `measure`  | `fonttools`      | Measure with real installed fonts instead of the built-in tables. |

SVG output needs nothing but the standard library — an SVG names fonts, it does not
embed them. PNG output *rasterises*, so it needs actual font files, and without them the
rasteriser quietly substitutes whatever the host has. `[fonts]` is 11 MB and is what
makes the same deck produce the same pixels on your laptop and on a build server —
on a laptop with Office, with `host_fonts=False` (below).

**On a Mac with Office, the faces PowerPoint itself draws are used by default**, with or
without `[fonts]`: Aptos and the rest from PowerPoint's application bundle, Aptos Display
from Office's cloud-font cache, read where they are installed and never copied, for
measurement as well as drawing — because that most closely resembles PowerPoint. **The
output then depends on the machine.** For the reproducible render, pass
`ConvertOptions(host_fonts=False)` (or `svg_to_png(..., host_fonts=False)`), or set
`PPTX2SVG_OFFICE_FONTS=0`; with `[fonts]` installed that gives the same pixels everywhere.
Where Office's folders do not exist (Linux, Windows, CI) nothing changes
([FONTS.md](../FONTS.md#a-mac-with-powerpoint-the-faces-powerpoint-draws-read-in-place)).

**Your own font folder** -- licensed faces kept where the system does not look -- is
`ConvertOptions(font_dirs=[...])` (or `--font-dir`), or the environment variable
**`OOXML_FONT_DIRS`** (folders separated by `os.pathsep`), which docx2svg reads too. The
explicit argument wins over the variable; either is added to the system's folders and
searched before them. Its faces are measured from and drawn with
([FONTS.md](../FONTS.md#your-own-folder-measured-and-drawn)).

## Python API

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
The codes are stable strings — `chart-unsupported-type`, `chart-3d-flattened`,
`diagram-no-cached-drawing`, `metafile-image`, `metafile-rasterizer-missing`,
`table-style-unknown`, `font-substituted`, `font-bundle-missing` and a handful more — so a
build can fail on the ones it cares about and ignore the rest. "Nothing was drawn" and
"something simplified was drawn" are deliberately different codes: `chart-unsupported-type`
is the first and `chart-3d-flattened` the second.

**Slide numbers and dates are evaluated per slide**, as PowerPoint does: an `a:fld` in a
slide's own placeholders or in a plain text box on the layout or master (a footer's
"‹#› | Copyright …") shows the slide's number (counting from `firstSlideNum`) and the
date, rather than the text cached in the file. The date is the clock at conversion, in the
field's language and format (`datetime1` … `datetime13`, measured for en-US, en-GB, nl-NL,
de-DE and fr-FR; any other language is drawn in en-US's formats with a
`field-date-format` warning). Output holding a date field therefore changes from day to
day; pin the clock for anything that must be reproducible -- tests, goldens, a cache:

```python
from datetime import datetime
convert_pptx_to_svg("deck.pptx", ConvertOptions(now=datetime(2026, 1, 31, 9, 0)))
```

or `--now 2026-01-31T09:00` on the command line.

`glyphs-missing` (PNG output) is text the rasteriser will not draw: no font it loads answers
to the text's `font-family`, or none has its characters -- Japanese rendered without
`pptx2svg-fonts` on a host with no CJK font, Korean with the bundle alone (Noto Sans JP has
no Hangul). It comes once per face and script, on the first slide that has it, and as a
`MissingGlyphsWarning` through Python's `warnings` too; `svg_to_png` raises only the latter.
`ConvertOptions(check_glyphs=False)` (or `svg_to_png(..., check_glyphs=False)`) skips the
check, which reads the fonts' `name` and `cmap` tables -- the host's only when the bundle
and the files passed in leave something unfound.


## Command line

```bash
pptx2svg deck.pptx -o out/                    # one .svg per slide
pptx2svg deck.pptx -o out/ -f png --width 1920
pptx2svg deck.pptx -o out/ -f both -s 1,3-5   # selected slides, both formats
```

`--height`, `--backend {auto,resvg,cairosvg}`, `--font-dir DIR`, `--system-fonts`,
`--no-bundled-fonts` and `-q` are the rest of it; `pptx2svg --help` lists them.

## The resolved model

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

A table's rows are laid out when it is drawn: `a:tr@h` is only a minimum, and PowerPoint
grows a row to fit its text and moves the rows below it down. `table_row_heights(table,
context)` is each row's height, EMU, as the renderer draws it -- for a caller that checks
whether a table runs off the slide without rendering it (pptx-agent's `overflows()` does).

## The agent view

For a language model to *read* a slide — its shapes, their boxes, colours and text, with
the ids an editor addresses them by — rather than to look at it:

```python
from pptx2svg import ConvertOptions, convert_pptx_to_agent_svg

svgs = convert_pptx_to_agent_svg("deck.pptx", ConvertOptions(slide_numbers=[3]))
```

Each slide is a compact SVG in **slide points** (`viewBox="0 0 960 540"` at 16:9), every
shape carrying `data-pptx-id` as the normal render does, numbers rounded to a tenth of a
point, and theme colours named beside the hex:

```xml
<g data-pptx-id="258.8"><rect x="142" y="125.3" width="99.7" height="21.6"
   fill="#DAE8F9" data-fill="dk2 lumMod=10% lumOff=90%" stroke="#022A2F"
   data-stroke="accent1 shade=15%" stroke-width="1.5"/>
 <text x="149.2" y="128.9" data-anchor="middle"><tspan x="191.9" dy="11"
   text-anchor="middle" font-size="11" font-weight="bold" fill="#0B2545"
   data-fill="dk2">Diagnose</tspan></text></g>
```

A preset is `data-preset` on the shape's box (an ellipse is an `<ellipse>`, a line a
`<line>`); text is one `<tspan>` per paragraph and run, unwrapped, with no fonts read or
embedded; pictures, charts, tables, SmartArt and media are placeholders
(`<rect data-kind="picture" …/>`; a table with its cell text, a chart with its type and
title, SmartArt with its nodes' text). The layout's and master's own shapes come first,
marked `data-layer="layout"`/`"master"`. There are no images, base64, fonts, filters or
`<defs>`, so a shape-heavy slide is a few thousand tokens, typically well under the normal
render — and far under it when the deck embeds pictures or fonts. It is a read view: it
is not meant to look like the slide, and nothing reads it back.

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


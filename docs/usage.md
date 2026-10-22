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

**A layout or master on its own** (`ConvertOptions.templates`, part paths), drawn after
the slides `slide_numbers` picks -- `slide_numbers=[]` for none:

```python
options = ConvertOptions(slide_numbers=[], templates=[
    "ppt/slideLayouts/slideLayout2.xml", "ppt/slideMasters/slideMaster1.xml"])
layout_svg, master_svg = convert_pptx_to_svg("template.potx", options)
```

A layout's view is what a slide on it starts as -- the background, the master's shapes
unless the layout's `showMasterSp` hides them, the layout's own shapes -- with each
placeholder drawn as a wireframe: a grey dashed box, its prompt in the text style a
slide's text there inherits, and a tag on its top edge with its type and `idx` (`title`,
`content 1`, `dt 10`); a banner in the top right corner names the layout. A master's
view is the same with its own placeholders. The SVG's root carries `data-template`
(`layout`/`master`), `data-template-part` and `data-template-name`, and the model's
`Slide.template` says the same, with `slide_number` 0. The wireframe is this library's
drawing, not PowerPoint's: PowerPoint draws no empty placeholder outside editing, and how
Slide Master view draws one has not been measured. A slide-number field shows its own
text (`‹#›`); a part that is no layout or master of the deck raises `ValueError`.

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
de-DE and fr-FR). A field with no `lang`, or one in any other language, is drawn in the
language of the conversion's locale (below) when that is one of the five, else in
en-US's, with a `field-date-format` warning. Output holding a date field therefore changes
from day to day; pin the clock for anything that must be reproducible -- tests, goldens, a cache:

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

### Rendering in two halves

A caller that keeps images by their SVG -- a render tool drawing only the slides an edit
changed -- writes the SVGs first and rasterises those it lacks:

```python
import tempfile
from pptx2svg import ConvertOptions, render_svgs, rasterise_svgs

options = ConvertOptions(width=1280)
svgs, resolved = render_svgs("deck.pptx", options)       # convert_pptx_to_svg's, and the model
with tempfile.TemporaryDirectory() as folder:
    embedded = resolved.embedded_fonts.write(folder) if resolved.embedded_fonts else None
    numbers = [(slide.slide_number, slide.template_part) for slide in resolved.slides]
    pngs = rasterise_svgs(svgs[:2], options, numbers=numbers[:2], font_files=embedded)
```

`rasterise_svgs` draws what `convert_pptx_to_png` draws for the same documents, byte for
byte; its `font_dirs` default to the conversion's (`user_font_dirs(options)`) and its
`host_fonts` to `uses_host_fonts(options)`, and `glyphs-missing` goes to `options.warnings`.

### A live, editable view

For a page that shows each slide as SVG, re-renders a slide after an edit and sends only
the shapes that changed, four opt-in options; with none of them the output is unchanged.

```python
from pptx2svg import ConvertOptions, render_svgs, rasterise_svgs, slide_groups

images: dict[str, bytes] = {}                     # filled in place, as `warnings` is
options = ConvertOptions(slide_numbers=[3], images=images, edit_marks=True, hyperlinks=False)
svgs, resolved = render_svgs("deck.pptx", options)
# svgs[0]: <svg data-pptx-size="960x540" ...><image href="img/3fa1c9....png" .../>...
# images:  {"img/3fa1c9....png": b"\x89PNG...", ...}
view = slide_groups(svgs[0])
# {"root": "<svg ...>", "order": ["background", "258.2", "258.7", ...],
#  "groups": {"258.7": "<g data-pptx-id=\"258.7\" ...>...</g>", ...},
#  "defs":   {"258.7": "<linearGradient id=\"grad-s258_7-1\" ...>...", ...}}
pngs = rasterise_svgs(svgs, options, images=images)  # the embedded form's PNG, byte for byte
```

**`images={}`: pictures by reference.** Every picture the SVG would carry as a `data:`
URI -- a picture, a picture fill, the background, a picture bullet, a metafile's drawing or
preview -- is written once as `href="img/<hash>.<format>"`: the first 32 hex digits of the
SHA-256 of its bytes and the format the bytes are (`png`, `jpg`, `gif`, `bmp`, `tif`,
`webp`, `svg`; else the media type's subtype). Its bytes go into the dict under that name,
so the same picture on several slides, or across renders sharing one dict, is one entry.
`image_prefix` sets the folder (`"img/"`). Nothing else is embedded in the SVG: fonts are
named, never embedded; gradients, patterns, masks, clip paths and filters are markup in
`<defs>`; what a referenced SVG picture (a metafile's drawing) embeds stays inside that
file. `embed_images(svg, images)` puts the pictures back exactly -- the plain render's
bytes -- and `rasterise_svgs(..., images=images)` does so before drawing, so the PNG is
the embedded one's; a name the dict lacks raises `KeyError` rather than drawing nothing.
On the fixture corpus a picture-heavy slide goes from 151,399 bytes to 3,529; a template
slide with full-bleed photographs from 2,453,564 to 4,908.

**`edit_marks=True`: the id contract.** Attributes and attribute-only groups are added
and definitions renamed; the drawing does not change (the PNG is the plain one's, byte for
byte, on every slide of the corpus):

- `data-pptx-id` is on every drawn shape: each top-level shape, each member of a group,
  each graphic frame (table, chart, SmartArt) and each piece of a SmartArt drawing
  (`<frame>/<drawing>/<shape>`). It is there without `edit_marks` too. pptx2svg writes
  `<sldId>.<cNvPr id>` for a slide's shapes, `lay:<id>`/`mst:<id>` for its layout's and
  master's. **pptx-agent's tools address shapes by their own ids** (a stamp, PowerPoint's
  `a16:creationId`, or `cNvPr@id` with the tree path where it is not unique: `257.c8060F849-…`,
  `257.3#5`); its `Slide.render_svg(**options)` takes these options and rewrites the ids
  to those, cells included. Use that render where ids go back to the tools.
- Each table cell drawn is a `<g data-pptx-id="<table id>/cell<row>,<column>">` holding its
  fill (a box drawn as nothing where it has none) and its text: grid positions from 0, as
  pptx-agent addresses cells; a position a merge hides is not drawn.
- `data-pptx-source` beside every `data-pptx-id`: `slide`, `layout` or `master`.
- `data-pptx-text` on the element holding a text body: `shape` (a shape's text frame),
  `cell` (a cell's), `diagram` (a SmartArt drawing's piece, whose text pptx-agent edits by
  node, not by this id). The text's address is the nearest `data-pptx-id` at or above it.
  A chart's text, which no id owns, is not marked.
- `data-pptx-size="960x540"` on the root: the slide's size in points.
- `data-pptx-path` is left out, and every shape's definitions get ids of their own
  (`grad-s258_7-1`, the background's `grad-bg-1`): neither then depends on what is drawn
  before the shape.

An id is the shape's for as long as the shape exists: editing its text, moving or resizing
it, restyling it, or adding and removing other shapes leaves it as it was. pptx-agent's id
for a shape whose `cNvPr@id` is duplicated carries its tree path until it is first edited
(and stamped), so adding or removing a shape before it can change that id.

Cost: a few dozen bytes a shape and about 60 a table cell -- 0.2% to 13.5% of a slide on
the fixture corpus, 3.3% at the median (`table test`, 25 cells, 12,956 to 14,702 bytes; a
42-shape slide 14,676 to 15,881), less where pictures are embedded. Small, but not
nothing for a renderer whose default output is held byte for byte by snapshots and read by
other tools, so the marks are opt-in.

**`slide_groups(svg)`: the slide by shape.** Each top-level shape's markup by its key, in
drawing order (the background first, as `"background"`), and the definitions each uses,
transitively. The key is the `data-pptx-id` (an id drawn twice gets `#2`, `#3`...). The
definitions stay in the document's one `<defs>`: a page swapping a group swaps its
definitions there too. Rendered with `edit_marks`, a shape that did not change has
byte-identical markup and definitions from render to render; without it, definitions are
numbered across the slide, and every shape after a changed one may differ. Measured
through pptx-agent on the corpus' slides with three or more shapes: an edit of one
shape's text, a move of another and a gradient on a third left 226 of the 227 other
groups byte-identical with `edit_marks` (the last was a connector glued to the moved
shape, which pptx-agent rerouted) and changed 18 without; removing one shape and adding
one at the back changed 6 of 268 (the shapes whose duplicated `cNvPr@id` gives them a
path in pptx-agent's id) with `edit_marks`, 243 without. It works on pptx-agent's
rewritten SVG too. A pass is string work on the renderer's own output: about 0.3 ms for
a slide with its pictures by reference, 26 ms for one embedding 2.4 MB of them.

**`hyperlinks=False`**: no `<a href>` around a shape or a run; what it covered is drawn as
before.

### Locale and clock

PowerPoint takes a few things from the machine it runs on rather than from the file: the
decimal and thousands separators of a chart's numbers (a data label in `#,##0.0` is drawn
`1,240.0` by default, `1.240,0` in a locale with a decimal comma) and the language of a
date field that states none. A renderer that read them from *its* machine would draw the
same deck differently on every server, so pptx2svg never does: the caller says which
locale a conversion is in, and without one it is en-US.

```python
from ooxml_common.locales import Locale

nl = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
            language="nl-NL", list_separator=";")
convert_pptx_to_svg("deck.pptx", ConvertOptions(locale=nl))
```

`Locale` is frozen and picklable, so the options travel into worker processes unchanged.
A field's own `lang` still wins over the locale's language.

The time zone is not part of the locale. `now` left to the default is the current time in
the **server's local time zone** (`datetime.now()`); a service drawing for users elsewhere
passes `now` as an aware datetime in the user's zone
(`datetime.now(ZoneInfo("Europe/Brussels"))`), or a fixed value for reproducible output.


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
The ids as a contract for an editor -- cells, sources, text bodies, what stays stable --
are under "A live, editable view" above.


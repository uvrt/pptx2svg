# Fonts: will my deck draw correctly?

One question, answered in one place: **the deck names font X — will pptx2svg draw it at
the widths PowerPoint laid it out at, and if not, what do I do?**

The invariant everything below serves is short. Layout — wrapping, autofit shrinking,
vertical centring — is computed from a table of advance widths *before* the SVG is handed
to a rasteriser. If the face that is finally drawn has different advance widths from the
face that was measured, every line break is in the wrong place and nothing says so. resvg
does not warn when it substitutes: rendering one string in Calibri, Carlito, Aptos, Noto
Sans JP and Lato on a machine with none of them produces five byte-identical PNGs.

So there is exactly one thing to know about your font: **is it measured and drawn with the
same widths?**

Skip to the answer: [find your font](#find-your-font) · [ask the tool](#ask-the-tool-pptx2svg-fonts---check)
· [Aptos](#aptos) · [what to do about a font we cannot draw](#the-escape-hatches-and-their-trap)

## Find your font

| The deck names… | How it is drawn | Grade | What to do |
| --- | --- | --- | --- |
| **Anything the deck embeds** (`<p:embeddedFontLst>`) | Measured *and* drawn from the deck's own file | `exact` | Nothing. This is the best outcome there is — [why](#1-the-deck-carries-it-best) |
| **Arimo, Caladea, Carlito, Cousine, Lato, Noto Sans JP, Raleway, Tinos** | The bundle ships that exact family | `exact` | `pip install 'pptx2svg[fonts]'` |
| **Calibri, Calibri Light** | Carlito | `compatible` | Nothing. Same advance widths; only outlines change |
| **Arial, Helvetica** | Arimo | `compatible` | Nothing |
| **Times New Roman, Times** | Tinos | `compatible` | Nothing |
| **Courier New, Courier** | Cousine | `compatible` | Nothing |
| **Liberation Sans / Serif / Mono** | Arimo / Tinos / Cousine | `compatible` | Nothing. Same designs under another name |
| **Cambria** | Measured as Cambria, drawn with Caladea | `approximate` | Caladea is **not** metric-compatible — [why this row is the important one](#3-the-bundle-supplies-a-metric-compatible-clone) |
| **Aptos, Aptos Display, Aptos Narrow** | Measured as Aptos, drawn with Carlito | `approximate` | No clone exists anywhere. Install real Aptos — [Aptos](#aptos) |
| **MS Gothic, MS Mincho, MS PGothic, MS PMincho** | Measured from their own tables, drawn with Noto Sans JP | `approximate` | The ideographs and kana already measure exactly; the Latin sub-run does not — [CJK](#a-note-on-cjk-the-approximate-grade-understates-it) |
| **Meiryo, Yu Gothic, Yu Mincho, Hiragino, Noto Serif JP** | Measured *and* drawn with Noto Sans JP | `approximate` | Same: exact on CJK, off on the Latin sub-run — [CJK](#a-note-on-cjk-the-approximate-grade-understates-it) |
| **Book Antiqua, Palatino Linotype, Century, Century Schoolbook, Century Gothic, Bookman Old Style, Monotype Corsiva, Arial Narrow, Symbol, Monotype Sorts, Comic Sans MS** | Widths guessed; drawn with the generic family | `missing` | A measured open clone exists but pptx2svg does not ship or map it — [tier 4](#4-a-clone-exists-but-is-not-bundled) |
| **Anything else** — Gill Sans MT, Verdana, Segoe UI, Georgia, Garamond, Tahoma, Trebuchet MS, Wingdings, the Indic and Thai faces… | Widths guessed; drawn with the generic family | `missing` | [Embed the font in the deck](#1-the-deck-carries-it-best), or supply it yourself — [escape hatches](#the-escape-hatches-and-their-trap) |

`exact` and `compatible` are faithful. `approximate` and `missing` are not, and
`pptx2svg fonts --check` exits non-zero on them.

Run `pptx2svg fonts` for the full table — 45 names, of which 20 grade faithful.

## The five ways a font gets drawn

Best outcome first.

### 1. The deck carries it (best)

A deck saved with *Embed fonts in the file* carries its typefaces in `ppt/fonts/*.fntdata`,
listed by `<p:embeddedFontLst>` in `ppt/presentation.xml`. pptx2svg decodes them (a pure-Python
MicroType Express decoder in the core — no extra needed) and uses them for **both
measurement and drawing**. Nothing is substituted, so no advance width is approximated.

```
$ pptx2svg fonts --check tests/fixtures/real-basic-theme.pptx
mode:   bundled from .../pptx2svg_fonts/files
        this machine's own fonts are ignored, so output is reproducible
        families: Arimo, Caladea, Carlito, Cousine, Lato, Noto Sans JP, Raleway, Tinos

face                       verdict      drawn with     note
Lato                       exact        Lato           drawn with the face the deck embedded
Raleway                    exact        Raleway        drawn with the face the deck embedded
Arial                      compatible   Arimo          Arimo has Arial's advance widths

all 3 faces resolve to the metrics they were measured at.
```

**This beats a bundled face of the same name**, which is not obvious and was decided on a
measurement. `real-basic-theme.pptx` embeds Raleway 4.026; the bundle ships a later
release. Over the 340 characters both tables describe, **338 advance widths differ**, and
`"Hamburgefonstiv 12345"` at 18 pt measures 261.000 px from the deck's own file against
251.184 px from the bundle's — 3.9 % apart. The author laid the deck out with the file
inside it, so that is the file to measure with. An embedded face therefore grades `exact`
and takes precedence over the substitution table even where a substitute exists.

It also beats *no* answer, which is the usual case. Template vendors pick arbitrary Google
Fonts, of which there are roughly 1,800 families; two commercial templates measured here
embed Anton, Arimo, Literata, Merriweather Sans, Merriweather Sans Light and Inclusive
Sans, and every one of them used to report "no substitute known; widths guessed" while
sitting inside the file being rendered. Scored against PowerPoint's own PDF export,
reading the embedded fonts raised SSIM on all twelve slides measured — the largest by
0.123, the mean by 0.058.

**If you control the deck, this is the fix for every `missing` row in the table above.**
Tick *Embed fonts in the file* in PowerPoint and the question stops being ours.

Three things it does not do:

- **Use a font whose licence refuses.** Every face is checked against its OS/2 `fsType`
  twice — in the EOT header, written by the embedder, and again in the decoded font,
  written by the foundry. The rule is LibreOffice's `EmbeddedFontsHelper::sufficientTTFRights`:
  refuse only a face whose permission field says *restricted licence* (0x0002) and nothing
  else, since rendering to a fixed image is a viewing use. `0x0200` (bitmap embedding only)
  is also refused, because we draw outlines. A refused face warns
  `font-embedded-restricted`, naming the restriction, and falls back to substitution.
- **Fail.** An undecodable payload warns `font-embedded-undecodable` (or
  `font-embedded-unreadable` where the part itself is absent) and the deck renders as it
  would have without it.
- **Cost nothing.** Decoding is 0.15–1.1 s per face, cached by content digest, and only
  the families a slide actually asks for are decoded. `--no-embedded-fonts`, or
  `ConvertOptions(use_embedded_fonts=False)`, turns it off.

### 2. The bundle supplies it exactly

Eight families, from `pip install 'pptx2svg[fonts]'`:

**Arimo, Caladea, Carlito, Cousine, Lato, Noto Sans JP, Raleway, Tinos.**

A deck that names one of these directly gets that exact file, measured from that exact
file. Decks reach these names by ordinary routes: Carlito and Caladea are LibreOffice's
own Calibri and Cambria substitutes, so anything round-tripped through LibreOffice names
them; Arimo, Tinos and Cousine are the Chrome OS core fonts that Debian packages as
`fonts-croscore`; Lato and Raleway are Google Fonts that Office offers from the cloud.

The list is not prose — it is
[`pptx2svg.fonts.BUNDLED_FAMILIES`](src/pptx2svg/fonts/__init__.py), and the identity rows
in the substitution table are generated from it, so a ninth family cannot ship without
being reachable by its own name.

Without the bundle, `pptx2svg` still renders, using the host's fonts, and emits one
`font-bundle-missing` warning saying that output is not reproducible. Every face then
grades `missing`, because in `system` mode we genuinely do not know what the rasteriser
will pick and claiming otherwise would be the confident-but-wrong answer this subsystem
exists to remove.

### 3. The bundle supplies a metric-compatible clone

*Metric-compatible* means one thing and one thing only: **the same advance widths**, so
line breaks, autofit and centring land where PowerPoint put them. Shape, colour, hinting
and x-height are all irrelevant to layout. A face can look nothing like the original and
still be a perfect substitute; a face can be a scholarly revival of the very same design
and still reflow every line.

| Office face | Drawn with | Verified |
| --- | --- | --- |
| Calibri | Carlito | ratio 1.0000 over four representative strings |
| Arial, Helvetica | Arimo | ratio 1.0000 |
| Times New Roman, Times | Tinos | ratio 1.0000 |
| Courier New, Courier | Cousine | ratio 1.0000 |
| Liberation Sans / Serif / Mono | Arimo / Tinos / Cousine | Liberation 2.x is built *from* the Chrome OS core fonts. The equality is inherited rather than re-measured here — 0 of 191 advances differ, and resvg draws Arimo at `wght=700` pixel-identically to static Liberation Sans Bold |

One caveat inside this tier: **Calibri Light** is a distinct face and Carlito has no light
cut, so Carlito draws it 1.3 % wide. It still grades `compatible`, because the verdict is
about internal consistency — we measure with Carlito and we draw with Carlito.

#### Cambria → Caladea is *not* in this tier

This is the clearest illustration in the codebase of why the distinction matters, and the
reason `Substitution.metric_compatible` is a field rather than an assumption.

Caladea is universally described as metric-compatible with Cambria. It is not. Measured
from the repository's own generated tables, over the 95 printable ASCII characters:

- **no advance matches exactly** — 95 of 95 differ;
- 57 of 95 differ by more than one design unit, which is the threshold this project uses
  to separate design from rounding;
- the **mean ratio is 0.9592**, range 0.530–1.089.

Against four representative strings the same error reads 0.9555, 0.9384, 0.9525, 0.9542 —
roughly 4.5 % narrow.

So Cambria gets the Aptos treatment instead: **measured with Cambria's own advance widths,
drawn with Caladea**, which is the closest serif that can be shipped. Line breaks match
PowerPoint; each drawn line is a few percent short. `fonts --check` grades it
`approximate` and `--check` exits non-zero on it.

Reproduce it:

```python
from pptx2svg.text.metrics import METRICS

cambria, caladea = METRICS["Cambria"], METRICS["Caladea"]
ratios = [
    (caladea.widths[c] / caladea.units_per_em) / (cambria.widths[c] / cambria.units_per_em)
    for c in map(chr, range(0x20, 0x7F))
]
print(sum(ratios) / len(ratios))   # 0.9592...
```

Publishing advance widths of a proprietary face is not redistributing it — they are facts
about a design, which is the footing on which Carlito, Arimo and Liberation exist at all.
No Cambria or Aptos outline, table or file is in this repository.

#### A note on CJK: the `approximate` grade understates it

Noto Sans JP is what every Japanese face is drawn with, and every one of them grades
`approximate`. For the CJK portion of a CJK run that grade is pessimistic. **Every
Microsoft CJK face gives every ideograph and every kana an advance of exactly 1.000 em** —
MS Gothic, MS Mincho, Meiryo, Yu Gothic, Microsoft YaHei, Malgun Gothic, DengXian, all of
them — and so does Noto Sans JP. Measured over 1,570 shared ideographs and kana, Noto Sans
JP differs from Meiryo, Yu Gothic, Microsoft YaHei, Malgun Gothic and DengXian on **zero**
codepoints. The divergence is confined to the Latin sub-run, which is 0.83 of Meiryo's
widths and 0.90 of Yu Gothic's.

The exceptions prove the rule: **MS PGothic makes kana proportional too** — its katakana
run from 0.648 em to 1.0 while Noto Sans JP's are uniformly 1.0 — which is why the four MS
faces are measured from their own tables rather than from Noto's. Details, and the thirteen
fixed-pitch families that Noto Sans Mono CJK JP matches on 0 of 1,106–2,264 shared
codepoints, are in
[ROADMAP ▸ CJK compatibility is nearly free](ROADMAP.md#cjk-compatibility-is-nearly-free-and-mostly-already-held).

### 4. A clone exists but is not bundled

A survey of the 186 family names a PowerPoint install can put in `a:latin/@typeface`
measured which of them have a genuine open metric clone. **26 do.** They are not in the
bundle, because the bundle *redistributes* and their licences (mostly AGPL-3.0 with a
PostScript/PDF-embedding exception that does not cover redistributing the files) do not
permit that. Installing them from Debian's archive is a different act, which is what
`tools/install-fonts-debian.sh --clones` does.

Headline rows, so you need not open the roadmap to learn your Palatino deck is covered:

| Family | Clone | Measured |
| --- | --- | --- |
| Book Antiqua | URW P052 | 0 of 74 text characters differ; 1 of 95 ASCII (`#`) |
| Palatino Linotype | URW P052 | 0 of 74 text; 10 of 95 ASCII — Microsoft regularised the math operators, URW kept Palatino's |
| Century Schoolbook, Century | URW C059 | 0 of 74 text; the 3 ASCII "differences" are one unit in 2048 |
| Century Gothic | URW Gothic | 0 of 74 text; 2 of 95 ASCII |
| Bookman Old Style | URW Bookman Light/Demi | 1 of 74 text (`Q`, +2.5 %) |
| Monotype Corsiva | URW Z003 | 0 of 74 text; 1 of 95. Surprising, and checked twice |
| Arial Narrow | Nimbus Sans Narrow, or TeX Gyre Heros Cn | 0 of 95 ASCII, all four cuts |
| Symbol | **Symbol Neu** (Apache-2.0) | **0 of 188 by legacy code, 0 of 189 by glyph name, mean 1.0000** |
| Monotype Sorts | URW D050000L | 9 of 202 differ, mean 0.9938 |
| Comic Sans MS | **Comic Relief** (OFL) | **0 of 95 ASCII, regular and bold.** The cleanest result in the survey |
| 13 CJK families (MS Gothic, MS Mincho, SimSun, NSimSun, SimHei, KaiTi, FangSong, MingLiU, MingLiU_HKSCS, BatangChe, GulimChe, DotumChe, GungsuhChe) | Noto Sans Mono CJK JP | **0 differing over 1,106–2,264 shared codepoints each** |

The full survey — including the faces the internet *suggests* and which all measurably
reflow (Gill Sans MT → Libre Franklin, mean 1.0774; Verdana → DejaVu Sans, 57 of 95 differ;
Garamond → EB Garamond, 91 of 95) and the families with no open equivalent at all — is
[ROADMAP ▸ The clone landscape](ROADMAP.md#the-clone-landscape--survey-measured).

**Installing a tier-4 clone does less than you would expect, and the reason is worth
knowing.** pptx2svg's substitution table has no row for any of these names, so:

- `pptx2svg fonts --check` still reports `missing`;
- measurement still guesses (see [tier 5](#5-nothing-matches-widths-are-guessed));
- the emitted SVG stack for a deck naming Palatino Linotype is
  `'Palatino Linotype', serif` — nothing in it names P052, so the rasteriser will not pick
  the clone up by itself.

Two different situations follow, and only one of them is a one-liner.

**You installed the *real* face.** Then the deck's own name is already first in the stack,
so `--system-fonts` draws with it. For Aptos and Cambria this is the **complete** fix,
because those two are already measured from their own tables — measurement and drawing
agree again:

```bash
tools/install-fonts-debian.sh --aptos        # Microsoft's own download, their terms
pptx2svg deck.pptx -f png --system-fonts     # measured as Aptos, now drawn as Aptos too
```

`fonts --check` will still say `approximate`: in `system` mode it reports the bundle's
answer, not fontconfig's, deliberately. Verify with the render, not the report.

**You installed a clone under a different name** (URW P052 for Palatino Linotype). Then you
need to say so twice — once for drawing, once for measurement:

```python
from pptx2svg import ConvertOptions, FontToolsTextMeasurer, convert_pptx_to_png

convert_pptx_to_png(
    "deck.pptx",
    ConvertOptions(
        font_mapping={"Palatino Linotype": "P052"},      # what to draw with
        measurer=FontToolsTextMeasurer(                   # what to measure with
            {"Palatino Linotype": "/path/to/P052-Roman.otf"}
        ),
    ),
    skip_system_fonts=False,
)
```

`font_mapping` alone gives right glyphs at guessed widths. See
[the escape hatches and their trap](#the-escape-hatches-and-their-trap).

### 5. Nothing matches: widths are guessed

No embedded face, no bundled family, no substitution row. This is where Gill Sans MT,
Verdana, Segoe UI, Georgia, Garamond and about 150 other PowerPoint families land.

**What actually happens.** `metrics_for()` returns `None` and every character is measured
by category instead: 0.3 em for a narrow character, **0.6 em for an ordinary one**, 1.0 em
for anything East Asian. The SVG names the requested family and a CSS generic
(`'Gill Sans MT', sans-serif`), and the generic points back at the bundle, so the text
draws at roughly the right size rather than not at all.

**What it costs.** Not glyph shapes — *line breaks*. Measured at 18 pt on
`"Hamburgefonstiv 12345"`, the guess returns 280.800 px for every unknown face. Against
the real tables:

| Face | Real width | The guess is |
| --- | --- | --- |
| Caladea | 231.312 px | **+21.4 %** |
| Tinos | 233.965 px | +20.0 % |
| Carlito | 235.055 px | +19.5 % |
| Raleway | 251.184 px | +11.8 % |
| Arimo | 254.824 px | +10.2 % |
| Lato | 258.996 px | +8.4 % |
| Cousine | 302.449 px | **−7.2 %** |

Cousine is the row that shows the guess is not even wrong in a consistent direction:
monospaced, it is 7.7 % *wider* than the fallback assumes, while Caladea is 17.6 %
narrower. A 20 % width error is not a cosmetic difference. It moves wrap points, changes
how many lines a paragraph occupies, changes what autofit shrinks to, and moves every
vertically centred block.

**How to detect it.** Two ways, both loud by design:

```
$ pptx2svg fonts --check deck.pptx
face                       verdict      drawn with     note
Arial                      compatible   Arimo          Arimo has Arial's advance widths
Courier New                compatible   Cousine        Cousine has Courier New's advance widths
ＭＳ Ｐゴシック                   approximate  Noto Sans JP   measured as ＭＳ Ｐゴシック, drawn as Noto Sans JP; no metric-compatible clone exists
Gill Sans MT               missing      -              no substitute known; widths guessed, drawn with the generic family

2 of 4 faces will not be drawn at the widths they were measured at.
```

and, for library callers, a warning on `ConvertOptions.warnings`:

```python
opts = ConvertOptions()
convert_pptx_to_svg("deck.pptx", opts)
[w.message for w in opts.warnings if w.code == "font-substituted"]
# ['Aptos: measured as Aptos, drawn as Carlito; no metric-compatible clone exists', ...]
```

`ConvertOptions(warn_on_font_substitution=False)` silences it, which you should not do.

**What to do.** In order of how well it works: embed the font in the deck
([tier 1](#1-the-deck-carries-it-best)); install the real face and render with
`--system-fonts`; or map and measure it yourself
([escape hatches](#the-escape-hatches-and-their-trap)).

## Which face draws the Japanese?

The five routes above answer *what do we draw this family with*. A run carrying kana or
ideographs has an earlier question to answer first — **which family** — and OOXML gives
it three places to look. In order:

1. the run's own `<a:ea typeface="..."/>`;
2. the theme font collection's `<a:ea>`;
3. the theme's `<a:font script="Jpan"/>` list (then `Hans`, `Hant`, `Hang`).

The Latin `<a:latin>` face is the last resort, not the default, and that ordering is
measured rather than assumed. PowerPoint's own PDF export of a deck whose charts name
`<a:latin typeface="Arial"/>` and nothing else, over a theme writing
`<a:ea typeface=""/>` and `<a:font script="Jpan" typeface="游ゴシック"/>`, embeds
**YuGothic-Regular** for the Japanese and **ArialMT** for the Latin runs *inside the same
labels*.

Two things to know when reading a deck's XML:

* **`typeface=""` is not a name.** Nearly every theme writes an empty `<a:ea>`; it means
  the collection names no East Asian face, and the script list is what answers.
* **A Latin face named as the East Asian one is ignored.** Decks really do write
  `<a:ea typeface="Raleway"/>`, and PowerPoint draws their Japanese in a Japanese face
  anyway. So does this library: a candidate that cannot draw kana loses to one that can.

One run can therefore be two faces, and it is split into separate `<tspan>` chunks that
each name their own, because rasterisers fall back per chunk rather than per glyph.
Measurement splits at exactly the same character boundaries, which is the same
measure-equals-draw rule the rest of this page is about.

## Ask the tool: `pptx2svg fonts --check`

This is the thing that answers the question for a *specific* deck, and it is the reason
this document does not have to enumerate 186 families.

```bash
pptx2svg fonts                       # every name the substitution table knows
pptx2svg fonts --check deck.pptx     # exit 1 if this deck cannot be drawn faithfully
```

A real run, on a fixture that uses Office's current default:

```
$ pptx2svg fonts --check tests/fixtures/authoring-integration.pptx
mode:   bundled from .../pptx2svg_fonts/files
        this machine's own fonts are ignored, so output is reproducible
        families: Arimo, Caladea, Carlito, Cousine, Lato, Noto Sans JP, Raleway, Tinos

face                       verdict      drawn with     note
Arial                      compatible   Arimo          Arimo has Arial's advance widths
Aptos                      approximate  Carlito        measured as Aptos, drawn as Carlito; no metric-compatible clone exists
Aptos Display              approximate  Carlito        measured as Aptos Display, drawn as Carlito; no metric-compatible clone exists

2 of 3 faces will not be drawn at the widths they were measured at.
$ echo $?
1
```

and one that passes:

```
$ pptx2svg fonts --check tests/fixtures/real-basic-theme.pptx
...
all 3 faces resolve to the metrics they were measured at.
$ echo $?
0
```

### The four grades

| Grade | Means | Faithful? |
| --- | --- | --- |
| `exact` | We draw the very family the deck asked for — bundled, or embedded in the deck | yes |
| `compatible` | We draw a clone built to the same advance widths. Outlines differ; nothing else does | yes |
| `approximate` | We draw *something*, but its widths are not the ones we measured with. Aptos drawn as Carlito, Cambria as Caladea, Mincho as a Gothic | **no** |
| `missing` | We know nothing about the face, or the bundle is not installed. The rasteriser gets the generic family; widths were guessed | **no** |

`--check` exits non-zero on `approximate` and `missing`, so a deck that cannot be rendered
faithfully fails a build instead of shipping wrong pixels.

Two properties are kept apart on purpose. `FontReport.faithful` asks "will every face draw
at the widths we measured?"; `FontReport.reproducible` asks "will two machines produce the
same pixels?" A deck can be faithful and not reproducible (system fonts that happen to be
right today), and perfectly reproducible while drawing Aptos as Carlito. Conflating the two
is how "it looks fine on my machine" becomes a production incident.

The faces reported are taken from the *resolved* model, not by grepping the XML for
`typeface="..."` — a theme lists forty-odd script fallbacks a deck never uses, and
`+mj-lt` is a pointer, not a face. What is reported is what a run ends up asking for after
the inheritance cascade has run, which is exactly what goes into the SVG.

## The escape hatches, and their trap

```python
# The host's fonts as well as the bundle
convert_pptx_to_png("deck.pptx", skip_system_fonts=False)

# Your own faces only
convert_pptx_to_png("deck.pptx", font_dirs=["./corporate-fonts"], use_bundled_fonts=False)

# Map a face to a substitute you do have
ConvertOptions(font_mapping={"Helvetica Neue": "Inter"})
```

On the command line: `--system-fonts`, `--no-bundled-fonts`, `--font-dir DIR`.

### The trap

**`font_dirs`, `font_files` and `--font-dir` feed the rasteriser only.** They reach
`svg_to_png`. Measurement has already happened by then, inside `convert_pptx_to_svg`,
which never sees them. So pointing `--font-dir` at a folder containing the exact face a
deck asks for produces **the right glyphs at guessed widths** — correct letters, wrong
line breaks. That is precisely the measure-with-one, draw-with-another failure this
subsystem exists to prevent, reachable through a documented flag.

It is not fully silent — the face is still absent from the bundle, so `font-substituted`
fires and its "widths guessed" clause is true. But its "drawn with the generic family"
clause is then false, so the warning gets the interesting half wrong in the reassuring
direction. The `font-bundle-missing` message compounds it by advising "or pass `font_dirs=`
explicitly", which is the parameter that fixes only drawing.

`font_mapping` has the same shape of limitation: it changes the `font-family` the SVG
names, not the table the layout was computed from.

### The complete route

Both halves, which needs `pip install 'pptx2svg[measure]'` (fontTools):

```python
import pathlib

from pptx2svg import ConvertOptions, FontToolsTextMeasurer, convert_pptx_to_png

fonts = pathlib.Path("./corporate-fonts")
measurer = FontToolsTextMeasurer({"Gill Sans MT": str(fonts / "GillSansMT.ttf")})

convert_pptx_to_png(
    "deck.pptx",
    ConvertOptions(measurer=measurer),   # measured from the real file
    font_dirs=[str(fonts)],              # and drawn from it
)
```

`FontToolsTextMeasurer` opens faces lazily and caches them; a family with no entry falls
back to the static tables, so a partial map is fine.

### The CLI cannot do this

**There is no measurer flag.** `pptx2svg --help` offers `--font-dir`, `--system-fonts`,
`--no-bundled-fonts` and `--no-embedded-fonts`, and none of them reaches measurement. A
command-line user's only route to a face pptx2svg does not know is `--font-dir`, and it is
the half-right one.

This is a current limitation, stated plainly rather than hidden: it is recorded in
[ROADMAP ▸ Fonts ▸ Left undone](ROADMAP.md#left-undone), where three candidate fixes are
listed and none is obviously best. The plumbing exists —
`DefaultTextMeasurer(extra_metrics=...)` takes a family → `FontMetrics` map that wins over
the static tables, and `fonts/sfnt.py` builds one from a font file using only the standard
library — so the change is small. It is unmade because it alters existing behaviour for
every caller of a documented parameter, which deserves its own decision.

Until then: **if you need both halves, use the Python API.**

## Aptos

Office's default face since 2023, shipped in six weights on a current PowerPoint install,
and the single most likely font in a deck saved this year.

**There is no metric-compatible open clone of Aptos.** Not Carlito, not Arimo, not Source
Sans, not anything. Pretending otherwise would be the same class of error as the Caladea
claim.

So the choice pptx2svg makes is explicit and is not a compromise between the two options —
it picks one:

- **Measured** with Aptos's own advance widths, so line breaks, autofit and centring match
  PowerPoint's.
- **Drawn** with Carlito, whose widths sit closest of anything shippable. On *"The quick
  brown fox jumps over the lazy dog"* at 18 pt, Carlito is **−3.6 %** against Aptos's own
  table, Arimo **+5.2 %** and Tinos **−2.8 %** — and Tinos is a serif, so it loses on shape
  what it gains on width.

Each drawn line is a few percent short, in the right place. That scores better than
re-wrapping the paragraph somewhere else. `fonts --check` grades it `approximate` and
exits non-zero, so nobody finds out from a screenshot.

**The question is settled, and settled negative.** Checked again in September 2026. The
Document Foundation's own public position concedes that its proposed replacements "are not
metrically compatible" — Aptos is a 2048-upem design and Source Sans, the usual suggestion,
is 1000. The open request to adopt one anyway (tdf#162872, September 2024) has produced
nothing. **Every non-Microsoft renderer reflows Aptos text today**, LibreOffice included,
whose built-in replacement table is empty by default. *(The bugzilla thread itself is behind
Anubis and could not be fetched; the TDF position statement is the citation here, not the
ticket discussion.)*

**What to do about it.** Microsoft publishes Aptos for download under its own terms. Install
it and render with `--system-fonts`, and the invariant is restored — pptx2svg already
measures with Aptos's table, so the real file makes measurement and drawing agree:

```bash
tools/install-fonts-debian.sh --aptos
pptx2svg deck.pptx -f png --system-fonts
```

The script pins the download by SHA-256 and refuses a changed payload rather than
substituting silently. You accept Microsoft's terms; this project does not grant you any
right to the file.

## The licensing boundary

Short, and it has no exceptions.

**No Microsoft font is bundled or committed, ever.** Not in the wheel, not in the
repository, not in a test fixture. What this project does publish is *measurements* —
Aptos's and Cambria's advance widths live in `text/metrics.py`. Advance widths are facts
about a design, which is the footing on which Carlito, Arimo and Liberation exist at all;
no outline, table or file is here.

**Everything the bundle ships is OFL.** `pptx2svg-fonts` declares `MIT AND OFL-1.1` and
carries eight SIL Open Font License families. A face is in it only if it can be
redistributed inside a wheel.

**The installer's non-free tiers are you accepting someone else's terms, not us
redistributing anything.** `tools/install-fonts-debian.sh` has five tiers. Tier 1 installs
the same designs the bundle ships, from apt. Tier 2 (`--clones`) adds open faces whose
licences permit installation but not redistribution-in-a-wheel — `fonts-urw-base35` (AGPL-3.0
with a PostScript/PDF-embedding exception that covers embedding a glyph in a PS or PDF file
and **not** shipping the font files), `fonts-liberation-sans-narrow` (GPL-2 with a font
exception), plus hash-pinned Comic Relief and Symbol Neu. Tiers 3–5 (`--aptos`,
`--mscorefonts`, `--ppviewer`, `--office-dir`) reach Microsoft's own downloads, Debian's
contrib packaging of Microsoft's EULA'd fonts, and a licensed Office installation you point
at. The script installs from an archive or copies files you already hold. It redistributes
nothing and grants you nothing.

That asymmetry is the whole reason tier 2 exists separately from the bundle: the wheel
*redistributes* and must be conservative; the script *installs* and need not be.

---

## Reproducing anything on this page

```bash
pip install 'pptx2svg[fonts,measure]'
pptx2svg fonts
pptx2svg fonts --check tests/fixtures/authoring-integration.pptx
```

Every measured number here comes from the repository: the substitution table is
[`src/pptx2svg/text/fontmap.py`](src/pptx2svg/text/fontmap.py), the grades are
[`src/pptx2svg/fonts/check.py`](src/pptx2svg/fonts/check.py), the advance-width tables are
[`src/pptx2svg/text/metrics.py`](src/pptx2svg/text/metrics.py) and are generated from the
shipped font files by `tools/extract_font_metrics.py` (a test re-runs it and fails on
drift). The clone survey is
[ROADMAP ▸ The clone landscape](ROADMAP.md#the-clone-landscape--survey-measured); the
installer's tiers are documented in the header of
[`tools/install-fonts-debian.sh`](tools/install-fonts-debian.sh).

A claim of metric compatibility with no measurement behind it is exactly what this project
does not ship.

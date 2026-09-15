# Roadmap to full PPTX support

Working plan for closing the gap between what `pptx2svg` renders today and what a PPTX
can contain. Written to be picked up cold: every item says what is missing, where the
code lives, how to approach it, and how to know it is done.

Phases are ordered by *dependency and payoff*. Phase 0 comes first because everything
after it needs a way to tell "better" from "different".

**Effort key:** S ≈ half a day · M ≈ 1–3 days · L ≈ 1–2 weeks · XL ≈ 3+ weeks.

**Sources.** Sizes are calibrated against two existing implementations, both MIT:

- [pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse) — the TypeScript library
  this one is ported from. Referenced for the pieces not yet ported.
- [aiden0z/pptx-renderer](https://github.com/aiden0z/pptx-renderer) — a browser-native
  PPTX renderer with materially wider coverage (187+ presets, charts, table styles,
  EMF previews) and a PowerPoint-based visual regression harness. Reviewed specifically
  for techniques worth adopting; the findings are marked **[pptx-renderer]** below.

### How the field handles fonts

Surveyed in September 2026, because this library's central bet — *compute layout from a
generated table of advance widths, and draw with the very same files* — is unusual enough
to be worth checking against everyone else's answer.

| | Embedded fonts | Metrics source | Missing-font strategy |
| --- | --- | --- | --- |
| LibreOffice ≥ 25.8 | **yes**, libEOT → temp font, `fsType`-gated | fontconfig / system | hardcoded metric map → fontconfig hook → `VCL.xcu` → attribute match |
| Apache POI XSLF | API exists, renderer ignores it | `java.awt` `FontRenderContext` | `Font.SANS_SERIF`, or a `FONT_MAP` you supply |
| python-pptx | no | Pillow, macOS/Windows only | n/a — it does no layout |
| Aspose.Slides | yes | own engine | `FontSubstRule` + PowerPoint-alike default |
| Syncfusion | not documented | own engine | Microsoft Sans Serif |
| PPTXjs / js-pptx | no | **none** — CSS plus a global fudge factor | whatever CSS falls back to |
| **[pptx-renderer]** | **yes**, EOT+MTX, `fsType`-gated | the browser | CSS stack, silently |
| ONLYOFFICE | rights-checked | **own FreeType/WASM + shipped `font_selection.bin`** | penalty-based match |
| Collabora Online | via LibreOffice | LibreOffice, server-side | LibreOffice's, plus remote font download |
| **pptx2svg** | **no** — see Phase 6 | **generated metric table** | clone set + `font-substituted` |

**The load-bearing result is not that this library compares well. It is why.** Two
projects arrived independently at "carry your own metrics": ONLYOFFICE compiles its own
FreeType engine to WebAssembly and feeds it a server-generated `font_selection.bin`
metrics blob, and the Rust project `Ryujiyasu/oxi` ships generated metric tables with
bundled clones and a resolved / DEGRADED / ABSENT report. Neither borrowed the idea from
the other. Two independent arrivals is the strongest available validation of the
principle the whole `text/` and `fonts/` subsystem rests on.

**The contrast worth keeping is silence.** Most of the field substitutes without saying
so. PPTXjs never measures at all and applies a global `fontSizeFactor` "browser rendering
adjustment" — a fudge constant is exactly what a renderer needs when it has no metrics,
and it is the clearest illustration available of what `font-substituted` exists to
prevent. Apache POI's own developers concede `getTextHeight()` "varies from
Windows/Linux/Mac". python-pptx closes every render-to-image issue with "use LibreOffice".

Two corrections to notes elsewhere: **Collabora Online is not a JS renderer** — server-side
LibreOffice rasterises PNG tiles and the browser is a viewport — and **nodeppt is a
Markdown-to-HTML authoring tool, not a PPTX renderer**. Neither belongs in a comparison of
rendering approaches.

---

## Where we are

| Area | State |
| --- | --- |
| Package, relationships, parts | Complete |
| Theme colours, colour maps, transforms | Complete |
| Placeholder / background / text inheritance | Complete |
| Shapes: all 187 presets PowerPoint accepts + custom geometry | Complete and **verified against PowerPoint**: every preset matches its own PDF export at two aspect ratios, median 0.999 silhouette overlap |
| Text: cascade, bullets, wrapping (Latin + CJK), autofit, vertical, tabs, columns | Complete for the common path |
| Fills, outlines, arrowheads, shadows, glow, soft edge | Complete |
| Pictures: crop, colour adjustments, tile, stretch | Complete |
| Tables: merged cells, borders, fills, **table styles** | Complete; **all 74** built-in styles carried, every one measured out of PowerPoint; an id in neither the deck nor the catalogue now warns `table-style-unknown` instead of rendering a bare grid in silence; cell text takes a table style's `tcTxStyle` over the master's `otherStyle` |
| Charts | `barChart`, `lineChart`, `pieChart`, `doughnutChart` and `radarChart` read and drawn with their data labels, **verified against PowerPoint** across 190 probe charts and every chart in the corpus; **no deck warns `chart-unsupported-type` any more**. Category labels wrap at whitespace and turn 45° only when their widest unbreakable token still will not fit. Every other chart type warns and draws an empty frame |
| SmartArt | Cached drawing rendered and verified against 46 real decks; **no layout engine**, so diagrams without a cache draw nothing and say so |
| EMF / WMF | Embedded previews rendered (Phase 4); **no vector interpreter** |
| 3-D, bevel, reflection | **Not rendered** |
| Shape identity on output (`data-pptx-id`) | Complete |
| Fonts: bundled, metric-generated, diagnosed | Complete; **Aptos and Cambria approximate** |

1,968 tests pass. The pipeline is `opc → parse → resolve → render → png`; each stage is
independently testable, and every phase below slots into exactly one of them.

### Shape identity

Every rendered element carries `data-pptx-id` -- `"<sldId>.<cNvPr id>"` for slide shapes,
`lay:`/`mst:` for shapes inherited from the layout or master -- plus `data-pptx-path`, the index
path through the shape tree. Both are set in `resolve/view.py:resolve_element`, the single
dispatch point, so every nesting depth is covered for free, and written out in
`render/svg.py:render_element`.

The path is not redundant: **`cNvPr@id` is not unique in real decks.** In
`real-financial-report.pptx` slide 2 a `p:sp` and a `p:graphicFrame` both carry `id="3"`. The
pair `(id, path)` is unique, and a downstream consumer needs both to address a shape.

### Measured baseline

Against real PowerPoint output, every slide of every fixture, produced by
`tools/fidelity.py` and recorded in `tests/fidelity-baselines.json`.

**These numbers mean something different from the ones they replace.** Every table that
used to sit here was partly a measurement of font availability: PowerPoint drew with
Microsoft's Calibri, Cambria and Aptos, our side drew with whatever resvg could find, and
for five of the seven fixtures that was a generic sans. The difference between the two
images was dominated by glyph shape before the renderer had done anything at all. The
harness now renders our side with the *same licensed faces* PowerPoint used, read in
place through a gitignored local profile, and **a deck whose faces PowerPoint did not have
either is skipped rather than scored against Microsoft's own fallback**.

| Fixture | SSIM | hist | Gate (≥0.95 / ≥0.80) |
| --- | --- | --- | --- |
| `table test.pptx` | **0.9895** | 0.9984 | pass |
| `authoring-integration.pptx` | 0.9327 | 0.9984 | SSIM |
| `real-college-template.pptx` (local only) | 0.8003 | 0.8753 | SSIM, hist |
| `real-basic-theme.pptx` | skipped | — | PowerPoint drew MS Gothic where the deck names ＭＳ Ｐゴシック |
| `sample.pptx` | skipped | — | same |
| `real-financial-report.pptx` | skipped | — | no Noto Sans JP on this machine, so PowerPoint substituted too |
| `real-product-page.pptx` | skipped | — | same |
| `sample-issue-387.pptx` | skipped | — | same |

**Five of eight fixtures cannot be scored at all, and that is the single biggest hole in
this project's feedback loop.** Not because the renderer is wrong on them — because the
oracle and the renderer disagree about which *face* to draw, so any number would measure
font resolution rather than layout. Two routes close it, neither taken here because both
change the developer's machine rather than the repository: install Noto Sans JP where
PowerPoint can see it (`~/Library/Fonts`) and re-export the three decks that name it, or
rewrite the two Japanese decks' themes to name `MS Gothic` — the face PowerPoint actually
resolves — instead of ＭＳ Ｐゴシック.

`real-college-template.pptx` escapes that trap: it names only Arial, Calibri and
Wingdings, all of which this machine has, so it is **the first real-world deck measurable
rather than skipped**. It is third-party and therefore **not committed** -- it lives in
the gitignored `scratch/`, and the tests that need it skip where it is absent (see
`tests/fixtures/README.md`). The defects it found are pinned by tests that do *not* need
it. Its nine slides:

| slide | SSIM | hist | what is left |
| --- | --- | --- | --- |
| 1 | 0.9931 | 1.0000 | — |
| 2 | 0.9870 | 0.9998 | — |
| 3 | 0.7627 | 0.1195 | CMYK in an EMF: pdfium and Quartz disagree on the logo's red by a visible amount. 2.9% coverage, nearly all logo, is why one hue moves the histogram this far. Needs an ICC transform in `metafile/pdf.py`, and the answer would be machine-specific. |
| 4 | 0.3721 | 0.9899 | `c:userShapes`, literal text in a number format, a manually laid out legend, the plot area's own `c:spPr` border. See Phase 3. |
| 5 | 0.9528 | 0.9997 | — |
| 6 | 0.7372 | 0.9611 | Line tops now agree within 1–2 px; at 20 pt over 5% coverage that alone costs most of the SSIM. |
| 7 | 0.5112 | 0.9715 | same |
| 8 | 0.8869 | 0.8365 | A 441 kB JPEG: our resampling and colour differ slightly from pdfium's at 58% coverage. |
| 9 | 1.0000 | 1.0000 | sparse — too little foreground to judge |

### Two limits of this oracle, worth knowing before chasing a number

**The oracle is localised.** This Office install renders decimals with a comma (`$8,0`,
`($1,0)`), so slide 4's numeric text can never match character-for-character regardless of
what the number-format code does.

**The metric is not monotone in geometric accuracy.** Measured, not assumed: setting slide
4's plot bottom to PowerPoint's own *to the pixel* scores 0.4704 SSIM / 0.9842 hist, while
reserving one label line instead of the two PowerPoint drew — a band 11.5 pt too short —
reproduces its old 0.9950 histogram exactly. The last half-percent on that slide is only
purchasable by being measurably wrong, so it was left unbought.

### What the corpus says is wrong now

With fonts eliminated as a variable, what remains is layout, and it is concentrated:

* **CJK label width** is the largest single error in any chart, and it is not a chart bug:
  `font_box` and `text_width` measure a Japanese label through the `<a:latin>` face its
  axis names rather than the CJK face it will be drawn in. On
  `real-financial-report.pptx`'s chart3 that is 19.8 pt. It lives in `text/` and `fonts/`.
* **A manually laid out legend** is the whole of slide 4's remaining chart error.
* **Text displacement of 1–2 px** on body copy is what slides 6 and 7 are made of, and
  `tools/fidelity.py`'s own docstring warns that SSIM is unusually sensitive to exactly
  that on thin high-contrast content.

The bold/italic and paragraph-spacing defects that used to head this list are **fixed**:
bold now inherits through the placeholder cascade, `spcBef` and `spcAft` add rather than
collapse, and `spcPct` is a share of the line height rather than the font size. Slide 6
went 0.0977 → 0.7372 on the spacing fix alone.

---

## Fonts — **done**

The problem, stated precisely: layout was computed from a table of advance widths, the
SVG named a `font-family`, and **nothing made those two agree**. On a machine without
Calibri the rasteriser substituted something else without a word — rendering one string
in Calibri, Carlito, Aptos, Noto Sans JP and Lato produced five byte-identical PNGs. Text
appeared, at widths nothing had computed.

### The shape of the fix

1. **Ship the faces**, in a separate `pptx2svg-fonts` distribution installed as
   `pptx2svg[fonts]`.
2. **Generate the metrics from them.** `tools/extract_font_metrics.py` reads the shipped
   files and rewrites `text/metrics.py`; a test re-runs it and fails on drift.
3. **Draw with them by default**, `skip_system_fonts=True`, so output does not depend on
   the host.
4. **Say so when we cannot** — `pptx2svg fonts --check`, and warnings on
   `ConvertOptions.warnings`.

### Why a separate distribution

Extras cannot conditionally add package data, so the choice was "font files in every
wheel" or "font files in a wheel you opt into". The main wheel is 150 kB; the fonts are
11 MB compressed. **SVG output embeds no fonts at all** — an SVG names them — so every
caller who only wants SVG, or who points `font_dirs` at their own corporate faces, would
have paid 11 MB for bytes they can never use. PNG already requires an extra
(`pptx2svg[png]`), so the audience that needs fonts is already typing an extra.

The cost is a second distribution to release in lockstep, and a bare `pip install
pptx2svg` that is *not* deterministic. That second cost is paid down by making the
degraded mode loud rather than silent: `pptx2svg.fonts.bundle_mode()` is the single source
of truth, `pptx2svg fonts` leads its report with it, and a render without the bundle emits
one `font-bundle-missing` warning naming the fix.

### What was found by measuring instead of trusting

Every one of these was believed to be true beforehand, on good authority, and was not:

* **Caladea is not metric-compatible with Cambria.** It is described that way
  everywhere. Measured against the installed Cambria over four representative strings it
  runs 4.5 % narrow (0.9555, 0.9384, 0.9525, 0.9542). Cambria therefore gets the Aptos
  treatment: measured with its own widths, drawn with Caladea.
* **The inherited Noto Sans JP table was wrong for 186 of 191 characters.** It came from
  pptx-glimpse and had never been checked against the font it claimed to describe. This
  is the single best argument for generating the table.
* **Bold was a flat 1.05 multiplier.** The true ratio is 1.000 for Cousine — monospace
  bold is the same width, and we were inflating every bold line of code by 5 % — 1.023 for
  Carlito, 1.049 for Tinos, 1.056 for Arimo, 1.089 for Caladea, 1.119 for Noto Sans JP.
  1.05 is the middle of that spread, which is another way of saying it was wrong for all
  of them. Bold now has its own generated table.
* **Aptos Display was not missing.** Office does not install it; it downloads it into a
  cloud-font cache under a numeric filename. Reading `/BaseFont` out of PowerPoint's own
  PDF export is what found it. Two of the seven fixtures use it, and they were the two
  worst-scoring.
* **resvg handles variable fonts correctly.** Arimo at `wght=700` renders pixel-identical
  to static Liberation Sans Bold, which is why the bundle ships one 1 MB variable Arimo
  instead of four static cuts.
* **Debian does not package Raleway.** The obvious `fonts-raleway` does not exist.
* **The substitution table only worked in one direction.** Every row was written from the
  Office side — "a deck asked for Calibri, what do we draw?" — which made a family a legal
  *input* only if some Office face happened to be spelled that way. The five faces the
  bundle ships, measures and draws with were therefore names we did not recognise:
  `metrics_for("Carlito")` returned `None` and the string was laid out from the 0.6 em
  per-character guess. At 18 pt on "Hamburgefonstiv 12345", Carlito, Arimo, Tinos, Cousine
  and Caladea each measured 280.800 px — the same 280.800 px as a face that does not exist
  — against 235.055, 254.824, 233.965, 302.449 and 231.312 px in their own tables, which
  were sitting in `METRICS` the whole time. Cousine shows the guess is not even wrong in a
  consistent direction: monospaced, it is 7.7 % *wider* than the fallback assumed, while
  Caladea is 17.6 % narrower. Decks reach these names by ordinary routes — Carlito and
  Caladea are LibreOffice's own Calibri and Cambria substitutes, `fonts-croscore` and
  `fonts-liberation2` put the rest in front of every Linux author — and `fonts --check`
  called each of them "no substitute known", the exact opposite of the truth. The identity
  rows now come from `BUNDLED_FAMILIES` itself, so a ninth family cannot ship without one.
* **Two heuristics were hiding behind the same gap.** With no row to consult,
  `generic_family` fell back to reading the *name*, and nothing in "Tinos", "Caladea" or
  "Cousine" says serif or monospace — all three ended their font stack in `sans-serif`.
  That is only the last resort in the stack, but the last resort is where it bites: in
  `system` mode a deck set in Tinos degraded to resvg's sans default. And `ascender_ratio`
  had no descender to work from either, so the first baseline came from
  `DEFAULT_ASCENDER_RATIO` — y=32 against Carlito's own y=29.806 on the probe deck.
* **Nothing in the corpus was affected, and that is the finding.** All seven fixtures
  re-render byte-identically, because not one of them names a substitute face directly;
  the fidelity scores cannot move. The bug was invisible precisely where it was measured.
* **Five of seven fixtures resolved most runs to `font_family=None`.** Nothing in OOXML
  obliges anyone to name a typeface, and plenty of decks name one nowhere. That meant no
  `font-family` in the SVG *and* no metrics table, so strings were measured with a 0.6 em
  per-character guess. It now falls back to the theme's body face, which is what
  PowerPoint draws.

### Aptos

Microsoft's Office default since 2023 has no open metric-compatible clone, and pretending
otherwise would be the same class of error as the Caladea claim above. The choice is
therefore explicit and is not a compromise between the two options — it picks one:

* **Measured** with Aptos's own advance widths, so line breaks, autofit and centring match
  PowerPoint's.
* **Drawn** with Carlito, whose widths are closest of anything shippable: −3.6 % on a
  representative sentence, against Arimo's +5.2 % and Tinos's −2.8 % (Tinos is a serif, so
  it loses on shape what it gains on width).

That deliberately breaks the measure-equals-draw invariant, which is why
`Substitution.metric_compatible` exists as a field rather than an assumption, why
`pptx2svg fonts` grades it `approximate` rather than `compatible`, and why `--check` exits
non-zero on it. A deck in Aptos can be rendered; it cannot be rendered *faithfully*, and
the tooling says so rather than letting someone find out from a screenshot.

Publishing measurements of a proprietary font is not redistributing it. Advance widths are
facts about a design, which is the footing on which Carlito, Arimo and Liberation exist at
all; no Aptos outline, table or file is in this repository.

**Settled, and negative — checked again in September 2026.** There is still no open
metric-compatible clone of Aptos, and the question does not need re-opening by the next
person who wonders. The Document Foundation's own public position concedes that its
proposed replacements "are not metrically compatible"; Aptos is a 2048-upem design and
Source Sans, the usual suggestion, is 1000. The open request to adopt one anyway
(tdf#162872, September 2024) has produced nothing. **Every non-Microsoft renderer reflows
Aptos text today** — LibreOffice included, whose built-in Tools ▸ Options replacement
table is empty by default.

That matters beyond this entry: Aptos ships in six weights on a current PowerPoint
install and is the default face a new deck gets, so it would be the single most important
row of any survey of open equivalents. It is already answered. Anyone running that survey
should skip it and spend the budget on the faces that might actually have clones.

*Source note:* the bugzilla thread itself is behind Anubis and could not be fetched by two
independent attempts, so the TDF position statement is the citation here, not the ticket
discussion. Flagged rather than laundered into a firmer claim than it is.

### Left undone

* **Emoji.** `real-product-page.pptx` renders through PowerPoint with AppleColorEmoji.
  Nothing open and redistributable is a drop-in for platform emoji, and colour emoji
  fonts are large. Emoji render as the substitute's glyph or not at all.
* **Complex scripts.** `a:cs` typefaces — Arabic, Hebrew, Thai, Devanagari — have no
  entry in the substitution table, so they fall through to the generic family and the
  per-category width guess. Adding them means shipping Noto for each script, which is
  another CJK-sized decision.
* **Italic is measured from the upright table.** Divergence is ≤2.5 % across the Office
  substitutes (Carlito 0.4 %, Tinos 2.2 %, Caladea 2.5 %, Arimo and Cousine 0.0 %), which
  is inside the noise of everything else. Lato is the outlier at 6 %.
* **Approximate mappings for non-Office faces.** Segoe UI, Verdana, Georgia, Consolas and
  friends are unmapped. Mapping them to a bundled family would be better than the 0.6 em
  guess, but only if the mapping is measured first — guessing is how the Caladea claim got
  in.
* **`font_dirs` / `font_files` / `--font-dir` feed the rasteriser only.** They reach
  `svg_to_png`; measurement has already happened by then, in `convert_pptx_to_svg`, which
  never sees them (`__init__.py`, where `convert_pptx_to_png` calls
  `convert_pptx_to_svg(source, options)` and only then passes the font arguments on). So
  pointing `--font-dir` at a folder containing the face a deck asks for produces **the
  right glyphs at guessed widths** — correct letters, wrong line breaks. That is the
  measure-with-one-draw-with-another failure this subsystem exists to prevent, reachable
  through a documented flag.

  It is not fully silent: the face is still absent from the bundle, so `font-substituted`
  fires and its "widths guessed" clause is true. But its "drawn with the generic family"
  clause is then false, and **the warning gets the interesting half wrong in the
  reassuring direction**. Worse, the `font-bundle-missing` message advises "or pass
  `font_dirs=` explicitly" — recommending the parameter that fixes only drawing.

  Callers can already get both halves right by combining `measurer=FontToolsTextMeasurer({...})`
  with `font_dirs=`, which needs `pptx2svg[measure]`. **The CLI cannot**: there is no
  measurer flag, so `--font-dir` is the only route a command-line user has, and it is the
  half-right one. Three candidate fixes, none obviously best: teach `font_dirs` to feed
  measurement as well (changes existing behaviour), add a CLI measurer option, or narrow
  the warning text so it stops claiming to know what the rasteriser drew with.

### Two open design questions, from the landscape review

Neither is a defect and neither is being acted on. Both are places where LibreOffice
encodes a distinction this library currently collapses, recorded so that whoever next
touches `text/fontmap.py` can decide deliberately rather than rediscover the question.

**1. Should the metric-compatible tier be unoverridable?**

LibreOffice resolves a missing font in five steps, and the *order* is the claim.
`PhysicalFontCollection::FindFontFamily` tries the exact name, then
**`FindMetricCompatibleFont` — a hardcoded C++ map** (`"Arial"` → `"Liberation Sans"`,
`"Times New Roman"` → `"Liberation Serif"`), then the fontconfig pre-match hook, then the
configurable `VCL.xcu` substitution lists, then attribute matching. The metric map sits
*deliberately ahead* of the configurable table, because metric compatibility is a stronger
claim than similarity and a user's preference list should not be able to override it
silently.

The `VCL.xcu` entry, verbatim, for reference:

```xml
<node oor:name="calibri" oor:op="replace">
  <prop oor:name="SubstFonts"><value>carlito;hiraginomarugothicpronw3;hiraginomarugothicprow3</value></prop>
```

`create_font_mapping()` merges a caller's dict straight over `DEFAULT_FONT_MAPPING` with
no such tier. So a caller can today override Calibri → Carlito with a face that reflows
the deck, and nothing says a word — the same silent-substitution failure the
`font-substituted` warning exists to close, reachable through a documented parameter. The
question is whether identity and metric-compatible rows should be structurally
unoverridable, or whether overriding one should at least warn. Not obvious: the escape
hatch exists because people have legitimate reasons to force a face.

**2. Should the "widths guessed" tier fall back per codepoint rather than per run?**

LibreOffice splits its substitution hook in two (`vcl/inc/font/fontsubstitution.hxx`):
`PreMatchFontSubstitution`, whole-font and configured, and
`GlyphFallbackFontSubstitution`, per-codepoint and driven by a set of missing codepoints.

This library has only the whole-font kind, so a face with no clone drops its entire run on
the generic family and the 0.6 em guess. Per-codepoint fallback would let most of a run
resolve properly and confine the guess to the characters that genuinely have nowhere to
go. That is the known better shape for the `missing` tier; it is not a small change, and
it interacts with measurement, since a run drawn from two faces must be measured from two
tables.

---

## Phase 0 — PowerPoint as the fidelity oracle

**Effort: M. Blocks everything else.**

Nothing currently catches a change that makes rendering *subtly worse*. The suite checks
structure (well-formed SVG, no `pt` units, text present) but cannot tell that a shadow
moved 4 px or a paragraph gained a line. Every phase below changes pixels.

**Microsoft PowerPoint 16.106 is installed on this machine**, which makes the reference
implementation itself available as ground truth — strictly better than LibreOffice, which
is not installed here anyway and is only an approximation of PowerPoint. **[pptx-renderer]**
uses exactly this approach and reports SSIM / colour-histogram / IoU metrics per slide
against PowerPoint exports.

### 0.1 The oracle — verified working

`tools/powerpoint_export_pdf.applescript` is in the repo and confirmed working on this
machine. It drives PowerPoint via AppleScript to export a deck to PDF; `pypdfium2` then
rasterises each page.

```bash
osascript tools/powerpoint_export_pdf.applescript "$PWD/deck.pptx" "$HOME/gt/deck.pdf"
```

```python
import pypdfium2 as pdfium
pdf = pdfium.PdfDocument(pdf_path)
page = pdf[0]
image = page.render(scale=1280 / page.get_size()[0]).to_pil()
```

**Constraints found by testing, all of which bite silently:**

| Constraint | Consequence |
| --- | --- |
| PowerPoint is sandboxed | Paths must be in a directory PowerPoint has **already been granted**. Under `$HOME` is *not* sufficient: a freshly created `~/pptx2svg-star/`, and a fresh directory under `~/Documents/`, both fail with **−9074** exactly as `/tmp` does, while the directory earlier exports used keeps working. The deck opens (a `~$` lock file appears) and only the save fails, so it reads as a broken deck rather than an unapproved path — which cost most of a session. Reuse the directory that already works. |
| **−9074 has a second cause** | A file PowerPoint wants to repair raises an app-modal dialog, and *every* export then fails −9074 until it is cleared — including known-good files. Check for a dialog before suspecting the path; one bad input otherwise looks exactly like a broken environment. |
| **−9074 has a third cause, and `pkill` is the fix** | After one export failed on a malformed deck, *every* subsequent export failed −9074 — including the corpus decks that had exported minutes earlier, so it reads as a revoked sandbox grant rather than a wedged app. AppleScript `quit saving no` returns `missing value` and does nothing, because PowerPoint is stuck on the half-open deck. `pkill -x "Microsoft PowerPoint"` clears it and the very next export succeeds. Before concluding the directory lost its grant, kill the app and retry — it turns a dead end into ten seconds. Remember to delete the `~$` lock the failed attempt left behind. |
| A fifth input defect, and it **repairs** rather than hangs | A content-type `Override` whose `PartName` starts `//`. `tests/deckbuilder.py` prepends the leading slash itself, so a caller passing `f"/{part}"` produces one; the three chart probe fixtures did exactly that. Our reader never looks at `[Content_Types].xml`, so the suite passed, but PowerPoint opens such a deck as `<name> [Repaired]` -- and the export script then fails with −2700 "no presentation matched", because the repaired presentation's full name is no longer the path it was asked for. Fixed in `tests/test_chart.py`; check any new caller. |
| **A third input defect with the hang signature** | Two series in one plot group both claiming `<c:idx val="0"/><c:order val="0"/>`. PowerPoint opens the deck and then never returns, exactly like the non-standard preset name and the partial `avLst` already listed. Out-of-order children of `c:ser` and `c:lineChart` (the schema's sequence is strict: `marker` before `dLbls`, the group-level `marker` *after* every `ser`) cost an earlier −9074 the same way. When a generated probe deck hangs, validate it against the schema sequence before suspecting the oracle. |
| The export script cannot clear that dialog | It is blocked inside `open` and never regains control. Dismissal has to run in a separate process, and **Escape does not work** — only a real button click does, matched across localisations (`Annuleren` on a Dutch install). |
| `count of presentations` is not a health check | A wedged PowerPoint answers `0` while still refusing every file. |
| Restarting re-raises the dialog | PowerPoint reopens the document it was killed over. Dismiss rather than restart; and after any restart, poll until it answers — an `open` sent mid-launch is refused instantly with −9074. |
| `save as PNG` is in the dictionary but **silently no-ops** | Returns success, writes nothing. PDF is the only export that works unattended. Do not spend time on it. |
| Per-slide PNG needs VBA | PowerPoint's `Slide.Export` is VBA-only, requiring a macro-enabled `.pptm` host. **[pptx-renderer]** does this; PDF + `pypdfium2` avoids the complexity and the macro-security friction. |
| First run prompts for automation permission | Fine interactively; in CI this must be pre-granted or the oracle skipped. Dismissing the repair dialog needs a *second* permission, Accessibility. |
| The script wraps its work in `with timeout of 45 seconds` | Otherwise a modal dialog costs the 120-second AppleEvent default on every attempt. A timeout here means "PowerPoint would not open this file", which is the verdict the oracle exists to give. |
| Must match the presentation by full path | Otherwise a concurrently open deck gets exported instead. The script does this. |

**Deliverable:** `tests/oracle/` with a `generate_ground_truth.py` that walks the fixture
corpus, exports each deck once, caches the PNGs by input SHA-256, and skips cleanly when
PowerPoint is absent (so `pytest` still passes on Linux CI and other machines).

### 0.2 Snapshot VRT (the regression net)

Ground truth needs PowerPoint; regression detection does not. Commit our *own* rendered
PNGs and fail on drift.

- `tests/vrt/` with committed snapshots and a `--update-snapshots` flag.
- Rendering is already deterministic (counter-based ids, generated font metrics) **and
  the fonts are now pinned**: `pip install 'pptx2svg[fonts]'` and the default render
  ignores the host entirely, which is what makes a committed PNG snapshot meaningful
  across macOS and CI. Before that, snapshots would have differed by machine.

Two layers, two jobs: snapshot VRT runs everywhere and catches regressions; the PowerPoint
oracle runs on this Mac and catches *being wrong in the first place*.

### 0.3 Metrics — **done**

`tools/fidelity.py`. Raw pixel-difference percentage is a blunt instrument — a 1 px text
baseline shift lights up every glyph, which is exactly how a *correct* gridline fix
scored as a regression below. Two gates replace it, both from **[pptx-renderer]**:

- **SSIM ≥ 0.95**, for "is the same thing in the same place" — catches wrong geometry,
  missing elements and layout shifts; ignores a hairline of antialiasing.
- **Colour histogram correlation ≥ 0.80**, position-blind, for theme resolution,
  gradients and tint/shade — the class of bug SSIM on greyscale barely sees.

Both are taken over foreground pixels only (grey < 245), and a slide under 1.5%
foreground is declared unscoreable rather than left to noise. Foreground IoU is
deliberately *not* used: upstream dropped it because a thin stroke loses half its IoU to
one pixel of antialiasing, and table gridlines are exactly that shape of problem.

A stored baseline fails on a drop of more than 0.02 SSIM. It is only compared against a
run with the same **font profile** — which faces the deck asks for, which of those the
host can supply, hashed — because a score taken with different fonts is not a score of
this library. That turned out to matter more than anything else in the file: see the
font column in the table above.

Pure NumPy, dev-only, behind the `fidelity` extra; the library stays standard-library.

### 0.4 Fixture corpus

Eight fixtures is thin, **only three of them can be scored** (see the baseline table
above), and **none contain SmartArt** — Phase 2 needed inputs before it needed code, and
got them from 46 real decks outside the repository instead. Generate a synthetic corpus with `python-pptx`, one feature per slide (each preset
family, each fill type, each bullet scheme, each table configuration). **[pptx-renderer]**
does this with a case generator and a support catalogue; the generated-corpus idea ports
directly even though their generator does not.

---

## Phase 1 — Parsed but not rendered — **done**

**Effort: M total. Highest fidelity-per-line in the whole plan.**

Fields the parser already read and the resolver already carried, which the renderer then
ignored. All of them now render, with two deliberate exceptions noted at the end.

| Gap | State |
| --- | --- |
| **Table styles** (`tableStyles.xml`, `a:tblStyle`) | Done — custom styles read, all 74 built-ins carried |
| **Tab stops** (`a:tabLst`, `defTabSz`) | Done |
| **Text highlight** (`a:highlight`) | Done — drawn as a rect behind the text |
| **Underline styles** (`u="dbl"`, `"wavy"`, `"dotted"` …) | Done — `text-decoration-style` |
| **Complex-script fonts** (`a:cs`) | Done — last in the `font-family` stack |
| **Text body rotation** (`a:bodyPr@rot`) | Done |
| **Hidden shapes** (`cNvPr@hidden`) | Done |
| **Image tile / stretch** | Done |
| **Multi-column text** (`a:bodyPr@numCol`) | Done, except mid-paragraph breaks |
| **Justified text** (`algn="just"`) | **Deferred** — see below |

### Correction: built-in table styles are not in the file

The earlier version of this section said banding, header row and borders "all come from
`ppt/tableStyles.xml` keyed by GUID". That is wrong, and it is the single most important
thing to know before touching tables.

**PowerPoint never writes a built-in style's definition into the file** — not even for a
style a table in the deck is actually using. Confirmed by round-tripping a deck through
PowerPoint 16.x: it rewrote `tableStyles.xml` as an empty element carrying nothing but
`def="{5C22544A-…}"`, the id of "Medium Style 2 - Accent 1". Three of the six fixtures
have exactly that shape. A reader for the part is necessary — Google Slides and Keynote
*do* write custom styles out — but it is nowhere near sufficient, and a catalogue of the
built-ins has to be carried.

The definitions are not on disk anywhere either; the application bundle was searched.
They were therefore **measured from PowerPoint's own rendering**, by
`tools/derive_table_styles.py`:

1. a sheet of 546 swatches, each filled with a known OOXML colour expression, gives an
   exact colour → expression dictionary — which is what turns a sampled `#cfd5ea` into
   `accent1 tint 40%` rather than into a guess;
2. a probe deck with one slide per candidate GUID and three tables per slide, laid out so
   every conditional region lands at a known row and column;
3. each probe rendered **twice, over a white background and over a black one**, so that
   `observed = colour × alpha + backdrop × (1 − alpha)` becomes two equations in two
   unknowns and alpha falls out. This matters: the whole Light Style family bands with
   20%-alpha black, which over a single background is indistinguishable from an opaque
   grey.

Two cross-checks keep guesses out of the result. A GUID PowerPoint does not recognise
renders exactly like "No Style, Table Grid" — that is its fallback — which caught three
candidate GUIDs that were wrong. And within a family the six accent variants must agree
once the accent number is factored out, which caught five stray colour matches.

**The catalogue is now complete: all 74.** The last two were absent only because their
GUIDs were unknown — the measurement was never the obstacle. "Medium Style 1"
(`{793D81CF-…}`) was read out of PowerPoint's own executable, which carries fifteen
brace-delimited GUIDs, fourteen of them table styles already catalogued here; that was
the fifteenth. "Light Style 1 - Accent 4" (`{D27102A9-…}`) came from
`aiden0z/pptx-renderer`'s published list and so was never more than a candidate. Both
then had to pass the same two cross-checks as everything else, and did: neither measures
like the unrecognised-GUID fallback, and each came out with its family's exact shape
carrying a colour the tool was never told to expect — accent 4 in the Light Style 1
shape, `dk1` where the accents sit in the Medium Style 1 shape, which is what an
accent-less base variant looks like elsewhere (compare "Dark Style 2" against "Dark Style
2 - Accent 1/Accent 2"). Re-deriving the whole catalogue in the same run reproduced the
other 72 entries byte for byte.

**Refuted: "Dark Style 2" is not missing its fourth accent.** PowerPoint pairs that
family's accents as "Accent 1/Accent 2", "Accent 3/Accent 4" and "Accent 5/Accent 6", so
it has three accent variants by design, not six, and all three are present. The family
count has to allow for it or the catalogue reads as two short forever.

A table naming a GUID that is in neither the deck nor the catalogue still falls back to
no style, exactly as PowerPoint does — but it now raises a `table-style-unknown` warning
while doing it. The render is defensible; the silence was not. An unstyled table is
pixel-identical to a table whose style genuinely carries nothing, so with no warning
there is no way to tell a correct render from one that dropped every band and header
rule. Same shape of bug as the substitutions `font-substituted` exists for.

### Also found on the way

Two things turned up that were not in this plan:

- **`tint` and `shade` were wrong in two independent ways** — the value was read as how
  far the colour *moves* rather than how much of it *survives*, and the blend was done in
  sRGB where PowerPoint does it in linear light. 284 of the 546 swatches differed from
  PowerPoint by more than 2/255 before the fix, several by more than 150. None do now.
  `lumMod`/`lumOff`, checked the same way, were already exact.
- **`a:clrChange` is parsed, resolved, and has no renderer at all** — a gap this section
  missed. It is now recorded in `UNRENDERED_FIELDS`, and is a Phase 5 item.

### Correction: the `real-basic-theme` regression was font metrics, not tables

Phase 1 recorded that this fixture's gridlines were drawn "a couple of pixels lower than
PowerPoint" because our CJK line height was "about 7% short". The regression was real
and reproduced exactly (1.84% → 2.07%, SSIM 0.946 → 0.938, all of it on slide 2). The
explanation was wrong in both direction and cause: the rules were drawn *higher*, not
lower, and the deck's Japanese text had nothing to do with it.

Measuring the rendered rules put them at y = 528/582/635.5/689 against PowerPoint's
528.5/584/639/695 — rows 53.7 px where PowerPoint draws 55.5 px, short by the same
amount every row and so cumulative. Working back through the cell margins, PowerPoint
wanted a 1.2016 em line box where we computed 1.1172 em, which is Liberation Sans's
`hhea` ascent + descent. The 7% was real; the reason was not.

So the line box was measured directly, by exporting probe decks and reading the line
advance off the raster:

> **PowerPoint's single-spaced line box is 1.2 × the font size, and the typeface has
> nothing to do with it.** Arial, Calibri, Times New Roman, Courier New, Aptos, Aptos
> Display, Lato, Raleway, MS Gothic, Meiryo and Noto Sans JP all measured 1.2 em, at
> 14 pt and 28 pt, with Latin text and with Japanese — although their real ascent +
> descent ranges from 1.00 em (MS Gothic) to 1.45 em (Noto Sans JP). `a:lnSpc`
> percentages multiply *that*: 150% measured 1.8 em, 90% measured 1.08 em. `a:spcPts` is
> literal and ignores it.

The first baseline has two rules, and they do not meet:

> At or below 100% the baseline hangs off the **bottom** of the line box, one font
> descent up, so the slack becomes leading above the text — 14 pt for 14 pt Arial,
> 14 pt for Times New Roman, 13 pt for Calibri. Above 100% the face drops out entirely
> and the baseline lands at **three quarters of the line box**: Arial and Calibri both
> measured 19 pt at 150% despite different descents. Going from 100% to 105% therefore
> moves the baseline *up*, which looked like a bad measurement until the second rule
> explained it.

Against the probe, our first baseline was out by up to 44 px (at 2560 px wide) and is
now out by at most 3. The old model — ascent × font size, line box from the face's own
metrics — was wrong for every font in the table, by −7% for Arial and +21% for Noto
Sans JP.

The corpus percentages *rejected* this fix at first, and that was the second lesson: on
a host with none of the decks' fonts installed, the end-to-end score cannot adjudicate a
layout change at all. The probe decks, which use only Arial and Times New Roman, could.

### The audit is now a test

`tests/test_render.py` walks the render model and fails on any field the renderer never
reads — which is how the table above was produced in the first place. Fields that are
deliberately not the renderer's business are listed with a reason each, and a second test
fails if one of those entries stops naming a real field, so the list cannot rot into a
place regressions hide.

### Left undone, on purpose

- **Justified text** (`algn="just"`). SVG has no `text-align: justify`, so it means
  per-word `x` positioning: the wrapper would have to hand out word positions instead of
  line segments, and every consumer of `LineSegment` would change with it. Deferred, as
  this plan suggested.
- **Mid-paragraph column breaks.** PowerPoint splits a paragraph across a column
  boundary; we move the whole paragraph. Seen side by side on a probe deck: PowerPoint
  starts paragraph 4 at the bottom of the left column and finishes it at the top of the
  right one. Fixing it means making the wrapper's line list the unit of column layout,
  and deciding what a continuation does about its bullet.
- **Tile flip and alignment** (`a:tile@flip`, `@algn`). SVG patterns cannot mirror
  alternate tiles.

### Fixtures

`table test.pptx` was added to cover the built-in style catalogue, and it now does: it
names `{5C22544A-...}` ("Medium Style 2 - Accent 1"), sets `firstRow` and `bandRow`, and
gives no cell an explicit fill, so every colour it renders comes out of
`parse/table_styles_builtin.py`. Exported through PowerPoint and sampled cell by cell,
all twenty-five fills are byte-identical — header `#156082`, bands `#CCD2D8` and
`#E7EAED` — with white rules throughout; the only differences anywhere in the table are
sub-pixel antialiasing on the rules themselves. That is **one** GUID of the seventy-two
carried. The rest came from the same measurement process but no fixture exercises them,
and they should not be described as verified.

None of the seven other fixtures uses a tab stop, a highlight, `bodyPr@rot`, a hidden shape, a
non-single underline or multiple columns — checked, not assumed. Those are validated
against PowerPoint using the deck `tools/make_feature_probe.py` builds, which is also
where the mid-paragraph column finding above comes from. A synthetic corpus of the kind
Phase 0.4 describes would subsume it.

---

## Phase 2 — SmartArt — **done, and validated against 46 real decks**

**Effort: S–M. The code landed at the S end -- ~120 lines in `resolve/view.py`, no new
parser -- but the estimate was only right about the code. Getting it to work on a file
PowerPoint actually wrote took a second pass, because the first one was verified only
against fixtures we built ourselves and those encoded our own misreading of the format.**

> **Correction to an earlier assessment.** I previously said SmartArt "needs a diagram
> layout engine". For the common case it does not. PowerPoint caches a fully laid-out
> DrawingML rendering of every diagram, and pptx-glimpse simply reads it —
> `computeSmartArtElement` is ~45 lines. Layout is only needed for files lacking the cache.

The lookup chain, confirmed against pptx-glimpse's source:

```
p:graphicFrame
  └─ a:graphicData[uri=…/diagram]/dgm:relIds@r:dm     ← we already parse this
       └─ ppt/diagrams/data1.xml                       ← the data model part
            └─ its _rels, type …/2007/relationships/diagramDrawing
                 └─ ppt/diagrams/drawing1.xml
                      └─ dsp:drawing/dsp:spTree        ← plain DrawingML!
```

That last shape tree is ordinary shapes, text and geometry — `parse/shapes.py:parse_shape_tree`
handles it as-is.

**Work:** all four steps done in `resolve/view.py:_resolve_diagram`. Both relationship
spellings (`schemas.microsoft.com/office/2007/…` and `purl.oclc.org/ooxml/…`) are accepted;
children resolve with `part_path` pointed at the drawing part, so a diagram's own images
are found; the warning and empty frame survive for a deck whose cached drawing is missing,
empty or corrupt.

**Two things worth knowing that the plan did not mention:**

- **Every cached shape has `cNvPr@id="0"`.** PowerPoint carries identity in `modelId`
  (a GUID), not the DrawingML id, so taking the id at face value would have given every
  node in a diagram the same `data-pptx-id`. Ids are rebuilt from the frame's id plus the
  child's index (`256.7001/0/0`), which is unique, stable and still addressable.
- **The `grpSpPr/a:xfrm` fallback is not `replace(transform)`.** A regular group with no
  `chOff` has children in absolute slide coordinates, so mapping it to the outer transform
  is right. A diagram's children are relative to the frame's own origin, so the fallback
  has to be `chOff=(0,0)`, `chExt=` the frame extent. Using the group default here shifts
  the whole diagram off the slide by the frame's offset.

**[pptx-renderer]** notes that diagram groups need "diagram-specific compensation requiring
matching layout provenance". Nothing beyond the above was needed for the cases tested —
but see the caveat, and expect this to be where a real file diverges. Their fallback when
no drawing exists is an EMF preview, which Phase 4 now unlocks for us too, though nothing
wires the two together yet (a `dgm:relIds` frame has no blip to hand to `_resolve_image`).

**Done when:** a SmartArt cycle/hierarchy renders its nodes and connectors ✅ (nodes, text,
fills, geometry and embedded pictures; connectors are ordinary shapes in the cache and
need no special handling); a deck with the drawing part deleted still renders and still
warns ✅.

**Now verified against real input, and it was broken.** The earlier caveat here said the
implementation was untested against genuine PowerPoint output. It was, and it rendered
*nothing at all* on every real deck. LibreOffice's test corpus
(`sd/qa/unit/data/pptx/smartart*`) has 46 SmartArt decks written by PowerPoint 12.0 to
16.0; all 46 produced an empty frame.

**The lookup was in the wrong place.** `dgm:relIds` names the data model, layout, quick
style and colours — and not the drawing, because the cached drawing was added to the
format after `relIds` was specified. Microsoft keyed it through an extension instead:

```
slide rels --r:dm--------------> ppt/diagrams/data1.xml
    data1.xml dgm:extLst/dsp:dataModelExt@relId = "rId6"
                                         |
slide rels --rId6 (diagramDrawing)-------+--> ppt/diagrams/drawing1.xml
```

So the id is written in the *data* part and resolved against the *slide's*
relationships. All 27 decks that carry a cached drawing do it this way, and exactly one
has a `ppt/diagrams/_rels/data1.xml.rels` at all — which is what the original code, and
the plan above, assumed was the only route. Two fallbacks remain for producers that do
something else: the data part's own relationships, then a lone diagram-drawing
relationship on the slide, used only when there is exactly one (two frames on a slide
with nothing to key them by would be a coin toss).

**Where the corpus stands now:**

| | Decks | Outcome |
| --- | --- | --- |
| Cached drawing with shapes in it | 13 | Renders shapes, text, fills, geometry |
| Cached drawing present but an empty `spTree` | 13 | Nothing to draw; warns `diagram-no-cached-drawing` |
| No drawing part at all | 20 | Nothing to draw; warns `diagram-no-cached-drawing` |

Every deck that carries usable cached content now renders it, and no frame comes out
blank without saying why. `smartart-font-size.pptx` carries three diagram frames on one
slide and warns three times, independently.

**`dsp:txXfrm` was the other thing only a real deck could show.** A diagram shape places
its text box separately from the shape — a Venn ring's label goes in the sliver that ring
does not share. 13 of the 26 cached drawings use it, exactly the 13 that render. Ignoring
it stacked every label at its shape's bounding-box corner.

#### What is left

Twenty decks have no cached drawing because Office 2007 did not always write one, and 13
more have an empty one. Drawing them means implementing the layout algorithms in
`dgm:layoutDef` — a diagram engine, a project in its own right, and **deliberately not
attempted**. The honest interim answer is an empty frame plus a warning naming the
reason, which is what happens now.

**Fixtures are still not committed.** The 46 decks are validated against locally but not
checked in: LibreOffice is MPL-2.0 / LGPLv3+ and redistribution with attribution is
defensible, but many of these files began life as bug-report attachments and that is a
call for the project to make, not for a contributor. The alternative that is unambiguously
ours to ship is `sld.Shapes.AddSmartArt` in VBA — PowerPoint authors the diagram, we own
the output — which needs a macro-enabled host `.pptm` built by hand. `tests/test_diagram.py`
still ends with a skipped test that activates when `tests/fixtures/real-smartart.pptx`
appears.

---

## Phase 3 — Charts

**Effort: XL. 3.1 and the `barChart` half of 3.2 are done; the rest is not.**

No shortcut: unlike SmartArt, PowerPoint does *not* cache a rendered chart. The
`c:chartSpace` part holds data plus styling, and the renderer must do axis scaling, tick
selection and plotting. pptx-glimpse spends ~1,400 lines on this plus ~450 on data.

A real fixture (`real-financial-report.pptx`, `ppt/charts/chart1.xml`, 6 KB) uses 67
distinct `c:` elements, a useful sense of the minimum vocabulary:
`barChart barDir grouping gapWidth overlap ser cat val numRef numCache strRef strCache pt
ptCount catAx valAx axId axPos scaling orientation crosses crossAx crossBetween
majorGridlines tickLblPos numFmt formatCode legend legendPos dLbls …`

### Build or delegate?

**[pptx-renderer]** delegates to [ECharts](https://echarts.apache.org/). The Python
equivalents are matplotlib (SVG backend) and pygal (pure-Python SVG). Hand-rolling was
the call and it held: the whole of `resolve/chart.py` lowers a chart to the
`ShapeElement`s and `ConnectorElement`s the renderer already draws, so there is no second
code path, no dependency, and the chart composes into the slide's transform stack for
free. `render/svg.py` gained one line.

The reference's `postProcess.ts` — 530 lines of pushing ECharts output back towards
PowerPoint — is the cost that decision avoids, and reading it was still worth it: it is
where the real PowerPoint behaviours are written down.

### 3.1 Chart data model and reader — **done**

`parse/chart.py` reads `c:chartSpace` into a `SourceChart`; `model.ChartData` /
`ChartSeries` / `ChartAxisScale` carry the result out to `convert_pptx_to_model` callers,
alongside the drawn primitives. Verified against all six charts in the corpus.

Three things the obvious reading gets wrong, each found in a fixture:

* **Chart booleans default to *true* when `val` is absent.** `<c:delete/>` deletes an axis
  and `<c:overlay/>` overlays a legend. The `is_true` used for ordinary DrawingML
  attributes reads both as false.
* **Categories are not always `c:strRef`.** Every category axis in
  `real-financial-report.pptx` is a `c:multiLvlStrRef`, and a numeric or date axis arrives
  as `c:numRef`. A reader that only looks at `c:strRef` silently draws a bare axis.
* **A missing `c:pt` is a blank cell, not a zero.** `c:dispBlanksAs` draws gap, zero and
  span differently, so blanks survive as `None` rather than being flattened at parse time.

`c:ptCount` is attacker-controlled and is clamped rather than trusted. The chart's own
`c:clrMapOvr` **replaces** the slide's rather than layering on it — a slide that remaps
`bg1`/`tx1` for its own shapes does not remap them for a chart that declares its own
mapping. Both **[pptx-renderer]** and this roadmap flagged it; it is invisible until a
deck does both at once.

### 3.2 Renderer — five types **done**, the rest not started

1. ✅ `barChart` — clustered, stacked, percentStacked, `barDir` col and bar
2. ✅ `lineChart` — markers, smoothing, blanks, the real fixture on slide 2
3. ✅ `pieChart` / `doughnutChart` — hole, rings, explosion, the real fixture on slide 3
4. ✅ `radarChart` — standard, marker and filled, the real fixture on slide 4
5. `areaChart`
6. `scatterChart` / `bubbleChart`
7. `stockChart`, `surfaceChart`, `ofPieChart` — long tail; defer

**No chart in the corpus warns `chart-unsupported-type` any more.**

Data labels are drawn for all five, at every `c:dLblPos` each type accepts — except
the radar's, whose placement no export has ever shown; see below.

#### Polar layout, measured

The core comes off `real-financial-report.pptx`'s own doughnut, which is exact enough to
pin it without a probe at all:

* the plot region is the **same** one a bar chart computes — edge insets plus the legend
  band, whose 113.98 pt on that deck is the identical number its bar charts reserve;
* the radius is **half the shorter side** of that region, centred in it;
* **angle zero is 12 o'clock and slices run clockwise**; its first slice ends at
  154.80° for a 43% share, which is 43% of 360;
* **`c:holeSize` is a percentage of the outer radius**, not of the frame or the diameter.

A twelve-chart probe adds the rest, and two of them contradict the obvious reading:

* **A `c:doughnutChart` stating no `c:holeSize` draws as a solid pie.** ECMA-376 documents
  a default of 10; the wedge PowerPoint emitted closes through the centre with no inner
  arc at all.
* **`varyColors` is the default for a pie and not for a bar.** Probe pies stating no
  `c:varyColors` came out accent1..accent4 across their slices, and a two-ring doughnut
  cycled the same four in *both* rings — per point, not per series.
* Several series make **concentric rings**, innermost first, splitting the space between
  the hole and the outer radius evenly.
* **`c:explosion` shrinks the radius by `1/(1+e)`** and offsets each slice by `e` of the
  *shrunk* radius along its own bisector — both halves measured to 0.01 pt.
* Slice labels sit along the bisector at a fraction of the radius: `ctr` exactly 0.5,
  `inEnd` 0.856, `outEnd` 1.020, `bestFit` 0.710.
* **`showPercent` percentages add up to 100**: three equal values are labelled 34%, 33%,
  33%, where rounding each on its own gives 33% three times and totals 99.
* A pie legends its **categories**, not its series.

#### Radar, measured

`radarChart` is done, and the corpus measured most of it. That is worth stating plainly
because the brief for this work assumed it could not: `real-financial-report.pptx` is
skipped by `tools/fidelity.py` for want of Noto Sans JP, but **vector geometry does not
move when a font is substituted**, so its chart5 — the only radar in the corpus — is
readable out of PowerPoint's own PDF at exact coordinates. A skipped *score* is not a
skipped *measurement*. Three probe decks of six charts each pinned the rest.

The polar conventions turn out to be the pie's, and that is a measurement rather than an
assumption — the four-category probe puts its vertices due north, east, south and west:

* **angle zero is twelve o'clock and categories run clockwise**, 360/n apart;
* the centre is the centre of the **same plot region a pie computes**, to within 0.21 pt
  on the legend probe and 0.1 pt on the corpus radar;
* a point sits at `(value − minimum) / (maximum − minimum)` of the radius.

Six things the schema does not say, each from a probe that contradicts the obvious
reading:

* **The web is polygonal and follows the category count** — 3, 5, 6 and 8 categories gave
  triangles, pentagons, hexagons and octagons — and **it is drawn whether or not the file
  asks for it**. A probe with no `c:majorGridlines` at all still drew every ring, and so
  did one with `<c:delete val="1"/>` on the value axis. Only the *styling* comes from
  `c:majorGridlines`.
* **The spokes do not.** No probe without a `c:spPr` on its category axis drew any; the
  corpus radar, whose category axis states `<a:ln w="12700">` in #888888, drew six in
  exactly that. So the radial lines are the category axis' own line and **its default is
  none** — the opposite of a bar chart, whose default axis line is black at 0.5 pt.
* **`standard` and `marker` draw an identical picture**, markers included, though
  ECMA-376 says a `standard` radar has none. Two probes, byte-identical output.
* **`filled` draws only the fill.** No markers, and no outline unless the series states an
  `a:ln` of its own: the probe, which states none, emits a bare `f`; the corpus radar,
  which states `w="25400"`, is stroked at 2 pt. That pair is the whole rule.
* **A radar's value axis stops at the data where a bar's goes a unit past it.** 0..5 of
  data draws five rings with the outermost through the largest point, where the same data
  on a bar gives 0..6. One discriminating observation, and it is the only difference —
  the unit is chosen identically.
* **Category labels wrap; they do not rotate.** Every `rot` in the long-label probe is
  zero and "Category Three" came back split as "Category" / "Three".

The value-axis labels sit up the twelve o'clock spoke, right-aligned, with their right
edge **two widths of the digit zero** to the left of it. Six charts across two faces and
three sizes plus the corpus radar's Arial 12 pt, worst residual 0.14 pt. Reading it as an
em fraction does not work — it is 1.069 em in Aptos and 1.112 em in Arial, and the
difference is exactly twice the difference between the two faces' digit widths.

The radius is the fitted part. Two constraints, smaller wins:

* **vertically**, it falls short of half the region by a reserve that is a function of the
  category label's line box — `1.2578 × line_height − 5.2501`, fitted to five probes with
  a worst residual of **0.089 pt**:

  | face | size | line box | reserve |
  | --- | --- | --- | --- |
  | Aptos | 8 | 9.766 | 7.071 |
  | Aptos | 10 | 12.207 | 10.081 |
  | Aptos | 14 | 17.090 | 16.191 |
  | Arial | 10 | 11.172 | 8.751 |
  | Arial | 14 | 15.641 | 14.511 |

  The slope is **not** 1, which is why no "leave one line of room" rule reproduces the
  set; and a sixth probe pins the other end — with the category axis deleted and so no
  labels at all the radius came out 79.44 pt against a half-region of 79.551, i.e. the
  reserve goes to zero.

* **horizontally**, the widest label must still fit the region, which for a spoke at angle
  θ bounds the radius by `(half_width − 2.85 − label_width) / |sin θ|`. Three probes are
  bound this way and land within 0.7 pt.

Ten cases land within 0.14 pt of PowerPoint. **The reserve counts one line even when a
label wraps**, and that is the measurement rather than an oversight: the one multi-line
observation has a radius of 60.24 pt, which the horizontal constraint reproduces exactly
and whose vertical reserve is therefore at most 19.31 pt — less than the 20.16 pt that two
lines of the fitted per-line reserve would ask for. The per-line rule is contradicted by
that probe, so it was not shipped.

Two smaller measurements that belong to other chart types too:

* **A line-style radar's legend key is a line with its marker on it**, not a swatch:
  19.200 pt of rule then 2.025 pt before the text, at 10 pt — against the swatch's
  5.492 + 2.371. That is 13.4 pt of band width. Four line-chart legends have since
  confirmed both numbers and refuted the *scaling*: they were carried as 1.920 and
  0.2025 ems on the assumption that a single 10 pt measurement scaled, and a 14 pt legend
  came back at the same 19.200 and 2.025. They are absolute points, and a `lineChart`
  now draws them too.
* **A radar series stating no `c:size` gets a 6 pt marker**, not ECMA-376's 7. The probe's
  second series measured 6.0 pt square; its first measured 5.76 pt across, which is the
  same 6 pt box with a diamond's tips falling inside PowerPoint's 0.24 pt output grid.
  `lineChart` still uses 7, which has never been measured either way.

What is wrong or unmeasured in the radar path:

* **The corpus radar's radius is 6.8 pt too large, and the cause is measured.** Its
  category labels are Japanese; PowerPoint laid them out in a substituted CJK face whose
  line box is about 1.57 em, while `font_box` reads the `<a:latin typeface="Arial"/>` the
  axis names, whose line box is 1.117 em. Feeding the fitted reserve the CJK line height
  reproduces PowerPoint's 18.44 pt exactly; feeding it Arial's gives 11.61. Fixing it
  means `font_box` knowing which face a CJK run actually resolves to, which is a
  `text/`-and-`fonts/` question, not a chart one.
* **Radar data labels are not measured.** The corpus radar's `c:dLbls` sets every
  `c:show*` to 0 and no probe turned one on, so where PowerPoint puts one is unknown. They
  are drawn — pushed radially out from the point, one marker clear — because silently
  dropping flags a file states is the worse failure, but that placement is a guess and the
  code says so. It is the only thing in the radar path without a measurement behind it.
* **`dispBlanksAs` on a radar is not measured.** A blank breaks the ring, mirroring the
  line chart. A ring is a cycle, so the stretch running through index 0 is one run and not
  two — that was a real defect the test caught, and `_rotate_past_blank` is why it is not
  one now.
* **The label's vertical anchoring is about 1.5 pt loose.** The four horizontal directions
  land within 0.15 pt of the fitted 2.85 pt gap; the two vertical ones measured 4.49 pt
  above the top vertex and 2.29 pt below the bottom one. The split is consistent with
  PowerPoint's line box being about a point taller than our metrics give — a font
  discrepancy rather than a second layout rule — so one constant is used for all four.
* **The wrap threshold is one bracket.** 43.72 pt of label stayed on one line and 59.10 pt
  wrapped, on a 198.47 pt region; 0.25 is the round number inside (0.2203, 0.2978].

#### Rotated category labels, measured

The category labels turn when they will not fit, and until this was measured it was the
worst-looking thing any corpus deck did: `real-financial-report.pptx` slide 3 reserved
27.3 pt under its plot against PowerPoint's 69.5 and printed five Japanese labels on top
of each other.

Twelve probe charts across two decks, five categories each in the same
220.4724 x 181.1024 pt frame with 10 pt Aptos labels, exported and read back as exact
vector coordinates. One deck sweeps the label from a fifth of its band to four times it;
the other straddles exactly one band, which is what turns the threshold from a guess into
a window.

* **The rule is that the widest label is wider than its own band** — and, once
  *Wrapped category labels* below is taken into account, specifically the widest
  unbreakable **token**, which is the same thing for every label in this sweep because
  none of them contains a space. On a 37.68 pt band a 36.62 pt label stayed level and a
  38.59 pt one turned, so the threshold is inside (0.972, 1.024] and one band width is
  the middle of it rather than a round number picked for tidiness. The *widest* label is
  what counts, not the first: the deck whose five labels differ only by a trailing letter
  left a 0.73 pt residual until that was fixed, which is exactly the 1.01 pt difference
  between Aptos' `A` and its `D` times sin 45.
* **The angle snaps to 45°.** Everything from 1.02 band widths to 4.18 came out at
  exactly 45, reading up to the right — `rot="-2700000"` in DrawingML terms. No
  intermediate angle appears anywhere in that range and nothing goes to 90. The corpus
  deck confirms it independently at 12 pt, where its export's text matrix is 8.4853,
  which is `12 cos 45`.
* **The band under the plot becomes `21.39 + widest x sin 45`**, replacing the level
  `6.5 + lineHeight + 0.615 em` outright rather than adding to it. Six probes, worst
  residual **0.015 pt**.
* **The label hangs off the far end of its rotated baseline**, 2.0 pt right of its band's
  centre and 12.7 pt below the axis. Six probes, spread under 0.15 pt.
* Neither fixture carries an explicit `rot=` on `a:bodyPr`, so all of this is
  PowerPoint's own decision rather than anything the file asks for.

And a contrast worth keeping, because it says this is a *category-axis* behaviour rather
than a general label one: **a radar facing the same problem wraps instead of rotating.**
Every `rot` in the radar long-label probe is zero and "Category Three" came back split
over two lines as "Category" / "Three". That observation turned out to be the thread
worth pulling: a bar chart wraps first too, and rotates only when wrapping cannot save
the label. See *Wrapped category labels* below.

Two pieces are measured and **not** shipped, both because a probe refutes the obvious
rule:

* **The cap.** PowerPoint reserved 85.63 pt for a label 4.18 bands wide — *less* than the
  86.36 pt it gave the 2.92-band label one step below it. No clamp on the width produces
  both, so whatever it does past about 90 pt of label was not identified. Ours keeps
  going up the fitted line; a test asserts that number so the divergence is recorded
  rather than latent.
* **The left inset.** It grows too once the first label reaches past the plot: 21.07 pt
  level, then 21.68, 36.44 and 54.19 as the label widens, with the last two probes
  sharing a value the way the bottom cap does. Solving it for the minimum pen position
  gives 10.02, 8.56 and 6.78 — not one number — so it is left alone and our plot comes
  out wider than PowerPoint's on the two most crowded probes.

**What it bought.** On chart3 the bottom inset goes 27.29 -> 89.3 against PowerPoint's
69.538, so the error more than halves and the labels stop colliding. The 19.8 pt left is
not this rule: it is the width we measure a CJK label at. `プラットフォーム` comes out
96 pt through the `<a:latin typeface="Arial"/>` the axis names, where PowerPoint laid it
out in a substituted CJK face at about 68 — the same `font_box`/`text_width` gap the
radar's radius has on the same deck, and the largest single error left in any chart.

`authoring-integration` holds at 0.9327 and `table-test` at 0.9734, unchanged to four
decimals: neither has a label wide enough to turn, so nothing this work did moves a
scored number. The probe decks that measured it are throwaway and were deleted.

#### Wrapped category labels, and why the rotation rule was right in the wrong domain

The rotation sweep above measured a real behaviour but measured it on labels built from
one repeated character. `MMMMM` has nowhere to break, so rotation was the only move
PowerPoint had, and the rule that came out — *turn when the widest label is wider than
its band* — is only half of what it does. `real-college-template.pptx` slide 4 is where
that showed: its eleventh category is `2012 (Proj)`, 54.41 pt on a 54.14 pt band, so
every one of its eleven labels turned, and PowerPoint's own export leaves all eleven
level and breaks that one over two lines.

**Wrapping comes first; rotation is the last resort.** Three throwaway decks, 57 bar
charts, same 220.4724 x 181.1024 pt frame and five categories as the rotation sweep, on a
37.761 pt band read off the axis rule.

* **The test is the widest unbreakable token, not the widest label.** `MMM MM` is
  44.43 pt — 1.18 bands — and came back level on two lines; `MMMMM` is 41.65 pt, 1.10
  bands, and turned. Straddling the band with a token instead of a label gives
  (0.9857, 1.0151] — `xxxxxxxi M` at 37.22 pt wrapped, `HHHHHi M` at 38.33 pt turned —
  and one band width is the only value inside both that window and the unbroken sweep's
  (0.972, 1.024].
* **The case that separates it from every other candidate rule** is
  `Fiscal MMMMMMMMMMMM 2012`. It has two spaces in it, so "any break opportunity means
  wrap" predicts level; PowerPoint turned it, because breaking it still leaves a 99.96 pt
  token. (It also truncated it to `Fiscal…`, which is its own unmeasured behaviour.)
* **The break set is whitespace, not Unicode line breaking.** `MMM-MM`, `MMM/MM`,
  `MMM,MM`, `MMM_MM` and `MMM–MM` were each 44–47 pt on that band and PowerPoint turned
  all five rather than breaking them, and `プラットフォーム` — 80 pt of CJK, which any
  line-breaking algorithm would break anywhere — turned as well. The surprise is U+00A0:
  `MMM<nbsp>MM` broke at the no-break space exactly as the plain space did. Tab and the
  rest of `str.split`'s whitespace are **not measured**.
* **One label that must turn turns all of them.** Four `MMM MM` beside one `MMMMM` came
  back with all five at 45°, none of them broken. Rotation is a decision for the axis.
* **A label that fits stays on one line** while its neighbour wraps, and sits on the
  *first* row of the block: the block is top-aligned, and the first baseline is where a
  one-line label's would be — 15.03, 15.05 and 15.07 pt under the axis at one, two and
  three lines, one number inside PowerPoint's 0.12 pt grid.
* **The band is the level band plus one line box per extra line.** A four-rung ladder in
  five faces, one token per line by construction. PowerPoint's band grew by the same
  amount from one line to two, two to three and three to four in every face, so this is a
  straight line rather than a fit:

  | face | per extra line | our line box | residual |
  | --- | --- | --- | --- |
  | Calibri | 12.205 | 12.207 | −0.002 |
  | Aptos | 12.205 | 12.207 | −0.002 |
  | Courier New | 11.330 | 11.328 | +0.002 |
  | Times New Roman | 11.075 | 11.074 | +0.001 |
  | Arial | 11.500 | 11.172 | **+0.328** |

**Arial is the one refutation, and it has a name.** PowerPoint's pitch is the face's full
`hhea` line spacing — ascender plus descender plus *lineGap* — and Arial is the only one
of the five whose lineGap is not zero: 67 units of 2048 is 0.328 pt at 10 pt, exactly the
residual. That rule is **not** shipped, because `text/metrics.py` carries no lineGap and
the file it would be generated from disagrees with the file PowerPoint used: our Tinos
substitute has 87 where Office's own `times.ttf`, inside the app bundle, has 0. Adding
the gap would trade a 0.33 pt error on Arial for a 0.42 pt one on Times New Roman, so the
line box alone is shipped and the Arial shortfall is pinned by a test.

Two more things measured and deliberately not shipped:

* **Past six lines PowerPoint stops wrapping.** Six tokens gave six lines and an 81.212 pt
  band, on the same straight line as one through four. Eight, twelve and twenty-four
  tokens all came back on **two** lines in a 35.212 pt band, each line four band widths
  wide and overlapping its neighbours. No rule reproduces both the linear part and that
  collapse, so the band stops growing at six and a test records that we are then 44 pt
  over. This is the same shape of problem as the rotated band's cap past 90 pt of label.
* **An explicit orientation on `a:bodyPr` forbids the turn, and PowerPoint drops labels
  instead.** `<a:bodyPr rot="0" vert="horz"/>` and `<a:bodyPr vert="horz"/>` both left
  `MMMMM` level on a band it does not fit — and PowerPoint printed only every *other*
  label, three of five, at twice the band pitch, rather than turning or overlapping them.
  With a break available (`MMM MM`) the same axis wrapped exactly as a free one did. So
  there is a third strategy under there — skipping to a wider tick-label interval — that
  no corpus deck reaches and that is not implemented. Note that
  `real-college-template.pptx`'s own chart carries `rot="0"`, so its labels would stay
  level under this rule too; the wrap rule above is what actually reproduces its output,
  and the two agree there.

**What it bought.** `real-college-template.pptx` slide 4 goes 0.3484 → **0.3721** SSIM
against the 0.3690 it scored before rotation existed, and its bottom band 59.9 → 37.8
against PowerPoint's own two-line block. Its histogram goes 0.9694 → **0.9899** against
0.9950, and it does not get all the way back: 0.9950 is reproduced *exactly* by reserving
one line instead of two, which is 11.5 pt shorter than the band PowerPoint drew. The
metric is not monotone in geometric accuracy on this slide — setting the plot bottom to
PowerPoint's own, to the pixel, scores 0.4704 SSIM and 0.9842 histogram — so the last
0.005 of histogram is only available by keeping a band that is measurably wrong. What is
actually left on that slide is the **manual legend layout**: its `c:legend` carries a
`c:manualLayout` we ignore, which is worth 4.8 pt of plot height and 9.5 pt of legend
baseline, and no rule was found that composes the plot area around a manually placed
legend (PowerPoint's plot bottom sits 13.4 pt above the legend's declared top, which is
neither the declared height nor our band).

`authoring-integration` holds at 0.9327 and `table-test` at 0.9895, unchanged to four
decimals. The probe decks are throwaway and were deleted.

#### The tick-density question, with four more observations that refute one more rule

The value-axis density rule is still unsolved, and radar adds four measurements that make
the "no single rule" verdict firmer rather than softer. Same frame, same data, the only
variable the label font:

| face | size | radius | PowerPoint | spacing | in line boxes | in ems |
| --- | --- | --- | --- | --- | --- | --- |
| Aptos | 10 | 69.47 | 0..5 by 1 | 13.89 | 1.138 | 1.389 |
| Arial | 10 | 70.80 | 0..5 by 1 | 14.16 | 1.267 | 1.416 |
| Aptos | 14 | 63.36 | 0..6 by 2 | 21.12 | refused 12.67 | refused 0.905 |
| Arial | 14 | 65.04 | 0..6 by 2 | 21.68 | refused 13.01 | refused 0.929 |

So a radar accepts 1.389 em of spacing and refuses 0.929 em, bracketing its threshold to
(0.929, 1.389). **The bar chart's bracket is (1.543, 1.610) em** — measured in the same
way, on the same machine — and the two do not overlap. Whatever the rule is, it is not one
number shared by both chart types. Today's behaviour is still "no density limit", which is
right for the 10 pt radars and draws five rings where PowerPoint draws three at 14 pt.
#### The corpus contains no drawn data label

Worth stating once, because two separate readings of the same files got it wrong. Five of
the six charts carry a `c:dLbls` block. Four state every `c:show*` flag as 0. The fifth,
chart4's doughnut, *does* set `showCatName` and `showPercent` — and then carries four
`c:dLbl` overrides, one per point, that set every flag back to 0. PowerPoint's export
confirms it: no label is drawn on any chart in the corpus. Counting `c:dLbls` elements
says nothing; the flags have to be read, and then the per-point overrides on top of them.

Data-label rendering is therefore verified **entirely against probes**.

Anything else warns `chart-unsupported-type` and draws an empty frame rather than a wrong
picture. The shared infrastructure — value domain, tick selection, number formatting,
gridlines, legend layout for all four `legendPos` values, plot-area rectangle — is built
and is what the other chart types will reuse.

#### Every constant was measured, and the measurement kept correcting the reasoning

Eighteen probe charts differing in one input each — title, legend on each of four sides,
8/10/14 pt text, gap width, tick marks, series count, horizontal bars, stacked and
percent-stacked grouping, negative values with and without `invertIfNegative`,
`varyColors` — were laid out at an identical frame size, exported by PowerPoint 16.106,
and read back out of the PDF as **exact vector coordinates**. Both sweeps are tests
(`tests/test_chart.py`), rebuilt from their XML generators, so no binary is committed.

Against `authoring-integration.pptx`, whose chart is half the slide:

| | SSIM | histogram |
| --- | --- | --- |
| before | 0.8379 | 0.8106 |
| after | **0.9327** | **0.9984** |
| the chart region alone, after | 0.9124 | 0.9957 (was 0.0499) |

Its plot rectangle, first bar, bar width and every text baseline land within 0.47 pt of
PowerPoint's; all seven gridlines and the category axis land on the same pixel row at
1280 px. The two bar charts in `real-financial-report.pptx` agree on all four insets to
within 0.14 pt and on the axis exactly.

What the measurement said that the reasoning did not:

* **The axis rule is not "aim for N ticks".** It is the plain power of ten below the span,
  halved when the span is under twice it. No tick target reproduces both 0..9 → 0..10 by 1
  (ten intervals) and 0..1842 → 0..2000 by 500 (four). Both ends round *strictly*
  outwards, so data topping out at 5 gets an axis to 6.
* **The tick-mark allowance is reserved whether or not tick marks are drawn.**
  `majorTickMark="none"` and `"out"` produced byte-identical plot rectangles.
* **The default axis and gridline colour is black at 0.5 pt, not grey.** Charts written by
  modern PowerPoint carry a `c:style` or a chart-style part that overrides this; none of
  the decks measured here does.
* **A chart's axis labels and its title resolve through different cascades.** With no
  `c:txPr` anywhere, PowerPoint drew `authoring-integration`'s axis labels in Aptos 10 pt
  (the theme's minor face, the chart default) and its title, whose `a:rPr` names nothing
  at all, in Arial 18 pt — the same fallback any unstyled text box gets.
* **`c:txPr` names a typeface as well as a size.** Reading only the size measured
  `real-financial-report`'s Arial labels in the theme face and put the plot area 1.7 pt
  out.
* **`varyColors` cycles the theme accents exactly.** **[pptx-renderer]** darkens them to
  88%; PowerPoint's export says #4472C4, #ED7D31, #A5A5A5 unmodified.
* **A negative bar is white with a *black* 0.75 pt outline**, and the series colour when
  `invertIfNegative` is 0.
* **`tickLblPos="nextTo"` means next to the axis, and the axis is at zero.** With negative
  values PowerPoint reserves no band under the plot at all and prints the category labels
  inside it, beside the zero line.
* **A horizontal value axis comes out coarser than a vertical one** for the same data on
  an axis of almost the same length, so it is not a density limit. One measurement only,
  and the code says so.

One rule resisted: **the vertical centring of a one-line label on a tick**. Five
measurements across two faces and three sizes fit none of half the cap height, half the
x-height, half the line box, or the centre of the digits' own ink. The constant is the
fitted mean and its worst residual is 0.61 pt.

#### The value axis' tick density is measured but **not solved**

`real-financial-report.pptx`'s line chart is the case that exposes it: 43 of data on a
73.6 pt plot, where PowerPoint draws **0..60 by 20** and the rule above gives 0..50 by 10.
Short plots draw fewer ticks, and the rule that decides how many is not known.

A six-cell probe, identical data (`3, 4, 5`) and 10 pt labels, shrinking frame:

| plot height | PowerPoint | intervals | spacing |
| --- | --- | --- | --- |
| 145.0 pt | 0..6 by 1 | 6 | 24.2 pt |
| 82.0 pt | 0..6 by 2 | 3 | 27.3 pt |
| 50.5 pt | 0..6 by 2 | 3 | 16.8 pt |
| 30.9 pt | 0..10 by 10 | 1 | 30.9 pt |
| 19.1 pt | 0..10 by 10 | 1 | 19.1 pt |
| 11.2 pt | 0..10 by 10 | 1 | 11.2 pt |

Those six alone are reproduced by "step the unit up the 1-2-5 ladder until the ticks are
at least 1.58 em apart", bracketed to (1.543, 1.610) by the 30.9 pt cell refusing 15.43 pt
and the 14 pt probe accepting 22.53 pt. **And the stacked probe refutes it**: a 0..10 axis
on a 145.0 pt plot with the same 10 pt font takes *ten* intervals at 14.5 pt spacing —
finer than the 15.43 pt the 30.9 pt cell rejected, on the same axis range and the same
font. No monotone spacing threshold produces both, so at least one more variable is
involved that this sweep did not vary.

A rule was written, measured against all of it, found to contradict the stacked case, and
**reverted rather than shipped**. Today's behaviour is "no density limit", which is right
everywhere except short plots. Whoever picks this up starts from the table above; the
discriminating pair is the 30.9 pt cell and the stacked probe.

#### Not done for the five types that draw

Each of these is known-missing rather than merely absent:

* **Data-label wrapping.** PowerPoint wraps a long category name onto two lines inside a
  multi-part label; we draw it on one. `c:separator`, `c:leaderLines` and a data label's
  own `c:layout` are read or ignored but never drawn.
* ~~**A line chart's legend key.**~~ — done. Four line-chart legend probes confirm the
  radar's 19.200 pt of rule with the marker at its midpoint and 2.025 pt before the text,
  and add that a series with `c:symbol val="none"` still gets the rule, and that the
  numbers are **points, not ems**: a 14 pt legend measured the same 19.200 and 2.025 as a
  10 pt one. `real-financial-report.pptx`'s line chart is the deck that wanted it, and it
  is skipped by the scorer on this machine for want of Noto Sans JP, so the improvement is
  not in a scored number.
* **Where a wrapped legend's rows sit.** The band cap and the opened row pitch are
  measured; how PowerPoint places the block vertically is not. Ours centres the rows and
  comes out about 5 pt high on chart4, whose measured baselines are 26.46, 43.50, 56.70,
  86.94 and 116.94 pt from the frame top.
* **`bestFit` is a fixed fraction of the radius.** PowerPoint's moves a label out of the
  way when it does not fit; the probe pie's labels all fit, so that behaviour was never
  exercised.
* **`c:smooth`'s tension.** Drawn as a Catmull-Rom spline, which has the right shape —
  the probe's control points are not collinear with its vertices, so it is a real spline —
  but PowerPoint's own tension was not measured and the curves will not coincide.
* ~~**Rotated category labels.**~~ — done and measured; see *Rotated category labels,
  measured* and *Wrapped category labels* below. Labels wrap at spaces before they turn,
  and the band grows a line box for every extra line. Four pieces are **not** done and
  are named there: PowerPoint caps the band it reserves past about 90 pt of a turned
  label in a way no clamp reproduces; it widens the *left* inset by up to 33 pt, which
  three measurements do not fit; it stops wrapping past six lines and lets the label
  overflow; and an explicit orientation on `a:bodyPr` makes it drop every other label
  rather than turn them, which nothing here implements.
* **Axis titles**, **minor gridlines and minor ticks**, **`c:dTable`**, and manual
  `c:layout` for the plot area or the legend.
* **Secondary axes.** A `c:barChart` group is tied to its axes through its own `c:axId`
  list, which is the hard part and is done; a second value axis is then mostly drawing.
* **Log scales** and `c:tickLblSkip` / `c:tickMarkSkip`. `c:crosses` and `c:crossesAt`
  move the category axis but have only been measured at zero. `c:crossBetween="midCat"`
  is implemented from the schema; nothing measured here uses it.
* **`dispBlanksAs="span"`** is treated as `gap`, which is right for a bar chart and will
  not be for a line one.
* The chart frame's rounded corners (`c:roundedCorners`) and effects.

#### What the probe sweeps could not catch

A review of the finished branch found ten defects, and the shape of them is worth keeping:
**both sweeps assert the plot rectangle, and four of the ten got the rectangle right while
drawing the wrong thing inside it.** A horizontal chart's gridlines ran across the bars
instead of up the plot; `barDir` swapped which line each axis' `c:delete` and `c:spPr`
applied to; `c:catAx/c:txPr` and `c:overlay` were parsed and then never read. Two more
silently lost data — a series longer than the labelled one lost its tail, and a
`c:multiLvlStrCache` without `c:ptCount` lost every label — and three were crashes on
numbers a file controls (`c:gapWidth="-100"` divides by zero; a datum near the float
ceiling overflows the axis rounding; a denormal span underflows the unit).

The lesson is not that measuring was wrong — it is that a measurement pins the *frame* and
says nothing about the *contents*, and that the four orientation bugs all lived in the one
variant with no corpus deck behind it. Anything drawn, not just the box it is drawn in,
needs its own assertion; a parsed field with no reader needs one too.

### 3.3 Combo charts (M)

Multiple `c:*Chart` groups sharing a category axis with a secondary value axis. The reader
already returns every group and each one's `c:axId` list, and the renderer picks the first
group it can draw — so a combo chart whose *second* group is a bar still draws the bar.
Drawing several groups at once, and the secondary axis, is not done.

### 3.4 3-D chart fallbacks (S)

`bar3DChart`, `line3DChart`, `pie3DChart`, `area3DChart` parse as their 2-D equivalents —
`parse/chart.flat_chart_kind` does this and `bar3DChart` therefore already draws flat.
**[pptx-renderer]** does the same and is explicit that it is not PowerPoint-perfect; no
3-D chart has been compared against real output here either.

---

## Phase 4 — EMF / WMF — **done (steps 1–3)**

**Effort: revised down from L to S–M for the cases that matter. Landed at the S end.**

> **Correction to an earlier assessment.** I previously scoped this as "write a subset EMF
> record interpreter" and flagged it as the riskiest phase. Reviewing **[pptx-renderer]**
> changed that: most Office EMFs do not need vector interpretation at all.

### The insight

PowerPoint embeds vector artwork as EMF, but those EMFs usually **carry a ready-made
preview inside them**:

- an **embedded PDF**, inside `EMR_COMMENT` (type 70) records with the `GDIC` comment id
  (`0x43494447`), public types `BEGINGROUP` (2) or `MULTIFORMATS` (`0x40000004`) — findable
  by scanning for `%PDF` … `%%EOF`;
- or an **embedded DIB bitmap**, in an `EMR_STRETCHDIBITS` (type 81) record.

`pptx-renderer/src/utils/emfParser.ts` extracts both in **294 lines** with no record
interpretation whatsoever. That is a direct, well-scoped port.

And we already have the rasteriser: **`pypdfium2`** is needed for Phase 0's oracle anyway,
so embedded-PDF EMFs cost nothing extra.

### Plan

1. ✅ **`pptx2svg/metafile/emf_preview.py`** — record walk, embedded PDF and DIB
   extraction. `metafile/dib.py` decodes 1/4/8/16/24/32 bpp DIBs and writes PNG with
   nothing but `zlib` and `struct`; `metafile/pdf.py` rasterises the PDF case.
2. ✅ **Wired into `resolve/view.py`** — and into the image *fill* and OLE-preview paths
   as well, which take metafiles just as legally as `p:pic` does.
3. ✅ **`metafile_converter` on `ConvertOptions`** — `(bytes, mime) -> (bytes, mime) | None`.
   It takes precedence over the built-in extraction, may return SVG, and a hook that
   raises warns instead of aborting the deck.
4. **Full vector interpreter** (L) — still not built, and the measurement that would
   justify it has not been possible: see below.

WMF has no equivalent preview convention. `extract_metafile_preview` accepts WMF bytes and
returns `None` for them, which costs nothing and correctly handles the real case of a
`.wmf` part that actually contains EMF bytes.

**Notes from the implementation:**

- **The corpus measurement could not be made.** No fixture contains an EMF or WMF at all,
  so there is no local evidence about how many real EMFs carry a preview. The "measure
  before building a vector interpreter" instruction stands, unexecuted; it needs a corpus
  of decks with pasted vector art, which this repo does not have.
- **Hard limits, with one substitution.** Input 8 MB, record 4 MB, 50,000 records are kept
  verbatim. The 200,000-geometry-point limit has no meaning here because no geometry is
  interpreted; the bound that does the equivalent job is a 64-megapixel cap on a decoded
  DIB, since a 40-byte header can otherwise demand a gigapixel allocation.
- **Only `EMR_STRETCHDIBITS` (81) is read for bitmaps.** `EMR_SETDIBITSTODEVICE` (79) and
  the `BitBlt`/`StretchBlt` records carry DIBs too, with different field offsets. Adding
  them is cheap but untestable against real output right now, so they were left out.
- **`EmrFormat` descriptors inside `MULTIFORMATS` are deliberately not trusted** —
  implementations disagree on whether `offData` is relative to the record or to the comment
  data. Accumulating the comment payloads and scanning for `%PDF` … `%%EOF` is immune to
  that, and also handles a PDF split across several `BEGINGROUP` records.
- **`pypdfium2` is an optional extra** (`pip install pptx2svg[metafile]`). Without it the
  PDF case warns `metafile-rasterizer-missing` and keeps the placeholder; DIB extraction
  and PNG encoding stay standard-library only.

**Done when:** a deck with pasted vector art renders it ✅ (against synthetic EMFs — see
the caveat below); a truncated or malformed EMF warns and falls back rather than raising
or hanging ✅; adversarial inputs are covered by tests ✅ (`tests/test_metafile.py`).

**Not yet verified against real input.** Every EMF in the test suite is built by
`tests/test_metafile.py` from the MS-EMF record layouts. That exercises the code paths and
the failure modes, but it cannot catch a misreading of the *specification* — if Office
writes a field somewhere other than where MS-EMF says, these tests will not notice. The
first real EMF-bearing deck should be run through
`convert_pptx_to_model` and checked for a `metafile-image` warning.

---

## Phase 5 — Coverage gaps found in review

**Effort: M–L total.** Individually small, collectively the difference between "renders
most decks" and "renders decks".

### 5.1 The missing preset shapes — **done, and then some**

**Effort: M. Came in under it, because the shapes stopped being transcribed by hand.**

All 53 are implemented, plus `lineInv`, which the original probe missed. The spec's 186
presets are now fully covered. (ECMA-376 itself omits `upArrow` — a known erratum — and we
keep our own; `bendUpArrow` stays registered as an alias for the `bentUpArrow` misspelling.)

**The approach changed partway through, and that is the part worth keeping.** The first
16 callouts and 12 action buttons were transcribed by hand from two third-party copies of
the spec that happened to agree. That works but cannot be checked. ECMA-376 publishes the
geometry as a data file — an electronic addendum to Part 1 — and
`tools/derive_preset_geometry.py` now compiles it directly into
`src/pptx2svg/render/preset_specs.py`, following `tools/extract_font_metrics.py`'s
pattern: pinned source SHA-256, generated output, `--check` that fails on drift.
Re-compiling the hand-transcribed 53 from the official file reproduced them exactly, so
the transcriptions were right — but they are now generated rather than merely lucky.
`render/geometry.py` went from 5426 lines to 1246.

The source file is **not vendored**: ECMA's text copyright policy governs it, so `--source`
points at a copy you obtain yourself. Vendoring it (as pptx-renderer does, with a notice)
is a reasonable alternative and would let `--check` run in CI; that is a project decision,
not a tooling one.

**Two real bugs fell out of doing it properly**, both of which predate this phase and both
of which affect `a:custGeom` as much as presets:

- `stAng`/`swAng` in an `arcTo` are **geometric** angles, not the ellipse's parametric
  angle. They coincide only when `wR == hR`, so every circular arc looked fine and every
  stretched one was wrong. `curvedUpArrow` made it visible: the band missed its own
  arrowhead. Confirmed by the spec's own named guides — after the fix the arcs land on
  `iy` and `x5` exactly.
- SVG draws **nothing** when an arc's endpoints coincide, and DrawingML writes a circle as
  one `arcTo` sweeping 360°. Every circle in the catalogue was being erased —
  `smileyFace` rendered as a bare mouth curve. Sweeps are now cut into half-turn pieces.

Neither changed any fixture's output: the corpus has no elliptical or full-circle custom
geometry. Both would bite on a real deck.

#### Every preset has now been measured against PowerPoint

The judgement-based ranking that used to live here is gone, replaced by a measurement.
All 187 presets PowerPoint accepts were laid out at their default adjustments in probe
decks, exported to PDF by PowerPoint 16.106, and scored by silhouette overlap against
both candidates — our generator and the compiled specification.

**The headline: PowerPoint never contradicts ECMA-376.** Across 187 shapes there was not
one case where our approximation beat the specification, and not one where the
specification failed to match PowerPoint. The standard is simply what PowerPoint draws,
which retires the question the stars raised.

| Outcome | Count |
| --- | --- |
| Specification matched, ours did not → promoted | **51** |
| Both matched (mostly shapes already compiled from the spec) | 135 |
| `upArrow`, which ECMA-376 omits as a known erratum — ours matches at 0.996 | 1 |
| Ours matched and the specification did not | **0** |
| Neither matched | **0** |

After promotion: **median 0.999, mean 0.997, minimum 0.962** across all 187. The seven
still below 0.98 are antialiasing-limited rather than wrong — `line` and `lineInv` both
score 0.962, and they are the same diagonal drawn by different code, one ours and one the
specification's.

142 presets are now compiled from the specification and 45 remain hand-written; the
specification does not beat any of the 45 by more than 0.01, and they are the shapes where
a native `<rect>`/`<ellipse>`/`<line>` is the better output anyway.

**Two aspect ratios, not one.** Every shape was measured in a square box *and* a 2:1 one,
because a generator can be exactly right when the box is square and wrong the moment it is
stretched. `chevron` is the proof: pixel-identical to the specification in a square box,
0.716 against PowerPoint when stretched. A single-aspect sweep would have passed it.

**What the sweep cost, and what it was worth.** Three things went wrong that are worth
knowing before repeating it:

- **Two different input defects present identically**, and the presentation is the problem:
  PowerPoint opens the deck, closes it again, and leaves the export blocked with the
  process idle at 0% CPU. That reads as a hung oracle, so it sends you to look at
  PowerPoint when the fault is in the file you handed it. When an export hangs, suspect
  the deck first and bisect it; the app is almost certainly fine. The two known causes:

  1. **A preset name OOXML does not define.** `bendUpArrow` — our alias for a misspelling
     that appears in real files — is one. Probe decks must contain only standard names.
  2. **A partial `a:avLst`.** Naming only the handle you want to move and leaving
     PowerPoint to fill in the rest is refused; the list must carry *every* handle the
     preset declares, with one value changed. Single-handle shapes such as `roundRect`
     never hit this, because for them a partial list is already complete — which is
     exactly why it stays hidden until a multi-handle shape meets it.
- Twenty shapes on one slide exports in three seconds; 120 across six slides silently
  produces nothing. Batch small.
- The sweep found exactly two shapes where *neither* candidate matched, `cloud` and
  `cloudCallout`, and both turned out to be our bug rather than PowerPoint's divergence:
  a path authored in its own coordinate space had its arc radii scaled before the start
  angle was converted, which rotates any arc not beginning on an axis. Invisible in a
  square box. That bug would not have been found any other way.

#### The adjustment values have been measured too

The sweep above moved no handles: every shape was drawn at its **default** adjustments, so
a preset that mishandles a non-default value passed it untouched. That gap is now closed.
298 handles across 122 presets were swept **one at a time**, others left at their
defaults — the two ends of the range the specification's own `pin` guides define, plus one
value outside each end, with a default-adjustment instance per preset as a control. 1102
cases, each at both aspect ratios.

| Outcome | Cases |
| --- | --- |
| Both matched PowerPoint | 890 |
| Specification matched, ours did not → **promoted** | 15 (10 presets) |
| Undecidable by this method, then re-probed → **promoted** | 8 (3 presets) |
| Neither matched | 38 |
| Ours matched, the specification did not | 8 |
| Undecidable even after re-probing | 2 |

**There is no clamp divergence, and that was the thing worth checking.** PowerPoint honours
the specification's `pin` exactly. Sweeping `parallelogram` through its published range
gives rendered areas of 1.000, 0.750, 0.600, 0.500, 0.400, 0.250 at adj = 0, 25000, 40000,
50000, 60000, 75000 — tracking the published geometry at every step.

**What PowerPoint does differently is narrower and stranger.** At the exact values where
the published geometry *degenerates to zero area*, PowerPoint draws the shape's bounding
rectangle instead of the collapsed path. `parallelogram` at adj = 100000 and `diagStripe`
at adj = 100000 both reduce to a line under the spec's own formulas, and both render as a
filled box. That accounts for most of the 38 "neither matched" cases and for all 8 where
our older approximation scored better — `leftBrace` and `rightBrace` collapse at both ends
of `adj2`, and our hand-written brace happened to land nearer PowerPoint's substitute.

This is **not implemented**. It only occurs at the extreme end of an adjustment range,
which PowerPoint's own UI will not let a user reach by dragging, and a
"degenerate → bounding box" fallback would mask real geometry bugs as readily as it would
match Office. It is recorded here as a known, bounded divergence.

**The promotions came almost entirely from out-of-range values**, which is the argument for
testing them: a generator that ignores the clamp is indistinguishable from one that
honours it until you hand it something out of range. `downArrow` at adj1 = -25000 scored
0.727 where the specification scores 1.000; `triangle`, `snip1Rect`, `snip2SameRect`,
`snipRoundRect`, `round2DiagRect`, `foldedCorner`, `bevel`, `frame` and `upDownArrow` the
same way.

**A blind spot found and then closed.** A callout's handles move its *leader line*, which
lies outside the shape's box — so cropping the box clips it away for PowerPoint and for us
alike, and 164 cases compared two identical pictures. Reporting those as agreement would
have been false; they were reported undecidable, then re-probed with the shape inset inside
a larger measured region. 162 became real measurements and 8 of them were failures:
`borderCallout3` on six handles, `wedgeRectCallout` and `wedgeRoundRectCallout` on `adj2`.
**When a preset's geometry can leave its own box, the box is the wrong crop.**

155 of the 188 registered presets are now compiled from the specification.

**The lesson that generalises.** The ranking this replaced was produced by eye and was
wrong in both directions: it listed `chevron` at 0.31 divergence (a square-box artefact of
measuring at 2:1) while calling `decagon` and `dodecagon` fine at a glance, when both were
systematically wrong. Shapes at 0.86–0.94 are exactly the ones a screenshot calls correct.

Shapes that should *stay* hand-written: `rect`, `ellipse`, `line`, `roundRect` and
friends emit a native `<rect>`/`<ellipse>`/`<line>`, which is smaller and strokes correctly
under a transform. All of them match PowerPoint already.

### 5.2 Other gaps (S–M each)

- **Picture bullets** (`a:buBlip`) — bullets as images. We handle `buChar` and `buAutoNum`
  only, and silently drop `buBlip`.
- **`a:clrChange` on a picture** — parsed and resolved onto `BlipEffects.clr_change`, and
  never drawn. Replacing one colour with another needs more than `feComponentTransfer`
  can express per channel; the usual trick is an `feColorMatrix` that isolates the source
  colour followed by an `feComposite`. Found by the model-coverage test, not by reading.
- **Compound lines** (`a:ln@cmpd`: `dbl`, `thickThin`, `thinThick`, `tri`) — currently render
  as a single stroke.
- **Rectangular gradients** — we treat every `a:path` gradient as radial;
  `a:path path="rect"` is a distinct shape gradient.
- **Pattern fills** — we implement ~25 presets, **[pptx-renderer]** has 52+. Same additive,
  low-risk shape as 5.1.
- **Prefer the SVG picture extension** — `mc:AlternateContent` may offer `asvg:svgBlip`
  (vector) in its `Choice` with a raster `Fallback`. `parse/shapes.py:parse_alternate_content`
  currently always prefers `Fallback`; since we can embed SVG directly, preferring a
  *supported* `Choice` is strictly better output. Small change, real quality win.
- **Audio / video placeholders** — media shapes currently vanish. Rendering the poster frame
  (`p:pic` inside the media frame) is nearly free.
- **Line-spacing edge semantics** — **[pptx-renderer]** notes percentage line spacing follows
  Office *line-unit* semantics, and that ordinary text boxes trim spacing outside the first
  and last visible paragraphs. We do neither; it shifts every multi-paragraph body slightly.

### 5.3 Advanced rendering (L, do last)

- **Embedded fonts** — **moved to Phase 6, and re-estimated from M to L.** This entry used
  to read "extract, undo the trivial obfuscation, pass via the existing `font_files`
  parameter". Both halves were wrong: the payload is EOT and in practice MicroType-Express
  compressed, not trivially obfuscated, and `font_files` feeds the rasteriser only, so it
  would have drawn the right glyphs while still measuring the wrong widths. Kept visible
  rather than deleted, because the happy path is what made it look cheap.
- **3-D (`a:sp3d`, `a:scene3d`), bevels, reflection** (L) — SVG has no 3-D model. Target an
  approximation: bevels as light/dark edge gradients, reflection as a flipped copy under an
  opacity gradient. **[pptx-renderer]** does not attempt these either.
- **Text-to-path** (M) — glyphs as `<path>` via fontTools, making output font-independent.
  `FontToolsTextMeasurer` already loads the faces.
- **Gradient stop overrides** (S) — known approximation: `resolve/view.py:_blend_stop`
  replaces every theme gradient stop with the `a:fillRef` override instead of re-deriving
  each stop's own tint/shade against it, so such gradients render flat. Fix by keeping stop
  colours unresolved in the format scheme until the override is known.
- **Bidi / RTL text** (L) — Arabic and Hebrew need reordering and shaping. Depends on 5.2's
  complex-script fonts. Large, and only matters for those scripts.

---

## Phase 6 — Embedded fonts — **not started; read the cost before the plan**

**Effort: L, with a hard dependency that could make it XL.** Phase 5.3 carried this at M
on the assumption of "trivial obfuscation". That assumption is false, and the correction
is the most important thing on this page.

### What is there, and that nothing reads it

`ppt/presentation.xml` carries `<p:embeddedFontLst>` — one `<p:embeddedFont>` per family,
each with `<p:regular>` / `<p:bold>` / `<p:italic>` / `<p:boldItalic>` children naming
`ppt/fonts/*.fntdata` parts by relationship id. All four faces, keyed by family name,
exactly the shape the font subsystem wants.

`grep -rln "embeddedFont\|fntdata" src/` returns nothing. The code has never heard of it.

### Why this beats a larger bundle

Two third-party template decks measured locally embed **Anton, Arimo, Literata,
Merriweather Sans, Merriweather Sans Light and Inclusive Sans**. Every one of those faces
raises `font-substituted` — "no substitute known; widths guessed, drawn with the generic
family" — while sitting inside the file being rendered. The deck handed us the font and we
walked past it.

That is the general case, not a curiosity. Template vendors pick arbitrary Google Fonts,
of which there are roughly 1,800 families; a bundle can never catch up, and each new deck
brings faces nobody anticipated. Reading the file always wins, and wins for faces that did
not exist when the release shipped.

Note also which face is in that list: **Arimo, which this project already bundles.** The
embedded-font path would have rendered it correctly all along.

### The payload is EOT, not a font file

Measured on both decks, and the identification is measurement rather than inference —
`EOTSize` matching the part length to the byte on two independent files settles it:

```
part bytes   70545
EOTSize      70545        <- matches the part length exactly
FontDataSize 70305
Version      0x00020002   <- EOT 2.2
Flags        0x00000004
bytes at (EOTSize - FontDataSize): 03 02 78 c0   <- not 00 01 00 00, not OTTO
```

### The flag reading is an inference, and must stay labelled as one

Apache POI's developer list states that PPTX embedded fonts are "always in EOT format…
subsetted and compressed… **MicroType Express (MTX) and Non-MTX**", so both variants occur
in the wild. Reading `Flags = 0x4` as `TTEMBED_TTCOMPRESSED` is consistent with that and
with the absent TrueType signature at the computed data offset — but **nothing has
actually decompressed one of these files**. Until something does, "these two decks are
MTX-compressed" is the leading hypothesis, not an established fact.

**Establishing it is the first task of this phase, and it is cheap.** Run one `.fntdata`
payload through any existing EOT decoder — libEOT, or FontForge, which reads EOT — and see
whether a usable font falls out. That single result decides the size of everything below,
so do it before estimating anything.

### If the inference holds, MTX is the whole problem

The uncompressed path is a header skip plus an optional XOR and would be an afternoon. It
is also **not what the decks that motivated this phase contain**, so shipping only that
path would fix nothing currently observable. Any plan that quietly scopes to non-MTX is
back to estimating from the happy path.

LibreOffice shipped PPTX embedded-font support in **25.8**, and the shape of its solution
is the cost estimate: `EmbeddedFontsHelper::addEmbeddedFont()` in
`vcl/source/gdi/embeddedfontshelper.cxx` writes a temporary font file, calls
`AddTempDevFont`, and **delegates the decode to libEOT**. That is a C library. A
pure-Python MTX decoder is a project in its own right, not a patch.

**This therefore lands as an optional extra, shaped exactly like `[metafile]`** — the core
stays standard-library-only, the decoder lives behind an extra, and a deck whose embedded
fonts cannot be decoded degrades to today's behaviour with a warning naming the reason.
That is the same degradation contract `metafile-rasterizer-missing` already implements, so
there is a pattern to copy rather than a policy to invent.

### The `fsType` gate is a constraint, not a nicety

An embedded font carries its own embedding-permission bits in the OS/2 table's `fsType`
field, and the foundry's restrictions travel with the file. **Both** LibreOffice
(`EmbeddedFontsHelper::sufficientTTFRights`) **and** **[pptx-renderer]** check them before
use, and LibreOffice 26.2 went further, replacing silent discard with a dialog telling the
user why a font was refused.

Given that this project will not put a Microsoft font in git, extracts nothing it has no
right to, and treats measured widths as facts precisely *because* font files are not —
using an embedded font without checking `fsType` would be out of character to the point of
inconsistency. **Write the gate into the design before any decode code exists.** It is far
harder to add a rights check to a working extractor than to build one that never had a
bypass.

### An extracted font must reach measurement, not just the rasteriser

The obvious wiring is to hand the extracted files to `svg_to_png` through the existing
`font_files` parameter. **That is half an implementation and the wrong half.** Those
arguments feed the rasteriser only; by the time they are used, `convert_pptx_to_svg` has
already measured every line from the static tables. The result would be a deck drawn in
the face the author chose and broken in exactly the way this library was built to avoid:
right glyphs, wrong line breaks, wrong autofit, wrong centring — and looking plausible
enough in a screenshot to pass review.

So an embedded face has to land in **both** halves: a measurer that can read the extracted
file (`FontToolsTextMeasurer` already does this for files on disk, behind
`pptx2svg[measure]`), and the rasteriser's font set. The measure-equals-draw invariant is
the whole point of the subsystem, and an embedded font is no more exempt from it than
Aptos is — where the invariant *is* deliberately broken, `Substitution.metric_compatible`
records it and `fonts --check` fails on it.

This also settles a question that would otherwise come up in review: extracting to a
temporary file on disk is not a workaround, it is what LibreOffice does
(`AddTempDevFont`), and it is what lets one extracted face serve both measurement and
drawing without inventing a second in-memory path.

The same gap exists today for `--font-dir` and is recorded under Fonts ▸ Left undone;
whoever fixes it there should look at this phase first, since the two want the same
plumbing.

### The bundle is not the answer here, and for the literal version cannot be

Recorded because it will be proposed again. "Put PowerPoint's default fonts in
`pptx2svg[fonts]`" splits into two requests with opposite answers:

- **Microsoft's actual font files — blocked, permanently.** Aptos, Calibri, Cambria,
  Candara, Consolas, Constantia, Corbel and the rest are licensed to Office. No
  Microsoft-licensed font file may enter this repository or any built artifact, and that
  rule is not negotiable for convenience. It is also why `tools/install-fonts-debian.sh`
  puts the ClearType set behind `--ppviewer` with an explicit licence warning instead of
  shipping it.
- **Open clones of them — legitimate, and partly done.** The bundle already carries six:
  Carlito, Arimo, Tinos, Cousine, Caladea and Noto Sans JP, plus Lato and Raleway, which
  are Office *cloud* fonts. That set covers 8 of the ~186 distinct typefaces a current
  PowerPoint installs. The remaining ~178 include entirely ordinary menu choices — Gill
  Sans MT, Rockwell, Franklin Gothic, Century Gothic, Garamond, Palatino Linotype, Tw Cen
  MT, Perpetua, Bookman Old Style, the Lucida family — and a deck naming one of those
  without embedding it gets the width guess today. Finding which have open equivalents is
  a real and separate piece of work.

The two paths are **disjoint, not competing**. Embedded fonts fix decks that carry their
faces; clones fix decks that name a PowerPoint face and carry nothing. Neither substitutes
for the other, and the bundle can never cover arbitrary vendor faces at all.

One trap for whoever scopes the clone work: a clone that merely *looks* similar is not a
clone that *measures* the same. This library measures with what it draws with, so a
non-metric-compatible substitute must report `approximate`, never `compatible`. Guessing
that distinction is how the Caladea 4.5 % error got in.

### Done when

- One real `.fntdata` payload has been decoded to a usable font file, and the MTX question
  is answered by measurement rather than by citation.
- `<p:embeddedFontLst>` is parsed into the source model, keyed by family and face.
- An embedded face is **both measured and drawn** from the extracted file — not drawn from
  it while measured from a guess, which is the failure `font_files` alone would have
  produced.
- `fsType` is checked, and a restricted font is refused with a warning that names the
  restriction.
- A deck whose embedded font cannot be decoded still renders, and says why.
- The two template decks in `scratch/` render without a single "no substitute known"
  warning, which is the observable this phase exists to move.

---

## Suggested order

```
Phase 0  (PowerPoint oracle + VRT)  ──┬─▶ Phase 1  (parsed-but-unrendered)   DONE
                                      ├─▶ Phase 2  (SmartArt — DONE)
                                      ├─▶ Phase 4  (EMF previews — DONE)
                                      ├─▶ Phase 5.1 DONE / 5.2  (small gaps)
                                      ├─▶ Phase 3  (charts — reader + 5 types done)
                                      └─▶ Phase 6  (embedded fonts — not started)
                                                              Phase 5.3 last
```

Phase 0 gates everything and is now much stronger than originally planned, because the
reference implementation is available locally.

Revised quick wins, in order of payoff per day:

1. ~~**Phase 1 table styles**~~ — done, though not in an afternoon: the built-in
   definitions are not in the file and had to be measured out of PowerPoint.
2. **Font metrics** — now the largest single component of the residual on every fixture,
   and the thing standing between `real-basic-theme.pptx`'s gridlines and PowerPoint's.
   Phase 6's embedded fonts attack one half of it, and are the only half a bundle can
   never reach; open clones for PowerPoint's own faces are the other, and the two do not
   overlap.
3. ~~**Phase 4 EMF previews**~~ — done; the S estimate held.
4. ~~**Phase 2 SmartArt**~~ — done; the S estimate held.
5. ~~**Phase 5.1 shapes**~~ — done. Not the additive, near-zero-risk job it looked
   like: it turned up two arc-conversion bugs that had been silently misdrawing
   custom geometry, and it replaced hand-transcription with a spec compiler.

Phase 3 is well along: the reader and five chart types -- `barChart`, `lineChart`,
`pieChart`, `doughnutChart` and `radarChart` -- are done and measured, and no chart in
the corpus warns `chart-unsupported-type` any more. The shared infrastructure the rest
need -- value domain, tick selection, number formatting, gridlines, legend layout,
plot-area rectangle, polar region -- is built. What is left, cheapest first:

1. **The CJK label width.** Now the largest single error left in any chart: on
   `real-financial-report.pptx`'s chart3 it is 19.8 pt of the 19.8 pt that remains after
   rotation, and on its radar it is the whole 6.8 pt of radius error. `font_box` and
   `text_width` measure a Japanese label through the `<a:latin>` face its axis names
   rather than the CJK face it will actually be drawn in. It lives in `text/` and
   `fonts/`, not in `resolve/chart.py`.
2. **A manually laid out legend.** `real-college-template.pptx` slide 4 carries a
   `c:legend/c:layout/c:manualLayout` we ignore, and it is now the whole of that slide's
   remaining chart error: 4.8 pt of plot height and 9.5 pt of legend baseline. Reading
   the element is trivial; what is not known is how PowerPoint composes the plot area
   around it, and one deck is not enough to fit that.
3. **Data-label wrapping**, **`bestFit` actually moving a label**, and **three-or-more
   line legend entries** -- each a known-missing detail with a named symptom above.
   Category-label wrapping is done; the data-label kind is a separate path.
4. **Combo charts** and **`areaChart` / `scatterChart`** -- new drawing rather than
   corrections, and the largest of what remains.

## Non-goals

Explicitly out of scope, to save re-litigating them:

- Animations, transitions, timing
- Speaker notes, comments, revision history
- Editing or writing `.pptx`. This library reads and renders; editing lives in the separate
  [`pptx-agent`](https://github.com/uvrt/pptx-agent) project, which depends on this one for
  rendering. The only concession made here is the shape identity described above — enough for a
  downstream editor to map a rendered group back to a shape, and nothing more.
- Equations (OMML) — **[pptx-renderer]** excludes these too
- Executing or editing embedded OLE objects (previews only)
- Pixel-exact PowerPoint reproduction — the goal is accurate text, shapes and layout

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

---

## Where we are

| Area | State |
| --- | --- |
| Package, relationships, parts | Complete |
| Theme colours, colour maps, transforms | Complete |
| Placeholder / background / text inheritance | Complete |
| Shapes: 134 presets + custom geometry with guide formulas | Common set done; **~53 presets missing** |
| Text: cascade, bullets, wrapping (Latin + CJK), autofit, vertical, tabs, columns | Complete for the common path |
| Fills, outlines, arrowheads, shadows, glow, soft edge | Complete |
| Pictures: crop, colour adjustments, tile, stretch | Complete |
| Tables: merged cells, borders, fills, **table styles** | Complete; 72 built-in styles carried |
| Charts | **Not rendered** |
| SmartArt | **Not rendered** |
| EMF / WMF | **Not rendered** |
| 3-D, bevel, reflection | **Not rendered** |
| Shape identity on output (`data-pptx-id`) | Complete |

292 tests pass. The pipeline is `opc → parse → resolve → render → png`; each stage is
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

Against real PowerPoint output, per fixture, at 1280 px wide. "Before" is the tree as of
commit 076a0f5, "after" is with Phase 1 landed.

| Fixture | >10/255 before | after | >64/255 before | after | mean before | after |
| --- | --- | --- | --- | --- | --- | --- |
| `authoring-integration.pptx` | 11.84% | **6.72%** | 5.66% | 5.55% | 9.80 | **7.87** |
| `real-basic-theme.pptx` | 1.84% | 2.07% | 0.98% | 1.12% | 1.72 | 1.86 |
| `real-product-page.pptx` | 6.80% | 6.80% | 4.77% | 4.77% | 8.40 | 8.40 |
| `real-financial-report.pptx` | 5.26% | 5.26% | 3.52% | 3.53% | 5.32 | 5.32 |
| `sample.pptx` | 2.61% | 2.58% | 1.96% | 2.04% | 4.16 | 4.18 |
| `sample-issue-387.pptx` | 1.91% | 1.91% | 1.44% | 1.44% | 2.80 | 2.80 |

Layout, colour and text positions match; the residual is almost entirely glyph
antialiasing and font substitution. **This is the number to drive down** — treat a
regression in it as a failing build.

Two entries need reading rather than scanning. `authoring-integration.pptx` is the only
fixture whose table names a built-in style, and it is the one Phase 1 moved.
`real-basic-theme.pptx` moved the *wrong* way, and the reason is worth keeping: its
custom table style's gridlines are now drawn, correctly and in the right colour, but a
couple of pixels lower than PowerPoint puts them, because our line height for its CJK
text comes out about 7% short. Drawing nothing scored better than drawing something
slightly misplaced. The next win on that fixture is font metrics, not tables.

The remaining three fixtures are flat because none of them uses anything Phase 1 touched
— checked, not assumed: `real-financial-report.pptx`'s tables carry explicit per-cell
fills and borders on every cell, so no style could change them.

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
| PowerPoint is sandboxed | Paths must be under the user's home. `/tmp` and `/private/tmp` fail with error **−9074**. |
| **−9074 has a second cause** | A file PowerPoint wants to repair raises an app-modal dialog, and *every* export then fails −9074 until it is cleared — including known-good files. Check for a dialog before suspecting the path; one bad input otherwise looks exactly like a broken environment. |
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
- Rendering is already deterministic (counter-based ids, static font metrics), **provided
  fonts are pinned**. Download Carlito / Liberation Sans / Liberation Serif / Noto Sans JP
  to a cache, verify SHA-256, and pass via `font_dirs=[...], skip_system_fonts=True` —
  `convert_pptx_to_png` already supports both. Without this, snapshots differ between
  macOS and CI.

Two layers, two jobs: snapshot VRT runs everywhere and catches regressions; the PowerPoint
oracle runs on this Mac and catches *being wrong in the first place*.

### 0.3 Metrics

Raw pixel-difference percentage is a blunt instrument — a 1 px text baseline shift lights
up every glyph. **[pptx-renderer]** uses SSIM, colour histogram distance, and IoU of
detected foreground regions, which separate "shifted slightly" from "structurally wrong".
Worth adopting; SSIM is a few lines over NumPy, no new dependency beyond what Pillow
already brings.

### 0.4 Fixture corpus

Six fixtures is thin, and **none contain SmartArt** — Phase 2 needs inputs before it needs
code. Generate a synthetic corpus with `python-pptx`, one feature per slide (each preset
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
| **Table styles** (`tableStyles.xml`, `a:tblStyle`) | Done — custom styles read, 72 built-ins carried |
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
once the accent number is factored out, which caught five stray colour matches. The
catalogue is **deliberately incomplete** rather than padded: "Light Style 1 - Accent 4"
and "Medium Style 1" are absent because their GUIDs are not known here, and a table
naming an unknown GUID falls back to no style, exactly as PowerPoint does.

### Also found on the way

Two things turned up that were not in this plan:

- **`tint` and `shade` were wrong in two independent ways** — the value was read as how
  far the colour *moves* rather than how much of it *survives*, and the blend was done in
  sRGB where PowerPoint does it in linear light. 284 of the 546 swatches differed from
  PowerPoint by more than 2/255 before the fix, several by more than 150. None do now.
  `lumMod`/`lumOff`, checked the same way, were already exact.
- **`a:clrChange` is parsed, resolved, and has no renderer at all** — a gap this section
  missed. It is now recorded in `UNRENDERED_FIELDS`, and is a Phase 5 item.

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

None of the six fixtures uses a tab stop, a highlight, `bodyPr@rot`, a hidden shape, a
non-single underline or multiple columns — checked, not assumed. Those are validated
against PowerPoint using the deck `tools/make_feature_probe.py` builds, which is also
where the mid-paragraph column finding above comes from. A synthetic corpus of the kind
Phase 0.4 describes would subsume it.

---

## Phase 2 — SmartArt

**Effort: S–M. Much cheaper than it looks.**

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

**Work:**

1. `parse/shapes.py` — `_diagram_drawing_rel_id` already returns the `r:dm` id. Keep it.
2. `resolve/view.py:_resolve_unsupported` — for `what == "diagram"`, hop: slide rels →
   data-model part → its rels → drawing part. Parse its `spTree`, resolve children using
   the *drawing part's* relationships (its images live there), return a `GroupElement`.
3. Wrap in a group whose `child_transform` comes from `dsp:spTree/dsp:grpSpPr/a:xfrm`
   `chOff`/`chExt`, so the diagram scales into the frame. The existing group renderer does
   the rest.
4. Keep the current warning + empty frame when no cached drawing exists.

Accept the relationship type under both `schemas.microsoft.com/office/2007/…` and
`purl.oclc.org/ooxml/…` namespaces.

**[pptx-renderer]** notes that diagram groups need "diagram-specific compensation requiring
matching layout provenance" — expect the child transform to need care. Their fallback when
no drawing exists is an EMF preview, which Phase 4 unlocks for us too.

**Done when:** a SmartArt cycle/hierarchy renders its nodes and connectors; a deck with the
drawing part deleted still renders and still warns.

---

## Phase 3 — Charts

**Effort: XL. The largest single piece of work here.**

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
equivalents are matplotlib (SVG backend) and pygal (pure-Python SVG). I recommend
**hand-rolling** anyway:

- Both bring heavy styling opinions that must then be fought back to Office defaults.
- matplotlib is a large dependency for a library whose selling point is having none; it
  would have to be optional, which means two code paths.
- The output must compose into an existing SVG document with our transform/clip stack, not
  stand alone.

This is a judgement call worth revisiting if chart work stalls — delegating would trade
fidelity and dependency weight for a large amount of saved time.

### 3.1 Chart data model and reader (M)

- New `parse/chart.py`: `c:chartSpace` → `SourceChart` (type, series, categories, axes,
  legend, title).
- Values live in `c:numCache` with `c:f` holding the spreadsheet formula. **Always read the
  cache**; parsing the embedded XLSX workbook is a separate project.
- New `model.py` types mirroring pptx-glimpse's `ChartData` / `ChartSeries` / `ChartLegend`.
- Resolve series colours through the theme (`a:solidFill` under `c:spPr`, else the
  `varyColors` accent cycle).
- **[pptx-renderer]**: chart-local colour maps must stay isolated from the parent slide's.
  Easy to get wrong, hard to notice.

### 3.2 Renderer, in order of real-world frequency (L)

1. `barChart` (clustered, stacked, percentStacked; `barDir` col/bar) — most common by far
2. `lineChart`
3. `pieChart` / `doughnutChart`
4. `areaChart`
5. `scatterChart` / `bubbleChart`
6. `radarChart`, `stockChart`, `surfaceChart`, `ofPieChart` — long tail; defer

Shared infrastructure, built once: zero-inclusive value domain, "nice" tick selection, tick
formatting honouring `c:numFmt@formatCode`, gridlines, legend layout for the four
`legendPos` values, plot-area rectangle.

**[pptx-renderer]** also handles: sparse scatter/bubble caches preserving missing
coordinates *and* explicit zeros (different things), gap/span/zero handling, and explicit
negative-bar inversion (`c:invertIfNegative`, present in our fixture). Worth reading their
implementation before writing ours.

### 3.3 Combo charts (M)

Multiple `c:*Chart` groups sharing a category axis with a secondary value axis. Do this
after single-type charts are solid — pptx-glimpse treats it as a distinct code path
(`renderCategoryComboChart`) for good reason.

### 3.4 3-D chart fallbacks (S)

`bar3DChart`, `line3DChart`, `pie3DChart`, `area3DChart` parse as their 2-D equivalents.
**[pptx-renderer]** does exactly this and is explicit that it is not PowerPoint-perfect —
but a flat bar chart beats an empty frame.

**Sequencing note:** 3.1 is independently valuable — it gives `convert_pptx_to_model`
callers the chart *data* before anything is drawn.

---

## Phase 4 — EMF / WMF

**Effort: revised down from L to S–M for the cases that matter.**

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

1. **`pptx2svg/metafile/emf_preview.py`** (S) — port the extractor: validate the EMF
   signature at offset 40 (`0x464d4520`), walk records to `EMR_EOF` (14), pull out embedded
   PDF or DIB. Keep pptx-glimpse's hard limits verbatim — metafiles are attacker-controlled
   binary: 8 MB input, 4 MB per record, 50,000 records, 200,000 geometry points.
2. **Wire into `resolve/view.py:_resolve_image`** (S) — replace the current grey placeholder:
   embedded PDF → rasterise with `pypdfium2` → `ImageElement`; DIB → encode PNG → `ImageElement`;
   neither → keep the placeholder and the warning.
3. **Pluggable converter hook** (S) — a `metafile_converter` callable on `ConvertOptions` so
   users can plug in Inkscape / `libemf2svg` for the remainder. Opt-in, so the pure-Python
   default holds.
4. **Full vector interpreter** (L) — *only if* the corpus shows EMFs without previews.
   Measure before building. There is no maintained pure-Python EMF→SVG library (`pyemf`
   writes rather than reads and has been dead since 2006), so this would be from scratch —
   but steps 1–3 may make it unnecessary.

WMF has no equivalent preview convention; it is rarer in modern decks, so defer it behind
the same measurement.

**Done when:** a deck with pasted vector art renders it; a truncated or malformed EMF warns
and falls back rather than raising or hanging; adversarial inputs are covered by tests.

---

## Phase 5 — Coverage gaps found in review

**Effort: M–L total.** Individually small, collectively the difference between "renders
most decks" and "renders decks".

### 5.1 The missing preset shapes (M)

We implement 134 presets; **[pptx-renderer]** implements 187+. Probing well-known OOXML
names found **53 we lack**, 48 of which they implement:

| Family | Missing |
| --- | --- |
| Action buttons | `actionButtonHome`, `actionButtonBackPrevious`, `actionButtonBeginning`, `actionButtonBlank`, `actionButtonDocument`, `actionButtonEnd`, `actionButtonForwardNext`, `actionButtonHelp`, `actionButtonInformation`, `actionButtonMovie`, `actionButtonReturn`, `actionButtonSound` |
| Callouts | `callout1‑3`, `accentCallout1‑3`, `accentBorderCallout1‑3`, `leftArrowCallout`, `rightArrowCallout`, `upArrowCallout`, `downArrowCallout`, `leftRightArrowCallout`, `upDownArrowCallout`, `quadArrowCallout` |
| Curved & circular arrows | `curvedUpArrow`, `curvedDownArrow`, `curvedLeftArrow`, `curvedRightArrow`, `circularArrow`, `leftCircularArrow`, `leftRightCircularArrow`, `swooshArrow` |
| Banners & scrolls | `verticalScroll`, `horizontalScroll`, `ellipseRibbon`, `ellipseRibbon2`, `leftRightRibbon` |
| Gears, tabs, misc | `gear6`, `gear9`, `funnel`, `pieWedge`, `cornerTabs`, `squareTabs`, `plaqueTabs`, `nonIsoscelesTrapezoid`, `chartPlus`, `chartStar`, `chartX`, `flowChartOfflineStorage` |

Purely additive to `render/geometry.py` — each is an independent function plus a registry
entry, so this parallelises across people or sessions and carries near-zero regression risk.
Callouts and action buttons are the common ones in real decks.

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

- **Embedded fonts** (`ppt/fonts/*.fntdata`) (M) — extract, undo the trivial obfuscation,
  pass via the existing `font_files` parameter. Best fidelity-per-effort in this group, and
  it directly attacks the font-substitution component of the 6.84% baseline.
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

## Suggested order

```
Phase 0  (PowerPoint oracle + VRT)  ──┬─▶ Phase 1  (parsed-but-unrendered)   DONE
                                      ├─▶ Phase 2  (SmartArt)
                                      ├─▶ Phase 4  (EMF previews — now cheap)
                                      ├─▶ Phase 5.1/5.2  (shapes + small gaps)
                                      └─▶ Phase 3  (charts — longest pole, start early)
                                                              Phase 5.3 last
```

Phase 0 gates everything and is now much stronger than originally planned, because the
reference implementation is available locally.

Revised quick wins, in order of payoff per day:

1. ~~**Phase 1 table styles**~~ — done, though not in an afternoon: the built-in
   definitions are not in the file and had to be measured out of PowerPoint.
2. **Font metrics** — now the largest single component of the residual on every fixture,
   and the thing standing between `real-basic-theme.pptx`'s gridlines and PowerPoint's.
   Phase 5.3's embedded fonts attack one half of it.
3. **Phase 4 EMF previews** — dropped from L to S–M by the embedded-PDF finding, and
   `pypdfium2` arrives with Phase 0 regardless.
4. **Phase 2 SmartArt** — dropped from XL to S–M by the cached-drawing finding.
5. **Phase 5.1 shapes** — additive, parallelisable, near-zero risk.

Phase 3 remains the long pole and should start in parallel rather than waiting.

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

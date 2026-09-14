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
| Shapes: all 186 ECMA-376 presets + custom geometry | Complete and **verified against PowerPoint**: every preset matches its own PDF export at two aspect ratios, median 0.999 silhouette overlap |
| Text: cascade, bullets, wrapping (Latin + CJK), autofit, vertical, tabs, columns | Complete for the common path |
| Fills, outlines, arrowheads, shadows, glow, soft edge | Complete |
| Pictures: crop, colour adjustments, tile, stretch | Complete |
| Tables: merged cells, borders, fills, **table styles** | Complete; 72 built-in styles carried, **1 verified** |
| Charts | `barChart` read and drawn, **verified against PowerPoint** across 18 probe charts and 3 real ones; every other chart type warns and draws an empty frame |
| SmartArt | Cached drawing rendered and verified against 46 real decks; **no layout engine**, so diagrams without a cache draw nothing and say so |
| EMF / WMF | Embedded previews rendered (Phase 4); **no vector interpreter** |
| 3-D, bevel, reflection | **Not rendered** |
| Shape identity on output (`data-pptx-id`) | Complete |
| Fonts: bundled, metric-generated, diagnosed | Complete; **Aptos and Cambria approximate** |

356 tests pass. The pipeline is `opc → parse → resolve → render → png`; each stage is
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

Against real PowerPoint output, every slide of every fixture, at 1280 px wide, produced
by `tools/fidelity.py`.

**These numbers mean something different from the ones they replace.** Every earlier
table on this page was partly a measurement of font availability: PowerPoint drew with
Microsoft's Calibri, Cambria and Aptos, our side drew with whatever resvg could find, and
for five of the seven fixtures that was a generic sans. The difference between the two
images was dominated by glyph shape before the renderer had done anything at all. The
harness now renders our side with the *same licensed faces* PowerPoint used, read in
place through a gitignored local profile, and a deck whose faces PowerPoint did not have
either is skipped rather than scored against Microsoft's own fallback.

So the columns below are not comparable with the "pre" and "P1" columns that used to be
here; they were taken under a different and less honest configuration. What *is*
comparable is before/after within this table, both measured with PowerPoint's faces on
both sides.

| Fixture | SSIM before / after | hist before / after | >10/255 before / after |
| --- | --- | --- | --- |
| `authoring-integration.pptx` | 0.7666 / 0.7661 | 0.797 / 0.796 | 6.80 / 6.76 |
| `real-basic-theme.pptx` | 0.9658 / **0.9674** | 1.000 / 1.000 | 1.58 / **1.52** |
| `sample.pptx` | 0.1178 / 0.1178 | 0.974 / 0.974 | 2.49 / 2.49 |
| `table test.pptx` | 0.9490 / **0.9531** | 0.998 / 0.998 | 1.26 / **1.19** |
| `real-financial-report.pptx` | skipped | — | — |
| `real-product-page.pptx` | skipped | — | — |
| `sample-issue-387.pptx` | skipped | — | — |

"before" is c57ca8f, "after" is the font work, both under the same font profile.

Two honest observations about that table.

**The gains are small, and that is the finding.** Bundling the right fonts and generating
the metrics from them removes a whole class of wrongness, but on a corpus scored with
PowerPoint's own faces on *both* sides there was never much font-related error left to
remove — the errors that remain are geometric. `table test.pptx` crossing the 0.95 gate
is the one visible win, and it comes from Aptos Display finally having a metrics table.

**Three fixtures are skipped, and that is fixable.** They name `Noto Sans JP`, which
PowerPoint does not have on this machine, so its export is already drawn with a
substitute of its own choosing. Installing Noto Sans JP where PowerPoint can see it
(`~/Library/Fonts`) and re-exporting would make all three comparable again. It was not
done here because it changes the developer's machine, not the repository.

### What the corpus says is wrong now

With fonts eliminated as a variable, `sample.pptx` at 0.118 SSIM is the loudest remaining
signal, and a side-by-side of slide 2 says plainly what it is:

* **The text block sits about 110 px too high** in a 720 px render. Vertical anchoring or
  the first-baseline rule, not fonts.
* **Bold and italic runs are not distinguished.** PowerPoint draws `PPTX` bold on one line
  and the whole of another in italic; ours renders both upright and uniform. The likely
  cause is that `msgothic.ttc` is a collection and resvg's face matching does not reach
  its bold member, but it has not been confirmed.
* Line breaks then differ, which is a consequence of the first two rather than a third
  bug.

None of that is font *selection*; all of it is layout. It belongs to whoever picks up
text positioning next, and it is now measurable, which it was not before.

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

None of the six other fixtures uses a tab stop, a highlight, `bodyPr@rot`, a hidden shape, a
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

### 3.2 Renderer — `barChart` **done**, the rest not started

1. ✅ `barChart` — clustered, stacked, percentStacked, `barDir` col and bar
2. `lineChart`
3. `pieChart` / `doughnutChart`
4. `areaChart`
5. `scatterChart` / `bubbleChart`
6. `radarChart`, `stockChart`, `surfaceChart`, `ofPieChart` — long tail; defer

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

#### Not done for `barChart`

Each of these is known-missing rather than merely absent:

* **Data labels.** `c:dLbls` is read in full — `showVal`, `showCatName`, `showSerName`,
  `showPercent`, `dLblPos`, per-point `c:dLbl` overrides, text and box styling — and
  nothing is drawn from it. No chart in the corpus switches any of them on.
* **Rotated category labels.** PowerPoint rotates them 45° when they will not fit, which
  is what `real-financial-report.pptx` slide 3 does; we draw them horizontally and they
  overlap. That deck's chart3 is the one place our layout is badly wrong (bottom inset
  27.3 pt against PowerPoint's 69.5 pt) and it is entirely this.
* **Axis titles**, **minor gridlines and minor ticks**, **`c:dTable`**, and manual
  `c:layout` for the plot area or the legend.
* **Secondary axes.** A `c:barChart` group is tied to its axes through its own `c:axId`
  list, which is the hard part and is done; a second value axis is then mostly drawing.
* **Log scales** and `c:tickLblSkip` / `c:tickMarkSkip`. `c:crosses` and `c:crossesAt`
  move the category axis but have only been measured at zero.
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
                                      ├─▶ Phase 2  (SmartArt — DONE)
                                      ├─▶ Phase 4  (EMF previews — DONE)
                                      ├─▶ Phase 5.1 DONE / 5.2  (small gaps)
                                      └─▶ Phase 3  (charts — reader + barChart done)
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
3. ~~**Phase 4 EMF previews**~~ — done; the S estimate held.
4. ~~**Phase 2 SmartArt**~~ — done; the S estimate held.
5. ~~**Phase 5.1 shapes**~~ — done. Not the additive, near-zero-risk job it looked
   like: it turned up two arc-conversion bugs that had been silently misdrawing
   custom geometry, and it replaced hand-transcription with a spec compiler.

Phase 3 is started: the reader and `barChart` are done and measured, and the shared
infrastructure the other chart types need -- value domain, tick selection, number
formatting, gridlines, legend layout, plot-area rectangle -- is built. `lineChart` is
the next one worth having and should be cheap now; data labels and rotated category
labels are the two gaps inside `barChart` itself.

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

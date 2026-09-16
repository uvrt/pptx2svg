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
| **pptx2svg** | **yes**, EOT+MTX, `fsType`-gated, pure Python | **generated metric table, or the embedded face's own** | clone set + `font-substituted` |

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

2,059 tests pass. The pipeline is `opc → parse → resolve → render → png`; each stage is
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

* ~~**CJK label width**~~ — the East Asian face now reaches measurement, and the item
  that used to head this list was **half right about the cause and wrong about the
  number**. See *The East Asian face cascade* below: the face was genuinely being
  dropped, and fixing it takes 6.83 pt off the corpus radar's radius — but the 19.8 pt
  attributed to chart3's *width* is not a width error at all. PowerPoint's own export
  draws `プラットフォーム` at 12.000 pt a glyph, which is the 96 pt we measure.
* **The rotated-label reserve has no cap**, and PowerPoint's does. That is what chart3's
  19.8 pt actually is, and there is now a candidate rule for it (below) that one probe
  deck would settle.
* **PowerPoint ellipsis-truncates a category label that will not fit** and nothing here
  does. Measured on `real-financial-report.pptx`: its rotated bar axis drew `プラット…`
  for an eight-character category, and its radar drew `海外売上…` and `従業員満…` for
  six-character ones. We draw all three in full.
* **A manually laid out legend** is the whole of slide 4's remaining chart error.
* **Text displacement of 1–2 px** on body copy is what slides 6 and 7 are made of, and
  `tools/fidelity.py`'s own docstring warns that SSIM is unusually sensitive to exactly
  that on thin high-contrast content.

The bold/italic and paragraph-spacing defects that used to head this list are **fixed**:
bold now inherits through the placeholder cascade, `spcBef` and `spcAft` add rather than
collapse, and `spcPct` is a share of the line height rather than the font size. Slide 6
went 0.0977 → 0.7372 on the spacing fix alone.

### The East Asian face cascade

The question a run has to answer before anything can measure it: *which face draws the
kana?* Three names compete — the run's `<a:ea>`, the theme font collection's `<a:ea>` and
its `<a:font script="Jpan"/>` list — and until now the answer was **none of them** unless
the run named one itself. Where it did not, `RunProperties.font_family_ea` came through as
`None` and `DefaultTextMeasurer` measured every ideograph with the *Latin* table.

That is worse than it sounds, and the reason is in `text/metrics.py`: **`cjk_width` is
1.0 em in every table, Latin ones included**, because `tools/extract_font_metrics.py`
writes `units_per_em` when the face has no glyph for its probe kanji. So a Latin face
never says "I cannot draw this". It answers "one em" and the layout is computed from a
constant while the rasteriser draws with whatever it finds.

Three places dropped it, and they had to be fixed together:

| where | what it dropped |
| --- | --- |
| `resolve/text.py` | `font_family_ea` was the run's own `<a:ea>` and nothing else |
| `render/text.py` | `_needs_script_split` needs *both* families, so a run with no `<a:ea>` never split — and the East Asian chunk is the only place `jpan_fallback_font` was ever consulted |
| `resolve/chart.py` | `ChartFont`, `font_box` and `text_width` carried one family, and `_label_body` emitted `RunProperties` with no `font_family_ea` at all |

#### What the precedence actually is

Read out of **PowerPoint's own PDF export of `real-financial-report.pptx`**, which is the
one place in the corpus where every competing explanation is ruled out. Its five charts
name `<a:latin typeface="Arial"/>` and stop; its theme writes `<a:ea typeface=""/>` in
both collections with `<a:font script="Jpan" typeface="游ゴシック"/>` beside it. The
export embeds **YuGothic-Regular** for every Japanese category label and **ArialMT** for
the Latin runs *inside the same labels*.

1. the run's `<a:ea>`, when it names a face that can draw East Asian text;
2. the theme collection's `<a:ea>`, same condition;
3. the theme's `<a:font script="Jpan"/>` (then `Hans`, `Hant`, `Hang`);
4. the Latin face, as a last resort so the emitted stack is never empty.

Two rules in there are measured rather than read off the spec:

* **An empty `typeface=""` is not a name.** Every theme in this corpus writes one.
* **A name with no East Asian glyphs loses to one that has them.**
  `real-basic-theme.pptx` writes `<a:ea typeface="Raleway"/>` on nine of its runs and
  PowerPoint drew their Japanese in MS Gothic and MS Mincho — so it did not honour the
  name either. `covers_east_asian` is the test, and it is a field on `Substitution`
  rather than something derived: `cjk_width` cannot distinguish a face that draws a kanji
  at one em from one that cannot draw it, and counting per-character rows calls Cambria
  Japanese (four bracket forms) and ＭＳ ゴシック not (monospace, so no row disagrees with
  `cjk_width`).

**Body before heading**, also measured: the theme offers `游ゴシック Light` as its major
Jpan face and `游ゴシック` as its minor, and the export used the Regular. `jpan_fallback`
in `__init__.py` had it the other way round and put a light weight behind body copy.

#### Before and after, against PowerPoint's own export

The oracle is the *pen span* — first glyph origin to last glyph origin — read out of the
PDFs with `pypdfium2`, which needs no knowledge of the final glyph's advance and is
therefore exact.

| string | face PowerPoint drew | drawn | ours before | ours after |
| --- | --- | --- | --- | --- |
| `デジタル` (12 pt) | YuGothic-Regular | 48.000 | 48.000 | 48.000 |
| `グローバル` (12 pt) | YuGothic-Regular | 60.000 | 60.000 | 60.000 |
| `その他` (12 pt) | YuGothic-Regular | 36.000 | 36.000 | 36.000 |
| `DX投資額` (12 pt) | ArialMT + YuGothic | 52.669 | 52.669 | 52.669 |
| `CO2削減` (12 pt) | ArialMT + YuGothic | 48.674 | 48.674 | 48.674 |
| `前年同期` (10 pt) | YuGothic-Regular | 30.000 | 30.000 | 30.000 |

**Every CJK string on every chart in that deck already measured to 0.000 pt, before and
after.** That is the refutation, and it is worth stating plainly because this file
asserted the opposite for two sections: `プラットフォーム` was said to come out 96 pt
where PowerPoint "laid it out in a substituted CJK face at about 68". The 68 was
`ROTATED_LABEL_INSET_PT + width · sin 45` inverted through PowerPoint's 69.538 pt inset —
our own formula run backwards, never a measurement. PowerPoint draws all eight glyphs at
exactly 12.000 pt, one em of Yu Gothic, for 96 pt.

The widths coincided because Arimo's `cjk_width` is 1.0 em and Yu Gothic's kana and
ideographs are too. The *line box* did not, and that is the part the fix collects:

| | line box at 12 pt | radar reserve | radius |
| --- | --- | --- | --- |
| Arial, as read before | 13.406 (1.117 em) | 11.61 | 52.39 |
| 游ゴシック → Noto Sans JP's table | 17.376 (1.448 em) | 16.61 | **47.39** |
| PowerPoint | — | 18.44 | **45.56** |

6.83 pt of radius error becomes 1.83. The residual is the fitted reserve, not the face:
`RADAR_LABEL_RESERVE_LINES`/`_PT` were fitted to five Latin probes, and reproducing 18.44
exactly wants a 1.5696 em line box, which is neither our Noto Sans JP table's 1.448 nor Yu
Gothic's own `hhea` figure. Whether the reserve has a term only CJK exercises is
unmeasured — one CJK radar in the corpus and no probe deck for it.

Where the face choice *does* move a width is `real-basic-theme.pptx`, whose `<a:ea>` is a
Latin face. Its kana went from Raleway's 1.0 em guess to ＭＳ Ｐゴシック's real
proportional advances: `かじょうがき１` at 13 pt goes 78.000 → 65.559, `たいとる` at 42 pt
goes 126.000 → 110.953. PowerPoint's export of that deck drew MS **Gothic** at a flat
1.0 em, so our numbers move *away* from that PDF — which is exactly why the harness skips
the deck: the exporting machine had no MS PGothic either. This is the same trade the
`text/fontmap.py` docstring already commits to for `sample.pptx`, and it is the one that
keeps measure-equals-draw true: the stack we emit now starts at a face we ship and can
draw, instead of at one with no kana in it.

#### What this turned up and did not fix

* **PowerPoint truncates, we overflow.** `プラット…`, `海外売上…`, `従業員満…` — a
  category label wider than its allowance is cut and ellipsised, not wrapped, because
  `wrap_label` breaks at whitespace and CJK has none. The radar's allowance looks like the
  `RADAR_LABEL_MAX_FRACTION` cap we already carry: 0.25 × 243.243 pt region = 60.8 pt, and
  both truncated labels came back at five cells (60.0 pt) where the full string is six
  (72.0). We wrap onto a second line instead.
* **A candidate cap for the rotated reserve, on three observations.** Inverting
  `ROTATED_LABEL_INSET_PT + width · sin 45` through PowerPoint's insets gives an implied
  label width of 90.85 and 91.88 pt for the two probes that exceeded the cap (frame
  181.1024 pt tall) and 68.09 pt for chart3 (frame 135 pt). Both ratios are 0.505 of the
  frame height, and a straight line through them has an intercept of 0.01. It is **not
  implemented**: three points with a 1.03 pt spread in two of them, and the two probes
  should be *equal* under a pure cap and are not. One probe deck sweeping frame height
  would settle it; this is exactly the shape of the Caladea claim, so it is recorded
  rather than shipped.
* **`lang` may select between the script list and an application default.**
  `real-financial-report.pptx`'s table cells name no typeface, inherit `+mn-ea` → 游ゴシック
  — and PowerPoint drew them in **MS Gothic**, not Yu Gothic, on a machine that has both.
  Every one of those runs is `lang="en-US"`. The chart text, which has no `lang` at all,
  went to the script list; so does `sample.pptx`, which has no `lang` anywhere and whose
  Japanese resolved through its theme's Jpan entry. The theory fits all three and cannot
  be acted on: "PowerPoint's default Japanese font for a non-Japanese run language" is a
  property of the machine, not of the deck, and the script list is the only deck-derived
  answer available. Recorded because the alternative is to pretend the cascade explains a
  case it does not.
* **`ascender_ratio` still prefers the Latin face** even for a run with no Latin character
  in it (`text/measure.py`), and chart `first_baseline` follows it deliberately so the two
  stay in step. Changing one without the other would unstitch measure-equals-draw.
* **The legend band still reads the Latin line box.** `ChartFont.box_for` is consulted by
  `_bottom_label_band` and `_radar_geometry` only. `real-financial-report.pptx`'s doughnut
  legends in Japanese, but its entries are laid out down a right-hand band where the line
  box is not what sets the width, and no probe measures a CJK legend.


---

## Fonts — **done**

The problem, stated precisely: layout was computed from a table of advance widths, the
SVG named a `font-family`, and **nothing made those two agree**. On a machine without
Calibri the rasteriser substituted something else without a word — rendering one string
in Calibri, Carlito, Aptos, Noto Sans JP and Lato produced five byte-identical PNGs. Text
appeared, at widths nothing had computed.

### The shape of the fix

0. **Read the deck's own fonts first**, where it carries any. Added in Phase 6, after the
   rest of this section was written: a `<p:embeddedFontLst>` face is decoded, rights-checked
   and used for both measurement and drawing, and it takes precedence over every rule below.
   Nothing a bundle ships can beat the file the author laid the deck out with -- measured,
   the bundled Raleway and the Raleway inside `real-basic-theme.pptx` disagree on 338 of 340
   advance widths.
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

  **Still unfixed, but no longer a design question.** Phase 6 needed exactly this
  plumbing and built it: `DefaultTextMeasurer(extra_metrics=...)` takes a
  family -> `FontMetrics` map that wins over the static tables, and `fonts/sfnt.py`
  builds one from a font file using only the standard library, so no extra is involved.
  Pointing `font_dirs` at that is now a small change. It is left alone here because it
  alters existing behaviour for every caller of a documented parameter, which deserves
  its own decision rather than being a side effect of the embedded-font work.

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
| **−9074 has a fourth cause, and it survives sessions** | A presentation left open from an *earlier* session — windowless, unlisted in the Window menu, raising no dialog — makes every export fail −9074 with nothing visible to blame. One was found on arrival as `zz-bisect [Repaired]`, days old. `get name of every presentation` is the diagnostic the other health checks miss: it names the zombie where `count of presentations` and a dialog sweep both come back clean. `pkill` and relaunch is again the only fix. **Run the name check first** — −9074 genuinely does not mean “unapproved path”. |
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

### 0.2 Snapshot VRT (the regression net) — **done**

Ground truth needs PowerPoint; regression detection does not. `tests/vrt/` carries one
committed SVG per slide of every deck in `tests/fixtures/` — 16 files, 195 kB — and
`tests/test_vrt.py` re-renders them and compares byte for byte. `--update-snapshots`
rebaselines, and skips rather than passes, so the flag cannot produce a green run.

**SVG, not the PNGs this section used to ask for.** A PNG additionally encodes resvg's
version and its antialiasing, so bumping `resvg-py` would fail every fixture at once and
read as a rendering regression — a gate that cries wolf on a dependency bump is a gate
people learn to rebaseline through. It is also 3x the bytes (587 kB against 195 kB on this
corpus) and, already compressed, those bytes never shrink, in the pack file or in history,
on every rebaseline, forever. The SVG is the part we author; the rasteriser is somebody
else's output and `tools/fidelity.py` already exercises it.

The font bundle turns out **not** to be what makes this work, which is worth correcting
here because this section claimed it was. The SVG path opens no font file at all: widths
come from the generated tables in `text/metrics.py` and the `font-family` is a name for a
rasteriser to resolve later. Measured — with the bundle and without it, all sixteen
documents are identical to the byte; only `ConvertOptions.warnings` differ, which is why
warnings are not snapshotted. The bundle is what makes a *PNG* reproducible. So these
snapshots are meaningful on a bare checkout too, and nothing has to skip.

What was verified rather than assumed, because a snapshot that differs per machine is
worse than none:

- **`PYTHONHASHSEED`.** Set iteration order varies between processes and never within one,
  so rendering twice in one process proves nothing. Two subprocesses under two fixed,
  different seeds agree, and a test keeps it that way.
- **Line endings.** `.gitattributes` pins `tests/vrt/**.svg` to `eol=lf` so Git on Windows
  cannot rewrite them on checkout; the test normalises on read as well, and asserts the
  render itself emits no `\r` so that normalisation cannot mask a real change.
- **Encoding.** Every read and write names UTF-8. These files carry Japanese, and this
  repository has already lost a Windows run to `read_text()` following the locale.
- **Float formatting across platforms.** Rounding is to three decimals and formatting is
  Python's own dtoa, both platform-independent. libm is not: `sin`, `cos`, `log10` and
  friends may differ by an ulp between glibc, macOS and MSVC. Measured rather than argued
  — perturbing *every* transcendental in the render by one ulp, systematically, in one
  direction, moved not a byte of fifteen of the sixteen snapshots. The sixteenth is the
  finding below, and it has since been fixed; nothing in the corpus moves either way now.

#### Found while proving that: `floor(log10(span))` amplified an ulp into a different chart — **fixed**

`resolve/chart.py` picked an axis unit with

    unit = 10.0 ** math.floor(math.log10(span))

`real-financial-report.pptx` slide 4 is a radar spanning exactly 0..100, so `log10` returns
exactly 2.0 and the unit is 100. One ulp low — `1.9999999999999998` — and `floor` gives 1,
the unit becomes 10, and the slide re-lays out: 29,659 characters of SVG become 34,639.
That was the *only* byte anywhere in the corpus that a 1-ulp libm difference could move,
and it did not move by a digit, it moved by a factor of ten.

Every mainstream libm returns 2.0 here, because the true value is exactly representable and
correct rounding therefore requires it; only a merely *faithful* implementation is allowed
to return the value below, which is why no user ever hit this. It was worth fixing anyway:
the failure mode is silent, total, and would have surfaced as one unexplained snapshot
mismatch on somebody's CI leg.

**The fix is `_decade()` in `resolve/chart.py`**, used by both sites that wanted a power of
ten (the axis unit and the 1-2-5 ladder in `_next_nice_unit`). It takes
`floor(log10(value))` as a *guess* and then checks it against the input, correcting by one
step in whichever direction the logarithm rounded; comparing a value with a power of ten is
an ordinary float comparison where comparing logarithms is a transcendental. Any `log10`
accurate to better than a whole decade — every real one — now gives the same axis. Two
details that were not free:

- The powers come from `float(f"1e{n}")`, CPython's own correctly-rounded parser, not from
  `10.0 ** n`, which is libm again and raises `OverflowError` at `1e309` where the parser
  returns `inf`. That `inf`, and the `0.0` at `1e-324`, are what let the correction walk
  off either end of the double range and stop.
- A span is a subtraction of authored decimals and lands *just under* round numbers:
  `0.24 - 0.14` is `0.09999999999999998` and `1.13 - 1.03` is `0.09999999999999987`. A
  strict comparison puts the second in the decade below and steps that axis by 0.005.
  `log10`'s rounding used to supply a forgiving window here by accident — which is why the
  old behaviour was right on the first of those and wrong on the second — so the window is
  now explicit and even: `_DECADE_SLACK = 1e-12`, wider than any such residue (a few ulps,
  ~1e-16 relative) and nine orders of magnitude narrower than a span like 9.99 that misses
  a decade because it was written that way. That window — between one part in 1e16 and one
  in 1e12 below a power of ten — is the one place the fix deliberately changes an answer,
  and it changes it towards the humane one. No snapshot and no fidelity score moves.

The tripwire that stood in for the fix (`test_the_platform_agrees_about_log10_of_a_power_of_ten`)
is **gone**, replaced by `test_a_one_ulp_error_in_log10_does_not_move_a_snapshot`: it
renders that same deck with `log10` nudged an ulp each way and demands the committed bytes
back. Asserting a property of the platform would now fail a *correct* build on a merely
faithful libm — the opposite of what a gate is for — while asserting the property of our
own code is what the snapshots actually depend on. `tests/test_chart.py` holds the
unit-level version, driving `_decade` with the perturbed logarithm directly.

`real-college-template.pptx` is **not** snapshotted: a render of it contains its text and
its images, so committing one would redistribute a deck that is not ours to redistribute.

Two layers, two jobs: snapshot VRT runs everywhere and catches regressions; the PowerPoint
oracle runs on this Mac and catches *being wrong in the first place*. The gap between them
is wider than it looks — `tests/fidelity-baselines.json` skips five of the seven committed
decks (PowerPoint substituted the same CJK face, or this machine lacks Noto Sans JP), so
for five of these seven, the snapshot is the *only* thing watching, and it is watching for
change rather than for correctness. `tests/vrt/README.md` says so at the top and lists the
known defects the committed bytes freeze in.

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

**Effort: XL. 3.1 is done and 3.2 has ten of its types; combo charts, the secondary axis
and `surfaceChart` are not.**

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

### 3.2 Renderer — ten types **done**, one deferred with its reasons

1. ✅ `barChart` — clustered, stacked, percentStacked, `barDir` col and bar
2. ✅ `lineChart` — markers, smoothing, blanks, the real fixture on slide 2
3. ✅ `pieChart` / `doughnutChart` — hole, rings, explosion, the real fixture on slide 3
4. ✅ `radarChart` — standard, marker and filled, the real fixture on slide 4
5. ✅ `areaChart` — standard, stacked, percentStacked, negatives, `crossBetween`
6. ✅ `scatterChart` — two value axes, all five `c:dLblPos`, splines
7. ✅ `bubbleChart` — `c:bubbleSize`, `c:bubbleScale`, `c:sizeRepresents`,
   `c:showNegBubbles`
8. ✅ `ofPieChart` — both `c:ofPieType` forms, all five splits, `c:serLines`
9. ✅ `stockChart` — `c:hiLowLines`, `c:upDownBars`, three and four series
10. ⛔ `surfaceChart` / `surface3DChart` — **measured and deferred**; see below

**Still warning `chart-unsupported-type`:** `surfaceChart` / `surface3DChart`, and the
whole ChartEx family. Their 3-D spellings degrade through `parse/chart.flat_chart_kind`
and then warn too. **No chart in the corpus warns**, and none ever did once 3.1 landed.

Data labels are drawn for all ten, at every `c:dLblPos` each type accepts — except
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

* **The corpus radar's radius was 6.83 pt too large; 5.00 of that is fixed.** Its
  category labels are Japanese, and `font_box` read the `<a:latin typeface="Arial"/>` the
  axis names — a face with no kana in it — for a label made of nothing else. PowerPoint's
  export embeds **YuGothic-Regular** for them, which is the theme's
  `<a:font script="Jpan"/>`; `ChartFont` now carries it, and its 1.448 em line box takes
  the reserve from 11.61 to 16.61 and the radius from 52.39 to 47.39 against PowerPoint's
  45.56. See *The East Asian face cascade*. **The 1.83 pt left is the reserve, not the
  face**: `RADAR_LABEL_RESERVE_LINES`/`_PT` were fitted to five Latin probes, and
  reproducing 18.44 exactly wants a 1.5696 em line box, which is neither our Noto Sans JP
  table's 1.448 nor Yu Gothic's own `hhea` figure. One CJK radar and no probe deck, so
  whether the reserve has a term only CJK exercises stays unmeasured.
* **A radar category label that will not fit is truncated by PowerPoint and wrapped by
  us.** Its export draws `海外売上…` and `従業員満…` where the categories are six
  characters. Five cells (60.0 pt) against the `RADAR_LABEL_MAX_FRACTION` allowance of
  0.25 × 243.243 = 60.8 pt, so the cap we already carry looks like the right one and the
  response to it is not: `wrap_label` breaks at whitespace and CJK has none.
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

#### Area, measured

Eighteen probe charts in the same 220.4724 x 181.1024 pt frame as the bar sweep, plus four
tie-breakers, exported by PowerPoint 16.106 and read back as exact path vertices.

**The plot rectangle is a bar chart's, to the last decimal.** All four insets agree on
every one of the eighteen — `bare` 21.073 / 11.000 / 11.102 / 24.965, `font14` 26.907 /
11.000 / 13.545 / 32.353, `title` top 40.803, `legend-b` bottom 49.048 — worst residual
across the sweep **0.23 pt**. So is the value axis: 3, 4, 5 of data gives 0..6 by 1, the
bar's strictly-outward rule and *not* the radar's stop-at-the-data one.

What is new, and six of the seven refute the obvious reading:

* **An area chart that states no `c:crossBetween` draws as `midCat`.** Three probes: with
  `val="between"` the three vertices land on the band centres (52.473 / 115.292 /
  178.072); with `val="midCat"` and with **no element at all** they are byte-identical to
  each other on the plot's edges and midpoint (26.433 / 108.015 / 189.632). A line chart's
  absent case is not measured — every line chart in the corpus and in every probe states
  `between` — so the default there is left alone and this one is scoped to the area.
* **`midCat` narrows the plot** because the first and last category label are now centred
  on its own edges and half of each hangs outside. Left becomes
  `max(label column, 11.0 + half the first label)` and right `11.0 + half the last`:
  `Reader` is 30.859 pt and `Renderer` 39.673 pt, giving 26.433 and 189.636 against
  PowerPoint's 26.433 and 189.632.
* **The fill closes to the zero line, not to the plot's floor.** The negative probe's
  polygon returns along y = 117.069, which is where its -3..6 axis puts 0, with the plot's
  own bottom 53 pt lower. The path therefore **crosses itself** where the line crosses
  zero and PowerPoint leaves the bow tie exactly as it falls — its fill rule is nonzero
  **winding**, which is what the renderer already does.
* **`c:invertIfNegative` does not exist on an area series.** CT_AreaSer has no such
  element; a deck carrying one makes PowerPoint open it `[Repaired]` and refuse to export.
  So the question the bar chart settled does not arise here — the file cannot ask.
* **The series are painted in series order, first at the back, and the fills are opaque.**
  Every one of the eighteen came back at alpha 255, and two probes say it is *order* and
  not size: swapping the two series' values swapped which one was hidden, and a probe
  whose first series covers the second entirely still emitted the first path first. A
  front series really does hide what is behind it. That is PowerPoint's picture.
* **No outline unless the file asks.** Eight probes stating none got a bare fill; one
  stating `<a:ln w="25400"/>` got a second, stroked copy of the whole closed path,
  baseline edge included.
* **A blank splits the run, and a run of one point draws nothing.** The
  `dispBlanksAs="gap"` probe, whose middle value is missing, drew **no area at all** —
  two runs of one point each, and a point has no area.

Stacked and percent-stacked accumulate exactly as a stacked bar does, with each band's
lower edge the running total *without* this series, traced backwards; percentStacked's
last series reaches 100% in every category. And:

* **An area data label sits at the vertical centre of its own band**, horizontally centred
  on the point. Unstacked 3/4/5 put its labels at 1.5, 2.0 and 2.5 on the value axis; a
  stacked pair put the second series' at 4, 6.5 and 5.5, the midpoints of its *segments*
  and not of the stack. Ten labels, worst residual 0.21 pt.
* **`c:dLblPos` is refused outright.** An area chart whose `c:dLbls` carries one — `ctr`
  included, which is the only value ECMA-376 lists for an area series — makes PowerPoint
  open the deck `[Repaired]`. An area label has exactly one placement.
* **The legend key is the square swatch**, 5.492 pt then 2.371 pt of gap: a bar's key, not
  the line chart's rule.

#### Scatter, measured

Eighteen probe charts plus six tie-breakers. **Both axes are value axes and there is no
category axis at all**, which is the one structural difference from every other Cartesian
chart here and the reason `_build_scatter` does not go through `_build_cartesian`.

**The brief's question — does the plot-area code assume a category axis? — is yes, and
specifically the bottom of it.** `_bottom_label_band` takes a category list and decides
between a level band, a wrapped one and a turned one; `_labels_rotate` and `wrap_label`
are written around a *band width*, which a row of numbers does not have; and
`_draw_labels` picks its placement off `barDir`. None of that applies to a scatter. What a
scatter's bottom actually is, measured, is the **horizontal bar chart's**: one plain line
of value labels centred on their ticks, with half the last one hanging past the plot's
right edge. So the band is `_bottom_label_band(font, [], width)` — the same formula with
an empty category list — and nothing else of that machinery is reached.

* **The plot rectangle** is the horizontal bar's on all six discriminating probes:
  `bare` 21.073 / 13.670 / 11.102 / 24.965 (the right inset is `11.0 + half of "6"`),
  `font14` 26.907 / 14.740 / 13.545 / 32.353, `x-deleted` 11.000 right and 11.102 bottom
  with no labels to reserve for. Worst residual **0.32 pt** over eighteen.
* **Neither axis is anchored at zero, and a scatter is the only type that is not.**
  `nice_axis_scale` forces `min(0, data)` because a bar that does not start at its axis is
  a different picture. Applied to a scatter it destroys the chart: a decade of years
  against a measurement collapses into a 1% sliver of a 0..3000 axis. Six probes bracket
  when PowerPoint lets go — the near end has to sit past **5/6** of the far one, measured
  to [0.80, 0.84) by the pair that straddles it:

  | x data | near/far | PowerPoint |
  | --- | --- | --- |
  | 10..50 | 0.20 | 0..60 by 20 |
  | 40..50 | 0.80 | **0..60 by 20** |
  | 42..50 | 0.84 | **40..55 by 5** |
  | 100..104 | 0.96 | 98..106 by 2 |
  | 2010..2020 | 0.995 | 2005..2025 by 5 |
  | −50..−10 | mirrored | −60..0 by 20 |

  The unanchored extent rounds strictly outward at **both** ends where an anchored one
  holds its low end at zero — 100..104 comes back 98..106, a whole unit clear each way —
  and the unit and the four-interval cap are unchanged. Ten of the eleven scatter axes
  measured now reproduce exactly. **Only a scatter is measured**: a line chart of
  temperatures has the same problem and no probe has ever shown what PowerPoint does with
  one, so every other type keeps the anchor.
* **`c:crosses` belongs to the axis it is written on**, on a scatter as on a bar: it says
  where *that* axis crosses the perpendicular one, so the vertical line's position is a
  question for the y axis even though the answer is an x coordinate. Only `autoZero` is
  measured, and the reading is `_category_axis_position`'s rather than a second one
  invented for scatters — which is also how `c:crossesAt` comes along for free.
* **The two negative corners mirror the bar chart's.** A negative *x* range floats the
  value axis into the plot — drawn at 111.088 pt, not at the plot's left edge 95.7 pt away
  — and the y labels go with it, right-aligned the same `descent + 0.645 em` from the axis
  the column always uses; the left inset is then the first x label's overhang, 15.373 pt
  measured against 15.373 predicted. A negative *y* range moves the x labels up beside the
  zero line and the bottom band disappears.
* **The x axis is the coarse one, and the cap is four intervals and not five.** 1..5 of
  data comes out 0..6 **by two** where the same span on the y axis takes ones. The
  constant was five on the single horizontal-bar observation, which only bounds it below
  six; the discriminating probe is `x-float`, whose 0.5..4.5 rounds to a *five*-interval
  0..5 at unit 1 and was coarsened to 0..6 by 2 anyway, against `x-neg`'s four-interval
  -4..4 by 2 which was kept. The bracket is [4, 5).
* **`c:scatterStyle` decides nothing.** `marker`, `line` and `lineMarker` produced
  byte-identical output — line *and* markers in all three. What turns either off is the
  series' own markup: `<a:ln><a:noFill/></a:ln>` for the line, `<c:symbol val="none"/>`
  for the marker, each measured. That was a live defect on the line chart too: a `noFill`
  stroke resolves to `None` exactly as an absent `c:spPr` does, and the default 1.5 pt
  line was being substituted for it.
* **A blank breaks the run**, and the vertex list alone does not say so. The path through
  a missing middle y has four points, which reads as unbroken; its segment *kinds* are
  move, line, move, line. Reading coordinates without reading the operators is how that
  gets missed — it was, until the render was put beside PowerPoint's. Points are joined in
  **the order the file lists them**, not sorted by x: the unsorted probe's path runs
  3, 1, 5, 2, 4.
* **All five `c:dLblPos` values are the line chart's geometry read off a different edge of
  the marker.** `r` — the default when the file states none — and `l` put the label's near
  edge a marker radius plus 0.6 em from the point, 9.000 pt measured and 9.000 predicted
  on both sides; `t` and `b` put its line box a radius plus the bar's 4.85 pt gap away;
  `ctr` is the ink centre. Worst residuals 0.20, 0.20, 0.21, **0.91** and 0.26 pt.
* **The legend key is the line chart's rule**, 19.200 pt with the marker on its midpoint.

#### Two constants the new probes corrected, and one they finally measured

Both of these are shared with `lineChart` and both were carried on assumption:

* **The default marker is 6 pt, not ECMA-376's 7.** The radar already used 6 from a probe
  whose *diamond* measured 5.76 pt across — a diamond, so the reading depended on the tips
  falling inside PowerPoint's 0.24 pt output grid. Two probes settle it with an
  axis-aligned shape and no such argument: the scatter two-series probe's second series
  states no `c:marker` and its **square** measured exactly 6.000 x 6.000, and a line chart
  with an explicit `c:size val="6"` square measured the same 6.000 on the same export. The
  scatter's `r` data label confirms it a third way: its 9.000 pt offset is radius plus
  0.6 em, which is 9.0 at a 6 pt marker and 9.5 at a 7 pt one.
* **An absent `c:smooth` smooths.** It is a chart boolean, so the element being missing is
  not the same as `val="0"` — the trap `parse/chart._flag` exists for, and this reader was
  using `bool(None)`. Three probes: `val="1"`, the element absent, and `val="0"` came back
  as four cubics, the *same* four cubics, and a four-segment polyline. PowerPoint's own
  writer always emits the element, so no corpus deck moves; a hand-written one does.
* **`c:smooth`'s tension is now measured, and both ends of the spline were wrong.** The
  roadmap recorded the curve as "a Catmull-Rom spline, which has the right shape" with the
  tension unmeasured. It is the plain 1/6: a five-point series exports as four cubics whose
  interior controls reproduce `c1 = p1 + (p2 - p0)/6` and `c2 = p2 - (p3 - p1)/6` to the
  0.001 pt the PDF prints, on a line chart and a scatter alike. **The terminal controls
  are a third of their own chord, not a sixth** — 87.073 against the 90.527 that
  duplicating the end point gives, on a chord of 20.72 pt — so the phantom point is a
  *reflection*, `p0 = 2*p1 - p2`, and reflecting it reproduces both ends exactly.

**What it bought.** No corpus deck holds an area or a scatter, so nothing scored moves and
that is the point: `authoring-integration` holds at 0.9327/0.9984, `table-test` at
0.9895/0.9984 and `real-college-template` at 0.8003/0.8753, all unchanged to four decimals
despite the marker size, the smooth default, the spline ends, the `noFill` stroke and the
bottom-axis interval cap all being shared with chart types those decks *do* hold. The
improvement is in the probe decks, scored the same way:

| probe deck | before | after |
| --- | --- | --- |
| area (18 charts) | 0.6165 / −0.0251 | **0.9562 / 0.9990** |
| scatter (18 charts) | 0.0175 / −0.1142 | **0.8695 / 0.7945** |
| tie-breakers (12) | 0.3148 / −0.0547 | **0.8921 / 0.8624** |
| smooth (6) | 0.4291 / 0.3285 | **0.8768 / 0.6820** |
| axis minimum (6) | 0.4870 / 0.9201 | **0.5806 / 0.8859** |

The axis-minimum deck is the one that stays low, and the reason is named above: four of
its six charts reproduce exactly and the other two are the 120..160 y axis, where we draw
five gridlines against PowerPoint's nine. That is the unsolved density rule and not the
zero anchor the deck was built to measure.

The three that stay under 0.95 are decks of thin curves and markers on white, where SSIM
is punishing about a pixel of antialiasing; put side by side at 1400 px the renders are
indistinguishable apart from the tick density noted below. The probe decks are throwaway
and were deleted; their generators are `tests/test_chart.py`'s `area_chart_xml` and
`scatter_chart_xml`.

#### Bubble, measured

Thirty probe charts across three decks in the same 220.4724 x 181.1024 pt frame as the
scatter sweep, exported by PowerPoint 16.106 and read back as exact path vertices. Every
drawn circle came back axis-aligned and square to 0.001 pt, so its bounding box *is* its
diameter, and every number below is PowerPoint's own.

**A bubble is a scatter plus a third dimension, and the "plus" is the whole of the work.**
The plot rectangle, the unanchored axes, the coarse x axis, the blank handling and the
`c:crosses` reading are the scatter's and were reused untouched: `bare` reproduces
21.073 / 13.670 / 11.102 / 24.965, the scatter sweep's own row.

* **The size-to-radius map is area-proportional, and the reference is the largest size.**
  Sizes 1, 4, 9 drew 13.162, 26.323 and 39.485 — exactly 1:2:3, so the *area* is
  proportional to the size and the diameter to its square root. `<c:sizeRepresents
  val="w"/>` on the same data drew 4.387, 17.549 and 39.485, exactly 1:4:9. The largest
  is 39.485 in **both**, and in four probes whose size distribution differs wildly
  (1,2,3 / 1,4,9 / 5,5,5 / 1,2,100), so the reference is the maximum rather than the sum.
  A two-series probe makes it **global**: a series topping out at 9 drew 27.920 beside
  one topping out at 18, which is 39.485 × sqrt(9/18).

* **What the largest bubble is sized against is not the plot rectangle.** This is the one
  that would have been guessed wrong: `font14` and `font8` move all four plot edges and
  draw the *same* 39.485 pt bubble, while a right legend (plot 137.805 x 145.035) draws
  37.511 and a top or bottom one (185.729 x 120.952) draws 33.928. The region is the
  **frame inset by 5 pt on every side**, less the title band and the legend band.

  The 5 pt is solved, not fitted. The legend band is 24.083 pt and takes the diameter from
  39.485 to 33.928, so the height it eats into is `24.083 / (1 − 33.928/39.485)` = 171.12,
  which is the 181.102 pt frame less **9.98**. The side-legend probe then falls out with
  no constant of its own: its reserve is the 47.924 pt the plot gives up, and
  `(210.472 − 47.924) × 0.2308` is 37.513 against 37.511 drawn. So does the title probe:
  `171.102 − 29.700 = 141.402`, times the same factor, is 32.631 against 32.631 drawn.

* **`c:bubbleScale` does not scale the diameter.** It is a soft clamp,
  `D = M · s / (s + 1000/3)` where `M` is the short side of that region — linear in *s*
  while small, and approaching the region's own short side as *s* grows, so a bubble can
  never fill more than the region no matter what the file asks for:

  | scale | PowerPoint | predicted |
  | --- | --- | --- |
  | 1 | 0.512 | 0.5117 |
  | 10 | 4.983 | 4.9836 |
  | 25 | 11.937 | 11.9374 |
  | 50 | 22.318 | 22.3177 |
  | 75 | 31.427 | 31.4269 |
  | 100 | 39.485 | 39.4851 |
  | 150 | 53.101 | 53.1007 |
  | 200 | 64.163 | 64.1634 |
  | 300 | 81.048 | 81.0485 |

  Worst residual **0.0005 pt** over nine observations. The obvious linear reading is
  refuted at both ends — it predicts 19.74 at 50 where PowerPoint drew 22.318, and 78.97
  at 200 where it drew 64.163 — and so is every power law: the implied exponent is 0.824
  on one side of 100 and 0.701 on the other.

* **A zero size draws nothing and a negative one draws hollow.** The `-4, 0, 9` probe
  emitted **two** circles: the 9 in the series colour and the −4 at its magnitude, white
  with a black 0.75 pt outline — the drawing a negative bar gets. `<c:showNegBubbles
  val="0"/>` removes it; the element absent and `val="1"` are identical, so the default is
  to draw it.

* **No line, no marker, no outline.** Every disc on every probe came back filled and
  unstroked, with no connecting stroke and no marker anywhere.

* **The legend key is the square swatch, not the scatter's rule.** Both legend probes drew
  a 5.492 pt key. This is not cosmetic: the legend reserve feeds the sizing region as well
  as the plot, so reading the line key there put the plot 13.4 pt narrow *and* the drawn
  bubble 3.1 pt small.

* **A bubble's data label stands further off its mark than a marker's does.** `r` puts the
  label box's near edge 8.494 pt past the **disc's** edge at 10 pt — 8.504, 8.484 and
  8.494 on discs of radius 6.581, 13.162 and 19.742 once each digit's own side bearing is
  taken out — against the scatter's 6.0 off a 3 pt marker; `t` and `b` put the line box
  7.25 and 6.76 pt out against the scatter's 4.85. `ctr` is the ink centre, 0.22 pt out,
  which is the same slack this file records everywhere else. One font size only, so
  whether these are points or ems is **unmeasured**; they are carried as points because
  the line chart's legend key turned out that way.

Two things measured and **not** implemented:

* **A label PowerPoint cannot fit is pushed back inside the frame.** The `t` probe's
  largest bubble would have put its label above the frame's top edge, and PowerPoint drew
  it at 4.140 pt instead of the 10.07 pt gap the other two took. Where it clamps to is not
  identified from one observation, so ours goes where the rule says and off the top.
* **`<c:bubble3D val="1"/>` hangs this PowerPoint.** Twice, reproducibly: the deck opens,
  a `~$` lock appears, AppleEvents stop being serviced and the 600-second timeout fires
  with no PDF written. So the 3-D bubble is not merely unmeasured — the oracle will not
  produce an answer for it, and the recovery is to kill PowerPoint and delete the lock.

#### Of-pie, measured

Twenty-four probe charts across two decks. A wedge's path closes through the pie's centre,
which is the second-to-last point PowerPoint emits, so the centre, the radius and both
edge angles come out of the PDF exactly rather than by fitting.

**The region is the pie's own** — `_polar_region`, edge insets plus the title and legend
bands — and the two plots are packed across its full width: the first plot's left edge and
the second's right edge land on the region's own edges in every probe, and both are
centred on its middle row to 0.001 pt.

* **The split rule, five spellings, four measured on 40/25/15/10/6/4.** `pos` moves the
  **last** `c:splitPos` points (`val="4"` kept 40 and 25 and moved the other four); `val`
  moves every point **below** `c:splitPos` (`val="12"` moved 10, 6 and 4 and kept 15);
  `percent` is the same test on the point's share and it is **strict** (`val="15"` moved
  10%, 6% and 4% and kept the 15%); `cust` moves exactly the `c:secondPiePt` indices, in
  their original order, and leaves everything else — colours included — where it was.

* **`auto`, which is also what an absent `c:splitType` means, moves the last ceil(n/3).**
  This is the part the brief expected to be subtly wrong, and it is where reading the
  schema would have left it wrong: ECMA-376 documents a `c:splitPos` default of 2, and a
  six-point chart does indeed move two — but three points moved **one** and eight moved
  **three**. Five counts, 3/4/6/7/8, moved 1/2/2/3/3. `round(n/3)` is refuted twice over,
  at n=4 (it predicts 1) and n=7 (it predicts 2); `floor(n/3)` fails at n=8.

* **The aggregated slice is centred at three o'clock**, pointing at the second plot, and
  that fixes the rotation of the whole chart. Five probes: the slice runs 72..108, 27..153,
  54..126, 0..180 and 54..126 degrees clockwise from twelve, every one centred on 90. The
  second plot starts at the **same** angle the first one does.

* **It takes the colour one past the last point.** Six points came out accent1..accent6
  with the slice in the next cycle's accent1; a four-point chart put it in accent5 and a
  three-point one in accent4.

* **The packing law.** With the region width `W` and `s = secondPieSize/100`,
  `r = W / (2 + 2s + g/100)` where `g` is `c:gapWidth` in percent — of the **first plot's
  radius**, which is the part the schema does not say. Six probes, every one exact:
  198.472/4.5 = 44.105 at the defaults, /4 = 49.618 at `secondPieSize=50`, /5 = 39.694 at
  100, /3.5 = 56.706 at 25, /6.5 = 30.534 at `gapWidth=300` and /3.5 = 56.706 at
  `gapWidth=0`. `c:secondPieSize` defaults to 75 and `c:gapWidth` to 100.

* **The bar form packs by a different divisor**, `r = W / (2 + s + g/200)` — the gap
  between the pie and the bar is **half** what it is between two pies. Two probes, both
  exact: 61.068 at `s=0.75` and 66.157 at `s=0.5`. The bar is `s·r` wide and `2·s·r` tall,
  the same vertical extent a second pie of that size would have, and it stacks the first
  moved point on **top**. Only `gapWidth=100` was measured on the bar form, so the `/200`
  is the natural reading of one observation rather than a fitted slope.

* **`c:serLines` is two tangents, and its presence is the switch.** A probe with no
  element drew no connector at all; a bare `<c:serLines/>` drew two lines in black at
  0.5 pt — the axis default — and `<a:ln w="28575">` in red drew them red at 2.25 pt. Each
  line leaves a **corner of the aggregated slice**, where its arc meets the circle, and is
  **tangent to the second pie**: the upper line ran (97.051, 76.922) to (168.104, 58.528),
  where the dot product of the second pie's radius and the line direction is 0.000 and the
  drawn length 73.395 pt is exactly `sqrt(d² − r²)`. The upper corner takes the upper
  tangent point. **The bar form is not measured** — no probe put `c:serLines` on one — so
  its lines run to the bar's two left corners, which is the natural analogue and is marked
  in the code as a guess rather than left undrawn.

* **Past six colours the accent cycle stops being the plain accents**, and this is a
  *pie-family* behaviour that an ofPie merely exposes, because an ofPie always needs one
  colour more than it has points. A seven-slice chart came back with accent1..accent6
  **darkened** and the seventh a light accent1; a nine-slice one repeated exactly the same
  darkened six and then three light ones, so the variation is per **cycle** of six and not
  a function of the count. Applied to the **linear-light** value of each channel —
  `L × 0.76` for the first cycle and `L + 0.23 × (1 − L)` for the second — it reproduces
  all 27 measured channels to the byte; the round numbers either side, 0.75 and 0.25, are
  off by up to 1 and 5.

  The colour *space* matters as much as the factor, and that is a second finding: the same
  modulation done in HLS on sRGB, which is what `resolve/color._apply_luminance` does for
  DrawingML's own `lumMod`, puts accent1's blue channel at 150 against the 173 PowerPoint
  drew. **That is a real defect in the general colour transform and it is not fixed here**
  — changing it moves every deck in the corpus — so the chart ramp carries its own
  conversion and says why. A third cycle is unmeasured: no probe had more than twelve
  points, and it repeats the second's tint.

Two divergences, both measured rather than guessed at:

* **PowerPoint shrinks both plots to make room for data labels and this does not.** The
  label probe's first radius came out 37.981 against the 44.105 the same chart draws
  without them, the same 0.861 factor on both plots. One observation does not say what the
  reserve is a function of, so the plots keep their full size and the labels are laid over
  them.
* **A split that moves nothing leaves PowerPoint drawing a dark disc where the second plot
  would be** — `#404040`, filled and stroked at 3 pt, at the radius the second plot would
  have had. Ours draws nothing there. An empty circle of flat dark grey is not a picture
  worth reproducing, and the *first* plot — which is the chart — is identical either way.

#### Stock, measured

Twelve probe charts. The plot rectangle, the axis, the category bands and the legend key
came back a line chart's in every one of them.

* **A stock chart with no decorations *is* a line chart.** This refutes the brief for this
  work, which had it as "a lineChart with the lines suppressed": a `c:stockChart` with
  neither `c:hiLowLines` nor `c:upDownBars` came back as one 1.5 pt polyline per series
  with the ordinary diamond/square/triangle marker cycle on it, drawn at 6 pt. What a real
  stock chart lacks is suppressed by the **file**, which writes `<a:ln><a:noFill/></a:ln>`
  on each series; nothing in the renderer hides anything. So `_is_line` answers yes for a
  stock chart and the two decorations are drawn on top, in that order — PowerPoint emitted
  the series first, then the hi-low lines, then the bars.

* **`c:upDownBars` takes the first and the last series, and the series *order* carries the
  meaning while the labels carry none.** That is the question the brief asked, and the
  discriminating probe is a three-series High/Low/Close chart with `c:upDownBars` on it:
  PowerPoint drew all five bars **down**, from each category's High to its Close. A
  four-series Open/High/Low/Close chart drew three up and two down, which is where close
  exceeds open and where it does not. Neither reading of the *names* survives that pair;
  `series[0]` and `series[-1]` reproduces both.

* **The bar width is `band / (1 + gapWidth/100)` and the default `gapWidth` is 150.** On a
  36.612 pt band the default drew 14.646, which is `36.612/2.5`; probes at 50 and 300 drew
  24.410 and 9.154 against 24.408 and 9.153 predicted. An empty `<c:upDownBars/>` is the
  default, byte-identical to `val="150"`.

* **The default up and down fills are #F9F9F9 and #3F3F3F**, both with a black 0.5 pt
  outline. They are **not** theme accents, which is what the brief suspected. One theme
  measured — the Office scheme, whose `lt1` is white and `dk1` black — so whether they are
  literal or derived from those two is unknown, and they are carried as literals rather
  than as a derivation nothing has tested. An explicit `c:upBars`/`c:downBars` fill wins.

* **`c:hiLowLines` is the vertical range at each category**, from the largest value there
  to the smallest — 18 down to 8 on a 0..20 axis, drawn at 25.606 and 98.123 pt, both on
  the axis to 0.001 pt — at the band centre, in black at 0.5 pt by default and in the
  file's own `a:ln` when it states one. Presence is the switch. Only well-formed data was
  probed, so "the largest and smallest of every series" and "the second and third series"
  are not separated by any observation here; the former is implemented, because it is the
  one that cannot pick the wrong pair when the series are ordered differently.

* **The legend key is the line chart's rule with the marker on it**, 19.200 pt, confirmed
  on the bottom-legend probe.

#### Surface — measured, and deferred

Six probe charts, exported and compared against our own render. **`surfaceChart` is not
drawn and should not be**, and this is the reasoning rather than an absence of effort.

What PowerPoint drew, on every one of the six:

* **A lit 3-D mesh, in perspective, including for the spelling without "3D" in it.**
  ECMA-376 calls `c:surfaceChart` a contour chart and `c:surface3DChart` a surface, and
  the obvious reading is that the first is a flat 2-D map. It is not: with no `c:view3D`
  the two spellings drew the **identical** projected 3-D surface, complete with a floor, a
  back wall, gridlines drawn in perspective and three axis label runs positioned inside
  that projection. Adding `<c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D>`
  turned the whole picture, so the *view*, not the element name, is what decides.
* **`c:wireframe val="1"` replaces the fill with a stroked mesh** and nothing else changes.
* **The surface is coloured by value band, not by series** — accent1 for 0–5, accent2 for
  5–10, accent3 for 10–15 — with each facet shaded by its orientation, so the same band
  appears in two or three different tones depending on which way the quad faces.
* **The legend is of those bands**, printed as `0-5`, `5-10`, `10-15`, which is a legend
  model no other chart type here has.

What that would take, none of which exists and none of which is shared with anything else:
a projection from `c:view3D` (`rotX`, `rotY`, `perspective`, `rAngAx`, `depthPercent`,
`heightPercent`), painter's-algorithm ordering of the quads, a lighting model to reproduce
the per-facet shading, `c:bandFmts` for the value bands, a projected axis frame with walls
and gridlines, axis labels placed in the projection, and a band legend. Each of those is
itself a fitted, measured thing; a surface drawn without the lighting or without the
hidden-surface ordering is not a rough version of the picture above, it is a different
picture that reads as a bug.

**What would close it**: the projection first, measured against the `view3d` probe, which
is the only one of the six whose camera differs and therefore the only one that constrains
the matrix. Until the projection reproduces that probe's floor and wall vertices, none of
the rest can be checked at all. The empty frame plus `chart-unsupported-type` stays, which
is the principle this file already applies everywhere else: a wrong picture is worse than
an honest gap.

#### What the new types cost the corpus, and what they bought

No corpus deck holds a bubble, an ofPie or a stock chart, so nothing scored moves and that
is the point: `authoring-integration` holds at 0.9327/0.9984, `table-test` at 0.9895/0.9984
and `real-college-template` at 0.8003/0.8753, all unchanged to four decimals despite the
accent ramp, the `_place_label` gap parameters and `_is_line` all being shared with chart
types those decks *do* hold. The probe decks, scored the same way:

| probe deck | before | after |
| --- | --- | --- |
| bubble 1 (12 charts) | 0.3640 / −0.0254 | **0.9338 / 0.9639** |
| bubble 2a (6) | 0.4053 / −0.0232 | **0.9252 / 0.9762** |
| bubble 2c (4) | 0.3605 / −0.0013 | **0.8900 / 0.9775** |
| bubble 2d (1) | 1.0000 / 1.0000 | 1.0000 / 1.0000 |
| bubble 3 (8) | 0.3677 / −0.0124 | **0.9246 / 0.9432** |
| ofPie 1 (12) | 0.6767 / −0.0299 | **0.9910 / 0.9994** |
| ofPie 2 (12) | 0.6638 / −0.0270 | **0.8949 / 0.9944** |
| stock 1 (12) | 0.0894 / 0.0098 | **0.5733 / 0.8683** |

The bubble decks sit where the scatter deck sits and for the same reason: circles on white
are a pixel of antialiasing per edge and SSIM is punishing about that. Put side by side at
1400 px they are indistinguishable.

**The two that stay low are both explained, and neither is a bubble, ofPie or stock
defect.**

* **ofPie 2 is the wrapped legend.** Its `legb` chart has six entries; PowerPoint laid
  them out as two rows of three and we lay them out as one row of six with the last name
  broken over two lines, which moves both plots 8 pt down. That is the *Where a wrapped
  legend's rows sit* gap already recorded below. Its `dlbl` chart is the label shrink
  recorded above. Every other chart on the deck is exact.
* **stock 1 is the unsolved value-axis rule**, and it adds one more observation to the
  section below rather than being a stock defect: 8..18 of data on a 145.035 pt axis comes
  back **0..20 by 2** — ten intervals at 14.504 pt — where the shipped rule gives 0..20 by
  5. Every stock-specific number on that deck reproduces: the hi-low lines land within
  0.07 pt, the bar widths within 0.005 pt, the bar tops and bottoms within 0.06 pt, the
  up/down assignment on all five categories, and the default fills exactly. What costs the
  SSIM is six missing gridlines and six missing axis labels on each of twelve charts.

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
69.538, so the error more than halves and the labels stop colliding.

**What the 19.8 pt left is — corrected.** This paragraph used to say it was the width:
that `プラットフォーム` came out 96 pt through the `<a:latin typeface="Arial"/>` the axis
names where PowerPoint laid it out "in a substituted CJK face at about 68". **The 68 was
this very formula inverted through PowerPoint's 69.538 pt inset, not a measurement.**
PowerPoint's own export draws all eight glyphs at exactly 12.000 pt — one em of the
YuGothic-Regular it embeds — so the label is 96 pt drawn and 96 pt measured, and
`デジタル`, `グローバル` and `その他` match to 0.000 pt as well. The error is this rule's
missing cap, and the export adds a clue the probes could not: **PowerPoint truncated the
label to `プラット…`**. See *The East Asian face cascade* for the implied-width numbers
and the 0.505-of-frame-height candidate they suggest.

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

`surfaceChart` and the ChartEx family warn `chart-unsupported-type` and draw an empty
frame rather than a wrong picture. The shared infrastructure — value domain, tick
selection, number formatting, gridlines, legend layout for all four `legendPos` values,
plot-area rectangle, the polar region, markers, data labels — is built, and the three types
that landed this week were mostly a matter of reusing it: the bubble is the scatter's
layout with a disc per point, the ofPie is the pie's region with two plots packed into it,
and the stock chart is the line chart with two decorations drawn over it.

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

**The area sweep adds one more datum and sharpens the contradiction rather than resolving
it.** Two of its probes carry the identical 0..10 axis and 10 pt labels and differ only in
plot height, because one of them has a bottom legend:

| plot height | PowerPoint | intervals | spacing |
| --- | --- | --- | --- |
| 145.03 pt | 0..10 by 1 | 10 | 14.503 pt |
| 120.95 pt | 0..10 by 2 | 5 | 12.095 pt |

So 14.503 pt is accepted and 12.095 pt refused, which brackets a putative threshold to
(12.095, 14.503] — and the 30.9 pt cell above refused **15.43 pt** on the same font. The
three cannot be ordered by spacing alone. Ours draws eleven labels where PowerPoint draws
six on the 120.95 pt probe, and that is the only visible difference between our render of
the area deck and PowerPoint's.

**A second rule was written against this and reverted, and the refutation is the corpus
radar.** The scatter axis-minimum probe threw up a case the *unit* rule gets wrong rather
than the density one: 120..160 of data on a 145 pt axis came back **0..180 by 20**, where
halving the power of ten gives 0..200 by 50. Stepping the unit down the 1-2-5 ladder while
the span holds fewer than about 3.5 units reproduces that **and** all five observations
`AXIS_HALVING_RATIO` was fitted to — six for six, and it brackets the threshold to
(3.2, 3.684] with 160/50 refused and 1842/500 accepted. Then `real-financial-report.pptx`
kills it: its chart5 has 65..100 of data and PowerPoint's own export draws **two** rings,
at radii 22.78 and 45.56, which is 0..100 by 50 — a ratio of exactly **2.0 accepted**
where the scatter refused 3.2.

What separates them is the **axis length**, which puts both observations back in this
section rather than in unit selection:

| chart | span | unit | intervals | axis | spacing |
| --- | --- | --- | --- | --- | --- |
| corpus radar | 100 | 50 | 2 | 45.56 pt | 22.78 pt |
| scatter y | 160 | 20 | 9 | 145.03 pt | 16.11 pt |

A target *band* of roughly 16 to 24 pt of spacing fits those two and every cell of the
six-cell table above — the radar's alternatives are 9.1 pt and 45.6 pt, both outside it,
and the scatter's are 8.1 pt and 36.3 pt — and then it dies on the same stacked probe that
killed the last candidate, which accepts 14.5 pt. So the ladder is **not shipped**, the
halving stays, and `test_the_axis_ladder_that_the_corpus_radar_refuted` pins what we
actually draw so the divergence is recorded rather than latent. Whoever picks this up now
has two long-axis observations to add to the six short-axis ones, and the discriminating
set is the 30.9 pt cell, the stacked probe and the corpus radar.

**The stock sweep adds a third long-axis observation, and it agrees with the stacked probe
rather than with anything shipped.** Its twelve charts all carry 8..18 of data on a
145.035 pt axis with 10 pt labels, and PowerPoint drew **0..20 by 2** on every one of them
— ten intervals at 14.504 pt, the same spacing the stacked probe accepted. The shipped
rule gives 0..20 by 5, four intervals at 36.26 pt, so we draw five gridlines where
PowerPoint draws eleven. That is the whole of the stock deck's SSIM shortfall and it is
not a stock behaviour: the same data on a line chart would do the same thing.

It also sharpens what the missing variable is *not*. This case and the stacked probe share
an axis length (145 pt), a font (10 pt) and an accepted spacing (14.5 pt); the area
probe's 120.95 pt cell refused 12.095 pt and the six-cell table's 30.9 pt cell refused
15.43 pt. So the accepted set now spans 14.50 to 27.3 pt and the refused set 12.10 to
15.43 pt, and they **overlap** — no threshold on spacing alone, at any axis length, orders
all nine observations. Whatever the rule is, it is not a spacing threshold.

#### Not done for the ten types that draw

Each of these is known-missing rather than merely absent:

* **Data-label wrapping**, and **`c:separator` on an area chart**. PowerPoint wraps a long
  category name onto two lines inside a multi-part label; we draw it on one. `c:separator`,
  `c:leaderLines` and a data label's own `c:layout` are read or ignored but never drawn.
  The area probe adds a measurement to that: a label showing category name *and* value
  came back joined with `"; "` on **one** line and wrapped mid-word when it did not fit
  (`Rendere` / `r; 5`), where a bar chart's multi-part label stacks its parts on separate
  lines. Ours stacks for both. The bar behaviour is measured and shipped; the area one is
  measured and **not** — changing the join would need the wrap that goes with it, and this
  is the same known gap either way.
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
* ~~**`c:smooth`'s tension.**~~ — done and measured; see *Two constants the new probes
  corrected* above. It is the plain Catmull-Rom 1/6, and the terminal control points are a
  *third* of their own chord rather than a sixth, which is the reflected phantom point and
  not the duplicated one this used to draw. Thirteen ordinates reproduced to 0.08 pt.
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
  move the category axis but have only been measured at zero — except on a scatter, where
  the negative-x and negative-y probes measure both crossings at the other axis' own zero.
  `c:crossBetween="midCat"` **is** measured now, on an area chart; the same placement is
  applied to a line chart stating it, which no probe has exercised.
* **`dispBlanksAs="span"`** is treated as `gap`, which is right for a bar chart and will
  not be for a line one. It is unmeasured on an area chart and on a scatter, where the
  line chart's reading is reused.
* ~~**A scatter's `c:bubbleSize`.**~~ — done and measured; see *Bubble, measured*. A
  scatter or bubble series' `c:trendline` and `c:errBars` are still neither read nor drawn,
  and neither is `c:dropLines` on a line or stock chart.
* **`c:crossBetween="midCat"` on a bar chart.** Excel writes it for the category axis'
  "Axis position: on tick marks" checkbox, so an ordinary column chart carries it, and
  `_draw_bars` still lays its bars into bands. Ours therefore keeps a `midCat` bar chart's
  labels in the bands with its bars — the one reading that cannot contradict itself — and
  what PowerPoint actually draws for that file is **not measured**. Only `areaChart` and
  `lineChart` put their marks on the ticks.
* **A scatter data label at `b` is 0.9 pt low**, the one loose number in the five
  placements. It is the slack this file records elsewhere: PowerPoint's line box runs about
  a point taller than our metrics give, so a placement hung off the *ascent* inherits all
  of the difference where one hung off the descent inherits none. One constant is used for
  both rather than two fitted ones.
* **A horizontal legend's entry pitch is 1.3 to 2.2 pt out** on both new types, and the
  measurement says so: the area probe's two bottom entries are 40.375 pt apart and the
  scatter's 55.127 pt, which is 7.58 and 8.97 pt of gap after the entry's own width where
  `LEGEND_ENTRY_GAP_EM` gives 5.0. No single number produces both, so the constant fitted
  to the bar legends is left alone and the residual recorded.
* The chart frame's rounded corners (`c:roundedCorners`) and effects.
* **Three things the new types measured and did not ship**, each with one observation
  behind it and named where it was measured: a bubble label PowerPoint pushes back inside
  the frame rather than letting it overflow; the shrink an ofPie applies to both its plots
  when it has data labels; and the dark disc PowerPoint draws where an empty second plot
  would be. See *Bubble, measured* and *Of-pie, measured*.
* **The ChartEx family** — treemap, sunburst, histogram, box-and-whisker, waterfall,
  funnel and map, all new in Office 2016. They are a `cx:chartSpace` part in a different
  namespace with a different data model, sharing no markup with `c:chartSpace`, so they are
  not a missing case in the chart reader but a second format. They draw an empty positioned
  frame and warn `chart-unsupported-type`. **They used to warn the wrong thing**: their
  graphic frame's first child is also called `chart`, so `parse/shapes` fell through to the
  ordinary chart path and reported "names no chart part" — true of the `c:` relationship
  and false about the file. `GRAPHIC_DATA_CHARTEX` tells them apart now. The picture was
  always right; only the diagnosis was wrong, and this is exactly the class of defect the
  review note below is about.

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

- ~~**Embedded fonts**~~ — **done, in Phase 6.** This entry once read "extract, undo the
  trivial obfuscation, pass via the existing `font_files` parameter". Both halves were
  wrong: the payload is EOT and in practice MicroType-Express compressed, and `font_files`
  feeds the rasteriser only, so it would have drawn the right glyphs while still measuring
  the wrong widths. Kept visible rather than deleted, because the happy path is what made
  it look cheap — and because the *second* estimate, "this needs a C library", was wrong
  in the other direction.
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

### 5.4 Group coordinate scaling does not scale text — **measured, and fixed**

A group's `a:ext`/`a:chExt` ratio maps its children's authored coordinate space onto its
on-slide box. We implemented that as an SVG `scale()` around the whole subtree, which
scales the glyphs with everything else. Google Slides exports make that catastrophic: they
author a slide in a small private space and stretch it to fit, so a 24 pt subtitle in a
group scaling 3.7969x drew at 91 pt, wrapped onto three lines and fell off the bottom of
the slide.

**Both readings were defensible before the measurement.** If PowerPoint scaled text, our
91 pt was right and the template was simply mis-authored. It does not. A probe deck of 21
groups was exported by PowerPoint 16.x and the drawn text read back out of the PDF with
`pypdfium2` — per-character font size is written as `1` with the size folded into the text
matrix, so the sizes below are cap heights and advance widths, not `/FontSize`:

| probe | group scale | drawn |
| --- | --- | --- |
| ungrouped control, 18 pt Arial | — | 12.89 pt cap height |
| uniform | 4.0 | 12.89 |
| nested, 2x inside 3x | 6.0 | 12.89 |
| shrinking group | 0.25 | 12.89 |
| non-uniform, x only | 4.0 / 1.0 | 12.89 tall, 58.67 pt wide — the control's width exactly |

The non-uniform probe is the one that settles it: scaled text would have drawn four times
as wide. **PowerPoint never scales text with a group, in either axis, at any depth.**

Everything else in the text frame is absolute too, each measured on the same deck:
`bodyPr@lIns` of 91440 EMU drew a 7.20 pt indent at scale 1 and 7.20 pt at scale 4;
`lnSpc` `spcPts` 3000 gave a 30.00 pt baseline pitch at both; `spcBef` `spcPts` 1200 gave a
32.88 pt paragraph pitch at both; `normAutofit` `fontScale` 50% gave a 6.44 pt cap height
at both; the first baseline sat 18.1 pt below the frame top at both. Only the *frame*
scales — a long paragraph broke into identical lines inside a 4x group and in an ungrouped
box of the same on-slide width.

So the fix is not "divide the font size by the scale", which would be wrong for a
non-uniform group and would wrap at the child-space width. The text is laid out in the
frame grown by the accumulated scale, at the authored size, and the result wrapped in the
transform that cancels the scale. `RenderContext.group_scale` carries the factor;
`tests/test_group_text_scale.py` holds the cases.

**That limitation is closed; see §5.5.** It was recorded here as "a rotated group with a
non-uniform scale composes to a shear, and a pair of per-axis factors cannot express
that", alongside the geometry half — "a rotated shape in a 4:1 group draws a
differently-skewed parallelogram than PowerPoint's". Both halves rested on an assumption
that turned out to be wrong: **PowerPoint never composes a shear at all.**

**A second defect surfaced on the same slide.** `_render_custom_path` emitted the path's
scale through the coordinate formatter, which rounds to three decimals. A custom path
authored in EMU — which is what every Google Slides export writes, `<a:path w="2387010"
h="161597">` on a shape 250 px wide — maps onto its shape by 1.05e-4, and three decimals
make that `scale(0, 0)`. **Every such shape drew as nothing.** It was invisible in the
corpus because the committed fixtures author custom paths in small path units, and in the
suite because the unit test used `w="100"`. On the sales template this was worth more than
the text fix: mean SSIM against PowerPoint's export went 0.7912 → 0.8015 with the group
rule alone and 0.7912 → 0.9101 with both.

### 5.5 A group's scale is not a matrix — **measured, and fixed**

§5.4 left one case open and framed it as "the composite is a shear and we cannot express
it". SVG *can* express a shear, with `matrix()`, so that framing made the work sound like
a notation problem. It was not. **PowerPoint does not draw a shear. There is no shear to
express.**

Three probe decks — 82 probes over 82 slides, built by `tools/make_group_shear_probe.py`
and exported by PowerPoint 16.x — were read back not by rasterising but by walking the
PDF's path objects for the drawn vertices (`tools/read_group_shear_probe.py`). The
discriminator is a square rotated 45° inside a 4:1 group. Scale-after-rotate gives a
rhombus with diagonals 178.2 and 44.5 pt and corners at 28° and 152°. PowerPoint drew a
**31.496 × 125.984 pt rectangle with an interior angle of 90.000°**, and so did all 43
sweep probes at every angle, ratio, flip and nesting order tried:

| authored | drawn | |
| --- | --- | --- |
| square, 4:1 group, rot 0 | 125.984 × 31.496, corner 90.000° | the control |
| square, 4:1 group, rot 30 | 125.984 × 31.496 at 30° | the 4 is on the width |
| square, 4:1 group, rot 44.9 | 125.984 × 31.496 | still the width |
| square, 4:1 group, rot 45 | 31.496 × 125.984 | the two factors change places |
| square, 4:1 group, rot 60 | 31.496 × 125.984, long side at 150° | an authored 60° draws at −30° |
| square, 4:1 group, rot 90 | 125.984 × 31.496, axis-aligned | identical to rot 0 |

So the rule is: **the group hands each of its two factors to one of the child's own axes,
choosing by which slide axis that axis currently lies nearer to, and rotates the grown box
afterwards.** The centre is still mapped by the plain scale — it landed on
`origin + S·(centre − chOff)` to the hundredth of a point on all 43. The rest of the deck
pins the rule down:

* **The crossing is at 45° and is a switch, not a blend.** 44.9 and 45.0 both drew exactly
  1× and 4× the authored side, nothing in between.
* **It does not move with the ratio.** 2:1 and 10:1 both flipped between 44 and 45.
* **It repeats every 90°, signed.** 134 swapped and 135 did not; 180 behaved as 0, 225 as
  45, 270 as 90, 315 as 135, 405 as 45, −30 as 150, −60 as 120.
* **Neither factor need be 1**, and the child need not be square: a 4:2 group put 4 and 2
  on a 400000 × 200000 child's own axes and swapped them past 45°.
* **Nesting composes.** A rotated nested *group* is treated exactly like a rotated shape,
  and a 2:1 inside a 4:1 scaled a rotated child by 8 — 251.968 pt against the authored
  31.496 — so the assignment sees the product.

**A reading the measurements refuted.** Since |sin| and |cos| cross at exactly 45°, the
obvious rule is "swap when |sin| > |cos|", and flips would be irrelevant because that test
takes magnitudes. Twenty-six flipped probes at 30° and 60° agree with it. The ties do not:
at exactly 45° an unflipped child swaps and one carrying `flipH` *or* `flipV` does not, and
at exactly 135° it is the other way round. One flip negates the angle the test is made on;
two flips cancel. The rule is therefore stated on a signed angle — `swaps_group_axes` in
`render/svg.py`.

**What text does inside it.** Glyphs never skew, and they never take the geometry's
effective angle either. Every character of an 18 pt Arial run in a 4:1, 1:4 or 2:1 group
came out with a text matrix of exactly `18 · R(θ)` for the **authored** θ — byte-identical
to the ungrouped control's matrix, no skew term, size untouched. Only the frame moves: the
first baseline of a run in a 4:1 group at 45° was predicted from the *swapped*
31.496 × 62.992 pt frame to within 0.07 pt, where the unswapped frame predicts a point
50 pt away. So text needs no counter-transform once the scale is folded into the box;
`RenderContext.group_scale` now carries the uniform case only.

**The fix.** A uniform scale commutes with rotation and stays an SVG `scale()` around the
subtree, byte-identical to before. A non-uniform one is folded into each child's own box.
The exception is a table, whose column widths and row heights are authored in EMU and do
not follow its frame; it keeps the wrapper, and with it the old approximation.

Uniformity is tested with a tolerance and a real deck says why: Google Slides rounds one
scale into two EMU pairs that no longer divide to the same number — 3.796875 across against
3.7968797 down in the meal planning template, and 0.75 against 0.7499996 — so four groups
that are uniform by intent would otherwise fold on float noise. 1e-4 is three orders of
magnitude above that and far below any deliberate stretch.

**Before and after, on the probe decks:** 14/72 → **72/72** drawn quads matching
PowerPoint's vertices to 0.05 pt, and 4/10 → **9/10** text baselines within a point at the
same angle and size. By mean SSIM: `group-shear` 0.7971 → 0.9968, `group-sweep`
0.6435 → 0.9997, `group-tie` 0.5115 → 0.9423. The scored corpus did not move at all
(`authoring-integration` 0.9327, `table-test` 0.9895, `real-college-template` 0.8003), nor
did either template deck in `scratch/`, which render byte-identically.

Only one deck anywhere changed: `real-basic-theme`, whose layouts hold a 0.7006 × 1.8185
group with children at −90°. Its pixels did not move, and the reason is worth recording —
at an exact multiple of 90° the scale and the rotation commute, `S·R(90) = R(90)·D`, so the
old code was accidentally right there. That is why the corpus never caught this.

**Found and not fixed: a group's scale does not reach the pen either.** A 6 pt outline drew
6.00 pt in every case measured — ungrouped, in a 4:1 group, in a *uniform* 4:4 group, and
on a rotated child. Line width is absolute the way font size is. Folding a non-uniform
group now gets this right as a side effect, but the uniform path still multiplies the
stroke by the scale: `group-tie` slide 9 scores 0.4702 against 0.9862+ for every other
slide on that deck, drawing a 6 pt pen at 24 pt. Fixing it means folding uniform groups
too, which changes every grouped fixture and contradicts four of the nine cases in
`tests/test_group_text_scale.py` that deliberately pin the `scale()` mechanism, so it is a
separate piece of work.

**Found and not fixed: we mirror text inside a flipped shape and PowerPoint does not.** The
`flipH` text probe drew unmirrored at its authored +45°; we emit `scale(-1, 1)` around the
`<text>` and draw it at −135°. This is independent of groups — any flipped text box does it
— and predates this work.

---

## Phase 6 — Embedded fonts — **done**

**Estimated L, with "a hard dependency that could make it XL". It came in under that, and
the reason is the second finding below.** The two cost drivers this phase was planned
around — "is it really MTX?" and "MTX needs a C library" — were an inference and a
prediction. Measuring both changed the shape of the work.

### 1. The payloads really are MTX, and that is now measured rather than inferred

The plan said: decode one payload before designing anything, because the whole estimate
rested on reading `Flags = 0x4` as `TTEMBED_TTCOMPRESSED` and nothing had actually
decompressed one. Done first. libEOT's `eot2ttf`, built from source as an oracle outside
this repository, decodes **all 21** `.fntdata` parts across the two local template decks
to valid TrueType files. So:

```
Version       0x00020002   (EOT 2.2)        on all 21
Flags         0x00000004   (TTCOMPRESSED)   on all 21 — no SUBSET, no XORENCRYPTDATA
fsType        0x0000       (installable)    on all 21
EOTSize       == the part length, to the byte, on all 21
```

The inference held. Two things that were not expected:

* **The payloads are not subsetted.** `TTEMBED_SUBSET` is clear, and the decoded Arimo
  carries all 3237 glyphs with a 3010-entry cmap. PowerPoint embedded the whole face.
  That is what makes measuring from the extracted file worth doing at all — a subset would
  only cover the characters already on the slides.
* **The part names lie.** `ppt/fonts/MerriweatherSansLight-bold.fntdata` decodes to a font
  whose `name` table says *Merriweather Sans / Regular* at `usWeightClass` 400. PowerPoint
  had no Light Bold cut and substituted one, and it still draws bold "Merriweather Sans
  Light" runs with it. The deck's slot assignment, not the file's own name table, is what
  governs. See §5.

### 2. A C dependency was not needed, and the extra had nothing to point at

The plan was `pptx2svg[eot]`, shaped like `[metafile]`, because LibreOffice 25.8 links
libEOT for exactly this. Two measurements killed that:

* **There is no EOT decoder on PyPI.** Not `libeot`, not a binding, not anything. The
  "optional extra" would have been a package this project first had to write and publish.
  That was never checked when the extra was proposed.
* **The decode-side subset of libEOT is about 700 lines of Python, and it is fast enough.**
  Per face: 0.15 s (Inclusive Sans, 23 kB payload) to 1.06 s (Arimo Bold Italic, 164 kB).
  The nine faces of the largest local deck total 5.9 s.

So `fonts/mtx.py` is pure standard library and lives in the core, and a bare
`pip install pptx2svg` reads embedded fonts. Only the decompressor is ported (`Decode`,
`DecodeLength`, `DecodeDistance2`, `InitializeModel`, AHUFF's read side, BITIO's read side,
`parseCTF`, `SFNTContainer`); libEOT's compressor half has no caller here.

**Correctness is pinned against libEOT, not against ourselves.** All 21 local payloads
decode **byte-identically** to `eot2ttf`'s output. Two further independent checks, because
the warning this file gives elsewhere — our own writer agreeing with our own reader proves
nothing — applies to the sfnt reader as much as to the decoder:

* The metrics table `fonts/sfnt.py` builds from the *embedded static* Arimo reproduces
  `METRICS["Arimo"]` — generated by fontTools from the *bundled variable* Arimo —
  character for character, both cuts, 349 of 349, plus unitsPerEm, ascender, descender,
  `default_width` and `cjk_width`. Different file, different tool chain, same numbers.
* `tests/fixtures/real-basic-theme.pptx` embeds eight real PowerPoint-written MTX payloads
  (see §3), so CI decodes PowerPoint's own output on every run.

### 2a. A bound the reference implementation does not have

Found by mutating bytes inside the font data of the real payloads: 19 of 1680 mutations
made an LZ copy item run past the declared output length, which escaped as an `IndexError`
rather than as this module's own error. In libEOT the same input is worse than an
exception -- `Decode` writes the whole copy item into a fixed-size buffer and only
*afterwards* asserts `pos == out_len`, so a corrupt stream overflows the heap before the
check runs. The port now refuses the item instead, and all 21 real payloads still decode
byte for byte with the bound in place.

A deck is untrusted input, so the surrounding code was swept the same way. 2520 truncation
and bit-flip cases through `decode_eot` and `read_sfnt`, 1680 through the compressed body,
960 through the CTF rebuild: no unexpected exception type, no case slower than eight
seconds. Two `cmap` subtable formats were also bounded — a format 12 header can declare
four billion groups and a format 4 segment can span 65,536 code points, so a kilobyte of
malicious `cmap` could otherwise ask for hours of work. The 120-mutation regression test
is pinned to a seed that reaches the copy-item bound three times, so it cannot quietly
stop covering it.

### 3. Correction: a corpus deck *does* embed fonts

`tests/fixtures/real-basic-theme.pptx` carries
`ppt/fonts/{Lato,Raleway}-{regular,bold,italic,boldItalic}.fntdata` — eight EOT 2.2
payloads, MTX compressed, written by PowerPoint. The instruction to treat a fidelity move
as "a finding, not a rebaseline" was sound, but its premise ("no corpus deck embeds fonts")
was not.

Two consequences, both worth keeping:

* **It made the feature testable in CI** without adding a byte to the repository.
* **It cost 6× on the test suite before it was cached.** Every test deriving a deck from
  that fixture decoded eight faces: 21 s → 126 s. A bounded content-digest cache in
  `fonts/embedded.py` brings it back to 25 s, and keying it by payload digest rather than
  by part path is what makes a *derived* deck hit.

Fidelity did **not** move: `authoring-integration` 0.9327, `table-test` 0.9895 and
`real-college-template` 0.8003 are all unchanged. `real-basic-theme` is skipped by the
harness for an unrelated reason — PowerPoint drew MS-Gothic where the deck asks for
ＭＳ Ｐゴシック — so the one corpus deck whose rendering *did* change is the one the corpus
cannot score. That gap is why the template decks were scored by hand instead; see §6.

### 4. The `fsType` policy, and why it is that one

Refuse a face **only** when its permission field says restricted-licence embedding and says
nothing else. That is LibreOffice's rule, adopted deliberately rather than invented:

```cpp
// EmbeddedFontsHelper::sufficientTTFRights, vcl/source/gdi/embeddedfontshelper.cxx
case FontRights::ViewingAllowed:
    return copyright == 0 || ( copyright & 0x0e ) != 0x02;
```

Rendering a deck to SVG or PNG is a viewing use: the output is a fixed image, it is not
editable, and it carries no font file — an SVG names fonts and embeds none. So `0x0000`
(installable), `0x0004` (preview & print) and `0x0008` (editable) all permit it, in any
combination, and `0x0002` alone does not.

Two bits beyond LibreOffice's check:

* **`0x0200`, bitmap embedding only — refused.** We draw outlines, which is exactly what
  the bit forbids. Refusing where LibreOffice permits is the direction to err in.
* **`0x0100`, no subsetting — permitted, and written down so nobody adds a check for it.**
  Nothing here subsets: the extracted face is written out whole.

**The gate is checked twice, and the second time is the point.** The EOT header carries its
own `fsType` copy, written by whatever produced the EOT; the decoded font's `OS/2` table is
the foundry's own statement. The header is checked first, purely so a refusal does not cost
a second of decompression, and the decoded table is checked after. Both must permit. There
is no bypass parameter, and `--no-embedded-fonts` only turns the whole feature off — it
cannot be used to ignore a restriction.

A refused face warns `font-embedded-restricted` naming the restriction and falls back to
substitution; an undecodable one warns `font-embedded-undecodable`. Neither is fatal, which
is the same degradation contract `metafile-rasterizer-missing` implements.

### 5. Relabelling, which is not cosmetic

`<p:embeddedFont><p:font typeface="X"/><p:bold r:id="Y"/>` asserts *Y is the bold cut of X
for this deck*, and the file behind Y need not agree — in the local corpus it frequently
does not. Left alone, such a file is **unaddressable**: the SVG asks for
`font-family: 'Merriweather Sans Light'; font-weight: bold`, and the rasteriser's font
database has that file under "Merriweather Sans" at weight 400. The face would be handed
over and then silently passed over, and a *similar* face would draw instead — the exact
silent-substitution failure this subsystem exists to prevent, and one that would have
looked fine in a screenshot.

So `fonts/sfnt.py:relabel` rewrites `name` 1/2/4/6 to the deck's typeface and the slot's
style, drops 16/17 (whose only purpose is to override 1/2, which are now correct), and sets
`OS/2.usWeightClass`, `OS/2.fsSelection` and `head.macStyle` to the slot. Copyright (0),
trademark (7), licence (13) and licence URL (14) survive byte for byte, which matters for a
file this library writes to disk. The measured widths come from the same file, so
measure-equals-draw holds whatever the label says.

Relabelling is deferred until a file is actually written, so an SVG-only render never pays
for it.

### 6. Measurement, and what it was worth

The hard requirement was that an embedded face reach measurement, not just the rasteriser.
It does, by a route that keeps the core standard-library-only: `fonts/sfnt.py` reads
`head`/`hhea`/`maxp`/`OS/2`/`hmtx`/`cmap` without fontTools and builds a `FontMetrics`,
which `DefaultTextMeasurer(extra_metrics=...)` consults ahead of the static tables.
Injecting a table rather than adding a measurer means every existing rule — bold cuts,
synthetic emboldening, East Asian advances, the first-baseline formula — applies to an
embedded face unchanged instead of being reimplemented beside it. Unlike the generated
tables, which store a ~450-character sample because a human reads them, the embedded table
stores the whole cmap, so nothing a deck can contain falls back to a guess.

`"Hamburgefonstiv 12345"` at 18 pt, guessed against measured:

| face | guessed | from the deck's own file | error removed |
| --- | ---: | ---: | ---: |
| Anton | 280.800 px | 225.141 px | **19.8 %** |
| Literata | 280.800 px | 270.360 px | 3.7 % |
| Inclusive Sans | 280.800 px | 270.744 px | 3.6 % |
| Merriweather Sans Light | 280.800 px | 273.408 px | 2.6 % |
| Arimo | 254.824 px | 254.824 px | 0 % — already bundled, and the two agree exactly |

Anton is the case that shows the guess is not merely imprecise: it is a condensed display
face, so the 0.6 em per-character fallback was laying out every headline on that deck a
fifth too wide.

**Scored against PowerPoint's own PDF export** of the two template decks, six slides each,
rendered with the bundle plus the extracted faces (SSIM, foreground pixels):

| | slide 1 | 2 | 3 | 4 | 5 | 6 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| nutrition, before | 0.7012 | 0.7398 | 0.8424 | 0.8199 | 0.8257 | 0.6942 |
| nutrition, after | **0.8190** | **0.8210** | **0.8879** | **0.9426** | **0.8731** | **0.7579** |
| sales, before | 0.8698 | 0.9267 | 0.8440 | 0.8710 | 0.7883 | 0.6468 |
| sales, after | **0.9124** | **0.9786** | **0.8754** | **0.9202** | **0.8363** | **0.6547** |

Twelve of twelve improved; mean +0.058, largest +0.123. Nothing regressed. That is the
measurement the corpus could not supply, and it is against PowerPoint rather than against
this library's own output.

### 7. An embedded face beats a bundled one of the same name

Not obvious, and decided on a measurement. `real-basic-theme.pptx` embeds Raleway 4.026 and
Lato 1.104; the bundle ships later releases of both. Over the characters the two tables
share, **338 of 340 Raleway advance widths differ** and 205 of 236 Lato ones do, and
`"Hamburgefonstiv 12345"` measures 251.184 px from the bundled Raleway against 261.000 px
from the deck's — **3.9 % apart, the same order as the 4.5 % Caladea error this file records as a
bug.** The author laid the deck out with the file inside it. So the embedded face wins, and
`fonts --check` grades it `exact` with "drawn with the face the deck embedded", because the
layout was measured from the very file the rasteriser is handed.

*Labelled as an inference:* Microsoft documents embedded fonts as a fallback for machines
that lack the face, which suggests PowerPoint itself prefers an *installed* copy over an
embedded one. That was not tested — `real-basic-theme` is unscoreable, and its Raleway text
does not reach the exported PDF's simple-font `/Widths` arrays, so the cheap check was not
available. It does not change the decision: this library has no installed fonts to prefer,
so the choice is between the deck's Raleway and the bundle's, and only one of them is the
file the author used.

### Done when — all met

- ✅ A real `.fntdata` payload decoded to a usable font file; the MTX question answered by
  decoding 21 of them, byte-identically to libEOT.
- ✅ `<p:embeddedFontLst>` parsed into the source model, keyed by family and face
  (`SourceEmbeddedFont`).
- ✅ An embedded face is **both measured and drawn** from the extracted file.
- ✅ `fsType` checked twice, and a restricted font refused with a warning naming the
  restriction.
- ✅ A deck whose embedded font cannot be decoded still renders, and says why.
- ✅ **The two template decks in `scratch/` render without a single "no substitute known"
  warning** — the observable this phase existed to move. Before: `Anton` and `Literata` on
  one, `Inclusive Sans` and `Merriweather Sans Light` on the other. After: none, on either.

  | deck | sha256 |
  | --- | --- |
  | `nutrition_templates-Meal Planning Slides.pptx` | `0f50e1c0f786abbaa459512cf602cc5992115741f275182265455103e013ef81` |
  | `sales_templates-IT Software Sales Proposal Slides.pptx` | `7b95c1f168d6102c8296aed23d7b58f26cdb3f4d8c5414d5197fad414c004204` |

  Both are third-party stock templates and **not committed**, on the same footing as
  `real-college-template.pptx` — see `tests/fixtures/README.md` and the `college_template`
  fixture in `tests/conftest.py` for the pattern a test should follow, which is to skip
  where the file is absent rather than to fail.

  Between them they embed Anton, Arimo, Literata, Merriweather Sans, Merriweather Sans
  Light and Inclusive Sans, and both carry EOT flags `0x00000004`. **Their provenance is
  not recorded**: they arrived in `scratch/` during the investigation and no source URL or
  download date was captured, which is a gap by the standard
  `tests/fixtures/README.md` sets for every other outside input. Whoever picks this phase
  up should either recover the provenance or substitute two template decks whose source is
  known, rather than treating the hashes above as sufficient on their own.

- **A synthetic deck exercising the uncompressed path** is still wanted. Nothing in
  `scratch/` covers non-MTX, so the cheap half of the format ships untested. It is
  additional to the real decks above, never a replacement: a fixture we write
  ourselves tests our EOT writer against our EOT reader, and those two can agree
  perfectly while both disagree with PowerPoint. It pins the branch; it proves
  nothing about the format.

### Left undone

* **Charts still measure from the static tables.** `resolve/chart.py` calls `font_box` and
  `text_width`, which reach `metrics_for` directly rather than through the measurer, and
  chart resolution runs *before* the embedded fonts are decoded — the decode is deferred
  until after resolution precisely so it can be limited to the families the resolved slides
  ask for. A chart set in an embedded face therefore still gets guessed widths. Neither
  template deck has one, so this is unobserved rather than known-broken. Fixing it means
  either decoding before resolution, and paying for families no slide uses, or routing
  chart text measurement through the measurer, which it should do anyway.
* **`--font-dir` still feeds the rasteriser only.** The gap recorded under Fonts ▸ Left
  undone is untouched. What changed is that the plumbing it wants now exists:
  `DefaultTextMeasurer(extra_metrics=...)` takes a family→`FontMetrics` map, and
  `fonts/sfnt.py` builds one from a font file with no fontTools dependency. Teaching
  `font_dirs` to feed measurement is now a small change rather than a design question, and
  the CLI could gain a measurer for nearly free. Deliberately not done here: it changes
  existing behaviour for every caller of a documented parameter, which is its own decision.
* **Italic is measured from the upright cut**, as everywhere else in this library. The
  italic file *is* extracted and drawn with; only the metrics table is shared, because
  `FontMetrics` has no italic column. Divergence is ≤2.5 % across the Office substitutes.
  Adding one is cheaper than it was, since the italic face is already parsed.
* **`hdmx` and `VDMX` are dropped**, following libEOT: they are keyed to the original glyph
  set and are not rebuilt, and carrying them through unchanged would let a rasteriser that
  trusts `hdmx` disagree with `hmtx`. Neither affects the advance widths measured here.
* **A TrueType collection inside an EOT is refused.** No `.fntdata` on hand contains one and
  the EOT specification does not describe one, so `read_sfnt` raises rather than guess.
* **The decode cache is per process and bounded at 32 faces.** A long-lived service
  rendering many embedded-font decks will re-decode. A disk cache would fix it and is not
  obviously wanted.

---

## The clone landscape — survey, measured

Phase 6 ends by deferring one question: "finding which have open equivalents is a real and
separate piece of work." This is that work. Nothing here is a plan; it is a survey with a
shortlist attached, and the decision about what the bundle ships is not made here.

It is done the way the Caladea finding was done — by reading the faces Office installed on
this machine and comparing advance widths character by character — because that is the only
method this project has any reason to trust. **A claim of metric compatibility with no
measurement behind it is exactly what this project does not ship.** Every row below either
carries a measurement or says plainly that it does not. No Microsoft font file was copied
anywhere; `/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts` was read in
place, and candidate clones were downloaded outside the repository, measured, and left
there. Measured advance widths are facts and are recorded; font files are not.

### The inventory, and why two different counts are both right

PowerPoint 16.106 ships its fonts in `Contents/Resources/DFonts`: **280 files, 322 faces**.
That is not an estimate — the app carries its own manifest at
`Contents/Resources/applicationfontmetadata.json`, whose header reads `"fc": 322`, and
enumerating the files with fontTools reproduces it exactly. **An earlier note on this page
put the face count at 388; that number is wrong and 322 replaces it.** The family counts in
it survive scrutiny: 238 distinct `name` id-1 strings across all languages, collapsing to
**186** once the 52 localised CJK aliases are dropped.

The manifest also answers a question the id-1 count cannot. Microsoft's own `fa` field —
the GDI-collapsed family, which is what the PowerPoint font menu shows — yields **155
families**. The two numbers differ because GDI folds weight and width variants into their
parent: the menu has one "Arial" holding Arial, Arial Black and Arial Narrow, and one
"Aptos" holding Light through Black.

**186 is the number that matters here**, because `text/fontmap.py` is keyed on the string a
deck writes into `a:latin/@typeface`, and a deck writes `Arial Narrow`, not `Arial`. The
bundle's eight families answer 16 of those 186 names. The other 170 reach the 0.6 em
per-character guess.

### What has to be true for a row to say "metric-compatible"

Two things, and conflating them is how Caladea got in.

1. **Same advance widths**, so line breaks, autofit and centring land where PowerPoint put
   them. This is measurable and is measured below.
2. **Nothing else.** Shape, colour, hinting and x-height are all irrelevant to layout. A
   face can look nothing like the original and still be a perfect substitute for our
   purposes, and a face can be a scholarly revival of the very same design and still
   reflow every line.

Two measurement conventions are used throughout. Differences of one design unit are
**quantisation, not design**: a Microsoft face is 2048 upem and most clones are 1000 or
2048, so a 0.05 % discrepancy means the clone was drawn to the same number and rounded
differently. And ASCII is reported in two slices — all 95 printable characters, and the 74
that are letters, digits and ordinary punctuation. The gap between those two slices is
almost always `#`, `|`, brackets and the math operators, which real slide text uses far
less than the 95-character count implies.

### Measured: families that do have a true metric clone

Each "differ" column counts characters differing by more than one design unit. "Text"
is the 74-character letters-digits-punctuation slice.

| Family | Clone | Kind | Licence | Evidence | Size |
| --- | --- | --- | --- | --- | --- |
| Arial Narrow | Nimbus Sans Narrow (URW) | **metric** | AGPL-3.0 + PS/PDF exception | 0/95 ASCII against Arial Narrow regular, italic and bold-italic | 617 kB /4 |
| Arial Narrow | TeX Gyre Heros Cn | **metric** | GUST (LPPL 1.3c) | 0/95 beyond rounding; max deviation 0.14 %, all of it 1000-upem quantisation | 467 kB /4 |
| Book Antiqua | URW P052 | **metric** | AGPL + exception | 0/74 text; 1/95 ASCII (`#`, 0.606 → 0.500) | 882 kB /4 |
| Palatino Linotype | URW P052 | **metric on text** | AGPL + exception | 0/74 text; 10/95 ASCII — Microsoft regularised `+ / < = > ^ \| ~` to 0.5 em, URW kept Palatino's | (same file) |
| Century Schoolbook | URW C059 | **metric** | AGPL + exception | 0/74 text; the 3 ASCII "differences" are one unit in 2048 | 838 kB /4 |
| Century | URW C059 | **metric** | AGPL + exception | Century and Century Schoolbook have identical ASCII advances; same result | (same file) |
| Century Gothic | URW Gothic | **metric** | AGPL + exception | 0/74 text; 2/95 ASCII (`#`, `^`) | 601 kB /4 |
| Bookman Old Style | URW Bookman Light/Demi | **metric** | AGPL + exception | 1/74 text (`Q`, +2.5 %); `#` also differs | 803 kB /4 |
| Monotype Corsiva | URW Z003 (Zapf Chancery) | **metric** | AGPL + exception | 0/74 text; 1/95 (`#`). Surprising, and checked twice | 215 kB /1 |
| Symbol | **Symbol Neu** (croscore) | **metric** | Apache-2.0 | **0/188 by legacy code and 0/189 by glyph name, mean 1.0000** | 69 kB /1 |
| Symbol | URW Standard Symbols PS | **metric** | AGPL + exception | 0/188 by code *and* by glyph name, mean 1.0001 — the same result under a worse licence | 24 kB /1 |
| Monotype Sorts | URW D050000L (Dingbats) | **near-metric** | AGPL + exception | 9/202 differ, mean 0.9938 | 37 kB /1 |
| Comic Sans MS | Comic Relief | **metric** | OFL-1.1 | **0/95 ASCII, regular and bold.** The cleanest result in the survey | 171 kB /2 |
| MS Gothic, MS Mincho, SimSun, NSimSun, SimHei, KaiTi, FangSong, MingLiU, MingLiU_HKSCS, BatangChe, GulimChe, DotumChe, GungsuhChe | Noto Sans Mono CJK JP | **metric** | OFL-1.1 | **0 differing over 1,106–2,264 shared codepoints each**, ideographs, kana and ASCII together | 16.0 MB /1 |
| MingLiU-ExtB, SimSun-ExtB | Noto Sans Mono CJK JP | metric on ASCII only | OFL-1.1 | 0/95 ASCII; the Ext-B ideographs they exist for are not in the JP font, so only ASCII was measurable | (same file) |

That is **26 of the 186 names** with a genuine metric-compatible open face, 11 of them
Latin or symbol and 15 CJK.

### Measured: what the usual suggestions actually do

These are the faces a search engine offers, measured rather than repeated. Every one of
them reflows. The mean ratio is over all 95 ASCII characters.

| Family | Suggested face | Verdict | Measurement |
| --- | --- | --- | --- |
| Franklin Gothic Book | Libre Franklin | approximate | 94/95 differ, mean 1.0508, range 0.534–1.524 |
| Franklin Gothic Medium | Libre Franklin | approximate | 94/95 differ, mean 1.0325 |
| Garamond | EB Garamond | approximate | 91/95 differ, mean 0.9722 |
| Century Gothic | Josefin Sans | approximate | 95/95 differ, mean 0.9093 — the URW row above is the answer here |
| Tw Cen MT | Josefin Sans | approximate | 95/95 differ, mean 1.0105, range 0.423–1.536 |
| Gill Sans MT | Libre Franklin | approximate | 95/95 differ, mean 1.0774 |
| Verdana | DejaVu Sans | approximate | 57/95 differ, mean 0.9613. The Bitstream Vera lineage is a design influence, not shared metrics |
| Tahoma | DejaVu Sans Condensed | approximate | 93/95 differ; mean 0.9994 is a coincidence of averaging, range 0.637–1.191 |
| Trebuchet MS | Arimo | approximate | 94/95 differ, mean 1.0216, range 0.495–1.515 |
| Candara | Carlito | approximate | 92/95 differ, mean 0.9950 |
| Corbel | Carlito | approximate | 94/95 differ, mean 0.9883 |
| Constantia | Caladea | approximate | 93/95 differ, mean 0.9574 — the same 4 % Caladea misses Cambria by |
| Rockwell | URW Bookman | approximate | 94/95 differ, mean 1.0658 |
| Perpetua | Nimbus Roman | approximate | 87/95 differ, mean 1.0782 |
| Lucida Bright | Nimbus Roman | approximate | 94/95 differ, mean 0.9207 |
| Arial Black | Nimbus Sans Bold | approximate | 85/95 differ, mean 0.8924 |
| Nyala (Ethiopic) | Abyssinica SIL | approximate | 355/358 differ, mean 1.1446. Noto Sans Ethiopic: 353/358, mean 1.1033 |
| David (Hebrew) | David Libre | approximate | 36/55 differ, mean 1.1223, and it covers only 55 of David's 88 Hebrew codepoints |
| Mangal (Devanagari) | Noto Sans Devanagari | approximate | 94/112 differ, mean 0.9455 |
| Latha (Tamil) | Noto Sans Tamil | approximate | 70/72 differ, mean 0.9943 |
| Gautami (Telugu) | Noto Sans Telugu | approximate | 82/93 differ, mean 1.0596 |
| Tunga (Kannada) | Noto Sans Kannada | approximate | 85/86 differ, mean 1.0150 |
| Kartika (Malayalam) | Noto Sans Malayalam | approximate | 95/98 differ, mean 0.8887 |
| TH SarabunPSK | Sarabun (Google Fonts) | approximate | 71/87 differ in the Thai block, **mean 1.5299** — Sarabun is drawn at a completely different optical size within the em |
| Angsana New | Noto Sans Thai | approximate | 71/87 differ, mean 1.3950 |
| Mongolian Baiti | Noto Sans Mongolian | approximate | 155/156 differ, mean 1.9053 |
| Dubai (Arabic) | Cairo | approximate | 86/101 differ, mean 1.1120 |

### None known

No open face claims, and none measured within reach of, the following. They are listed so
the next person does not re-search them: the Monotype and ITC display and script faces
(Abadi MT, Baskerville Old Face, Bauhaus 93, Bell MT, Bernard MT, Braggadocio, Britannic,
Calisto MT, Colonna MT, Cooper, Copperplate Gothic, Curlz MT, Desdemona, Edwardian Script
ITC, Engravers MT, Eurostile, Footlight MT, Gabriola, Gill Sans Ultra Bold, Gloucester MT,
Goudy Old Style, Haettenschweiler, Harrington, Imprint MT Shadow, Kino MT, Matura MT,
Mistral, Modern No. 20, News Gothic MT, Onyx, Perpetua Titling MT, Segoe Print, Segoe
Script, Stencil, Wide Latin); the whole Lucida family, which is Bigelow & Holmes and has
never been cloned; the Microsoft symbol and dingbat fonts (Bookshelf Symbol 7, Marlett, MS
Reference Specialty, MT Extra, Webdings, Wingdings 1/2/3, Segoe UI Symbol, Segoe UI
Historic) — **Wingdings against D050000L differs on all 202 shared codes, mean 0.8632**,
which is the measurement behind "Wingdings is not Dingbats"; Cambria Math, where STIX Two
Math is the free stand-in and shares nothing but purpose; the minority-script faces
(Microsoft Himalaya, Tai Le, New Tai Lue, Yi Baiti, Myanmar Text); and the nine HG* faces,
which are Ricoh designs bundled for Japanese Office.

### Four things that turned up that were not the question

* **Verdana and MS Reference Sans Serif are the same metrics.** 0 differing over all 191
  Latin-1 characters. Neither has a clone, so this changes nothing today, but any future
  Verdana clone answers two names, and any substitution rule written for one must be
  written for both.
* **TeX Gyre is not URW, and the difference is exactly 14 characters.** The TeX Gyre
  families are redrawn from the URW base-35 and are usually described as metric-compatible
  with them. Measured, Bonum and Heros are byte-identical to URW Bookman and Nimbus Sans;
  Pagella, Schola, Adventor and Termes are identical on letters and digits and **deliberately
  rewidened on `( ) [ ] { } + < = > | / \ *`** — `(` goes from 0.333 to 0.456 em in Pagella,
  0.333 to 0.483 in Schola. On a realistic sentence the cost is small and quantified:
  "Revenue grew 18% in Q3 (year on year), driven by EMEA." measures **1.0092** in TeX Gyre
  Pagella against Book Antiqua and **1.0000** in URW P052; sentences without brackets come
  back 0.9996–1.0001. So TeX Gyre is metric-compatible for prose and carries roughly a 0.9 %
  penalty per parenthetical. That is a fifth of the error that disqualified Caladea, and it
  is a real error rather than none.
* **TeX Gyre Chorus is not Zapf Chancery.** All 95 ASCII differ from Z003 by a uniform
  1.081, so it is the same proportions at a different design size — which still reflows.
  Monotype Corsiva's only metric clone is the AGPL one.
* **Selawik is metric-compatible with Segoe UI — for three of its five cuts.** Microsoft's
  own OFL release claims the compatibility flatly. Measured against the Segoe UI that
  Microsoft Remote Desktop installs: Regular 0/188 over Latin-1, Light 0/95 — exact.
  **Semilight is not: 51 of 95 differ, and `1` is 19 % out.** Bold and Semibold could not be
  measured, no Segoe UI Bold being present on this machine. It is also a 352-glyph font,
  which is Latin-1 and little else. So the claim is true where it is most used and false
  where nobody checked, which is the shape of every finding in this survey.

### CJK compatibility is nearly free, and mostly already held

The single most useful structural finding in the survey. **Every Microsoft CJK face gives
every ideograph and every kana an advance of exactly 1.000 em** — MS Gothic, MS Mincho,
Meiryo, Yu Gothic, Yu Mincho, DengXian, Microsoft YaHei, JhengHei, Malgun Gothic, SimSun,
SimHei, MingLiU, Batang, Gulim, all of them — and so does the Noto Sans JP the bundle
already ships. Measured over 1,570 shared ideographs and kana, Noto Sans JP differs from
Meiryo, Yu Gothic, Microsoft YaHei, Malgun Gothic and DengXian on **zero** codepoints.

So the comment in `text/fontmap.py` that reads "Noto Sans JP is not metric-compatible with
any of these (nothing is; the MS faces are proprietary and were never cloned)" is too
strong, and the survey contradicts it. For the CJK portion of a CJK run it *is* exactly
compatible, and has been all along. The divergence is confined to the Latin sub-run, which
is 0.83 of Meiryo's widths and 0.90 of Yu Gothic's. Three faces are the exception and they
prove the rule: **MS PGothic makes kana proportional too** (176 kana differ, which is why its
own table already exists and is right), and Gungsuh and STZhongsong do the same.

The fixed-pitch families are the strong case. MS Gothic, MS Mincho, SimSun, NSimSun, SimHei,
KaiTi, FangSong, MingLiU, MingLiU_HKSCS, BatangChe, GulimChe, DotumChe and GungsuhChe are all
(1.0 em ideograph, 1.0 em kana, 0.5 em ASCII), which is precisely Noto Sans Mono CJK JP's
profile, and the measurement is 0 differing codepoints for every one of them.

**Which makes the cheapest fix on this page a metrics-only one.** Those thirteen families
need no font file at all: their entire advance table is three constants. `fonts/` already
draws them with Noto Sans JP, and the only thing missing is a measured table, which costs
nothing to carry and does not enlarge the wheel by a byte. The same trick works for the
monospaced Latin faces, where the whole table is one number:

| Family | Advance | Nearest shippable | Drawn-width error |
| --- | --- | --- | --- |
| Lucida Console | 0.602539 em | Cousine (0.600098, already bundled) | **0.41 %** |
| Lucida Sans Typewriter | 0.602539 em | Cousine | **0.41 %** |
| Consolas | 0.549805 em | Anonymous Pro (0.545898) | 0.71 %, but needs a new four-cut family (158 kB regular) |
| Consolas | 0.549805 em | Cousine | 9.1 % — too wide to draw with |

Inconsolata is 0.5 and JetBrains Mono, Source Code Pro and Noto Sans Mono are all 0.6; no
open monospace was found at Consolas's 1126/2048.

### Licence: the URW result is the best measurement and the worst licence

The URW base-35 set produces the strongest numbers in this survey and is the one candidate
group that **cannot be adopted without a deliberate decision about the package licence.**
Read first-hand from `LICENSE` in `ArtifexSoftware/urw-base35-fonts`:

> The font and related files in this directory are distributed under the GNU AFFERO GENERAL
> PUBLIC LICENSE Version 3 [...] with the following exemption: As a special exception,
> permission is granted to include these font programs in a Postscript or PDF file [...]

**The exception covers embedding a glyph in a PS or PDF document. It does not cover
redistributing the font files.** This is narrower than the GPL font exception people assume
when they see "URW fonts are basically free". `pptx2svg-fonts` currently declares
`MIT AND OFL-1.1`; shipping URW would make it `MIT AND OFL-1.1 AND AGPL-3.0-only`, with the
network clause attached, inside a wheel that an MIT library pulls in as an extra. It also
would not help Phase 6's eventual SVG font embedding, since SVG is neither PostScript nor
PDF. **Flagged, not decided** — it is a licensing judgement, not a technical one, and it
belongs to whoever owns the project's licence policy.

The alternative is TeX Gyre. Its GUST Font License is LPPL 1.3c plus a *request* (explicitly
"not legally required") to rename derived works — permissive, non-copyleft, and compatible
with an MIT distribution as long as the licence text travels with the files. Its cost is the
14 rewidened punctuation characters measured above.

Licences confirmed by reading the shipped text: URW (AGPL + the quoted exception), GUST
(fetched from gust.org.pl), Symbol Neu (Apache-2.0, from its own `name` id 13), Selawik
(OFL-1.1, from `LICENSE.txt` in `microsoft/Selawik`), and Comic Relief and the Noto/Google
Fonts candidates (OFL-1.1, the `OFL.txt` beside each family in `google/fonts`).

### What the reference implementation does about this, which is almost nothing

`aiden0z/pptx-renderer` is this project's standing reference and was checked for a mapping
table worth borrowing. It has one, in `src/renderer/fontResolver.ts`, and it is eleven
entries long: `FONT_FAMILY_ALIASES` maps `calibri`, `calibri light`, `aptos`, `aptos
display` and seven CJK spellings (`microsoft yahei`, `微软雅黑`, `dengxian`, `等线`,
`simhei`, `黑体`, `heiti sc`) to **CSS font stacks**, not to faces it ships. The only
Western clone named anywhere in it is **Carlito**, third in Calibri's stack behind Calibri
and Aptos; there is no Arimo, Tinos, Cousine, Caladea or Liberation, and no concept of
metric compatibility at all. Everything else falls through `cssFontFamilyStack()` to
`sans-serif` and whatever the browser has.

That is not a criticism of it — a browser renderer cannot measure, so a font stack is the
only tool it has — but it does mean **there is no table here to copy.** It also makes the
same point the Fonts section opens with from the other direction: aiden0z solved its font
problem by reading `<p:embeddedFontLst>`, which is Phase 6, and left the clone problem to
the user's machine.

One correction and one further negative, so neither is re-searched. **Liberation Sans
Narrow exists only in the 1.x line** — the 2.x OFL releases ship Sans, Serif and Mono and
no Narrow — but it is not moot, as first written here: Debian packages it as
`fonts-liberation-sans-narrow` 1:1.07.6-4 in main, and measured against the Arial Narrow
macOS installs it is **0 of 95 in all four cuts**, the cleanest Arial Narrow result of the
three. Its licence is **GPL-2 with the Liberation font exception**, and that exception
covers embedding a glyph in a document, not redistributing the files — the same shape as
URW's, and the same reason it stays out of the wheel while being the best `apt` route.
TeX Gyre Heros Cn remains the shortlist pick for the bundle. And croscore has **a sixth face that this project has never heard of**:
**Symbol Neu**, Google's Apache-2.0 metric clone of Microsoft Symbol, which measures 0/188
against the installed `symbol.ttf` by legacy code and 0/189 by glyph name. It is not in any
current package — Google dropped it after croscore 1.23.0 on the reasoning that browsers
map Symbol themselves, and Debian dropped it with them — so it has to come out of the
archived `croscorefonts-1.23.0.tar.gz` on `chromeos-localmirror`, which is a supply-chain
caveat rather than a licensing one. Beyond it, croscore and crosextra hold only the five
faces already bundled (Arimo, Tinos, Cousine; Carlito, Caladea).

### The shortlist

Ranked by names rescued from the width guess per megabyte added to a bundle that is
currently 11 MB packed and 20.5 MB on disk for eight families. This is a shortlist for the
*wheel*, which redistributes and must therefore be conservative. `tools/install-fonts-debian.sh`
is under no such constraint — it installs from Debian's archive and redistributes nothing —
so its `--clones` tier takes the measured-best face in every row, `fonts-urw-base35`
included. The bundle's inclusion rule
— every family is there because Office will draw with it — is respected: everything below
is a clone of a face PowerPoint installs.

**Ship. No licence question, measured exact, and cheap.**

| | Adds | Rescues | kB per name |
| --- | --- | --- | --- |
| 1 | **Metrics-only CJK and mono tables** (0 kB) | MS Gothic, MS Mincho, SimSun, NSimSun, SimHei, KaiTi, FangSong, MingLiU, MingLiU_HKSCS, BatangChe, GulimChe, DotumChe, GungsuhChe, Lucida Console, Lucida Sans Typewriter | **0** |
| 2 | **Symbol Neu** (Apache-2.0, 1 file, 69 kB) | Symbol | 69 |
| 3 | **Comic Relief** (OFL, 2 cuts, 171 kB) | Comic Sans MS | 171 |
| 4 | **TeX Gyre Heros Cn** (GUST, 4 cuts, 467 kB) | Arial Narrow | 467 |
| 5 | **TeX Gyre Bonum** (GUST, 4 cuts, 524 kB) | Bookman Old Style | 524 |
| 6 | **TeX Gyre Adventor** (GUST, 4 cuts, 693 kB) | Century Gothic | 693 |
| 7 | **TeX Gyre Pagella** (GUST, 4 cuts, 875 kB) | Book Antiqua, Palatino Linotype | 438 |
| 8 | **TeX Gyre Schola** (GUST, 4 cuts, 813 kB) | Century Schoolbook, Century | 407 |

Rows 1–8 together are **3.6 MB on disk for 24 names**, against 20.5 MB for the 16 the
bundle answers today. Rows 1–5 carry no caveat at all: the metrics-only tables and Symbol
Neu are exact, Comic Relief is exact, and Heros Cn and Bonum are metrically identical to
the URW originals. Rows 6–8 carry the 0.9 %-per-parenthetical note, which `Substitution`
already has a field for — this is what `caveat` is for, and they should still grade
`compatible`, because we would measure and draw with the same face.

Row 1 is first on the list because it is free, and it is free for a reason worth stating
plainly: **for a fixed-pitch face the advance table is a constant, and a constant is a fact
about a design, not a copy of it.** The same argument that lets `METRICS` carry Aptos's
widths lets it carry MS Gothic's — and MS Gothic's is one number repeated. Nothing has to
be downloaded, licensed or shipped.

**Decide, do not drift into.** With Symbol Neu taking Symbol, two AGPL rows are left that
have no alternative anywhere: **D050000L** (37 kB) for Monotype Sorts and **Z003** (215 kB)
for Monotype Corsiva, TeX Gyre Chorus having been measured out of contention above. 252 kB
for two more names is a good ratio *and* it is the AGPL. Worth putting to whoever owns the
licence question precisely because the amount is small enough that it would otherwise get
waved through.

**Do not ship.**

* **Noto Sans Mono CJK JP, 16 MB.** It is exactly right for thirteen families, and the
  metrics-only row above already captures most of that value for nothing. Sixteen megabytes
  — enough to double the bundle — buys only the correct *drawn* Latin inside a fixed-pitch
  CJK run. Wrong trade at this size.
* **Every "approximate" row in the second table.** Libre Franklin, EB Garamond, Josefin Sans,
  DejaVu Sans, Abyssinica SIL, the Noto Indic set. Shipping one of these would mean drawing
  a face that reflows and grading it `approximate`, which is the Aptos arrangement — and
  Aptos earns that arrangement by being Office's *default*, which Tw Cen MT is not. For
  everything else, a `font-substituted` warning naming the truth beats a substitute that
  quietly moves every line break.
* **Gelasio and Selawik**, despite both measuring exact — Gelasio 0/95 against Georgia,
  Selawik 0/188 against Segoe UI — because **neither Georgia nor Segoe UI is in `DFonts`**.
  macOS supplies Georgia, as it supplies Courier New, Trebuchet MS and the Comic Sans
  regular cut; Segoe UI is a Windows system face that Office draws with there and that
  PowerPoint for Mac carries only as `.woff` add-in chrome under `Resources/sdx/`. Both are
  out of scope for a survey of what PowerPoint *ships* and are recorded so the results are
  not re-derived.

  **Selawik is the strongest argument for widening that rule**, and should be put to whoever
  owns it rather than settled here. It is 216 kB for five cuts, OFL, published by Microsoft,
  and Segoe UI is named by real decks far more often than Bookman Old Style is. The rule it
  fails — "Office ships it" — is already stretched by Lato and Raleway, which are in the
  bundle because Office *offers* them from the cloud. If the rule becomes "Office will draw
  with it", Selawik is the first row and Gelasio the second.

### What could not be settled

* **Whether the URW AGPL exception's reach can be argued wider.** The text is unambiguous
  about what it *grants*; whether ordinary aggregation of unmodified font files in a wheel
  needs the exception at all is a legal question this survey is not competent to answer,
  and guessing at it would be the licensing version of asserting metric compatibility.
* **Coverage beyond Latin-1 for the TeX Gyre and URW candidates.** All the numbers above are
  ASCII, Latin-1 or the relevant script block. A deck in Book Antiqua with Greek or Cyrillic
  in it has not been measured.
* **Whether `Century` and `Century Schoolbook` being metrically identical is by design.**
  They are two menu entries with the same 95 ASCII advances. Both map to C059 either way, so
  nothing depends on the answer.
* **Every number here is from one machine's PowerPoint 16.106.** A different Office build
  could ship a different cut of the same family. Nothing in the survey depends on a version
  Microsoft could quietly change, but a re-check is cheap and has never been done twice.
* **Kerning, everywhere.** This survey compares `hmtx` advances only. `pptx2svg` does not
  apply kerning and PowerPoint does, which is a pre-existing gap and not one the clone
  choice changes — but a clone with different kern pairs will diverge from PowerPoint by
  more than these tables suggest once that gap is closed.

---

## Suggested order

```
Phase 0  (PowerPoint oracle + VRT)  ──┬─▶ Phase 1  (parsed-but-unrendered)   DONE
                                      ├─▶ Phase 2  (SmartArt — DONE)
                                      ├─▶ Phase 4  (EMF previews — DONE)
                                      ├─▶ Phase 5.1 DONE / 5.2  (small gaps)
                                      ├─▶ Phase 3  (charts — reader + 10 types done)
                                      └─▶ Phase 6  (embedded fonts — DONE)
                                                              Phase 5.3 last
```

Phase 0 gates everything and is now much stronger than originally planned, because the
reference implementation is available locally.

Revised quick wins, in order of payoff per day:

1. ~~**Phase 1 table styles**~~ — done, though not in an afternoon: the built-in
   definitions are not in the file and had to be measured out of PowerPoint.
2. **Font metrics** — still the largest single component of the residual on every
   fixture. Phase 6's embedded fonts have closed the half a bundle can never reach: a
   deck that carries its faces is now measured and drawn from them, which moved twelve
   of twelve slides closer to PowerPoint on the two template decks that motivated it.
   Open clones for PowerPoint's own faces are the other half; they do not overlap, and
   only a deck that names an Office face and embeds nothing still gets the width guess.
3. ~~**Phase 4 EMF previews**~~ — done; the S estimate held.
4. ~~**Phase 2 SmartArt**~~ — done; the S estimate held.
5. ~~**Phase 5.1 shapes**~~ — done. Not the additive, near-zero-risk job it looked
   like: it turned up two arc-conversion bugs that had been silently misdrawing
   custom geometry, and it replaced hand-transcription with a spec compiler.

Phase 3 is well along: the reader and ten chart types -- `barChart`, `lineChart`,
`areaChart`, `scatterChart`, `bubbleChart`, `pieChart`, `doughnutChart`, `ofPieChart`,
`radarChart` and `stockChart` -- are done and measured, and no chart in the corpus warns
`chart-unsupported-type` any more. `surfaceChart` is measured and deliberately deferred;
see *Surface -- measured, and deferred* for what PowerPoint actually draws and what would
close it. The shared infrastructure -- value domain, tick selection, number formatting,
gridlines, legend layout, plot-area rectangle, polar region -- is built, and the three
types that landed this week reused nearly all of it. What is left, cheapest first:

1. ~~**The CJK label width.**~~ Done, and the diagnosis was half wrong — see *The East
   Asian face cascade*. `ChartFont` now carries the East Asian face as well as the Latin
   one, measures per character and names both in the emitted `font-family`, which takes
   5.00 pt of the radar's 6.83 pt radius error. Chart3's 19.8 pt turned out not to be a
   width at all: PowerPoint draws `プラットフォーム` at the same 96 pt we measure. What
   is left there is **the rotated reserve's missing cap**, with a candidate rule on three
   observations and a probe deck needed to settle it.
2. **A manually laid out legend.** `real-college-template.pptx` slide 4 carries a
   `c:legend/c:layout/c:manualLayout` we ignore, and it is now the whole of that slide's
   remaining chart error: 4.8 pt of plot height and 9.5 pt of legend baseline. Reading
   the element is trivial; what is not known is how PowerPoint composes the plot area
   around it, and one deck is not enough to fit that.
3. **Data-label wrapping**, **`bestFit` actually moving a label**, and **three-or-more
   line legend entries** -- each a known-missing detail with a named symptom above.
   Category-label wrapping is done; the data-label kind is a separate path.
4. **Combo charts and the secondary axis.** `bubbleChart`, `ofPieChart` and `stockChart`
   have landed; `surfaceChart` is measured and deliberately deferred, and what would
   close it is written down under *Surface -- measured, and deferred*. Drawing several
   `c:*Chart` groups at once, and the second value axis that usually comes with them, is
   now the largest chart item left.
5. **The value-axis tick density.** Still unsolved, and the stock sweep has made it
   firmer rather than looser: the accepted spacings now run 14.50 to 27.3 pt and the
   refused ones 12.10 to 15.43 pt, and they **overlap**, so no threshold on spacing alone
   orders the nine observations at any axis length. It is the whole of the stock probe
   deck's shortfall and the only visible difference between our render of the area deck
   and PowerPoint's.

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

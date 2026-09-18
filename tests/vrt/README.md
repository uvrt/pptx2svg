# Committed render snapshots

One SVG per slide of every deck in `tests/fixtures/`, exactly as `convert_pptx_to_svg`
produced it on the day it was captured. `tests/test_vrt.py` renders the same decks and
compares byte for byte.

## Read this before you read a file in here as an assertion of correctness

**These files are not ground truth. They are a photograph of our own output, bugs
included.** A snapshot can tell you that the render *changed*; it can never tell you that
it is *right*. Correctness comes from the other layer — `tools/fidelity.py` scores a
render against PowerPoint's own PDF export — and that layer needs PowerPoint, so it runs
on one Mac while this one runs on all twelve CI legs.

The distinction is not academic here, because **the oracle scored only three of these
decks for most of this project's life**. `tests/fidelity-baselines.json` records where
each one stands now. Installing the bundled Noto Sans JP into `~/Library/Fonts` and
re-exporting (see FONTS.md ▸ *Making the oracle draw Japanese*) brought
`real-financial-report` and `sample-issue-387` in; `sample-cjk` was derived to bring a
Japanese deck in. Three are still skipped, each for a difference in the two renderers'
*inputs* rather than their output: `real-basic-theme` and `sample` resolve their Japanese
to ＭＳ Ｐゴシック, which this PowerPoint cannot use and which is not ours to install, and
`real-product-page` carries emoji that no face either deck names can draw. Their
snapshots record what we do, and that is all they record.

Nothing scores 1.0: `authoring-integration` is at SSIM 0.9327, `table test` at 0.9895,
`sample-issue-387` at 0.9897, `real-financial-report` at 0.9112, `sample-cjk` at 0.6788
and `chart-gallery` at 0.6747 — the last being a deck built to be scored, whose low mean
is mostly the chart types we knowingly do not draw plus SSIM's severity on sparse line
art. ROADMAP.md *3.2a* breaks it down per slide.

## Defects that are baked into these files

From `ROADMAP.md`, *What the corpus says is wrong now*. Two of that list are in these
files, both on `real-financial-report.pptx` and both checked against the committed bytes
rather than assumed:

| Defect | Frozen into |
| --- | --- |
| The rotated-label reserve has no cap, where PowerPoint's does — slide 3 reserves too little under its plot | `real-financial-report/slide-03.svg` |
| PowerPoint ellipsis-truncates a category label that will not fit; we draw it in full | `slide-03.svg` draws `プラットフォーム` where PowerPoint's export drew `プラット…`; `slide-04.svg` draws `海外売上比率` and `従業員満足度` where it drew `海外売上…` and `従業員満…` |

`chart-gallery.pptx` adds three more, all of them deliberate — it is a fixture built to
pin what charts do now, including where that is wrong. Named here so nobody reads one of
its snapshots as an assertion that we are right:

| Defect | Frozen into |
| --- | --- |
| A combo chart draws only its first group — no line series and no secondary value axis | `slide-17.svg` |
| A `pie3DChart` and a stacked `area3DChart` are drawn flat; PowerPoint draws a real scene under both. The other three 3-D spellings draw theirs — `bar3DChart`, `line3DChart` and, since the surface landed, `surfaceChart` on slide 12 | `slide-15.svg`, `slide-16.svg` |
| A `stockChart`'s gridlines and axis sit about a pixel off PowerPoint's, which is the plot-rectangle defect slide 1 carries and not a legend one | `slide-11.svg` |

The three defects this list used to name on slides 6, 9 and 11 are **fixed** — see
ROADMAP.md 3.2b — and those bytes are an assertion that we are right about them now.

The rest of that list — the manually laid out legend, and body copy landing 1–2 px off —
was measured on `real-college-template.pptx`, which is not committed and not snapshotted.
Whether the same 1–2 px displacement is in these files is *unknown*, because five of the
seven decks have never been scored against PowerPoint at all.

Fixing any of this will turn the suite red. That is the suite working. Rebaseline, and say
in the commit message which defect the new bytes fix.

## Rebaselining

```
python -m pytest tests/test_vrt.py --update-snapshots
```

It skips rather than passes, so the flag can never produce a green run. Read the diff
before committing it: every changed line is a change in what a user sees. On a failure
the new render is written to `tests/vrt/diffs/` (gitignored) so the two can be opened side
by side in a browser.

## Why SVG, and not PNG

The SVG is our own output, from standard-library code alone. A PNG additionally encodes
the rasteriser: bump `resvg-py` and every fixture goes red at once, which teaches people
to rebaseline without looking. Measured on this corpus, PNG is also 587 kB against the
SVG's 195 kB, and being already compressed those bytes never shrink — in the pack file or
in history, on every rebaseline, forever. resvg 0.5.0 does render byte-identically
between processes here, so a PNG layer with a tolerance is *possible*; it is just a
second, weaker gate paid for in permanent repository weight.

## Why there is no snapshot of `real-college-template.pptx`

It is not ours to redistribute (see `tests/fixtures/README.md`), and a rendered SVG of it
contains its text and its images — redistributing the render would redistribute the deck.
`tools/fidelity.py` measures it locally; nothing about it is committed.

## What is deliberately not snapshotted

`ConvertOptions.warnings`. They depend on whether the `pptx2svg-fonts` bundle is
installed, which is a property of the machine, not of the deck. The SVG itself does not —
verified across the whole corpus and held there by
`test_the_render_does_not_depend_on_the_font_bundle`.

## If one CI leg goes red and the others do not

That is not a rendering change — nothing in the render depends on the platform by design.
Look at `tests/test_vrt.py`'s own tests first; they name the three ways it could happen
(line endings, `PYTHONHASHSEED`, and a libm that disagrees about `log10` of a power of
ten — which used to move `real-financial-report/slide-04.svg` and nothing else, and no
longer moves anything: `resolve/chart.py` verifies that exponent rather than trusting it,
and `test_a_one_ulp_error_in_log10_does_not_move_a_snapshot` renders that deck with
`log10` nudged both ways to prove it). Do not rebaseline: a snapshot captured on one
platform to satisfy another is a snapshot that has stopped meaning anything.

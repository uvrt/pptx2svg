# Test fixtures

Real `.pptx` files produced by PowerPoint, Google Slides and python-pptx, used as
end-to-end inputs. They exercise structures that are awkward to synthesise by hand:

- theme font references (`+mj-lt`, `+mn-lt`)
- `defaultTextStyle` in `presentation.xml`
- `txStyles` on the slide master
- style references (`a:lnRef` / `a:fillRef` / `a:effectRef`)
- version-specific XML quirks from real Office builds

| File                          | Produced by    | Slides | Covers                                             |
| ----------------------------- | -------------- | ------ | -------------------------------------------------- |
| `real-basic-theme.pptx`       | Google Slides  | 2      | title/content layouts, tables, images, theme fonts |
| `real-product-page.pptx`      | hand-authored  | 1      | rounded rectangles, ellipses, text boxes           |
| `real-financial-report.pptx`  | hand-authored  | 4      | CJK text, tables, charts, dense layouts            |
| `sample.pptx`                 | PowerPoint     | 6      | assorted shapes and text                           |
| `sample-cjk.pptx`             | `tools/make_cjk_deck.py`, from `sample.pptx` | 6 | the same deck with a Japanese face both renderers draw — the only scorable CJK deck |
| `sample-issue-387.pptx`       | PowerPoint     | 1      | regression case for text layout                    |
| `authoring-integration.pptx`  | python-pptx    | 1      | one of each element type: shape, picture, connector, table, chart |
| `chart-gallery.pptx`          | `tools/make_chart_gallery.py` | 17 | one chart type per slide: every `c:*Chart` group element the reader knows |

## The chart gallery

`chart-gallery.pptx` is **generated, not collected**: `tools/make_chart_gallery.py` writes
every byte of it, so it can be regenerated and reviewed rather than taken on trust. Read
that file's module docstring for what each slide carries and why.

It exists because charts are the largest body of measured behaviour in this project — ten
drawable types, an axis rule settled over 616 probe readings, legend and plot-area layout
— and **no other committed deck is chart-heavy**. `real-financial-report.pptx` has four
charts and the oracle skips it; `authoring-integration.pptx` has one.

Two properties were designed in and are worth keeping:

* **The oracle scores it.** Its only faces are the theme's Aptos and Aptos Display, which
  PowerPoint draws natively here, and there is no CJK text anywhere — so
  `tools/fidelity.py` compares two renders of the same outlines rather than measuring
  which fonts this machine happens to have. It is the first chart-heavy deck that has ever
  been scored; see ROADMAP.md *3.2a* for what the score is made of.
* **It names its own defects.** Slide 17 (a combo chart) is expected to draw only its
  first group, and slides 15 and 16 to draw their 3-D scenes flat. All three are pinned by
  `tests/test_chart.py` and by the snapshots, so a change either way is visible. Slide 12
  (`surfaceChart`) used to be the fourth and is not: its mesh, its value bands and its band
  legend are drawn, and the test there now asserts that **no** slide in the deck refuses.

There is deliberately **no ChartEx (`cx:chartSpace`) slide**: four hand-written ones all
hang PowerPoint on open, and a deck PowerPoint will not open is not a fixture. ROADMAP.md
0.1 records what was tried.

## Provenance

Most of these were copied from
[pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse)'s `shared-fixtures/`
directory (MIT licensed). Keeping the same inputs means rendering differences between the
two implementations can be compared directly.

## A deck that is used but not committed

`real-college-template.pptx` is measured against but **deliberately not in this
directory**. It is Dickinson College's public sample presentation -- a third-party
document that is ours to render, not to redistribute -- so it lives outside the
repository, in the gitignored `scratch/`.

It matters because it is the only real-world deck whose typefaces this machine can supply
in full, and therefore the only one `tools/fidelity.py` scores rather than skips. Four
defects were found through it and are now covered by tests that do not need it: the
`c:invertIfNegative` default, the `c:idx` accent cycle, bold inheriting through the
placeholder cascade, and `spcBef`/`spcAft` adding rather than collapsing.

To run the handful of tests that do need it, put a copy in `scratch/`:

```
https://www.dickinson.edu/download/downloads/id/1076/sample_powerpoint_slides.pptx
sha256 ac7f2627645042190df3244cc25929f4b006d144fc2cac520e79ab376197bbbf
```

Without it they skip. `tests/conftest.py`'s `college_template` fixture is the single
place that looks for it.

`FIXTURES-README.md` carries pptx-glimpse's own listing, in Japanese, and records each
file's provenance -- source URL, download date and sha256 for anything added since.

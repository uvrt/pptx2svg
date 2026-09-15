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
| `sample-issue-387.pptx`       | PowerPoint     | 1      | regression case for text layout                    |
| `authoring-integration.pptx`  | python-pptx    | 1      | one of each element type: shape, picture, connector, table, chart |

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

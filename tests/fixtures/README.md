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
| `real-college-template.pptx`  | PowerPoint 2007 | 9     | EMF logos, a stacked column chart with a negative value, a table, bullet lists, a JPEG |

## Provenance

Most of these were copied from
[pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse)'s `shared-fixtures/`
directory (MIT licensed). Keeping the same inputs means rendering differences between the
two implementations can be compared directly.

`real-college-template.pptx` is not one of theirs: it is a public sample deck from
Dickinson College, added because it is the first real-world deck whose faces this machine
can supply in full, and so the first that `tools/fidelity.py` scores rather than skips.

`FIXTURES-README.md` carries pptx-glimpse's own listing, in Japanese, and records each
file's provenance -- source URL, download date and sha256 for anything added since.

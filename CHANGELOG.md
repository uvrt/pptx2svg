# Changelog

pptx2svg has no tagged release yet; every change so far is on the way to 0.1.0.
The full history is the [merged pull requests](https://github.com/uvrt/pptx2svg/pulls?q=is%3Apr+is%3Amerged).

## 0.1.0 -- unreleased (in development since 2026-09-11)

- A layout's or master's background picture is drawn: an inherited `p:bg`'s `r:embed` is
  now looked up in the part that owns it, not in the slide's relationships (where the
  same id is usually the notes slide, so the background was dropped with an
  `unsupported-fill-image` warning, or a wrong picture was drawn without one). The same
  holds for a theme's image fill reached through `p:bgRef`/`a:fillRef`, and for an
  `a:buBlip` picture bullet a slide inherits from its layout's or master's list styles.
- Python 3.14 and 3.15: CI runs the suite on both, on Linux, macOS and Windows, and the
  classifiers declare them. `requires-python` stays `>=3.10`; no library change, and the
  `png` extra's resvg-py has wheels for both. CPython 3.14's Windows builds deflate with
  zlib-ng, whose bytes differ from zlib's for the same entries (both valid; what a package
  holds is unchanged), so the test that `sample-cjk.pptx` is what `tools/make_cjk_deck.py`
  writes compares every entry everywhere and the archive's bytes where deflate is zlib's
  (`tests/zip_content.py`).
- An application's own font folder is measured from as well as drawn with:
  `ConvertOptions.font_dirs` (new), `convert_pptx_to_png(font_dirs=...)` and `--font-dir`
  now reach measurement (`fonts.office.user_layout_metrics`) -- a face there is measured
  from its file unless the tables measure its family as itself (Aptos, Calibri) -- and a
  face found there grades `exact` instead of raising `font-substituted`. Without an
  explicit `font_dirs`, the `OOXML_FONT_DIRS` environment variable (`os.pathsep`-separated,
  shared with docx2svg) is read, `svg_to_png` included; an explicit empty list means none.
  Added to the system's folders, searched before them. Changes the layout for existing
  `font_dirs` callers whose folder holds a face the tables do not measure as itself: it was
  drawn from the folder at guessed widths before. Requires ooxml-common 0.8.
- First port of pptx-glimpse: PPTX slides to SVG and PNG in pure Python, with the full
  placeholder, colour and text inheritance cascades, all ECMA-376 preset geometries,
  tables with PowerPoint's 74 built-in styles, and charts laid out from measurements.
- Embedded fonts read and used for measurement and drawing; the `pptx2svg-fonts` bundle
  of SIL OFL faces for reproducible PNG output; `pptx2svg fonts --check`.
- The format-neutral half (OPC, units, fonts, text metrics, DrawingML, charts, SmartArt,
  text bodies) moved to [ooxml-common](https://github.com/uvrt/ooxml-common) (#1, #5, #6, #9).
- Fidelity harness scored with one rasteriser on both sides, using every core (#2, #3, #4).
- Measured against PowerPoint: axis tick marks (#8), theme style fills and outlines (#10),
  colour transforms and run placement (#11), autofit on open (#12), kerning and Office's
  own faces by default on a Mac (#13, #14), Japanese face resolution and colour emoji (#15).
- The agent view: a compact per-slide SVG for a model to read (#16).
- Chart labels as PowerPoint draws them: in their `c:txPr` colour (`tx1` at 65%, data
  labels at 75%, as Office writes them), data labels in their own number format
  (`€12.4m`), and a radar's category labels 4% of the radius off their vertex.
- Text the PNG will not draw is said out loud: a `glyphs-missing` warning and a
  `MissingGlyphsWarning`, once per face and script, when no font the rasteriser loads
  answers to a run's `font-family` or has its characters (`pptx2svg.glyphs`) -- every CJK
  character, rendered without `pptx2svg-fonts` on a host with no CJK font, vanished
  silently. The README says what `pptx2svg-fonts` holds and under which licence.
- `table_row_heights`: a table's rows as the renderer lays them out, grown to fit their text,
  for a caller that checks a slide without drawing it (pptx-agent's table overflow fact).

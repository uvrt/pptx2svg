# Changelog

pptx2svg has no tagged release yet; every change so far is on the way to 0.1.0.
The full history is the [merged pull requests](https://github.com/uvrt/pptx2svg/pulls?q=is%3Apr+is%3Amerged).

## 0.1.0 -- unreleased (in development since 2026-09-11)

- **Rasterising loads only the fonts a slide can use** (round 10, O2; ooxml-common 0.23.0,
  now the floor). resvg-py builds a new font database on every call from every file of
  every folder it is handed; `svg_to_png` now hands it only the files a slide can use --
  the faces filed under a name its runs can resolve to and those its missing characters
  fall back to -- in the order resvg loads them (`ooxml_common.fonts.rasterdb`): 26 bundle
  files become 10, a 280-file Office folder plus the bundle 22. And a slide drawn in two
  passes (Aptos Display and Aptos answer to the same name in resvg) no longer lays out in
  each pass the texts that pass draws nothing of (`fonts.office.without_undrawn_text`;
  kept under a filter, mask, clip, opacity or in a definition). Same pixels: 535 slide
  renders (fixtures, local decks, five font setups) byte-identical to before. On this Mac,
  where opening a file is cheap, one 26-text slide 159 -> 147 ms and 41 slides 6.8-7.3 ->
  6.4-6.5 s; the work removed is file opening and text layout, which is what was slow
  where production measured it.
- EMF+ pictures are drawn from their EMF+ records (round 7, L9 step 2; ooxml-common 0.22.0,
  now the floor), as Windows PowerPoint draws them: a dual file's GDI fallback is no longer
  what is drawn, and an EMF+-only file is no longer a grey box. Gradients (linear and path),
  hatches, textures, pens with caps, joins, dashes and compound bands, curves, paths, clips,
  containers, bitmaps and text. A file whose EMF+ records draw nothing is still drawn from
  its GDI records, with a `metafile-records` warning saying so. Goldens and VRT snapshots: no
  fixture carries a metafile, none changed.
- EMF and WMF pictures are drawn from their GDI records (round 7, L9; ooxml-common 0.21.0,
  now the floor): pens, brushes, hatches, lines, polygons, rectangles, ellipses, arcs,
  paths, clips, mapping modes, world transforms, bitmaps and text, as **Windows**
  PowerPoint draws them (`ooxml_common.metafile`; Mac PowerPoint's player differs). A
  metafile with no embedded preview -- a template's dot grid, corner brackets -- was a
  grey box; it is now drawn. The order (`resolve/view.py`, `_metafile_preview`): a
  `metafile_converter`; the records, when every one was understood; an embedded PDF
  (Office's own rendition beats a drawing that left records out); the partial drawing,
  with a new `metafile-records` warning naming what it left out (EMF+ among them: a dual
  file is drawn from its GDI records, not yet its EMF+ ones); the embedded bitmap; the
  grey box. An EMF that is one bitmap is now an SVG placing it where its record says,
  not the bitmap lifted out. `svg_to_png` writes such pictures inline, since resvg draws
  no text in an SVG `<image>`. `pptx2svg.metafile.dib` moved to ooxml-common (re-exported).
  Goldens and VRT snapshots: no fixture carries a metafile, none changed.
- A rich chart title is drawn in PowerPoint's own fonts (ooxml-common 0.18.0, now the
  floor), measured on PowerPoint 16 for Mac's PDF export, 16 October 2026, in a deck whose
  text boxes are Aptos. A `c:title/c:tx/c:rich` whose paragraphs state no `a:defRPr` and
  whose runs name no face is **Arial, regular**, at the run's size or 18 pt — whatever the
  chart-wide `c:txPr` states, `+mn-lt` included — where it took the text box's face. A
  title whose paragraph states an `a:defRPr` is the theme's **minor face** unless a face
  is stated, **bold** unless `b` is stated (`sz` alone: Aptos bold; `b="0"`: regular). A
  run naming a face keeps it. Colour and stated sizes are unchanged.
  The probe decks' plot area now matches to 0.1 pt (190.9 pt from the 6 to the 0 tick
  label, PowerPoint 191.0; 188.2 before).
  **VRT snapshots changed: `chart-gallery` slides 1, 4, 6, 7, 9, 11, 12, 13, 15 and 17.**
  Their titles state `a:defRPr sz="1400" b="0"` and were drawn Arial (that deck's text
  boxes are Arial); they are now Aptos regular, which is what PowerPoint's own export of
  that deck draws (`chart-gallery.pdf`: every one Aptos 14 regular). The title band is
  Aptos's, so each plot moves: slide 1's value-axis step is now 21.89 pt against
  PowerPoint's 21.74 (22.19 before). Chart glyphs at whole points (PowerPoint draws 18.62
  pt as 19) are measured and not applied (ROADMAP.md).
- A slide layout or master can be drawn on its own (round 6, R11):
  `ConvertOptions(templates=[part path, ...])` adds a template view of each after the
  slides (`slide_numbers=[]` for none), for PNG too. What a slide on the layout starts as
  -- background, the master's shapes as `showMasterSp` says, the layout's shapes -- with
  every placeholder a grey dashed box holding its prompt in its inherited text style, a
  tag with its type and `idx`, and a banner naming the layout or master
  (`pptx2svg.resolve.templates`). A wireframe of this library's, not PowerPoint's drawing
  (Slide Master view is not measured). The SVG root carries `data-template`,
  `data-template-part` and `data-template-name`; `Slide.template` and friends in the
  model; a `glyphs-missing` warning on a view names its part (`part_path`), its slide
  number 0. Layouts and masters carry their `cSld` name in the source model. Slides draw
  as before: goldens and VRT snapshots unchanged.
- A chart's date categories are drawn as dates (ooxml-common 0.17.0, now the floor): a
  category kept as a serial under a date code -- `[$-413]d\ mmmm\ yyyy`, as ooxml-edit's
  `date_format` writes it -- reads `5 januari 2026` where it read `46027`, as PowerPoint
  16 drew it (measured 16 October 2026; `[$-407]d\ mmm\ yyyy` `15 Mär 2026`,
  `[$-409]dd/mm/yyyy` `05/01/2026`).  A code without a language prefix takes its month
  names from `ConvertOptions.locale`'s language, as PowerPoint takes its interface's; a
  `/` is drawn as `/` (PowerPoint shows the viewer's region separator, which the locale
  does not hold -- as a date field's pattern is its language's, not the region's).
- A chart's automatic title is drawn: a `c:title` with no text over one series showed
  nothing; it is the series' name, bold, in the theme's minor face, at 1.2 times the
  chart's text size rounded to the point, 18 pt where the chart states none (measured on
  PowerPoint 16's PDF export: 8, 10, 11 and 24 pt chart text gave 10, 12, 13 and 29 pt).
  Over several series PowerPoint's own words are not measured and nothing is drawn, as
  before.  Goldens and VRT snapshots unchanged: none holds such a chart.
- `ConvertOptions(locale=...)` takes ooxml-common's shared `Locale` (0.15.0, now the
  floor): what PowerPoint takes from its machine rather than from the deck.  A chart's
  numbers -- data labels, axis ticks, percentages -- are printed in its decimal and
  thousands separators, and a date field with no `lang` (or one whose language has no
  measured table) is drawn in its language when `DATE_FORMATS` has it, else in en-US's as
  before; a field's own `lang` still wins.  The default is `EN_US`, never the operating
  system's locale, and draws byte for byte what it drew before: goldens and VRT
  snapshots unchanged.  The agent view takes it too.  The docs now say that `now` left
  to the default is the server's local time zone, which is not a locale field.
- The built-in "Themed Style 1" and "Themed Style 2" tables (all six accents of each) are
  drawn as PowerPoint draws them.  They draw on the theme's format scheme -- the table on
  a fill style with an effect style under it, every line a line style -- which the
  measured catalogue had flattened to one solid accent fill and a few white rules.
  `tools/derive_themed_table_styles.py` measures, on PowerPoint 16's PDF export under five
  themes built to tell the indices and the `phClr` colour apart, which fill, line and
  effect style each region uses and in which colour, the fills' alpha, the text's colour,
  weight and face, and the corner cells' exceptions; the twelve styles are written from
  that and read like a deck's own `tableStyles.xml`, which now reads `a:lnRef` borders and
  `a:tblBg` (fill or fill style, effects or effect style) too.  The background is drawn
  once under the whole table and its shadow under the table (ooxml-common 0.14); a bevel
  in the effect style is not drawn, as PowerPoint draws none on a table.  The tool's
  `check` finds no cell fill, edge, shadow or text that differs from PowerPoint's.
  Two neighbouring cells now share one style line on their common edge -- the one the
  region of higher precedence states, whichever side it is on (a header's bottom rule
  over the first body row's gridline), where the cell drawn last used to win.
  Under the 2013 and 2007 Office themes, every accent with six flag sets: 109 of 144
  tables match PowerPoint's export in fill, every grid edge and shadow (0 before); the
  other 35 are "Themed Style 2" on the 2007 theme, whose gradient PowerPoint interpolates
  lighter mid-way than a linear SVG gradient (as it does on a shape with the same fill
  style, not a table matter) and whose shadow is drawn a little lighter.  Two VRT
  snapshots change only in the first body row's top line under a Medium Style 2 header,
  now the header's 3 pt rule rather than a 1 pt one drawn over it: the same pixels but
  for anti-aliasing.

- An SVG picture (`asvg:svgBlip`) is drawn in PNG output whatever unit its root's size is
  in: resvg-py drew nothing for `width="72pt"` -- or `pc`, `in`, `cm`, `mm` -- ("SVG has an
  invalid size") and nothing for an SVG that is not UTF-8, so the picture vanished from the
  PNG while the SVG output showed it.  The page now reaches the rasteriser with each
  embedded SVG's `width` and `height` restated in pixels as PowerPoint reads them
  (`ooxml_common.svgimage`, 0.12): a length at 96 px to the inch, an `em` of 12 px unless
  the root states a font-size, the `viewBox`'s size for a percentage or none, 500 px
  without a `viewBox`; and a root with no `viewBox` gets one of its own size, which
  CairoSVG needs to scale it.  Measured on PowerPoint 16's export of 30 SVG pictures
  (`tools/make_svg_size_probe.py`, `tools/read_svg_size_probe.py`): resvg now draws all 30
  as PowerPoint does (8 before), CairoSVG 27 (10 before; it reads no root's
  `preserveAspectRatio`, so a root out of proportion with its `viewBox` is stretched
  rather than letterboxed).  The SVG output is unchanged.
- Three things PowerPoint 16 draws that were missing, each measured from its export
  (production feedback round 3): an empty placeholder's own fill and outline (the prompt
  text still is not drawn; one with neither is still left out); the "Light Style 2" family's
  1 pt lines -- the table's outline and the edges of the odd banded rows and columns, in the
  style's colour -- which the generated catalogue lacked (an overlay in
  `table_styles_builtin`, kept apart from the generated table); and a deck's own table
  style's unstated `wholeTbl` edges, which PowerPoint draws as 1 pt black lines (black
  whatever the theme's `dk1`; an edge stated as `a:noFill` stays undrawn).
- A data label with text of its own is drawn (ooxml-common): a typed or "Value From Cells"
  label (`c:dLbl/c:tx`, `c15:datalabelsRange`) prints its text with its runs' formatting
  instead of the point's value, on scatter and bubble charts too.
  `tools/make_point_label_probe.py` builds the deck it was measured on.
- A symbol face this conversion cannot draw -- Wingdings, Wingdings 2/3, Webdings, Symbol;
  not embedded, not in `font_dirs`, not on the host where the host's faces are in use --
  is drawn as the Unicode characters its codes stand for (ooxml-common's
  `text.symbol_fonts`): a Wingdings `q` box bullet as ❑, `ü` as ✓, `§` as ▪, `Ø` as ➢, in
  runs and bullets, spelled as the code or as its U+F0xx alias, with a
  `symbol-font-mapped` warning. Symbol and Wingdings are laid out at their own advances,
  recorded from Word's copies; the glyph choices were compared with PowerPoint 16's export
  of `tools/make_symbol_probe.py`. A face that is available still draws itself. The
  `pptx2svg-fonts` bundle (0.2.0) ships Noto Sans Symbols 2 (SIL OFL 1.1) to draw them
  with; no deck is measured in it.
- Slide numbers and dates are evaluated per slide, as PowerPoint draws them: an `a:fld`
  in a plain text box on the layout or master (the corporate footer "‹#› | Copyright …")
  and in the slide's own `sldNum`/`dt` placeholders shows the slide's number, counted from
  `p:presentation@firstSlideNum`, and the date, instead of the text cached in the file.
  `datetime1` to `datetime13` (and `datetime`, `datetimeFigureOut`) are formatted as
  PowerPoint 16 drew them for en-US, en-GB, nl-NL, de-DE and fr-FR
  (`tools/make_field_probe.py`); another language is drawn in en-US's formats with a
  `field-date-format` warning. The clock is `ConvertOptions.now` (new; `--now` on the
  command line), the time of the conversion unless pinned, so output with a date field
  changes from day to day unless it is; the snapshot tests pin it. Requires ooxml-common
  with `SourceTextRun.field_type`.
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

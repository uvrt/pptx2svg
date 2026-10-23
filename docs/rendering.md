# What gets rendered

The pipeline, and what is and is not drawn. Back to the [README](../README.md).

## How it works

```
.pptx  ─▶  OPC package  ─▶  source model  ─▶  render model  ─▶  SVG  ─▶  PNG
           pptx2svg.opc    pptx2svg.parse   pptx2svg.resolve   .render  .png
```

| Module              | Responsibility                                                           |
| ------------------- | ------------------------------------------------------------------------ |
| `pptx2svg.opc`      | ZIP container, content types, relationship graph                         |
| `pptx2svg.parse`    | OOXML → source model, deliberately **unresolved** (theme refs, rel ids intact) |
| `pptx2svg.resolve`  | Theme colours + the placeholder/background/text inheritance cascades      |
| `pptx2svg.render`   | Text measurement, line breaking, geometry, SVG output                    |
| `pptx2svg.text`     | Font metrics, measurement, wrapping                                      |
| `pptx2svg.fonts`    | Which faces we can draw, and what happens when we cannot                 |
| `pptx2svg.metafile` | EMF/WMF preview extraction                                               |
| `pptx2svg.png`      | Rasterisation via an external backend                                    |

Each stage is usable on its own.

The parse/resolve split is what makes inheritance work: a `<a:schemeClr val="tx1"/>`
means nothing until you know which colour map applies, and a placeholder's font size may
live on the slide, the layout, the master's `txStyles`, or the presentation's
`defaultTextStyle`. The parser records what the XML says; the resolver decides what it
means.

## What gets rendered

The goal is accurate text, shapes and spatial layout — not a pixel-exact reproduction of
every PowerPoint feature.

**Shapes and text.** All 186 preset geometries in ECMA-376, plus `lineInv`, plus full
custom geometry with guide formulas — every one of them verified against PowerPoint's own
PDF export at two aspect ratios. Text carries the complete inheritance cascade, bullets
and auto-numbering, line wrapping (Latin and CJK), autofit (the stored `normAutofit`
scale, as PowerPoint draws a file it has not re-fitted), vertical text, tabs and columns. Solid, gradient, pattern and
image fills; outlines with dashes and arrowheads; shadows, glow and soft edges; pictures
with cropping, tiling and colour adjustments; groups with nested coordinate spaces;
hyperlinks; and alt text as `aria-label`.

**Tables** render with merged cells, borders and fills, and with **PowerPoint's built-in
table styles** — all 74 of them, measured out of PowerPoint itself, because it keeps
those definitions inside the application and never writes them into the file. A table
that names one therefore renders banded and headed rather than as a bare grid. "Themed
Style 1" and "2" draw on the theme's format scheme instead -- the table on a theme fill
style (a gradient in Office's themes) with an effect style's shadow, every line a theme
line style -- so for them it is the index and the colour of each that were measured
(`tools/derive_themed_table_styles.py`).

**Charts** — `barChart`, `lineChart`, `areaChart`, `scatterChart`, `bubbleChart`,
`pieChart`, `doughnutChart`, `ofPieChart`, `radarChart`, `stockChart` and `surfaceChart`,
which is every group element ECMA-376 defines — are read and drawn with their axis range,
tick interval, gridlines, legend and data labels, laid out from constants measured out of
PowerPoint's PDF export. No deck in the test corpus warns `chart-unsupported-type`.

A **surface** is the one whose marks are not its series: it draws a lit lattice over
(category, series, value) through the scene's depth, cut into bands by the value axis' own
intervals and coloured band by band, with a legend of those value ranges rather than of
the series. `c:bandFmts` and `c:wireframe` are both drawn.

The 3-D spellings draw their scene where it is measured — `bar3DChart`'s prisms,
`line3DChart`'s ribbons, a clustered `area3DChart`'s slabs and a surface's mesh, each with
its floor and two walls, laid out through `c:view3D`'s camera. A `pie3DChart`, a stacked
`area3DChart` and anything under a perspective camera (`c:rAngAx="0"`, or no `c:view3D` at
all) **draw flat and say so** — one `chart-3d-flattened` warning each — with their data,
categories, axis and legend in full. A 3-D value axis is **not** padded the way a flat one
is — measured, and it is the difference between the 0–50 by 5 PowerPoint draws and the
0–60 by 10 the flat rule asks for — so the numbers on the axis are PowerPoint's own even
where the picture is not.

**SmartArt** renders from the DrawingML drawing PowerPoint caches beside the diagram —
shapes, text, fills and geometry, each label placed by its own `dsp:txXfrm`. Where that
cache is missing or empty there is nothing to draw, and the renderer says so rather than
leaving a blank rectangle unexplained.

**EMF/WMF pictures** render the preview Office embeds in the metafile. A DIB preview
needs nothing extra; a PDF preview needs `pptx2svg[metafile]`.

## Where inherited text formatting comes from

A run that states no colour, size or face takes each, property by property, from the
first layer below it that states one. Measured on PowerPoint 16.106 for Mac (round 15,
23 October 2026: made-up decks giving each candidate layer its own colour and size, read
from the slide show on screen, the PDF export and PowerPoint's own `font color`, which
agreed on every probe; PowerPoint's resave writes no colour, so it is resolved when drawn):

| Text in | Layers, nearest first |
|---|---|
| a text box, an autoshape, a group's member | run, paragraph (`a:pPr/a:defRPr`), the shape's `a:lstStyle` at the paragraph's level, the shape's `p:style/a:fontRef`, the presentation's `defaultTextStyle` at the paragraph's level (`lvl1pPr` for level 0); then black, 18 pt |
| a placeholder | run, paragraph, its own list style, the layout's and the master's placeholder, the master's `titleStyle`/`bodyStyle`/`otherStyle` by type, `defaultTextStyle` |
| a table cell | run, paragraph, its list style, the table style's `a:tcTxStyle`, the master's `otherStyle`, `defaultTextStyle` |

For the first row: the master's `otherStyle` is **not** on the way (a `defaultTextStyle`
whose levels say white draws a text box white over an `otherStyle` that says `tx1`), and
neither a list style's nor `defaultTextStyle`'s `defPPr` stands in for a level the style
does not have. A theme colour goes through the slide's colour map (master, layout, slide
override). The placeholder and table rows were not re-measured and keep their order. The
order is `ooxml_common.drawingml.text_chain`, which pptx-agent reports colours from.

## Not rendered

Four things, and each of them is a real gap rather than a rough edge:

- **3-D effects, bevels and reflections.** `a:scene3d`, `a:sp3d` and `a:reflection` are
  ignored; the shape draws flat. This is the one item on the list that passes silently.
- **SmartArt with no cached drawing.** There is no diagram layout engine here, and
  implementing `dgm:layoutDef` is a project in its own right. Such a frame draws nothing
  and warns `diagram-no-cached-drawing`. This is common in older files: of 46 real
  SmartArt decks measured, 13 carried a drawing with shapes in it and 33 did not.
- **Vector EMF/WMF content.** Only the embedded preview is read; the metafile's own
  drawing records are not interpreted, so a metafile without a preview draws a
  placeholder and warns. `ConvertOptions(metafile_converter=...)` is the hook for
  shelling out to Inkscape or `libemf2svg` if you need the vectors.
- **Every chart in the ChartEx family.** Office 2016's newer types (treemap, sunburst,
  histogram, box-and-whisker, waterfall, funnel, map) live in a `cx:chartSpace` part in a
  different namespace with a different data model; they draw an empty frame and warn
  `chart-unsupported-type`. Every `c:*Chart` group ECMA-376 defines *is* drawn, the
  surface included. A combo chart draws whichever of its groups is a type we know and
  ignores the others.

See [ROADMAP.md](../ROADMAP.md) for what closing each of these involves.


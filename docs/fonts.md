# Fonts: the summary

> **[FONTS.md](../FONTS.md) answers one question end to end: *the deck names font X — will it
> be drawn correctly, and if not, what do I do?*** It has a scan table keyed on font name,
> all five ways a face gets drawn, and the escape hatches with their trap. What follows here
> is the summary.

Office's typefaces are proprietary and cannot be redistributed. Layout is therefore
computed from the advance widths of open fonts built to match them, and **those same
files are what the rasteriser draws with** — measuring with one face and drawing with
another is the failure this whole subsystem exists to prevent. resvg does not warn when
it substitutes: rendering one string in Calibri, Carlito, Aptos, Noto Sans JP and Lato on
a machine with none of them produces five byte-identical PNGs.

| Office face        | Drawn with     | Debian package              | Licence     |
| ------------------ | -------------- | --------------------------- | ----------- |
| Calibri            | Carlito        | `fonts-crosextra-carlito`   | SIL OFL 1.1 |
| Arial, Helvetica   | Arimo          | `fonts-croscore`            | SIL OFL 1.1 |
| Times New Roman    | Tinos          | `fonts-croscore`            | SIL OFL 1.1 |
| Courier New        | Cousine        | `fonts-croscore`            | SIL OFL 1.1 |
| Cambria            | Caladea †      | `fonts-crosextra-caladea`   | SIL OFL 1.1 |
| Aptos †            | Carlito †      | —                           | —           |
| Japanese Gothic    | Noto Sans JP † | `fonts-noto-cjk`            | SIL OFL 1.1 |
| Lato, Raleway      | themselves     | `fonts-lato`, — ‡           | SIL OFL 1.1 |

The first four are exact: measured character by character against the copies Office
installs, the advance widths match to the unit, so substituting them changes glyph shapes
and nothing else — not one line break moves.

**The right-hand column is also a left-hand column.** A deck may name Carlito, Caladea,
Arimo, Tinos or Cousine directly — LibreOffice ships the first two as *its* Calibri and
Cambria substitutes, so anything round-tripped through it does, and Arimo/Tinos/Cousine
are the Chrome OS core fonts that Debian packages as `fonts-croscore`. Each resolves to
itself and reports `exact`. So do the three Liberation aliases: Liberation Sans → Arimo,
Liberation Serif → Tinos, Liberation Mono → Cousine, which are the same designs under
another name.

† **Approximate, and reported as such.** Caladea is universally described as
metric-compatible with Cambria; measured, it runs 4.5 % narrow. **Aptos**, Microsoft's
Office default since 2023, has no open clone at all — we measure it with its own widths
(so line breaks match PowerPoint) and draw it with Carlito, whose widths sit closest of
anything shippable, −3.6 % on a representative sentence. Noto Sans JP is not
metric-compatible with MS Gothic or Meiryo either; nothing is, and a Gothic standing in
for a Gothic beats a Latin fallback. The measurements behind all three, and what to do
about each, are in [FONTS.md](../FONTS.md#the-five-ways-a-font-gets-drawn).

‡ Debian does not package Raleway. Use the pip bundle, or Google Fonts.

## Fonts the deck brought with it

A deck saved with *Embed fonts in the file* carries its typefaces in `ppt/fonts/*.fntdata`,
and pptx2svg reads them. They are used for **both measurement and drawing**, so an
embedded face lays out at its own advance widths rather than at a substitute's — which is
the entire point, and the half that handing font files to the rasteriser alone would miss.

That beats any bundle, and not by a little. Template vendors pick arbitrary Google Fonts,
of which there are roughly 1,800 families; two commercial templates measured here embed
Anton, Arimo, Literata, Merriweather Sans, Merriweather Sans Light and Inclusive Sans, and
every one of them used to report "no substitute known; widths guessed" while sitting
inside the file being rendered. Scored against PowerPoint's own PDF export of those two
decks, reading the embedded fonts raised SSIM on all twelve slides measured — the largest
by 0.123, the mean by 0.058.

It also beats a bundled face of the *same name*: `real-basic-theme.pptx` embeds Raleway
4.026, 338 of 340 advance widths differ from the release the bundle ships, and the same
string measures 3.9 % apart between them. The author laid the deck out with the file
inside it, so that is the file to measure with.

Three things this does not do:

- **Use a font whose licence refuses.** Every face is checked against its OS/2 `fsType`
  before it is touched — in the EOT header and again in the decoded font, because the
  first is written by the embedder and the second by the foundry. A restricted-licence
  face is refused with a `font-embedded-restricted` warning naming the restriction, and
  the deck falls back to substitution. The rule is LibreOffice's
  (`EmbeddedFontsHelper::sufficientTTFRights`).
- **Fail.** A payload that cannot be decoded warns `font-embedded-undecodable` and the
  deck renders as it did before.
- **Cost nothing.** Decoding is 0.15–1.1 s per face, and only the families a slide
  actually asks for are decoded. `--no-embedded-fonts` (or
  `ConvertOptions(use_embedded_fonts=False)`) turns it off.

No extra is needed: the MicroType Express decoder is pure Python and lives in the core.

## Will my deck render faithfully?

```bash
pptx2svg fonts                       # what the bundle covers
pptx2svg fonts --check deck.pptx     # exit 1 if this deck cannot be drawn faithfully
```

```
mode:   bundled from .../site-packages/pptx2svg_fonts/files
        this machine's own fonts are ignored, so output is reproducible
        families: Arimo, Caladea, Carlito, Cousine, Lato, Noto Sans JP, Raleway, Tinos

face                       verdict      drawn with     note
Arial                      compatible   Arimo          Arimo has Arial's advance widths
Aptos                      approximate  Carlito        measured as Aptos, drawn as Carlito; no metric-compatible clone exists
Aptos Display              approximate  Carlito        measured as Aptos Display, drawn as Carlito; no metric-compatible clone exists

2 of 3 faces will not be drawn at the widths they were measured at.
```

A face the deck embeds grades `exact` — "drawn with the face the deck embedded" — because
the layout was measured from the very file the rasteriser is handed. That is a stronger
guarantee than any clone offers, so it wins even where a substitute exists.

`exact` and `compatible` are faithful; `approximate` and `missing` are not, and `--check`
exits non-zero on them so a deck that cannot be rendered faithfully fails a build instead
of shipping wrong pixels. The same information reaches library callers as
`font-substituted` warnings on `ConvertOptions.warnings`. What each grade means and what to
do about a bad one: [FONTS.md](../FONTS.md#the-four-grades).

## Without the bundle

`pptx2svg` alone renders PNGs with the host's fonts and emits one
`font-bundle-missing` warning saying so. That is a deliberate downgrade, not a silent
one — output is then whatever the machine happens to have installed.

## Escape hatches

```python
# The host's fonts as well as the bundle, for a deck naming a face we do not carry
convert_pptx_to_png("deck.pptx", skip_system_fonts=False)

# Your own faces only
convert_pptx_to_png("deck.pptx", font_dirs=["./corporate-fonts"], use_bundled_fonts=False)

# Layout measured from specific files on this machine
from pptx2svg import ConvertOptions, FontToolsTextMeasurer
measurer = FontToolsTextMeasurer({"Calibri": "/path/to/Calibri.ttf"})
convert_pptx_to_svg("deck.pptx", ConvertOptions(measurer=measurer))

# Map a face to a substitute you do have
ConvertOptions(font_mapping={"Helvetica Neue": "Inter"})
```

On the command line: `--system-fonts`, `--no-bundled-fonts`, `--font-dir DIR`.

**Your own folder is measured and drawn.** `font_dirs` (on `ConvertOptions` or
`convert_pptx_to_png`) and `--font-dir` reach measurement as well as the rasteriser: a
face there is measured from its file unless the tables measure its family as itself
(Aptos, Calibri), and grades `exact`. Without them, **`OOXML_FONT_DIRS`** (folders
separated by `os.pathsep`) is read -- the one variable pptx2svg and docx2svg share.
Precedence: the explicit argument, then the variable; added to the system's folders and
searched before them. `font_files` still reaches the rasteriser alone. See
[FONTS.md](../FONTS.md#your-own-folder-measured-and-drawn).

## Debian and other Linux hosts

Two routes, and only one of them is reproducible.

**The bundle (recommended).** `pip install 'pptx2svg[fonts]'` pins the exact font files.
Two machines running the same version produce the same pixels, which is the only way to
get that guarantee.

**System packages.** `tools/install-fonts-debian.sh` installs the same designs from apt
(`fonts-croscore`, `fonts-crosextra-carlito`, `fonts-crosextra-caladea`,
`fonts-liberation2`, `fonts-noto-cjk`, `fonts-lato`) and verifies each family with
`fc-list`. Render with `--system-fonts` to use them. This is fine for correct *layout* —
Liberation Sans/Serif/Mono are derived from Arimo/Tinos/Cousine and measure identically,
checked character by character — but apt gives you whatever version the distribution
shipped, so two machines on different releases can still differ.

The same script reaches the faces no bundle can legally contain, behind explicit flags:

```bash
tools/install-fonts-debian.sh                  # open substitutes from apt
tools/install-fonts-debian.sh --clones         # + the rest of the open metric clones
tools/install-fonts-debian.sh --aptos          # + Aptos, from Microsoft's own download
tools/install-fonts-debian.sh --mscorefonts    # + Arial, Georgia, Verdana … (EULA)
tools/install-fonts-debian.sh --ppviewer       # + Calibri, Cambria and the ClearType set
tools/install-fonts-debian.sh --office-dir auto   # + everything else PowerPoint ships
```

`--clones` needs no licence from anyone. It adds `fonts-urw-base35`, `fonts-texgyre`,
`fonts-liberation-sans-narrow` and two faces Debian does not package (Comic Relief and
Symbol Neu, both hash-pinned), which between them cover Book Antiqua, Palatino Linotype,
Century Schoolbook, Century, Century Gothic, Bookman Old Style, Arial Narrow, Monotype
Corsiva, Symbol, Monotype Sorts and Comic Sans MS. Each of those pairings was measured
character by character against the PowerPoint face it substitutes for; the numbers are in
`ROADMAP.md` under *The clone landscape*. These are not in the wheel because the wheel
*redistributes* and their licences (AGPL and GPL-2, both with document-embedding
exceptions) do not permit that; installing them from Debian's archive is a different act.
Installing a clone under its own name does not by itself make pptx2svg use it — see
[FONTS.md](../FONTS.md#4-a-clone-exists-but-is-not-bundled) for the two lines it needs.

`--mscorefonts` installs Debian's `ttf-mscorefonts-installer` from **contrib**, which
presents Microsoft's EULA through debconf — use `--accept-eula` to preseed it for a fleet.
`--ppviewer` extracts the ClearType set from the PowerPoint Viewer installer, the route
documented at [wiki.debian.org/ppviewerFonts](https://wiki.debian.org/ppviewerFonts)
(hash-pinned; needs `cabextract`). **These are non-free Microsoft fonts — you must already
hold a licence to use them.** Nothing extracted is ever committed or shipped.

`--office-dir PATH` is for organisations that hold Office licences and want the ~150
families nothing above reaches — Gill Sans MT, Rockwell, Franklin Gothic, Tw Cen MT,
Perpetua, Garamond, the Lucida family, the CJK and Indic faces. **It downloads nothing.**
It copies from a licensed Microsoft Office installation you point it at (`--office-dir
auto` probes the usual places, including a mounted macOS `PowerPoint.app/Contents/
Resources/DFonts` and `/mnt/c/Windows/Fonts`) into a directory fontconfig indexes, and
then names what landed. Whether those files may be copied onto a given machine is your
organisation's licensing decision; the script makes no part of it for you.

Back to the [README](../README.md).

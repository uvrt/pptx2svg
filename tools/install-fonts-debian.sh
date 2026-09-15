#!/usr/bin/env bash
#
# Install the fonts a Debian box needs to rasterise PPTX decks the way PowerPoint does.
#
# Five tiers, in decreasing order of how freely they can be used:
#
#   1. OPEN SUBSTITUTES (default).  Carlito, Caladea, Arimo, Tinos, Cousine, Noto CJK --
#      the same faces `pip install 'pptx2svg[fonts]'` bundles, from apt.  Redistributable,
#      metric-compatible with the Office faces -- except Caladea and Noto CJK, which
#      are not, and which `pptx2svg fonts --check` grades `approximate` for that
#      reason.  See FONTS.md.  Enough for correct *layout* everywhere else.
#
#   2. THE REST OF THE OPEN CLONES (--clones).  URW base-35, TeX Gyre, Liberation Sans
#      Narrow, Comic Relief and Symbol Neu.  Also open, also metric-compatible, and worth
#      roughly a dozen more PowerPoint families -- Book Antiqua, Palatino Linotype,
#      Century Schoolbook, Century, Century Gothic, Bookman Old Style, Arial Narrow,
#      Monotype Corsiva, Symbol, Monotype Sorts, Comic Sans MS.  Not in tier 1 only
#      because the pip bundle cannot ship all of them; see ROADMAP "The clone landscape".
#
#   3. APTOS (--aptos).  Microsoft's Office default since 2023 has no open clone, so
#      nothing in tiers 1-2 draws it correctly.  Microsoft publishes it for download under
#      its own terms; you accept those, not us.
#
#   4. THE MICROSOFT CORE FONTS (--mscorefonts, --ppviewer).  Arial, Times New Roman,
#      Courier New, Georgia, Verdana, Trebuchet MS, Comic Sans MS, Impact, Webdings and
#      Andale Mono from Debian's contrib packaging of Microsoft's own EULA'd download;
#      Calibri, Cambria, Candara, Consolas, Constantia, Corbel and Meiryo out of the
#      PowerPoint Viewer cabinet.  Non-free.  You must already hold a licence.
#
#   5. EVERYTHING ELSE POWERPOINT SHIPS (--office-dir PATH).  Gill Sans MT, Rockwell,
#      Franklin Gothic, Tw Cen MT, Perpetua, Garamond, the Lucida family, the CJK and
#      Indic faces -- about 186 families in all, and the ~150 of them that no tier above
#      reaches.  There is no download here and there never will be: this tier copies from
#      a licensed Microsoft Office installation that YOU point it at.  If your
#      organisation holds Office licences, it already has these files; this is the step
#      that puts them where fontconfig can see them.
#
# Why this script may use fonts the pip bundle may not.  `pptx2svg[fonts]` *redistributes*
# font files inside a wheel, so it can only carry licences that permit that -- OFL,
# Apache-2.0.  This script *installs* from the distribution's own archive or from files
# you already hold, and redistributes nothing.  That is why `fonts-urw-base35` (AGPL with
# a PostScript/PDF-embedding exception) and `fonts-liberation-sans-narrow` (GPL-2 with a
# font exception) are fine here and are deliberately absent from the bundle.
#
# Reproducibility note.  Only the pip bundle pins exact files.  apt gives you whatever
# version the distribution shipped, so two machines on different Debian releases can
# still differ by a hinting change.  If you need byte-identical PNGs, install
# `pptx2svg[fonts]` and let pptx2svg use it -- that is the default -- and treat this
# script as the way to get the *licensed* faces that no bundle can legally contain.
#
# Usage:
#   tools/install-fonts-debian.sh                          # open substitutes only
#   tools/install-fonts-debian.sh --clones                 # + the rest of the open set
#   tools/install-fonts-debian.sh --aptos                  # + Aptos from Microsoft
#   tools/install-fonts-debian.sh --mscorefonts            # + Arial/Georgia/... (EULA)
#   tools/install-fonts-debian.sh --ppviewer               # + Calibri/Cambria (non-free)
#   tools/install-fonts-debian.sh --office-dir auto        # + everything Office ships
#   tools/install-fonts-debian.sh --office-dir /mnt/office/DFonts
#   tools/install-fonts-debian.sh --all --system           # every downloadable tier
#
set -euo pipefail

# --------------------------------------------------------------------------------------
# Arguments
# --------------------------------------------------------------------------------------

WANT_OPEN=1
WANT_CLONES=0
WANT_APTOS=0
WANT_MSCORE=0
WANT_PPVIEWER=0
OFFICE_DIR=""
# XDG user fonts by default: no root needed, and nothing licensed lands in a place a
# container image build would pick up by accident.  --system is for servers.
FONTDIR="${HOME}/.local/share/fonts"
SYSTEM=0
# Preseeding the mscorefonts EULA has to be a deliberate act, not a side effect of -y.
ACCEPT_EULA="${MSCOREFONTS_ACCEPT_EULA:-0}"

while [ $# -gt 0 ]; do
  case "$1" in
    --clones)      WANT_CLONES=1 ;;
    --aptos)       WANT_APTOS=1 ;;
    --mscorefonts) WANT_MSCORE=1 ;;
    --ppviewer)    WANT_PPVIEWER=1 ;;
    --accept-eula) ACCEPT_EULA=1 ;;
    --office-dir)  shift; OFFICE_DIR="${1:-}" ;;
    # --all is every tier this script can fetch on its own.  It deliberately does NOT
    # include --office-dir, which needs a path only the operator knows.
    --all)         WANT_CLONES=1; WANT_APTOS=1; WANT_MSCORE=1; WANT_PPVIEWER=1 ;;
    --no-open)     WANT_OPEN=0 ;;
    --system)      SYSTEM=1; FONTDIR="/usr/local/share/fonts" ;;
    --dir)         shift; FONTDIR="$1" ;;
    -h|--help)     sed -n '2,58p' "$0"; exit 0 ;;
    *)             echo "unknown option: $1" >&2; exit 2 ;;
  esac
  shift
done

SUDO=""
if [ "$SYSTEM" = "1" ] && [ "$(id -u)" != "0" ]; then SUDO="sudo"; fi

say() { printf '\n== %s\n' "$*"; }

# One scratch directory and one trap.  The previous version set an EXIT trap per tier,
# and the second silently replaced the first, so the Aptos temp directory leaked whenever
# --ppviewer also ran.
TMPROOT="$(mktemp -d)"
trap 'rm -rf "$TMPROOT"' EXIT

need() { command -v "$1" >/dev/null || { echo "missing: $1 (apt install $2)" >&2; exit 1; }; }

# Three tiers can reach for apt and any of them may be the first to do so -- `--no-open
# --clones` and a bare `--mscorefonts` both used to run against whatever index the box
# happened to have.  Refresh once, on demand, and never twice.
APT_REFRESHED=0
apt_refresh() {
  [ "$APT_REFRESHED" = "1" ] && return 0
  $SUDO apt-get update -qq
  APT_REFRESHED=1
}

# Fetch $1 to $2 and check it against sha256 $3.  An empty $3 prints the digest and
# carries on, which is how you bootstrap a new pin; it is never silently skipped.
fetch_pinned() {
  local url="$1" dest="$2" want="$3"
  curl -fsSL "$url" -o "$dest"
  local got; got="$(sha256sum "$dest" | cut -d' ' -f1)"
  if [ -n "$want" ] && [ "$got" != "$want" ]; then
    echo "  $url" >&2
    echo "  expected sha256 $want" >&2
    echo "  got      sha256 $got" >&2
    echo "  refusing to install a payload that is not the one this script was written against." >&2
    exit 1
  fi
  [ -n "$want" ] || echo "  note: unpinned; sha256 is $got"
}

# --------------------------------------------------------------------------------------
# 1. Open, redistributable substitutes
# --------------------------------------------------------------------------------------

if [ "$WANT_OPEN" = "1" ]; then
  say "Open metric-compatible substitutes (apt)"
  # fonts-croscore is Arimo/Tinos/Cousine (Arial/Times/Courier); the two crosextra
  # packages are Carlito (Calibri) and Caladea (Cambria); fonts-liberation2 is the same
  # designs again under different names and is included because plenty of decks and
  # themes name "Liberation Sans" directly.
  apt_refresh
  # Package names checked against sources.debian.org rather than assumed.  Note what
  # is *not* here: there is no fonts-raleway in Debian.  Raleway comes from the pip
  # bundle (`pip install 'pptx2svg[fonts]'`) or from Google Fonts by hand; a deck that
  # names it gets the generic sans otherwise, and `pptx2svg fonts --check` will say so.
  $SUDO apt-get install -y --no-install-recommends \
    fonts-croscore \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-liberation2 \
    fonts-noto-cjk \
    fonts-lato \
    fontconfig
fi

mkdir -p "$FONTDIR" 2>/dev/null || $SUDO mkdir -p "$FONTDIR"

# --------------------------------------------------------------------------------------
# 2. The rest of the open clone set
# --------------------------------------------------------------------------------------

# Every face in this tier was measured against the PowerPoint copy it substitutes for,
# character by character; the numbers are in ROADMAP "The clone landscape".  Summary of
# what each package buys, so nobody has to rediscover it:
#
#   fonts-urw-base35              P052 -> Book Antiqua and Palatino Linotype (0 of 74
#                                 letters/digits/punctuation differ); C059 -> Century
#                                 Schoolbook and Century (0/74); URW Gothic -> Century
#                                 Gothic (0/74); URW Bookman -> Bookman Old Style (1/74,
#                                 'Q'); Z003 -> Monotype Corsiva (0/74); Standard Symbols
#                                 PS -> Symbol (0/188); D050000L -> Monotype Sorts
#                                 (9/202); Nimbus Sans Narrow -> Arial Narrow (0/95).
#   fonts-texgyre                 The same designs under a permissive licence, but Debian
#                                 ships only Bonum, Heros, HerosCn and Pagella -- no
#                                 Schola and no Adventor -- which is why urw-base35 is
#                                 the package that actually covers the set.
#   fonts-liberation-sans-narrow  Arial Narrow, 0 of 95 in all four cuts.
#   Comic Relief                  Comic Sans MS, 0 of 95 in both cuts.  Not in Debian.
#   Symbol Neu                    Symbol, 0 of 188.  Google dropped it after croscore
#                                 1.23.0, so it comes out of the archived tarball.

COMICRELIEF_COMMIT="${COMICRELIEF_COMMIT:-993d03b988291b8609cd3fc24e10b6a982c79de9}"
COMICRELIEF_BASE="https://raw.githubusercontent.com/google/fonts/${COMICRELIEF_COMMIT}/ofl/comicrelief"
COMICRELIEF_REGULAR_SHA256="${COMICRELIEF_REGULAR_SHA256:-9ef4958ab06385d91f635f83a15569786b243f9d010cfb3f9a5cdda593c7bc22}"
COMICRELIEF_BOLD_SHA256="${COMICRELIEF_BOLD_SHA256:-33aee007212b696afd57d50c901135cd723cd188f7ad98f3a1c572ea64e8f0d5}"

# chromeos-localmirror keeps the old croscore releases; 1.23.0 is the last one with
# SymbolNeu.ttf in it.  Pinned to the tarball, not to the file inside it, so a repacked
# archive fails loudly.
CROSCORE_URL="${CROSCORE_URL:-https://storage.googleapis.com/chromeos-localmirror/distfiles/croscorefonts-1.23.0.tar.gz}"
CROSCORE_SHA256="${CROSCORE_SHA256:-b469b5457b093a9d8878ef6ff6868f54e258441b88983b1866f64c8995584b4c}"

if [ "$WANT_CLONES" = "1" ]; then
  say "The rest of the open metric-compatible clones (apt)"
  apt_refresh
  $SUDO apt-get install -y --no-install-recommends \
    fonts-urw-base35 \
    fonts-texgyre \
    fonts-liberation-sans-narrow

  say "Comic Relief and Symbol Neu (open, but not packaged by Debian)"
  need curl curl
  need sha256sum coreutils
  need tar tar
  target="$FONTDIR/pptx2svg-clones"
  $SUDO mkdir -p "$target"

  for cut in Regular:"$COMICRELIEF_REGULAR_SHA256" Bold:"$COMICRELIEF_BOLD_SHA256"; do
    name="ComicRelief-${cut%%:*}.ttf"
    fetch_pinned "$COMICRELIEF_BASE/$name" "$TMPROOT/$name" "${cut##*:}"
    $SUDO cp "$TMPROOT/$name" "$target/$name"
  done
  # The OFL obliges you to carry the licence with the files.  Doing that is one line, and
  # a font directory with no licence text in it is how an obligation gets lost.
  curl -fsSL "$COMICRELIEF_BASE/OFL.txt" -o "$TMPROOT/ComicRelief-OFL.txt" \
    && $SUDO cp "$TMPROOT/ComicRelief-OFL.txt" "$target/ComicRelief-OFL.txt"

  fetch_pinned "$CROSCORE_URL" "$TMPROOT/croscore.tar.gz" "$CROSCORE_SHA256"
  tar xzf "$TMPROOT/croscore.tar.gz" -C "$TMPROOT" croscorefonts-1.23.0/SymbolNeu.ttf
  $SUDO cp "$TMPROOT/croscorefonts-1.23.0/SymbolNeu.ttf" "$target/SymbolNeu.ttf"
fi

# --------------------------------------------------------------------------------------
# 3. Aptos, from Microsoft's own download
# --------------------------------------------------------------------------------------

# Microsoft's own download, taken from the versioned download.microsoft.com path rather
# than the aka.ms short link.  The short link is not a stable source -- at the time of
# writing https://aka.ms/AptosFonts redirects to Bing rather than to a font bundle -- and
# an unversioned redirect is exactly what cannot be pinned.  This path can be, and is:
# the digest below is of the 2,979,784-byte zip fetched from it, which unpacks to 34 flat
# .ttf files (Aptos, Display, Narrow, Mono, Serif, each in its weights and italics).
APTOS_URL="${APTOS_URL:-https://download.microsoft.com/download/8/6/0/860a94fa-7feb-44ef-ac79-c072d9113d69/Microsoft%20Aptos%20Fonts.zip}"
# Pinned, so a changed payload is a hard failure rather than a silent substitution.
# Override both of these together if Microsoft publishes a newer bundle and you have
# verified it yourself; clearing APTOS_SHA256 downgrades to the size check below.
APTOS_SHA256="${APTOS_SHA256:-6528fd120e719a9f985e94214eca6887d1653b88456916a792a630b02e95b025}"
# The fallback when the hash is deliberately cleared.  Not integrity, but it still catches
# the failure that actually happens: the URL serving an HTML error page instead of a zip.
APTOS_EXPECTED_BYTES="${APTOS_EXPECTED_BYTES:-2979784}"

if [ "$WANT_APTOS" = "1" ]; then
  say "Aptos (Microsoft, proprietary -- you accept Microsoft's terms)"
  need curl curl
  need unzip unzip

  curl -fsSL "$APTOS_URL" -o "$TMPROOT/aptos.zip"
  if [ -n "$APTOS_SHA256" ]; then
    echo "$APTOS_SHA256  $TMPROOT/aptos.zip" | sha256sum -c - || {
      echo "Aptos download does not match the pinned hash; refusing to install." >&2
      exit 1
    }
  else
    got="$(wc -c < "$TMPROOT/aptos.zip" | tr -d ' ')"
    if [ -n "$APTOS_EXPECTED_BYTES" ] && [ "$got" != "$APTOS_EXPECTED_BYTES" ]; then
      echo "Aptos download is $got bytes, expected $APTOS_EXPECTED_BYTES." >&2
      echo "Microsoft has changed what this URL serves.  Re-verify before installing:" >&2
      echo "  sha256 $(sha256sum "$TMPROOT/aptos.zip" | cut -d' ' -f1)" >&2
      echo "Then set APTOS_SHA256 (and APTOS_EXPECTED_BYTES) and run again." >&2
      exit 1
    fi
    echo "note: APTOS_SHA256 is unset, so only the size was checked ($got bytes)."
    echo "      Its sha256 is: $(sha256sum "$TMPROOT/aptos.zip" | cut -d' ' -f1)"
    echo "      Set APTOS_SHA256 to that value to pin it for your fleet."
  fi
  $SUDO unzip -jo "$TMPROOT/aptos.zip" '*.ttf' -d "$FONTDIR"
fi

# --------------------------------------------------------------------------------------
# 4a. The Microsoft core fonts, via Debian contrib
# --------------------------------------------------------------------------------------

# ttf-mscorefonts-installer (source package msttcorefonts, section contrib) downloads
# Microsoft's own "core fonts for the web" and presents Microsoft's EULA through debconf.
# It is the packaged, licence-respecting route to Arial, Times New Roman, Courier New,
# Georgia, Verdana, Trebuchet MS, Comic Sans MS, Impact, Webdings and Andale Mono.
#
# Two of those matter to pptx2svg beyond the obvious: Verdana, which has no open clone at
# all (and is metrically identical to MS Reference Sans Serif, so it answers two names),
# and Georgia, which Gelasio clones exactly but which the bundle does not carry.
if [ "$WANT_MSCORE" = "1" ]; then
  say "Microsoft core fonts (NON-FREE -- Microsoft's EULA applies)"
  echo "Package: ttf-mscorefonts-installer, from Debian contrib."
  echo "This script does not grant you any right to these fonts."
  apt_refresh
  # contrib is not enabled on every Debian install, and the failure mode without it is a
  # bare "unable to locate package", which sends people looking in the wrong place.
  if ! apt-cache policy ttf-mscorefonts-installer 2>/dev/null | grep -q 'Candidate: [0-9]'; then
    echo "ttf-mscorefonts-installer is not available." >&2
    echo "It lives in contrib; enable it in /etc/apt/sources.list (add 'contrib' to the" >&2
    echo "component list for your suite), run apt-get update, and try again." >&2
    exit 1
  fi
  if [ "$ACCEPT_EULA" = "1" ]; then
    # Preseeding is how a fleet installs this unattended.  It is behind an explicit flag
    # because accepting a licence on someone's behalf should never be a default.
    need debconf-set-selections debconf-utils
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" \
      | $SUDO debconf-set-selections
    $SUDO env DEBIAN_FRONTEND=noninteractive apt-get install -y ttf-mscorefonts-installer
  else
    echo "debconf will show you Microsoft's EULA; accept it there, or re-run with"
    echo "--accept-eula (or MSCOREFONTS_ACCEPT_EULA=1) to preseed it for a fleet."
    $SUDO apt-get install -y ttf-mscorefonts-installer
  fi
fi

# --------------------------------------------------------------------------------------
# 4b. Calibri and Cambria, via the PowerPoint Viewer installer
# --------------------------------------------------------------------------------------

# https://wiki.debian.org/ppviewerFonts -- the ClearType collection that shipped with
# Office 2007 and Windows 7, still distributed inside the PowerPoint Viewer setup.  Hash
# verified against the wiki.
PPVIEWER_URL="https://archive.org/download/PowerPointViewer_201801/PowerPointViewer.exe"
PPVIEWER_SHA256="249473568eba7a1e4f95498acba594e0f42e6581add4dead70c1dfb908a09423"

if [ "$WANT_PPVIEWER" = "1" ]; then
  say "Calibri, Cambria and friends (NON-FREE -- you must hold a licence)"
  echo "Source: $PPVIEWER_URL"
  echo "These are Microsoft fonts.  This script does not grant you any right to them."
  need cabextract cabextract
  need curl curl

  curl -fsSL "$PPVIEWER_URL" -o "$TMPROOT/PowerPointViewer.exe"
  echo "$PPVIEWER_SHA256  $TMPROOT/PowerPointViewer.exe" | sha256sum -c - || {
    echo "PowerPointViewer.exe does not match the pinned hash; refusing to extract." >&2
    exit 1
  }

  # Two nested cabinets: the installer carries ppviewer.cab, which carries the fonts.
  # -L lowercases the names, which matters because the cabinet stores them shouting
  # (CALIBRI.TTF) and fontconfig does not care but humans reading ls -1 do.
  ( cd "$TMPROOT" && cabextract -q PowerPointViewer.exe -F ppviewer.cab )
  target="$FONTDIR/ppviewer"
  $SUDO mkdir -p "$target"
  ( cd "$TMPROOT" && $SUDO cabextract -q -L -F '*.TT[CF]' -d "$target" ppviewer.cab )
  echo "extracted:"
  ls -1 "$target"
fi

# --------------------------------------------------------------------------------------
# 5. Everything else PowerPoint ships, from a licensed Office installation
# --------------------------------------------------------------------------------------

# The other ~150 families.  Tiers 1-4 between them reach Arial, Times New Roman, Courier
# New, Calibri, Cambria, Candara, Consolas, Constantia, Corbel, Aptos, Georgia, Verdana,
# Trebuchet MS, Comic Sans MS, Meiryo and the clone set -- which leaves Gill Sans MT,
# Rockwell, Franklin Gothic, Tw Cen MT, Perpetua, Garamond, Calisto MT, Bell MT, Goudy Old
# Style, News Gothic MT, Eurostile, the whole Lucida family, the Monotype display faces,
# the CJK set, the Indic set, Thai, Hebrew and Ethiopic without any substitute at all.  A
# deck naming one of those gets guessed advance widths today.
#
# There is exactly one legitimate source for them and it is not a download: an Office
# installation your organisation is licensed for.  Point this at it.
#
#   macOS host, fonts shared into a container or copied to the Debian box:
#     /Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts
#   Windows host:
#     C:\Windows\Fonts  (and the Office private font directory, if your deployment uses one)
#
# Nothing is fetched here, nothing is committed, and the files stay yours.  What this tier
# does is copy them somewhere fontconfig indexes and then tell you what landed.

office_probe() {
  # Ordered most-specific first: an Office DFonts directory is a better answer than a
  # whole Windows font directory, which also contains faces Office does not ship.
  local candidates=(
    "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts"
    "/mnt/c/Windows/Fonts"
    "/media/office/DFonts"
    "$HOME/office-fonts"
  )
  local d
  for d in "${candidates[@]}"; do
    [ -d "$d" ] && { printf '%s\n' "$d"; return 0; }
  done
  return 1
}

if [ -n "$OFFICE_DIR" ]; then
  if [ "$OFFICE_DIR" = "auto" ]; then
    if ! OFFICE_DIR="$(office_probe)"; then
      echo "--office-dir auto found no Office font directory in the usual places." >&2
      echo "Pass the path explicitly: --office-dir /path/to/DFonts" >&2
      exit 1
    fi
    echo "auto-detected Office fonts at: $OFFICE_DIR"
  fi

  say "Everything else PowerPoint ships (NON-FREE -- from YOUR licensed installation)"
  [ -d "$OFFICE_DIR" ] || { echo "not a directory: $OFFICE_DIR" >&2; exit 1; }

  # -maxdepth 1 because Office keeps these flat, and because a recursive sweep of, say,
  # C:\Windows would pick up far more than was asked for.  No bash arrays here on
  # purpose: `mapfile -d` needs bash 4.4, and this script is also useful run from a mac to
  # stage a directory before shipping it to the Debian box.
  office_find() {
    find "$OFFICE_DIR" -maxdepth 1 -type f \
      \( -iname '*.ttf' -o -iname '*.ttc' -o -iname '*.otf' -o -iname '*.otc' \) "$@"
  }
  # Count first, so "0 fonts" is reported as the mistake it is rather than as a silent
  # successful no-op -- which is what pointing this at the wrong directory looks like.
  office_count="$(office_find -print | wc -l | tr -d ' ')"
  if [ "$office_count" = "0" ]; then
    echo "no font files in $OFFICE_DIR" >&2
    echo "Expected a flat directory of .ttf/.ttc files, such as an Office DFonts folder." >&2
    exit 1
  fi

  echo "These are Microsoft fonts from your own installation.  Copying them onto other"
  echo "machines is a licensing decision your organisation owns; this script does not make"
  echo "it for you and grants you no rights."
  echo "source: $OFFICE_DIR"
  echo "files:  $office_count"

  target="$FONTDIR/office"
  $SUDO mkdir -p "$target"
  # Copy the matched files only, rather than `cp -r dir/.`, so a source that is a whole
  # Windows font directory brings the fonts and leaves the .fon, .pfm and desktop.ini
  # beside them.  -print0/xargs -0 because Office font filenames contain spaces
  # ("Book Antiqua Bold Italic.ttf").
  office_find -print0 | $SUDO xargs -0 -I{} cp -f {} "$target/"
  echo "copied to: $target"
fi

# --------------------------------------------------------------------------------------
# Verify
# --------------------------------------------------------------------------------------

say "Rebuilding the font cache"
$SUDO fc-cache -f "$FONTDIR" >/dev/null || fc-cache -f "$FONTDIR" >/dev/null

say "Verifying"
# fc-list is the only honest check: it asks the same fontconfig database a rasteriser
# will.  Checking that a file exists on disk proves nothing -- it is finding the *family*
# that matters, and a family name is inside the file, not in its name.
status=0
check() {
  if fc-list : family | tr ',' '\n' | grep -qxF "$1"; then
    printf '  %-24s ok\n' "$1"
  else
    printf '  %-24s MISSING\n' "$1"
    status=1
  fi
}

if [ "$WANT_OPEN" = "1" ]; then
  # Raleway is deliberately absent: Debian does not package it.  See the apt call above.
  for family in Carlito Caladea Arimo Tinos Cousine "Noto Sans CJK JP" Lato; do
    check "$family"
  done
fi
if [ "$WANT_CLONES" = "1" ]; then
  # Family names read out of the files themselves rather than guessed from the package
  # names -- URW ships "P052", not "URW Palatino", and TeX Gyre spells the condensed cut
  # "TeX Gyre HerosCn" with no space.
  for family in P052 C059 "URW Gothic" "URW Bookman" "Nimbus Sans Narrow" \
                "Standard Symbols PS" D050000L Z003 \
                "TeX Gyre Pagella" "TeX Gyre Bonum" "TeX Gyre HerosCn" \
                "Liberation Sans Narrow" "Comic Relief" "Symbol Neu"; do
    check "$family"
  done
fi
[ "$WANT_APTOS" = "1" ] && check "Aptos"
# Verdana and Georgia are the two that no open face replaces, so they are the ones worth
# naming individually if this tier half-fails.
[ "$WANT_MSCORE" = "1" ] && for family in Arial "Times New Roman" "Courier New" Georgia \
                                          Verdana "Trebuchet MS" "Comic Sans MS" Impact; do
  check "$family"
done
# Calibri and Cambria are what pptx2svg maps; the other four come along for free and
# are worth naming so a missing one is visible rather than mysterious.
[ "$WANT_PPVIEWER" = "1" ] && for family in Calibri Cambria Candara Consolas Constantia Corbel Meiryo; do
  check "$family"
done
# A spread of the families that only tier 5 can supply: one Monotype text face, one slab,
# one grotesque, one geometric, one Lucida, one CJK.  If these are present the copy
# worked; checking all 150 would be noise.
[ -n "$OFFICE_DIR" ] && for family in "Gill Sans MT" Rockwell "Franklin Gothic Book" \
                                      "Tw Cen MT" "Book Antiqua" "Lucida Sans" "MS Gothic"; do
  check "$family"
done

say "Done"
echo "Point pptx2svg at these with:   pptx2svg deck.pptx -f png --font-dir $FONTDIR --system-fonts"
echo "Or check what a deck needs with: pptx2svg fonts --check deck.pptx --system-fonts"
exit $status

#!/usr/bin/env bash
#
# Install the fonts a Debian box needs to rasterise PPTX decks the way PowerPoint does.
#
# Three tiers, in decreasing order of how freely they can be used:
#
#   1. OPEN SUBSTITUTES (default).  Carlito, Caladea, Arimo, Tinos, Cousine, Noto CJK --
#      the same faces `pip install 'pptx2svg[fonts]'` bundles, from apt.  Redistributable,
#      metric-compatible with the Office faces, and enough for correct *layout*.
#
#   2. APTOS (--aptos).  Microsoft's Office default since 2023 has no open clone, so
#      nothing in tier 1 draws it correctly.  Microsoft publishes it for download under
#      its own terms; you accept those, not us.
#
#   3. THE CLEARTYPE COLLECTION (--ppviewer).  Extracted from the PowerPoint Viewer
#      installer, the route documented at https://wiki.debian.org/ppviewerFonts.  The
#      cabinet was opened and its contents listed rather than taken on trust; it carries
#      26 files -- Calibri, Cambria, Candara, Consolas, Constantia and Corbel in four
#      cuts each, plus Meiryo and Meiryo Bold, which is a real Japanese Gothic and more
#      than the wiki advertises.  These are non-free Microsoft fonts: you must already
#      hold a licence to use them.  Never the default, never committed, never shipped.
#
# Reproducibility note.  Only the pip bundle pins exact files.  apt gives you whatever
# version the distribution shipped, so two machines on different Debian releases can
# still differ by a hinting change.  If you need byte-identical PNGs, install
# `pptx2svg[fonts]` and let pptx2svg use it -- that is the default -- and treat this
# script as the way to get the *licensed* faces that no bundle can legally contain.
#
# Usage:
#   tools/install-fonts-debian.sh                       # open substitutes only
#   tools/install-fonts-debian.sh --aptos               # + Aptos from Microsoft
#   tools/install-fonts-debian.sh --ppviewer            # + Calibri/Cambria (non-free)
#   tools/install-fonts-debian.sh --all --system        # everything, system-wide
#
set -euo pipefail

# --------------------------------------------------------------------------------------
# Arguments
# --------------------------------------------------------------------------------------

WANT_APTOS=0
WANT_PPVIEWER=0
WANT_OPEN=1
# XDG user fonts by default: no root needed, and nothing licensed lands in a place a
# container image build would pick up by accident.  --system is for servers.
FONTDIR="${HOME}/.local/share/fonts"
SYSTEM=0

while [ $# -gt 0 ]; do
  case "$1" in
    --aptos)     WANT_APTOS=1 ;;
    --ppviewer)  WANT_PPVIEWER=1 ;;
    --all)       WANT_APTOS=1; WANT_PPVIEWER=1 ;;
    --no-open)   WANT_OPEN=0 ;;
    --system)    SYSTEM=1; FONTDIR="/usr/local/share/fonts" ;;
    --dir)       shift; FONTDIR="$1" ;;
    -h|--help)   sed -n '2,30p' "$0"; exit 0 ;;
    *)           echo "unknown option: $1" >&2; exit 2 ;;
  esac
  shift
done

SUDO=""
if [ "$SYSTEM" = "1" ] && [ "$(id -u)" != "0" ]; then SUDO="sudo"; fi

say() { printf '\n== %s\n' "$*"; }

# --------------------------------------------------------------------------------------
# 1. Open, redistributable substitutes
# --------------------------------------------------------------------------------------

if [ "$WANT_OPEN" = "1" ]; then
  say "Open metric-compatible substitutes (apt)"
  # fonts-croscore is Arimo/Tinos/Cousine (Arial/Times/Courier); the two crosextra
  # packages are Carlito (Calibri) and Caladea (Cambria); fonts-liberation2 is the same
  # designs again under different names and is included because plenty of decks and
  # themes name "Liberation Sans" directly.
  $SUDO apt-get update -qq
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
# 2. Aptos, from Microsoft's own download
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
  need() { command -v "$1" >/dev/null || { echo "missing: $1 (apt install $2)" >&2; exit 1; }; }
  need curl curl
  need unzip unzip

  tmp="$(mktemp -d)"
  trap 'rm -rf "$tmp"' EXIT
  curl -fsSL "$APTOS_URL" -o "$tmp/aptos.zip"
  if [ -n "$APTOS_SHA256" ]; then
    echo "$APTOS_SHA256  $tmp/aptos.zip" | sha256sum -c - || {
      echo "Aptos download does not match the pinned hash; refusing to install." >&2
      exit 1
    }
  else
    got="$(wc -c < "$tmp/aptos.zip" | tr -d ' ')"
    if [ -n "$APTOS_EXPECTED_BYTES" ] && [ "$got" != "$APTOS_EXPECTED_BYTES" ]; then
      echo "Aptos download is $got bytes, expected $APTOS_EXPECTED_BYTES." >&2
      echo "Microsoft has changed what this URL serves.  Re-verify before installing:" >&2
      echo "  sha256 $(sha256sum "$tmp/aptos.zip" | cut -d' ' -f1)" >&2
      echo "Then set APTOS_SHA256 (and APTOS_EXPECTED_BYTES) and run again." >&2
      exit 1
    fi
    echo "note: APTOS_SHA256 is unset, so only the size was checked ($got bytes)."
    echo "      Its sha256 is: $(sha256sum "$tmp/aptos.zip" | cut -d' ' -f1)"
    echo "      Set APTOS_SHA256 to that value to pin it for your fleet."
  fi
  $SUDO unzip -jo "$tmp/aptos.zip" '*.ttf' -d "$FONTDIR"
fi

# --------------------------------------------------------------------------------------
# 3. Calibri and Cambria, via the PowerPoint Viewer installer
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
  command -v cabextract >/dev/null || {
    echo "missing: cabextract (apt install cabextract)" >&2; exit 1; }
  command -v curl >/dev/null || { echo "missing: curl" >&2; exit 1; }

  tmp2="$(mktemp -d)"
  trap 'rm -rf "$tmp2"' EXIT
  curl -fsSL "$PPVIEWER_URL" -o "$tmp2/PowerPointViewer.exe"
  echo "$PPVIEWER_SHA256  $tmp2/PowerPointViewer.exe" | sha256sum -c - || {
    echo "PowerPointViewer.exe does not match the pinned hash; refusing to extract." >&2
    exit 1
  }

  # Two nested cabinets: the installer carries ppviewer.cab, which carries the fonts.
  # -L lowercases the names, which matters because the cabinet stores them shouting
  # (CALIBRI.TTF) and fontconfig does not care but humans reading ls -1 do.
  ( cd "$tmp2" && cabextract -q PowerPointViewer.exe -F ppviewer.cab )
  target="$FONTDIR/ppviewer"
  $SUDO mkdir -p "$target"
  ( cd "$tmp2" && $SUDO cabextract -q -L -F '*.TT[CF]' -d "$target" ppviewer.cab )
  echo "extracted:"
  ls -1 "$target"
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
    printf '  %-22s ok\n' "$1"
  else
    printf '  %-22s MISSING\n' "$1"
    status=1
  fi
}

if [ "$WANT_OPEN" = "1" ]; then
  # Raleway is deliberately absent: Debian does not package it.  See the apt call above.
  for family in Carlito Caladea Arimo Tinos Cousine "Noto Sans CJK JP" Lato; do
    check "$family"
  done
fi
[ "$WANT_APTOS" = "1" ] && check "Aptos"
# Calibri and Cambria are what pptx2svg maps; the other four come along for free and
# are worth naming so a missing one is visible rather than mysterious.
[ "$WANT_PPVIEWER" = "1" ] && for family in Calibri Cambria Candara Consolas Constantia Corbel Meiryo; do
  check "$family"
done

say "Done"
echo "Point pptx2svg at these with:   pptx2svg deck.pptx -f png --font-dir $FONTDIR --system-fonts"
echo "Or check what a deck needs with: pptx2svg fonts --check deck.pptx --system-fonts"
exit $status

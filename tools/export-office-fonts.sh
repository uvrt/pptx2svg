#!/usr/bin/env bash
#
# Copy the fonts a licensed Microsoft Office installation carries into one directory,
# ready to move to another machine and install with `tools/install-fonts-debian.sh
# --office-dir <that directory>`.
#
# WHY THIS EXISTS.  About 150 of the ~186 families PowerPoint ships have no open clone
# at all -- Gill Sans MT, Rockwell, Franklin Gothic, Tw Cen MT, Perpetua, Bookman Old
# Style, the Lucida family, and every Indic and CJK face Office bundles.  A deck naming
# one of them is laid out from *guessed* advance widths, which moves line breaks rather
# than merely changing glyph shapes.  See FONTS.md.  No bundle can fix that, because
# these files are not redistributable; the only lawful source is an Office installation
# somebody already licensed.
#
# WHAT THIS IS NOT.  It is not a download, and it never will be.  It copies from a
# licensed install that YOU point it at, onto a machine that is YOURS.  Microsoft's font
# licence governs where those files may go -- the Office EULA licenses the fonts for use
# with your licensed Office installation, and putting them on an unrelated machine is
# generally outside it.  Whether your situation is covered is between you and Microsoft;
# this script takes no view and grants you nothing.  It exists so that an operator who
# *is* covered does not have to hand-copy 280 files and guess at what they got.
#
# NOTHING HERE MAY ENTER THE REPOSITORY.  The default target is outside the checkout for
# that reason, and the script refuses to write inside it.  No Microsoft font is committed
# to this project, ever, in any form -- only measured advance widths, which are facts
# about a file rather than the file.
#
# Usage:
#   tools/export-office-fonts.sh                       # auto-discover, to ~/pptx2svg-office-fonts
#   tools/export-office-fonts.sh --out /mnt/usb/fonts  # somewhere else
#   tools/export-office-fonts.sh --source /path/DFonts # a specific install
#   tools/export-office-fonts.sh --tar                 # also write a .tar.gz for transfer
#
# Then, on the Debian host:
#   tools/install-fonts-debian.sh --office-dir /path/to/that/directory
#
set -euo pipefail

OUT="${HOME}/pptx2svg-office-fonts"
SOURCES=()
WANT_TAR=0

while [ $# -gt 0 ]; do
  case "$1" in
    --out)     shift; OUT="${1:?--out needs a path}" ;;
    --source)  shift; SOURCES+=("${1:?--source needs a path}") ;;
    --tar)     WANT_TAR=1 ;;
    -h|--help) sed -n '2,37p' "$0"; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
  shift
done

say() { printf '\n== %s\n' "$*"; }

# --------------------------------------------------------------------------------------
# Refuse to write inside the checkout
# --------------------------------------------------------------------------------------
#
# A .gitignore line is one edit away from not protecting you, and `git add -f` ignores it
# outright.  Licensed fonts belong outside the working tree, on the same footing as the
# oracle exports and the font profile.
REPO="$(cd "$(dirname "$0")/.." && pwd -P)"
mkdir -p "$OUT"
OUT="$(cd "$OUT" && pwd -P)"
case "$OUT/" in
  "$REPO"/*)
    echo "refusing to write licensed fonts inside the repository: $OUT" >&2
    echo "pick a path outside $REPO -- the default ~/pptx2svg-office-fonts is fine." >&2
    exit 1 ;;
esac

# --------------------------------------------------------------------------------------
# Where Office keeps them
# --------------------------------------------------------------------------------------
#
# Two places, and the second is the one everybody forgets.  DFonts holds what the
# installer laid down.  The cloud-font cache holds faces Office downloads on demand --
# Aptos Display among them, which is why a machine with Office can *use* Aptos Display
# while `fc-list` has never heard of it.
if [ "${#SOURCES[@]}" -eq 0 ]; then
  case "$(uname -s)" in
    Darwin)
      for app in PowerPoint Word Excel; do
        d="/Applications/Microsoft ${app}.app/Contents/Resources/DFonts"
        [ -d "$d" ] && SOURCES+=("$d")
      done
      for d in "${HOME}/Library/Group Containers/"*.Office/FontCache/*/CloudFonts; do
        [ -d "$d" ] && SOURCES+=("$d")
      done
      ;;
    *)
      for d in "/usr/share/fonts/truetype/msttcorefonts" "/c/Windows/Fonts" "/mnt/c/Windows/Fonts"; do
        [ -d "$d" ] && SOURCES+=("$d")
      done
      ;;
  esac
fi

if [ "${#SOURCES[@]}" -eq 0 ]; then
  echo "no Office font directory found; pass --source <path> explicitly." >&2
  echo "On macOS it is inside the application bundle:" >&2
  echo "  /Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts" >&2
  exit 1
fi

say "Sources"
for d in "${SOURCES[@]}"; do echo "  $d"; done

# --------------------------------------------------------------------------------------
# Copy
# --------------------------------------------------------------------------------------

say "Copying to $OUT"
copied=0
skipped=0
for dir in "${SOURCES[@]}"; do
  # -print0/read -d '' so a family with a space in its name survives, which on the cloud
  # cache is most of them.
  while IFS= read -r -d '' file; do
    base="$(basename "$file")"
    target="$OUT/$base"
    # Same name and same size means the same file; two Office apps ship overlapping
    # DFonts directories and copying 516 MB twice helps nobody.
    if [ -f "$target" ] && [ "$(wc -c < "$target")" = "$(wc -c < "$file")" ]; then
      skipped=$((skipped + 1))
      continue
    fi
    cp -p "$file" "$target"
    copied=$((copied + 1))
  done < <(find "$dir" -type f \( -iname '*.ttf' -o -iname '*.ttc' -o -iname '*.otf' -o -iname '*.otc' \) -print0)
done
echo "  $copied copied, $skipped already present"

# --------------------------------------------------------------------------------------
# Manifest
# --------------------------------------------------------------------------------------
#
# A directory of 280 opaque filenames is not an inventory.  The manifest records what
# family each file actually answers to -- read out of the file's own name table, because
# a family name is inside the file and not in its name -- so the operator on the far end
# can tell what arrived and diff it later.
say "Writing MANIFEST.tsv"
python3 - "$OUT" <<'PY'
import hashlib, pathlib, struct, sys

root = pathlib.Path(sys.argv[1])

def families(data: bytes) -> list[str]:
    """nameID 1 of every face in the file, TrueType collections included."""
    def table_dir(offset: int):
        count = struct.unpack_from(">H", data, offset + 4)[0]
        return {
            data[offset + 12 + 16 * i: offset + 16 + 16 * i].decode("latin1"):
            struct.unpack_from(">II", data, offset + 20 + 16 * i)
            for i in range(count)
        }

    heads = []
    if data[:4] == b"ttcf":
        n = struct.unpack_from(">I", data, 8)[0]
        heads = list(struct.unpack_from(f">{n}I", data, 12))
    else:
        heads = [0]

    out = []
    for head in heads:
        try:
            name_off, _ = table_dir(head)["name"]
            count, string_off = struct.unpack_from(">HH", data, name_off + 2)
            for i in range(count):
                rec = name_off + 6 + 12 * i
                pid, _, _, nid, length, off = struct.unpack_from(">HHHHHH", data, rec)
                if nid != 1:
                    continue
                raw = data[name_off + string_off + off:][:length]
                # Platform 0 (Unicode) and 3 (Windows) are both UTF-16BE; only
                # platform 1 (Macintosh) is a byte encoding.  Reading a Unicode record
                # as latin1 yields "T i m e s   N e w   R o m a n", which is how this
                # was caught.
                encoding = "mac-roman" if pid == 1 else "utf-16-be"
                value = raw.decode(encoding, "replace").strip()
                if value and value not in out:
                    out.append(value)
        except Exception:
            continue
    return out

rows = []
for path in sorted(root.iterdir()):
    if path.suffix.lower() not in {".ttf", ".ttc", ".otf", ".otc"}:
        continue
    data = path.read_bytes()
    rows.append((
        path.name,
        "; ".join(families(data)) or "(unreadable)",
        str(len(data)),
        hashlib.sha256(data).hexdigest(),
    ))

manifest = root / "MANIFEST.tsv"
with manifest.open("w", encoding="utf-8") as handle:
    handle.write("file\tfamilies\tbytes\tsha256\n")
    for row in rows:
        handle.write("\t".join(row) + "\n")

seen = sorted({f for _, fams, _, _ in rows for f in fams.split("; ") if f != "(unreadable)"})
print(f"  {len(rows)} files, {len(seen)} distinct families")
PY

cat > "$OUT/READ-ME-FIRST.txt" <<'NOTICE'
These are Microsoft-licensed font files, copied from an Office installation.

They are NOT redistributable.  They were placed here by
tools/export-office-fonts.sh from pptx2svg, for the operator's own use with
machines they control and licences they hold.  Nothing in pptx2svg grants any
right to these files, and none of them is committed to that project.

Do not put this directory in version control, in a container image, in a build
artifact, or on a machine you are not licensed for.

To install them on a Debian host:

    tools/install-fonts-debian.sh --office-dir /path/to/this/directory

MANIFEST.tsv lists every file with the families it answers to and its sha256.
NOTICE

say "Summary"
echo "  directory : $OUT"
echo "  size      : $(du -sh "$OUT" | cut -f1)"
echo "  manifest  : $OUT/MANIFEST.tsv"

if [ "$WANT_TAR" = "1" ]; then
  tarball="${OUT%/}.tar.gz"
  say "Writing $tarball"
  tar czf "$tarball" -C "$(dirname "$OUT")" "$(basename "$OUT")"
  echo "  $(du -sh "$tarball" | cut -f1)"
fi

say "Next"
echo "  On the Debian host:"
echo "    tools/install-fonts-debian.sh --office-dir /path/to/$(basename "$OUT")"
echo "  Then check a deck against it:"
echo "    pptx2svg fonts --check deck.pptx --font-dir ~/.local/share/fonts --system-fonts"

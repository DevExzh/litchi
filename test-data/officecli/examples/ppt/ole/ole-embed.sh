#!/bin/bash
# PowerPoint OLE embedded objects — embed .xlsx / .docx binary payloads in a deck.
#
# CLI twin of ole-embed.py (officecli Python SDK). Both produce an equivalent
# ole-embed.pptx. This one drives the officecli binary directly: one
# `officecli add ... --type ole --prop k=v` invocation per embedded object.
#
#   - build tiny source docs first (a spreadsheet + a Word memo) via officecli
#   - slide 1: embed the .xlsx (Excel.Sheet.12) and the .docx (Word.Document.12),
#     each with an explicit preview= thumbnail so they render visibly
#   - slide 2: WITH preview= vs WITHOUT — a deliberate teaching contrast
#   - each ole object is read back with `get` (contentType / fileSize / progId / relId)
#
# IMPORTANT: officecli does NOT auto-generate the Office live-preview image for
# an embedded OLE object. Without preview=, the object embeds and validates
# fine, but real PowerPoint renders it as a BLANK rectangle in static/print
# view — it only becomes visible after the user double-clicks to activate it.
# Supply preview= for any OLE object you want visible in a static view.
#
# The embedded source files (data.xlsx, data.docx) and the preview PNGs live in
# this example dir alongside the deck — they are the meaningful inputs the embed
# step consumes, kept in-dir.
#
# Requirements: Pillow (pip install Pillow) to synthesize the preview thumbnails.
# Usage: ./ole-embed.sh [officecli path]

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
DIR="$(dirname "$0")"
FILE="$DIR/ole-embed.pptx"

# Source payloads embedded into the deck (kept in-dir — they are the demo inputs).
XLSX="$DIR/data.xlsx"
DOCX="$DIR/data.docx"
THUMB_XLSX="$DIR/thumb-xlsx.png"
THUMB_DOCX="$DIR/thumb-docx.png"

# ── Build the .xlsx payload (a tiny spreadsheet) ──────────────────────────────
rm -f "$XLSX"
$CLI create "$XLSX"
$CLI open "$XLSX"
$CLI set "$XLSX" '/sheet[1]/A1' --prop value="Q1 Revenue"
$CLI set "$XLSX" '/sheet[1]/A2' --prop value="North"
$CLI set "$XLSX" '/sheet[1]/B2' --prop value="1200"
$CLI set "$XLSX" '/sheet[1]/A3' --prop value="South"
$CLI set "$XLSX" '/sheet[1]/B3' --prop value="980"
$CLI close "$XLSX"

# ── Build the .docx payload (a tiny Word memo) ────────────────────────────────
rm -f "$DOCX"
$CLI create "$DOCX"
$CLI open "$DOCX"
$CLI add "$DOCX" /body --type paragraph \
    --prop text="Quarterly Memo" --prop bold=true --prop size=16
$CLI add "$DOCX" /body --type paragraph \
    --prop text="Revenue is up 12% quarter over quarter."
$CLI close "$DOCX"

# ── Synthesize two distinct preview thumbnails (Excel-green, Word-blue) ───────
python3 - "$THUMB_XLSX" "$THUMB_DOCX" <<'PY'
import sys
from PIL import Image, ImageDraw

xlsx_path, docx_path = sys.argv[1], sys.argv[2]

# Excel-styled thumbnail (green) for the .xlsx embed
img = Image.new("RGB", (240, 160), (33, 115, 70))
d = ImageDraw.Draw(img)
d.rectangle((8, 8, 231, 151), outline=(255, 255, 255), width=4)
d.text((30, 40), "Q1 Revenue", fill=(255, 255, 255))
d.text((30, 90), "(embedded .xlsx)", fill=(200, 230, 210))
img.save(xlsx_path)

# Word-styled thumbnail (blue) for the .docx embed
img = Image.new("RGB", (240, 160), (43, 87, 154))
d = ImageDraw.Draw(img)
d.rectangle((8, 8, 231, 151), outline=(255, 255, 255), width=4)
d.text((30, 40), "Quarterly Memo", fill=(255, 255, 255))
d.text((30, 90), "(embedded .docx)", fill=(205, 218, 240))
img.save(docx_path)
PY

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ══════════════════════════════════════════════════════════════════════════════
# Slide 1: embed an .xlsx and a .docx, each with a preview= so they render
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" / --type slide
$CLI add "$FILE" "/slide[1]" --type textbox \
    --prop text="Embedded OLE objects — Excel + Word packages" \
    --prop size=24 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# Embed the spreadsheet. Double-clicking the object in PowerPoint opens the
# workbook in-place. progId=Excel.Sheet.12 marks it as a modern .xlsx package.
# preview= gives it a visible face in static view (officecli won't bake one).
# Features: --type ole, src (embedded payload), progId, preview, x/y/width/height
$CLI add "$FILE" "/slide[1]" --type ole \
    --prop src="$XLSX" \
    --prop progId=Excel.Sheet.12 \
    --prop preview="$THUMB_XLSX" \
    --prop x=1.5in --prop y=1.8in --prop width=3.5in --prop height=2.5in
$CLI add "$FILE" "/slide[1]" --type textbox \
    --prop text="src=data.xlsx  progId=Excel.Sheet.12  preview=thumb-xlsx.png" \
    --prop size=12 --prop italic=true \
    --prop x=1.5in --prop y=4.4in --prop width=5in --prop height=0.4in

# Embed the Word memo. progId=Word.Document.12 marks it as a modern .docx.
# preview= gives it a visible face in static view (officecli won't bake one).
# Features: --type ole, src (embedded payload), progId, preview, x/y/width/height
$CLI add "$FILE" "/slide[1]" --type ole \
    --prop src="$DOCX" \
    --prop progId=Word.Document.12 \
    --prop preview="$THUMB_DOCX" \
    --prop x=6.5in --prop y=1.8in --prop width=3.5in --prop height=2.5in
$CLI add "$FILE" "/slide[1]" --type textbox \
    --prop text="src=data.docx  progId=Word.Document.12  preview=thumb-docx.png" \
    --prop size=12 --prop italic=true \
    --prop x=6.5in --prop y=4.4in --prop width=5in --prop height=0.4in

# ══════════════════════════════════════════════════════════════════════════════
# Slide 2: WITH preview= vs WITHOUT — a deliberate teaching contrast
# ══════════════════════════════════════════════════════════════════════════════
$CLI add "$FILE" / --type slide
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="preview= is the reliable path to a visible OLE object" \
    --prop size=24 --prop bold=true \
    --prop x=0.5in --prop y=0.3in --prop width=12in --prop height=0.6in

# WITH preview= — the thumbnail is drawn in the object frame in static view.
# add-time only (Set ignores this key).
# Features: --type ole, src, progId, preview (thumbnail image), x/y/width/height
$CLI add "$FILE" "/slide[2]" --type ole \
    --prop src="$XLSX" \
    --prop progId=Excel.Sheet.12 \
    --prop preview="$THUMB_XLSX" \
    --prop x=1.5in --prop y=1.8in --prop width=3.5in --prop height=2.5in
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="src=data.xlsx  preview=thumb-xlsx.png  (renders visibly)" \
    --prop size=12 --prop italic=true \
    --prop x=1.5in --prop y=4.4in --prop width=4.5in --prop height=0.4in

# WITHOUT preview= — embeds and validates fine, but real PowerPoint renders a
# BLANK rectangle in static view; it only appears once double-clicked/activated.
# officecli does NOT auto-generate the Office live-preview image.
# Features: --type ole (no preview — blank until activated)
$CLI add "$FILE" "/slide[2]" --type ole \
    --prop src="$XLSX" \
    --prop progId=Excel.Sheet.12 \
    --prop x=6.5in --prop y=1.8in --prop width=3.5in --prop height=2.5in
$CLI add "$FILE" "/slide[2]" --type textbox \
    --prop text="src=data.xlsx  (no preview — blank until opened in PowerPoint)" \
    --prop size=12 --prop italic=true \
    --prop x=6.5in --prop y=4.4in --prop width=5in --prop height=0.4in

$CLI close "$FILE"

# ══════════════════════════════════════════════════════════════════════════════
# Inspect: Get surfaces read-only readbacks (src is NOT echoed — use relId).
# ══════════════════════════════════════════════════════════════════════════════
echo "── query all OLE objects ──"
$CLI query "$FILE" ole
echo "── get slide 1 / ole 1 (the .xlsx) ──"
$CLI get "$FILE" "/slide[1]/ole[1]"
echo "── get slide 1 / ole 2 (the .docx) ──"
$CLI get "$FILE" "/slide[1]/ole[2]"

$CLI validate "$FILE"
echo "Generated: $FILE"

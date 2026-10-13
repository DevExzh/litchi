#!/bin/bash
# Excel Shapes Gallery — generates shapes.xlsx exercising the xlsx `shape`
# property surface (officecli help xlsx shape).
#
# CLI twin of shapes.py (officecli Python SDK). Both produce an equivalent
# shapes.xlsx: a single "Gallery" sheet laid out as a shape gallery — each
# demo shape sits in a readable grid with a label cell above it.
#
# Shapes are anchored with cell-unit x/y/width/height (TwoCellAnchor
# column/row indices — NOT points/inches). Get reads position back as those
# same integer indices.
#
# Usage: ./shapes.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/shapes.xlsx"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# Title cell + column widths so labels are legible.
$CLI set "$FILE" /Sheet1/A1 --prop value="Excel Shapes Gallery" --prop font.bold=true --prop font.size=16 --prop fill=1F4E79 --prop font.color=FFFFFF
$CLI set "$FILE" /Sheet1/A2 --prop value="Each shape sits below its label. Shapes are cell-anchored (x/y/width/height in column/row units)." --prop font.italic=true --prop font.color=595959

# Label helper cells sit on a header row above each shape band.
# Band 1 labels (row 4), shapes span rows 5-8.
$CLI set "$FILE" /Sheet1/A4 --prop value="geometry=rect" --prop font.bold=true
$CLI set "$FILE" /Sheet1/D4 --prop value="geometry=roundRect" --prop font.bold=true
$CLI set "$FILE" /Sheet1/G4 --prop value="geometry=ellipse" --prop font.bold=true
$CLI set "$FILE" /Sheet1/J4 --prop value="geometry=triangle" --prop font.bold=true

# ─────────────────────────────────────────────────────────────────────────────
# Band 1 — Geometry preset gallery (solid theme-accent fill, white text)
# ─────────────────────────────────────────────────────────────────────────────
# Features: geometry=rect, solid fill, text, color, bold, valign
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=rect \
    --prop x=0 --prop y=4 --prop width=2 --prop height=3 \
    --prop fill=4472C4 --prop text="rect" --prop color=FFFFFF --prop bold=true

# Features: geometry=roundRect
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=roundRect \
    --prop x=3 --prop y=4 --prop width=2 --prop height=3 \
    --prop fill=2A9D8F --prop text="roundRect" --prop color=FFFFFF --prop bold=true

# Features: geometry=ellipse
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=ellipse \
    --prop x=6 --prop y=4 --prop width=2 --prop height=3 \
    --prop fill=E76F51 --prop text="ellipse" --prop color=FFFFFF --prop bold=true

# Features: geometry=triangle
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=triangle \
    --prop x=9 --prop y=4 --prop width=2 --prop height=3 \
    --prop fill=E9C46A --prop text="triangle" --prop color=264653 --prop bold=true

# Band 2 labels (row 9), shapes span rows 10-13.
$CLI set "$FILE" /Sheet1/A9 --prop value="geometry=diamond" --prop font.bold=true
$CLI set "$FILE" /Sheet1/D9 --prop value="geometry=parallelogram" --prop font.bold=true
$CLI set "$FILE" /Sheet1/G9 --prop value="geometry=rightArrow" --prop font.bold=true
$CLI set "$FILE" /Sheet1/J9 --prop value="geometry=star5 (name=MyStar)" --prop font.bold=true

# Features: geometry=diamond
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=diamond \
    --prop x=0 --prop y=9 --prop width=2 --prop height=3 \
    --prop fill=8E44AD --prop text="diamond" --prop color=FFFFFF --prop bold=true

# Features: geometry=parallelogram
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=parallelogram \
    --prop x=3 --prop y=9 --prop width=2 --prop height=3 \
    --prop fill=457B9D --prop text="parallelogram" --prop color=FFFFFF --prop bold=true

# Features: geometry=rightArrow
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=rightArrow \
    --prop x=6 --prop y=9 --prop width=2 --prop height=3 \
    --prop fill=F4A261 --prop text="rightArrow" --prop color=264653 --prop bold=true

# Features: geometry=star5, name (override cNvPr @name)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=star5 \
    --prop x=9 --prop y=9 --prop width=2 --prop height=3 \
    --prop fill=E63946 --prop name=MyStar

# Band 3 labels (row 14), shapes span rows 15-18.
$CLI set "$FILE" /Sheet1/A14 --prop value="flipH=true (mirrored arrow)" --prop font.bold=true
$CLI set "$FILE" /Sheet1/D14 --prop value="flipV=true" --prop font.bold=true
$CLI set "$FILE" /Sheet1/G14 --prop value="flipBoth=true" --prop font.bold=true
$CLI set "$FILE" /Sheet1/J14 --prop value="rotation=30" --prop font.bold=true

# ─────────────────────────────────────────────────────────────────────────────
# Band 3 — Flips & rotation
# ─────────────────────────────────────────────────────────────────────────────
# Features: flipH (horizontal mirror; Office-API alias of flip=h)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=rightArrow \
    --prop x=0 --prop y=14 --prop width=2 --prop height=3 \
    --prop fill=4472C4 --prop flipH=true

# Features: flipV (vertical mirror)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=triangle \
    --prop x=3 --prop y=14 --prop width=2 --prop height=3 \
    --prop fill=2A9D8F --prop flipV=true

# Features: flipBoth (both axes)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=parallelogram \
    --prop x=6 --prop y=14 --prop width=2 --prop height=3 \
    --prop fill=E76F51 --prop flipBoth=true

# Features: rotation (degrees clockwise; stored as 60000ths on @rot)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=rightArrow \
    --prop x=9 --prop y=14 --prop width=2 --prop height=3 \
    --prop fill=E9C46A --prop rotation=30 --prop text="30°" --prop color=264653 --prop bold=true

# Band 4 labels (row 19), shapes span rows 20-23.
$CLI set "$FILE" /Sheet1/A19 --prop value="glow=FFD700" --prop font.bold=true
$CLI set "$FILE" /Sheet1/D19 --prop value="gradientFill (2-stop)" --prop font.bold=true
$CLI set "$FILE" /Sheet1/G19 --prop value="reflection=true" --prop font.bold=true
$CLI set "$FILE" /Sheet1/J19 --prop value="shadow=000000 + softEdge=4" --prop font.bold=true

# ─────────────────────────────────────────────────────────────────────────────
# Band 4 — Effects: glow, gradient fill, reflection, shadow, soft edge
# ─────────────────────────────────────────────────────────────────────────────
# Features: glow (color halo; pass a color or 'true' for accent blue)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=ellipse \
    --prop x=0 --prop y=19 --prop width=2 --prop height=3 \
    --prop fill=2A9D8F --prop glow=FFD700

# Features: gradientFill (two-stop linear, C1-C2:angle; mutually exclusive with fill)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=roundRect \
    --prop x=3 --prop y=19 --prop width=2 --prop height=3 \
    --prop gradientFill=FF6B6B-4ECDC4:45 --prop text="gradient" --prop color=FFFFFF --prop bold=true

# Features: reflection (mirror-below effect; 'true' enables a default reflection)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=roundRect \
    --prop x=6 --prop y=19 --prop width=2 --prop height=3 \
    --prop fill=457B9D --prop reflection=true --prop text="reflection" --prop color=FFFFFF --prop bold=true

# Features: shadow (outer shadow) + softEdge (feathered border radius)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=rect \
    --prop x=9 --prop y=19 --prop width=2 --prop height=3 \
    --prop fill=8E44AD --prop shadow=000000 --prop softEdge=4 --prop text="shadow" --prop color=FFFFFF

# Band 5 labels (row 24), shapes span rows 25-28.
$CLI set "$FILE" /Sheet1/A24 --prop value="line=C00000:2:dash (color:width:style)" --prop font.bold=true
$CLI set "$FILE" /Sheet1/D24 --prop value="styled text (font/size/align)" --prop font.bold=true
$CLI set "$FILE" /Sheet1/G24 --prop value="fill=accent1 (theme color)" --prop font.bold=true
$CLI set "$FILE" /Sheet1/J24 --prop value="anchor=K25:L28 (cell range)" --prop font.bold=true

# ─────────────────────────────────────────────────────────────────────────────
# Band 5 — Outline, text styling, theme fill, cell-range anchor
# ─────────────────────────────────────────────────────────────────────────────
# Features: line compound form 'color[:width[:style]]' (width in pt; style is a
#           prstDash token: solid/dash/dot/dashdot/longdash) — same grammar as
#           pptx shape line. fill=none (no fill → outline-only shape).
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=rect \
    --prop x=0 --prop y=24 --prop width=2 --prop height=3 \
    --prop fill=none --prop line=C00000:2:dash --prop text="outline" --prop color=C00000

# Features: font, size, italic, align (paragraph), valign, margin (text inset)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=roundRect \
    --prop x=3 --prop y=24 --prop width=2 --prop height=3 \
    --prop fill=E9C46A --prop text="Georgia italic center-top, inset 6pt" \
    --prop font=Georgia --prop size=12 --prop italic=true \
    --prop align=center --prop valign=top --prop margin=6 --prop color=264653

# Features: fill with a theme (scheme) color name — follows the workbook theme
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=ellipse \
    --prop x=6 --prop y=24 --prop width=2 --prop height=3 \
    --prop fill=accent1 --prop text="accent1" --prop color=FFFFFF --prop bold=true

# Features: anchor (cell-range form B2:F7; readback normalizes to x/y/width/height)
$CLI add "$FILE" /Sheet1 --type shape --prop geometry=roundRect \
    --prop anchor=K25:L28 \
    --prop fill=1D3557 --prop text="anchored K25:L28" --prop color=FFFFFF --prop bold=true

# Widen columns so the label cells are readable.
for c in 1 2 3 4 5 6 7 8 9 10 11 12; do
    $CLI set "$FILE" "/Sheet1/col[$c]" --prop width=13
done

$CLI close "$FILE"

# ─────────────────────────────────────────────────────────────────────────────
# Round-trip readback (fresh, from disk) — confirm props survive the write.
# ─────────────────────────────────────────────────────────────────────────────
$CLI query "$FILE" shape
$CLI get "$FILE" '/Sheet1/shape[8]'  --json   # star5 + name=MyStar
$CLI get "$FILE" '/Sheet1/shape[9]'  --json   # flipH
$CLI get "$FILE" '/Sheet1/shape[12]' --json   # rotation=30
$CLI get "$FILE" '/Sheet1/shape[13]' --json   # glow
$CLI get "$FILE" '/Sheet1/shape[20]' --json   # anchor → x/y/width/height

$CLI validate "$FILE"
echo "Generated: $FILE"

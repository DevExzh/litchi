#!/bin/bash
# Cell Formatting Showcase — generates cell-formatting.xlsx exercising the full
# xlsx `cell` property surface (schemas/help/xlsx/cell.json).
#
# CLI twin of cell-formatting.py (officecli Python SDK). Both produce an
# equivalent cell-formatting.xlsx.
#
# 6 sheets, one property group each:
#   Fonts    — font.name/size/bold/italic/color, underline, strike, super/subscript
#   Fills    — fill (hex/named/rgb), alignment.* + textRotation/indent/shrinkToFit
#   Borders  — border shorthand, border.all, per-side styles, border.color, diagonals
#   Numbers  — numberformat codes (thousands, %, currency, date, scientific, accounting)
#   Data     — value/type, formula, link + tooltip, locked, merge, arrayformula
#   RichText — runs (multi-format text within one cell; add-time only)
#
# `set` auto-creates the target cell, so no explicit `add` is needed per cell.
#
# Usage: ./cell-formatting.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/cell-formatting.xlsx"

rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==========================================================================
# Sheet1: Fonts — font.* family + underline/strike
# ==========================================================================
$CLI set "$FILE" /Sheet1/A1 --prop value="Cell font properties" --prop font.bold=true --prop font.size=14 --prop fill=1F4E79 --prop font.color=FFFFFF
$CLI set "$FILE" /Sheet1/A2 --prop value="Property" --prop font.bold=true --prop fill=D9E1F2
$CLI set "$FILE" /Sheet1/B2 --prop value="Rendered sample" --prop font.bold=true --prop fill=D9E1F2

# (A = label, B = rendered sample)
$CLI set "$FILE" /Sheet1/A3  --prop value="font.name=Georgia"
$CLI set "$FILE" /Sheet1/B3  --prop value="Georgia serif" --prop font.name=Georgia
$CLI set "$FILE" /Sheet1/A4  --prop value="font.size=18"
$CLI set "$FILE" /Sheet1/B4  --prop value="18pt text" --prop font.size=18
$CLI set "$FILE" /Sheet1/A5  --prop value="font.bold=true"
$CLI set "$FILE" /Sheet1/B5  --prop value="Bold text" --prop font.bold=true
$CLI set "$FILE" /Sheet1/A6  --prop value="font.italic=true"
$CLI set "$FILE" /Sheet1/B6  --prop value="Italic text" --prop font.italic=true
$CLI set "$FILE" /Sheet1/A7  --prop value="font.color=C00000"
$CLI set "$FILE" /Sheet1/B7  --prop value="Red text" --prop font.color=C00000
$CLI set "$FILE" /Sheet1/A8  --prop value="underline=single"
$CLI set "$FILE" /Sheet1/B8  --prop value="Underlined" --prop underline=single
$CLI set "$FILE" /Sheet1/A9  --prop value="underline=double"
$CLI set "$FILE" /Sheet1/B9  --prop value="Double underline" --prop underline=double
$CLI set "$FILE" /Sheet1/A10 --prop value="strike=true"
$CLI set "$FILE" /Sheet1/B10 --prop value="Struck out" --prop strike=true
$CLI set "$FILE" /Sheet1/A11 --prop value="superscript=true"
$CLI set "$FILE" /Sheet1/B11 --prop value="Superscript cell" --prop superscript=true
$CLI set "$FILE" /Sheet1/A12 --prop value="subscript=true"
$CLI set "$FILE" /Sheet1/B12 --prop value="Subscript cell" --prop subscript=true
$CLI set "$FILE" /Sheet1/A13 --prop value="combined"
$CLI set "$FILE" /Sheet1/B13 --prop value="Bold + italic + blue + 14pt" --prop font.bold=true --prop font.italic=true --prop font.color=2E75B6 --prop font.size=14

$CLI set "$FILE" "/Sheet1/col[1]" --prop width=22
$CLI set "$FILE" "/Sheet1/col[2]" --prop width=32

# ==========================================================================
# Sheet2: Fills & alignment
# ==========================================================================
$CLI add "$FILE" / --type sheet --prop name=Fills
$CLI set "$FILE" /Fills/A1 --prop value="Fills & alignment" --prop font.bold=true --prop font.size=14 --prop fill=548235 --prop font.color=FFFFFF

$CLI set "$FILE" /Fills/A2 --prop value="fill=E63946 (hex)" --prop fill=E63946 --prop font.color=FFFFFF
$CLI set "$FILE" /Fills/A3 --prop value="fill=gold (named)" --prop fill=gold
$CLI set "$FILE" /Fills/A4 --prop value="fill=rgb(46,157,182)" --prop fill="rgb(46,157,182)" --prop font.color=FFFFFF

$CLI set "$FILE" /Fills/A6 --prop value="left"   --prop fill=F2F2F2 --prop alignment.horizontal=left
$CLI set "$FILE" /Fills/A7 --prop value="center" --prop fill=F2F2F2 --prop alignment.horizontal=center
$CLI set "$FILE" /Fills/A8 --prop value="right"  --prop fill=F2F2F2 --prop alignment.horizontal=right
$CLI set "$FILE" /Fills/C6 --prop value="top"    --prop fill=FCE4D6 --prop alignment.vertical=top
$CLI set "$FILE" "/Fills/row[6]" --prop height=34
$CLI set "$FILE" /Fills/C7 --prop value="middle" --prop fill=FCE4D6 --prop alignment.vertical=center
$CLI set "$FILE" "/Fills/row[7]" --prop height=34
$CLI set "$FILE" /Fills/C8 --prop value="bottom" --prop fill=FCE4D6 --prop alignment.vertical=bottom
$CLI set "$FILE" "/Fills/row[8]" --prop height=34

$CLI set "$FILE" /Fills/A10 --prop value="This is a long sentence that wraps inside one cell via alignment.wrapText." --prop fill=E2EFDA --prop alignment.wrapText=true
$CLI set "$FILE" /Fills/A12 --prop value="RTL reading order" --prop fill=DDEBF7 --prop alignment.readingOrder=rtl

# textRotation / indent / shrinkToFit — set directly on alignment (canonical keys).
$CLI set "$FILE" /Fills/A14 --prop value="rotated 45deg" --prop fill=FFF2CC --prop alignment.textRotation=45
$CLI set "$FILE" "/Fills/row[14]" --prop height=40
$CLI set "$FILE" /Fills/A16 --prop value="indented 3" --prop fill=F2F2F2 --prop alignment.indent=3
$CLI set "$FILE" /Fills/A18 --prop value="ThisLongLabelShrinksToFit" --prop fill=E2EFDA --prop alignment.shrinkToFit=true

$CLI set "$FILE" "/Fills/col[1]" --prop width=30
$CLI set "$FILE" "/Fills/col[3]" --prop width=14

# ==========================================================================
# Sheet3: Borders
# ==========================================================================
$CLI add "$FILE" / --type sheet --prop name=Borders
$CLI set "$FILE" /Borders/A1 --prop value="Border styles" --prop font.bold=true --prop font.size=14 --prop fill=7030A0 --prop font.color=FFFFFF

$CLI set "$FILE" /Borders/B3  --prop value="border=thin (all)" --prop border=thin
$CLI set "$FILE" /Borders/B5  --prop value="border.all=medium" --prop border.all=medium
$CLI set "$FILE" /Borders/B7  --prop value="border + color" --prop border=thick --prop border.color=C00000
$CLI set "$FILE" /Borders/B9  --prop value="double bottom" --prop border.bottom=double
$CLI set "$FILE" /Borders/B11 --prop value="dashed box" --prop border.top=dashed --prop border.bottom=dashed --prop border.left=dashed --prop border.right=dashed
$CLI set "$FILE" /Borders/B13 --prop value="mixed sides" --prop border.left=thick --prop border.top=thin --prop border.right=medium --prop border.bottom=double
# Diagonal borders — direction via diagonalUp/Down, color requires a diagonal line.
$CLI set "$FILE" /Borders/B15 --prop value="diagonal up" --prop border.diagonal=thin --prop border.diagonalUp=true
$CLI set "$FILE" /Borders/B17 --prop value="diagonal down + color" --prop border.diagonal=medium --prop border.diagonalDown=true --prop border.diagonal.color=C00000

$CLI set "$FILE" "/Borders/col[1]" --prop width=18
$CLI set "$FILE" "/Borders/col[2]" --prop width=24

# ==========================================================================
# Sheet4: Number formats
# ==========================================================================
$CLI add "$FILE" / --type sheet --prop name=Numbers
$CLI set "$FILE" /Numbers/A1 --prop value="numberformat codes" --prop font.bold=true --prop font.size=14 --prop fill=C55A11 --prop font.color=FFFFFF
$CLI set "$FILE" /Numbers/A2 --prop value="Format code" --prop font.bold=true --prop fill=FCE4D6
$CLI set "$FILE" /Numbers/B2 --prop value="Result" --prop font.bold=true --prop fill=FCE4D6

# label cell: show the (short) code as literal text — type=string keeps codes
# like "0.00E+00" from being parsed as a scientific-notation number.
$CLI set "$FILE" /Numbers/A3 --prop value="#,##0" --prop type=string
$CLI set "$FILE" /Numbers/B3 --prop value=1234567 --prop numberformat="#,##0"
$CLI set "$FILE" /Numbers/A4 --prop value="#,##0.00" --prop type=string
$CLI set "$FILE" /Numbers/B4 --prop value=1234.5 --prop numberformat="#,##0.00"
$CLI set "$FILE" /Numbers/A5 --prop value="0.00%" --prop type=string
$CLI set "$FILE" /Numbers/B5 --prop value=0.1834 --prop numberformat="0.00%"
$CLI set "$FILE" /Numbers/A6 --prop value='$#,##0.00' --prop type=string
$CLI set "$FILE" /Numbers/B6 --prop value=29999.9 --prop numberformat='$#,##0.00'
$CLI set "$FILE" /Numbers/A7 --prop value="yyyy-mm-dd" --prop type=string
$CLI set "$FILE" /Numbers/B7 --prop value=45413 --prop numberformat="yyyy-mm-dd"
$CLI set "$FILE" /Numbers/A8 --prop value="0.00E+00" --prop type=string
$CLI set "$FILE" /Numbers/B8 --prop value=602214 --prop numberformat="0.00E+00"
$CLI set "$FILE" /Numbers/A9 --prop value='_(* #,##0.00_)' --prop type=string
$CLI set "$FILE" /Numbers/B9 --prop value=-4250 --prop numberformat='_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_)'

$CLI set "$FILE" "/Numbers/col[1]" --prop width=28
$CLI set "$FILE" "/Numbers/col[2]" --prop width=18

# ==========================================================================
# Sheet5: Data — value/type, formula, link, locked, merge
# ==========================================================================
$CLI add "$FILE" / --type sheet --prop name=Data
$CLI set "$FILE" /Data/A1 --prop value="Values, formulas, links" --prop font.bold=true --prop font.size=14 --prop fill=2E75B6 --prop font.color=FFFFFF

$CLI set "$FILE" /Data/A3 --prop value="Qty"
$CLI set "$FILE" /Data/B3 --prop value=12
$CLI set "$FILE" /Data/A4 --prop value="Price"
$CLI set "$FILE" /Data/B4 --prop value=4.5 --prop numberformat='$#,##0.00'
$CLI set "$FILE" /Data/A5 --prop value="Total" --prop font.bold=true
$CLI set "$FILE" /Data/B5 --prop formula="B3*B4" --prop numberformat='$#,##0.00' --prop font.bold=true

$CLI set "$FILE" /Data/A7 --prop value="type=string on a numeric value" --prop type=string
$CLI set "$FILE" /Data/B7 --prop value=007 --prop type=string

$CLI set "$FILE" /Data/A9 --prop value="OfficeCLI on GitHub" --prop link="https://github.com/iOfficeAI/OfficeCLI" --prop tooltip="Open the repo" --prop underline=single --prop font.color=0563C1

$CLI set "$FILE" /Data/A11 --prop value="locked cell (effective when sheet is protected)" --prop locked=true

$CLI set "$FILE" /Data/A13 --prop value="Merged title across A13:C13" --prop merge=A13:C13 --prop fill=DDEBF7 --prop alignment.horizontal=center --prop font.bold=true

# Dynamic-array formula — spills the result across the ref range.
$CLI set "$FILE" /Data/A15 --prop value="arrayformula = B3*2" --prop font.italic=true
$CLI set "$FILE" /Data/B15 --prop arrayformula="B3*2"

$CLI set "$FILE" "/Data/col[1]" --prop width=40
$CLI set "$FILE" "/Data/col[2]" --prop width=16

# ==========================================================================
# Sheet6: Rich-text — runs (multi-format text within one cell)
# ==========================================================================
# `runs` is an add-time property (requires --type cell + type=richtext). Each
# run is a JSON object with "text" plus any font props (bold, italic, color,
# size, underline). `set` does not support rich-text; use `add`.
$CLI add "$FILE" / --type sheet --prop name=RichText
$CLI set "$FILE" /RichText/A1 --prop value="runs — rich-text within one cell" --prop font.bold=true --prop font.size=14 --prop fill=5B2C8B --prop font.color=FFFFFF

# Each add creates the cell with multi-format text in a single SST entry.
$CLI add "$FILE" /RichText --type cell --prop ref=A3 --prop type=richtext \
    --prop 'runs=[{"text":"Bold + Red  ","bold":true,"color":"C00000"},{"text":"Italic + Blue","italic":true,"color":"2E75B6"},{"text":"  Normal"}]'
$CLI add "$FILE" /RichText --type cell --prop ref=A5 --prop type=richtext \
    --prop 'runs=[{"text":"H","bold":true,"color":"1F4E79","size":18},{"text":"2","superscript":true,"size":10},{"text":"O water formula","color":"1F4E79"}]'
$CLI add "$FILE" /RichText --type cell --prop ref=A7 --prop type=richtext \
    --prop 'runs=[{"text":"Strike","strike":true},{"text":" | "},{"text":"underline","underline":"single"},{"text":" | "},{"text":"size 14pt","size":14}]'

$CLI set "$FILE" "/RichText/col[1]" --prop width=50

$CLI close "$FILE"

# ==========================================================================
# Set -> Get round-trip: confirm canonical keys read back (fresh, from disk)
# ==========================================================================
$CLI get "$FILE" /Sheet1/B11 --json
$CLI get "$FILE" /Numbers/B6 --json
$CLI get "$FILE" /Borders/B9 --json

$CLI validate "$FILE"
echo "Generated: $FILE"

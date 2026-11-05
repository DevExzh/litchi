#!/bin/bash
# sparklines.sh — exercise the full xlsx `sparkline` element (in-cell mini charts)
# (schemas/help/xlsx/sparkline.json) using the officecli CLI.
#
# One dashboard sheet: a label column + 12 months of trend data per row, with a
# sparkline placed in the cell adjacent to each data row so the result reads like
# a real KPI dashboard. Demonstrates all three sparkline kinds (line / column /
# winLoss) plus every point-highlight, marker, colour and line-weight prop the
# element declares.
#
# CLI twin of sparklines.py (officecli SDK); both produce an equivalent
# sparklines.xlsx.
#
# Each sparkline is one `add` against the sheet: type= picks the kind,
# dataRange= is the source values, location= is the target cell. The group is
# stored under the x14 extension list and renders a tiny inline chart.
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/sparklines.xlsx"
echo "Building $FILE ..."
rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"

# helper: write a data row — label in A, 12 monthly values across B..M
row() {  # row <r> <label> <v1> ... <v12>
  local r="$1" label="$2"; shift 2
  officecli set "$FILE" "/Sheet1/A$r" --prop value="$label" --prop font.bold=true
  local cols=(B C D E F G H I J K L M) i=0
  for v in "$@"; do officecli set "$FILE" "/Sheet1/${cols[$i]}$r" --prop value="$v"; i=$((i+1)); done
}
sp() { officecli add "$FILE" /Sheet1 --type sparkline "${@:1}"; }   # sp --prop ...

# ===== Header row: label · Jan..Dec · Trend =====
officecli set "$FILE" /Sheet1/A1 --prop value="Region / Product" --prop font.bold=true --prop fill=1F4E79 --prop font.color=FFFFFF
mcols=(B C D E F G H I J K L M); mnames=(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec); i=0
for m in "${mnames[@]}"; do
  officecli set "$FILE" "/Sheet1/${mcols[$i]}1" --prop value="$m" --prop font.bold=true --prop fill=1F4E79 --prop font.color=FFFFFF
  i=$((i+1))
done
officecli set "$FILE" /Sheet1/N1 --prop value="Trend" --prop font.bold=true --prop fill=1F4E79 --prop font.color=FFFFFF

# ===== Data rows (rows 2..8) =====
row 2 "North"    45 52 48 61 58 67 72 69 74 81 78 90
row 3 "South"    88 84 79 72 68 61 55 49 44 40 38 35
row 4 "East"     30 55 20 70 35 82 40 90 25 60 45 100
row 5 "West"     12 15 14 18 22 25 24 28 30 33 31 40
row 6 "Central"  50 48 55 52 60 58 63 61 68 66 72 70
row 7 "Online"   -20 15 -35 40 -10 55 -50 30 -25 60 -15 80
row 8 "Kiosk"    5 -8 12 -3 20 -15 25 -6 30 -18 35 -10

# ===== Line sparklines (rows 2-3) =====
# Features: type=line, plain series colour, custom line weight
sp --prop type=line --prop dataRange=B2:M2 --prop location=N2 --prop color=#4472C4 --prop lineWeight=1.5
# Features: line + all point highlights + per-point marker colours + markers toggle
sp --prop type=line --prop dataRange=B3:M3 --prop location=N3 --prop color=#ED7D31 \
   --prop markers=true --prop highPoint=true --prop lowPoint=true --prop firstPoint=true --prop lastPoint=true \
   --prop highMarkerColor=#00B050 --prop lowMarkerColor=#FF0000 --prop firstMarkerColor=#7030A0 \
   --prop lastMarkerColor=#0070C0 --prop markersColor=#808080 --prop lineWeight=2.25

# ===== Column sparklines (rows 4-6) =====
# Features: type=column, high/low point highlight with marker colours
sp --prop type=column --prop dataRange=B4:M4 --prop location=N4 --prop color=#70AD47 \
   --prop highPoint=true --prop lowPoint=true --prop highMarkerColor=#00B050 --prop lowMarkerColor=#C00000
# Features: column, first/last point highlight
sp --prop type=column --prop dataRange=B5:M5 --prop location=N5 --prop color=#5B9BD5 \
   --prop firstPoint=true --prop lastPoint=true --prop firstMarkerColor=#264478 --prop lastMarkerColor=#0070C0
# Features: column, plain single-colour bars
sp --prop type=column --prop dataRange=B6:M6 --prop location=N6 --prop color=#A5A5A5

# ===== WinLoss sparklines (rows 7-8) =====
# Features: type=winLoss, negative points highlighted in their own colour
sp --prop type=winLoss --prop dataRange=B7:M7 --prop location=N7 --prop color=#4472C4 \
   --prop negative=true --prop negativeColor=#C00000
# Features: type=win-loss alias (maps to winLoss), high/low + negative
sp --prop type=win-loss --prop dataRange=B8:M8 --prop location=N8 --prop color=#7030A0 \
   --prop highPoint=true --prop lowPoint=true --prop negative=true --prop negativeColor=#FF0000

officecli close "$FILE"
officecli validate "$FILE"
echo "Created: $FILE"

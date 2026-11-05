#!/bin/bash
# slicers.sh — exercise the full xlsx `slicer` element (schemas/help/xlsx/slicer.json)
# using the officecli CLI.
#
# Slicers are the interactive button panels that filter a PivotTable. In OOXML a
# slicer is NOT free-standing: it is anchored to a *pivot cache field*, so it
# always binds to an existing PivotTable via `pivotTable=` + `field=`. This
# script therefore builds the prerequisites first (source data → PivotTable),
# then adds several slicers on different fields of that pivot.
#
# CLI twin of slicers.py (officecli Python SDK). Both produce an equivalent
# slicers.xlsx. See slicers.md for a per-slicer guide.
#
# Usage: ./slicers.sh [officecli path]
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
CLI="${1:-officecli}"
FILE="$(dirname "$0")/slicers.xlsx"

echo "Building $FILE ..."
rm -f "$FILE"
$CLI create "$FILE"
$CLI open "$FILE"

# ==========================================================================
# Source data — a realistic sales table: Region / Product / Quarter / Sales
# ==========================================================================
$CLI set "$FILE" /Sheet1/A1 --prop text=Region
$CLI set "$FILE" /Sheet1/B1 --prop text=Product
$CLI set "$FILE" /Sheet1/C1 --prop text=Quarter
$CLI set "$FILE" /Sheet1/D1 --prop text=Sales

# rows: Region Product Quarter Sales
write_row() {
  local r="$1"; shift
  local cols=(A B C D)
  local i=0
  for v in "$@"; do
    $CLI set "$FILE" "/Sheet1/${cols[$i]}${r}" --prop text="$v"
    i=$((i + 1))
  done
}

write_row 2  North Laptop Q1 12500
write_row 3  North Phone  Q2 8900
write_row 4  North Tablet Q3 6200
write_row 5  South Laptop Q1 22000
write_row 6  South Phone  Q2 18500
write_row 7  South Tablet Q4 7800
write_row 8  East  Laptop Q2 19500
write_row 9  East  Phone  Q3 13800
write_row 10 East  Tablet Q1 5400
write_row 11 West  Laptop Q4 25000
write_row 12 West  Phone  Q2 16800
write_row 13 West  Tablet Q3 8900

# ==========================================================================
# The slicer SOURCE — a PivotTable. Slicers anchor to this pivot's cache fields.
# ==========================================================================
# Features: source range, 1-level rows + column axis, single sum value field,
#   named so slicers can reference it by bare name (pivotTable=SalesPivot)
$CLI add "$FILE" / --type sheet --prop name=Dashboard
$CLI add "$FILE" /Dashboard --type pivottable \
  --prop source=Sheet1!A1:D13 \
  --prop rows=Region \
  --prop cols=Quarter \
  --prop values=Sales:sum \
  --prop layout=outline \
  --prop grandtotals=both \
  --prop name=SalesPivot \
  --prop style=PivotStyleMedium9

# ==========================================================================
# Slicers — each binds to SalesPivot via a different cache field.
# ==========================================================================

# Slicer 1: Region
# Features: pivotTable= (full path reference), field=Region, custom caption,
#   columnCount=2 (two-column button grid), rowHeight in EMU, explicit name
$CLI add "$FILE" /Dashboard --type slicer \
  --prop pivotTable=/Dashboard/pivottable[1] \
  --prop field=Region \
  --prop caption='Filter by Region' \
  --prop columnCount=2 \
  --prop rowHeight=250000 \
  --prop name=RegionSlicer

# Slicer 2: Product
# Features: pivotTable= by BARE NAME (resolves against the host sheet's pivots),
#   columnCount=3 (wide grid), caption defaulting shown on Slicer 3 instead
$CLI add "$FILE" /Dashboard --type slicer \
  --prop pivotTable=SalesPivot \
  --prop field=Product \
  --prop caption='Filter by Product' \
  --prop columnCount=3 \
  --prop name=ProductSlicer

# Slicer 3: Quarter
# Features: caption OMITTED — defaults to the field name ("Quarter"); rowHeight
#   OMITTED — defaults to 225425 EMU (~17.5pt). Minimal single-column slicer.
$CLI add "$FILE" /Dashboard --type slicer \
  --prop pivotTable=SalesPivot \
  --prop field=Quarter \
  --prop columnCount=1 \
  --prop name=QuarterSlicer

# ==========================================================================
# Modify an existing slicer with `set` (caption + columnCount are settable;
# `field` is add-time only and Set intentionally ignores it).
# ==========================================================================
# Features: set caption + set columnCount round-trip on an existing slicer
$CLI set "$FILE" '/Dashboard/slicer[1]' \
  --prop caption='Region' \
  --prop columnCount=1

$CLI close "$FILE"

$CLI validate "$FILE"
echo "Generated: $FILE"
echo "  Sheet1 (source data) + Dashboard (1 PivotTable + 3 slicers)"

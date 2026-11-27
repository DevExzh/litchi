#!/bin/bash
# sheet-settings.sh — exercise the xlsx `sheet` element's sheet-level property
# surface (schemas/help/xlsx/sheet.json) using the officecli CLI directly.
#
# These are the per-*worksheet* settings that live on <sheetView>, <pageSetup>,
# <headerFooter>, <sheetPr>, <sheetProtection> and the sheet's defined-names —
# distinct from the workbook-level settings in workbook-settings.{sh,py}. Four
# themed sheets: freeze panes, print setup, headers/footers, display+protection.
# CLI twin of sheet-settings.py (officecli SDK); both produce an equivalent
# sheet-settings.xlsx.
#
# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
FILE="$(dirname "$0")/sheet-settings.xlsx"
echo "Building $FILE ..."
rm -f "$FILE"
officecli create "$FILE"
officecli open "$FILE"          # resident mode: many sets in one process

# column letters, indexed positionally by the helpers below
COLS=(A B C D E F G H)
# helper: header row in row 1 (bold).   hdr <sheet> <c1> <c2> ...
hdr() {
  local sheet="$1"; shift
  local i=0
  for title in "$@"; do
    officecli set "$FILE" "/$sheet/${COLS[$i]}1" --prop value="$title" --prop font.bold=true
    i=$((i+1))
  done
}
# helper: one data row.   row <sheet> <rownum> <c1> <c2> ...
row() {
  local sheet="$1" r="$2"; shift 2
  local i=0
  for v in "$@"; do
    officecli set "$FILE" "/$sheet/${COLS[$i]}${r}" --prop value="$v"
    i=$((i+1))
  done
}

# =====================================================================
# Sheet 1 — Freeze Panes  (rename Sheet1)
# =====================================================================
officecli set "$FILE" /Sheet1 --prop name=1-Freeze-Panes
hdr 1-Freeze-Panes Date Region Product Units Revenue
row 1-Freeze-Panes 2 2026-01 North Widget 120 1140
row 1-Freeze-Panes 3 2026-01 South Gadget 95  1045
row 1-Freeze-Panes 4 2026-02 East  Widget 140 1225
row 1-Freeze-Panes 5 2026-02 West  Gizmo  88  968
row 1-Freeze-Panes 6 2026-03 North Gadget 132 1452

# Features: freeze panes — freeze the header row + the first (Date) column so
# they stay visible while scrolling. B2 = freeze row 1 AND column A.
officecli set "$FILE" /1-Freeze-Panes --prop freeze=B2

# =====================================================================
# Sheet 2 — Print Setup
# =====================================================================
officecli add "$FILE" / --type sheet --prop name=2-Print-Setup
hdr 2-Print-Setup Item Qty Unit Total
row 2-Print-Setup 2 Screws  500 0.02 10.00
row 2-Print-Setup 3 Bolts   300 0.05 15.00
row 2-Print-Setup 4 Washers 800 0.01 8.00
row 2-Print-Setup 5 Nuts    450 0.03 13.50
row 2-Print-Setup 6 Anchors 120 0.12 14.40

# Features: print setup — orientation, paper size (9=A4), fit-to-page, margins,
# a print area, and repeat-at-top print titles. These are print-only; verify via
# `get`, not visual render (the memo note applies).
officecli set "$FILE" /2-Print-Setup \
  --prop orientation=landscape \
  --prop paperSize=9 \
  --prop fitToPage=1x1 \
  --prop printArea=A1:D6 \
  --prop printTitleRows=1:1 \
  --prop printTitleCols=A:A \
  --prop margin.top=1.0in --prop margin.bottom=1.0in \
  --prop margin.left=0.5in --prop margin.right=0.5in \
  --prop margin.header=0.3in --prop margin.footer=0.3in

# =====================================================================
# Sheet 3 — Headers & Footers
# =====================================================================
officecli add "$FILE" / --type sheet --prop name=3-Headers-Footers
hdr 3-Headers-Footers Quarter Sales Target
row 3-Headers-Footers 2 Q1 45000 40000
row 3-Headers-Footers 3 Q2 52000 48000
row 3-Headers-Footers 4 Q3 61000 55000
row 3-Headers-Footers 5 Q4 58000 60000

# Features: header/footer — Excel format codes pass through verbatim.
#   &L left  &C center  &R right   &P page num  &N page count  &D date
officecli set "$FILE" /3-Headers-Footers \
  --prop header="&LQuarterly Report&C2026 Sales&R&D" \
  --prop footer="&LConfidential&CPage &P of &N&R&F"

# =====================================================================
# Sheet 4 — Display & Protection
# =====================================================================
officecli add "$FILE" / --type sheet --prop name=4-Display-Protection
hdr 4-Display-Protection Metric Value
row 4-Display-Protection 2 Users     1250
row 4-Display-Protection 3 Sessions  8400
row 4-Display-Protection 4 Bounce    32
row 4-Display-Protection 5 Retention 68

# Features: display — tab color, hide gridlines + row/col headings, zoom, an
# AutoFilter over the used range, and RTL layout.
officecli set "$FILE" /4-Display-Protection \
  --prop tabColor=C0392B \
  --prop gridlines=false \
  --prop headings=false \
  --prop zoom=125 \
  --prop autoFilter=A1:B5 \
  --prop direction=rtl

# Features: protection — enable sheet protection with a legacy password hash.
# (Do this last on this sheet; a protected sheet can't be sorted, so sorting
#  is demonstrated on its own sheet below.)
officecli set "$FILE" /4-Display-Protection \
  --prop protect=true \
  --prop password=secret123

# =====================================================================
# Sheet 5 — a hidden sheet + a sorted sheet (sort can't coexist w/ protect)
# =====================================================================
officecli add "$FILE" / --type sheet --prop name=5-Sorted --prop tabColor=27AE60
hdr 5-Sorted Name Score
row 5-Sorted 2 Carol 88
row 5-Sorted 3 Alice 95
row 5-Sorted 4 Bob   72
row 5-Sorted 5 Dave  60
# Features: sort — reorder rows by column B descending (highest score first).
officecli set "$FILE" /5-Sorted --prop sort="B desc"

# Features: hidden — a sheet hidden at creation time.
officecli add "$FILE" / --type sheet --prop name=6-Hidden --prop hidden=true
officecli set "$FILE" /6-Hidden/A1 --prop value="Hidden data sheet"

officecli close "$FILE"          # flush resident to disk
officecli validate "$FILE"
echo "Created: $FILE"

#!/bin/bash
# Generate complex table test documents (Word + Excel + PowerPoint)
# Includes merged cells, multi-level headers, formulas, charts, and other complex scenarios
# For testing officecli's table processing capabilities

# NOTE: intentionally NO `set -e`. Like the SDK twin's doc.batch, this script
# tolerates forward-compat 'UNSUPPORTED props' warnings (officecli exit 2) and
# keeps building so the full document is produced.
echo "Using CLI: officecli"

DIR="$(dirname "$0")"

###############################################################################
# 1. Word Complex Table Document
###############################################################################
DOCX="$DIR/tables.docx"
echo ""
echo "=========================================="
echo "Generating Word complex table document: $DOCX"
echo "=========================================="

rm -f "$DOCX"
officecli create "$DOCX"
officecli open "$DOCX"
officecli add "$DOCX" /body --type paragraph --prop text="Complex Table Examples" --prop style=Heading1 --prop align=center
officecli add "$DOCX" /body --type paragraph --prop text=""

# -- Table 1: Project Progress Tracker (vertical merge vmerge) --
echo "  -> Table 1: Project Progress Tracker"
officecli add "$DOCX" /body --type paragraph --prop text="1. Project Progress Tracker" --prop style=Heading2
officecli add "$DOCX" /body --type table --prop rows=7 --prop cols=6

# Header
officecli set "$DOCX" '/body/tbl[1]/tr[1]/tc[1]' --prop text="Project Name" --prop bold=true --prop shd=4472C4 --prop color=FFFFFF --prop valign=center
officecli set "$DOCX" '/body/tbl[1]/tr[1]/tc[2]' --prop text="Phase" --prop bold=true --prop shd=4472C4 --prop color=FFFFFF
officecli set "$DOCX" '/body/tbl[1]/tr[1]/tc[3]' --prop text="Owner" --prop bold=true --prop shd=4472C4 --prop color=FFFFFF
officecli set "$DOCX" '/body/tbl[1]/tr[1]/tc[4]' --prop text="Start Date" --prop bold=true --prop shd=4472C4 --prop color=FFFFFF
officecli set "$DOCX" '/body/tbl[1]/tr[1]/tc[5]' --prop text="End Date" --prop bold=true --prop shd=4472C4 --prop color=FFFFFF
officecli set "$DOCX" '/body/tbl[1]/tr[1]/tc[6]' --prop text="Progress" --prop bold=true --prop shd=4472C4 --prop color=FFFFFF

# Project A - Smart Office System (merge 3 rows)
officecli set "$DOCX" '/body/tbl[1]/tr[2]/tc[1]' --prop text="Smart Office System" --prop vmerge=restart --prop valign=center --prop shd=D9E2F3
officecli set "$DOCX" '/body/tbl[1]/tr[2]/tc[2]' --prop text="Requirements"
officecli set "$DOCX" '/body/tbl[1]/tr[2]/tc[3]' --prop text="John"
officecli set "$DOCX" '/body/tbl[1]/tr[2]/tc[4]' --prop text="2025-01-05"
officecli set "$DOCX" '/body/tbl[1]/tr[2]/tc[5]' --prop text="2025-02-15"
officecli set "$DOCX" '/body/tbl[1]/tr[2]/tc[6]' --prop text="100%" --prop color=00B050

officecli set "$DOCX" '/body/tbl[1]/tr[3]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=D9E2F3
officecli set "$DOCX" '/body/tbl[1]/tr[3]/tc[2]' --prop text="Development"
officecli set "$DOCX" '/body/tbl[1]/tr[3]/tc[3]' --prop text="Sarah"
officecli set "$DOCX" '/body/tbl[1]/tr[3]/tc[4]' --prop text="2025-02-16"
officecli set "$DOCX" '/body/tbl[1]/tr[3]/tc[5]' --prop text="2025-06-30"
officecli set "$DOCX" '/body/tbl[1]/tr[3]/tc[6]' --prop text="75%" --prop color=FFC000

officecli set "$DOCX" '/body/tbl[1]/tr[4]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=D9E2F3
officecli set "$DOCX" '/body/tbl[1]/tr[4]/tc[2]' --prop text="Testing"
officecli set "$DOCX" '/body/tbl[1]/tr[4]/tc[3]' --prop text="Mike"
officecli set "$DOCX" '/body/tbl[1]/tr[4]/tc[4]' --prop text="2025-07-01"
officecli set "$DOCX" '/body/tbl[1]/tr[4]/tc[5]' --prop text="2025-08-31"
officecli set "$DOCX" '/body/tbl[1]/tr[4]/tc[6]' --prop text="0%" --prop color=FF0000

# Project B - Data Platform Upgrade (merge 3 rows)
officecli set "$DOCX" '/body/tbl[1]/tr[5]/tc[1]' --prop text="Data Platform Upgrade" --prop vmerge=restart --prop valign=center --prop shd=E2EFDA
officecli set "$DOCX" '/body/tbl[1]/tr[5]/tc[2]' --prop text="Architecture"
officecli set "$DOCX" '/body/tbl[1]/tr[5]/tc[3]' --prop text="Emily"
officecli set "$DOCX" '/body/tbl[1]/tr[5]/tc[4]' --prop text="2025-03-01"
officecli set "$DOCX" '/body/tbl[1]/tr[5]/tc[5]' --prop text="2025-04-15"
officecli set "$DOCX" '/body/tbl[1]/tr[5]/tc[6]' --prop text="100%" --prop color=00B050

officecli set "$DOCX" '/body/tbl[1]/tr[6]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=E2EFDA
officecli set "$DOCX" '/body/tbl[1]/tr[6]/tc[2]' --prop text="Migration"
officecli set "$DOCX" '/body/tbl[1]/tr[6]/tc[3]' --prop text="David"
officecli set "$DOCX" '/body/tbl[1]/tr[6]/tc[4]' --prop text="2025-04-16"
officecli set "$DOCX" '/body/tbl[1]/tr[6]/tc[5]' --prop text="2025-07-31"
officecli set "$DOCX" '/body/tbl[1]/tr[6]/tc[6]' --prop text="40%" --prop color=FFC000

officecli set "$DOCX" '/body/tbl[1]/tr[7]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=E2EFDA
officecli set "$DOCX" '/body/tbl[1]/tr[7]/tc[2]' --prop text="Acceptance"
officecli set "$DOCX" '/body/tbl[1]/tr[7]/tc[3]' --prop text="Lisa"
officecli set "$DOCX" '/body/tbl[1]/tr[7]/tc[4]' --prop text="2025-08-01"
officecli set "$DOCX" '/body/tbl[1]/tr[7]/tc[5]' --prop text="2025-09-30"
officecli set "$DOCX" '/body/tbl[1]/tr[7]/tc[6]' --prop text="0%" --prop color=FF0000

# -- Table 2: Financial Statement (gridspan horizontal merge + vmerge vertical merge) --
echo "  -> Table 2: Financial Statement"
officecli add "$DOCX" /body --type paragraph --prop text=""
officecli add "$DOCX" /body --type paragraph --prop text="2. Financial Statement" --prop style=Heading2
officecli add "$DOCX" /body --type table --prop rows=8 --prop cols=5

# Header row 1 - gridspan=2 automatically removes merged tc
officecli set "$DOCX" '/body/tbl[2]/tr[1]/tc[1]' --prop text="Category" --prop bold=true --prop shd=2E75B6 --prop color=FFFFFF --prop vmerge=restart --prop valign=center
officecli set "$DOCX" '/body/tbl[2]/tr[1]/tc[2]' --prop text="Line Item" --prop bold=true --prop shd=2E75B6 --prop color=FFFFFF --prop vmerge=restart --prop valign=center
officecli set "$DOCX" '/body/tbl[2]/tr[1]/tc[3]' --prop text="Amount (10K USD)" --prop bold=true --prop shd=2E75B6 --prop color=FFFFFF --prop gridspan=2 --prop align=center
# gridspan=2 removed original tc[4], original tc[5] becomes tc[4]
officecli set "$DOCX" '/body/tbl[2]/tr[1]/tc[4]' --prop text="Notes" --prop bold=true --prop shd=2E75B6 --prop color=FFFFFF --prop vmerge=restart --prop valign=center

# Header row 2
officecli set "$DOCX" '/body/tbl[2]/tr[2]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=2E75B6
officecli set "$DOCX" '/body/tbl[2]/tr[2]/tc[2]' --prop text="" --prop vmerge=continue --prop shd=2E75B6
officecli set "$DOCX" '/body/tbl[2]/tr[2]/tc[3]' --prop text="Budget" --prop bold=true --prop shd=5B9BD5 --prop color=FFFFFF --prop align=center
officecli set "$DOCX" '/body/tbl[2]/tr[2]/tc[4]' --prop text="Actual" --prop bold=true --prop shd=5B9BD5 --prop color=FFFFFF --prop align=center
officecli set "$DOCX" '/body/tbl[2]/tr[2]/tc[5]' --prop text="" --prop vmerge=continue --prop shd=2E75B6

# Revenue (merge 3 rows)
officecli set "$DOCX" '/body/tbl[2]/tr[3]/tc[1]' --prop text="Revenue" --prop vmerge=restart --prop valign=center --prop shd=DEEAF6 --prop bold=true
officecli set "$DOCX" '/body/tbl[2]/tr[3]/tc[2]' --prop text="Product Sales"
officecli set "$DOCX" '/body/tbl[2]/tr[3]/tc[3]' --prop text="500.00" --prop align=right
officecli set "$DOCX" '/body/tbl[2]/tr[3]/tc[4]' --prop text="523.50" --prop align=right --prop color=00B050
officecli set "$DOCX" '/body/tbl[2]/tr[3]/tc[5]' --prop text="Exceeded"

officecli set "$DOCX" '/body/tbl[2]/tr[4]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=DEEAF6
officecli set "$DOCX" '/body/tbl[2]/tr[4]/tc[2]' --prop text="Consulting Services"
officecli set "$DOCX" '/body/tbl[2]/tr[4]/tc[3]' --prop text="200.00" --prop align=right
officecli set "$DOCX" '/body/tbl[2]/tr[4]/tc[4]' --prop text="185.30" --prop align=right --prop color=FF0000
officecli set "$DOCX" '/body/tbl[2]/tr[4]/tc[5]' --prop text="Below target"

officecli set "$DOCX" '/body/tbl[2]/tr[5]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=DEEAF6
officecli set "$DOCX" '/body/tbl[2]/tr[5]/tc[2]' --prop text="Tech Licensing"
officecli set "$DOCX" '/body/tbl[2]/tr[5]/tc[3]' --prop text="80.00" --prop align=right
officecli set "$DOCX" '/body/tbl[2]/tr[5]/tc[4]' --prop text="92.00" --prop align=right --prop color=00B050
officecli set "$DOCX" '/body/tbl[2]/tr[5]/tc[5]' --prop text="New partners"

# Expenses (merge 3 rows)
officecli set "$DOCX" '/body/tbl[2]/tr[6]/tc[1]' --prop text="Expenses" --prop vmerge=restart --prop valign=center --prop shd=FFF2CC --prop bold=true
officecli set "$DOCX" '/body/tbl[2]/tr[6]/tc[2]' --prop text="Labor Cost"
officecli set "$DOCX" '/body/tbl[2]/tr[6]/tc[3]' --prop text="320.00" --prop align=right
officecli set "$DOCX" '/body/tbl[2]/tr[6]/tc[4]' --prop text="335.00" --prop align=right --prop color=FF0000
officecli set "$DOCX" '/body/tbl[2]/tr[6]/tc[5]' --prop text="New hires"

officecli set "$DOCX" '/body/tbl[2]/tr[7]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=FFF2CC
officecli set "$DOCX" '/body/tbl[2]/tr[7]/tc[2]' --prop text="Operating Expenses"
officecli set "$DOCX" '/body/tbl[2]/tr[7]/tc[3]' --prop text="150.00" --prop align=right
officecli set "$DOCX" '/body/tbl[2]/tr[7]/tc[4]' --prop text="142.80" --prop align=right --prop color=00B050
officecli set "$DOCX" '/body/tbl[2]/tr[7]/tc[5]' --prop text="Cost savings"

officecli set "$DOCX" '/body/tbl[2]/tr[8]/tc[1]' --prop text="" --prop vmerge=continue --prop shd=FFF2CC
officecli set "$DOCX" '/body/tbl[2]/tr[8]/tc[2]' --prop text="R&D Investment"
officecli set "$DOCX" '/body/tbl[2]/tr[8]/tc[3]' --prop text="180.00" --prop align=right
officecli set "$DOCX" '/body/tbl[2]/tr[8]/tc[4]' --prop text="195.50" --prop align=right
officecli set "$DOCX" '/body/tbl[2]/tr[8]/tc[5]' --prop text="Strategic investment"

# -- Table 3: Skill Assessment Matrix (color heatmap) --
echo "  -> Table 3: Skill Assessment Matrix"
officecli add "$DOCX" /body --type paragraph --prop text=""
officecli add "$DOCX" /body --type paragraph --prop text="3. Skill Assessment Matrix" --prop style=Heading2
officecli add "$DOCX" /body --type table --prop rows=6 --prop cols=7

# Header
officecli set "$DOCX" '/body/tbl[3]/tr[1]/tc[1]' --prop text="Name/Skill" --prop bold=true --prop shd=002060 --prop color=FFFFFF --prop align=center
for col_data in "2:Python" "3:Java" "4:Frontend" "5:Database" "6:DevOps" "7:AI/ML"; do
    col="${col_data%%:*}"; name="${col_data#*:}"
    officecli set "$DOCX" "/body/tbl[3]/tr[1]/tc[$col]" --prop text="$name" --prop bold=true --prop shd=002060 --prop color=FFFFFF --prop align=center
done

# Colors: Expert=00B050(dark green) Proficient=92D050(light green) Familiar=FFC000(yellow) Beginner=FF0000(red)
fill_skill_row() {
    local row=$1 person=$2; shift 2
    officecli set "$DOCX" "/body/tbl[3]/tr[$row]/tc[1]" --prop text="$person" --prop bold=true --prop shd=D6DCE4 --prop align=center
    local col=2
    for cell in "$@"; do
        local text="${cell%%:*}" color="${cell#*:}"
        officecli set "$DOCX" "/body/tbl[3]/tr[$row]/tc[$col]" --prop text="$text" --prop shd="$color" --prop color=FFFFFF --prop align=center --prop bold=true
        ((col++))
    done
}
fill_skill_row 2 John   Expert:00B050 Proficient:92D050 Familiar:FFC000 Expert:00B050 Familiar:FFC000 Expert:00B050
fill_skill_row 3 Sarah  Proficient:92D050 Expert:00B050 Expert:00B050 Proficient:92D050 Familiar:FFC000 Beginner:FF0000
fill_skill_row 4 Mike   Familiar:FFC000 Familiar:FFC000 Expert:00B050 Familiar:FFC000 Expert:00B050 Proficient:92D050
fill_skill_row 5 Emily  Expert:00B050 Beginner:FF0000 Familiar:FFC000 Expert:00B050 Proficient:92D050 Familiar:FFC000
fill_skill_row 6 David  Proficient:92D050 Proficient:92D050 Proficient:92D050 Expert:00B050 Expert:00B050 Expert:00B050

# -- Table 4: Property Coverage Table (missing table/row/cell props) --
echo "  -> Table 4: Property coverage — border/layout/direction/cell-formatting"
officecli add "$DOCX" /body --type paragraph --prop text=""
officecli add "$DOCX" /body --type paragraph --prop text="4. Property Coverage (border / layout / direction / cell formatting)" --prop style=Heading2

# Table with border.all + cellSpacing + colWidths + direction + indent + layout + padding
officecli add "$DOCX" /body --type table \
    --prop rows=3 --prop cols=4 \
    --prop "border.all=single;8;2E74B5" \
    --prop "colWidths=2500,2500,2500,2500" \
    --prop "cellSpacing=20" \
    --prop "indent=200" \
    --prop "layout=fixed" \
    --prop "padding=80"

# Override outer-edge borders after creation
officecli set "$DOCX" '/body/tbl[4]' --prop "border.top=double;8;1F3864"
officecli set "$DOCX" '/body/tbl[4]' --prop "border.bottom=double;8;1F3864"
officecli set "$DOCX" '/body/tbl[4]' --prop "border.left=double;8;1F3864"
officecli set "$DOCX" '/body/tbl[4]' --prop "border.right=double;8;1F3864"
officecli set "$DOCX" '/body/tbl[4]' --prop "border.horizontal=single;4;9DC3E6"
officecli set "$DOCX" '/body/tbl[4]' --prop "border.vertical=single;4;9DC3E6"

# Header row: header=true + height.exact
officecli set "$DOCX" '/body/tbl[4]/tr[1]' --prop header=true --prop height.exact=400
officecli set "$DOCX" '/body/tbl[4]/tr[1]/tc[1]' --prop text="Cell Borders" --prop bold=true --prop fill=2E74B5 --prop color=FFFFFF
officecli set "$DOCX" '/body/tbl[4]/tr[1]/tc[2]' --prop text="Run Formatting" --prop bold=true --prop fill=2E74B5 --prop color=FFFFFF
officecli set "$DOCX" '/body/tbl[4]/tr[1]/tc[3]' --prop text="Merge / Flow" --prop bold=true --prop fill=2E74B5 --prop color=FFFFFF
officecli set "$DOCX" '/body/tbl[4]/tr[1]/tc[4]' --prop text="Padding / Grid" --prop bold=true --prop fill=2E74B5 --prop color=FFFFFF

# Data row 2: cell borders (border.all, tl2br, tr2bl), direction, nowrap
officecli set "$DOCX" '/body/tbl[4]/tr[2]/tc[1]' \
    --prop text="border.all + tl2br + tr2bl" \
    --prop "border.all=single;8;FF0000" \
    --prop "border.tl2br=single;4;0000FF" \
    --prop "border.tr2bl=single;4;0000FF"
officecli set "$DOCX" '/body/tbl[4]/tr[2]/tc[2]' \
    --prop text="font + italic + strike + underline + highlight" \
    --prop "font=Times New Roman" \
    --prop italic=true \
    --prop strike=true \
    --prop underline=single \
    --prop highlight=yellow
officecli set "$DOCX" '/body/tbl[4]/tr[2]/tc[3]' \
    --prop text="direction=rtl + nowrap + textDirection=btlr" \
    --prop direction=rtl \
    --prop nowrap=true \
    --prop textDirection=btlr
officecli set "$DOCX" '/body/tbl[4]/tr[2]/tc[4]' \
    --prop text="padding per side + skipGridSync" \
    --prop padding.top=50 \
    --prop padding.bottom=150 \
    --prop padding.left=80 \
    --prop padding.right=80

# Data row 3: border.top/bottom/left/right per cell, fitText, skipGridSync
officecli set "$DOCX" '/body/tbl[4]/tr[3]/tc[1]' \
    --prop text="border.top + border.bottom" \
    --prop "border.top=single;8;FF0000" \
    --prop "border.bottom=single;8;0000FF"
officecli set "$DOCX" '/body/tbl[4]/tr[3]/tc[2]' \
    --prop text="border.left + border.right" \
    --prop "border.left=single;8;00FF00" \
    --prop "border.right=single;8;FF00FF"
officecli set "$DOCX" '/body/tbl[4]/tr[3]/tc[3]' \
    --prop text="fitText squeezes text to cell width" \
    --prop fitText=true
officecli set "$DOCX" '/body/tbl[4]/tr[3]/tc[4]' \
    --prop text="width + skipGridSync" \
    --prop width=2500 \
    --prop skipGridSync=true

# Demonstrate hmerge (horizontal merge) in a separate small 3-col table
# hmerge=restart on tc[1] spans 2 cols and absorbs tc[2]; tc[3]→tc[2] after
officecli add "$DOCX" /body --type table \
    --prop rows=2 --prop cols=3 \
    --prop "border.all=single;4;808080"
# Set the non-merged cell before applying hmerge=restart (which removes tc[2])
officecli set "$DOCX" '/body/tbl[5]/tr[1]/tc[3]' --prop text="normal tc"
officecli set "$DOCX" '/body/tbl[5]/tr[1]/tc[1]' \
    --prop text="hmerge restart (spans 2 cols)" \
    --prop hmerge=restart
# After hmerge=restart, original tc[3] is now tc[2]
officecli set "$DOCX" '/body/tbl[5]/tr[2]/tc[1]' --prop text="row 2 col 1"
officecli set "$DOCX" '/body/tbl[5]/tr[2]/tc[2]' --prop text="row 2 col 2"
officecli set "$DOCX" '/body/tbl[5]/tr[2]/tc[3]' --prop text="row 2 col 3"

# Also demonstrate table direction=rtl on a separate small table
officecli add "$DOCX" /body --type table \
    --prop rows=2 --prop cols=2 \
    --prop "direction=rtl" \
    --prop "border.all=single;8;C00000"
officecli set "$DOCX" '/body/tbl[6]/tr[1]/tc[1]' --prop text="RTL table" --prop bold=true
officecli set "$DOCX" '/body/tbl[6]/tr[1]/tc[2]' --prop text="column order mirrored"
officecli set "$DOCX" '/body/tbl[6]/tr[2]/tc[1]' --prop text="row 2 col 1"
officecli set "$DOCX" '/body/tbl[6]/tr[2]/tc[2]' --prop text="row 2 col 2"

# data — inline shorthand: rows separated by ';', cells by ',' (builds the whole
# grid in one prop instead of rows+cols+per-cell set commands)
officecli add "$DOCX" /body --type table \
    --prop "data=Region,Q1,Q2;North,120,150;South,90,110" \
    --prop "border.all=single;4;808080"

officecli validate "$DOCX"
officecli close "$DOCX"
echo "  Done: Word document: $DOCX"

###############################################################################
# 2. Excel Sales Report
###############################################################################
XLSX="$DIR/tables.xlsx"
echo ""
echo "=========================================="
echo "Generating Excel sales report: $XLSX"
echo "=========================================="

rm -f "$XLSX"
officecli create "$XLSX"
officecli open "$XLSX"

# Sheet1: Sales Data
echo "  -> Sheet1: Sales Data"
officecli set "$XLSX" '/Sheet1/A1' --prop value="2025 Annual Sales Report"
officecli set "$XLSX" '/Sheet1/A2' --prop value="Department"
officecli set "$XLSX" '/Sheet1/B2' --prop value="Q1"
officecli set "$XLSX" '/Sheet1/C2' --prop value="Q2"
officecli set "$XLSX" '/Sheet1/D2' --prop value="Q3"
officecli set "$XLSX" '/Sheet1/E2' --prop value="Q4"
officecli set "$XLSX" '/Sheet1/F2' --prop value="Annual Total"

for entry in "3:Engineering:128000:156000:189000:210000" \
             "4:Marketing:95000:112000:138000:165000" \
             "5:Operations:76000:89000:102000:118000" \
             "6:Sales:230000:275000:310000:356000" \
             "7:HR:45000:48000:52000:55000"; do
    IFS=':' read -r row dept q1 q2 q3 q4 <<< "$entry"
    officecli set "$XLSX" "/Sheet1/A$row" --prop value="$dept"
    officecli set "$XLSX" "/Sheet1/B$row" --prop value="$q1"
    officecli set "$XLSX" "/Sheet1/C$row" --prop value="$q2"
    officecli set "$XLSX" "/Sheet1/D$row" --prop value="$q3"
    officecli set "$XLSX" "/Sheet1/E$row" --prop value="$q4"
    officecli set "$XLSX" "/Sheet1/F$row" --prop formula="SUM(B${row}:E${row})"
done

# Total row
officecli set "$XLSX" '/Sheet1/A8' --prop value="Total"
for col in B C D E F; do
    officecli set "$XLSX" "/Sheet1/${col}8" --prop formula="SUM(${col}3:${col}7)"
done

# Growth rate
officecli set "$XLSX" '/Sheet1/A9' --prop value="Quarterly Growth Rate"
officecli set "$XLSX" '/Sheet1/C9' --prop formula="(C8-B8)/B8"
officecli set "$XLSX" '/Sheet1/D9' --prop formula="(D8-C8)/C8"
officecli set "$XLSX" '/Sheet1/E9' --prop formula="(E8-D8)/D8"

# Sheet2: Employee Performance
echo "  -> Sheet2: Performance"
officecli add "$XLSX" / --type sheet --prop name="Performance"

officecli set "$XLSX" '/Performance/A1' --prop value="Employee Performance Review"
officecli set "$XLSX" '/Performance/A2' --prop value="Name"
officecli set "$XLSX" '/Performance/B2' --prop value="Department"
officecli set "$XLSX" '/Performance/C2' --prop value="Performance Score"
officecli set "$XLSX" '/Performance/D2' --prop value="Capability Score"
officecli set "$XLSX" '/Performance/E2' --prop value="Attitude Score"
officecli set "$XLSX" '/Performance/F2' --prop value="Total Score"
officecli set "$XLSX" '/Performance/G2' --prop value="Grade"

declare -a EMP_DATA=(
    "3:John:Engineering:92:88:95"
    "4:Sarah:Marketing:85:90:78"
    "5:Mike:Operations:78:82:90"
    "6:Emily:Sales:96:75:88"
    "7:David:Engineering:88:92:85"
    "8:Lisa:HR:72:85:92"
    "9:Tom:Sales:91:78:80"
    "10:Amy:Marketing:65:70:88"
    "11:Chris:Engineering:95:93:90"
    "12:Kate:Operations:80:86:75"
)

for emp in "${EMP_DATA[@]}"; do
    IFS=':' read -r row name dept s1 s2 s3 <<< "$emp"
    officecli set "$XLSX" "/Performance/A$row" --prop value="$name"
    officecli set "$XLSX" "/Performance/B$row" --prop value="$dept"
    officecli set "$XLSX" "/Performance/C$row" --prop value="$s1"
    officecli set "$XLSX" "/Performance/D$row" --prop value="$s2"
    officecli set "$XLSX" "/Performance/E$row" --prop value="$s3"
    officecli set "$XLSX" "/Performance/F$row" --prop formula="C${row}*0.4+D${row}*0.35+E${row}*0.25"
    officecli set "$XLSX" "/Performance/G$row" --prop formula="IF(F${row}>=90,\"A\",IF(F${row}>=80,\"B\",IF(F${row}>=70,\"C\",\"D\")))"
done

# Sheet3: Summary
echo "  -> Sheet3: Summary"
officecli add "$XLSX" / --type sheet --prop name="Summary"

officecli set "$XLSX" '/Summary/A1' --prop value="Metric"
officecli set "$XLSX" '/Summary/B1' --prop value="Value"
officecli set "$XLSX" '/Summary/A2' --prop value="Highest Score"
officecli set "$XLSX" '/Summary/B2' --prop formula="MAX(Performance!F3:F12)"
officecli set "$XLSX" '/Summary/A3' --prop value="Lowest Score"
officecli set "$XLSX" '/Summary/B3' --prop formula="MIN(Performance!F3:F12)"
officecli set "$XLSX" '/Summary/A4' --prop value="Average Score"
officecli set "$XLSX" '/Summary/B4' --prop formula="AVERAGE(Performance!F3:F12)"
officecli set "$XLSX" '/Summary/A5' --prop value="Grade A Count"
officecli set "$XLSX" '/Summary/B5' --prop formula="COUNTIF(Performance!G3:G12,\"A\")"
officecli set "$XLSX" '/Summary/A6' --prop value="Annual Total Sales"
officecli set "$XLSX" '/Summary/B6' --prop formula="Sheet1!F8"

officecli close "$XLSX"
echo "  Done: Excel document: $XLSX"

###############################################################################
# 3. PowerPoint Data Report
###############################################################################
PPTX="$DIR/tables.pptx"
echo ""
echo "=========================================="
echo "Generating PowerPoint data report: $PPTX"
echo "=========================================="

rm -f "$PPTX"
officecli create "$PPTX"
officecli open "$PPTX"

# Slide 1: Title Page — HIGH-LEVEL (slide background + two text shapes)
echo "  -> Slide 1: Title Page"
officecli add "$PPTX" / --type slide
officecli set "$PPTX" '/slide[1]' --prop background=1F3864
officecli add "$PPTX" '/slide[1]' --type shape --prop text="2025 Annual Data Analysis Report" \
  --prop x=1500000 --prop y=2000000 --prop width=9192000 --prop height=1200000 \
  --prop size=40 --prop bold=true --prop color=FFFFFF --prop align=center --prop valign=middle
officecli add "$PPTX" '/slide[1]' --type shape --prop text="Dept Comparison | Performance Overview | Financial Summary" \
  --prop x=2500000 --prop y=3500000 --prop width=7192000 --prop height=800000 \
  --prop size=20 --prop color=BDD7EE --prop align=center --prop valign=middle

# Slide 2: Data Table — HIGH-LEVEL (title shape + styled table via add table + per-cell set).
# Cell values carry thousands separators ("128,000"), so they can't go through the
# comma-delimited `data=` seed — the table is created empty (rows/cols) and each cell
# is set individually. headerFill styles row 0; first column gets a light-blue band.
echo "  -> Slide 2: Data Table"
officecli add "$PPTX" / --type slide
officecli add "$PPTX" '/slide[2]' --type shape --prop text="Quarterly Sales by Department" \
  --prop x=500000 --prop y=200000 --prop width=11192000 --prop height=600000 \
  --prop size=28 --prop bold=true --prop color=1F3864
officecli add "$PPTX" '/slide[2]' --type table --prop rows=6 --prop cols=5 \
  --prop headerFill=2E75B6 --prop firstRow=true \
  --prop x=500000 --prop y=1000000 --prop width=11192000 --prop height=4500000

# Header row (white, bold, centered).
c=1
for h in Department Q1 Q2 Q3 Q4; do
  officecli set "$PPTX" "/slide[2]/table[1]/cell[1,$c]" --prop text="$h" --prop bold=true --prop color=FFFFFF --prop align=center
  c=$((c+1))
done

# Body rows: "label|Q1|Q2|Q3|Q4" ('|' avoids clashing with the values' commas).
# First column carries a light-blue band fill; the rest are plain centered.
r_idx=2
for row in \
  "Engineering|128,000|156,000|189,000|210,000" \
  "Marketing|95,000|112,000|138,000|165,000" \
  "Operations|76,000|89,000|102,000|118,000" \
  "Sales|230,000|275,000|310,000|356,000" \
  "HR|45,000|48,000|52,000|55,000"; do
  IFS='|' read -r label q1 q2 q3 q4 <<< "$row"
  officecli set "$PPTX" "/slide[2]/table[1]/cell[$r_idx,1]" --prop text="$label" --prop fill=DEEAF6 --prop align=center
  officecli set "$PPTX" "/slide[2]/table[1]/cell[$r_idx,2]" --prop text="$q1" --prop align=center
  officecli set "$PPTX" "/slide[2]/table[1]/cell[$r_idx,3]" --prop text="$q2" --prop align=center
  officecli set "$PPTX" "/slide[2]/table[1]/cell[$r_idx,4]" --prop text="$q3" --prop align=center
  officecli set "$PPTX" "/slide[2]/table[1]/cell[$r_idx,5]" --prop text="$q4" --prop align=center
  r_idx=$((r_idx+1))
done

# Slide 3: Pie Chart Analysis
echo "  -> Slide 3: Pie Chart Analysis"
officecli add "$PPTX" / --type slide
officecli add "$PPTX" '/slide[3]' --type shape --prop text="Annual Sales Share by Department" --prop size=28 --prop bold=true --prop x=500000 --prop y=200000 --prop width=11192000 --prop height=600000
officecli add "$PPTX" '/slide[3]' --type shape --prop text="Engineering 683,000 (24.4%)" --prop x=1000000 --prop y=1200000 --prop width=10000000 --prop height=500000
officecli add "$PPTX" '/slide[3]' --type shape --prop text="Marketing 510,000 (18.2%)" --prop x=1000000 --prop y=1900000 --prop width=10000000 --prop height=500000
officecli add "$PPTX" '/slide[3]' --type shape --prop text="Operations 385,000 (13.7%)" --prop x=1000000 --prop y=2600000 --prop width=10000000 --prop height=500000
officecli add "$PPTX" '/slide[3]' --type shape --prop text="Sales 1,171,000 (41.8%)" --prop x=1000000 --prop y=3300000 --prop width=10000000 --prop height=500000
officecli add "$PPTX" '/slide[3]' --type shape --prop text="HR 200,000 (7.1%)" --prop x=1000000 --prop y=4000000 --prop width=10000000 --prop height=500000

officecli close "$PPTX"
echo "  Done: PowerPoint document: $PPTX"

###############################################################################
# Verification
###############################################################################
echo ""
echo "=========================================="
echo "Verifying all files"
echo "=========================================="
officecli view "$DOCX" outline
echo ""
officecli view "$XLSX" outline
echo ""
officecli view "$PPTX" outline
echo ""
ls -lh "$DOCX" "$XLSX" "$PPTX"
echo ""
echo "All done!"

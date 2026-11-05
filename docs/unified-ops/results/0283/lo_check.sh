#!/bin/bash
# Build one DOCX with every classic and extended chart type, and one PPTX
# and one XLSX with every extended type, through the litchi CLI; convert them
# with headless LibreOffice and report which chart titles are rendered.
# Usage: lo_check.sh <litchi binary> <out dir>
set -u
L="$1"; OUT="$2"; mkdir -p "$OUT"; rm -f "$OUT"/*.docx "$OUT"/*.pptx "$OUT"/*.xlsx "$OUT"/*.pdf
D="$OUT/charts.docx"; P="$OUT/slides.pptx"; X="$OUT/sheets.xlsx"
CLASSIC="column bar line pie doughnut area scatter bubble radar stock combo stackedColumn percentStackedBar column3d pie3d"
EXTENDED="waterfall funnel treemap sunburst histogram pareto boxWhisker"
data_for() {
  case $1 in
    pie|pie3d|doughnut) echo "Share:30,25,28,17";;
    stock) echo "High:130,135,140,138;Low:118,122,128,125;Close:125,130,135,132";;
    waterfall) echo "Cash:100,-30,50,120";;
    histogram|boxWhisker) echo "First:1,2,2,3,9,4;Second:4,5,5,6,1,3";;
    funnel|treemap|sunburst|pareto) echo "Values:50,30,15,5";;
    *) echo "East:120,135,148,162;West:95,108,115,128";;
  esac
}
"$L" create "$D" >/dev/null; "$L" create "$P" >/dev/null; "$L" create "$X" >/dev/null
n=0; fail=0
for t in $CLASSIC $EXTENDED; do
  n=$((n+1))
  "$L" add "$D" /body --type paragraph --prop text="Chart $n" >/dev/null || fail=1
  "$L" add "$D" "/body/p[$n]" --type chart --prop chartType=$t --prop title="D-$t" --prop categories=Q1,Q2,Q3,Q4 --prop data="$(data_for $t)" >/dev/null || { echo "docx add $t failed"; fail=1; }
done
m=0
for t in $EXTENDED; do
  m=$((m+1))
  "$L" add "$P" / --type slide --prop layout=Blank >/dev/null || fail=1
  "$L" add "$P" "/slide[$m]" --type chart --prop chartType=$t --prop title="P-$t" --prop categories=Q1,Q2,Q3,Q4 --prop data="$(data_for $t)" >/dev/null || { echo "pptx add $t failed"; fail=1; }
  "$L" add "$X" /Sheet1 --type chart --prop chartType=$t --prop title="X-$t" --prop categories=Q1,Q2,Q3,Q4 --prop data="$(data_for $t)" --prop x=0 --prop y=$(( (m-1)*16 )) --prop width=8 --prop height=15 >/dev/null || { echo "xlsx add $t failed"; fail=1; }
done
for f in "$D" "$P" "$X"; do
  timeout 300 soffice -env:UserInstallation=file://$OUT/profile --headless --convert-to pdf --outdir "$OUT" "$f" >/dev/null 2>&1 || { echo "convert $f failed"; fail=1; }
done
rendered() { pdftotext "$OUT/$1" - 2>/dev/null | tr -d "\f" | grep -qx "$2"; }
list() { local out=""; for t in $2; do rendered "$1" "$3-$t" && out="$out\"$t\","; done; echo "[${out%,}]"; }
echo "{\"docx_classic_rendered\": $(list charts.pdf "$CLASSIC" D), \"docx_extended_rendered\": $(list charts.pdf "$EXTENDED" D), \"pptx_extended_rendered\": $(list slides.pdf "$EXTENDED" P), \"xlsx_extended_rendered\": $(list sheets.pdf "$EXTENDED" X), \"pdfs\": [$(ls "$OUT"/*.pdf 2>/dev/null | xargs -n1 basename | sed 's/.*/"&"/' | paste -sd,)], \"add_failures\": $fail}"

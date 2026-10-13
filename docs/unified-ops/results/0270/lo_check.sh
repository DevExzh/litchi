#!/bin/bash
# Build one PPTX and one XLSX with every authored chart variant through the
# litchi CLI, convert both with headless LibreOffice and check that every
# chart title is rendered. Usage: lo_check.sh <litchi binary> <out dir>
set -u
L="$1"; OUT="$2"; mkdir -p "$OUT"; rm -f "$OUT"/*.pptx "$OUT"/*.xlsx "$OUT"/*.pdf
P="$OUT/charts.pptx"; X="$OUT/sheets.xlsx"
TYPES="column bar line pie doughnut area scatter bubble radar stock combo stackedColumn percentStackedBar column3d bar3d line3d pie3d area3d stackedArea percentStackedLine"
"$L" create "$P" >/dev/null; "$L" create "$X" >/dev/null
n=0; fail=0
for t in $TYPES; do
  n=$((n+1))
  data="East:120,135,148,162;West:95,108,115,128"
  case $t in pie|pie3d|doughnut) data="Share:30,25,28,17";; stock) data="High:130,135,140,138;Low:118,122,128,125;Close:125,130,135,132";; esac
  "$L" add "$P" / --type slide --prop layout=Blank >/dev/null || fail=1
  "$L" add "$P" "/slide[$n]" --type chart --prop chartType=$t --prop title="T-$t" --prop categories=Q1,Q2,Q3,Q4 --prop data="$data" --prop colors=4472C4,ED7D31 --prop dataLabels=value >/dev/null || { echo "pptx add $t failed"; fail=1; }
  "$L" add "$X" /Sheet1 --type chart --prop chartType=$t --prop title="X-$t" --prop categories=Q1,Q2,Q3,Q4 --prop data="$data" --prop x=0 --prop y=$(( (n-1)*16 )) --prop width=8 --prop height=15 >/dev/null || { echo "xlsx add $t failed"; fail=1; }
done
for f in "$P" "$X"; do
  timeout 300 soffice -env:UserInstallation=file://$OUT/profile --headless --convert-to pdf --outdir "$OUT" "$f" >/dev/null 2>&1 || { echo "convert $f failed"; fail=1; }
done
missing=0
for t in $TYPES; do
  pdftotext "$OUT/charts.pdf" - | tr -d "\f" | grep -qx "T-$t" || { echo "pptx title T-$t not rendered"; missing=$((missing+1)); }
done
xmissing=0
for t in $TYPES; do
  pdftotext "$OUT/sheets.pdf" - | tr -d "\f" | grep -qx "X-$t" || { echo "xlsx title X-$t not rendered"; xmissing=$((xmissing+1)); }
done
echo "{\"variants\": $n, \"pptx_titles_missing\": $missing, \"xlsx_titles_missing\": $xmissing, \"add_failures\": $fail}"

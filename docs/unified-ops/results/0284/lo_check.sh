#!/bin/bash
# Build one PPTX with every authored SmartArt layout through the litchi CLI,
# edit nodes (text, add, remove, layout change), edit a foreign SmartArt's
# text in place, convert with headless LibreOffice and check that every node
# text is rendered. Usage: lo_check.sh <litchi binary> <out dir> <foreign.pptx>
set -u
L="$1"; OUT="$2"; FOREIGN="$3"; mkdir -p "$OUT"; rm -f "$OUT"/*.pptx "$OUT"/*.pdf "$OUT"/*.txt
P="$OUT/smartart.pptx"
LAYOUTS="list process cycle hierarchy pyramid matrix venn"
"$L" create "$P" >/dev/null
n=0; fail=0; missing=0
for l in $LAYOUTS; do
  n=$((n+1))
  "$L" add "$P" / --type slide --prop layout=Blank >/dev/null || fail=1
  "$L" add "$P" "/slide[$n]" --type smartart --prop layout=$l --prop text="$l-A
	$l-A1
$l-B
$l-C" --prop width=8in --prop height=4.5in >/dev/null || { echo "add $l failed"; fail=1; }
done
# Edits on slide 2 (process): text, add, remove, then on slide 3 a layout change.
"$L" set "$P" "/slide[2]/smartart[1]/node[2]" --prop text=process-B2 >/dev/null || fail=1
"$L" add "$P" "/slide[2]/smartart[1]" --type node --prop text=process-D >/dev/null || fail=1
"$L" remove "$P" "/slide[2]/smartart[1]/node[3]" >/dev/null || fail=1
"$L" set "$P" "/slide[3]/smartart[1]" --prop layout=matrix >/dev/null || fail=1
F="$OUT/foreign.pptx"; cp "$FOREIGN" "$F"
"$L" set "$F" "/slide[1]/smartart[1]/node[2]" --prop text=EDITED >/dev/null || fail=1
for f in "$P" "$F"; do
  timeout 300 soffice -env:UserInstallation=file://$OUT/profile --headless --convert-to pdf --outdir "$OUT" "$f" >/dev/null 2>&1 || { echo "convert $f failed"; fail=1; }
done
pdftotext "$OUT/smartart.pdf" "$OUT/smartart.txt"; pdftotext "$OUT/foreign.pdf" "$OUT/foreign.txt"
expect() { grep -qw -- "$2" "$OUT/$1.txt" || { echo "$1: '$2' not rendered"; missing=$((missing+1)); }; }
absent() { grep -qw -- "$2" "$OUT/$1.txt" && { echo "$1: '$2' rendered but removed"; missing=$((missing+1)); }; }
for l in list cycle hierarchy pyramid matrix venn; do expect smartart "$l-A"; expect smartart "$l-A1"; expect smartart "$l-B"; expect smartart "$l-C"; done
expect smartart process-A; expect smartart process-A1; expect smartart process-B2; expect smartart process-D
absent smartart process-C
expect foreign EDITED; absent foreign cc
echo "{\"layouts\": $n, \"texts_missing\": $missing, \"failures\": $fail}"

#!/usr/bin/env bash
# Replays a few command lines of OfficeCLI's example scripts through `litchi`
# (record 0256). It is the seed of the OfficeCLI example-corpus harness of
# plan step 17: each replayed line is cited as `<file>:<line>` of the copies
# vendored under `test-data/officecli/examples/` (record 0257; provenance in
# `test-data/officecli/PROVENANCE.md`) and kept as OfficeCLI wrote it, with
# `officecli` replaced by the binary under test. Every `step` checks that its
# cited line exists and runs the same verb. A line whose status is not 0 is a
# recorded delta; the status says how Litchi refuses it.
#
# Usage: officecli_smoke.sh <litchi binary> <scratch directory>
set -u

CLI=${1:?usage: officecli_smoke.sh <litchi binary> <scratch directory>}
DIR=${2:?usage: officecli_smoke.sh <litchi binary> <scratch directory>}
EXAMPLES="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.." && pwd)/test-data/officecli/examples"
failures=0

# cited <citation> <verb>: the vendored example line must run the same verb.
cited() {
    local citation=$1 verb=$2
    local file=${citation%%:*} number=${citation##*:}
    local line
    line=$(sed -n "${number}p" "$EXAMPLES/$file" 2>/dev/null)
    if [[ " $line " != *" $verb "* ]]; then
        echo "FAIL $citation: the vendored example line does not run '$verb': $line" >&2
        failures=$((failures + 1))
    fi
}

# step <expected status> <citation> <arguments...>
step() {
    local expected=$1 citation=$2
    shift 2
    cited "$citation" "$1"
    local status=0
    "$CLI" "$@" >"$DIR/last.out" 2>"$DIR/last.err" || status=$?
    if [ "$status" -ne "$expected" ]; then
        echo "FAIL $citation: expected $expected, got $status: $*" >&2
        cat "$DIR/last.err" >&2
        failures=$((failures + 1))
    fi
}

# check <citation> <text stdout must contain> <arguments...>
check() {
    local citation=$1 expected=$2
    shift 2
    local actual
    actual=$("$CLI" "$@" 2>&1)
    if [[ "$actual" != *"$expected"* ]]; then
        echo "FAIL $citation: expected '$expected', got '$actual'" >&2
        failures=$((failures + 1))
    fi
}

# --- examples/word/run-formatting.sh -------------------------------------
DOCX="$DIR/run-formatting.docx"
step 0 word/run-formatting.sh:18 create "$DOCX"
heading() { step 0 word/run-formatting.sh:20 add "$DOCX" /body --type paragraph --prop "text=$1" --prop bold=true --prop size=14 --prop color=1F4E79 --prop spaceBefore=8pt; }
step 0 word/run-formatting.sh:22 add "$DOCX" /body --type paragraph --prop "text=Run / Character Formatting Showcase" --prop align=center --prop bold=true --prop size=20
heading "Weight & style"
step 0 word/run-formatting.sh:26 add "$DOCX" /body --type paragraph --prop "text=Bold text" --prop bold=true
step 0 word/run-formatting.sh:27 add "$DOCX" /body --type paragraph --prop "text=Italic text" --prop italic=true
step 0 word/run-formatting.sh:28 add "$DOCX" /body --type paragraph --prop "text=Bold + italic" --prop bold=true --prop italic=true
step 0 word/run-formatting.sh:36 add "$DOCX" /body --type paragraph --prop "text=wave (red)" --prop underline=wave --prop underline.color=FF0000
step 0 word/run-formatting.sh:41 add "$DOCX" /body --type paragraph --prop "text=double strike" --prop dstrike=true
step 0 word/run-formatting.sh:46 add "$DOCX" /body --type paragraph --prop "text=small caps rendering" --prop smallcaps=true
step 0 word/run-formatting.sh:50 add "$DOCX" /body --type paragraph --prop "text=E = mc"
step 0 word/run-formatting.sh:51 add "$DOCX" "/body/p[last()]" --type run --prop "text=2" --prop superscript=true
step 0 word/run-formatting.sh:58 add "$DOCX" /body --type paragraph --prop "text=Red 16pt" --prop color=C00000 --prop size=16
step 0 word/run-formatting.sh:59 add "$DOCX" /body --type paragraph --prop "text=Highlighted" --prop highlight=yellow
step 0 word/run-formatting.sh:67 add "$DOCX" /body --type paragraph --prop "text=emboss" --prop emboss=true
step 0 word/run-formatting.sh:74 add "$DOCX" /body --type paragraph --prop "text=expanded spacing" --prop charSpacing=2pt
# Deltas: per-script fonts and language tags are not in Litchi's paragraph
# add vocabulary yet; the strict default refuses them (unsupported_property).
step 4 word/run-formatting.sh:63 add "$DOCX" /body --type paragraph --prop "text=Latin Georgia + CJK 宋体" --prop font.latin=Georgia --prop font.eastAsia=SimSun --prop size=14
step 4 word/run-formatting.sh:79 add "$DOCX" /body --type paragraph --prop "text=Tagged en-US for spellcheck" --prop lang=en-US
check word/run-formatting.sh:51 '"text": "2"' get "$DOCX" "/body/p[9]/r[2]" --json --depth 0
check word/run-formatting.sh:26 'Bold text' view "$DOCX" text
check word/run-formatting.sh:26 '"matches": 1,' query "$DOCX" 'p:contains(Bold text)' --json --depth 0

# --- examples/excel/cell-formatting.sh -----------------------------------
FILE="$DIR/cell-formatting.xlsx"
step 0 excel/cell-formatting.sh:26 create "$FILE"
step 0 excel/cell-formatting.sh:32 set "$FILE" /Sheet1/A1 --prop value="Cell font properties" --prop font.bold=true --prop font.size=14 --prop fill=1F4E79 --prop font.color=FFFFFF
step 0 excel/cell-formatting.sh:33 set "$FILE" /Sheet1/A2 --prop value="Property" --prop font.bold=true --prop fill=D9E1F2
step 0 excel/cell-formatting.sh:38 set "$FILE" /Sheet1/B3  --prop value="Georgia serif" --prop font.name=Georgia
step 0 excel/cell-formatting.sh:46 set "$FILE" /Sheet1/B7  --prop value="Red text" --prop font.color=C00000
step 0 excel/cell-formatting.sh:48 set "$FILE" /Sheet1/B8  --prop value="Underlined" --prop underline=single
step 0 excel/cell-formatting.sh:52 set "$FILE" /Sheet1/B10 --prop value="Struck out" --prop strike=true
step 0 excel/cell-formatting.sh:54 set "$FILE" /Sheet1/B11 --prop value="Superscript cell" --prop superscript=true
step 0 excel/cell-formatting.sh:60 set "$FILE" "/Sheet1/col[1]" --prop width=22
step 0 excel/cell-formatting.sh:66 add "$FILE" / --type sheet --prop name=Fills
step 0 excel/cell-formatting.sh:69 set "$FILE" /Fills/A2 --prop value="fill=E63946 (hex)" --prop fill=E63946 --prop font.color=FFFFFF
step 0 excel/cell-formatting.sh:70 set "$FILE" /Fills/A3 --prop value="fill=gold (named)" --prop fill=gold
check excel/cell-formatting.sh:32 '"text": "Cell font properties"' get "$FILE" /Sheet1/A1 --depth 0 --json
check excel/cell-formatting.sh:66 '"preview": "Fills"' get "$FILE" "/sheet[2]" --depth 0 --json

# --- examples/ppt/presentation.sh ----------------------------------------
PPTX="$DIR/presentation.pptx"
step 0 ppt/presentation.sh:9 create "$PPTX"
# Delta: OfficeCLI picks a layout for a bare slide; Litchi requires an exact
# layout name (missing_property, exit 4).
step 4 ppt/presentation.sh:18 add "$PPTX" / --type slide

if [ "$failures" -ne 0 ]; then
    echo "$failures replayed lines diverged" >&2
    exit 1
fi
echo "all replayed lines matched"

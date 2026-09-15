#!/bin/bash
# Paired timing for change 0603 on CPU 24, both legs from the identical probe
# source built against the before and after checkouts.
#   A = before (6c4c1469b), B = after (this branch)
#   order A1 B1 B2 A2, then S1..S4 as four before legs so S1/S2 and S3/S4 give
#   the A/A floor in the same window.
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0603
A=/home/zhuhe/code/litchi-worktrees/targets/0603-before/release/xlsx-admission-probe
B=/home/zhuhe/code/litchi-worktrees/targets/0603-after/release/xlsx-admission-probe
F=$S/derived
mkdir -p "$S/bench/legs"
leg() { # binary tag sheetfile sheet cell samples legname
  taskset -c 24 "$1" bench "$F/$3" "$4" "$5" 5 "$6" > "$S/bench/legs/$2-$7.txt" 2>&1
}
row() { # tag file sheet cell samples
  leg "$A" "$1" "$2" "$3" "$4" "$5" A1
  leg "$B" "$1" "$2" "$3" "$4" "$5" B1
  leg "$B" "$1" "$2" "$3" "$4" "$5" B2
  leg "$A" "$1" "$2" "$3" "$4" "$5" A2
  for n in S1 S2 S3 S4; do leg "$A" "$1" "$2" "$3" "$4" "$5" $n; done
  echo "bench done $1"
}
row fct  FormatConditionTests-proj-plain.xlsx      "Flags"          C4   40
row dvtr dataValidationTableRange-proj-plain.xlsx  "County Ranking" C31  40
row sss  sheet-state-show-proj-plain.xlsx          "стр1"           ER25 40
row mfe  MatrixFormulaEvalTestData-proj-plain.xlsx "Sheet1"         A1   40
row ndp  no_drawing_patriarch-proj-plain.xlsx      "Лист 1"         A1   30
echo BENCHDONE

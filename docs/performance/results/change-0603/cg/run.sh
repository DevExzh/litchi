#!/bin/bash
# Callgrind isolation pairs for change 0603, following change 0602's method:
# profiles of N and N+M planning-plus-commit cycles against one retained
# editor differ by exactly M operations, so the open, the process start and
# the harness cancel and the residue is M complete plan-and-commit operations.
# --separate-callers=1 keeps raw::worksheet::parse attributable to its caller.
set -u
LEG=$1          # before | after
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0603
B=/home/zhuhe/code/litchi-worktrees/targets/0603-$LEG/release/xlsx-admission-probe
mkdir -p "$S/cg/$LEG"
run() {
  taskset -c 24 valgrind --tool=callgrind --separate-callers=1 --cache-sim=no --branch-sim=no \
    --callgrind-out-file="$S/cg/$LEG/$1-n$5.out" "$B" edit "$S/derived/$2" "$3" "$4" "$5" \
    > "$S/cg/$LEG/$1-n$5.log" 2>&1
  callgrind_annotate --inclusive=yes --threshold=100 "$S/cg/$LEG/$1-n$5.out" > "$S/cg/$LEG/$1-n$5.incl" 2>/dev/null
  rm -f "$S/cg/$LEG/$1-n$5.out"
}
pair() { run "$1" "$2" "$3" "$4" 1; run "$1" "$2" "$3" "$4" 4; echo "done $1"; }
pair fct  FormatConditionTests-proj-plain.xlsx      "Flags"          C4
pair dvtr dataValidationTableRange-proj-plain.xlsx  "County Ranking" C31
pair sss  sheet-state-show-proj-plain.xlsx          "стр1"           ER25
pair mfe  MatrixFormulaEvalTestData-proj-plain.xlsx "Sheet1"         A1
pair ndp  no_drawing_patriarch-proj-plain.xlsx      "Лист 1"         A1
echo ALLDONE-$LEG

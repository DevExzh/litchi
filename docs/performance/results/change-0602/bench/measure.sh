#!/bin/bash
# Change 0602 measurement driver.
#
# Legs, per fixture, all on CPU 18 with the identical binary:
#   A1 B1 B2 A2   A = a one-cell `set` (change 0525's reduced readback applies)
#                 B = a one-cell `insert` on an absent row, which leaves the
#                     rewrite's omission list empty and so takes the complete
#                     candidate parse -- the only in-tree route that disables
#                     the reduced readback.
#   S1 S2 S3 S4   four `set` legs, so S1/S2 and S3/S4 give the A/A floor in the
#                 same window.
# Then callgrind isolation pairs (N=1 and N=4, M=3) for both kinds.
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0602
B=/home/zhuhe/code/litchi-worktrees/targets/0602-probe/release/xlsx-admission-probe
F=$S/derived
bench() { # tag file sheet a1 samples leg kind
  taskset -c 18 "$B" bench "$F/$2" "$3" "$4" 5 "$5" $7 > "$S/bench/$1-$6.txt" 2>&1
}
row() { # tag file sheet setcell insertcell samples
  for leg in A1 B1 B2 A2; do
    case $leg in A*) k="" ; c=$4 ;; B*) k="--insert"; c=$5 ;; esac
    bench "$1" "$2" "$3" "$c" "$6" "$leg" "$k"
  done
  for leg in S1 S2 S3 S4; do bench "$1" "$2" "$3" "$4" "$6" "$leg" ""; done
  echo "bench done $1"
}
row fct  FormatConditionTests-proj-plain.xlsx      "Flags"          C4   A99    40
row dvtr dataValidationTableRange-proj-plain.xlsx  "County Ranking" C31  A999   40
row sss  sheet-state-show-proj-plain.xlsx          "стр1"           ER25 A9999  40
row ndp  no_drawing_patriarch-proj-plain.xlsx      "Лист 1"         A1   A99999 30
echo BENCHDONE

cg() { # tag file sheet cell n kind suffix
  taskset -c 18 valgrind --tool=callgrind --separate-callers=1 --cache-sim=no --branch-sim=no \
    --callgrind-out-file="$S/cg/$1$7-n$5.out" "$B" edit "$F/$2" "$3" "$4" "$5" $6 \
    > "$S/cg/$1$7-n$5.log" 2>&1
  callgrind_annotate --inclusive=yes --threshold=100 "$S/cg/$1$7-n$5.out" > "$S/cg/$1$7-n$5.incl" 2>/dev/null
  rm -f "$S/cg/$1$7-n$5.out"
}
pair() { # tag file sheet setcell insertcell
  cg "$1" "$2" "$3" "$4" 1 ""         "-set";    cg "$1" "$2" "$3" "$4" 4 ""         "-set"
  cg "$1" "$2" "$3" "$5" 1 "--insert" "-insert"; cg "$1" "$2" "$3" "$5" 4 "--insert" "-insert"
  echo "cg done $1"
}
pair fct  FormatConditionTests-proj-plain.xlsx      "Flags"          C4   A99
pair dvtr dataValidationTableRange-proj-plain.xlsx  "County Ranking" C31  A999
pair sss  sheet-state-show-proj-plain.xlsx          "стр1"           ER25 A9999
pair ndp  no_drawing_patriarch-proj-plain.xlsx      "Лист 1"         A1   A99999
echo ALLDONE

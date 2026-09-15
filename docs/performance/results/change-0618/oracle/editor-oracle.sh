#!/usr/bin/env bash
# Editor-level oracle for change 0618: the real `tabs`, `edit_cells` and
# `append_plain_paragraph` routes over the whole OOXML fixture corpus, one
# line per fixture and operation carrying the output digest or the typed error.
set -u
S="$1"; LEG="$2"; BIN="$3"
TD=/home/zhuhe/code/litchi-worktrees/before-8fe9efa55/test-data
WORK=$(mktemp -d "${S}/editor-work-XXXXXX")
trap 'rm -rf "$WORK"' EXIT
digest() { [ -f "$1" ] && sha256sum "$1" | cut -c1-32 || echo none; }
run() { # run <label> <fixture> -- <argv...>
  local label="$1" fixture="$2"; shift 3
  rm -f "$WORK/out."*
  local err
  err=$("$@" 2>&1 >/dev/null) || true
  local out
  out=$(ls "$WORK"/out.* 2>/dev/null | head -1)
  if [ -n "$out" ]; then
    echo "$LEG $label ${fixture#$TD/} -> $(digest "$out")"
  else
    echo "$LEG $label ${fixture#$TD/} -> none :: ${err//$'\n'/ | }"
  fi
}
for f in $(find "$TD" \( -name '*.xlsx' -o -name '*.xlsm' \) | sort); do
  sheet=$(basename "$f" | sed 's/\..*//')
  run "tabs-hide-first" "$f" -- "$BIN/tabs" "$f" "$WORK/out.xlsx" Sheet1 hide
  run "edit-cells"      "$f" -- "$BIN/edit_cells" "$f" "$WORK/out.xlsx"
done
for f in $(find "$TD" \( -name '*.docx' -o -name '*.docm' \) | sort); do
  run "docx-append-paragraph" "$f" -- "$BIN/append_plain_paragraph" "$f" "$WORK/out.docx" "change 0618 paragraph"
done

#!/usr/bin/env bash
# Change 0744 matched phase profiles: frame-pointer `profiling` builds of both
# legs (identical flags), CPU 12. perf.data and scripts are deleted after the
# summaries are written; only the summaries are retained.
set -euo pipefail
declare -A BIN=(
  [before]=/home/zhuhe/code/litchi-worktrees/targets/0744-before/profiling/litchi-perf-baseline
  [after]=/home/zhuhe/code/litchi-worktrees/targets/0744/profiling/litchi-perf-baseline
)
PHASES=(
  store_parse=Worksheet::store
  rewrite_scan=snapshot::scan::scan
  rewrite_write=package::rewrite
  rewrite_write=rewrite_value_only_with_provenance
  compact=compact::changed_worksheet
  reduced_readback=reduced_readback
  verify_parse=raw::worksheet::parse
  audit=audit_published_xml
  deflate=Compress::compress
  preserve_copy=PreservationIndex
  write_other=write_to_stream
  commit_other=Edit::commit
)
for leg in before after; do
  bin=${BIN[$leg]}
  sha256sum "$bin" >> binaries.sha256
  taskset -c 12 perf record -m 128 -F 4000 -g -o "$leg-onecell.data" "$bin" --case xlsx_one_cell_commit_save --xlsx-shape dense-wide --samples 30 --warmup 2 --json "$leg-onecell.json" > /dev/null 2> "$leg-onecell.perf.stderr"
  perf script -i "$leg-onecell.data" -F ip,sym --no-inline 2>/dev/null > "$leg-onecell.script"
  python3 attrib.py "$leg-onecell.script" xlsx_commit_save_operation "${PHASES[@]}" > "$leg-onecell-phases.txt"
  taskset -c 12 perf record -m 128 -F 4000 -g -o "$leg-first.data" "$bin" --case xlsx_first_cell --xlsx-shape dense-wide --samples 100 --warmup 3 --json "$leg-first.json" > /dev/null 2> "$leg-first.perf.stderr"
  perf script -i "$leg-first.data" -F ip,sym --no-inline 2>/dev/null > "$leg-first.script"
  python3 attrib.py "$leg-first.script" Worksheet::store materialize=semantic::materialize from_unsorted=Store::from_unsorted lane_walk=lane::walk lane_recognize=lane::recognize reader=read_event_impl namespaces=process_event resolve=resolve_event > "$leg-first-phases.txt"
  rm -f "$leg-onecell.data" "$leg-first.data" "$leg-onecell.script" "$leg-first.script"
  taskset -c 12 valgrind --tool=callgrind --toggle-collect='*xlsx_commit_save_operation*' --callgrind-out-file="$leg-onecell.callgrind" "$bin" --case xlsx_one_cell_commit_save --xlsx-shape dense-wide --samples 2 --warmup 1 --json /dev/null > /dev/null 2> "$leg-onecell.callgrind.stderr"
  callgrind_annotate --inclusive=yes "$leg-onecell.callgrind" > "$leg-onecell-callgrind-inclusive.txt" 2>/dev/null
  taskset -c 12 valgrind --tool=callgrind --toggle-collect='*Worksheet*store*' --callgrind-out-file="$leg-first.callgrind" "$bin" --case xlsx_first_cell --xlsx-shape dense-wide --samples 2 --warmup 1 --json /dev/null > /dev/null 2> "$leg-first.callgrind.stderr"
  callgrind_annotate --inclusive=yes "$leg-first.callgrind" > "$leg-first-callgrind-inclusive.txt" 2>/dev/null
  rm -f "$leg-onecell.callgrind" "$leg-first.callgrind"
  echo "$(date +%T) $leg profiles done"
done

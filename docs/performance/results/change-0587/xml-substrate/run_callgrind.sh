#!/usr/bin/env bash
# Survey-only scratch runner: callgrind the probe for {eager,source} x {real,control}
# and extract inclusive Ir for the XML passes of interest. Deletes raw callgrind
# outputs after extraction (keeps only the annotated symbol tables).
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/xml-substrate
cd "$S" || exit 1
SYMS='read_event_impl|process_markup_compatibility|capture_inner|validate_xml|Parser>::parse|parse_source_with_observer|worksheet_xml_and_parse_source|process_markup_compatibility_stream|clone_bounded_name_part|Namespaces>::with_local|from_utf8|memmem|find_bytes|resolve_event|scan_range|x14ac::capture|selected::scan|PROGRAM TOTALS'
for m in eager source; do
  for v in real control; do
    out="cg-$m-$v.out"
    taskset -c 3 valgrind --tool=callgrind --callgrind-out-file="$out" --compress-strings=no --compress-pos=no \
      ./target/release/xmlprobe "$m" "$v.xlsx" H680 >/dev/null 2>"cg-$m-$v.stderr"
    callgrind_annotate --inclusive=yes "$out" 2>/dev/null | grep -E "$SYMS" | head -n 40 > "cg-$m-$v.inclusive.txt"
    callgrind_annotate --inclusive=no "$out" 2>/dev/null | head -n 45 > "cg-$m-$v.self.txt"
    rm -f "$out"
  done
done
echo done

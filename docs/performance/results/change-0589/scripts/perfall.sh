#!/usr/bin/env bash
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589
leg=$1; bin=$2
: > $SC/out/perfstat-$leg.jsonl
while read -r f; do
  bytes=$(stat -c%s "$f")
  if [ "$bytes" -gt 500000 ]; then lo=20; hi=120; else lo=100; hi=600; fi
  $SC/perfstat.sh "$leg" "$bin" doc-snapshot-open "$f" $lo $hi >> $SC/out/perfstat-$leg.jsonl
done < $SC/out/doc-fixtures.txt
while read -r f; do
  bytes=$(stat -c%s "$f")
  if [ "$bytes" -gt 500000 ]; then lo=20; hi=120; else lo=100; hi=600; fi
  $SC/perfstat.sh "$leg" "$bin" ppt-textedit-open "$f" $lo $hi >> $SC/out/perfstat-$leg.jsonl
done < $SC/out/ppt-fixtures.txt

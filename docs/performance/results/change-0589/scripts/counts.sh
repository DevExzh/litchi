#!/usr/bin/env bash
# $1 = leg name, $2 = binary
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589
leg=$1; bin=$2
: > $SC/out/counts-$leg.jsonl
while read -r f; do taskset -c 9 "$bin" counts doc-snapshot-open "$f" >> $SC/out/counts-$leg.jsonl; done < $SC/out/doc-fixtures.txt
while read -r f; do taskset -c 9 "$bin" counts doc-snapshot-resolve "$f" >> $SC/out/counts-$leg.jsonl; done < $SC/out/doc-fixtures.txt
while read -r f; do taskset -c 9 "$bin" counts ppt-textedit-open "$f" >> $SC/out/counts-$leg.jsonl; done < $SC/out/ppt-fixtures.txt

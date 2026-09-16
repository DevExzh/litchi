#!/bin/bash
set -u
S=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0655
until grep -q FINAL-COUNTS-DONE $S/out/final-counts.txt 2>/dev/null; do sleep 20; done
$S/timing-all.sh

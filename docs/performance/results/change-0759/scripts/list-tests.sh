#!/bin/bash
# usage: list-tests.sh <worktree> <target-dir> <out-file> <packages...>
WT=$1; TD=$2; OUT=$3; shift 3
ARGS=""; for p in "$@"; do ARGS="$ARGS -p $p"; done
cd "$WT" || exit 1
CARGO_TARGET_DIR=$TD TMPDIR=/home/zhuhe/code/litchi-worktrees/merge-scratch/tmp CARGO_BUILD_JOBS=16 cargo test --locked $ARGS -- --list 2>&1 | python3 -c '
import sys,re
cur="?"
for line in sys.stdin:
    line=line.rstrip("\n")
    m=re.match(r"\s+Running (unittests )?(\S+) \(.*/deps/([A-Za-z0-9_\-]+)-[0-9a-f]+\)", line)
    if m:
        cur=m.group(3)+":"+m.group(2); continue
    m=re.match(r"\s+Doc-tests (\S+)", line)
    if m:
        cur="doc:"+m.group(1); continue
    m=re.match(r"(.+): (test|bench)$", line)
    if m:
        name=m.group(1)
        if cur.startswith("doc:"):
            name=re.sub(r" \(line \d+\)"," ",name)
        print(cur+" :: "+name)
' | sort -u > "$OUT"
echo "$OUT $(wc -l < $OUT)"

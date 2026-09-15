#!/usr/bin/env bash
# args: leg bin
SC=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589
leg=$1; bin=$2
run() { # stem mode path n
  taskset -c 9 valgrind --tool=callgrind --callgrind-out-file=$SC/cg/$leg-$1-s$4.out \
    --cache-sim=no --branch-sim=no "$bin" profile "$2" "$3" 0 "$4" \
    > $SC/cg/$leg-$1-s$4.stdout 2> $SC/cg/$leg-$1-s$4.stderr
}
for n in 5 25; do run docpic doc-snapshot-open /home/zhuhe/code/litchi/test-data/ole/doc/picture.doc $n; done
for n in 10 60; do run docdup doc-snapshot-open /home/zhuhe/code/litchi/test-data/ole/doc/duplicate-style-names.doc $n; done
for n in 10 60; do run pptmid ppt-textedit-open /home/zhuhe/code/litchi/test-data/poi/test-data/slideshow/45543.ppt $n; done
echo "CGDONE-$leg"

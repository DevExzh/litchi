#!/usr/bin/env bash
cpu=24
for leg in "docsnap-dup:doc-snapshot-open:/home/zhuhe/code/litchi/test-data/ole/doc/duplicate-style-names.doc" "docsnap-pic:doc-snapshot-open:/home/zhuhe/code/litchi/test-data/ole/doc/picture.doc" "docfacade-pic:doc-facade-open:/home/zhuhe/code/litchi/test-data/ole/doc/picture.doc"; do
  IFS=: read -r stem mode path <<<"$leg"
  for n in 10 110; do
    ( setarch x86_64 -R taskset -c $cpu valgrind --tool=callgrind --callgrind-out-file=/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/doc-ppt/cg/$stem-s$n.out --cache-sim=no --branch-sim=no /tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/doc-ppt/target/release/docppt_survey profile $mode "$path" 1 $n > /tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/doc-ppt/cg/$stem-s$n.stdout 2> /tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/doc-ppt/cg/$stem-s$n.stderr; echo "rc=$? $stem-s$n" >> /tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/doc-ppt/cg/done2.log ) &
    cpu=$((cpu+1))
  done
done
wait; echo ALLDONE >> /tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/doc-ppt/cg/done2.log

# Direct (unsymlinked) runs

Supplementary single runs cited by the record, each pinned to core 20:

- `fe-{base,cand}-{40,1040}.perf`: `perf stat -x,` of the harness
  `cfb_open --shape few-large` with 40 and 1,040 samples, events
  `instructions:u,cycles:u,branch-misses:u,ic_cache_fill_l2,ic_cache_fill_sys,bp_redirects.resync`,
  started through the round-0 symlinks (`argv0/<arm>/litchi-perf-baseline-p`).
  The per-owner values in the record are (1,040-sample − 40-sample) / 1,000.
- `{base,cand}.rec`: the `perf record -e page-faults -c 1 --call-graph dwarf`
  summaries of `doc_semantic_one_edit_save --writer-shape large --samples 220
  --warmup 5`, started as `bin/<arm>/litchi-perf-baseline`. With `-c 1` every
  page fault of the whole process is an event: perf processed 132,364 (base)
  and 52,494 (cand), and kept 71,640 and 32,839 after losing 45.9% and 37.4%
  while writing call chains. The `perf.data` files were deleted after
  summarizing.
- `strace-{base,cand}-{20,220}.txt`: `strace -f -c -e trace=mmap,munmap,brk,madvise,mremap`
  of the same selector with 20 and 220 samples, from the same unsymlinked path.

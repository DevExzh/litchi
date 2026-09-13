# change-0563 evidence packet

Change record: [`docs/performance/0563-opc-single-warm-part-observation.md`](../../0563-opc-single-warm-part-observation.md).
Disposition: retained. `performance_claim: none`; no claim-registry entry is added.

## Contents

| Path | What it is |
| --- | --- |
| `syscalls/` | Whole-child `strace -f -c` counts for four OOXML file-source cases, baseline and candidate. |
| `latency/` | The 16 A/H/H/A children of the non-regression matrix. |

## The primary evidence is not in this directory

It is three counted-observation tests in `crates/litchi-opc/src/source_backed.rs`
— `a_warm_part_read_observes_the_source_twice`,
`a_cold_part_read_keeps_its_complete_observation_bracket` and
`a_changed_source_outranks_a_missing_part`. The first two were confirmed to fail
with the removed observation restored. They assert the warm path takes two
observations and zero source reads and the cold path keeps its five, which is the
result this change is about; the syscall captures below only show how little of
that reaches a cold-dominated selector.

## Replay

```sh
cargo test --offline -p litchi-opc --all-features -- \
  a_warm_part_read_observes_the_source_twice \
  a_cold_part_read_keeps_its_complete_observation_bracket \
  a_changed_source_outranks_a_missing_part
```

## What is not here

No cold-cache, physical-device, remote/range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform capture, and no
file-backed repeated-part-read selector — which is the coverage gap that leaves
the warm path's end-to-end effect unmeasured.

# Evidence: change 0590, PPTX opened-transaction revision reuse

Change record:
[`0590-pptx-opened-transaction-revision-reuse.md`](../../0590-pptx-opened-transaction-revision-reuse.md).

Disposition: retained. `performance_claim: none`, no claim-registry entry. This
packet holds both legs of every measurement the record cites and nothing else.

## Contents

| Path | What it is |
| --- | --- |
| `counts/calls-{before,after}-<case>-s1.txt` | Callgrind call counts and whole-child `Ir` for `package_fingerprint`, the capture entry point, `notes::load_snapshot` and (after only) `packages_equal` / `Snapshot::rebound_to`, per selector, at `--warmup 0 --samples 1`. |
| `counts/calls-{before,after}-<case>-s3.txt` | The same counts at `--samples 3`, for `pptx_eager_batch_edit_save`, `pptx_slide_remove_boundary_save` and `pptx_slide_move_boundary_save`. The s3 − s1 difference over the two extra samples is the per-lifecycle isolation pair the record quotes. |
| `counts/incl-{before,after}-<case>-s{1,3}.txt` | `callgrind_annotate --inclusive=yes --threshold=99.5` for each leg, selector and sample count. |
| `counts/cg-{before,after}-*.log` | The valgrind run logs, each ending in `exit=0` and the collected `Ir`. |
| `timing/time-<case>-{A1,B1,B2,A2}.json` | The harness reports for the paired timing, in run order: before, after, after, before. 30 measured samples after 3 warmups per leg, pinned to CPU 10. |
| `timing/time-<case>.txt` | The paired summary the record quotes: per-leg p50/mean/p95/p99, the pooled before and after legs, the deltas in both directions, and the A/A and B/B floors from the repeated legs. |
| `timing/time100-pptx_slide_remove_boundary_save-*` | A second, independent four-leg run of the one selector that regressed, at 100 measured samples after 5 warmups. |
| `timing/phases-<case>.txt` | The two boundary selectors' own per-phase clocks (plan, commit, publication, reopen), paired across the four legs. This is what localizes the move selector's change to its commit phase and shows the remove selector drifting uniformly across phases that contain no changed code. |
| `timing/perf-<case>-{A1,B1,B2,A2}.txt` | `perf stat -e cycles,instructions,task-clock` around the same four legs. |
| `timing/perf-<case>.txt` | The cycles and instructions summary per leg. |
| `scripts/` | `cg-run.sh` (one callgrind leg plus its annotation and counts), `callcounts.py` (sums callgrind `calls=` per callee over the shared `fn`/`cfn` name-compression table), `time-run.sh`, `time-run-100.sh` and `perf-run.sh` (the A1 B1 B2 A2 legs), `time-report.py`, `time-report-100.py`, `perf-report.py` and `phase-report.py` (the summaries). |
| `gates.txt` | The tail of every gate run in the worktree. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; the coordinator merges them. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad after this packet was assembled. |

## Provenance

| | |
| --- | --- |
| Base commit | `08d968f8ec7db27cf1187d01911fd08b9d014d91` (`feat/office-format-completeness`) |
| Branch | `perf/0590-pptx-opened-transaction-revision-reuse` |
| Before leg | built `--release --locked` from the read-only checkout of the base at `/home/zhuhe/code/litchi-worktrees/before-08d968f8e/tools/perf-baseline`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0590-before` |
| Before binary sha256 | `6cb5c84f88b8ee27fcc85099065bb79ec034e3b7abbfce601b4bb7c072d62c95` |
| After leg | built `--release --locked` from this branch's `tools/perf-baseline` |
| After binary sha256 | `b1c0e1be7b2efde61340dc9fc3cf42e749abd2aba5b9baf06f003635aba08f3f` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 |
| Pinning | every measured process `taskset -c 10` |
| Host load | seven other agents of the same wave were building and measuring on the other cores throughout; the A/A and B/B floors in `timing/*.txt` are the only statement about that |

## What this packet does not establish

No speedup, regression, allocation, RSS, cold-cache, range-source, concurrency,
real-producer or cross-platform result. The call counts and instruction
differentials are exact for these four selectors, on their fixed synthetic
corpora, on this host and these two builds. Callgrind runs SHA-256 in software
because valgrind masks the SHA CPUID bit, so any instruction share attributed to
hashing is about five times its native cycle share. The eager PPTX corpora are
built without physical source provenance and re-deflate all media on save, so
their wall-clock region is inflated by recompression a production open would not
perform. One selector, `pptx_slide_remove_boundary_save`, is slower in this
window; its numbers are here in full and the record says why they are not
attributable to the change.

The raw `callgrind.out` files were deleted after the annotations and counts were
extracted; `cleanup.json` records them.

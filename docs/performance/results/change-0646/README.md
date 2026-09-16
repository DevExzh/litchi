# Evidence: change 0646, PPTX cross-package copy candidate retention (design)

Change record:
[`0646-pptx-cross-copy-candidate-retention-design.md`](../../0646-pptx-cross-copy-candidate-retention-design.md).

Disposition: **design, retained**. `performance_claim: none`, no claim-registry
entry. **No production code is changed on the branch.** The implementation the
measurements below describe is a scratch patch kept in `patch/`; it exists to
size the design and is not merged.

## Contents

| Path | What it is |
| --- | --- |
| `patch/0646-retained-candidate-scratch.patch` | The complete scratch implementation, as `git diff` against the base: one doc-hidden `Arc` accessor in `litchi-opc`, the `RetainedCandidate` field and its threading in `litchi-pptx`, and the three binding-proof tests. Applies to `c7326f680` with `git apply`. This is the code every number below was measured on. |
| `counts/incl-{before,after}-<case>-s{1,3}.txt` | `callgrind_annotate --threshold=99.9` for each leg, selector and sample count. |
| `counts/calls-{before,after}-<case>-s{1,3}.txt` | Callgrind call counts per callee for the twelve symbols the record names, summed over the shared `fn`/`cfn` name-compression table. The `(s3 − s1) / 2` difference is the per-lifecycle isolation pair the record quotes. |
| `counts/cg-*.runlog.txt` | The valgrind run log for each profile, naming the exact binary and arguments and ending in the collected `Ir`. (`.txt`, because the repository gitignores `*.log`.) |
| `counts/counts-summary.txt` | The generated before/after tables the record's count and instruction sections quote: calls per lifecycle from the `(s3 − s1) / 2` isolation pair, inclusive `Ir` per lifecycle, and the whole child at `--samples 1`. |
| `alloc/retention-{before,after}-{plain,media-rich}.json` | `litchi-perf-baseline-alloc retention --api owned`, 5 samples after 1 warmup per leg, with the nine ownership checkpoints, the region peak and the absolute allocator counters. |
| `alloc/alloc-summary.txt` | The generated checkpoint, region-peak and allocator-counter comparison the record's memory section quotes. |
| `timing/timing-{A1,B1,B2,A2}.json` | Harness reports for the paired timing in run order (before, after, after, before), all four cross-copy selectors in one process, pinned to CPU 21. Each carries its own `binary_identity.binary_sha256`, per-phase clocks and the published `output_sha256` of every sample. |
| `timing/timing-summary.txt` | Per-leg p50/mean/p95/p99, pooled before and after, deltas in both directions, the A/A and B/B floors from the repeated legs, the distinct published digests, and the per-phase (plan, apply, publication, reopen) medians. |
| `timing/perf-{A1,B1,B2,A2}.txt` | `perf stat -e cycles,instructions -x,` around the same four legs, whole child. |
| `timing/perf-summary.txt` | Cycles and instructions per leg with the A/A and B/B floors. |
| `timing/phase-per-leg.txt` | The per-leg medians of every phase clock (plan, apply, publication, reopen, lifecycle) for all four selectors, in run order. This is what the record's bimodal-publication finding is read off. |
| `tests/` | `litchi-pptx-before.txt` and `litchi-pptx-scratch-debug.txt`: the `litchi-pptx` suite output for the untouched before checkout and for the scratch tree, which is the "verdicts unchanged" evidence, plus `litchi-opc-scratch.txt` for the crate the patch's one accessor lives in. |
| `scripts/` | `cg.sh` (one callgrind leg plus its annotation), `callcounts.py` (sums callgrind `calls=` per callee over the shared name-compression table), `report.py` (the isolation-pair count and instruction tables), `timing.sh` and `perfstat.sh` (the A1 B1 B2 A2 legs), `timing_report.py`, `perf_report.py` and `alloc_report.py` (the summaries). `callcounts.py`, `timing_report.py` and `perf_report.py` are reused verbatim from `results/change-0598/scripts/`. |
| `gates.txt` | The tail of every gate run in the worktree. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; the coordinator merges them. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad after this packet was assembled. |

## Provenance

| | |
| --- | --- |
| Base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` (`feat/office-format-completeness`) |
| Branch | `perf/0646-pptx-cross-copy-candidate-retention-design` |
| Before leg | built `--release --locked` from the shared read-only checkout at `/home/zhuhe/code/litchi-worktrees/before-c7326f680/tools/perf-baseline`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0646-before` |
| After leg | the same command in the branch worktree with the scratch patch applied, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0646-after` |
| Binaries | staged outside both target directories before any measurement (change 0627: a concurrent build relinked one mid-run) |
| `litchi-perf-baseline` before | `900816be96595f095fec6b9a0c765360bb8331f161bb82f78f3d6d5c80abf89f` |
| `litchi-perf-baseline` after | `ac6a1a3b086ad8b0192dbdfac8da6229a2d00c3b815bdb258000e53b54d424a5` |
| `litchi-perf-baseline-alloc` before | `e5df08e5116a2718888e1922cb5a762c15bb977ad2380047a649e33bc1f29a85` |
| `litchi-perf-baseline-alloc` after | `2d13631ce0224d4b62436aafc47f342605249f95a7987dc8751c861b3d90d2b1` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 |
| Pinning | every measured process `taskset -c 21` |
| Host load | seven other agents of the same wave were building and measuring on the other cores; the run-window load average is recorded beside the timing summary, and the A/A and B/B floors are the only statement about it |

## What this packet does not establish

No speedup is claimed and none is registered. **No production code changed**, so
nothing here is a statement about the shipped library: it is a statement about
what a specific scratch implementation of a specific design costs and saves on
two synthetic corpora, on this host, in these two builds.

The call counts, instruction differentials and allocator counters are exact and
deterministic. The wall-clock and `perf stat` numbers are not, and the floors are
reported beside them. Callgrind runs SHA-256 in software because valgrind masks
the SHA CPUID bit, so any instruction share attributed to hashing is about five
times its native cycle share; `perf stat` is the counterweight and it is
whole-child, so it dilutes rather than isolates.

Both corpora are generated by `litchi-pptx-cross-slide-copy-evidence-v1` and the
media-rich one is deliberately incompressible. The largest real `.pptx` under
`test-data/` is 972,788 bytes, so no real deck here exercises the large end of the
retention ceiling. No source-backed cross-copy selector, no external-package
fixture, no cold-cache, range-source, concurrency or cross-platform measurement
was taken.

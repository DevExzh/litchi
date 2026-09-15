# Evidence: change 0598, PPTX cross-package copy revision reuse

Change record:
[`0598-pptx-cross-copy-revision-cache.md`](../../0598-pptx-cross-copy-revision-cache.md).

Disposition: retained. `performance_claim: none`, no claim-registry entry. This
packet holds both legs of every measurement the record cites and nothing else.

## Contents

| Path | What it is |
| --- | --- |
| `counts/calls-{before,after}-<case>-s{1,3}.txt` | Callgrind call counts per callee for `package_fingerprint`, `physical_package_fingerprint`, `snapshot_physical_revision`, `bounded_package_bytes`, `capture_internal`, `prepare_cross_slide_copy_for_slides` and `PackageWriter::write_to_stream`, with the whole-child `Ir`, per leg, selector and sample count. The s3 − s1 difference over the two extra samples is the per-lifecycle isolation pair the record quotes. |
| `counts/incl-{before,after}-<case>-s{1,3}.txt` | `callgrind_annotate --inclusive=yes --threshold=99.5` for each leg, selector and sample count. |
| `counts/cg-{before,after}-*.runlog.txt` | The valgrind run logs for each profile, each naming the exact binary and arguments and ending in the collected `Ir`. (`.txt`, because the repository gitignores `*.log`.) |
| `counts/counts-summary.txt` | The generated before/after tables the record's count and instruction sections quote, including the per-lifecycle differentials. |
| `timing/timing-{A1,B1,B2,A2}.json` | The harness reports for the paired timing, in run order: before, after, after, before. 30 measured samples after 3 warmups per leg, all four cross-copy selectors in one process, pinned to CPU 21. Each carries its own `binary_identity.binary_sha256`, the per-phase clocks and the published `output_sha256` of every sample. |
| `timing/timing-summary.txt` | The paired summary the record quotes: per-leg p50/mean/p95/p99, the pooled before and after legs, the deltas in both directions, the A/A and B/B floors from the repeated legs, the distinct published digests, and the per-phase (plan, apply, publication, reopen) medians. |
| `timing/perf-{A1,B1,B2,A2}.txt` | `perf stat -e cycles,instructions -x,` around the same four legs at `--warmup 1 --samples 10`, whole child. |
| `timing/perf-summary.txt` | The cycles and instructions summary per leg with the A/A and B/B floors. |
| `scripts/` | `cg.sh` (one callgrind leg plus its annotation), `callcounts.py` (sums callgrind `calls=` per callee over the shared `fn`/`cfn` name-compression table), `counts_report.py` (the before/after count and instruction tables), `timing.sh` and `perfstat.sh` (the A1 B1 B2 A2 legs), `timing_report.py` and `perf_report.py` (the summaries). |
| `gates.txt` | The tail of every gate run in the worktree. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; the coordinator merges them. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad after this packet was assembled. |

## Provenance

| | |
| --- | --- |
| Base commit | `f22f9393598da39272f2185d2dac3d36c4b61d23` (`feat/office-format-completeness`) |
| Branch | `perf/0598-pptx-cross-copy-revision-cache` |
| Before leg | built `--release --locked` from the read-only checkout of the base at `/home/zhuhe/code/litchi-worktrees/before-f22f93935/tools/perf-baseline`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0598-before` |
| Before binary sha256 | `9946dc388dddb4937e5c4ebe9e0fc4a50f75dcd698f9968f08593ad8ae5e7b4f` |
| After leg | built `--release --locked` from this branch's `tools/perf-baseline`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0598-after` |
| After binary sha256 | `a224eaecd7c32f4e43d8dd2ed69a8b2a4d774749298ea0b3c0e7ec323b3efbdb` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 |
| Pinning | every measured process `taskset -c 21` |
| Host load | other agents of the same wave were building and measuring on the other cores throughout; the A/A and B/B floors in `timing/timing-summary.txt` are the only statement about that |

## What this packet does not establish

No speedup, regression, allocation, RSS, cold-cache, range-source, concurrency,
real-producer or cross-platform result. The call counts and instruction
differentials are exact for these two selectors, on their fixed synthetic
corpora, on this host and these two builds; the harness's own corpus
construction and its ten untimed refusal gates dominate the whole-child totals.
Callgrind runs SHA-256 in software because valgrind masks the SHA CPUID bit, so
any instruction share attributed to hashing is about five times its native cycle
share; the `perf stat` rows are the counterweight and they are whole-child.

The wall-clock floors in this window are large: the media-rich A/A floor is
−4.99% at p50 and the plain lifecycle's B/B floor is −4.68%, so two of the four
selectors' wall-clock deltas are inside their own noise and are reported rather
than relied on. The record rests on the counts.

The external-package copy binary was not run; its pinned LibreOffice QA fixture
is not in the tree (change 0454 fetched it over the network and the artifact
remains outside tracked source).

The raw `callgrind.out` files were deleted after the annotations and counts were
extracted; `cleanup.json` records them.

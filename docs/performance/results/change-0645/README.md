# Evidence: change 0645, the memoized per-part revision proof (frozen design)

Change record:
[`0645-pptx-memoized-revision-proof-design.md`](../../0645-pptx-memoized-revision-proof-design.md).

Disposition: **retained, design only.** `performance_claim: none`, no
claim-registry entry. No file under `crates/` or `tools/` is modified on the
branch. The scratch implementation that produced every number here lives only
in the measurement worktree and is retained under `patches/`.

## Contents

| Path | What it is |
| --- | --- |
| `counts/calls-{before,after,afterb}-<case>-s{1,3}.txt` | Callgrind call counts and whole-child `Ir` for the fingerprint entry points, the capture entry point, `notes::load_snapshot`, `packages_equal` and `sha2::sha256::compress256`, per selector and sample count. The s3 − s1 difference over the two extra samples is the per-lifecycle isolation pair. |
| `counts/incl-{before,after,afterb}-<case>-s{1,3}.txt` | `callgrind_annotate --inclusive=yes --threshold=99.5` for each leg, selector and sample count. |
| `counts/callsites-{before,after,afterb}-pptx_eager_batch_edit_save-s1.txt` | The fingerprint's inclusive `Ir` split by **call site** (`capture_internal` versus `Transaction::commit`), which is what separates a cold fingerprint from a memoized one. |
| `counts/isolation-pairs.txt` | The per-lifecycle table the record quotes: total `Ir`, fingerprint subtree, `sha2` subtree and call counts, for all three legs and all four selectors. |
| `counts/memo-per-fingerprint.txt` | The probe build's per-fingerprint memo accounting for one eager lifecycle: parts hit, parts hashed, payload bytes hashed and total bytes fed to SHA-256. |
| `counts/memo-hitrate.txt` | The same probe's per-commit summary (228 of 229 parts hit). |
| `counts/cg-*.log` | The valgrind run logs, each ending in `exit=0`. |
| `timing/time-<case>-{A1,B1,B2,A2}.json` | Harness reports for the paired timing, in run order before, A+B, A+B, before. 30 measured samples after 3 warmups per leg, pinned to CPU 20. |
| `timing/time-<case>.txt` | The paired summary: per-leg p50/mean/p95/p99, pooled legs, deltas in both directions, and the A/A and B/B floors from the repeated legs. |
| `timing/phases-<case>.txt` | The selectors' own per-phase clocks (plan, commit, publication, reopen), paired across the four legs. |
| `timing/perf-<case>-{A1,B1,B2,A2}.txt` | `perf stat -e cycles,instructions,task-clock` around the same four legs. |
| `timing/perf-<case>.txt` | The cycles and instructions summary per leg with the A/A floor. |
| `patches/scratch-design-a.patch` | **Scratch leg A** — the memo on `Snapshot` alone. Applies to `c7326f680`. Not `cargo fmt`-clean (it predates the formatting pass); leg A+B is the formatted one. |
| `patches/scratch-design-ab.patch` | **Scratch leg A+B** — A plus the facade `Package`'s retained memo. Applies to `c7326f680`, `cargo fmt --all --check` clean, and passes the gates in `gates.txt`. This is the implementation every A+B number came from. |
| `patches/scratch-probe-instrumentation.patch` | The per-fingerprint hit/miss counters and the harness print behind `LITCHI_0645_MEMO`, applied on top of A+B for the probe build only. **Never present in a timed binary**; the timed binaries' sha256s are in `binaries.txt` and were taken from uninstrumented builds. |
| `scripts/` | `cg-run.sh`, `counts-all.sh`, `counts-b.sh` (the callgrind legs), `callcounts.py` (call counts per callee over the shared name-compression table), `callsites.py` (inclusive `Ir` per caller/callee pair), `summarize.py` (the isolation-pair table), `time-run.sh`, `perf-run.sh`, `timing-all.sh` (the A1 B1 B2 A2 legs), `time-report.py`, `perf-report.py`, `phase-report.py`. |
| `binaries.txt` | sha256 of the three staged binaries. |
| `gates.txt` | The scratch implementation's gates, then this branch's documentation gates. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; the coordinator merges them. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad after this packet was assembled. |

## Provenance

| | |
| --- | --- |
| Base commit | `c7326f680` (`feat/office-format-completeness`) |
| Branch | `perf/0645-pptx-revision-proof-format-design` |
| Before leg | built `--release --locked` from the shared read-only checkout at `/home/zhuhe/code/litchi-worktrees/before-c7326f680/tools/perf-baseline`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0645-before` |
| Before binary sha256 | `a494c98a7cfd7014dafd201cc079ef4c5f6e7862b1d85d6cbea1f8f9173e618c` |
| After leg **A** sha256 | `8f048eb0883516471d69cc1ff3087987214b3c84456608ea1a3cd4fdd51eb46c` |
| After leg **A+B** sha256 | `f82f6b900f4d4078135a4bf5481916f69d90ef354c5564ca80f16151ebe018f2` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 |
| Pinning | every measured process `taskset -c 20`; binaries staged outside every Cargo target directory before measurement (0627's rule) |
| Host load | seven other agents of the same wave were building and measuring on the other cores throughout; the A/A and B/B floors in `timing/*.txt` are the only statement about that |

## What this packet does not establish

No speedup, no registered claim, and no production behaviour at all: the branch
changes only documentation. The scratch implementation does **not** bump either
durable magic and does **not** carry the typed refusal, so this packet says
nothing about how an `LPRM0001` or `LPCP0002` patch behaves under the new
proof — that is design, and admission gate 1 is where it gets measured.

The call counts and instruction differentials are exact for four selectors on
their fixed synthetic corpora, on this host and these builds. Callgrind runs
SHA-256 in software because valgrind masks the SHA CPUID bit, so every
instruction share attributed to hashing is roughly five times its native cycle
share; the native `perf stat` cycles are the counterweight and they are a
quarter to a fifth of the callgrind figure. The eager corpora carry no physical
source provenance and re-deflate all media on save, so their wall-clock
denominator is inflated by recompression a production `open` would not perform
(change 0590's caveat, unchanged).

Two selectors ran in a noisy window: `pptx_slide_move_boundary_save`'s own B/B
p50 floor is −10.17% and its A/A p95 floor is −22.01%, and
`pptx_slide_remove_boundary_save`'s p50 delta sits on its own A/A floor. Their
wall clock is reported in full and is not the result; their counts are.

One selector is **slower** because of this design:
`pptx_slide_remove_boundary_save` is +2.45% in callgrind instructions per
lifecycle, +0.32% in native instructions and +0.43% in cycles. The mechanism is
measured and named in the record.

The raw `callgrind.out` files were deleted after the annotations, call counts
and per-call-site splits were extracted; `cleanup.json` records them.

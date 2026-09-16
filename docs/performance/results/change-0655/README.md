# Evidence: change 0655, the memoized per-part revision proof, landed

Change record:
[`0655-pptx-memoized-revision-proof.md`](../../0655-pptx-memoized-revision-proof.md).

Disposition: **retained, implemented in `crates/litchi-pptx`.**
`performance_claim: none`, no claim-registry entry.

## Contents

| Path | What it is |
| --- | --- |
| `counts/calls-{before,after}-<case>-s{1,3}.txt` | Callgrind call counts and whole-child `Ir` for the fingerprint entry points, the capture entry point, `notes::load_snapshot`, `packages_equal`, `PartDigests::project` and `sha2::sha256::compress256`, per leg, selector and sample count. The s3 − s1 difference over the two extra samples is the per-lifecycle isolation pair. |
| `counts/incl-{before,after}-<case>-s{1,3}.txt` | `callgrind_annotate --inclusive=yes --threshold=99.5` for each leg, selector and sample count. |
| `counts/callsite-pairs.txt` | Per-lifecycle inclusive `Ir` and call count for each *call site* of the fingerprint, as an isolation pair over the two sample counts. This is what separates a cold fingerprint from a memoized one on the same leg. |
| `counts/isolation-pairs.txt` | The per-lifecycle table the record quotes: total `Ir`, fingerprint subtree, `sha2` subtree and call counts, for both legs and every selector. |
| `counts/memo-per-fingerprint-<case>.txt` | The probe build's per-fingerprint memo accounting: parts hit, parts hashed, payload bytes hashed and total bytes fed to SHA-256, one line per fingerprint, at `--samples 1` and `--samples 3` so the fingerprints-per-lifecycle count is an isolation pair. |
| `counts/memo-per-lifecycle.txt` | The same probe output differenced into an exact per-lifecycle table: how many fingerprints each lifecycle issues and, for each one, how many parts the memo answered and how many bytes were hashed and fed. |
| `counts/cg-*.txt` | The valgrind run logs, each ending in `exit=0`. They are named `.txt` rather than `.log` because the repository ignores `*.log` everywhere. |
| `timing/time-<case>-{A1,B1,B2,A2}.json` | Harness reports for the paired timing, in run order before, after, after, before. 30 measured samples after 3 warmups per leg, pinned to CPU 10. |
| `timing/time-<case>.txt` | The paired summary: per-leg p50/mean/p95/p99, pooled legs, deltas in both directions, and the A/A and B/B floors from the repeated legs. |
| `timing/phases-<case>.txt` | The selectors' own per-phase clocks, paired across the four legs. |
| `timing/perf-<case>-{A1,B1,B2,A2}.txt` | `perf stat -e cycles,instructions,task-clock` around the same four legs. |
| `timing/perf-<case>.txt` | The cycles and instructions summary per leg with the A/A floor. |
| `instrumentation/memo-accounting.patch` | The per-fingerprint hit/miss and bytes-fed counters behind `LITCHI_0655_MEMO`, applied on top of the landed implementation for the probe build only and **reverted before any timed binary was built**. The timed binaries' sha256s are in `binaries.txt`. |
| `scripts/` | `cg-run.sh`, `cg-run-real.sh`, `counts-all.sh`, `counts-after.sh`, `counts-real-before.sh`, `final-counts.sh` (the callgrind legs), `memo-per-lifecycle.py` (differences the probe's per-fingerprint lines into an exact per-lifecycle table), `callsite-pairs.py` (per-call-site isolation pairs), `make-gates.sh` (assembles `gates.txt`), `callcounts.py` (call counts per callee over the shared name-compression table), `callsites.py` (inclusive `Ir` per caller/callee pair), `summarize.py` (the isolation-pair table), `time-run.sh`, `perf-run.sh`, `timing-all.sh`, `run-timing.sh` (the A1 B1 B2 A2 legs), `time-report.py`, `perf-report.py`, `phase-report.py`. Adapted from change 0645's packet; the differences are CPU 10 rather than 20, two legs rather than three, and the `--ooxml-file` argument the real-deck selectors need. |
| `binaries.txt` | sha256 of the three staged binaries (before, after, probe). |
| `gates.txt` | The tail of every gate. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; the coordinator merges them. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree tree after this packet was assembled. |

## Provenance

| | |
| --- | --- |
| Base commit | `70d7768cc` (`feat/office-format-completeness`) |
| Branch | `perf/0655-pptx-memoized-revision-proof` |
| Before leg | built `--release --locked` from the shared read-only checkout at `/home/zhuhe/code/litchi-worktrees/before-70d7768cc/tools/perf-baseline`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0655-before` |
| After leg | built `--release --locked` from this branch's worktree, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0655-after` |
| Probe leg | the after leg plus `instrumentation/memo-accounting.patch`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0655-probe`; used only for the hit/miss accounting, never timed |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0, valgrind 3.26.0 |
| Pinning | every measured process `taskset -c 10`; binaries staged outside every Cargo target directory before measurement (change 0627's rule) |
| Host load | seven other agents of the same wave were building and measuring on the other cores throughout; the A/A and B/B floors in `timing/*.txt` are the only statement about that |

## What this packet does not establish

No registered claim and no speedup. `lto = true` makes release binaries
non-reproducible byte for byte (change 0635), so the sha256s identify the
binaries that were measured rather than promising a rebuild reproduces them.

Callgrind runs SHA-256 in software because valgrind masks the SHA CPUID bit, so
every instruction share attributed to hashing is several times its native cycle
share; the `perf stat` cycles are the counterweight.

The raw `callgrind.out` files were deleted after the annotations and call counts
were extracted; `cleanup.json` records them.

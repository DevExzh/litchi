# Evidence: change 0656, the PPTX cross-package copy's retained candidate archive and its budget

Change record:
[`0656-pptx-cross-copy-candidate-budget.md`](../../0656-pptx-cross-copy-candidate-budget.md).

Disposition: **retained, implemented**. `performance_claim: none`, no
claim-registry entry. The paired medians and counts below are evidence, not
registered claims.

## Contents

| Path | What it is |
| --- | --- |
| `counts/incl-{before,after}-<case>-s{1,3}.txt` | `callgrind_annotate --threshold=99.9` for each leg, selector and sample count. |
| `counts/calls-{before,after}-<case>-s{1,3}.txt` | Call counts per callee for the twelve symbols the record names, summed over the shared `fn`/`cfn` name-compression table. The `(s3 − s1) / 2` difference is the per-lifecycle isolation pair the record quotes. |
| `counts/cg-*.runlog.txt` | The valgrind run log for each profile, naming the exact binary and arguments and ending in the collected `Ir`. (`.txt`, because the repository gitignores `*.log`.) The eight raw `callgrind.out` profiles themselves are **not** retained (4.4 MB); `scripts/run-counts.sh` regenerates them and `scripts/report.py` rebuilds every table from them. |
| `counts/counts-summary.txt` | The generated before/after tables the record's count and instruction sections quote, for both selectors: calls per lifecycle from the `(s3 − s1) / 2` isolation pair, inclusive `Ir` per lifecycle, and the whole child at `--samples 1`. |
| `counts-final/` | The same annotations, per-callee call counts and run logs for the **after leg re-run on the committed tree's binary**, which is the check that the three behaviour-free edits made after the first windows (two doc comments, a `derive(Debug)` removed from a private struct nothing formats, test code) did not move the production code. The call counts are identical and the per-lifecycle totals agree within 0.11% (plain) and 0.0001% (media-rich). |
| `alloc/retention-{before,after}-{plain,media-rich}.json` | `litchi-perf-baseline-alloc retention --api owned`, 5 samples after 1 warmup per leg, with the nine ownership checkpoints, the region peak and the absolute allocator counters. |
| `alloc/alloc-summary.txt` | The generated checkpoint, region-peak and allocator-counter comparison for the first after build. |
| `alloc-final/` | The same probe on the committed tree's binary, which is the table the record quotes. The two differ only in the binaries' constant static-allocation difference (+20 bytes instead of −4); the `planned` deltas net of it are identical. |
| `timing/w4/` | **The quoted window**, and the only one whose after binary is the committed tree's. Same shape as `w2` below; load average 19.5 falling to 18.2. Its A/A p50 floors are −1.35%, −2.36%, −3.12% and +4.71%, all under 5%; its B2 leg took excursions that push two selectors' p95/p99 positive, which the record reports rather than averages away. |
| `timing/w2/` | **Quoted beside w4** for its floors, which are the tightest of the four (A/A p50 under 2% on every selector) and whose after binary predates the three behaviour-free edits `counts-final/` checks. Harness reports for the paired timing in run order A1 B1 B2 A2 (before, after, after, before), all four cross-copy selectors in one process, pinned to CPU 11, plus the `perf stat` legs over the same order. Each JSON carries its own `binary_identity.binary_sha256`, per-phase clocks and the published `output_sha256` of every sample. `timing-summary.txt` has per-leg p50/mean/p95/p99, the pooled deltas in both directions, the A/A and B/B floors and the distinct published digests; `phase-per-leg.txt` has the per-leg medians of every phase clock; `perf-summary.txt` has cycles and instructions per leg with their floors; `load.txt` records the run-window load average. |
| `timing/w1/` | The first timing window, run under load average 26 rising to 59. Its A/A p50 floors are +145%, +3.4%, −57% and −6.0%, so its selector totals say nothing and the record says so; it is retained because the programme's rule is to report every scenario run, including the ones that got worse, and because its `perf stat` legs (cycles −27.08% against a +0.82% A/A floor) agree with window 2's independently. |
| `timing/w3/` | A media-rich-only window at 40 samples per leg, run to settle two phases: whether the planning phase regressed (it does not: the before and after pairs' ranges overlap, and window 4 reads the same phase negative) and whether `publication_ns` follows the binary (it does not: one leg of each binary lands in each of its two modes inside this single window, and window 4 splits an after pair across the two modes). |
| `tests/litchi-pptx-before.txt` | `cargo test -p litchi-pptx` on a detached worktree of the untouched base `70d7768cc`: 875 passed, 0 failed, 2 ignored across 78 suites. |
| `tests/litchi-pptx-after.txt` | The same on the branch: 885 passed, 0 failed, 2 ignored across 78 suites. The difference is the ten tests this change adds, and every pre-existing test runs on the reuse path with its `debug_assert` re-serializing the candidate. |
| `tests/litchi-pptx-release-retention.txt` | `cargo test -p litchi-pptx --release --lib retention`: the release half of the substituted-archive proof, where the `debug_assert` is compiled out and the refusal has to come from the recomputed revision. |
| `tests/litchi-opc-after.txt` | `cargo test -p litchi-opc`, the crate the new shared-handle accessor lives in: 712 passed, 0 failed, 1 ignored. |
| `tests/litchi-facade-after.txt` | `cargo test -p litchi --features docx,xlsx,pptx,xls`: 265 passed, 0 failed, 7 ignored. |
| `tests/perf-baseline-after.txt` | `cargo test` in `tools/perf-baseline`, which the constructor change touches: 540 passed, 0 failed, 1 ignored. |
| `scripts/` | `cg.sh` and `run-counts.sh` (one callgrind leg plus its annotation, and the serial sequence over both selectors, both legs and both sample counts), `callcounts.py` and `dump_calls.py` (per-callee call counts), `report.py` (the isolation-pair count and instruction tables), `timing.sh`, `timing-mr.sh`, `perfstat.sh` and `window2.sh` (the A1 B1 B2 A2 legs and the quiet-window watcher), `timing_report.py`, `phase_per_leg.py`, `perf_report.py` and `alloc_report.py` (the summaries). `callcounts.py`, `report.py`, `timing_report.py`, `perf_report.py` and `alloc_report.py` derive from `results/change-0646/scripts/`, with the CPU pin and the change number changed. |
| `binary-sha256.txt` | The sha256 of every binary that was measured, as staged outside the Cargo target directories. Six: the before harness and allocator binaries, and two builds each of the after ones. |
| `gates.txt` | The tail of every gate run in the worktree. |
| `log-sections.md` | The four log paragraphs for `HOTSPOTS.md`, `GOAL_AUDIT.md`, `REPORT.md` and `ADR_COMPLIANCE.md`; the coordinator merges them. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `cleanup.json` | What was removed from the session scratchpad and the worktree tree after this packet was assembled. |

## Provenance

| | |
| --- | --- |
| Base commit | `70d7768cc6dada420ede063f72c88dc99ad30383` (`feat/office-format-completeness`) |
| Branch | `perf/0656-pptx-cross-copy-candidate-budget` |
| Before leg | built `--release --locked` from the shared read-only checkout at `/home/zhuhe/code/litchi-worktrees/before-70d7768cc/tools/perf-baseline`, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0656-before` |
| After leg | the same command in the branch worktree, `CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/0656-after` |
| Binaries | staged outside both target directories before any measurement (change 0627: a concurrent build relinked one mid-run) |
| `litchi-perf-baseline` before | `325c9d6e1aa18a5425a97eb7a90f42ab59c0c871d4f43de41039e4cdc062a51b` |
| `litchi-perf-baseline` after (windows 1-3) | `2818d8913ebf77d4493499eb25ff2855ede5aa2dee41add31c5e511ad3214281` |
| `litchi-perf-baseline` after (committed tree, window 4 and `counts-final/`) | `96c7a42e96eb20a65685eeb20cd1840662050bd15631c026c98c6943fc6c4bef` |
| `litchi-perf-baseline-alloc` before | `6e288e4ab87f4a22437ea32b36dfdb41744596bfeff1700801111cea2c35a770` |
| `litchi-perf-baseline-alloc` after (`alloc/`) | `10e39fc6fb30941669327795afa61be6cf978c640a7d8f6c5efbe67edc7e8b8e` |
| `litchi-perf-baseline-alloc` after (committed tree, `alloc-final/`) | `c5efdf2aa6beb473e5d375c5bc66085e55ab72a0ec14c6389b09d41fa474aaa2` |
| Host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14), valgrind 3.26.0 |
| Pinning | every measured process `taskset -c 11` |
| Host load | seven other agents of the same wave were building and measuring on the other cores; each timing window records its own load average beside its summary, and the A/A and B/B floors are the only statement about it |

`lto = true` makes release binaries non-reproducible byte for byte (change
0635), which is why the sha256 of every binary that was timed is recorded rather
than the build command alone. The two legs' harness sources differ by one
argument, because `Limits::new` gained a sixth parameter and
`tools/perf-baseline/src/pptx_slide_boundaries.rs` calls it; that call site is a
boundary probe outside every measured selector.

## What this packet does not establish

No speedup is claimed and none is registered. Everything here describes one
implementation on two synthetic corpora, on this host, in these two builds.

The call counts, instruction differentials and allocator counters are exact and
deterministic. The wall-clock and `perf stat` numbers are not, and the floors are
reported beside them; window 1's floor exceeds 5% and its selector totals are
disclaimed rather than used. Callgrind runs SHA-256 in software because valgrind
masks the SHA CPUID bit (change 0649), so any instruction share attributed to
hashing is about five times its native cycle share, and it prices
`rep movsb`-class copies per byte (change 0604); `perf stat` is the
counterweight and it is whole-child, so it dilutes rather than isolates.

Both corpora are generated by `litchi-pptx-cross-slide-copy-evidence-v1` and the
media-rich one is deliberately incompressible. The largest real `.pptx` under
`test-data/` is 972,788 bytes, so no real deck here exercises the large end of
the retention budget. No source-backed cross-copy selector, no external-package
fixture, no repeated application of one plan, no cold-cache, range-source,
concurrency or cross-platform measurement was taken.

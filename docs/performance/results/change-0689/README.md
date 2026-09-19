# Change 0689 evidence packet

Final candidate retained after source, correctness, performance and evidence
review. See [review.md](review.md) for the disposition and limitations.

The [change record](../../0689-xls-sst-chain-checkpoint.md) describes the
XLS retained SST chain checkpoint experiment. Baseline is `9ba2c0b6b`.
The fresh baseline profile assigns 57.42% self samples to the checked chain
walk. Diagnostic route tracing identifies 478 repeated SST links versus 14
worksheet links on the 54016 first-cell route.

## Measurement separation

- `native/`: unchanged 0686 direct-source probe, 12 cases and two sources,
  100 fresh owners per leg after three warmups, eight selected queries.
  Baseline A/A then candidate A/B/B/A: 14,400 owner records, 115,200 queries.
  Source construction and outcome projection are outside native timers.
- `costs/`: unchanged 0686 allocation companion and 0684 counted-source
  companion. There are 96 allocation groups with three identical repeats per
  binary (576 captures), and 12 counted eight-query routes. These instrumented
  timings are not native latency. Measured owners/results remain alive through
  allocation capture; requested bytes and live deltas have different meanings.
- `repeat/`: unchanged 0684 same-owner probe, 12 groups, six legs, nine process
  samples per leg, 50,000 queries per sample. Every run uses the default 2 MiB
  cache ceiling, including the missing target whose native case uses 1 MiB.
  Per-query values are loop means; these are not 50,000 independent latency
  samples. Two preparatory queries occur before its loop timer.
- `diagnostics/`: the same repeat probe at N=10 and N=100,010, three repeats
  per binary on eight groups. Counter subtraction divides by 100,000 extra
  queries. Whole-process counters include setup and the timing wrapper;
  native-child RSS uses `perf stat -- time -v probe`. RSS and allocator gauges
  are separate quantities, neither proving a process memory bound.
- `corpus/`: unchanged 0684 differential probe, all 126 real XLS fixtures on
  owned/file sources, plus the unchanged 0686 generated 70,001-cell fixture
  with full visitor count and semantic-digest parity.
- Matched profiles run two million selected queries after two preparatory
  queries, with perf sampling at 997 Hz. Symbol reports, stderr and command
  bindings remain; raw perf data is removed after review.

CPU 12, two Cargo jobs, warm OS file caches and a shared host are recorded in
`environment.json`. No cold-device, remote, scaling or cross-platform claim.

## Reproduction and auditing

All five probe packages and their locks are reused unchanged from 0684/0686.
`build.py baseline` was run before changing production sources; candidate
builds ran after final checks. The driver uses a shared compilation target and
copies each revision's binaries into separate before/after directories.
For reproduction, use the baseline revision plus copies of the 0689 drivers,
then the candidate revision, retaining the before binaries. Adapt absolute
paths and CPU affinity consistently. Separate worktrees may also be used.

Run `build.py PHASE`, then `measure-native.py PHASE`, `measure-costs.py PHASE`,
`measure-corpus.py PHASE`, `measure-repeat.py PHASE` and
`measure-diagnostics.py PHASE`, for baseline and candidate. Run
`profile.py PHASE` and `inspect-assembly.py PHASE` for each revision; these
retain matched profile reports, helper/cursor/walk disassembly and section sizes.
Every driver records source, binary, probe, fixture and raw-output bindings.

Run `audit-costs.py`, `audit-native.py`, `audit-extra.py` and `audit-final.py`;
`summarize.py` renders phase medians, all paired median regressions and separate
mean/tail triggers. The native audit includes all control drift and 1,000-draw
paired median bootstrap intervals. Intervals do not remove shared-host drift;
100-sample p99 remains descriptive. No geometric mean hides individual results.

`run-integration.py` retains six owner/facade quality gates;
`check-consumers.py` adds DOC/PPT tests. `run-evidence.py` invokes the existing
boundary, claim, coverage and non-iWork gates. Final source and command bindings
are independently checked by `audit-final.py`. No iWork or coverage promotion
is part of this change. `performance_claim: none`.

Use `LITCHI_GATE_OUTPUT=docs/performance/results/change-0689/final-verified`
with `run-integration.py`; run `final-doc-gates.py` after final report edits.
Cargo durations record incremental build logistics, not clean-build latency.
All 96 allocation groups and all 12 counted routes retain complete raw
comparisons. Added index storage is charged and reported separately.

`trace-routes.py` and `trace-candidate-routes.py` temporarily instrument route
positions, build a separate diagnostic binary in the shared target, then restore
all edited source bytes in `finally`. Run only after source-bound native
captures have finished, with no concurrent source editor. `audit-traces.py`
compares route counts; no instrumented timing enters the latency report.

`measure-followup.py` uses the frozen native binaries for four flagged groups,
1,000 fresh owners per leg after ten warmups, A/A then A/B/B/A: 24,000 owners
and 192,000 queries. `followup/` contains deterministic gzip JSON raw captures;
compression happens outside probe timers. `audit-followup.py` verifies all
hashes, outcomes and reported statistics. Original native captures are retained
and not pooled with this follow-up. See `followup-summary.md` for costs as well
as improvements. `initial-checks/` preserves the initial lifecycle-test failure
and differing test source; final passing checks use the corrected test lifetime.

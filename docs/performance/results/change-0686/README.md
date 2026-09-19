# Change 0686 evidence packet

Final quality and evidence audits pass. See the [change record](../../0686-xls-oversized-index-admission.md) for results and limitations, and [review](review.md) for retention disposition.

Baseline `fb0416104` includes the worksheet occurrence cache and immutable CFB
checkpoints. The hypothesis is that successful full-scan occurrence counts can
prove optional index construction cannot fit the fixed local limit, removing
repeated allocation without memoizing errors or temporary pressure.

`cases.json` freezes large stored/missing queries, medium and tiny sheets,
formula refusals, and disabled/too-small/default limits. `measure-costs.py`
uses a separate allocation companion and the unchanged 0684 counted-source
budget probe; their timings are not native latency. Allocation queries keep
owner and measured result/error alive through gauge capture. Successful
reallocation contributes its requested size, while live gauges use its delta.
Three allocation repeats must agree exactly. Counted outcomes must match exactly. Open and first/build-query I/O must match;
later I/O changes from newly admitted indexes are retained individually.

`measure-corpus.py` reuses 0684's corpus differential probe for owned and file
sources. It compares complete JSON, including refused fixture outcomes.
`run-integration.py` records final quality command/source bindings. Root is the
serial Cargo lane. CPU 12 is used for measurements; file inputs use warm OS
caches on a shared host. No physical cold-storage or parallel scaling claim.

Each standalone probe has a committed Cargo.lock. Build identical sources in
separate baseline/candidate worktrees with `--release --locked --offline`, using
separate target directories. Drivers retain exact commands and absolute binary
paths; adapt those paths and CPU affinity consistently when reproducing.
`audit-costs.py` verifies immutable baseline sources, live candidate sources,
probe/fixture/raw hashes, identical repeats and outcome/first-query I/O parity.

## Retained evidence and reproduction

The final native matrix has 26 case/source groups, six legs per group,
100 fresh owners per leg and eight queries per owner: 15,600 owner records
and 124,800 query records. `native-comparison.json` retains phase distributions,
A/A drift, paired comparisons and bootstrap intervals. `regressions.md` and
`tail-regressions.md` retain review triggers, including unchanged controls.

Allocation companions cover 104 case/source/phase groups, three repeats per
binary (624 captures). Counted-source comparisons cover 13 eight-query routes.
`generator-manifest.json` binds two deterministic fixtures and two identical
generation runs; full visitors check 70,001 and 100,001 cells for both sources.
These generated inputs do not establish native Office compatibility.

Run `build.py`, the `measure-*.py` drivers and `profile.py` with their recorded
arguments after adapting only worktree/target paths and available CPU affinity.
Each driver retains its exact commands in manifests. Then run
`audit-costs.py`, `audit-native.py`, `audit-diagnostics.py`, `bind-profiles.py`
and `audit-final.py`; `summarize.py` renders comparison tables. Source changes
require new builds, checks and captures. Archived absolute paths identify the
original run; binaries and external perf data are cleaned after auditing.

Final diagnostic RSS uses `perf stat -- time -v probe`, so the RSS figure
belongs to the native child. The initial enclosing-perf RSS captures are kept
in `initial-perf-wrapper-rss/` and excluded from the final comparison. Counter
subtraction accounts for both the warmup and measured owner (2,000 extra
queries). Counters include projection/reporting and the small timing wrapper.
Profiles have 11,414 baseline versus 75 candidate samples; candidate function
percentages are not precise evidence after the workload became much shorter.

`initial-check/` retains the corrected test type-inference failure;
`initial-budget-expectation/` retains the obsolete 1 MiB fallback assertion;
`initial-before-review/` retains the rejected post-failure reservation fallback
and its normal-allocator checks. None supplies final candidate timings.
`final-verified/` binds all six passing quality gates to the measured source.
There are 1,942 passing Rust tests and two existing ignored tests.

Warm OS caches, shared-host drift, higher construction peaks and retention,
formula/build latency regressions and the short-run RSS increase limit the
claim. No physical cold-storage, parallel scaling, universal speedup or process
memory bound is asserted. `performance_claim: none`.

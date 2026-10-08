# 0970 historical-baseline affinity diagnostic design v3

Status: fresh bounded design, 2026-10-08. This document authorizes no source
edit, Cargo command, build, benchmark launch, capture, or qualification
rerun. The role of 0970 is to diagnose the exact historical ordinary baseline
binary that produced the 0969 A/A packets. It does not measure a current
revision, a candidate, or a candidate effect. Reusing that identical binary
isolates launcher affinity as the planned factor; rebuilding it would add
another variable and is forbidden.

## Custody and scope

The exact binary is the retained ordinary baseline executable used by every
0969 A/A child:

```
/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile
SHA-256 0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb
```

The 0969 bounded baseline custody/reuse review is
`/home/zhuhe/litchi-performance-0969/review/baseline-reuse-v1.json`, SHA-256
`dbd362dae8e31dfe5ff54d6af5c37883df829bede3a5e5fa19e04337e0a8312e`, with
status `pass_bounded_output_oracle_and_matched_baseline_reuse`. Its binary
identity and focused runtime/harness/configuration/fixture scope are evidence
for reusing this historical executable; they are not permission to rebuild or
to claim whole-tree identity.

Baseline admission must use the 0969 v1 per-file manifest and its
`baseline_sha256` entries:

* `/home/zhuhe/litchi-performance-0969/manifests/source-file-manifest-v1.json`,
  SHA-256 `99e65d43e0feed89133eec60b095130d68058a071a02e8a15b05dea01209bb6c`;
* `/home/zhuhe/litchi-performance-0969/manifests/source-manifest-v1.json`,
  SHA-256 `f1818af5c306fb6ea01b59ca8df29710e773386b246071202bd4aa896f4e8e2b`.

The runner validates all 924 v1 entries against
`/home/zhuhe/litchi-performance-0969/source-baseline`, using each entry's
`baseline_sha256`, byte count, mode, and relative path. It must never use a
v3 `candidate` entry to admit the historical baseline. A mismatch stops
before a child. This is a focused per-file custody check, not a blanket
whole-checkout equivalence claim.

The actual 0969 fixture paths and hashes are:

* `/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/NoHeadFoot.doc`,
  26,112 bytes, SHA-256
  `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48`;
* `/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc`,
  335,360 bytes, SHA-256
  `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d`.

The sealed 0969 validator inputs are reused read-only:

* plan `/home/zhuhe/litchi-performance-0969/plans/plan-0969-aa-v6.json`,
  SHA-256 `e853030180d7464c06cf1ed5d525957ce2d50514b8337a1b168392374d00b5df`;
* `tools/protocol_common_v4.py`, SHA-256
  `5ce0f2ac08fc7909df054a4cc2d5137b3279af4cd726acf5bfb4558d97b05c5c`;
* `tools/run_0969_v4.py`, SHA-256
  `c7315d5c5311c5643aebca16e8cf2b931ee1b74663eb5a8a31e8185f504ddd20`;
* `tools/reduce_0969_v4.py`, SHA-256
  `cbbbb6d09a981546a9db216c18e4c5e8da4df7202d56dbe3dc7d1c999f2ec2a9`.

The 0969 raw A/A evidence used for offline rationale is
`/home/zhuhe/litchi-performance-0969/analysis/aa-independent-audit-v1.json`,
SHA-256 `f0d0d0b63a896be41dab84550ff39f59954c26538423be2a3d3d017df9f26775`,
and the sealed offline variance summary is
`/home/zhuhe/litchi-performance-0970/analysis/aa-variation-v1.json`,
SHA-256 `9e41ff784382ae5d225ce751d3398eac1ee1ff56500d6e229d68508ef07f710a`.
Neither is a new capture or a candidate result.

## Fixed host/tool bindings

The reviewed host snapshot is
`/home/zhuhe/litchi-performance-0970/host-audit-v1.md`, SHA-256
`37e204363bee72eb7d3080300dd32ce5ce21d53be86dee2c9e726b086f6901c5`.
It records 32 allowed CPUs, current cpuset `0-31`, no observed current
throttling, and unavailable cpufreq/thermal interfaces. It does not recreate
the 0969 host state. The executable identity addendum is
`/home/zhuhe/litchi-performance-0970/host-tool-addendum-v1.json`, SHA-256
`0e013628d436d00b2ca13b17c89458cf6288c56655b442e1b20965c90cc79a81`.
The future runner binds these exact tool identities:

* `/usr/bin/perf`, SHA-256
  `174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b`;
* `/usr/bin/time`, SHA-256
  `919efdcc04c1dbc7a1479c2f3310d491c95f028e1762a84c47810ac0ec1c0d88`;
* resolved `/usr/bin/python3.14`, SHA-256
  `be9a2a5eada8c89c1c399fdfb8397179e877c20d8db0739c802f726d4d0e69fd`.

The fixed six-event capability probe is
`/home/zhuhe/litchi-performance-0970/telemetry-probe-v1.json`, SHA-256
`8cb97d377c6162c57e1ae092c7c8e33f9b2b23c38aa2d44f188d3831e20614dc`.
It successfully enabled exactly `task-clock,cycles,instructions,
context-switches,cpu-migrations,page-faults` against `/usr/bin/true`.
The child self-report affinity probe is
`/home/zhuhe/litchi-performance-0970/affinity-probe-v1.json`, SHA-256
`0f2e97080d2ba9898a3c0c74660f3280a4cbb66e4546997fa93749d476ce6a17`.
It selected the highest allowed CPU, observed `[31]`, restored the child's
mask, and changed no parent or global setting.

Before any future root GO, the runner must take a fresh process-quiescence
snapshot and verify that no Cargo, rustc, test, benchmark, or other workload
is active. It must repeat that check before the first child. A failed check,
changed allowed mask, missing tool, or changed tool hash stops before capture.

## Diagnostic question and affinity conditions

The four fixed cells are `docnohf/e1/default`, `docnohf/e32/default`,
`docfloat/e1/default`, and `docfloat/e32/default`. The diagnostic asks:

1. Does a single fixed CPU change each child's within-child tail, using
   `p95/p50` over the 80 native samples?
2. Does it change dispersion between fresh children, using the distribution of
   each child's p50 and p95 summaries?
3. Do process CPU time, context switches, migrations, and page faults move
   with either effect?

After the fresh audit, derive `pin_cpu = max(sorted(os.sched_getaffinity(0)))`.
The current audited host therefore proposes CPU 31, but the fresh audit is
authoritative. Pinned children inherit the one-element mask `{pin_cpu}`.
Unpinned children inherit the full audited allowed mask. A pin outside the
allowed set prevents capture.

A short Python self-report child runs once for each condition before any
benchmark child. Its exact program prints one JSON object with `before`,
`target`, `after`, and `restored` affinity lists; the parent checks that the
child's own `after` list is the full audited mask for unpinned or `[pin_cpu]`
for pinned. This is the affinity proof; no fast-PID `/usr/bin/true` race is
used. A malformed report, wrong mask, nonzero exit, or timeout stops the run.

## Frozen schedule and exact wrapper ownership

There are 12 pairs per cell, 24 serial children per cell, and 96 children in
total. Pair numbers 1,3,5,7,9,11 run unpinned then pinned; 2,4,6,8,10,12
run pinned then unpinned. Each child invokes the retained CLI with exactly
`--case`, `--input`, `--edits`, `--retention default`, `--warmup 5`, and
`--samples 80`; there are five warmups and 80 measured samples. The analysis
seed is `970001` with 10,000 paired resamples. Packet and sidecar paths use
exclusive creation and contain the unique cell, pair, condition, and order.
No retry, replacement, or selective rerun is permitted.

The exact benchmark wrapper argv is frozen as this nesting, with placeholders
expanded only to the packet's unique paths and cell values:

```
/usr/bin/perf stat -x, -o <perf-sidecar> \
  -e task-clock,cycles,instructions,context-switches,cpu-migrations,page-faults -- \
  /usr/bin/time -f <fixed-time-format> -o <time-sidecar> \
  /usr/bin/python3.14 /home/zhuhe/litchi-performance-0970/tools/launch_baseline_v1.py \
  --metadata <launch-sidecar> --requested-affinity <expected-mask> -- \
  <historical-baseline-binary> --case <case> --input <fixture> \
  --edits <edits> --retention default --warmup 5 --samples 80
```

`launch_baseline_v1.py` is a tiny external launcher, not a monitor. It checks
its requested mask, records its own PID, actual `os.sched_getaffinity(0)`,
inherited cgroup path, UTC timestamp, and the exact `execve` argv in an
exclusive, fsynced launch sidecar immediately before replacing itself with
the exact historical baseline binary. The sidecar records the helper hash and
its unique packet identity. `execve` preserves the PID and affinity, and the
retained baseline source has no affinity mutation. The runner binds the helper
source/hash and its self-test receipt before capture; every packet binds the
launch-sidecar hash, PID, requested mask, actual mask, and cgroup. No wrapper
PID inference or polling monitor is used. A requested/actual mask mismatch,
missing sidecar, wrong helper hash, or failed `execve` stops the run.

The outer `perf stat` measures `/usr/bin/time`, the launcher startup, and the
execed historical baseline in one process tree; its counters therefore include
that wrapper scope. GNU time likewise includes launcher startup and then
continues across the same-PID `execve`, while its sidecar is owned by that
baseline launch. The harness's `Instant` samples exclude perf/time/launcher
startup and sidecar I/O. These scopes are recorded explicitly and are never
presented as interchangeable timing metrics.

The fixed GNU time format is:

```
{"user_s":%U,"system_s":%S,"voluntary_context_switches":%w,"involuntary_context_switches":%c,"minor_faults":%R,"major_faults":%F,"max_rss_kb":%M}
```

The capability probe freezes perf enabled for every child. If a later
admission finds the frozen capability invalid, the run stops before the first
benchmark; it does not fall back to a new event list or run mixed wrappers.

The launcher receives the expected mask through its explicit argument. For a
pinned child the parent launches the wrapper under a one-element
`preexec_fn` mask `{pin_cpu}`; for an unpinned child it leaves the audited full
mask unchanged. The launch sidecar's self-observation immediately before
`execve`, together with the separate child self-report control, is the
launch-path proof. There is no race-prone observation of a short-lived outer
`perf` PID. A missing observation or requested/actual mismatch invalidates the
packet and stops capture.

## Packet gates and sealed validator reuse

The runner must call the read-only 0969 packet validator for the existing
report/oracle fields, then add only the diagnostic custody and telemetry
checks. Every child must have return code zero, no timeout, the exact bound
binary and fixture hashes, and a unique packet identity. The report must have
`warmups == 5`, `samples_requested == 80`, exactly 80 sequential sample
records indexed 0 through 79, and 80 finite strictly positive `elapsed_ns`
values. The 20 existing correctness/oracle fields must remain booleans and
true, output/reference hashes and route identities must match, and semantic
change count must equal the requested edit count. No output is accepted merely
because the process returned zero.

Each packet must have one time sidecar, one perf sidecar, and one launch
sidecar. The launch sidecar must contain the bound helper hash, exact baseline
argv, self PID, requested mask, actual mask, inherited cgroup path, and a
pre-`execve` UTC timestamp; its PID must be the PID whose `execve` becomes the
baseline. The time sidecar must parse as one object with finite nonnegative
values and integer switch/fault fields. The perf sidecar must have exactly one
valid CSV row for each frozen event, with no `<not counted>` or missing value.
Sidecar names, packet IDs, binary/fixture hashes, cell, pair, condition,
order, helper hash, and observed affinity must agree. Any missing, duplicate,
malformed, failed, or wrong-identity packet makes the entire diagnostic
`incomplete`; no partial result or retry is accepted.

## Offline reducer and interpretation

The reducer is diagnostic-only and does not call `perf_decision`, emit a
qualification status, or apply a performance threshold. It preserves all raw
packets and sidecars. For each child it reports elapsed p50, p95, mean, sample
spread, and `within_tail = p95/p50`. For each condition and cell it reports
between-child distributions over the 12 child p50 and p95 summaries: median,
nearest-rank p25, p75, p95, IQR/median, MAD/median, p95/median, and the
between/within CV ratio. It reports first-versus-last child summaries and
order-stratified values so chronological drift is visible without a
bimodality or host-cause claim.

The paired comparison reports pinned/unpinned ratios for child p50/p95,
ratios of between-child IQR and MAD measures, within-child tail changes, and
paired differences for GNU time user/system seconds, voluntary and
involuntary switches, minor/major faults, max RSS, and each perf event. It
also reports whole-process descriptive summaries for `task-clock`, `cycles`,
and `instructions`, plus `cycles/instructions` and `cycles/task-clock` where
both operands are present. These ratios can describe signatures of equal-work
elapsed variation; they are not a CPU-frequency measurement and are never
attributed to the harness `Instant` region. Missing, not-counted, or
multiplexed counters remain explicit and are never replaced with zero. It
also reports the exact ownership note that perf includes the time wrapper
while GNU time covers the baseline launch. Paired bootstrap intervals use seed
`970001` and 10,000 resamples; no pass/fail boundary is applied.

The only diagnostic statuses are `complete` and `incomplete`. A complete
report may describe an affinity-associated pattern when paired summaries and
telemetry point in one direction, or report no separating signal. It must
retain host limitations: the 0969 quiescence receipt has no historical child
affinity, cgroup, frequency, thermal, migration, or scheduler counters, and
current cpufreq/thermal interfaces are unavailable. No report may claim a
host root cause, speedup, slowdown, candidate rejection, or adoption result.
The 0969 `uninterpretable` readiness decision remains unchanged.

The offline 0969 variance analysis found that fresh-child dispersion can be
much larger than within-child sample dispersion and documented a docfloat
high/low sequence candidate. Those are descriptive inputs for the metrics
above, not a reason to select a favorable cell, relax a gate, or declare hash
seed, migration, or frequency as causal.

## Review and implementation boundary

The accepted ADR set was reread and revalidated before this design: 32 files,
3,147,627 bytes, root
`/home/zhuhe/litchi-performance-0969/source-baseline/docs/adr`, expected set
SHA-256 `bf84c9e75149dda400f9f93c4c72f07b26afa5969ff133c9f4b4427421a9bc40`.
The read-only source-variability review is
`/home/zhuhe/litchi-performance-0970/source-variability-review-v1.md`, SHA-256
`36fb53596eb1b4c635291532092510552d4c38b1ee272eaf29cc9262fe7c34ef`.

After independent review of this v3 design, a small external runner and
reducer may be written under `/home/zhuhe/litchi-performance-0970/tools`.
They must remain workload-inert in self-tests, reuse the sealed validator,
validate the v1 baseline per-file hashes, freeze the exact wrapper and
telemetry ownership above, and use fresh exclusive output roots. No helper
may start a native child until separate review and root GO bind this v3,
the host/probe/addendum receipts, all custody hashes, and the final helper
hashes.

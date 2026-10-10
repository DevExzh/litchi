# 0179 — XLSB sheet-command coalescing

## Contract

Closure transactions and atomic JSON batches now coalesce consecutive XLSB
sheet-visibility commands into one native structural edit. An eligible command
uses an absolute `/sheet[N]` path and one `hidden` or `visibility` assignment.
One run admits distinct tabs assigning the same exact native state. Native
staging supplies each item's changed/no-op result; native publication remains
the authority for signatures, view layouts, preservation and durable patches.

Runs close on repeated tabs, state changes, named/filtered/query targets,
name or combined assignments, empty assignments, reads, other edit families,
and expanded-target or staged-byte boundaries. A changed hide of the active
tab also closes its run before the next command. This last barrier is required:
hiding A can require selecting B, whose view may be unsupported; combining a
later hide of B must not bypass that failure by selecting C directly.

A shadow visible-tab count preserves each command's last-visible admission
and existing `invalid_argument` error. Plans are recomputed after a flush,
including when an initial plan failed against the old snapshot: revealing B
and then hiding the previously sole visible A must succeed. Deferred native
failures are replayed to find their originating command index; an earlier
pending failure takes precedence over a later malformed/admission failure.
The existing up-front operation-quota policy is unchanged. Snapshot accounting,
source fences, transaction poisoning and best-effort isolation retain the
existing transaction contract.

No property vocabulary, capability, native public API or OfficeCLI behavior
changes. Sheet-name coalescing and more general state/selector planning remain
follow-ups. The full all-format/Python/CLI/acceptance goal remains open.

## Verification

Independent implementation, regression-test, benchmark and review agents worked
in parallel, with implementation artifacts integrated from isolated worktrees.
Regression coverage includes exact sequential output and view parity, changed
flags and inverse restoration, the missing intermediate chart-view failure,
last-visible item attribution, named/query/cell/empty/combined barriers,
reveal-then-hide replanning, best-effort rollback, quota boundaries and 1,000
non-active commands publishing in one commit.

Verification passed: 1,229 combined facade tests/doctests using
`--no-default-features --features docx,xlsx,pptx,xlsb,encryption`, followed by
all twelve final focused sheet-batch regressions after the best-effort test was
added. Combined and minimal `xlsb` all-target Clippy passed with warnings denied,
as did warning-denied facade rustdoc, scoped rustfmt, crate boundaries and the
generated OfficeCLI parity check. The production code remained unchanged after
these final checks and before the after measurements.

Cargo verification used an external copied root lockfile with cached `cc` 1.6.0
through `resolver.lockfile-path`. Root and shared performance lockfiles remain
unchanged. Temporary implementation/test/benchmark worktrees were removed after
integration. The base branch has no commits missing from this branch.

## Measurement scope

The new `tools/perf-baseline/unified-xlsb-sheet-batch` harness compares N prepared
direct-path commands in one `File::edit` closure with N native state setters in
one native edit. Its first tab is an active visible sentinel; the remaining N
tabs are hidden. This is a homogeneous non-active-tab workload. Active-tab
relocation barriers and atomic JSON dispatch are tested for correctness but
are not separately timed by this harness. Queries are not used.

The before matrix covers 1, 10 and 100 commands with three ABBA rounds. Larger
pre-coalescing cases repeat whole-package publication per command and were
not run. At 100 commands the unified/native ratio was 122.65 eager and 91.02
source. Subsequent evidence records source and binary identities, raw samples,
corpus hashes, actual facade/native commit counts and exact output comparison.


## Results

The [before](results/0179-xlsb-sheet-batch-before.json),
[after](results/0179-xlsb-sheet-batch-after.json), and
[repeat](results/0179-xlsb-sheet-batch-repeat.json) matrices retain complete
machine-readable evidence. Both after matrices cover 1, 10, 100, 1,000 and
10,000 commands for eager/source owners, with three ABBA rounds (six samples
per leg) on CPU 0 and no concurrent build/test jobs. Every output comparison
is byte-exact, every failure matrix is empty, and every after/repeat facade
batch records one native commit.

| Owner, command count | After native / unified ms | Repeat native / unified ms | Unified/native ratios |
|---|---:|---:|---:|
| eager, 1,000 | 137.91 / 156.50 | 139.76 / 157.26 | 1.135 / 1.125 |
| eager, 10,000 | 8799.22 / 9963.86 | 8693.65 / 10064.88 | 1.132 / 1.158 |
| source, 1,000 | 171.93 / 188.66 | 171.37 / 188.35 | 1.097 / 1.099 |
| source, 10,000 | 10065.54 / 11684.79 | 10049.46 / 11841.32 | 1.161 / 1.178 |

Measured 1k and 10k overhead stays below 2×. Single-command ratios are
1.028/1.079 eager and 1.027/1.052 source, below the measured 10% threshold.
At 100 commands eager facade time drops from 940.43 ms to 7.66/7.60 ms,
and source facade time drops from 944.45 ms to 10.36/10.31 ms.

Linear scaling is **not** accepted: unified 1k→10k growth is about 62–64×,
and native publication itself grows about 59–64×. Whole-package publication
needs profiling rather than attributing this native growth to the facade.
The facade also still performs a full catalog validation scan per direct
command; replacing that repeated scan with the run's admitted catalog/count
is a separate follow-up. The full goal remains active.

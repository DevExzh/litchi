# Allocator peak-counter correction

This batch repairs the benchmark allocator's process high-water calculation
and labels corrected reports with `post_update_peak_v2`. The exact candidate
revision and clean source manifest are in `build-candidate.json`. Document
production behavior is unchanged. See the [change record](../../changes/0421-allocator-peak-counter.md)
and [affected-evidence notice](affected-evidence.md).

The frozen protocol selects the existing media-rich and plain owned PPTX
lifecycles, two fresh processes per selector, 30 samples and three warmups.
CPU 2, one worker, warm generated input, Rust 1.98.1 and the established
release/debug/frame-pointer flags are fixed. The host is shared and background
activity is uncontrolled. Allocator elapsed times are not compared.

`summary.json` and `result-table.md` describe a corrected current baseline.
Reports retain source/corpus/output bindings, preservation/refusal gates and
allocation vectors. The replay checks high-water is at least live bytes at
each quiescent boundary. This necessary invariant alone cannot certify an old
report: old values can satisfy it accidentally. The counter-revision identity
is required separately. Raw historical reports are neither repaired nor paired
with these observations.

## Replay

From the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0421/summarize.py --replay
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0421/check-report-guards.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0421/check-portable-tool-binding.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0421/portable-replay.py
```

The portable export uses this bundle, including four retained validator modules
under `replay-tools/`. It verifies them against their pinned manifest before
export and requires no original worktree, repository checkout, or copied binary. `SHA256SUMS` inventories the
retained files; `compression.json` binds compressed logs to their original
bytes, including failing pre-fix tests. `cleanup.json` records task temporary
removal while preserving shared caches.

The portable checks replay all four baseline reports, run five actual-report
guards, and validate the separate normal-binary identity report. That normal
report has one sample and no warmup; it is a functional serialization check
excluded from the baseline. The first isolated export exposed an absolute-path
assumption in the summary verifier. Its failed receipt is preserved under
`checks/portable-replay-path-failure.*`; the verifier now binds recorded command
paths to their historical root while reading hash-checked artifacts from the
current export. No captured report or journal was changed.

[Validation](validation.md) distinguishes the pre-fix failures, passing tests,
final-code checks, and the existing Clippy warnings.

## Fresh capture

Create a detached clean worktree at the recorded candidate revision. Copy the
frozen `protocol.json` into a fresh output directory, then invoke with Python
and `PYTHONDONTWRITEBYTECODE=1`:

```text
build.py candidate WORKTREE --root OUTPUT --binary-prefix /tmp/UNIQUE_PREFIX
capture.py --root OUTPUT
summarize.py --root OUTPUT
```

The build script reuses the checked-in 0418 source/host identity helpers and
shared `tools/perf-baseline/target` cache. It builds both normal and allocator
binaries; only the allocator binary supplies this baseline. Keep tests,
builds and measurements serialized. Drivers refuse existing output. Remove
only reproduction-owned worktrees and binaries.

## Scope and next work

The fix changes only benchmark accounting and report compatibility. It adds
no unsafe code, dependency, production parallelism or document API. ADR 0005's
measurement contract requires correcting the counter before using it for
optimization decisions. Existing overflow and unavailable statuses remain.

Corrected peaks describe logical request sizes in allocator-observer order.
They include setup and historical allocations, exclude hidden allocator-internal
copy overlap, and are not an operation-local peak or aggregate memory budget.
Next work should implement a separate region tracker without resetting global
counters, measure retained objects through explicit drop boundaries and near
limits, then compare the admitted source-backed media lifecycle against the
owned path with matched semantics and timer boundaries.

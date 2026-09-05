# Bounded PPTX archive allocation experiment

This bundle compares `7c7561160` (control) with `0320a6a88` (candidate).
The harness is identical. The candidate changes the private cross-copy archive
buffer to capped geometric growth, followed by one fallible exact-size
reservation and copy when spare capacity exists. It preserves the existing
archive byte limit, typed allocation errors and exact-source authorization.

The experiment separates three scopes:

- Heaptrack normal-binary captures: one sample, zero warmups, **whole command**
  including corpus generation, semantic/refusal gates and teardown.
- Normal ABBA: 100 samples and 10 warmups per process for media-rich and plain
  owned cross-copy lifecycles. These are timing diagnostics, below the
  500-sample release latency claim threshold.
- Allocator ABBA: 30 samples and three warmups, with operation counters. Its
  elapsed times are not compared. Requested bytes include the full requested
  size of successful reallocations, not just growth or physical copying.

Both builds use Rust 1.98.1, identical release/debug/frame-pointer flags and a
shared build cache. Each measured process uses CPU 2 and one worker on the
recorded AMD EPYC 9R45 KVM host. Sources are generated in memory and warm;
shared-host background activity is uncontrolled. `host-tools.json`, both build
identities and the run journals retain the exact environment and commands.

`summary.json` is generated from all 16 successful run journals and raw reports.
It retains each leg, both paired directions, same-revision drift, allocation
means/ranges and whole-process RSS. No pooled speedup or general memory
improvement is claimed. Live and peak allocator counters are process snapshots;
they do not establish an operation-local peak. The final-copy transient can
coexist with the grown buffer, and the output bound is not an aggregate memory
budget.

The first capture wrapper incorrectly rejected empty successful stdout. Its
complete available artifacts are retained under `failed-attempt-1`, excluded
from the final ABBA. The corrected wrapper explicitly accepts empty logs and
has a retained self-test. The build-start protocol's illustrative CLI typo was
corrected before tracing; `protocol-amendment.md` records both original bytes
and the narrow accepted correction.

## Replay retained evidence

From the repository root:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0419/summarize.py --replay
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0419/analyze-heaptrack.py --replay
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0419/verify-probes.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0419/capture-selftest.py
PYTHONDONTWRITEBYTECODE=1 python3 docs/performance/results/change-0419/portable-replay.py
```

Post-cleanup replay passed after removing both build worktrees and all four
copied binaries. The isolated export also passed with only this bundle and
the four hash-bound validator modules listed in `checks/portable-replay.json`.

The numerical heap replay needs `zstd`, Python and the retained symbolized
traces; it does not need the removed binaries or worktrees. Its streaming
parser reproduces the unfiltered Heaptrack histogram's allocation count and
requested-byte sum, then attributes events by retained stack ancestry. The
concrete writer implementation is classified separately from generic function
names containing `BoundedVecWriter`. Each top-trace list is explicitly capped;
aggregate totals include every matched trace.

Heaptrack 1.5.0's `--filter-bt-function` does **not** filter the `-H` histogram:
the histogram is filled before backtrace filtering. The parser and source
binding in `analyze-heaptrack.py` make this limitation explicit. Counts are
allocation events; a zero net live delta at command exit does not mean zero
transient memory use. The raw producer stderr files may use gzip sidecars;
their original decompressed hashes and lengths remain authoritative.

## Reproduce fresh measurements

Create clean detached worktrees at the two recorded revisions. Choose a fresh
output directory and binary prefix, copy `protocol.json` and
`measurement-protocol.json` into that directory, then run the scripts from
this checkout in order:

```text
build.py control CONTROL_WORKTREE --root OUTPUT --binary-prefix /tmp/UNIQUE_PREFIX
build.py candidate CANDIDATE_WORKTREE --root OUTPUT --binary-prefix /tmp/UNIQUE_PREFIX
trace.py control --root OUTPUT
trace.py candidate --root OUTPUT
analyze-heaptrack.py --root OUTPUT
capture.py --root OUTPUT
summarize.py --root OUTPUT
```

Invoke each with Python and `PYTHONDONTWRITEBYTECODE=1`. `build.py` shares the
checkout's existing `tools/perf-baseline/target` cache and copies each binary
before the next build. Do not run tests, builds or other profilers alongside
measurement. The scripts refuse existing capture/build outputs. Hardware,
OS, allocator and source path differences may change results; fresh identities
are mandatory. Remove only the worktrees and copied binaries created for the
reproduction, retaining shared caches.

The source and ADR review is in `source-review.md`; the broader outstanding
program work is in `goal-audit.md`. This experiment does not establish native
producer breadth, physical cold/range I/O, scaling, or source-backed lifecycle
coverage.

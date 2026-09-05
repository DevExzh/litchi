# 0428 managed PPTX cache and budget observations

This bundle adds fallible diagnostics to existing PPTX source owners and
captures their cache, budget, source-read, and process-RSS observations. It
does not implement or claim a workload optimization. The allocator checkpoint
results in 0427 remain a separate experiment.

Implementation revision: `6d0dc4ff0d3f54b0afae09a3e8fe684593741478`. The completed
release capture retains 480 samples and 5,460 phase points. Every final
releasable caller budget is zero, and non-RSS numeric phase observations match
within and across repeats. All 27 repeat flags are process RSS points; see
[resource-review.md](resource-review.md) for their retained, non-causal
interpretation and [resource-audit.json](resource-audit.json) for derived checks.

`protocol.json` specifies two repeats of eight lanes, with thirty retained
samples and three warmups in each fresh release process. Plain and media-rich
cross-copy lifecycles share the existing deterministic corpus builders and
exact output gates. Media-only rows exercise selected-image memory admission,
pinning/eviction, and oversized bypass. Repeated-publication rows use an exact
cumulative output ceiling for three publications.

Source and destination have separate caller budgets. Memory/object/depth
reservation gauges are distinct from cumulative input/output/work charges.
Diagnostics become explicitly unavailable after their owner is consumed or
dropped; caller budgets remain observable without retaining a package.
Process RSS and VmHWM include runtime, corpus, and observer state and cannot
be attributed to a specific owner. See `checks/cache-design.md` for the
ownership model and limitations.

Root serializes all builds, tests, and captures with Rust 1.98.1, four build
jobs, incremental compilation disabled, and one test thread. Workspace checks
use source custody excluding standalone tools while the harness is being
written. Harness and release checks bind both source trees. Failed commands
remain in the evidence; final status is recorded separately from initial
attempts.

`initial-lifecycle-plain.json` and `initial-lifecycle-media.json` are preliminary
one-sample debug executions of the lifecycle-only implementation, bound to
`checks/debug-lifecycle-build.json`. Their schema predates the final resource
rows. They are retained as development history and are excluded from formal
release counts and final-schema validation.
The one-under control and the failed/verified CLI preflight directories likewise
remain outside formal release counts.

Reproduction after the implementation commit:

```sh
python3 -B docs/performance/results/change-0428/check.py --tag release-build -- env RUSTFLAGS=-Cforce-frame-pointers=yes CARGO_PROFILE_RELEASE_DEBUG=1 cargo build --locked --release --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
python3 -B docs/performance/results/change-0428/prepare-build.py
python3 -B docs/performance/results/change-0428/check.py --tag release-capture -- python3 -B docs/performance/results/change-0428/capture.py --binary /tmp/litchi-goal-0428-binaries/litchi-perf-baseline --revision <implementation-commit>
python3 -B docs/performance/results/change-0428/summarize.py
```

These drivers create outputs exclusively; reproduce in a fresh evidence
directory or checkout with generated receipts and captures removed. The
retained bundle can be checked without the original executable or repository:

```sh
python3 -B docs/performance/results/change-0428/verify.py --portable
```

`verify-report.py` independently checks each report, `probe-report.py` exercises
adversarial mutations, and `verify.py` binds source manifests, executable and
corpus identities, command receipts, all raw observations, summary derivation,
and the full file inventory. `seal.py` compresses logs losslessly with original
and stored hashes. Cleanup removes only the hash-bound copied executable and
preserves both existing target directories.

The recorded portable replays before and after cleanup passed all 16 formal reports and
272 formal mutation probes, plus changed-output and pinned-validator rejection
checks. `cleanup.json` records removal of the hash-bound temporary executable
and preservation of the original build executable and both target directories.

The global performance goal remains open. See `next-work.md` for the next
evidence gap; no latency, throughput, physical-copy, allocator, general leak,
or causal optimization claim follows from these observations.

# DOCX replayable paragraph stream evidence

This batch characterizes the implemented DOCX logical tail-stream lifecycle.
It retains 120 replay-route and 108 input/sink/compression formal processes,
each with 30 samples and three warmups: 6,840 measured operations in total.
The 114 pilot processes are excluded. All formal children and both capture
gates pass with unchanged, matching normal/allocator source manifests.

These are contemporaneous route/profile measurements, not a historical
optimization speedup. This batch changes measurement tooling and evidence;
production Rust is unchanged. The full non-iWork performance/CRUD goal remains
open, including missing workload intersections, broader before/after evidence,
cold-cache/atomic-save studies and the remaining optimization program.

## Results and scope

The [complete tables and scaling figure](route-analysis/formal1/measurements.md)
retain every workload/profile/repeat, normal latency tails, allocator operation
heap increments, requested allocations, and whole-process RSS. The
[machine-readable summary](route-summary.json) also contains throughput,
logical I/O counters and histograms, repeat spread, comparison flags, and
exact source/authored/candidate identities.

For 64 existing and 16,384 short authored paragraphs, normal p50 latency in
repeats 1/2 is 118.783/116.986 ms for deterministic replay,
90.297/90.194 ms for the memory store, and 96.711/97.078 ms for the file store.
Allocator operation peak increments are respectively 989,190, 9,184,134 and
795,759 bytes in both repeats. The memory-store configuration reserves its
explicit 8 MiB ceiling; reservation and allocation requests are not physical
resident-memory or copy measurements.

The file store has an adverse small-case result: 64 existing/64 short authored
paragraphs take 3.793/3.797 ms versus deterministic 0.796/0.867 ms. At 131,072
existing/64 authored paragraphs all three routes have p50 values near
475–479 ms. The source-varying short-text operation peak is 989,190 bytes for
deterministic replay at each selected source size, but normal process RSS
grows from about 6.2 MiB to 72 MiB because it includes fixture and other process
owners. This does not prove a whole-process constant-memory property.

Two process repeats provide descriptive uncertainty, not a strong confidence
interval. Normal and allocator timing are kept separate. All individual
comparisons remain visible; differences above 5% trigger review, including
slower profiles and repeat drift. No aggregate is used to hide adverse cases.

## Profiles and next investigation

The [24 external profile receipts](route-profiles/profiles1/result.json) all
pass and retain perf counters/samples, positional-I/O syscall traces, and
heaptrack data for source-heavy and authored-heavy cases across three replay
routes. These profiles cover whole children, including setup and profiler
overhead, separately from formal latency samples. The
[profile plan](route-profile-plan.md) records commands and scope. The
[validated profile summary](route-profiles/profiles1-summary.json) includes
all six exported perf stacks with no pending exports or validation errors;
[profile methods](profile-methods.md) explain inclusive, overlapping markers.
The VM reports zero generic/L1 cache counters and unsupported LLC events;
those values do not establish cache benefits.

Formal file-input measurements identify a further investigation: at 64
existing/16,384 authored paragraphs the first-repeat normal p50 is 443.859 ms,
versus 118.783 ms for owned input. Both measured operations make 60 logical
reads returning 7,651 bytes. The [six focused diagnostic profiles](input-metadata-profiles/metadata2/result.json)
all pass. The authored-heavy file-input child makes 3,735,939 `statx` calls,
versus 12 for owned input; a separate raw trace attributes 3,735,927 calls to
the exact prepared source descriptor. Source-heavy file input makes 15,415
`statx` calls versus 12 for owned input. These whole-child counts include one
warmup, one sample and setup, so they are not per-operation counters.
`FileSource::len` and `version` each query metadata; the measurements identify
metadata-call amplification as a concrete optimization target. They do not
yet isolate each caller or establish a safe optimization. Freshness and
mutation detection must remain intact.

The raw metadata traces are retained losslessly as gzip with a
[compression custody receipt](input-metadata-profiles/metadata2/compressed-custody.json).
Verification decompresses them and checks the original byte hashes. The
failed metadata1 attempt used an invalid `pread` syscall filter and is
retained; metadata2 uses `pread64`.

The [repeated-audit investigation](repeated-audit-investigation.md) maps six
decoded source passes and rules out directly deleting the standalone source
XML audit. A valid candidate can repair malformed source XML in the generic
OPC API. Any optimization must preserve independent source and candidate
validity, source-first errors, limits, freshness and publication guards.

## Reproduction and verification

[Formal execution](formal-execution.md) lists the build, source preparation,
pilot and capture commands. [Analysis methods](analysis-methods.md) describe
the exact inventory, frozen validators, corpus checks, metric scopes and
repeat calculations. The workload protocol, machine inventory, matched build
records, execution-input manifest and cleanup receipt are all retained.

The normal and allocator binaries are copied under the explicit cache paths
in `route-attempts/formal1/`. Keep those copies for verification, or reproduce
with fresh build/output attempts and a newly recorded machine/protocol;
commands refuse to overwrite existing evidence. The recorded machine and
filesystem capabilities are part of the scope. Source manifests cover the
workspace broadly, while the benchmark feature set excludes iWork.

The analyzer independently revalidates every report using the frozen capture
validators before extracting statistics. The final verifier checks artifact
integrity and recomputes the formal summary:

```sh
python3 -B docs/performance/results/change-0484/verify_route_seal.py verify
```

Generated corpus archives and failed historical diagnostics remain evidence.
Owned replay files and their 60 empty pilot/formal directories have been
removed with authenticated cleanup records. Profiler scratch is also removed.

# 0485: OPC splice consumed-window batching

The OPC splice auditor now batches hashing and sink emission for bytes already
consumed by the parser, using its existing bounded adapter window. Previously,
small XML parser fragments each triggered source metadata checks and sink
calls. The change-0484 authored-heavy file-input diagnostic observed
3,735,927 `statx` calls on the source descriptor despite only 60 logical
source reads in the corresponding formal operation.

Per-fragment `Resource::Work` charging remains unchanged. Source-first
freshness checks guard external source, replay, and sink callbacks; replay
EOF and exact source/candidate/replay digests remain authoritative. The
standalone source XML audit remains required. Oversized replay read counts
are rejected before indexing the buffer. No public API, source version
policy, dependency, buffer allocation, or ambient provider is added.

The [implementation and ADR review](../results/change-0485/implementation-review.md)
records the proof obligations. The [measurement methods](../results/change-0485/methods.md)
define three workloads and 18 route/input arms, two separately built
normal/allocator roles, two reversed process repeats and 30 measured
samples per process. Before executables are the retained change-0484 builds;
after executables are built from this change under the same compiler flags.

The [completed comparison](../results/change-0485/results-review.md) retains
144 formal processes and 4,320 samples, with identical candidate archive
bytes. Authored-heavy file-input p50 falls from 443.916 / 440.383 ms to
239.499 / 238.865 ms across two repeats (45.76–46.05% lower). Source-heavy
owned p50 falls from 482.325 / 476.215 ms to 385.058 / 385.006 ms.
Operation heap peaks remain effectively unchanged. The complete review keeps
three adverse latency quantiles, nine adverse whole-child RSS observations,
and the source-heavy metadata-call increase visible.

All-feature tests pass 1,989 cases (32 ignored); no-default-feature tests pass
1,948 (32 ignored). Formatting, Clippy/rustdoc with warnings denied, crate
boundaries, five benchmark tests, and nine Python tests pass. Twelve external
profiles and sanitizer fuzz checks (27 required-positive cases, two 10,000-run
campaigns) pass. The [bundle](../results/change-0485/README.md) retains raw
receipts, historical failures, cleanup evidence, and seal verification.

The full non-iWork performance/CRUD goal remains open. This batch does not
close native Office certification, cold-cache, atomic-save, concurrency, or
all workload-intersection requirements.

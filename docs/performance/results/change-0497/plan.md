# 0497: atomic publication for bounded DOCX logical-tail append

The existing replayable paragraph stream already bounds parser and replay
windows and retains source, authored, candidate, and inverse proofs. It has no
format-owned atomic destination method. Add consuming `write_to_path` methods
on the existing plan and commit products, using OPC's existing atomic helper
and the exact stream publication path. This is a required filesystem-publication
enabler, not a speculative speed optimization.

Preserve source identity/version and artifact checks, authored replay admission,
finite limits, cancellation, candidate authentication, durable patch and inverse
proof products, and typed errors. A callback failure must leave the destination
unchanged and remove its private sibling temporary file. Preserve the existing
`OpcError::Committed` distinction after replacement and failed directory sync.
Test destination aliases, permission preservation on supported platforms,
symlink/nonregular refusal, source/replay/cancellation/limit failures, exact
candidate bytes, reopening, patch/inverse replay, and resource release.

## Measurements being prepared

Use the original `66af25e2fc3c208e6819fffa5fa29e102942bc8e` harness as the
before control. Compare its default hashing sequential sink against the same
default route after the change. The new counting/non-retaining sequential sink
and atomic destination are after-only capability baselines; neither has an
invented before implementation or authorizes an atomic speedup claim. The
counting sink uses the production publication artifact fingerprint/length and
untimed fixture oracles. The atomic route additionally reads back its actual
destination outside timing. Preserve legacy default report compatibility.

The planned matrix reuses the existing 0489 small, authored-heavy, and
source-heavy workloads and its 18 explicit provider/workload arms: deterministic
owned/file/short-read/delayed input plus owned memory-store and file-store
authored input. Two reversed repeats in normal and allocator roles use three
warmups and 30 measured samples. There are 72 before/default children and
216 after/default/counting/atomic children, or 288 children and 8,640 samples.
The legacy comparison is balanced within each arm/role pair: before then
after in repeat one, followed by reversed pair order and after then before
in repeat two. A separate capability block similarly balances counting then
atomic against atomic then counting. This controls ordering without creating
an atomic before comparison. A separate 72-child pilot uses normal binaries,
one warmup and three samples across the same 18 arms and four version/route
coordinates; its 216 samples are excluded from formal analysis.
The detailed command/schema contract must be frozen before capture. Record
all source/build hashes, raw reports, terminal states, private-path cleanup,
whole-child RSS, lifecycle allocations, read/replay/output counters, uncertainty,
and every adverse repeat above the five-percent review threshold. Never fill
unobservable atomic write-call counts with synthetic measurements.

The atomic interval includes sibling-file creation, publication, synchronization,
replacement, and the existing parent-directory synchronization policy. Source
file preparation, output readback/hash/semantic/media checks, and measurement
report writing remain explicitly outside timing. Source-profile file inputs
are prepared positional descriptors, not verified-cold filesystem lifecycles.
The legacy hashing sink keeps its original timing scope. Separate syscall traces
will identify destination synchronization and replacement activity without
counting profiler overhead as benchmark latency.

## Validation and retention

Use one serial Cargo lane for the isolated before and after clones. Run the
focused route tests, relevant DOCX default and feature gates, full harness
library tests, warning-denied Clippy/rustdoc, formatting, boundary checks,
existing applicable fuzz/preservation tests, and Python evidence-validator tests.
Preserve unsuccessful development attempts. Retain measured executables and
reproducible source/build/capture evidence; remove owned source clones, Cargo
targets, output/replay files, and scratch once terminal custody is verified.

The full non-iWork goal remains open. This batch does not establish native
producer round trips, genuine borrowed-source lifetimes, cold intersections,
parallel scaling, arbitrary append semantics, or broad CRUD/security completion.
Preserve unrelated primary changes. Do not access `~/code/litchi-spec-gaps`.

The full DOCX gate reproduced an existing section event-limit test failure
on the original before checkout. Its 16 MiB fixture budget refused the scan's
roughly 192 MB reservation before reaching the event limit. A test-only sixth
overlay file retains that typed memory refusal and admits the event case under
a finite 256 MiB budget. Production scanner limits and reservations are unchanged.
The tail-stream tests also explicitly fail an authored replay handle after
preparation and verify atomic destination preservation and temporary cleanup.

# 0823 real-file probe static review

This review covers the packet-local source under `real-probe-src/`. It was
prepared without running Cargo, rustfmt, a build, or the probe. No reader
invocation was attempted, so there are no failed reader logs to retain.

## Scope and custody

The probe is copied from the committed 0822 real-file probe. The pinned
package paths, input and reference sizes and SHA-256 identities, marker,
target `(slide 0, shape 0)`, direct and wrapped edit helpers, output oracle,
reopen checks, semantic text digest, slide count check, per-sample checks, and
the Snapshot drop point are retained. The package and binary names remain
`pptx-edit-profile-0822`; the lockfile is copied unchanged. The trial binds
its source baseline to `b76786208d`, the current committed 0822 profile head.
The report schema is `litchi.performance.0823.pptx-edit-trial.v1` because this
is a new trial packet and the report now describes optional allocator evidence.

The only new source files are the packet-local probe manifest and the three
probe modules. No production crate or workspace performance harness file is
changed by this packet-local addition.

## Feature and measurement review

The `allocator-metrics` feature is opt-in and is absent from the default
feature set. `allocation_metrics` is compiled in both modes so the native
path can call the same region API, but `begin()` returns a disabled region
unless the executable enables the feature. The global allocator wrapper is
compiled only under `#[cfg(feature = "allocator-metrics")]` and is therefore
not installed by the native binary.

The allocator callback API and callback guard are feature-gated. The split
region helpers and raw `Snapshot`/counter test surface are gated by both
`test` and `allocator-metrics`, because the real probe uses one unsplit region
and those APIs exist only for allocator-module tests. This keeps native and
feature non-test builds free of unused instrumentation warnings while
retaining the full observer and allocator tests in the feature test build; it
does not add blanket dead-code allowances to the native path. The existing
documented allowance on the split-only `segment_before` field is narrowly
scoped with `cfg_attr(not(all(test, feature = "allocator-metrics")),
allow(dead_code))`, so it covers only the field's intentionally uncompiled
test-only readers.

Each measured sample creates the allocation region after `Package::open` and
immediately before `Instant::now()`. It invokes the existing direct or
wrapped edit body unchanged, reads the elapsed duration, and calls
`Region::finish()` after that elapsed read. The observer boundary snapshots
are therefore outside the elapsed clock. Serialization, output hashing,
package destruction, and semantic readback remain outside the clock as in the
0822 probe. A feature report includes one allocation sample per measured
sample; a native report omits the `allocation` member from each sample. The
report-level allocator identity remains as provenance and reports
`instrumentation: "none"` for native runs.

The `edit_helper_0822` body still binds the commit result, applies it, and
discards the returned `Snapshot` at the existing `apply_opened_presentation_commit`
semicolon. The region guard is not held by the commit result and cannot move
that drop point. The timed call boundary is consequently the same operation
for elapsed and allocation evidence.

The local allocator wrapper is based on the current
`tools/perf-baseline` implementation. Its tests include the 0820 repair:
process-wide `live_bytes` is checked through allocation/deallocation
conservation and monotonic high-water properties rather than assuming that a
test owns every process callback. The wrapper delegates all allocation
ownership to `std::alloc::System`; the observer only records post-system-call
metrics. The wrapper uses `cfg_attr(not(test), global_allocator)`: direct
`GlobalAlloc` tests exercise the wrapper without installing it as the test
harness allocator, while the actual allocator observer executable still has
the global wrapper. The wrapper tests and the real probe's direct/wrapped
parity test share the allocation observer `TEST_LOCK`; synthetic counter
tests therefore have no ambient test-harness allocator callbacks, and the
real parity test cannot overlap those observer mutations.

## Suggested test and evidence gate

Root should run the following as separate, retained command records:

1. `cargo fmt --manifest-path docs/performance/results/change-0823/real-probe-src/Cargo.toml -- --check`
2. `cargo check --manifest-path docs/performance/results/change-0823/real-probe-src/Cargo.toml`
3. `cargo check --manifest-path docs/performance/results/change-0823/real-probe-src/Cargo.toml --features allocator-metrics`
4. `cargo test --manifest-path docs/performance/results/change-0823/real-probe-src/Cargo.toml`
5. `cargo test --manifest-path docs/performance/results/change-0823/real-probe-src/Cargo.toml --features allocator-metrics`
6. `cargo clippy --manifest-path docs/performance/results/change-0823/real-probe-src/Cargo.toml --all-targets --all-features -- -D warnings`

After those gates, the probe driver should build and run both direct and
wrapped modes in the native and allocator configurations with the exact 0822
input/reference arguments. The native JSON must have the retained file
identities and verification fields, must omit per-sample `allocation`, and
must report allocator instrumentation `none`. The allocator JSON must retain
the same output bytes and all verification booleans, report
`system_allocator_operation_scoped`, and contain measured or explicitly
fail-closed allocation samples with checked status fields. The driver should
reject any output whose pinned input, reference, output bytes, reopened text,
target marker, full-text digest, or slide count differs from the existing
oracle.

The reviewer suggestion is to admit both direct and wrapped modes as valid
probe modes while using real/direct for the warm qualification lane. The
existing real-probe unit test exercises direct and wrapped output and semantic
parity under the shared observer lock, so this recommendation does not add a
new gate or change the timed workflow.

Allocation results are resource evidence only. They must not be combined with
native timing samples as if they were the same binary or treated as proof of
a latency improvement; the global wrapper can perturb allocator scheduling.

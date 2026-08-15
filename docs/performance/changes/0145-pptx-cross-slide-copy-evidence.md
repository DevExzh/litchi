# Change 0145: PPTX cross-presentation slide-copy evidence

Date: 2026-08-16

Status: validated: focused selector test, two-case release smoke, strict
deprecation check, and all-target harness Clippy pass.

## Scope

The performance harness now exposes two opt-in PPTX selectors:

- `pptx_cross_copy_plain`
- `pptx_cross_copy_media_rich`

Both selectors use the public opened-PPTX cross-presentation APIs. The fixed
source package has three slides and the fixed destination package has two
slides. The selected source slide is slide 3, the selected destination slide is
slide 2, and the copy is inserted at zero-based position 1. The media-rich corpus adds
eight deterministic 2 MiB PNG resources to both packages, forcing the copied
dependency closure to exercise deterministic collision-avoidance remapping. The plain corpus keeps the
same slide graph without media.

The harness measures these phases independently:

1. `Snapshot::plan_cross_slide_copy` planning;
2. `Package::apply_cross_slide_copy_plan` commit/publication into the in-memory
   destination; and
3. `OpcPackage::to_stream` sequential publication into a bounded forward-only
   sink.

Reopen timing is retained as a separate diagnostic. Corpus construction,
reopen, semantic readback, package and presentation relationship topology,
slide relationship target/count, `[Content_Types].xml`, dependency-closure Part
identity/target/content-type/payload checks (including exact slide XML and
media-owner elements), untouched destination Part/member bytes,
collision-remap, source-immutability, output-hash, durable-patch,
stale-source, stale-destination, borrowed-provenance, and foreign-source
refusal gates are outside the measured phases. Sink counters and per-sample
output hashes are emitted with the phase vectors.

## Evidence boundary

The selectors are correctness/evidence controls only. They do not claim a
speedup, release ABBA result, allocation/RSS result, cold-cache result,
physical-I/O result, or production optimization. No CFB, OPC, or iWork
production code is changed by this tranche. The selectors do not enter the
default matrix: the default remains 36 cases and 198 records.

The implementation is in the [performance harness](../../../tools/perf-baseline/src/main.rs)
and its [operator README](../../../tools/perf-baseline/README.md). The
production capability under test is the bounded [PPTX cross-copy planner](../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs);
this change does not modify that production crate.

Example:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 \
  --case pptx_cross_copy_plain,pptx_cross_copy_media_rich \
  --json target/perf/pptx-cross-slide-copy.json
```

## Verification commands

The focused checks are:

```sh
cargo fmt --manifest-path tools/perf-baseline/Cargo.toml -- --check
git diff --check -- tools/perf-baseline/src/main.rs \
  tools/perf-baseline/README.md \
  docs/performance/BASELINE.md \
  docs/performance/CRUD_COVERAGE.md \
  docs/performance/ADR_COMPLIANCE.md \
  docs/performance/changes/0145-pptx-cross-slide-copy-evidence.md
cargo test --locked --manifest-path tools/perf-baseline/Cargo.toml \
  pptx_cross_copy_selectors_emit_separated_phase_evidence_and_gates -- --nocapture
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 0 --samples 1 \
  --case pptx_cross_copy_plain,pptx_cross_copy_media_rich --json -
RUSTFLAGS="-D deprecated" cargo check --locked \
  --manifest-path tools/perf-baseline/Cargo.toml
cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-targets -- -D warnings
```

The focused test and smoke must show two deterministic selectors, separate
plan/commit/publication/reopen vectors, bounded sink writes, exact hashes, and
all computed semantic/topology/closure/refusal gates. These checks are
correctness and counter evidence only; they do not establish a speedup,
allocation/RSS, release-ABBA, cold-cache, physical-I/O, or production
optimization result.

# Probe provenance and pre-capture repairs

The probe starts from 0782's public PPT fixtures, authoring and semantic/raw
OfficeArt text oracle. `main.rs` adds the shared non-inlined `profile_write`
owner and output retention only for the final sample when `--artifact` is
requested. Normal native/profile loops retain the historical per-iteration
output destruction order. Qualification emits the final artifact after every
sample's verification, using exclusive file creation. Source/output identities,
shape dimensions and text oracles must agree with sealed historical fixtures.

The named wrapper includes the public write plus two `black_box` barriers. Its
barrier overhead is shared by both legs and is not measured against a wrapper-
free executable. Its call/return must survive optimization. The new current-
source timings therefore diagnose these binaries; they do not pool with 0782.

The shared allocation module is inherited except for its report labels. It is
compiled in disabled mode in every native and Callgrind workload; no allocator
metric is collected in this packet. The optional counter wrapper is only
built/tested by probe quality gates. No production unsafe boundary changes.

Before any release build or workload capture:

- The first standalone lock generation selected newer cached registry versions.
  Its lock is retained as `probe-initial-resolution.lock`. A fresh copy of the
  root lock was pruned by Cargo metadata instead, and every resolved registry
  version/checksum was checked against the root lock in `lock-alignment.json`.
- The first probe Clippy gate refused unused shared counter APIs. A module-local
  dead-code allowance documents that this probe uses a subset of those APIs.
  There is no runtime algorithm change; initial source and receipt remain.
- The next all-feature probe test exposed the inherited process-global allocator
  observing test-runner allocations inside synthetic counter assertions. The
  wrapper now installs itself globally only outside `cfg(test)`, matching the
  already audited isolation in 0840. Direct `GlobalAlloc` delegation tests still
  call the same wrapper; synthetic counter tests run without test-runner events.
  The failed tests and v2 sources remain retained. The final default/all-feature
  tests are both required to pass before the release executable is admitted.

All twelve current production files touched by the archived 0782 patch are
byte-identical to its historical before files (`candidate-source-parity.json`).
Other current code and dependency changes remain part of both new legs. The
candidate must be restored exactly after the diagnostic, irrespective of timing.

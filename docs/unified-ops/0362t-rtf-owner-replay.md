# 0362t — RTF owner replay commits

This slice closes the three RTF facade owner replay methods that still
published a bare `Package`.  It is based on `b96137acb4` on
`uo2/w8-rtf-replay`; the coordinator must integrate it with the later shared
replay scope and tail-diagnostic fixes rather than treating those separate
changes as part of this record.

## API and ownership

* `litchi_rtf::Package::apply(&edit::Patch)` now returns the RTF
  `edit::Commit`.  The commit owns the published package, the reversible RTF
  patch, and the existing operation-count/changed diagnostics.
* `litchi_rtf::Package::apply_durable<Mode>(&core::patch::Patch<Mode>)` now
  returns `litchi_core::edit::Commit<Package, core::patch::Patch<Mode>>`.
  Durable forward and inverse replay continue to authenticate the source,
  enforce the recipient's options, and return the same bounded package.  The
  shared diagnostics report whether the exact artifact changed.
* `litchi_rtf::Package::apply_external_reference_redaction` now returns the
  RTF redaction commit.  Its patch remains `ForwardOnly`; no reversible patch
  or forbidden pre-redaction payload is constructed.  The existing redaction
  diagnostics remain available through the commit.

All RTF tests and the perf-baseline RTF consumers were migrated to consume
`Commit::into_package()` where an owner is needed.  Tests also inspect commit
shape/diagnostics on durable and redaction paths.  Source checks, stale/foreign
refusals, inverse replay, exact no-op behavior, and recipient policy remain
covered by the existing tests.  The separate coordinator repair for the RTF
tail inverse `inserted_bytes` diagnostic is intentionally not duplicated here.

## Validation

The following commands were run with the capped W8 environment and returned
zero unless stated otherwise:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/w8-file/env.sh
CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2 cargo check -p litchi-rtf --all-targets
CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2 cargo test -p litchi-rtf --test canonical_primary --test paragraph_split_merge --test external_reference_redact
CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2 cargo test -p litchi-rtf --all-targets
cargo clippy -p litchi-rtf --all-targets -- -D warnings
cargo fmt --all -- --check

cd tools/perf-baseline
CARGO_BUILD_JOBS=2 CARGO_TEST_JOBS=2 cargo check --all-targets
cargo clippy --all-targets -- -D warnings
```

The full RTF test receipt is kept outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-rtf-replay/rtf-tests.log`
(159,639 bytes, SHA-256
`b3db17e0ccaa365129ce7df36cc64ba456de8995c711d0db264619b77385f36d`).  It
contains 190 test suites with 1,534 passed and zero failed or ignored tests.
The focused command covered 26 tests (3 suites), all passed.  The perf-baseline
all-target check passed after its RTF consumers were migrated.

The RTF crate's unfiltered all-target Clippy command passed with `-D
warnings`.  Perf-baseline's unfiltered command retains four pre-existing
`clippy::while_let_on_iterator` warnings in `src/lib.rs` (lines 22714,
22747, 29390, and 31277) and one pre-existing
`clippy::needless_borrow` in `examples/ppt_elements_probe.rs:45`; the same
command with those existing diagnostics allowed passed.  No lint suppression
was added by this slice.  No performance timing or binary is claimed here;
protected measurement artifacts remain outside the repository.

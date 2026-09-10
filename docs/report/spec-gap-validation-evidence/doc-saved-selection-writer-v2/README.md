# DOC saved-selection validation v2

This evidence draft records the refreshed isolated candidate at
`/var/tmp/litchi-doc-saved-selection-publication-validation-20260910`, based on
`ce0be842b60a28db142b795389e24be5bde8ae07`. The seven owned source paths and
their SHA-256 values are bound by `source-manifest.json` and `receipt.json`.
The v1 evidence directory remains historical and is not reused as v2 gate
evidence.

The strict DOC owner validates active Selsf main-story CPs before a table
clone. Length-changing main-story edits remap `cpFirst`, `cpLim`, `cpAnchor`,
and active text-block `cpAnchorShrink` when each value has a deterministic
splice-boundary mapping. A CP strictly inside replaced text, or an insertion
at an existing CP, returns the typed `PositionDependency` refusal before
candidate publication. Exact no-ops and edits wholly after the selection keep
the original Selsf bytes. Durable text replay uses the same body operation and
therefore performs the same remap; explicit Selsf byte patches remain outside
durable and semantic three-way merging.

Semantic no-ops preserve ignored `blktblSel` and shape `fInsEnd` bytes. A
transition to a partial `TableSel` normalizes invalid previously ignored table
edges before parsing, and table-edge/anchor-shrink setters enforce their
applicable flags without setter-order dependence.

## Gates

With Rust 1.95.0 and dedicated targets, the refreshed candidate passed:

- `saved_selection`: 8 passed;
- `parts::saved_selection`: 9 passed, 1,018 filtered;
- `auxiliary_tables`: 2 passed;
- `doc_body_transaction`: 6 passed, including the three fixtures that
  previously failed under the blanket-refusal implementation;
- `--all-features --all-targets`: 1,044 library tests ran (1,042 passed and 2
  ignored); 1,232 tests passed total, 0 failed, 2 ignored across 45 test
  binaries;
- warning-denied all-feature/all-target Clippy;
- warning-denied all-feature rustdoc;
- pinned rustfmt for the six owned Rust files and `git diff --check`.

Independent `calc_chain_review` remains pending. This artifact makes no review
clearance claim.

## Scope and limits

The batch covers the Selsf parser/transaction, strict body and tracked-revision
publication seams, public exports, feature-matrix row, and focused tests. It
does not apply selection state to document navigation or host UI, render
selection geometry, relocate FIB ranges, insert/delete Selsf records, or claim
layout behavior. Changed publication continues to use the existing
signed-source refusal policy; exact no-ops remain allowed.

## Reproduction

Restore the captured lockfile and use fresh dedicated targets. The logged
replay used CPU affinity 8–31 while the root gate reserved the other CPUs.

```sh
gzip -dc docs/report/spec-gap-validation-evidence/doc-saved-selection-writer-v2/Cargo.lock.gz > Cargo.lock
export RUSTUP_HOME=/tmp/litchi-spec-gap-rustup
export TMPDIR=/var/tmp/litchi-spec-gap-gate-tmp

CARGO_TARGET_DIR=/var/tmp/litchi-doc-saved-selection-v2-focused-target-195 \
  taskset -c 8-31 cargo +1.95.0 test -p litchi-doc --test saved_selection --offline --locked
CARGO_TARGET_DIR=/var/tmp/litchi-doc-saved-selection-v2-lib-target-195 \
  taskset -c 8-31 cargo +1.95.0 test -p litchi-doc --lib parts::saved_selection --offline --locked
CARGO_TARGET_DIR=/var/tmp/litchi-doc-saved-selection-v2-focused-target-195 \
  taskset -c 8-31 cargo +1.95.0 test -p litchi-doc --test auxiliary_tables --offline --locked
CARGO_TARGET_DIR=/var/tmp/litchi-doc-saved-selection-v2-focused-target-195 \
  taskset -c 8-31 cargo +1.95.0 test -p litchi-doc --test doc_body_transaction --offline --locked
CARGO_TARGET_DIR=/var/tmp/litchi-doc-saved-selection-v2-full-target-195 \
  taskset -c 8-31 cargo +1.95.0 test -p litchi-doc --all-features --all-targets --offline --locked
CARGO_TARGET_DIR=/var/tmp/litchi-doc-saved-selection-v2-clippy-target-195 \
  taskset -c 8-31 cargo +1.95.0 clippy -p litchi-doc --all-features --all-targets --offline --locked -- -D warnings
CARGO_TARGET_DIR=/var/tmp/litchi-doc-saved-selection-v2-doc-target-195 \
  RUSTDOCFLAGS='-D warnings' taskset -c 8-31 cargo +1.95.0 doc --offline --locked --all-features -p litchi-doc --no-deps

rustfmt +1.95.0 --check --edition 2024 \
  crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/lib.rs \
  crates/litchi-doc/src/parts/saved_selection.rs \
  crates/litchi-doc/src/tracked_revision/package.rs \
  crates/litchi-doc/tests/auxiliary_tables.rs \
  crates/litchi-doc/tests/saved_selection.rs
git diff --check
```

The compressed logs, lockfile, and their hashes are recorded in `receipt.json`.

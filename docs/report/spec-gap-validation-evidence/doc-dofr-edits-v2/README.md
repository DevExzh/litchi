# DOC DOFR source-bound publication validation v2

This validation evidence covers the six-path DOC `RgDofr` publication
overlay in `/var/tmp/litchi-doc-dofr-publication-validation-20260909`, based
on commit `463ab2a5989bba49ffd95277dc29422a363de7f8`. The source overlay and
its exact hashes are defined by `source-manifest.json`; publication state is
handled separately.

The feature exposes bounded, inert `RgDofr` reads and source-checked,
same-length complete-record replacements through the DOC body facade. It
retains the complete source owner for facade publication, refuses changed
publication when opaque signature metadata could become stale, and refuses
semantic three-way plans containing auxiliary-table changes.

## Scope and boundaries

The six source hashes are bound in `source-manifest.json` and `receipt.json`.
The four implementation/test paths are copied from the frozen DOFR feature
overlay; `lib.rs` and `FEATURE_MATRIX.md` contain only the public export and
matrix context hunks described by that manifest. Unrelated CFB, saved-selection,
VBA, and other changes are excluded.

The publication limits are deliberate:

- **Complete-source binding:** `DofrPatch` first checks the FIB-selected table
  range and exact `RgDofr` bytes. The high-level `body_text::Edit` then requires
  the patch owner to match the complete immutable DOC source by shared-allocation
  identity or exact complete-source bytes. Thus a patch made from a different
  DOC body is rejected even when its `RgDofr` range has the same length. An
  independently reopened byte-identical DOC may authorize the patch. Detached
  low-level `DofrArray`/`DofrPatch` operations intentionally remain component
  scoped and do not claim complete-document ownership.
- **Same-length records:** replacement preflight checks the record index and
  exact serialized length before retaining a transaction draft. The candidate
  array is reparsed before the draft is replaced. Record insertion/deletion,
  `RgDofr` resizing, and FIB range relocation are outside this batch. Exact
  no-ops return without charging an operation or changing source bytes.
- **Signed-source refusal:** changed publication is refused when PIDDSI has a
  `DigitalSignature` property, when selected `StwUser` contains a valid
  `Sign`, `SigAgile`, or `SigV3` Word signature, or when selected `StwUser`
  metadata is malformed and cannot be safely classified. PIDDSI directory-name
  matching is CFB case-insensitive. Direct and tracked lower editors apply the
  same changed-source guard. Exact byte no-ops remain allowed. This is a
  conservative stale-metadata policy, not cryptographic verification or trust
  evaluation; signature payloads remain opaque.
- **Auxiliary three-way refusal:** `Snapshot::plan_three_way` requires both
  patches to target the exact source and then refuses any patch carrying the
  auxiliary `RgDofr` change marker. This covers DOFR-versus-body and
  DOFR-versus-DOFR histories. The marker is intentionally conservative: the
  planner does not infer merge safety from a history that is otherwise
  byte-equivalent or from a net-no-op auxiliary operation; opaque auxiliary
  bytes are never silently dropped. Exact no-op replacement itself produces no
  auxiliary change marker.

The batch does not render frames, apply list styles, resolve external frame
files, write `StwUser`, or provide arbitrary-length `RgDofr` editing. Unknown
record payloads remain inert and source-preserved under the existing bounded
reader policy.

## Gate results

Root's dedicated Rust 1.95.0 gate passed 1,223 unit/integration tests and 14
doctests, with 2 unit/integration and 12 doctest ignores. Strict all-target,
all-feature warning-denied Clippy, warning-denied rustdoc, pinned rustfmt, and
`git diff --check` passed. The authoritative raw log hashes and exit codes are
recorded in `receipt.json`; the compressed logs are included here. No native
producer acceptance, frame rendering, or cryptographic validity claim is made.

## Reproduction

Restore the captured ignored lockfile and use a fresh dedicated target:

```sh
gzip -dc docs/report/spec-gap-validation-evidence/doc-dofr-edits-v2/Cargo.lock.gz > Cargo.lock

export RUSTUP_HOME=/tmp/litchi-spec-gap-rustup
export CARGO_TARGET_DIR=/var/tmp/litchi-doc-dofr-publication-root-target-195
export CARGO_BUILD_JOBS=4
export TMPDIR=/var/tmp/litchi-spec-gap-gate-tmp

cargo +1.95.0 test --offline --locked --all-features -p litchi-doc
cargo +1.95.0 clippy --offline --locked --all-features --all-targets -p litchi-doc -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo +1.95.0 doc --offline --locked --all-features -p litchi-doc --no-deps

rustup run 1.95.0 rustfmt --check --edition 2024 --config skip_children=true \
  crates/litchi-doc/src/body_text.rs crates/litchi-doc/src/lib.rs crates/litchi-doc/src/parts/dofr.rs crates/litchi-doc/src/tracked_revision/package.rs crates/litchi-doc/tests/dofr_edits.rs
git diff --check
```

`tests.log.gz`, `clippy.log.gz`, and `rustdoc.log.gz` preserve the root gate
outputs. The receipt records the commands actually executed by root and the
separate offline/locked replay commands shown above. `Cargo.lock.gz` is the
exact lockfile used for the reproducible offline command. `receipt.json` binds
each archive's compressed and raw SHA-256, the six source hashes, and the root
gate receipt provenance.

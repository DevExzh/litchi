# 0362ugb — ODB and ODM lazy edit entry

This slice closes the confirmed edit-construction cost debt for the ODB
authoring transaction and the ODM title family.  It starts at
`7559f44aa975254bd96f6e28eee17366b0ebeb6a`; no public signature or consumer
route changes.

## Entry-state design

`litchi_odb::Package::edit` and `edit_with_policy` now construct
`authoring::Edit` with `Cow::Borrowed(source.database.content_xml())`.  All
existing mutating splices use `Cow::to_mut()` at the write boundary.  The
constructor, read-only staging and discard remain borrowed.  Some helper
call sites take the mutable buffer before proving a move is needed or before
returning a validation failure (for example, a self-move through
`relocate_sibling`), so those paths may materialize; this does not change
their semantic no-op result.  The commit path still compares the staged XML
with the source, rebuilds and validates only a changed candidate, and
preserves the existing policy refusal, patch, inverse and exact-source
behavior.

`litchi_odm::Package::edit_title` no longer reads `meta.xml` at entry.  Its
bounded title state borrows the title already retained by the immutable
metadata projection (`before` and the initial `after` are `Cow::Borrowed`).
`set` changes the staged value to `Cow::Owned`; only a changed commit then
materializes and parses the raw `meta.xml`, applies the bounded patch and
performs the existing readback validation.  Exact no-op commits continue to
return the source artifact without touching the raw part.

No lazy failure is represented by a panic or an exposed debug payload.  The
existing fallible metadata, XML, limit and publication errors remain on their
original commit paths.

## Lazy-state proofs

The ODB advanced transaction test asserts that the staged XML pointer is the
source XML pointer immediately after `edit()`, then is a distinct allocation
after the first successful query mutation.  The ODM semantic reference test
asserts the same borrowed-to-owned transition for a projected title.  The
existing ODM exact-source no-op test remains in the same suite and proves the
deferred raw-part route returns byte-identical source bytes.

## Validation

With the capped `odb-odm-lazy-entry` environment (one Cargo job, offline,
locked, debug information disabled, incremental disabled and external
`TMPDIR`), these checks passed:

```text
cargo fmt --package litchi-odb --package litchi-odm -- --check
cargo check --offline --locked -p litchi-odb -p litchi-odm --all-targets --all-features --message-format=short
cargo test --offline --locked -p litchi-odb --all-features --test advanced_transactions
cargo test --offline --locked -p litchi-odm --all-features --test semantic_references
cargo clippy --offline --locked -p litchi-odb -p litchi-odm --all-targets --all-features -- -D warnings
```

The target directory was 422,254,400 bytes after the focused checks and was
kept outside the worktree.  No native timing or baseline qualification is
claimed by this record.

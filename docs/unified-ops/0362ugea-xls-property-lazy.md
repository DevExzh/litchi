# 0362ugea — XLS property and worksheet-feature lazy entry

This slice closes the remaining XLS edit-entry debt from 0362uge. The
package metadata, OLE property-set, and worksheet-feature families now keep a
cheap clone of the immutable package owner at `edit_*()` entry. Their source
bytes, property-set parser/editor, typed metadata value, Workbook stream, and
BIFF globals are created only by the first accessor, verb, or commit.

The lazy shells retain the package's admitted policy. Property-set materialize
uses the CFB profile returned by `Package::edit_profile`, including the
source-backed CFB ceilings. A failed first materialization installs no partial
state, so a retry uses the same owner and policy. Metadata accessors are
fallible (`properties()` and `properties_mut()`), so malformed property
streams cannot be hidden behind an infallible reference. The XLS unified
metadata adapter, examples, and tests pass the new accessor result through.

Each family keeps its existing publication behavior: encrypted input is
refused before an edit is created; a source mutation after first materialize
is checked before property or worksheet-feature publication; exact no-ops keep
the source package and identity patch; changed unsigned commits republish
through the retained eager or source-backed policy; and dropping a cold edit
performs no source capture. Signed inputs can still be read and committed as
exact no-ops after a projection has been resolved, including metadata and
worksheet features, while a changed property, metadata, or feature commit is
refused before any publication. Signed no-op patches replay byte-for-byte;
explicit signature removal leaves the returned unsigned package editable. The
worksheet-feature edit uses private lazy stream state, so all feature,
drawing, note, hyperlink, validation, format and filter verbs materialize
through one gate.

## Scope and consumers

- `crates/litchi-xls/src/package/metadata.rs`: deferred typed properties and
  fallible property accessors and signed no-op/change policy.
- `crates/litchi-xls/src/package.rs`: crate-private signed-state projection
  used by the scoped publication guards.
- `crates/litchi-xls/src/workbook/package/property_set.rs`: deferred CFB
  snapshot/editor, retained CFB limits, source freshness check, and signed
  no-op/change policy.
- `crates/litchi-xls/src/sheet_features/edit.rs` and `objects.rs`: deferred
  Workbook stream/BIFF globals, all feature-object consumers, and typed
  signed-write refusal.
- `crates/litchi-ops/src/format/xls/metadata.rs` and
  `crates/litchi-ops/examples/ole_metadata_bench.rs`: accessor-result
  migration.
- `crates/litchi-xls/tests/xls_metadata.rs`: accessor-result migration,
  signed read/no-op/change/replay checks, and existing exact no-op, patch,
  inverse, and fidelity checks.
- `crates/litchi-xls/tests/sheet_features.rs`: signed read/no-op/change and
  replay checks.

Private unit checks assert that all three package edit constructors retain no
source payload at entry. Focused signed tests passed for metadata/property-set
and worksheet-feature read, no-op commit, changed-write refusal, and no-op
replay. Existing XLS metadata and worksheet-feature suites continue to cover
changed output, exact no-op, replay and inverse behavior; the coordinator must
run them after assigning the capped Cargo slot.

The source fix was validated in the isolated capped target after the initial
source-only commit. The generated workspace lock is intentionally ignored and
kept outside Git at `/home/zhuhe/code/litchi-worktrees/uo2-xls-property-lazy/Cargo.lock`
(SHA-256 `6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`).
The signed metadata/property-set test passed 1/1 and the signed
worksheet-feature test passed 1/1. Their raw logs remain outside the checkout:
`signed-metadata-4f08-v3.log` (under `.coord/xls-property-lazy`) and
`signed-sheet-4f08-v2.log` (same directory). The broader focused XLS suite,
strict Clippy, and affected `litchi-ops` checks remain queued. No native
qualification, benchmark, or performance claim is made here; no binary or
large evidence artifact belongs in this record.

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
the source package and identity patch; changed commits republish through the
retained eager or source-backed policy; and dropping a cold edit performs no
source capture. The worksheet-feature edit uses private lazy stream state, so
all feature, drawing, note, hyperlink, validation, format and filter verbs
materialize through one gate.

## Scope and consumers

- `crates/litchi-xls/src/package/metadata.rs`: deferred typed properties and
  fallible property accessors.
- `crates/litchi-xls/src/workbook/package/property_set.rs`: deferred CFB
  snapshot/editor, retained CFB limits, and source freshness check.
- `crates/litchi-xls/src/sheet_features/edit.rs` and `objects.rs`: deferred
  Workbook stream/BIFF globals and all feature-object consumers.
- `crates/litchi-ops/src/format/xls/metadata.rs` and
  `crates/litchi-ops/examples/ole_metadata_bench.rs`: accessor-result
  migration.
- `crates/litchi-xls/tests/xls_metadata.rs`: accessor-result migration and
  existing exact no-op, patch, inverse, and fidelity checks.

Private unit checks assert that all three package edit constructors retain no
source payload at entry. Existing XLS metadata and worksheet-feature suites
continue to cover changed output, exact no-op, replay and inverse behavior;
the coordinator must run them after assigning the capped Cargo slot.

No Cargo build, benchmark, native qualification, or performance claim was
made in this source-only slice. The intended validation is `cargo fmt --all
--check`, the focused XLS metadata/property-set/sheet-feature tests, the full
XLS all-target/all-feature suite, strict XLS Clippy, and the affected
`litchi-ops` XLS consumer checks under the coordinator's isolated capped
target. No binary or large evidence artifact belongs in this record.

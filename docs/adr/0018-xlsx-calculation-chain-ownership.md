# ADR 0018: Typed XLSX calculation-chain ownership

- Status: Accepted
- Date: 2026-08-03

## Context

The migration host owned SpreadsheetML calculation-chain parsing, package
discovery, and mutation even though the capability is concrete XLSX grammar.
Its public model exposed long format-prefixed names and several independent
booleans for attributes whose combinations are constrained by the Office
profile. Keeping that implementation in `litchi-ooxml` also made the target
crate split less truthful and left the concrete owner without a complete
read/write/remove boundary.

Checked-in `[MS-OI29500]` section 2.1.688 requires the first cell to carry a
sheet identifier, limits that identifier to `1..=65534`, lets later cells
inherit it, and makes the `l` and `s` attributes mutually exclusive. Formula
text is not present in this part, and reading a chain must never evaluate a
workbook.

## Decision

`litchi-xlsx::chain` is the canonical owner of calculation-chain grammar,
semantic state, and OPC topology. Its contextual public vocabulary is
`Conformance`, `Sheet`, `Step`, `Flags`, `Cell`, and `Chain`, with preserved
non-schema attributes under `chain::raw::Attr`. The migration-host module and
all long forwarding names are removed.

`Sheet` proves the native `1..=65534` domain. `Cell::new` accepts the shared
checked `At` selector and stores its resolved grid address. `Step::{Same,
Level, Child}` makes the `l`/`s` exclusion structural, while the independent
thread and array markers occupy one `Flags: u8`. `Chain` is nonempty by
construction, keeps the first and every changed sheet boundary explicit on
the wire, and exposes semantic `get`, `put`, and `remove` as the primary CRUD
operations. Checked order operations remain available through `at`, `insert`,
`replace_at`, `remove_at`, and `move_at`; none indexes or unwinds on invalid
input.

Malformed producer input with duplicate semantic sheet/address keys remains
available in source order for inspection and repair. Semantic lookup or
mutation rejects that state as ambiguous instead of choosing one occurrence.
The reader accepts Strict and Transitional namespaces, applies the common MCE
processor, preserves bounded extension markup and qualified attributes, and
uses strict UTF-8 for accepted qualified names. Resource, nesting, attribute,
cell-count, and output limits are checked before publication or serialization.

The physical package verbs are the short `load`, `put`, and `remove`. They
require exactly one internal workbook relationship and a coherent part set,
reject external, duplicate, orphaned, wrong-content-type, or relationship-
bearing chain parts, and never infer dependencies from worksheet formulas.
An exact `put` is a no-op and retains signatures. A changed store updates the
part and relationship conformance together; creation preflights names and
rolls back the defensive part insertion if the workbook owner cannot be
reacquired. Removal retains a target that another package relationship still
references. Only an actual commit invalidates signatures.

The temporary workbook host caches `Chain` and `Conformance`, exposes only
`chain`, `chain_conformance`, consuming `put_chain`, and `remove_chain`, and
restores the concrete owner's part after its legacy materializer rebuilds
workbook relationships. That adapter is migration debt, not a second owner.

## Consequences

- Invalid sheet IDs, addresses, positions, mutually exclusive states, graph
  multiplicity, and ambiguous semantic keys are typed failures.
- Ordinary callers use checked semantic sheet/address keys; numeric order is a
  secondary repair and import path, while raw relationship and part IDs stay
  below the facade.
- `put_chain` moves the caller's model into the workbook cache. Serialization
  borrows that model and writes sheet identifiers through a stack buffer; this
  is a structural copy-avoidance property, not a benchmark result.
- The module has no async-runtime or public lock dependency. Exclusive package
  mutation is expressed through `&mut OpcPackage`.
- The legacy workbook materialization sequence can still mutate the in-memory
  package before a later restoration failure. Resolving that larger save
  transaction remains migration work.
- No new native Excel compatibility claim follows from moving a byte-preserving
  owner. Native evidence is required for behavior that changes emitted Office
  semantics.

## Verification

Nine owner tests cover Strict and Transitional round trips, MCE and extension
preservation, exact bounds, malformed/no-unwind rejection, semantic and numeric
CRUD, ambiguity, graph multiplicity, signature-preserving no-ops, conformance
replacement, shared-target removal, rollback, and real fixture loading. Two
host tests cover consuming cache integration and restoration after writer
materialization.

Warning-denied Clippy and rustdoc pass for `litchi-xlsx` and the focused
`litchi-ooxml` host. The executable boundary checker accepts 35 packages, 106
direct internal dependencies, and 13 explicit debts. Formatting, manifest,
diff, and protected-checklist checks are part of the combined slice gate; per
explicit direction, the previously green full-workspace suite is not repeated.

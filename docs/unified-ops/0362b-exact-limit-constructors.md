# 0362b — Exact limit-constructor names

- Date: 2026-10-21.
- Status: in progress; the iWork raw-wire, Graph/XLS/PPT/ODS constructor slices,
  BIFF framing slice and inventory fix are complete.
  The remaining constructors below are required work, not exceptions.
- Contract: 0351 §2.5 requires options on owner ingress/edit routes and an
  explicit limits argument on raw/record APIs, with no limit-less twin.

## Inventory gap

The inventory checked only the suffix `_with_limits`, so a method named exactly
`with_limits` escaped the inventory. It now reports both forms. A real-rustdoc
regression fixture proves both names are reported and that an exact glossary
ban finds the constructor. All 28 glossary checker tests pass.

The explicit `with_limits` ban is registered for `iwa-common` as that API
migrates. Other profiles will receive it with their corresponding migrations;
the newly detected rows remain visible in the final whole-workspace inventory.

## iWork raw wire

`RawWireFields::new(source, RawWireLimits)` is the sole constructor. The old
`with_limits` name and default-only constructor are removed. Production callers
keep their existing explicit policy. Every test and the Pages hidden-axis fuzz
target passes the same default profile explicitly where it previously used the
default constructor. No alias or deprecation bridge remains.

The constructor body is byte-identical to the former configured constructor;
`results/0362b/iwa-validation.json` records its SHA-256. This is an API
consolidation, with unchanged field/work/depth/input admission, source borrowing,
allocation behavior and scan algorithm. No new runtime layer is introduced.

Full suites for iwa-common, iwa-core, iwa-text-wire and numbers-wire pass (546
tests). The three affected Numbers format suites and the Pages hidden-axis suite
pass; their counts are in the receipt. All-target/all-feature clippy passes for
all six crates and the `iwa-common` glossary profile passes. The migrated Pages
hidden-axis fuzz target compiles with its separate manifest (`cargo check
--offline`); its generated lock file remains ignored.

## Graph, XLS, PPT and ODS constructors

Graph's `Package`, `PackageRef`, `Workbook`, `WorkbookRef` and `Payload` now
share `OpenOptions` and `open_with_options`. Defaults retain the previous
bounds. The PPT `builder::graph::Draft` uses these same Graph ingress
options. Getter `limits_value()` exposes the configured value. The draft and
its `View` move from `chart` into `builder::graph`; the former `GraphPackage`
name and `chart::View` export are deleted. The read projection still constructs
the draft through `draft_package`/`into_draft_package`, and its replacement
still stages through the embedded-object edit. The redundant `SemanticChart`
public alias is removed in favor of `chart::semantic::Model`; its iterator is
`chart::semantic::Models`, with the former prefixed public name removed.

The raw chart `Ref`, `Stream` and `Book` constructors take `Limits` in their
sole `open` entry; discovery uses `Refs::open_workbook(bytes, limits)`. XLS
`ole_object::EmbeddedPayload::new(bytes, limits)` likewise requires its object
bounds. ODS `document::History::with_options(package, EditOptions)` preserves
its selected package/patch/composition/history bounds; `EditOptions` gains its
`limits_value()` getter. Every constructor consumer migrates in this slice.

The four Graph raw constructor bodies are identical after updating their calls
to the canonical names; the XLS payload constructor body is byte-identical. Their normalized body hashes are recorded alongside
validation in `results/0362b/owner-validation.json`. No parsing, graph traversal,
allocation, ownership or publication algorithm changes. Existing tests cover
configured bounds, zero-copy pointer/capacity preservation, exact unknown-record
replay, nested OLE payload admission and history eviction/redo.

All affected targets compile with every feature. The full Graph suite (70),
XLS OLE tests (53), PPT chart tests (14 unit + 5 integration) and ODS history
suite (7) pass: 149 tests, none ignored. All-target/all-feature clippy passes for
the four crates, and the independently locked performance harness compiles.
The exact-name ban is registered in their five glossary profiles.

## BIFF framing and its consumers

The raw APIs have one constructor each: `Records::new(bytes, Limits)`,
`Encoder::new(Limits)`, `Record::open(bytes, Limits)` and
`RecordRef::own(Limits)`. `Encoder::default`, `with_limits` and `own_with`
are deleted. Every XLS and Graph frame consumer passes its previous configured
policy or the previous default explicitly. Constructor admission errors return
before traversal; malformed frame errors still occur at the same frame and
fuse the iterator. The error variants, resource kinds and frame bytes are
unchanged.

Because all iterator construction now checks input admission, the redundant
pending-error field and per-record branch are removed. `Record::open` retains
its private constructor for already checked input, so moving one owned frame
still avoids allocation and a duplicate admission check. The encoder and owned
copy algorithms are unchanged.

Graph record traversal, workbook rediscovery and physical edit scans use their
retained bounds. Previously these paths constructed a default BIFF iterator,
which could reject an already admitted stream with more than one million
records. A regression test opens such a stream under explicit bounds and
checks rediscovery, borrowed traversal, owned traversal and exact copying.
Graph `records()`/`charts()` and PPT chart projections propagate the typed
constructor result; all projection consumers migrate. The raw Graph copy call
is likewise solely `Ref::own(Limits)`.

The complete BIFF and Graph suites pass (15 + 71 tests), as does the complete
XLS library suite (1,311 tests, including corpus and mutation comparisons).
PPT chart validation adds 10 unit and 9 integration tests. All affected targets
compile with every feature; all-target clippy and the independently locked
performance harness check pass. The new `biff` glossary profile bans both limit
twin forms and `own_with`; Graph bans its removed copy twin. BIFF and Graph
produce no inventory rows. Other workspace categories remain open below.

Performance evidence and reproduction commands are in
`results/0362b/biff-validation.json`; its focused raw framing probe is retained
alongside the receipt. It checks configured and default-policy traversal and
encoding, with exact output checksums. This probe does not stand in for the
final integrated workspace gate or the other waves' family measurements.

Three isolated ABBA blocks compare against `90b586ca78`, with the old default
constructors used on the default-policy routes and old configured constructors
on the configured routes. Each sample verifies its generated output checksum.

| Raw route | Median after/before | Three-block range |
|---|---:|---:|
| scan-default | 0.986 | 0.983–0.986 |
| scan-configured | 0.794 | 0.793–0.795 |
| encode-default | 0.993 | 0.985–0.994 |
| encode-configured | 0.994 | 0.982–1.015 |

The iterator occupies 72 bytes instead of 104 on this x86-64 toolchain. The
encoder's configured constructor, owned frame admission and frame-copy bodies
are byte-identical; the receipt records their hashes. Borrowed framing still
allocates nothing. The raw samples include all three blocks, including the
1.015 encoding block; no end-to-end performance claim follows from this probe.

## Remaining required constructors

The IWA raw metadata constructor row is implemented in
[0361f](0361f-iwa-raw-constructors.md): `FieldPath`, `FieldInfo`,
`MessageInfo` and `ArchiveInfo` now require their validated `Limits` profile,
and every workspace consumer migrated in that slice. Its direct raw fields and
identifier-zero behavior remain intentionally public and are checked again at
publication. The source sweep also found exact-name constructors/setters in the legacy
shared/raw layers (ODraw, OLE user properties), signing and DrawingML
helpers. These need canonical required-limit raw calls or options-based
owner/edit routes, with every consumer migrated. Existing option setters need
noun setters and `*_value()` getters. The low-level ZIP and CFB writer setters
also need review for public reachability and the applicable raw/writer
configuration contract.

W4 owns its remaining form-control owner policy setter. W8 owns the two DOCX
source-tail edit setters alongside its OPC API migration. Root coordinates the
other exact-name migrations and the final profile bans. This record does not
close the W4/W4b/W8 program or replace its integrated workspace gate.

The typed-error sweep also remains required. Record
[0359c](0359c-typed-legacy-edit-budgets.md) migrates XLS embedded-payload,
compound-publication and primary staging budgets to typed kinds, followed by
Graph ingress, authoring and fixed-record budgets with every host consumer and
unified classifier migrated. [0359e](0359e-object-capture-budgets.md) migrates
shared OLE core object-capture budgets and their host conversions; raw
directory/OLEDS/link/protection policies remain required work.
[0359d](0359d-graph-package-owner.md)
consolidates Graph's separate `Package` and edit `Snapshot` owners onto the
canonical returned owner. Untouched-stream sharing, nested raw chart models
and shared OLE publication remain required semantic work. The constructor
slices do not accept these as exceptions or complete that audit.

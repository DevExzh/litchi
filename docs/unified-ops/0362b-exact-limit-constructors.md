# 0362b — Exact limit-constructor names

- Date: 2026-10-21.
- Status: in progress; the iWork raw-wire, Graph/XLS/PPT/ODS constructor slices
  and inventory fix are complete.
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

## Remaining required constructors

The source sweep also found exact-name constructors/setters in the legacy
shared/raw layers (BIFF, ODraw, OLE user properties), signing and DrawingML
helpers. These need canonical required-limit raw calls or options-based
owner/edit routes, with every consumer migrated. Existing option setters need
noun setters and `*_value()` getters. The low-level ZIP and CFB writer setters
also need review for public reachability and the applicable raw/writer
configuration contract.

W4 owns its remaining form-control owner policy setter. W8 owns the two DOCX
source-tail edit setters alongside its OPC API migration. Root coordinates the
other exact-name migrations and the final profile bans. This record does not
close the W4/W4b/W8 program or replace its integrated workspace gate.

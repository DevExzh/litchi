# ADR 0023: Dedicated ODF family crates

- Status: Accepted
- Date: 2026-08-04

## Context

`litchi-odf` currently contains every OpenDocument family: text, spreadsheet,
presentation, drawing, chart, image, master, web-template, database, formula,
and flat-document handling. That topology makes unrelated features compile
together, forces family code through a monolithic root, and makes shared XML
and package logic look format-specific. It also prevents the snapshot and
semantic facades described by ADRs 0003 and 0004 from being owned by the
family that can validate them.

## Decision

The workspace is split into one common substrate, one crate per ODF family,
and a deliberately thin umbrella:

```text
litchi-odf-common
├── litchi-odt
├── litchi-ods
├── litchi-odp
├── litchi-odg
├── litchi-odc
├── litchi-odi
├── litchi-odm
├── litchi-oth
├── litchi-odb
└── litchi-odf       (detection and optional ergonomic umbrella)
```

`litchi-odf-common` owns package/archive transactions, manifest and safe path
handling, namespace and scalar vocabularies, bounded common XML utilities,
metadata, shared inert media/object discovery, and genuinely family-neutral
element/style primitives. It never depends on a family crate and does not
publish a family-specific model merely to avoid a migration.

Each family crate owns its package content grammar, semantic model, immutable
snapshot, detached builders, and transactional edit/patch boundary. Its root
facade is small; nested modules use contextual names such as
`document::Document`, `sheet::Sheet`, `slide::Slide`, and `chart::Chart` rather
than `OdtDocument`, `OdsSheet`, or other prefix-expanded names. Large codecs,
models, package graphs, and tests live in separate semantic submodules.

No family crate depends on `litchi-odf` or on another concrete family crate.
The umbrella depends optionally on selected family crates and owns only
detection, feature selection, and concise cross-family open helpers. It may
re-export canonical family modules, but it does not define aliases, duplicate
models, or compatibility forwarding layers. Users who need minimal compile
and memory footprints can depend directly on a family crate.

Flat formats remain in their owning family crate (`.fodt`/`.oth` with text,
`.fods` with spreadsheets, and so on). Formula documents and database
front-ends are independent owners even when they reuse common XML primitives.
ODT-derived master and web documents may reuse shared text capabilities only
through `litchi-odf-common` or a neutral lower layer; they never import the
ODT crate as a peer.

Migration is structural and behavior-preserving in stages:

1. move common substrate ownership downward and make its public vocabulary
   contextual;
2. migrate ODT, ODS, and ODP as complete family slices;
3. migrate ODG, ODC, ODI, ODM, OTH, ODB, and formula/flat paths;
4. reduce `litchi-odf` to detection and optional facade wiring;
5. add ADR 0003 `Snapshot`/`Edit`/`Commit` and reversible patch surfaces at
   each family owner, deleting the old attached mutable root paths as each
   replacement is verified.

During a slice, a re-export is permitted only for a canonical common owner
while call sites move. It must not introduce a second semantic type or a
prefix-expanded name, and it is deleted before that slice is marked complete.

## Consequences

- Family-only builds avoid unrelated parsers, codecs, dependencies, and
  feature code, reducing compile work and resident state.
- Shared logic has one owner, while family-specific validation and package
  graphs remain close to the format that can prove them.
- ADR 0003 edits can return immutable snapshots and reversible patches without
  leaking a monolithic package handle across all ODF families.
- The umbrella remains ergonomic for ordinary applications, but it no longer
  determines the memory footprint or dependency closure of specialized users.
- Native Office/LibreOffice verification and malformed-corpus gates are
  required per family; a green common crate cannot certify a family artifact.

## Verification

Every migrated family must pass its owner unit/property/malformed-input tests,
public API compile tests, package round trips, formatting and diff checks, and
the executable dependency-boundary checker. The umbrella must prove that
feature-disabled family crates are absent from its dependency closure and that
enabled facades expose canonical owner types without aliases. Native evidence
is recorded per family and is not inferred from another family’s artifact.

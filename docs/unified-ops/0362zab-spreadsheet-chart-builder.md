# 0362zab — Spreadsheet chart authoring placement

This leaf applies 0351 §2.1.7 to the shared spreadsheet chart model. A
worksheet `Chart` is freshly authored state: its `add_series` and related
construction methods belong to a builder surface, not to a package owner or
read-only projection. The change keeps the existing model, codec, resource
relationships and serialized behavior; it moves their public authoring
entry point to a focused builder namespace.

## Canonical paths

The shared definitions are re-exported from one place:

```text
litchi_spreadsheet_drawing::chart::builder::{
    Anchor, Chart, ExternalDataPart, ExternalDataTarget, Relationship, Series,
    Target, UserShapesPart,
}
```

Each spreadsheet host exposes those exact definitions at its canonical
authoring path:

```text
litchi_xlsx::builder::chart::{Anchor, Chart, ExternalDataPart,
    ExternalDataTarget, Relationship, Series, Target, UserShapesPart}
litchi_xlsb::builder::chart::{Anchor, Chart, ExternalDataPart,
    ExternalDataTarget, Relationship, Series, Target, UserShapesPart}
```

The old direct authoring re-exports from `litchi_xlsx::chart` and
`litchi_xlsb::chart` are removed, not aliased. Their `chart::read` and
`chart::write` codec entry points remain. Low-level `anchor`, `relationship`
and codec modules remain where the host implementation needs them. The
XLSX `drawing::Chart` and the XLSB drawing/package chart projections remain
read-oriented values and are not replaced by the builder model.

## Consumer and policy migration

Every chart authoring consumer was migrated in this change: spreadsheet
codec tests and fixtures, XLSX and XLSB writer internals, examples, the
performance harness, XLSB integration tests, `litchi-ops` tests and the
unified facade tests. No duplicate root path or renamed `ChartAnchor` alias
was introduced; the chart anchor is consistently `builder::chart::Anchor`.

The XLSX and XLSB minimal/full glossary profiles now ban each removed direct
child (`Anchor`, `Chart`, `ExternalDataPart`, `ExternalDataTarget`,
`Relationship`, `Series`, `Target` and `UserShapesPart`) under their `chart`
scope, with the corresponding `builder::chart` replacement. The new
builder paths are classified as authoring drafts, not immutable owners. The
profiles also ban the former `litchi_xlsx::drawing::Anchor` and
`litchi_xlsb::builder::{Chart, ChartAnchor}` paths; the XLSB root
`builder::Anchor` remains the unrelated shape/image authoring value. A fresh
inventory should therefore report the exact builder paths and no old
host-root authoring re-export.

## Validation

The focused locked offline build used the pinned workspace lock
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`, one
Cargo job and the 20 GiB capped environment:

```text
cargo check --locked --offline -p litchi-spreadsheet-drawing -p litchi-xlsx -p litchi-xlsb --lib
```

The command passed. Focused library and consumer tests, formatting, glossary
JSON validation and the staged-file size guards are run before this leaf is
committed. No performance qualification is claimed here: the implementation
reuses the same chart model and codec paths, so this namespace-only change
does not intentionally alter serialization or allocation behavior.

The completed focused validation also includes:

```text
litchi-spreadsheet-drawing: 31 tests passed
litchi-xlsx:              1,876 tests passed
litchi-xlsb:                768 tests passed
cargo check -p litchi-ops --tests
cargo check -p litchi --tests
cargo check -p litchi --features xlsx --example xlsx_charts
cargo check -p litchi --features xlsx --example xlsx_charts_comprehensive
unified-chart-add-baseline: cargo check --offline
check_api_glossary.py: xlsx-minimal, xlsx-full, xlsb-minimal and xlsb-full passed
```

The host rustdoc captures expose the eight canonical `builder::chart` uses,
with their source identity bound to
`litchi_spreadsheet_drawing::chart::builder`; the generated JSON and inventory
TSVs remain outside the repository. The workspace-wide formatter still has
pre-existing 2024-edition files outside this leaf; all changed Rust files
were formatted with the repository's edition and `git diff --check` passed.
The standalone performance harness lock needed its existing dependency graph
refresh to include the already-present `litchi-sheet-text` workspace package;
that generated lock update is committed with the consumer import changes.
The lock diff changes only local workspace dependency edges; registry package
versions and checksums remain unchanged, so the locked offline harness remains
reproducible.

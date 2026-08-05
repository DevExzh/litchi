# ADR 0024: Current post-migration workspace topology

- Status: Accepted — current-state inventory
- Date: 2026-08-05
- Scope: Documentation only; this record does not rewrite historical migration
  slices.

## Authority

The workspace is defined by [`Cargo.toml`](../../Cargo.toml), whose current
membership is `crates/*`. The package inventory below is based on the current
manifests and `cargo metadata --no-deps`; it describes package ownership, not a
compatibility promise.

## OOXML

There is no current package named `litchi-ooxml`: no such package appears in
workspace metadata and there is no `crates/litchi-ooxml/Cargo.toml`. The
standalone format owners are [`litchi-docx`](../../crates/litchi-docx/Cargo.toml),
[`litchi-pptx`](../../crates/litchi-pptx/Cargo.toml),
[`litchi-xlsx`](../../crates/litchi-xlsx/Cargo.toml), and
[`litchi-xlsb`](../../crates/litchi-xlsb/Cargo.toml).

The current dependency layers are:

```text
litchi-opc
└── litchi-ooxml-common
    └── litchi-drawingml
        ├── litchi-docx
        ├── litchi-pptx
        ├── litchi-xlsx
        └── litchi-xlsb
```

The diagram shows the shared foundation direction; the concrete format crates
also depend directly on the common package and OPC owners where their manifests
require them. [`litchi-ooxml-common`](../../crates/litchi-ooxml-common/Cargo.toml)
owns shared OOXML package vocabulary and services,
[`litchi-drawingml`](../../crates/litchi-drawingml/Cargo.toml) owns
host-neutral DrawingML, and [`litchi-opc`](../../crates/litchi-opc/Cargo.toml)
owns physical OPC packaging.

The root [`litchi` manifest](../../crates/litchi/Cargo.toml) retains an
`ooxml` feature and its [`ooxml` facade](../../crates/litchi/src/lib.rs), but
that facade re-exports the standalone owners; it is not a replacement package
named `litchi-ooxml`.

## OLE2 and legacy binary formats

The current legacy container and shared-object layers are:

```text
litchi-cfb
└── litchi-ole-common
    ├── litchi-doc
    ├── litchi-ppt
    └── litchi-xls
```

This is reflected by the [`litchi-ole-common` manifest](../../crates/litchi-ole-common/Cargo.toml)
and the manifests for [`litchi-doc`](../../crates/litchi-doc/Cargo.toml),
[`litchi-ppt`](../../crates/litchi-ppt/Cargo.toml), and
[`litchi-xls`](../../crates/litchi-xls/Cargo.toml). The common crate owns
format-neutral validated OLE2 structures; host metadata and semantic records
remain in the concrete legacy format crates.

## ODF

[`litchi-odf-common`](../../crates/litchi-odf-common/Cargo.toml) is the shared
OpenDocument substrate. Dedicated family packages currently present in the
workspace are:

`litchi-odt`, `litchi-ods`, `litchi-odp`, `litchi-odg`, `litchi-odc`,
`litchi-odi`, `litchi-odm`, `litchi-oth`, `litchi-odb`, and
`litchi-odf-formula`.

Their split and ownership are recorded in
[ADR 0023](0023-odf-family-crate-split.md). The
[`litchi-odf` manifest](../../crates/litchi-odf/Cargo.toml) makes `odt`, `ods`,
and `odp` the default family features and provides `all` for the remaining
families. Its [`facade`](../../crates/litchi-odf/src/lib.rs) owns detection and
feature-gated family re-exports only; family package, model, and authoring
ownership remains in the dedicated crates. The top-level [`litchi` manifest](../../crates/litchi/Cargo.toml)
similarly exposes the primary ODF families through its `odf` facade feature.

## Historical terminology

References to `litchi-ooxml` in ADR 0002, ADR 0008, ADR 0011, ADR 0013, ADR
0014, ADR 0015, ADR 0017, and ADR 0018 describe the former migration host or a
verification slice. They are intentionally retained as historical evidence;
they do not describe a current package or dependency. This record supplies
the current terminology without deleting or rewriting those records.

## Verification

The audit used current workspace manifests, `cargo metadata --no-deps
--format-version 1`, focused reference searches, and the existing topology
decisions in [ADR 0002](0002-crate-topology.md) and
[ADR 0023](0023-odf-family-crate-split.md). No source files were changed.

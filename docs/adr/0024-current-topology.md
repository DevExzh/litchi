# ADR 0024: Current post-migration workspace topology

- Status: Accepted — current-state inventory
- Date: 2026-08-06
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

Within the concrete owners, large package and semantic domains are layered
under contextual folders rather than kept as one source file. DOCX now has
`document` and `paragraph` owners; PPTX has `presentation`; and XLSX has
`workbook::{edit,worksheet,data_model,comments}` plus `views`. Each facade
keeps its semantic model separate from XML/package codecs and focused tests.
XLSB's workbook owner follows the same structure.

The current continuation applies the same topology to RTF's codec and
document owners, DrawingML chart reader/writer owners, OGraph chart records,
DOC writer core, DOCX fields, ODT parser, ODS tracked changes, ODP parser, XLS
writer core, XLSB workbook writing, and the XLSX chart-sheet package. These
folders preserve their existing public owner paths while making model, codec,
package, and test responsibilities explicit. The PPTX notes writer retains
the same layered ownership and now emits quoted namespace attributes.

The latest continuation adds the same boundaries to PPT writer core and
Escher, PPTX ChartEx, XLSB conditional formatting and workbook codecs, ODT
fields, ODS content, DOCX document writing, XLSX worksheet snapshot editing,
DOC PAP/TAP, and XLS pivot tables. The DOC package facade exposes contextual
`OpenOptions`, `EncryptionKind`, and `Error` names without compatibility
aliases; section-border validation is surfaced separately as `BorderError`.

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
format-neutral validated OLE2 structures, including the typed `property_set`
model/codec/editor shared by DOC, PPT, and XLS; host metadata and
format-specific semantic records remain in the concrete legacy format crates.
The property-set editor keys staged mutations by complete CFB stream paths, so
equal leaf names in different storage trees remain distinct during rewrites.

The current legacy owners also use nested semantic folders: DOC fields, PPT
animation parser/types/writer, XLS list objects, and ODraw properties expose
facades over model/codec/package/test seams. These are source-organization and
ownership boundaries; they do not imply a broader compatibility promise.

The latest owner pass adds nested `parts/chp`, `parts/fields/codec`, and
`writer/core/package` seams to DOC; moves the shared OfficeArt wire model into
ODraw; and layers RTF lexer/parser/writer, IWA media and Numbers editor,
DrawingML chart reader, OGraph chart records, PPTX animations, XLS list and
writer codecs, XLSB worksheet writing, and XLSX catalog editing. Each owner
keeps a small contextual facade while separating semantic models, wire/XML or
binary codecs, package operations, validation, and focused tests. The public
surface remains prefix-free within each format context: DOC exports concise
`Leniency`, `ToleranceReport`, `StylesheetDefect`, `EncryptionProfile`,
`Element`, and `Section` names without compatibility aliases.

The DOC writer facade now exports contextual `Writer`, `WriteError`,
`HeaderKind`, `Picture`, `SmartTagEntry`, `StyleRevision`, and
`StyleWriteError`; writer-only `StyleDefinition` remains under
`litchi_doc::writer` because the root reader facade already owns that name.
Tracked revisions, MTEF equation options, text boxes, and small writer I/O
errors use the same prefix-free rule. DOC also has a nested
`parts/route_slip` owner for typed, lossless MS-DOC routing-slip metadata; its
FIB/table-stream parser and serializer are exposed without claiming Document
lifecycle integration or protection-policy enforcement.

The OLE2 owner now also has `parts/ole_controls`, which layers the inert
`OcxInfo`/`RgxOcxInfo` metadata model, binary codec, FIB/table-stream seam, and
tests without creating a control runtime or activation API.

The shared `litchi-ole-common::toolbar` owner now layers the bounded,
format-neutral `[MS-OSHARED]` `WString`, toolbar/control headers, flags, and
dimensions into model and codec seams. It preserves borrowed payloads and
reserved bits and remains inert: DOC/PPT/XLS command-bar lifecycle wiring and
macro/UI execution are intentionally outside this common owner.

DOC now adds a contextual `parts/command_bars` owner on top of that common
codec. Its public `CommandBars` facade reads and writes the optional FIB
`fcCmds`/`lcbCmds` table range, exposing bounded macro-command, allocated-
command, key-map, and CTBWRAPPER metadata without activating any command. The
owner intentionally refuses variable TBC data and unknown Tcg records when a
safe boundary cannot be recovered.

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

ODT's field, builder, and mutable owners and ODS's content codec are layered
inside their family crates. The ODS content owner is parser-only and therefore
does not own package assembly; package ownership remains with the family
facade. The OOXML-common web-extension codec is likewise layered into
semantic, XML, relationship, and package owners while preserving the compact
public web facade.

ODP's parser and ODT's index writer now follow the same semantic/XML/
validation/package/test organization. The current continuation layers ODS
content traversal and ODT mutable editing under nested semantic/validation or
snapshot/package codec facades. The same wave layers DOCX document packages,
web extensions, and section writing; PPT writer-core models; PPTX ChartEx
semantic records; XLS revision records; XLSB host cell reading; and XLSX pivot
reading. These changes are source topology and ownership evidence only; they
do not broaden the format conformance claims in ADR 0023.

The latest dense-owner pass additionally layers ODS data-pilot and ODT
graphic-property models, OGraph chart models, and the DOC/DOCX/PPT/PPTX/XLSB/
XLSX owners listed in ADR 0008. Their facades retain typed snapshots and
format-specific ownership while moving semantic, wire/XML, validation, and
test responsibilities into contextual folders.

The current continuation extends that inventory with DOC document and
writer-core models, DOC field tests, DOCX field tests, DrawingML chart-reader
semantic domains, IWA Numbers editor semantics, ODS traversal, ODT parser
codec, PPTX ChartEx validation, RTF content fields, XLSB pivot writing, and
XLSX workbook-edit tests. These are nested ownership boundaries; they do not
make the RTF lint backlog or Office conformance claims disappear.

The next owner continuation layers DOC numbering, DrawingML diagram data, IWA
editor tables, ODS style protection, ODT mutable semantics, PPT animation test
domains, PPTX shape tags, XLS workbook codecs, XLSB conditional-formatting
binary codecs, and XLSX data-validation codecs into contextual model, codec,
validation, and test folders. XLS also exposes a bounded `toolbar` facade for
the `[MS-XLS]` XCB stream. It reuses the shared `[MS-OSHARED]` toolbar model,
preserves reserved and fixed visual bytes, and refuses variable or unknown
control payloads without activating macros, UI, or ActiveX behavior. These
owners remain format-local and prefix-free while shared wire logic stays in
the common crates.

The following continuation applies the same topology to the shared
`property_set` binary codec, DOC image writing, PPT comparison and embedded
objects, ODraw images, OGraph package assembly, XLS query tables, DOCX field
tests, PPTX tag packages and animation XML, XLSB workbook-writer tests, XLSX
raw worksheet and snapshot editors, ODS sheet traversal, and ODT field
codecs. The XLS toolbar owner is now package-integrated: its `Workbook` and
`Writer` facades own the optional root `XCB` stream while the common toolbar
model owns borrowed-to-owned lifetime conversion. All control and command
behavior remains inert.

The latest continuation also adds the typed DOC `ObjInfo`/`ODTPersist2`
metadata layer and a bounded PPT `animation::diagram_build` owner for
`DiagramBuildContainer`/`DiagramBuildAtom` records. Both retain fixed-width
unknown values and reserved bytes where safe, reject malformed boundaries,
and expose no activation or playback runtime. DOC document semantics,
sections, and form fields; PPT writer records and animation editing; ODraw
property groups; OGraph chart aggregates; XLS list-object semantics and pivot
writing; XLSB pivot writing; XLSX package metadata; DOCX package tests; and
PPTX shape anchors now follow the same nested facade/model/codec/validation/
test organization.

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
[ADR 0023](0023-odf-family-crate-split.md). The layered owner paths are
verified by the affected-crate all-target compile and boundary-policy check.

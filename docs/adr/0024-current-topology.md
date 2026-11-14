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

Spreadsheet Data Model storage is a separate shared branch owned by
[`litchi-xldm`](../../crates/litchi-xldm/Cargo.toml):

```text
litchi-xldm
├── litchi-xlsb
└── litchi-xlsx
```

The XLDM crate is dependency-free within the workspace and owns the bounded,
source-sharing storage, metadata, identity, and OLAP-proof layers. The XLSB
and XLSX crates retain their format-specific package bindings and expose the
shared owner only through their contextual data-model/package surfaces. Both
edges are normal runtime dependencies; neither workbook crate reimplements or
owns a second XLDM storage model.

The root [`litchi` manifest](../../crates/litchi/Cargo.toml) retains the
`ooxml` feature gate, but its public facade exposes the standalone owners
directly as `litchi::{docx, pptx, xlsx, xlsb}` (alongside `opc` and
`ooxml_common`); no public `litchi::ooxml` wrapper remains. This is not a
replacement package named `litchi-ooxml`.

Within the concrete owners, large package and semantic domains are layered
under contextual folders rather than kept as one source file: DOCX has
`document` and `paragraph` owners, PPTX has `presentation`, XLSX has
`workbook::{edit,worksheet,data_model,comments}` plus `views`, and XLSB's
workbook owner follows the same structure. The layering rule for every format
is in [Owner layering](#owner-layering) below.

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

The DOCX drawing inventory also exposes the checked Word 2010 `AnchorId` value
on `wp:inline` and `wp:anchor` objects. Its `[MS-DOCX]`/`[MS-ODRAWXML]`
eight-digit hexadecimal range is enforced at the XML boundary; object/pict
authoring and layout remain outside this small inert inventory owner.

The PPT bookmark-summary owner is layered as
`bookmark_summary/{model,codec,validation,tests}.rs`; its semantic collection is
`bookmark_summary::Summary`. The record codec is inert and bounded, while
summary-to-text-bookmark identity checks stay in the validation layer.

The shared OfficeArt wire model is owned by ODraw.

The DOC writer facade exports contextual `Writer`, `WriteError`,
`HeaderKind`, `Picture`, `SmartTagEntry`, `StyleRevision`, and
`StyleWriteError`; writer-only `StyleDefinition` remains under
`litchi_doc::writer` because the root reader facade already owns that name.
Tracked revisions, MTEF equation options, text boxes, and small writer I/O
errors use the same prefix-free rule. DOC has a nested `parts/route_slip`
owner for typed, lossless MS-DOC routing-slip metadata. The public
`litchi_doc::route_slip` facade layers its FIB/table-stream codec, bounded
validation, package editor, exact recipient selectors, and snapshot
transactions. `Document::route_slip()` exposes the optional metadata through a
deferred `Result`, while the package editor publishes reversible route and OLE
patches and rejects protected lifecycle edits. This remains passive metadata
ownership: authentication, mail transport, and host routing are not
implemented.

DOC's `parts/ole_controls` layers the inert
`OcxInfo`/`RgxOcxInfo` metadata model, binary codec, FIB/table-stream seam, and
tests without creating a control runtime or activation API.

The shared `litchi-ole-common::toolbar` owner layers the bounded,
format-neutral `[MS-OSHARED]` `WString`, toolbar/control headers, flags,
dimensions, and typed `TBCGeneralInfo`/`TBCExtraInfo` payloads into model and
codec seams. It preserves borrowed strings, typed merge modes, reserved bits,
and format-specific tails without allocating decoded source data. It remains
inert: DOC/PPT/XLS command-bar lifecycle wiring and macro/UI execution are
intentionally outside this common owner. XLS's `toolbar` facade owns the optional root
`[MS-XLS]` `XCB` stream through its `Workbook` and `Writer` facades.

DOC has a contextual `parts/command_bars` owner on top of that common
codec. Its public `CommandBars` facade reads and writes the optional FIB
`fcCmds`/`lcbCmds` table range, exposing bounded macro-command, allocated-
command, key-map, and CTBWRAPPER metadata without activating any command. It
also decodes bounded variable TBC data through the common model and rejects
ambiguous boundaries or unknown Tcg records when a safe boundary cannot be
recovered.

DOC has a contextual `parts/envelope` owner for the optional
`fcMsoEnvelope`/`lcbMsoEnvelope` FIB range. Its typed `Envelope` facade models
the documented `[MS-OSHARED]` Office 6/8 message body, recipient property bags,
and attachment metadata, while retaining unknown CLSID payloads as bounded
opaque bytes. `Document::envelope()` is read-only and inert; no mail transport,
recipient resolution, attachment activation, or package-writer emission is
part of this owner.

DOC captions use the matching `parts/captions/{model,codec,validation,
transaction,package,tests}` boundary. The `captions::Editor` owns atomic CRUD
over the `[MS-DOC]` `SttbfCaption`/`SttbfAutoCaption` FIB ranges and publishes
reversible semantic and CFB byte patches; new payloads are appended and clear
operations only clear pointers, so unrelated table-stream bytes remain opaque.
Caption fields and host automation remain inert.

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

## Owner layering

Every format owner keeps a concise, prefix-free contextual facade over nested
semantic model, wire/XML or binary codec, package or transaction, validation,
and focused-test modules. Facades expose no compatibility aliases or
module-prefixed duplicate names. Editable owners follow the
`Snapshot`/`Transaction`/`Commit`/`Patch` shape of ADR 0003: edits are staged
on a clone, source-checked, published atomically, and reversible, and a patch
applied to a different artifact is a typed conflict. Unknown XML, BIFF and
BIFF12 records, CFB topology, and reserved wire bits stay opaque and are
preserved through edits.

All such owners are inert: none activates macros, ActiveX or OLE objects,
controls, add-ins, DDE, external links or data connections, media playback,
collaboration or broadcast features, rendering or layout, schema retrieval, or
formula and revision replay. Invalid state returns typed errors rather than
panicking.

Owners shared across formats live in the common crates:

- `litchi-ole-common`: the `property_set` model, codec, and editor (including
  SummaryInformation, DocumentSummaryInformation, arrays, vectors, versioned
  streams, `HeadingPairs`, and `DocParts`), the `toolbar` model, `object`
  snapshots with `Commit`/`Patch` and OLEDS link metadata, CFB directory
  catalogs, smart tags, Custom XML stores, DataSpaces, and VBA-signature
  metadata. DOC, PPT, and XLS publish these through host-validated package
  snapshots.
- `litchi-ooxml-common`: typed relationship identifiers, bounded MCE
  `AlternateContent` choices, Custom XML Data Storage items and properties,
  web extensions, and the other shared OOXML package vocabulary.
- `litchi-drawingml`: colors and ordered color transforms, transforms,
  model3d resources, and the host-neutral chart snapshot/editor for series,
  axis, label, and metadata edits; concrete packages keep only placement and
  relationship ownership.
- `litchi-odf-common`: `style::master` (shared by ODP master-page editing) and
  `chart::authoring` (the chart definition, cached table, validation, and XML
  writer shared by standalone ODC and embedded ODT charts). `litchi-odc` keeps
  only its standalone package builder, and ODT does not depend on ODC.

The top-level `litchi` facade exposes `docx`, `pptx`, `xlsx`, `xlsb`, `odp`,
`ods`, and `odt` directly (alongside `opc`, `ooxml_common`, and
`odf_common`), and `litchi-odf` remains a thin feature-gated detector and
family umbrella.

## IWA and iWork

`litchi-pages`, `litchi-numbers`, and `litchi-keynote` are the concrete
application package owners. The shared physical and wire substrate is
`litchi-iwa-archive` (ZIP/package preservation), `litchi-iwa-core`
(Snappy/IWA framing and neutral archive metadata), `litchi-iwa-cache`,
`litchi-iwa-detect`, `litchi-iwa-index`, `litchi-iwa-graph`,
`litchi-iwa-package`, `litchi-iwa-protos`, and `litchi-numbers-wire`, a
low-level BNC adapter excluded from the supported format and root facades.
Shared engines and neutral values live in `litchi-iwa-common`,
`litchi-iwa-style`, `litchi-iwa-table`, `litchi-iwa-text`,
`litchi-iwa-text-wire`, and `litchi-iwa-structured`.

`litchi-iwa-style` is the shared drawing-style and media engine of the three
application owners: the style-variation graph engine and the drawing and media
field passes, which would otherwise be copied into each leaf.
`litchi-iwa-table` holds the table-family engines (cell formats, controls and
the engines that follow them) in format-neutral form over one resolved table
model, which each leaf supplies through its `TableSource`; the leaves keep
selectors, publication, and their public transaction types. Neither crate is
re-exported by a facade: their signatures name native archive positions, which
the supported APIs do not expose.

The `litchi-iwa` migration host is deleted
([ADR 0028](0028-iwa-monolith-exit.md)). There is no migration host and no
migration debt.

## Conversion and interchange

[`litchi-markdown`](../../crates/litchi-markdown/Cargo.toml) is the
dependency-light owner of Markdown configuration, Unicode helpers, and the
format-neutral `ToMarkdown` trait. It does not parse documents and does not
depend on any concrete Office or OpenDocument format crate.

Concrete format adapters currently live in the top-level
[`litchi` facade](../../crates/litchi/src/markdown), where the selected format
features and their semantic models are available. This placement is an adapter
boundary, not a claim that each format crate implements `ToMarkdown`, and it
does not make Markdown a document-format owner or a bidirectional conversion
layer. Rendering, pagination, external retrieval, and active-content execution
remain outside this helper crate.

## Shared spreadsheet data model

[`litchi-xldm`](../../crates/litchi-xldm/Cargo.toml) is the neutral shared
Spreadsheet Data Model owner. It owns bounded storage inspection and writing
plus metadata, generated path, native, identity, and OLAP-closure projections
while retaining borrowed source bytes. `litchi-xlsx` keeps its `package::xldm`
facade and error mapping; `litchi-xlsb` uses the same owner for its data-model
identity and closure proof. Package binding and workbook-specific transactions
remain in those two concrete crates. The shared owner has no workspace
dependencies, and only `litchi-xlsx` and `litchi-xlsb` depend on it.

## Unified operations, command line, and text workbooks

[`litchi-ops`](../../crates/litchi-ops/Cargo.toml) is the facade-tier package of
[ADR 0033](0033-unified-operations-api.md): the flat `File` enum, unified error,
detection, open options, node view, property registry, batch executor,
capability reporting, Markdown output, formula evaluation, and the per-format
handlers. Its edges point downward to the leaf and common crates and are
optional, each owned by the format or capability feature that needs it
(including the Pages, Keynote, and Numbers leaves and `litchi-iwa-detect` and
`litchi-iwa-archive` for iWork). The `litchi` facade depends on it, and
`litchi::open` returns its `File`. It is runtime-neutral.

[`litchi-cli`](../../crates/litchi-cli/Cargo.toml) is the `litchi` binary: the
command-line verbs over the public `litchi-ops` API, with no document
behaviour of its own. It depends only on `litchi-ops` and `litchi-core`.

[`litchi-sheet-text`](../../crates/litchi-sheet-text/Cargo.toml) owns the CSV,
TSV, PRN, SYLK, DIF, and fixed-width text-workbook readers and writers. It
depends on `litchi-core` and `litchi-sheet`; the facade re-exports it as
`litchi::sheet::text` behind its `sheet-text` feature, and `litchi-ops` uses its
delimited tokenizer for `File::import`.

## Historical terminology

References to `litchi-ooxml` in ADR 0002, ADR 0011, ADR 0014, ADR 0015,
ADR 0017, and ADR 0018 name the former migration host. They do not describe a
current package or dependency; this record supplies the current terminology.

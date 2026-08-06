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

The root [`litchi` manifest](../../crates/litchi/Cargo.toml) retains the
`ooxml` feature gate, but its public facade exposes the standalone owners
directly as `litchi::{docx, pptx, xlsx, xlsb}` (alongside `opc` and
`ooxml_common`); no public `litchi::ooxml` wrapper remains. This is not a
replacement package named `litchi-ooxml`.

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

The DOCX drawing inventory also exposes the checked Word 2010 `AnchorId` value
on `wp:inline` and `wp:anchor` objects. Its `[MS-DOCX]`/`[MS-ODRAWXML]`
eight-digit hexadecimal range is enforced at the XML boundary; object/pict
authoring and layout remain outside this small inert inventory owner.

The current legacy owners also use nested semantic folders: DOC fields, PPT
animation parser/types/writer, XLS list objects, and ODraw properties expose
facades over model/codec/package/test seams. These are source-organization and
ownership boundaries; they do not imply a broader compatibility promise.

The PPT bookmark-summary owner is now layered as
`bookmark_summary/{model,codec,validation,tests}.rs`. Its canonical semantic
collection is `bookmark_summary::Summary`; the former module-prefixed
`BookmarkSummary` spelling is removed rather than retained as an alias. The
record codec remains inert and bounded, while summary-to-text-bookmark identity
checks stay in the validation layer.

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
errors use the same prefix-free rule. DOC also has a nested `parts/route_slip`
owner for typed, lossless MS-DOC routing-slip metadata. The public
`litchi_doc::route_slip` facade layers its FIB/table-stream codec, bounded
validation, package editor, exact recipient selectors, and snapshot
transactions. `Document::route_slip()` exposes the optional metadata through a
deferred `Result`, while the package editor publishes reversible route and OLE
patches and rejects protected lifecycle edits. This remains passive metadata
ownership: authentication, mail transport, and host routing are not
implemented.

The OLE2 owner now also has `parts/ole_controls`, which layers the inert
`OcxInfo`/`RgxOcxInfo` metadata model, binary codec, FIB/table-stream seam, and
tests without creating a control runtime or activation API.

The shared `litchi-ole-common::toolbar` owner now layers the bounded,
format-neutral `[MS-OSHARED]` `WString`, toolbar/control headers, flags,
dimensions, and typed `TBCGeneralInfo`/`TBCExtraInfo` payloads into model and
codec seams. It preserves borrowed strings, typed merge modes, reserved bits,
and format-specific tails without allocating decoded source data. It remains
inert: DOC/PPT/XLS command-bar lifecycle wiring and macro/UI execution are
intentionally outside this common owner.

DOC now adds a contextual `parts/command_bars` owner on top of that common
codec. Its public `CommandBars` facade reads and writes the optional FIB
`fcCmds`/`lcbCmds` table range, exposing bounded macro-command, allocated-
command, key-map, and CTBWRAPPER metadata without activating any command. It
also decodes bounded variable TBC data through the common model and rejects
ambiguous boundaries or unknown Tcg records when a safe boundary cannot be
recovered.

DOC also adds a contextual `parts/envelope` owner for the optional
`fcMsoEnvelope`/`lcbMsoEnvelope` FIB range. Its typed `Envelope` facade models
the documented `[MS-OSHARED]` Office 6/8 message body, recipient property bags,
and attachment metadata, while retaining unknown CLSID payloads as bounded
opaque bytes. `Document::envelope()` is read-only and inert; no mail transport,
recipient resolution, attachment activation, or package-writer emission is
part of this owner.

DOC captions now add the matching `parts/captions/{model,codec,validation,
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
preserves reserved and fixed visual bytes, and now round-trips `TBCCmd` plus
bounded variable `TBCData` through the shared typed general metadata model.
Ambiguous or unknown control payloads are still rejected without activating
macros, UI, or ActiveX behavior. These owners remain format-local and
prefix-free while shared wire logic stays in the common crates.

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

The current owner continuation further layers DOC embedded-object
transactions, field parsing, writer package semantics, and writer tests; ODS
data-pilot parsing; PPT embedded objects, animation behavior, text-format, and
text-style writers; XLS OLE objects and writer streams; XLSB formula text and
worksheet writing; and XLSX ActiveX and XLDM package owners. DOC's
`parts/ole_controls` facade now owns the specified 20-byte `OcxInfo` body,
ObjectPool metadata, and the live document FIB seam without retaining the old
`parts/ole/controls` owner. These additions keep semantic, wire/XML or BIFF,
validation, package, and test ownership nested; they remain inert with respect
to control activation, macros, and external behavior.

The current continuation extends the topology with layered ODT document and
text-element owners; ODS database-range and table-template style owners; ODP
parser XML and authoring-builder owners; DOC OLE metadata; PPT animation
timing; XLS differential formats; DOCX glossary codecs; PPTX animation and
modern-comment codecs; XLSB formula/resolution; and XLSX worksheet-snapshot
and workbook-transaction owners. Each keeps a concise contextual facade over
semantic, wire/XML, package, validation, and test modules.

`litchi-ole-common::object` now additionally owns an immutable `Snapshot`
read facade. It shares captured stream buffers across clones and creates
independent transactional editors, keeping large OLE payloads out of format
neutral copies while leaving DOC/PPT/XLS interpretation in their owners.
Its public `object::Commit` now pairs a validated post-edit `Snapshot` with a
reversible, source-checked `object::Patch`; applying a patch to a different
artifact is a typed conflict rather than a last-writer-wins replacement.

The current continuation adds typed multidimensional `property_set::Array` and
scalar-typed `property_set::Vector` models, including checked
`VT_ARRAY|VT_VARIANT` and `VT_VECTOR|VT_VARIANT` element headers. Their binary
codec and validation remain solely in `litchi-ole-common`; unsupported or
malformed OLE Property Set types are inert or rejected at that boundary. The
common toolbar owner is split into semantic subdomains, while DOC field and
OfficeArt snapshots, PPT chart/ODraw, XLS chart/OLE controls, and XLSB pivot
definition/record validation stay in their contextual owners.

The current OLE2 continuation extends that common boundary with code-page-aware
`[MS-OSHARED]` `HeadingPairs` and `DocParts` composite values, while DOC,
ODraw, PPT, and XLS keep their format-specific owners. DOC now authors checked
`Asumyi`/`PlcfAsumy` ranges; ODraw exposes typed solver rules; PPT master
metadata authors `SlideNameAtom`; and XLS owns a layered `[MS-XLS]` `XML`
stream model with typed schema/map/data-binding identities and list-column
dependency validation. These are bounded metadata operations: no macros,
external binding, schema resolution, layout, or rendering is activated.

ODT mutable editing, ODS authoring and formula evaluation, ODP authoring,
DOCX paragraph codecs, PPTX presentation properties, and XLSX chart-sheet
package operations now use nested model, codec, package/transaction,
validation, and test folders. These changes preserve the prefix-free facade
rule and the standalone OOXML/ODF crate topology; no compatibility wrapper or
duplicate shared-format grammar was introduced.

The current OLE2/OOXML/ODF continuation adds DOC route-slip lifecycle edits,
DOCX run effects, ODP handout masters, ODS metadata and calculation settings,
ODraw custom geometry, OGraph chart patches, PPT master inventories and chart
host replacement, PPTX model3d resources, and bounded slicer/timeline owners in
XLSB and XLSX. Each owner keeps semantic, wire/XML or BIFF, package,
validation, and focused-test layers; common OLE snapshot commits stay at the
artifact boundary, while host crates retain semantic dependency closure.
Unsupported records and active behavior remain inert or lossless, and invalid
state returns typed errors rather than panicking.

This continuation adds ODP master-page editing through the shared
`litchi-odf-common::style::master` model, a transactional ODS worksheet graph,
DOC captions and ObjectPool/ActiveX metadata, DOCX settings extensions, PPT
diagram inventories, PPTX modern-comment V2 commands and zoom owners, XLS
chart snapshots, XLSB scenario and threaded-comment owners, XLSX rich-value
and feature-property-bag owners, and shared DrawingML model3d resources. Each
new owner keeps semantic, wire, validation, package/transaction, and focused
test seams nested behind a concise prefix-free facade; unknown records remain
opaque where the specification requires preservation, and no runtime activates
macros, controls, links, collaboration, rendering, or external code.

The current boundary cleanup makes DOC's parsed `document` owner private and
exports `Document` only from the `litchi_doc` root facade; no public
`litchi_doc::document` compatibility path remains. The shared OLE Property Set
owner additionally types `[MS-OLEPS]` `VT_VERSIONED_STREAM` values with checked
indirect property names, bounded code-page strings, and inert version GUIDs.
The referenced CFB stream remains host/package data and is never opened or
executed by the common semantic layer.

The ODraw picture-property owner now layers `[MS-ODRAW]` `pibName` and
`pibFlags` into `prop::picture::{Metadata, Snapshot, Edit}`. Picture names are
checked, bounded UTF-16LE views; valid flag dependencies are typed while
undefined producer bits remain exact. A committed edit returns an owned
snapshot and reversible patch, rewrites only the modeled descriptors, and
preserves source order plus every untouched opaque property payload.

The ODF chart-content authoring owner now lives under
`litchi-odf-common::chart::authoring`. It owns the typed definition, cached
table, extension, validation, and deterministic XML writer layers shared by
standalone ODC and embedded ODT charts. `litchi-odc` retains only its
standalone package builder/facade, and ODT no longer depends on the peer ODC
family crate; package topology and embedded-object mutation remain in their
owning family crates.

The current migration adds typed relationship identifiers in
`litchi-ooxml-common`, shared DrawingML colors in `litchi-drawingml`, ODS
embedded-chart transactions, ODP/ODT annotation owners, DOCX paragraph
collapse snapshots, PPTX shape classification, XLSB cell-watch snapshots,
and inert OLEDS object-link metadata in `litchi-ole-common`. Each owner keeps
semantic values, bounded codecs, package integration, validation, and focused
tests in nested modules; edits are clone-staged and source-checked, while
unknown XML, BIFF12 records, and OLE wire tails remain opaque and inactive.

This turn extends the same topology with typed OLE Document Summary
Information (`property_set::document_summary`) over the shared PIDDSI codec;
DOC `parts/annotation_bookmarks`; DOCX `section/footnote_columns`; PPT
`document_comparison`; PPTX `shape::designer` `p15:designElem` metadata; XLS
`picture_compression`; and ODP embedded-chart package transactions. OOXML
common now also owns bounded MCE `AlternateContent` choices, while
`litchi-drawingml::color` owns checked color choices and ordered transforms.
The top-level `litchi` facade exposes `docx`, `pptx`, `xlsx`, `xlsb`, `odp`,
`ods`, and `odt` directly (alongside `opc`, `ooxml_common`, and `odf_common`),
and `litchi-odf` remains a thin feature-gated detector/family umbrella. Every
new edit path retains opaque source material, uses bounded validation, and
publishes only source-checked snapshots or package transactions; no macro,
link, rendering, collaboration, or external-code behavior is activated.

This continuation completes another cross-format wave: shared OLE2 now owns
typed SummaryInformation metadata; ODT owns inert protection-policy snapshots
with opaque-settings preservation; ODS owns DataPilot package transactions;
DrawingML owns typed transform snapshots; DOC owns MsoEnvelope package edits;
DOCX footnote columns retain inherited namespace context and authored lexical
values; PPT document-comparison edits publish through the live OLE record; and
XLS chart-area edits patch only the fixed BIFF payload. These owners keep
semantic, codec, validation, package, transaction, and focused-test layers
nested by responsibility, with reserved wire bits and unknown XML/records
preserved through edits.

The DOC package now exposes the shared `litchi_doc::spaces` facade and
`Package::data_spaces` structural inspection for MS-OFFCRYPTO DataSpaces and
legacy-binary IRM graphs. The owner remains deliberately inert: it validates
transform/license topology, labels, integrity sidecars, and custom-XML
promotion markers without evaluating rights, decrypting streams, or contacting
external policy services.

The subsequent migration wave completes five more contextual owners. DOCX
numbering now exposes Word 2012 `restartNumberingAfterBreak` edits with
namespace-aware source preservation; ODS exposes cell-anchored annotation
snapshots and package transactions; XLSB exposes threaded comments, people,
mentions, and worksheet relationship edits; PPTX exposes inert non-Ink
`p:contentPart` payload inventories; and DOC, PPT, and XLS publish the shared
OLE Property Set editor through host-validated package snapshots. Each owner
keeps unknown XML, BIFF12 records, or CFB topology opaque, returns exact
no-op sources, and rejects stale or protected edits before publication.

The next continuation adds DOC subdocument package publication with checked
FIB/table-pointer relocation; PPT native diagram build transactions; XLS OLE
object metadata edits; OGraph series metadata transactions; DOCX OpenType
run-property extensions; and XLSB shared-workbook revision metadata. These
owners remain nested by semantic, wire, validation, package, and transaction
responsibility, preserve unknown source material, and keep referenced files,
SmartArt layout, collaboration, activation, and external execution inert.

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

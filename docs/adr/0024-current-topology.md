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

The following continuation adds source-checked BIFF8 Revision Log metadata
transactions, SpreadsheetML revision package snapshots and CRUD, legacy-PPT
diagram publication through the owning slide envelope, and PresentationML
media-track/caption/narration transactions. The XLSX revision owner validates
relationship and orphan topology before atomic publication; the PPTX owner
retains unknown extension XML and treats WebVTT/media targets as inert. No
owner activates collaboration, playback, SmartArt layout, external links, or
formula/revision replay.

The fourth continuation extends the same boundary to OLE2 and remaining OOXML
payload owners. Shared OLE smart tags and toolbar controls now have bounded
source-checked property-bag/control transactions; DOC embedded ObjectPool
entries, legacy-PPT external media, and BIFF8 external links publish inert
metadata and lifecycle edits with opaque payload preservation. DOCX settings
extensions, PPTX change metadata and ActiveX controls, and XLSX OLE objects
follow contextual `Snapshot`/`Transaction`/`Commit`/`Patch` facades with stale,
atomic, relationship/orphan, and MCE validation. These owners never activate
OLE, ActiveX, media, external links, macros, or collaboration behavior.

The fifth continuation completes additional external and embedded-object seams.
OGraph now publishes source-checked chart-package transactions with retained
CFB/compression envelopes; DOC smart-tag host tables and DOCX web settings gain
typed edits with FIB/OPC topology validation. XLSX and XLSB external links,
PPTX content parts, and slide-owned PPTX OLE objects expose the same layered
snapshot/transaction/patch boundary with opaque XML, BIFF12, MCE, and binary
payload preservation. All targets remain inert: no link refresh, DDE/OLE
activation, web fetch, content execution, rendering, or formula recalculation
is introduced.

The sixth continuation moves more shared semantic mutation below host facades.
DOCX mail-merge settings and recipient metadata, PPTX slide-show events, and
XLSX external connections now use source-checked typed transactions with
relationship and opaque-XML preservation. `litchi-drawingml::chart` owns the
host-neutral chart snapshot/editor boundary for typed series, axis, label, and
metadata edits; concrete packages retain only placement and relationship
ownership. None of these owners contacts a data source, replays an event,
calculates or renders a chart, or activates an external provider.

The seventh continuation extracts another shared package boundary and closes
more OOXML host seams. `litchi-ooxml-common::custom_xml` now owns bounded
source-checked Custom XML Data Storage item/properties CRUD; DOCX glossary
catalogs, PPTX notes graphs, and XLSB host connections publish contextual
transactions with opaque XML/BIFF12 retention and relationship validation.
These services remain inert and never retrieve schemas, render notes, contact
providers, execute add-ins, or refresh external data.

The eighth continuation extends the same OLE2/OOXML boundary with DOC tracked
revision snapshots, DOCX chart-graph and document-variable transactions, PPT
broadcast and terminal document-structure edits, XLS RTD topic transactions,
XLSB web-extension binding edits, XLSX XML-map edits, and PPTX structure,
guide, color-map, and custom-show inverse edits. Shared `litchi-ole-common`
now also owns CFB directory catalogs, Custom XML stores, and VBA-signature
metadata. Each owner is nested by semantic model, bounded codec, package
integration, validation, and source-checked `Snapshot`/`Transaction`/`Commit`/
`Patch` layers; opaque XML, BIFF, and CFB tails remain preserved and external
links, macros, add-ins, broadcasts, and formula/runtime behavior remain inert.

## IWA and iWork

The current shared physical substrate is split among `litchi-iwa-archive`
(ZIP/package preservation), `litchi-iwa-core` (Snappy/IWA framing and neutral
archive metadata), `litchi-iwa-detect`, `litchi-iwa-index`,
`litchi-iwa-graph`, `litchi-iwa-package`, `litchi-iwa-protos`,
`litchi-iwa-text`, and `litchi-iwa-text-wire`. `litchi-pages`,
`litchi-numbers`, and `litchi-keynote` are the concrete application package
owners. `litchi-numbers-wire` is a low-level BNC adapter excluded from the
supported format and root facades.

Historical 2026-08-06 snapshot (superseded by the 2026-08-23 current-topology
amendment, which records 64 workspace packages, 239 internal dependency
declarations, and 13 ordered migration debts):

`litchi-iwa` still exists only as the migration host for editors and
compatibility tests that have not reached the concrete packages. Its 17
internal workspace dependencies are all explicit ordered debt, with no
canonical edge. Physical preservation and comparison examples have moved to
`litchi-iwa-archive`. The root `litchi::iwork` coordinator now owns supported,
immutable cross-format reading from regular ZIP files, borrowed/shared bytes,
and frozen app-authored package directories, and publishes only root-owned
archive-free semantic views. Directory ingress is an index-only semantic
snapshot and is deliberately not exact-package or edit provenance. The host
structured path remains temporary migration debt for uncovered frozen logical-
entry coordination plus outstanding parity/property/fuzz execution and editor
ownership; it is not the supported root boundary.
[ADR 0028](0028-iwa-monolith-exit.md) is the authoritative exit gate.

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

### 2026-08-08 Keynote ordering continuation (historical snapshot; superseded by
the 2026-08-23 current-topology amendment)

The root prepared-source coordinator now also admits validated frozen logical
entries through an internal, semantic-only route; the preceding reference to
"uncovered frozen logical-entry coordination" is retained as historical state
and is superseded by this paragraph and ADR 0028's later amendment. That route
does not publish member names, storage builders, or edit provenance.

Within the concrete Keynote owner, the document-root and show topology use
narrow private Buffa lazy projections after format-owned wire preflight.
Ordered slide references are streamed from the embedded slide tree, while raw
source field records, including their encoded keys, encoded lengths, and
payloads, remain the preservation authority. The public structural writer is
selector-first: `Package::edit_slide_order()` directly returns an edit that
moves one selected slide to a checked final semantic position and produces a
separate reversible `SlideOrderPatch`. The skip-state transaction remains
source compatible.

The migration host's move method, focused example, and move-specific tests have
moved to `litchi-keynote`, but no host dependency edge is removed by that
vertical capability. At this historical snapshot, the boundary ledger
contained all 17 ordered `litchi-iwa` debts; later debt-retirement amendments
supersede that count. Remaining editor, Prost graph, example, test, fuzz,
durable-patch, and atomic-save ownership prevents host deletion.

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
[ADR 0023](0023-odf-family-crate-split.md), plus the current iWork exit in
[ADR 0028](0028-iwa-monolith-exit.md). The layered owner paths are
verified by the affected-crate all-target compile and boundary-policy check.

## 2026-08-08 focused settings and graph-boundary continuation

The preceding 17-debt snapshot is historical and is superseded here.
`litchi-keynote::Package` now owns a direct bounded `show_settings` reader and
exact-source `edit_show_settings` transaction. The reader validates the full
known Show/SlideTree envelope without initializing full slide semantics or
retaining slide identifiers. A changed exact source rewrites one owning
component and is fully reopened under retained options; a null show permits
only its exact semantic no-op. Raw source field records, not Buffa, remain the
unknown-content authority.

Changed legacy nested-`Index.zip` settings edits are still a host compatibility
capability because the ordinary focused edit must not silently normalize its
physical provenance. The host method, example, and compatibility tests remain
until that behavior gains an explicit preservation-safe owner. This slice is
therefore not complete show-settings host retirement.

The redundant direct `litchi-iwa -> litchi-iwa-graph` manifest edge is retired.
The host consumes graph identities and snapshots through `litchi-iwa-index`,
whose own graph dependency remains canonical. Debt identity 007 is deleted and
later identities are not renumbered. The current checker inventory is 63
workspace packages, 223 internal dependency declarations, and 16 ordered
migration debts.

## 2026-08-08 Pages section-name continuation

The concrete Pages owner now contains its first exact-source mutation family.
`Package::edit_section_name` selects an existing section by exact semantic name
or checked position, distinguishes absent and explicitly empty native names,
and returns a separate reversible exact-artifact patch. One selected section
payload and its owning IWA member form the mutation closure; complete raw field
records and object-header bytes remain the preservation authority, and the
candidate is fully reopened before publication.

The root facade exposes these canonical `litchi-pages` types directly. The
migration host's raw-ID rename example has moved to a semantic focused-crate
example, but its legacy-normalizing compatibility writer remains. This vertical
ownership move changes no manifest edge: the current inventory remains 63
packages, 223 internal declarations, and 16 ordered debts.

## 2026-08-08 current-status amendment: archive cache state and focused gates

The preceding 16-debt inventory is historical and is superseded by this
current-status amendment. `litchi-iwa-archive` now owns cache-backed
`PackageState` and its bounded physical parsed-component state. The
`litchi-iwa-cache` crate remains a dependency-free leaf; `litchi-iwa` retains
format and error policy. Direct host-to-cache debt identity 003 is retired
without renumbering later identities. The current checker inventory is 63
packages, 223 internal dependency declarations, and 15 ordered debts.

Numbers narrows one read seam only: `TableInfo.tableModel` uses a strict small
private Buffa projection after bounded raw preflight, rejects a zero reference,
does not encode or retain unknown data, and stores no repeated fields. Raw
source remains authoritative. The broader table model and Numbers graph have
not migrated by this change.

Pages 14.4 opened the focused clear and range outputs without repair:
`/private/tmp/litchi-pages-example.KdlErn/clear.pages`
(`63c2aa20f6064b9a8c5a536475d1a71b34175f4c6924a4d384f24c39fd5155e6`)
was visibly empty, and `range.pages`
(`dd0405249a56e3e2b535e6a9541f02feda6299ce1a0959f4d68f7e44a0ae307a`)
rendered exactly `Range prefix: Litchi native Pages fixture`, `Buffa lazy-view
migration verification`, and `2026-08-07`. Native Save As/close/reopen yielded
`clear-native-resaved-20260808.pages`
(`3ba278e1934688c653ab73f1ee2a194f670545dd160aa5d8e33c2054463a9676`)
and `range-native-resaved-20260808.pages`
(`74072d9d813282618db8e47f7ebc26cc59f7c17b1abf9d22c5bbf5473b942a9f`).
Focused semantic reread matched, and no-op and inverse output over each
native-resaved artifact remained byte-identical. This closes the clear/range
native gates, not the broader Pages or monolith-exit gates.

## 2026-08-08 current-status amendment: Keynote existing speaker notes

The concrete Keynote owner now provides selector-first speaker-notes reads and
checked UTF-16 text transactions for an existing slide-to-note-to-storage
graph. Its public facade contains archive-free semantic edit, commit, patch,
diagnostic, error, limit, position, and span types; native identities and
generated protobuf values remain private. The migration host's raw-ID example
is removed in favor of the focused format example, while its wider notes
compatibility and graph-creation/deletion surfaces remain.

Only the selected owner-reference projection uses a private Buffa lazy view.
Strict bounded wire preflight, package-wide ownership scans, exact note and
reference shapes, preserved raw records and IWA headers, one-component changed
reassembly, complete retained-limit reopening, and semantic/topology readback
form the publication boundary. This is not whole-Keynote-graph lazy decoding.

Keynote 14.4 opened the public example's set, range, and clear artifacts
without repair, saved each natively, closed it, and reopened the exact saved
path. Focused reread matched the requested Unicode values or empty note, exact
no-ops over all three native-resaved files were byte-identical, and temporary
edits inverted to each native hash exactly.

Development-only internal edges are now classified and stale-checked as such;
two redundant Numbers/Pages test-only ZIP edges were removed. The checker
reports 63 workspace packages, 221 internal declarations, and 15 ordered
migration debts. No host debt is retired by the notes transfer, so durable
patches, atomic publication, aggregate memory policy, fuzz/sanitizer gates,
remaining format ownership, and deletion of `litchi-iwa` are still required.

## 2026-08-08 current-status amendment: aggregate contracts stay owner-side

The dependency graph is unchanged at 63 workspace packages, 221 internal
dependency declarations, and 15 ordered migration debts. The neutral
`litchi-iwa-structured` boundary now enforces complete retained-owned-text
accounting for Pages and Keynote, while `litchi-pages` owns the distinction
between retained semantic UTF-8 and synthesized rendered separators. Exact
Pages observations are preserved by the root facade instead of reconstructed
there. These are contract hardenings within the existing owners and add no
manifest edge.

Debt 011 and the host structured adapter remain deliberately present. Five
Numbers compatibility cases still live only in the migration host: detached
models, type-9 numeric cells, package-global ordering, canonical/legacy model
precedence with deduplication, and exact/over table limits. Removing the edge
before those oracles move would erase migration evidence rather than complete
ownership transfer.

## 2026-08-08 current-status amendment: neutral aggregate without a host adapter

The preceding debt-011 status is historical. `litchi-numbers` now owns both
the rooted projection and the explicitly allocating package-global
compatibility projection. Focused and root tests own the five migration
oracles: detached models, exact finite type-9 values, global order, canonical
type-6001 precedence over legacy type-6000 with deduplication, and exact versus
exceeded table limits. The root facade composes concrete owner results into the
neutral `litchi-iwa-structured` model; the neutral crate remains the aggregate
model and budget owner.

`litchi-iwa` no longer publishes or implements `StructuredData` or
`extract_structured_data`, and it no longer depends on
`litchi-iwa-structured`. The host adapter, its tests, and support-only Numbers
hooks are gone. Debt identity 011 is retired without renumbering later
identities. The current checked topology is 64 workspace packages, 238
internal dependency declarations, and 14 ordered migration debts.

The host is still required for the remaining compatibility and edit surfaces.
Focused eager Prost payload paths, full Buffa migration, root preparation's
unrelated-sidecar peak memory, durable publication, and the final host removal
remain outside this cutover.

## 2026-08-09 current-status amendment: Keynote title/body text owner

The concrete Keynote package now owns reads and checked UTF-16 transactions for
text in an existing slide's existing title or body placeholder. Its public API
is role-aware and selector-first, distinguishes absent placeholders from empty
storages, and exposes no native identifier, component name, generated message,
or raw record. A changed edit commit proves exclusive role-correct ownership,
rewrites the selected text storage and invalidates the selected slide node's
rendered-thumbnail state, including its preview object references and
preview-owned selected-message aggregate/field data-reference occurrences.
Proven unrelated data references remain exact; ambiguous aggregate-only
ownership fails closed. These occupy one or two
distinct IWA components; diagnostics report that unique component count. The
commit also deletes any
root `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg` through the
archive owner's bounded deletion-aware reassembly path. Preview entry deletions
are not counted as touched IWA components. All mutations publish atomically in
one candidate. Retained ZIP entries and all IWA
objects other than the selected storage and slide node remain the preservation
set.

The candidate fully reopens under the retained limits and checks selected
text, slide-node invalidation, root-preview absence, physical object
preservation, and unselected slide semantics. Applying a changed patch reopens
and verifies its stored target bytes but does not reassemble them. An exact
no-op relies on the immutable selected snapshot established when the edit
started, leaves previews and caches untouched, reports zero components, skips
whole-source validation and candidate reparse, and shares the source allocation.
A changed inverse patch is exact-source checked and restores the
complete original artifact, including preview/cache state, after reopen and
verification. Slides with native field-37/38 cached title/body strings remain
fail-closed because this vertical does not yet own their rewrite.

The exact format-ownership seam reuses the private speaker-notes Buffa view for
`KN.SlideArchive` title/body fields 5 and 6 and adds a private placeholder view
for the required placeholder/shape inheritance chain, optional owned-storage
field 4, and placeholder kind. The selected read forces the slide view, while
package-wide proof raw-scans every slide and note candidate. A slide candidate
is forced through the slide view only when its raw edge matches the selected
placeholder. Placeholder candidates are likewise raw-scanned before the
relevant owner is forced through the placeholder view.
The scanner also polices deprecated storage, text-flow, standalone shape-info,
and embedded reference aliases. It does not force the Buffa `NoteArchive`
view. Text decoding and splicing continue through `litchi-iwa-text-wire`. Raw
source remains the unknown-content and rewrite authority.

The migration host's nine title/body/notes set, replace, and clear methods are
removed. This is intentionally breaking: callers replace raw indices and
mutable editor calls with `SlideSelector`, checked UTF-16 spans, and focused
`SlideTextEdit` or `SlideNotesEdit` commits. It is not a source- or
behavior-compatible alias for malformed, ambiguous, or shared graphs that the
focused owners reject.

The cache-hardened sequential output has SHA-256
`f3b13cd5bd614d93493cc6780ff177e6a203d990d15b9d5c592687ef40a48263`.
Apple Keynote opened it without repair, displayed both requested Unicode values
and the untouched date, regenerated all three root previews on native Save As,
and completed close/reopen without warning. The native copy has SHA-256
`cb3f9b05613505bb422942ca43e237a731454f58753ee65f26ae639187b96a6c`;
focused reread matched, and both same-value edits were exact zero-component
no-ops. ADR 0008 records the complete gate and inverse hashes.

No title/body placeholder creation or deletion, arbitrary text-box editing,
durable patch serialization, atomic filesystem publication, whole-Keynote
Buffa conversion, or host deletion follows from this transfer. The current
metadata/policy inventory is 64 workspace packages, 235 internal declarations,
and 14 ordered migration debts.

## 2026-08-10 current-status amendment: Numbers table-lock owner

`litchi-numbers::Package` now owns effective lock reads and immutable
exact-source transactions for an existing attached table selected by semantic
sheet and table selectors. `litchi-numbers::table::lock::State` re-exports the
canonical archive-free iWork table-lock value. Edit, commit, reversible patch,
diagnostics, error, and limit types remain format-owned and expose no native
identifiers or protobuf state.

The package maps the selected semantic position back through the rooted native
sheet/drawable topology and requires one unambiguous canonical or legacy
`TableInfo` payload. A strict bounded raw-wire pass owns field presence and the
canonical optional drawable `locked` Boolean. The private Buffa lazy views for
both the required drawable `super.locked` value and nonzero table-model
ownership reference are forced, and their complete snapshot is checked against
preflight. Raw source, rather than Buffa, remains the rewrite and
unknown-content authority; table-model, tile, sidecar, and formula payloads
remain on their existing paths.

This supersedes the old two-message/opaque-super/64 KiB TableInfo description:
the current three-message TableInfo/Drawable/Reference projection forces both
lock and model lazy values and measures 83,529 generated bytes under 84 KiB.

Exact no-ops share and preserve the source, including absent versus explicit
false lock encodings. A changed commit patches one nested scalar, rewrites one
IWA component, reassembles the flat package under retained limits, completely
reopens it, and verifies the requested state. Competing rooted sheet ownership,
contradictory selected-owner metadata, noncanonical object-length prefixes, and
selected merge/diff metadata fail closed on a changed edit. Detached/unrooted
pseudo-sheet and view-state dependent references remain opaque and preserved.
Changed patch application
reopens its already-stored exact target instead of reassembling it; inverse
application restores the complete original artifact. Legacy nested packages
remain readable and support exact no-ops, but refuse changed publication.
Typed diagnostics distinguish zero-component/no-reparse no-ops from
one-component/full-reparse changes.

The Numbers host read and mutation seam is retired in full: its direct
getter/setter, private selector context, `NumbersTableInfo.lock_state` and its
field-population branch inside `tables()`, model-specific shared read/write
helpers, and Numbers-only model-ID
matching branch are gone. Numbers readback now uses the focused package API.
The boundary ratchet covers five exact functions under a three-host plus
two-shared scope and separately rejects `NumbersTableInfo.lock_state`. Pages
and Keynote retain the generic shared table-lock
getter/setter and codec. This is therefore one Numbers read/mutation owner
transfer, not removal of the shared codec, a dependency edge, or the monolith.
The current source inventory has two
semantic-state tests, nine codec tests, and 15 transaction tests, including a
checked-in native-fixture case and rooted `FormBasedSheet` field path `[1, 2]`;
the focused transaction suite passed 15/15. Flat legacy type-6003 TableInfo
change/inverse and partial-sink write accounting are included.
The bounded `numbers_table_lock` fuzz target compiles, and all 57 boundary
policy regressions pass. The full policy command still reports the 14
pre-existing soapberry-zip/xml-minifier annotations. A Numbers-only fuzz
package and sustained sanitizer execution remain open.

Apple Numbers 14.4 (7043.0.93) completed the current-writer gate without
warning: the source, Rust-locked, and native-resaved SHA-256 values are
respectively
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`,
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`,
and `8aa87a3afcb145b66c5c6f4e10645cd1cf658f4b65f0976612ac6d62d4652995`.
Numbers showed the table locked with disabled cells, retained the B2 marker and
B3 value 42, saved, closed, and reopened it; focused reread remained locked and
an equal-state transaction preserved the native-resaved hash exactly. The
inverse restored the exact source hash.

Open topology work remains: no aggregate transaction peak-memory or total-work
contract covers all retained artifacts, buffers, traversals, hashing, and
candidate reopen; a complete fallible-allocation audit remains unfinished; and
the library has exact `write_to` with partial-byte failure accounting but no
atomic durable filesystem-save owner.
The process-local patch has no versioned semantic operation envelope,
read/write sets, composition, three-way merge, or bounded history.
Resource/allocation errors do not yet include the selected semantic table
path, and exact source bytes remain ordinary `Package` surface rather than an
explicit advanced/raw API.
The flattened `TableLock*` transaction names remain migration debt against the
focused-module short-name rule.
The archive-free `Table` snapshot does not yet carry lock state, remaining host
table/cell mutations do not enforce that state by default, and the private
Numbers locator has not converged on the neutral IWA index owner.

## 2026-08-10 current-status amendment: Pages page-layout owner

`litchi-pages::Package` now owns reading and immutable exact-source mutation of
the document-wide, presence-preserving `page_layout::Layout`. The focused API
is `page_layout`, `edit_page_layout`, and `apply_page_layout`, plus
format-owned edit, commit, reversible patch, diagnostics, error, and limit
types. No physical identifier, component, generated message, or wire-field
vocabulary crosses the public boundary.

The private locator requires one object 1 and one type-10000
`TP.DocumentArchive`. A bounded canonical raw pass reads required opaque
`super` field 15 and scalar fields 30 through 39 and 42; every projected layout
scalar is then forced on the existing document-body Buffa lazy view and the
complete result is cross-checked. The projection remains read-only, repeated-
view-free, and preservation-free, with five generated files measuring 122,114
bytes under 124 KiB. Raw records remain authoritative for unknown fields and
rewrites.

A changed transaction patches the selected document and follows a raw rooted
cache graph: required `TP.DocumentArchive.super` field 15 to
`TSA.DocumentArchive.view_state` field 5, its unique type-210 object to field 1,
and that reference to the unique type-10147 view-state root. Deprecated
document fields 11 and 12 are rejected. Both followed local edges require one
aggregate metadata occurrence and optional unique field metadata at `[15, 5]`
and `[1]`, respectively. The transaction removes the rooted
layout-state field 1 and its uniquely proven aggregate and optional path-`[1]`
reference metadata, while preserving UI-state field 2, unrelated metadata,
unknown fields, the now-detached opaque layout-state object, the intermediate
bridge, and detached/unrooted view-state candidates. Missing, ambiguous, or
contradictory rooted objects or metadata, a layout/UI alias, selected
merge/diff state, and noncanonical object lengths fail closed. The document and
rooted view-state root can share one component or occupy two; diagnostics
report that exact one-or-two component count.

The same atomic reassembly deletes root `preview.jpg`, `preview-micro.jpg`, and
`preview-web.jpg`, reported separately from components. The complete candidate
is reopened under retained limits and checked for requested layout, absent
cache edge and previews, stable statistics, and unchanged section semantics.
Canonical unknown protobuf groups remain readable and exact on no-op paths,
but changed layout splicing currently fails closed on a group-bearing document
payload.
Exact semantic no-ops preserve cache and preview bytes, share the source,
report zero components, and skip reassembly and reopen. Changed patch
application reopens its exact stored target; inverse application restores the
whole source artifact. Legacy nested packages retain reads and exact no-ops but
refuse changed publication.

The migration host no longer has `PagesEditor::page_layout`,
`set_page_layout`, the private page-layout module/source, or the old host
example. A focused example demonstrates validated immutable chaining,
no-clobber temporary publication, and optional exact inverse output. Boundary
ratchets protect both the retirement and the archive-free facade. Remaining
host Pages editors and other settings/cache owners are unchanged, so no
manifest edge or ordered migration debt is retired.
The current inventory remains 64 workspace packages, 235 internal dependency
declarations, and 14 ordered debts.

Verification is current: all 92 Pages tests/doctests pass, including 10/10
focused transactions, as do 6/6 private codec tests, the Pages package check,
focused warnings-denied Clippy, and all 63 boundary-policy tests. The fuzz
target compiles and completed 32 generated smoke inputs plus a fixed changed
corpus; sanitizer execution remains pending because the installed stable
toolchain cannot run cargo-fuzz's sanitizer flags and nightly is unavailable.
On the checked-in native Pages fixture, the 792 by 612 point landscape
candidate touched two components, deleted three previews, retained semantic
text, and inverted exactly. Source and candidate SHA-256 values are
`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`
and `79e00545ef6e2e30e366e3160b7d9126bf06cffac5fbbd5551e3d3789cc298e4`.
Apple Pages 14.4 (7043.0.93) opened the candidate without warning, showed US Letter
landscape with Document Body and all three fixture lines intact, then completed
native Save As, close, and reopen. Save As regenerated all three previews and
produced SHA-256
`8228e7518bb080bd8e5ec134d0abc7484c8825ad3cde3d16cabf76c5dbd8ef82`;
a focused equal-layout transaction reproduced that artifact exactly with zero
components and preview deletions.

Open topology debt includes the unowned opaque layout-state object, other
render/settings caches, whole-Pages Buffa coverage, aggregate transaction
peak-memory and total-work accounting, a complete fallible-allocation proof,
durable patch serialization, and a library-owned atomic durable filesystem
save. Exact bytes remain ordinary `Package` surface, and flattened
`PageLayout*` transaction names remain focused-module naming debt.

## 2026-08-10 current-status amendment: combined Pages document settings

`litchi-pages::document_settings` now owns an archive-free composite
`Settings`, formed from `document_options::Options` and `footnote::Settings`,
with canonical short `Edit`, `Commit`, `Patch`, `Diagnostics`, `Error`, and
`LimitKind` transaction names. The new focused
`Package::{document_settings, edit_document_settings,
apply_document_settings}` method and type signatures expose no native
identifiers, source bytes, archive/IWA types, Prost messages, Buffa views, or
generated types.

The private owner is the unique rooted `TP.DocumentArchive.settings` reference
at field 7 to a unique local type-10012 `TP.SettingsArchive`. The locator
requires the nonzero local reference exactly once in aggregate metadata and
accepts only optional unique matching field metadata at path `[7]`. A strict
raw preflight and forced Buffa lazy projection agree on SettingsArchive fields
1/2/3/9/10/30-34: body, headers, footers, hyphenation, ligatures, footnote
kind/format/numbering/gap, and facing pages. The five generated files total
174,682 bytes under the 176 KiB limit; their deterministic aggregate SHA-256
is `7618a60db84b87e28eea67a8acd85ce8eb19513cf4cee7654c1c4e78f405f824`.
The projection has neither a repeated view nor a production encoder; raw
records retain rewrite and preservation authority.

Exact semantic no-ops share the source and skip reassembly, reopen, cache
traversal, and preview deletion. A changed edit rewrites the selected settings
component, invalidates the rooted document view-state cache chain, and
atomically deletes root `preview.jpg`, `preview-micro.jpg`, and
`preview-web.jpg`; the settings and cache roots can occupy one or two IWA
components, reported separately from the three deleted previews. Reopen checks
the requested settings, cache/previews, statistics, and preserved semantics.
Changed patch application reopens its exact stored target, conflicts reject,
and the inverse restores the exact source artifact.

Canonical unknown scalar fields are preserved. Bounded canonical groups are
readable and exact on no-op paths, but a changed splice of group-bearing
settings fails closed. Noncanonical or wrong-wire encodings, duplicates,
invalid booleans/int32/references, contradictory rooted ownership metadata,
merge/diff state, and malformed object framing are rejected. Legacy nested
`Index.zip` remains readable and byte-exact for no-ops, but changed edits now
return `UnsupportedSource`; this intentionally removes the former host's
changed normalization behavior.

The migration deleted `PagesEditor::{document_options,
set_document_options, footnote_settings, set_footnote_settings}`, the three
private host sources `document_options.rs`, `document_options/wire.rs`, and
`footnote_settings.rs`, and two old host examples with their duplicate tests.
One focused example now owns read/edit/apply, immutable chaining, synced
no-clobber publication, and optional inverse output. The combined boundary
ratchet passes 70/70 tests. The live repository boundary command still reports
14 unrelated pre-existing findings: 12 for six `soapberry-zip` dev-only edges
and two for `xml-minifier` normal edges. The workspace inventory remains 64
packages, 235 internal dependency declarations, and 14 ordered debts.

Verification passes all 108 Pages tests/doctests, including 14/14 focused
transactions, 4/4 codec cases, and 6/6 facade cases. Package check,
documentation, and no-dependencies warnings-denied Clippy are green. The fuzz
target compiles and its no-op/changed smoke cases pass; full sanitizer runtime
remains unavailable because the installed stable toolchain rejects the
required flags and nightly is unavailable.

Apple Pages 14.4 (7043.0.93) supplied a fresh footnote-bearing seed with
SHA-256 `9da01e2805459e05450551827140069eefe8049aeeacc7625d3c62d7e00ffeab`.
The Rust candidate, SHA-256
`3d052e7f1ec86e57ea0553e46f628de1d9fa5bdda615ded9410fca29c93f0995`,
reported changed, two touched components, and three deleted previews; its
inverse restored the exact seed. Pages opened it without warning and showed
body/header/footer and facing pages enabled, hyphenation and ligatures
disabled, Roman footnotes restarting each page with an 18-point gap, and all
three body markers plus the note intact. Save As, close, and reopen preserved
those semantics, regenerated the previews, and produced SHA-256
`803167e2479c459f9a33c8ecfc4d713f596fdc5d5d337090ab3c90e467a0cba6`.
A focused equal-settings transaction on that native resave reported zero
components and preview deletions and was byte-exact; its inverse was exact too.

Remaining shared debt includes aggregate transaction peak-memory and total-work
accounting, the retained infallible `ArchiveInfo` clone in the shared archive
encoder, a complete fallible-allocation proof, group-aware changed splicing,
exact streaming/partial-output accounting and a robust Pages `Package::write_to`,
library-owned atomic durable filesystem replacement, and versioned deterministic
patch serialization with semantic operations, read/write sets, composition,
merge, and history. Exact source bytes remain ordinary `Package` surface;
opaque cache objects and other Pages settings/render state remain unowned.

## 2026-08-10 current-status amendment: hardened Keynote show settings

The earlier Keynote show-settings section is superseded. The archive-free
semantic and transaction family is now canonically grouped as
`show::{Settings, Edit, Patch, Commit, Diagnostics, Error, LimitKind}`.
`Package::{show_settings, edit_show_settings, apply_show_settings}` exposes no
native or raw-source values in its focused signatures; consuming `Edit::set`
makes immutable chaining explicit. Exact output is streamed through
`Package::write_to`, which keeps the retained source private and reports a
precise sink offset on failure without allocating another package-sized
buffer.

The private locator selects the unique root `Document.iwa`, object 1, and
`KN.DocumentArchive`, then follows required local show reference field 2. A
nonzero selected identifier must occur exactly once in aggregate metadata;
optional field metadata must match unique path `[2]` and cannot assign the
selected identifier elsewhere. It resolves to one object in one component
with one `KN.ShowArchive` message. A null show remains readable as default
settings and supports only an exact no-op.

Strict raw and forced Buffa lazy passes cross-check both hops. The root's five
generated files measure 58,630 bytes under 60 KiB with aggregate SHA-256
`7918aad2578cf3bd07eb0be36f2e31d11f93391584308c1e4adc1fd86ed065fd`.
The Show/SlideTree projection validates all known reference/size/settings
fields and the slide ceiling without retaining the hand-routed repeated slide
list; its five files measure 138,661 bytes under 140 KiB with aggregate
SHA-256
`747fe9f99dc5bb1855aae1bfcb16065a5fe6305bdbf8730a21ef24bb75e915ee`.
Both generated surfaces are repeated-view-free and encoder-free. Raw records,
not Buffa, own preservation and rewriting.

Mutation adds canonical selected-component framing and rejects selected
`should_merge`, base-message, and all diff/merge metadata. A changed edit
raw-splices only size and eight optional settings fields in the selected Show,
rewrites one component, then fully reopens and verifies the candidate. Size or
slide-number-visibility changes delete the existing zero-to-three root
previews; playback-only changes preserve them exactly. Every slide component
and slide-node thumbnail/playback cache remains exact in both cases.
Diagnostics separate the one component from preview deletions.

No-ops share the exact source and skip cache inspection, reassembly, and
reopen. Changed patch application authorizes exact bytes and reopens its stored
target; the inverse restores the entire source artifact. Legacy nested
`Index.zip` reads and exact no-ops remain supported, while changed edits now
return `show::Error::UnsupportedSource` rather than invoking the retired
normalizing host behavior.

`KeynoteEditor::{show_settings, set_show_settings}`, its module/source,
`examples/edit_keynote_show.rs`, and direct editor mutation tests are deleted.
The focused example owns consuming semantic staging, inverse verification,
no-clobber temporary publication, and `write_to`. The host read-only
`KeynoteDocument::show` still Prost-decodes `KN.ShowArchive`; this is direct
editor-mutation retirement, not complete host/native Show retirement. No
manifest edge or ordered debt changes.

Current evidence passes 19/19 focused transactions, 106/106 complete codec
tests, 49/49 focused Keynote codec tests, Keynote all-target checking,
`litchi-iwa` library checking, umbrella facade compilation, strict rustdoc,
and 80/80 boundary tests. Focused live retirement/leak audits are empty; the
general boundary command retains 14 unrelated pre-existing diagnostics.
The fuzz target passes `cargo check`, and its stable-built executable completed
32 bounded cases with expected missing-sanitizer-symbol warnings; cargo-fuzz
sanitizer execution still requires unavailable nightly.

Apple Keynote 14.4 (7043.0.93) opened and auto-played both final Rust candidates
without repair/recovery/conversion. From source
`f3adcde9315b6df580805bcb63c995cc1e1ef569a4befa06a102485e13c883b2`,
the slide-number candidate/resave hashes were
`6d28d461c1203f00384fe6a758df1f903c7555b90ff02d2dc32d856aa9056c13`
and `031a701040ed1ea9a5111fe3e298bcddcf33d498891f827b703d01328ba17224`;
the 1280-by-720 candidate/resave hashes were
`67e9ff0557683af105dfe57f999acabcde23f121f7aebb06102c93e03121c027`
and `a3a2f6e072db4bd952f2c02e528f25c3656dba5810fbff75e93b5a699aac0eda`.
Both inverses restored the source exactly. Inspector, Save As, close, and
exact-path reopen retained Self-Playing, Loop, Play on Open, five-/two-second
delays, and the respective 1920-by-1080 Widescreen/1280-by-720 Custom sizes.
Rust deleted three root previews and Keynote regenerated them. All four
`Index/Slide*.iwa` hashes stayed exact from each Rust candidate through native
resave.

Keynote normalized explicit slide-number true to absence. Restaging absence is
an exact `031a7010...` no-op; restaging true changes it. The native size
artifact's same-settings no-op and inverse remain exact at `a3a2f6e0...`.
Thus the native evidence supports conservative preview invalidation and exact
slide-cache preservation, not persistence of the slide-number scalar.

Open debt includes the host Prost Show reader and other generated graph
consumers, aggregate peak-memory/total-work accounting, a complete fallible-
allocation proof, group-aware changed splicing, durable/versioned semantic
patches with read/write sets and composition/merge/history, and a
library-owned atomic durable filesystem save. `write_to` neither flushes nor
syncs or renames a destination. A full sanitizer-backed fuzz campaign remains
a verification gate.

## 2026-08-10 current-status amendment: Numbers names owner

`litchi-numbers::names` now owns atomic sheet/table renaming with canonical
short `Edit`, `Patch`, `Commit`, `Diagnostics`, `Error`, and `LimitKind` types.
The root `litchi::numbers::names` facade preserves that nesting and forbids flat
aliases. `Package::edit_names` is an infallible `O(1)` empty batch;
`rename_sheet` and `rename_table` consume the edit and resolve selectors against
the immutable base, while `apply_names` owns exact patch application. Focused
signatures expose no physical vocabulary. Source bytes are crate-private and
`Package::write_to` is the exact streaming output seam.

Changed ownership follows root document field 1 to the ordered local
Sheet/FormBasedSheet objects, then sheet drawable path `[2]` or form path
`[1, 2]` to TableInfo and its field 2 to TableModel. Each local edge requires
one aggregate metadata occurrence and optional unique matching field metadata;
selected table models require exactly one rooted TableInfo owner. Strict raw
preflight and forced Buffa views cross-check ordinary/nested sheet names and
TableModel identity/name. The generated projection is five files/82,641 bytes,
aggregate SHA-256
`944b7637fd6bf0eb895174b1e9229aa9eb9c393e05c666a86dd2843792eefe3e`.
Raw records retain preservation authority.

Final-state validation makes a batch atomic: sheet names are workbook-unique,
table names are sheet-local, swaps/collision-away work, and repeated targets or
final collisions fail before mutation. Changed table renames refuse a locked
selected table, any rooted pivot owner, and rooted nonempty volatile name-cell
dependencies; sheet-only rename remains valid with an unselected locked table.
The conservative native Θ(T²) pivot scan is charged against a preflight work
limit before native work. Each touched component is rewritten once and the
candidate is fully reopened and locality-checked.

Every changed batch deletes the existing zero-to-three root previews while
preserving `Index/ViewState.iwa` and every unrelated record/object/message
exactly. No-ops share the source and bypass changed guards/reassembly/reopen.
Changed apply reopens its exact retained target; inverse restores the entire
source and previews. Canonical/form and accepted legacy TableInfo/TableModel
variants work when unambiguous. Nested legacy packages read/no-op exactly but
refuse changed publication as `UnsupportedSource`.

Host `NumbersEditor::{rename_sheet, rename_table}`, direct host tests, and
`examples/rename_numbers_items.rs` are retired; the focused example owns
semantic batching, inverse checking, `write_to`, and synced no-clobber output.
The private `rename_attached_table_in_package` remains for Numbers sheet
duplication, and its `rename_table_in_package` wrapper remains for
Pages/Keynote attached tables. No crate edge is removed, so ordered debt 015
(`litchi-iwa -> litchi-numbers`) remains. Current inventory is unchanged at 64
packages, 235 internal dependency declarations, and 14 ordered debts.

Verification is green: 10/10 focused, 105/105 library, 1/1 facade with
`--features numbers`, 89/89 boundary regressions, both live focused audits,
`litchi-numbers --all-targets`, `litchi-iwa --lib`, and rustdoc. Host
`litchi-iwa --all-targets` is not claimed because unrelated examples remain
red. The stable fuzz build completed an eight-case bounded control-flow smoke
with expected missing sanitizer symbols; it was not ASan.

Apple Numbers 14.4 (7043.0.93) accepted source
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
and Rust candidate
`22f8bc21223317318ec23ec764b8998af77a2c7800c68cbe88351abdb26b6e56`
without warning/repair/conversion; inverse restored the source. It displayed
sheet `Líneas 你好 🧪`, table `表 Café №42`, exact B2 marker and B3=42, with the
ordinary table selectable/editable. Save As/close/exact-path reopen produced
`e1803b0568454a345f7962c5b4c72e8cb3d78adb2c87d5db1e6c58288a9413c4`,
regenerated three previews, and retained the values. Equal restage/no-op/inverse
were exact at that hash.

The separate lock oracle
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`
showed `Locked`/`Locked items cannot be edited`, disabled cells, enabled Unlock,
and no title change from the Edit action. It supports the focused locked-table
refusal and sheet-only exception.

Open debt remains the bounded native Θ(T²) dependency scan, aggregate
peak-memory/total-work accounting, complete fallible-allocation proof,
process-local full-artifact patch storage and missing versioned semantic patch
operations/composition/merge/history, library-owned durable atomic save, and a
sanitizer-backed fuzz campaign. `write_to` does not flush/sync/rename.

## 2026-08-10 current-status amendment: Keynote transition mutation

Keynote slide transitions now use canonical nested
`transition::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` transaction
types through selector-first package read/edit/apply methods. Focused public
signatures expose no physical/native vocabulary or source bytes; exact output
uses `Package::write_to`.

Changed ownership follows the rooted Show/SlideTree `[3, 2]` reference to one
selected SlideNode, then its required local field 2 to one SlideArchive. Both
edges require unique resolution, exact aggregate metadata, and optional unique
matching field-path evidence. Strict transition and node-marker projections
must agree with the semantic record. Changed-only canonical framing and
merge/base/diff guards protect every selected message/component.

Rooted uniqueness walks the Show's slide-node list once and resolves nodes
through the package's sorted object index, yielding `O(slides log objects)`
lookup cost. Aggregate node-message and reference-payload bytes share the
`LimitKind::WireWork` charge rather than a per-node reset.

Strict raw preflight precedes a private five-message Buffa lazy-view
cross-check. The 2,347-byte derived schema is provenance-checked against KN,
has no repeated projection or production encoder, and generates five
files/208,052 bytes under 224 KiB. The validated raw records, not generated
views, retain exact preservation and splice authority. One aggregate field
counter and one strict-plus-Buffa work counter cover the selected SlideArchive,
transition, attributes, and animation envelopes; nested envelopes cannot reset
those resource ceilings.

The mutation closure is SlideArchive transition field 4 plus SlideNode marker
field 7 only when effect presence changes. The owners may share one component
or occupy two; every touched component is rewritten once, followed by full
reopen and exact locality verification. Everything unselected, including
unknowns, the three root previews, `Index/ViewState.iwa`, and slide/node
playback caches, remains byte-exact; playback-only transition edits do not use
root-preview deletion. No-ops share the source. Clearing an
already absent transition is an idempotent exact no-op; changed legacy nested
sources return `UnsupportedSource`. Exact apply reopens the stored target and
inverse restores the source.

Host methods `slide_transition`, `set_slide_transition`, and
`clear_slide_transition`, the `transition_lifecycle` module/source, the three
clear/edit/set-effect examples, and five whole direct mutation tests are
retired. The exact host scope change is +120/-998 lines, net -878. The focused
edit example becomes the mutation owner. `KeynoteSlideInfo.transition` and
host slide readers remain; `transition_wire.rs` is retained for
`KeynoteEditor::slides()` aggregate decoding and no-op validation, while
creation uses the separate `creation.rs::transition()` helper and retained
creation example. This is not complete host transition deletion.

No edge/debt changes: debt 014 remains and inventory stays 64 packages, 235
internal dependency declarations, 14 `litchi-iwa` dependency declarations,
and 14 ordered debts.

The deterministic gate passes 8/8 focused transition tests, 79/79 Keynote
library tests, 6/6 warning-denied doctests, 7/7 facade tests with
`--features keynote`, 6/6 codec tests, and retained host conversion/reader
tests at 3/3 and 7/7. Common infrastructure passes 10/10 focused and 140/140
full tests plus strict library Clippy; archive exact-artifact coverage reports
79 unit and 2 integration tests. `cargo check -p litchi-keynote --all-targets`,
`cargo check -p litchi-iwa --lib`, host no-run, formatting, diff checks, and
101/101 boundary regressions pass. All fuzz bins check; generated no-op,
fixed-clear, and fixed-set stable smokes completed six bounded cases each, with
expected missing-sanitizer-symbol
warnings and therefore no ASan claim.

Apple Keynote 14.4 (7043.0.93) opened disposable copies without warning,
repair, recovery, or conversion. Source SHA-256 was
`ab186d8d59c858e1b3c2596fd45463cec75ddd92e9fda9032da656a940e68dca`;
pristine Magic Move and clear candidates were
`d5d24386cb544374f4c26da4349f7be961be34180a4536578616886a56af8c1a`
and `5235a3d03dbabced6d06a03b4873826da8602d97f478c61f6467b35d732a08e5`,
and each inverse restored the source exactly. Magic Move showed 2 seconds,
Automatic, and 2.25 seconds of delay; clear showed No Transition Effect while
retaining Automatic and 2.25 seconds. Both states survived Save As,
close, and exact-path reopen.

The native resave hashes were
`dda5049cf431b5c88ea0a9fb209c67edc0d7f0764c23a17eb4e9fdf947d786f6`
for Magic Move and
`784069ca8bd2729829bcf204cccdced93f7fbea2b5f8c6b3e4965b47ef423e94`
for clear. Equal restaging on each native artifact reported
`changed=false`/`touched_components=0`; output, comparison, and no-op inverse
were exact at that native hash. Remaining shared debt covers aggregate
peak-memory/work and complete fallible-allocation accounting, process-local
complete-artifact patches without stable semantic serialization/read-write
sets/composition/merge/history, durable atomic library publication, and a
sanitizer-backed fuzz campaign.

## 2026-08-10 current-status amendment: Numbers table headers

The existing archive-free `table::headers::{Count, Settings}` remains the
semantic owner; this work adds nested
`table::headers::transaction::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind, Path, InvalidReason}` types rather than duplicating header settings.
`Package::{table_header_settings, edit_table_headers, apply_table_headers}`
uses an explicit sheet selector plus sheet-scoped table selector and keeps
native IDs, components, generated/wire types, and raw artifacts out of the
focused signatures. `Edit::settings` borrows the staged state;
`Edit::set(self, Settings) -> Self` is consuming and infallible. Exact output
uses `write_to`.

Changed ownership follows the rooted Document-to-Sheet/FormBasedSheet-to-
TableInfo-to-TableModel chain, including sheet path `[2]`, form path `[1, 2]`,
and TableInfo field 2. Unique resolution, exact aggregate reference metadata,
optional unique matching field evidence, and one rooted selected TableInfo
owner are required. Competing rooted ownership and selected metadata
contradictions fail closed; detached/unrooted references remain opaque.

Changed edits refuse an interactively locked selected table. Present counts
remain in `1..=5`; header rows plus footer rows must fit declared rows and
header columns must fit declared columns. Optional count/Boolean presence in
TableModel fields 9/10/11/12/13/29/32 is semantic state, not a default to
normalize. Selected raw framing and bounded work must be validated before the
candidate is rewritten.

Strict raw/Buffa cross-checking uses five generated files/51,480 bytes, no
repeated views, with SHA-256
`5a94caa4620c56bb464792084c01325cef01744bebac97ef948466b9dea105dd`;
raw records remain authoritative.

Field-85 pivot state blocks any change. Fields 81/84/86 or nonempty 83 block
header counts; active field-81/83/86 category/group state also blocks section
counts. Strict TableInfo role aliases 4/5/7/8/15/16/17 gate counts according
to header versus section role; rooted HeaderNameMgr gates header counts, and
deprecated sheet field 4 gates repetition. These are typed
`UnsupportedDependency` refusals. Footer/freeze/repeat and dependency-free
counts remain supported; admitted locality is not a general TableModel-only
count-parity claim.

Changed publication fully reopens and locality-checks the candidate, deletes
the existing zero-to-three root previews, and preserves `Index/ViewState.iwa`
plus all unrelated ZIP/IWA state. No-ops share exact source state, preserve
previews, and perform no changed-only lock/reassembly/reopen work. Exact patch
apply conflicts on the wrong artifact or selected source payload, charges
source-plus-target work before reopening a changed retained target, and inverse
restores the complete source.

The host cut removes exactly the two public Numbers editor methods, two whole
dedicated mutation tests, one duplicated `Count` test, and
`edit_numbers_table_headers.rs`. Ten mixed structural/sort tests and seven
creation/topology examples survive through private helpers or focused package
handoffs. The `table_headers` module/source, wire codec, attached helpers,
package bridge, structural/sort callers, and Pages/Keynote owners remain.

Within `litchi-numbers`, the private package owner is now split into `api`,
`dependencies`, `error`, `ownership`, `resolve`, and `rewrite` modules; every
file is under 600 lines and the public `table::headers::transaction` surface is
unchanged. Category-owner group declarations are indexed once and resolved
under linear aggregate work, preserving exact aggregate/path metadata rules
without repeated full metadata scans.

The sheet-scoped selector is a deliberate break from the old workbook-wide
catalog. Rooted canonical/legacy roles remain supported when unambiguous;
changed nested legacy physical packages refuse as `UnsupportedSource`. Locked
reads/no-ops remain valid, while changes refuse and delete root previews. No
edge or debt changes: debt 015 remains and inventory stays 64 packages, 235
internal dependency declarations, and 14 ordered debts.

The native refusal oracle changed source
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
in Numbers 14.4 to 2/2 header rows/columns without warning and preserved B2/B3.
The 136,213-byte save
`5c2323b509e5ea9a975b5f254bbd46cf42657aa1c3858d2c7e98f30f07e4b40c`
changed TableModel, HeaderNameMgr, a new manager tile object, and CalcEngine
formula/dependency state. It justifies typed dependency refusal and is not a
Rust writer/count-parity gate.

The compatible freeze oracle toggled Freeze Header Rows off from the same
source, retained 1/1 counts and B2/B3, and saved 136,199 bytes at
`015568e6b922e80fbfb760491dc49994ccc2218356ed197131beb46c1bd75850`.
Only TableModel 904538 field 12 changed from true-present to absent;
HeaderNameMgr stayed exact. The native off-to-on control hash was
`df44ed7d0b12c1d372dad7ad7361ed1140d41967921ee42b71a4072b78615721`.
Both saves regenerated equivalent ViewState with different IDs, so no native
raw-byte equality is claimed.

Verification passes 8/8 focused tests with defaults and 8/8 without default
features, 4/4 codec tests, 2/2 root-facade tests with `--features numbers`, and
114/114 boundary regressions. `cargo check -p litchi-numbers --all-targets`,
formatting, diff, warning-denied no-dependency rustdoc, and the doctest gate
(one compile-fail pass, one ignored example) are green. Strict Clippy has no new
header-file
finding, but full-crate Clippy remains blocked by unrelated baseline warnings.

The fuzz bin checks; its stable fixed-input smoke completed eight runs with
expected missing-sanitizer-symbol warnings, so no fuzzing/sanitizer
claim is made. Focused CLI source/inverse
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
produced changed artifact
`a8b88d21806b547a5265c60662610f68f524173cac1ca4252d368596c8ef8d2a`
with changed=true, one touched component, and three deleted previews. It was
not a native UI-open gate.

A distinct post-split freeze-row-only candidate, SHA-256
`c938d74bcf04be692097488af838f5105a8470e337eafa06fdc8b94b36231d6a`,
did pass a Numbers 14.4 Computer Use open: no repair/warning, Table 1 at 22 by
7, header columns/rows/footer rows 1/1/0, Freeze Header Rows unselected, and
B2/B3 preserved as the fixture text and 42. Its inverse was byte-exact.

Remaining debt is aggregate memory/work and fallible-allocation accounting,
process-local complete-artifact patches without stable semantic operations,
composition/merge/history, durable atomic library save, baseline Clippy
cleanup, and sanitizer-backed fuzzing.

## 2026-08-10 current-status amendment: Keynote placeholder visibility

Title/body placeholder visibility is owned behind
`slide::placeholder::{Kind, State, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and `Package::{slide_placeholder_visibility,
edit_slide_placeholder_visibility, apply_slide_placeholder_visibility}`. These
focused signatures expose no generated or raw-source types. Missing roles read
as `None` and cannot be synthesized by the consuming `set`/`show`/`hide` edit.
The public selector is canonically the shared `slide::placeholder::Kind` for
both slide-text and visibility operations. Replacing `SlideTextRole` is an
intentional source break; the common discriminator does not merge the
operations' distinct ownership and mutation contracts.

The format contract retains title/body stable references in SlideArchive fields
5/6 and represents visibility as exact membership in both owned-drawables field
7 and z-order field 42. The rooted Document `[2]` -> Show/SlideTree `[3,2]` ->
SlideNode field 2 chain, exact aggregate reference metadata, slide/placeholder
co-location, and the strict placeholder Buffa view jointly prove ownership.
Changed admission also refuses aliases, conflicting list metadata, merge state,
noncanonical framing, selected cache/layering state, layout overrides, and
builds targeting the selected placeholder.

Changed edits touch the slide component and, when separate, the SlideNode
component; they invalidate only that node's rendering cache and remove the
three root previews. `Index/ViewState.iwa`, other roles, date objects, content,
unknown fields, and unrelated components remain preserved. No-op and inverse
artifacts are exact, and changed patch application validates its retained
source before reopening the target. This does not move slide-number, layout,
placeholder creation, text-box, or style mutation.
Ownership uses linear payload occurrence/kind and metadata declaration indexes.
The bounded 4,096-to-8,192-object step remains within 2.3x recorded work. A
budget-aware single SlideNode pass conditionally invalidates and exact-verifies
the direction-aware delta. Verification uses only bounded, fallibly allocated
occurrence/declaration indexes, with no full node/payload clone or verification
rewrite; zero allowance fails atomically before publication. Structural work
includes every `MessageInfo`/`FieldInfo`, even empty records. A fixture with
4,096 empty `FieldInfo` records is atomically rejected by zero and payload-only
allowances, and the slide router precharges exact
`source + output + 2 * fields` work before allocation.
Full precharge includes selected/nonselected payload bytes, metadata vectors,
paths, features, and bases, every aggregate/`FieldInfo` reference in both
`Work` and `References`, and `header_length`. Low allowances atomically reject
the 256-KiB sibling plus 2,048 references/vectors.

Native Keynote 14.4 confirmed the list convention and UI behavior. The pristine
500,058-byte fixture is
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Title hidden
`d61a92b212d8a0f001bdfc24490d846e065b96885f0d0d0b86ef0be9f10e7580`
to reshown
`9d914ea25a42aaced4459a429e776b09b2024e2858133369f159dad7bce67325`
appended title after body; body hidden
`05ca9617ea5a23c57252c28c3029af96d4ec54345de331571d89b612566b8416`
to reshown
`8ee6ac8230273def64450b4cee86c9678849d77b5a7fbd11eb88e0c786279eee`
appended body. Checkboxes, canvas, date/other role, and reopen were confirmed.
Apple regenerated caches, so this is semantic rather than raw-cache evidence.

The Rust title-hidden candidate
`df119410433b97b9993d46619764a8ffb75f257b16c0680cd54faabd9a453cdd`
reported changed=true, two touched components, and three deleted previews; its
inverse exactly restored the pristine hash. Keynote 14.4 opened it warning-free
with Title off, Body on, and body/date retained. Save As, close, and reopen
preserved that state in the 475,102-byte native resave
`c5c996415191758b9fc638a8fdf024a912a6fe2ac4c3989970f0cb611e0670e3`.

Two-way Rust gates also pass exactly: Apple-hidden title
`d61a92b212d8a0f001bdfc24490d846e065b96885f0d0d0b86ef0be9f10e7580`
became shown
`3d36d31c6222b7622cab180f6dd9559ccf43f4b481e6b245c9d2c56fe8852b2c`,
and Apple-hidden body
`05ca9617ea5a23c57252c28c3029af96d4ec54345de331571d89b612566b8416`
became shown
`3e8855e954c16bd32350e057665b5ee4758a02e85ad23c3c6543f1caef177b13`;
each inverse restored its exact hidden source. Both shows reported
changed=true, two touched components, and three deleted previews.

The host cut is exact: the three
`KeynoteEditor::{set_slide_text_placeholder_visible, set_slide_title_visible,
set_slide_body_visible}` mutators, public `KeynoteSlideTextPlaceholder`, the
complete 150-line `keynote/editor/placeholder_visibility.rs` source/module, two
whole direct tests plus one exclusive constant, and the 30-line
`set_keynote_placeholder_visibility` example are gone. Five mixed layout
assertions use focused reads. Shared ownership and the layout and slide-number
paths remain.

Verification is 94/94 Keynote library, 18/18 slide-preview, 5/5 focused
visibility, 25/25 slide-text, 8/8 `--features keynote` facade, 7/7 doctest, and
129/129 boundary tests. Keynote all-target and host-lib checks plus strict
library Clippy/rustdoc, formatting, and diff checks pass. The expanded
`keynote_slide_text` fuzz target compiles and completes a bounded stable smoke;
missing sanitizer symbols make that control-flow evidence, not sanitizer-backed
fuzzing. Native and exact inverse gates pass. No edge or debt item closes.

## 2026-08-11 current-status amendment: per-slide Keynote slide-number visibility

Per-slide slide-number visibility has moved into the existing focused
`slide::placeholder` owner. This supersedes the preceding title/body section's
slide-number exclusion only. `Kind::SlideNumber` is the shared canonical
visibility discriminator in
`slide::placeholder::{Kind, State, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and the existing Package read/edit/apply methods; the slide-text
owner rejects it. Presentation-wide `KN.ShowArchive.slideNumbersVisible` field
6 remains independently owned by `show::Settings`, and layout, creation, text,
and style ownership do not move.

The focused projection proves Document field 2 -> Show/SlideTree `[3,2]` ->
SlideNode field 2 -> SlideArchive. SlideArchive field 20 names the selected
native-kind-1 placeholder. Visible requires canonical Node field 18 true and
one selected reference in each Slide field 7 and field 42; hidden requires
false/absent and no selected membership. Showing appends after the existing
field-7 and field-42 entries, and hiding removes only the selected entries.
Competing rooted slide ownership, role/closure aliases, contradictory
membership, noncanonical field 18, a missing selected placeholder, style
visibility overrides, or unsupported storage fail closed. Exact no-ops retain
absent versus explicit false; the process-local patch retains exact source
artifacts for inverse restoration.

The native storage-zero representation is accepted without inventing a
metadata reference. A nonzero storage is limited to the same-component strict
type-2001 storage/type-2043 slide-number-attachment closure: kind absent/3,
`in_document=true`, text one U+FFFC, one attachment at character zero, exact
metadata/dependency paths, empty textual super, and absent/zero attachment
kind. Other objects, styles, geometry, content, dependencies, and unknowns are
preserved rather than normalized.

The implementation is split between
`package/slide_placeholder_visibility/slide_number.rs` for rooted ownership and
storage closure and `package/slide_preview/slide_number.rs` for the strict
field-18 splice and exact delta. A new Buffa projection covers the node,
storage, borrowed attachment table, and attachment super; handwritten code
performs strict raw parsing first and forces/cross-checks the lazy views. Build
evidence is five generated files/112,101 bytes, zero repeated views, under
116 KiB, SHA-256
`eacce4103b5c9f9f32fd98639b81249ae1d15fcd63da6fe636569e0a2a324c30`.
Raw source artifacts, not generated output, remain the preservation authority.

Codec and transaction budgets cover bytes, fields, nesting, aggregate work,
rooted object/payload/metadata scans, references, selected/nonselected payload
bytes, output allocation, exact forward/inverse delta, and physical archive
reassembly. Bounded fallible indexes avoid a full node/payload clone and a
second verification rewrite. Failure is typed, redacted, and atomic.

Changed output touches the Node and Slide components (one if co-located, two if
split) and deletes the three existing root previews. It does not invalidate
the Node thumbnail/cache: only field 18, the selected field-7/field-42
membership, permitted metadata lengths/records, and preview deletions may
differ. ViewState, other slides and roles, storage/attachment closure, and
global Show field 6 remain exact. No-op skips reassembly/reopen; changed commit
reopens its candidate, and changed apply exact-checks source and target before
reopening the target. Output is through `write_to`; patch serialization and
durability remain debt.

The host cut removes `KeynoteEditor::set_slide_number_visible`, the complete
172-line `slide_number` source/module, one 23-line mutation example, and two
whole direct tests plus their four constants and fixture helper. The 53-line
creation example remains and hands the edit to the focused Package.
`KeynoteSlideInfo` read state, creation, shared placeholder ownership, layout,
title/body visibility, and global show settings remain. No edge or debt item
closes.

The 500,058-byte pristine native fixture is
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
Rust produced the 455,859-byte visible candidate
`a2dafcd4ffc57bafc3bbf7d7cd4ee8131bab2c06dd52adc292632d4208c126be`,
reported changed=true/touched=2/deleted=3, and exactly inverted to pristine.
Keynote 14.4 opened warning-free with Slide Number checked, attachment `1`
visible, and title/body/date exact. Save As/close/exact reopen preserved state
and content at 500,192 bytes,
`b1edd073d309157d27508baf4aedbe93d6dee0687f727dd71f1e8232f6171882`.
Keynote regenerated the root previews while Data9074 stayed byte-exact at
`575645e2455199d7cc0c65fab8002b9e025765ba19b8b03c6e51c000f4915e89`;
Apple-only controls independently confirmed the exact field-18 plus
field-7/field-42 membership delta and unchanged global Show field 6.

Frozen-tree verification passes 8/8 focused slide-number codec, 98/98 Keynote
library, 7/7 focused visibility, 22/22 slide-preview, 9/9 `--features keynote`
facade, and 7/7 doctests. Keynote all-target checking, strict Keynote library
Clippy/rustdoc, host library check/no-run and examples, formatting, and diff
checks are green. The fuzz target compiles and completes a bounded 16-run
stable control-flow smoke, but missing sanitizer symbols mean this is not
sanitizer-backed fuzz evidence. The boundary unit suite passes 138/138, the
live slide-number host, placeholder host, and focused audits are clean, and the
full checker retains only the unchanged 14 dependency-policy baselines. Native
compatibility and exact inverse are complete.

## 2026-08-11 current-status amendment: focused Keynote soundtrack settings

Keynote soundtrack playback settings now have a direct focused owner at
`soundtrack::{Mode, Settings, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and
`Package::{soundtrack_settings, edit_soundtrack_settings,
apply_soundtrack_settings}`. The rooted transaction selects Document field 2,
Show field 17, and one type-21 Soundtrack object with strict reference metadata,
nonexternal/nonaliasing identifiers, unique messages, non-merge state, and
bounded component framing. Absence reads `None` and cannot be changed into a
new soundtrack.

Only optional volume field 1 and mode field 2 are in the semantic delta.
Volume is finite and in `0.0..=1.0`; known modes have canonical enum variants,
while truly unknown discriminants round-trip. Field-3 movie-media references
are streamed by strict raw decoding and matched against message metadata,
PackageMetadata ownership/counts, safe data locators, and unique `Data/`
members. The item order, payloads, data-reference metadata, and unknown fields
remain exact.

A scalar-only Buffa projection is forced and cross-checked after strict raw
preflight; no repeated generated view or encoder owns field 3 or publication.
Build provenance is five generated files/27,753 bytes under a 32-KiB cap, zero
repeated views, and aggregate SHA-256
`458206e0b57d8ec5ae4c3fc706bf793ccd385ab867b7e92ac30d66ab1858b4d3`.
Codec reports and transaction work share bounded bytes/fields/nesting/work,
reference/media, component, compression/output, reassembly, reopen, and exact
comparison accounting. This policy does not close the remaining shared
allocation, peak-memory, work-bound, output, patch-serialization, or durable
save debts.

An exact no-op shares its package, touches nothing, and skips reopen. A change
rewrites one soundtrack component, reopens once, and permits only canonical
field-1/field-2 and selected length changes inside that archive, plus the
corresponding selected ZIP CRC/size/offset bookkeeping. Apply requires exact
source/target artifacts; inverse is byte-exact. Changed non-exact/legacy
provenance returns `UnsupportedSource`, while read/no-op compatibility remains.

The settings path is playback-only: previews, ViewState, slides and node
caches, field-3 items, media/data files and metadata, and unknowns remain exact.
The IWA soundtrack-item reader and add/insert/replace/move/remove mutations,
`KeynoteSoundtrackItemInfo`, creation, resource allocation/reclamation, and
their shared wire helper remain outside this owner.

The native source is 506,640-byte
`69795554212651b261f5ffd71dd5cf511544f285cab680d724a9de7d3f04b14d`.
Rust's same-size Loop/0.35 candidate is
`6367e38a2edeebe6e65b148d0fd2aae555ee219dc1a65c339954047eb533ce1a`;
only `Index/Document.iwa` changed and its inverse restored the source. Keynote
opened without warning, showed Loop/0.3499999940395355, retained `ringin`
00:00:01, and played it. Native Save As produced 506,651-byte
`e264f4e714b0c44fca420b2c7b43e18f2ed1be99a766d25fe901f68d5f8bc299`.
The media payload stayed exact at
`5a08f48c4f86074e14a763d4f19f49ca31196a7a5f52fb48960e76b6f3d3d96b`,
the slide and three previews were exact, and the normalized post-native restage
was an exact no-op.

The host cut removes the two direct settings methods, the entire 68-line
`soundtrack.rs` editor module/source, its settings-only wire patch helper and
dead decoded-native record field, two whole direct settings tests and their
exclusive support (157 test lines), and the 29-line mutation example. The
production delta is +2/-91 lines. The mixed inspector and README use the
focused Package. Item CRUD, shared soundtrack wire/media code, creation,
resource lifecycle, the item example, and item tests remain.

Topology remains 64 workspace packages/235 internal declarations/14
`litchi-iwa` dependency declarations/14 ordered debts. No edge closes and debt
014 (`litchi-iwa -> litchi-keynote`) remains.

Current verification is 5/5 codec, 1/1 focused scaling unit, 4/4 focused
settings, 99/99 Keynote library, 10/10 `--features keynote` facade, and 8/8
doctests. Keynote all-target, strict
Clippy/rustdoc, example, host, formatting, and diff gates are green. Performance
review is P0/P1-clean. The test-only `media.rs` gate exercises realistic
4,096/8,192 metadata/media states through the actual streaming path;
references double exactly and fields/work/references remain within 2.3x. This
is resource-accounting evidence without a wall-clock claim. Boundary
regressions pass 152/152; host and focused audits each report zero diagnostics,
and the full checker retains only the unchanged 14 baselines: six dev-only
annotation findings and eight edge classifications.

## 2026-08-11 current-status amendment: Numbers sheet-order owner

Numbers now owns one exact sheet move through
`sheet::order::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` and
`Package::{edit_sheet_order, apply_sheet_order}`; existing Document sheet
iteration remains the read path. A semantic selector moves once to a checked
final zero-based destination after removal. Positional no-op, missing/invalid
staging, unsupported source, resource/allocation, verification, and conflict
outcomes remain typed.

The native order is dual. Root type-1 Document field 1 orders sheet references
and field 5 selects a type-205 sidebar root; that root's field 2 orders one
child per sheet, and each child's field 3 associates it with the corresponding
sheet. The Document and sidebar order references must be unique ordered
subsequences in their aggregate metadata. The selected subsequences move in
lockstep; any selected order reference in `FieldInfo` is refused. Optional
sidebar/child declarations must use exact field-5/field-3/field-2 paths.
Root, sidebar, children, descendants, and ordinary type-2 sheets must be
nonexternal, disjoint, canonical, non-merge, and co-located in
`Index/Document.iwa`. FormBasedSheet and split-component mutation remain
native-unproven `UnsupportedSource` cases.

`TNNumbersSheetReferenceArchive.proto` is the sole scalar Reference projection.
Strict handwritten Document-field-1/5 and TreeNode-field-2/3 passes own all
repeated routing and force Buffa scalar parity without a generated repeated
view or encoder. The five-file closure is 32,579 bytes under 33 KiB, has zero
`RepeatedView`/`LazyRepeatedView`, and digest
`2a0850fd82cfbf337ed48e582d4a998bd27e5046eb63c61f6939fa5ff1a09854`.
Raw records remain authoritative.

Codec bytes/fields/depth/work/references and transaction lookup, metadata,
archive allocation/extent, compression/output, preview deletion, reopen, and
exact locality share finite budgets and fallible allocation. No-op shares
source and reports 0/0/0/false. Changed publication requires exactly one of
each of the three canonical source previews; missing or repeated members fail
closed. Commit rewrites one component, deletes all three, reopens once, and
verifies the dual move. Forward apply proves 3 -> 0 previews and inverse proves
0 -> 3; apply exact-authorizes and precharges source/retained-target work
before reopening. Conflicts and inverse are exact, while changed
legacy/non-exact sources fail closed.

Child IDs/nodes/associations/descendants, CalcEngine, ViewState, ordinary
sheet/table/drawable graphs, global table order, sidecars, and unknowns are
exact. The host retains sheet add/duplicate/remove, FormBasedSheet/general
Document-reference substrate, table/drawable CRUD, and allocation/reclamation.
Only previews are deliberately deleted beyond the two order sequences and
necessary owner/message/ZIP bookkeeping.

Independent performance review is P0/P1-clean with no release blocker or
O(S²). Strict codec, raw reorder, and core aggregate-header work at 4,096 and
8,192 references stays within 2.3x plus a fixed 32-unit production allowance;
the codec-only bound is strict 2.3x. There is no wall-clock claim. P2 remains:
about four snapshots per sheet (cap 4,096) trade bounded memory for no source
reselection/O(1) inverse; Vec-to-Arc publication may transiently duplicate the
target; and separately allocated byte-equal patch sources may incur one bounded
O(package-bytes) authorization comparison before charging (identity is O(1)).

Matched Apple control/reorder artifacts are 133,594-byte
`f9c5cbec4f422484c63d1d39bd8d09da122d011596561a5feb2ad1e812574990`
and 153,498-byte
`7b3bcbc853346a433e84ee815d28671d01fc3da857e43b8b7d29b310f94e7e1a`.
They establish simultaneous Document/sidebar reversal with child associations
exact and 93/103 decompressed members, including table sidecars, unchanged.
Apple cache/subgraph/tree/ViewState/ID/metadata/property/timestamp churn is
native normalization, not the minimal focused delta.

Rust candidate
`97c76894503a2628c1828babd93d9a9a891794d86c86177cab60f09333997a68`
opened warning-free in Numbers 14.4 with `FirstCreated`, `SecondCreated` and
`A-new`/`A-old`/`B-only` associations correct and CalcEngine benign. Save As,
close, and exact reopen produced the semantically identical 103-member
`4aa257e4db61a3c03950360b29267c9495985d460ae22b6f679bee31f2693217`.
Its three regenerated previews exactly matched the Apple reorder. A focused
same-position restage and inverse remained exact at that hash with diagnostics
0/0/0/false.

The implementation inventory is five sources: `sheet/order.rs`,
`package/sheet_order.rs`, and frozen private
`package/sheet_order/{error,resolve,rewrite}.rs`. The host cut removes the move
method and exclusive `sheet_index` (-58 production lines), changes tests
+2/-43, deletes the 23-line move example, and migrates the retained remove
example +2/-6 to a semantic selector. Sheet add/duplicate/remove and shared
substrate remain.

Codec/protobuf gates pass 7/7 and 132/132; Numbers passes 109/109 library, 4/4
private sheet-order, and 1/1 public integration tests. Boundary regressions pass
165/165; Python compilation/diff are green; host/focused audits are empty. The
full checker retains only 14 unchanged baselines: six missing dev-only
`soapberry-zip` annotations plus eight unclassified edges (those six and the
`litchi-odf-common`/`litchi-opc` edges to `xml-minifier`). Topology remains 64
packages/235 internal declarations/14 `litchi-iwa` declarations/14 ordered
debts, including debt 014 (`litchi-iwa -> litchi-keynote`).

## 2026-08-11 current-status amendment: Numbers table-title owner

Numbers now exposes the focused
`table::title::{Settings, Edit, Patch, Commit, Diagnostics, Error, LimitKind,
Path}` family through
`Package::{table_title_settings, edit_table_title, apply_table_title}`. The
API is selector-first and archive-free; the focused signatures expose no raw
source, native identity, component, or generated value, and publication uses
`write_to`. `Settings` independently preserves presence for TableModel
field-22 visibility and field-37 outline; the consuming
infallible `Edit::set` stages the complete value.

The changed owner follows the rooted Document/Sheet-or-FormBasedSheet/
TableInfo/TableModel chain, requires exact local reference metadata and one
canonical selected message, and rejects a locked table. Effective visibility
also requires valid field-33 height and distinct exact field-30/field-36
references to canonical type-2022 paragraph and type-2025 shape styles. Missing,
external, aliased, or malformed prerequisites fail closed. Changed admission
scans `Index/ViewState.iwa` and returns `UnsupportedSource` for any native
type-6284 table-name-selection message because that transient selection state
is an unsupported dependency. Reads and exact no-ops remain broad. Accepted
changed sources preserve every other ViewState byte exactly.

The scalar-only private projection covers fields 22/33/37 and reuses the
existing scalar Reference view for fields 30/36. Strict raw validation precedes
forced Buffa parity; raw records retain preservation authority. Generated
evidence is five files/32,332 bytes under 33 KiB, digest
`56cfd70666ffa6079175bdab0a63a4ddd055099edf3c771ed3ad8b3051596ee1`,
with 9/9 focused codec and 141/141 full protobuf tests.

Exact no-op shares source and skips reopen. Changed publication rewrites one
selected `Index/CalculationEngine.iwa` component, deletes each existing
canonical preview (zero to three), reopens, and verifies semantic state and
exact locality while preserving accepted ViewState and all other components.
Exact source/target apply, conflict, and inverse semantics remain process-local
and byte-exact; legacy changed sources fail closed.

The native control resave is 136,204 bytes/SHA-256
`25c9fc858ca4fb4f1fedeafb944e96afb81af03a082a41be297ecf6f2542dbdb`;
the native title-hidden artifact is 136,273 bytes/
`ac8a7117ad6256b0da2e6d191b9e64f721b689d71696a89ac0f78bc6aa513a28`.
Numbers removes raw field 22 for the hidden form instead of encoding false;
field 37 retains its independent presence contract. The native comparison is
not evidence for mutating type-6284 ViewState; changed admission rejects it.

The exact Rust source is the 136,357-byte
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`;
the 136,351-byte hidden candidate is
`4c7f6340b6f2675240577c5b59d5c154de24c8a7e763a31257c56a9899a8e40c`,
and its inverse restores the source exactly. Numbers 14.4 opened it warning-free
with Table Title off and retained the 22-by-7 table, B2 fixture marker, and B3
value 42. Warning-free Save As, close, and exact-URL reopen preserved that UI
state in the 136,353-byte resave
`5b162f8431f45333f0ae9a8654dfa724794f2ec2b391ea11f6a5eee7822cbb10`.

Performance review is final with no P0/P1 finding. At 4,096 -> 8,192 objects,
the real rooted Package path records fields 53,307 -> 108,363 (2.0326x),
`WireWork` 315,936 -> 636,752 (2.0155x), references 16,386 -> 32,770 (exact
`2 + 4N`, 1.9999x), and `TransactionWork` 9,084,384 -> 18,298,157 (2.0142x).
All are at most 2.3x; maximum-minus-one work refuses before output. P2 consists
only of linear selector temporary vectors and redundant changed decode passes.

The host retirement removes two public NumbersEditor methods, 32 production
lines, 245 direct-test lines, and the legacy title example. Private cross-format
package helpers and wire code stay for Pages and Keynote. Boundary regressions
pass 173/173. Final gates are 111/111 Numbers library, 2/2 private title, 5/5
public title, 9/9 codec, and 141/141 full protobuf tests. The full checker has
only 14 unchanged dependency-policy baselines. Inventory is 64 packages/237
internal declarations/14 `litchi-iwa` dependency declarations/14 ordered
debts; debt 015 remains and no edge closes.

## 2026-08-11 current-status amendment: aggregate Pages section settings

The prior Pages section-name and pagination records remain historical evidence,
but their independent-writer descriptions and statements that the
legacy-normalizing settings/name writer remains are superseded. The concrete
current owner is `litchi-pages::section::settings`, operating on the
archive-free, presence-preserving `section::Settings` through
`Package::{section_settings, edit_section_settings,
apply_section_settings}`. Exact names and checked positions are the only public
selection identities. Generated values, raw IDs, physical member names, wire
records, and exact artifacts remain private.

One strict raw/Buffa projection covers optional fields 17--22, 26, and 28. Raw
records preserve and splice those settings; Buffa is a bounded borrowed
semantic cross-check and has no production encode path. The legacy
pagination-only projection remains a scoped facade reader, while both the
section-name and pagination transaction facades delegate their physical work
to the aggregate core. The aggregate generated-output ceiling and digest are
80,202 bytes under 80 KiB and
`2202f4b1d394346450cb9f88a41c2784ab476cff23b181fffbab6f37b4a42b62`;
its five generated files contain no repeated lazy view, and the focused
protobuf suite passes 149/149.

Target-sensitive dependency validation covers the previous section's template
closure and fields 23--25 without rewriting them. A changed edit also proves
and exactly preserves rooted layout/cache state, its metadata, and every root
preview. Only the selected section component is rewritten; template payloads,
field-30 background, fields 29/31, sibling sections, and all unrelated physical
state remain exact. An exact no-op precedes dependency planning and remains
byte-identical with zero touched components, preview deletions, or reopen.

The host no longer owns `section_settings`, `set_section_settings`, or
`set_section_name`, and the raw-ID settings example is retired. The separate
background writer remains under a private background seam. Changed legacy
nested sources are refused instead of normalized; reads and exact no-ops remain
supported. Matched Pages 14.4 pairs prove independent field-17, field-19, and
field-28 false-to-true changes as one exact scalar delta on section
1732889/type 10011, with header/references, field 18, templates, storages, names,
caches, and previews exact. Warning-free close/reopen showed the expected
inherited, alternating, and hidden-first-page header/footer behavior. The seed
hash is `19b8a24c7bc0d57d87614a0f08215072c9c61519b15629827f5a448b29218422`;
full pair hashes are recorded in ADR 0008. Production scaling at 4,096 and
8,192 rooted real-package objects keeps selected fields/wire work/references at
77/564/4 and scales `TransactionWork` from 292,154 to 587,222 (2.0100x), with
one output allocation and reopen at each size. A maximum-minus-one work budget
fails before output with both counters zero. Focused integration is 7/7, two
private production tests cover observation and scaling, and two private
security tests cover alias metadata plus repeated-reference scaling/refusal;
the projection suite is 149/149, and locality review is clean. The complete Pages
library/integration gate is 118/118; boundary regressions are 181/181; focused
facade and host audits are empty; and the live checker retains only the 14
unchanged baselines. The matched native pairs are the UI oracle; no distinct
Rust-authored UI artifact is claimed.

No manifest edge or ordered debt changes. The inventory remains 64 workspace
packages, 237 internal declarations, 14 `litchi-iwa` dependency declarations,
and 14 ordered debts, including debt 017 (`litchi-iwa -> litchi-pages`).

## 2026-08-12 current-status amendment: Numbers table-cell mutation

The concrete Numbers package now contains the focused cell-batch owner behind
`table::cells::{Input, Change, Edit, Patch, Commit, Diagnostics, Error,
LimitKind, Path, DependencyKind}` and
`Package::{edit_table_cells, apply_table_cells}`. The previously documented
eager semantic read remains unchanged. The mutation owner adds strict raw and
Buffa validation, physical scalar/tile and string-list planning, bounded
non-text scalar sparse growth, in-place authored-text replacement in uniquely
owned rich backing,
exact style-reference release, final-overlay formula-cache refresh,
grouped publication, exact locality, and process-local reversible patches.

This is not the whole table-cell aggregate. The supported set is finite scalar
set/clear, direct/unsegmented string-list assignment/release with exact
refcounts, synthetic 513-row finite non-text scalar sparse growth, in-place
authored-text replacement that retains unique rich key/storage identity, and
strict supported cache propagation.
Canonical payload field-1-to-storage and storage field-2-to-style FieldInfo
metadata may be present and remains exact; a rewrite requiring any FieldInfo
reference transition, or noncanonical/ambiguous FieldInfo rich ownership,
refuses as `RichText`.
Formula compilation and formula-cell construction, arbitrary rich text,
formatting, controls, merge/pivot/category/spill/hidden/conditional state,
comments, and Pages/Keynote attached-table mutation remain outside this owner.
CalculationEngine field 14 is projected and its rooted HeaderNameMgr reference
validated; only the referenced manager payload/update semantics are not, so a
manager-backed header change refuses as `HeaderNameIndex`. Sparse text to a
missing tile refuses as `SharedString`. Segmented string lists, existing
formula/error cells, and modeled unsupported ownership/formula closure fail as
`UnsupportedDependency`; impacted active merge, pivot, category, spill,
hidden, or conditional-style state refuses by its matching kind while
unrelated/inert state remains exact. Malformed routes fail as `InvalidSource`,
a modeled missing storage prerequisite fails as
`UnsupportedDependency { CellStorage }`, an unmodeled stored BNC value/source
kind fails as `UnsupportedSource`, and locked ownership fails atomically. Reads
and exact no-ops stay broad, while changed packages without
an exact physical `SourceCatalog`, including nested legacy sources, fail as
`UnsupportedSource`.

Storage and dependency projections currently measure five files/465,932 bytes
and five files/544,538 bytes, with SHA-256
`1a894fd5d22b004db664bc7c348d9591a4608ab9263a8122c726c8a1ecb0c3b3`
and `2fba7c22aef58ed3cfe6eba1f77e5eaf79d2597dd79966e05d20e50c0e2b33b3`;
both generate zero repeated views. The full protobuf inventory is currently
178/178. The strict formula projection remains five files/201,539 bytes with
SHA-256
`ccd972b3dcd76b6142342d36435f2f76a305c029265853ced04d64c1e2bf1752`,
and its focused codec gate passes 7/7. Exact patches privately retain both
verified package snapshots so apply can borrow the patch and run directional
locality without reopening; that memory and the
lack of durable serialization/composition/merge/history remain debt.
The PackageMetadata projection is five files/145,681 bytes, has no repeated
generated view, and has SHA-256
`ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31`;
its focused gate passes 7/7.

Final rooted transaction-work ratios are 1.1899x numeric, 1.2245x unique text,
1.1396x same-tile, and 1.8021x formula when fixtures double 4,096-to-8,192;
governed subterms are at most 2.0x. Required-minus-one formula/sparse cases
reject before component, reassembly, output, reopen, or locality work. The
numeric B3=43 scalar and unique-rich no-impact candidates pass Numbers 14.4
open/Save As/reopen and
exact inverse gates; the latter preserves its independent formula/cache and is
not impacted-formula native proof.

The host cut retires three direct NumbersEditor cell mutators, two Numbers-only
raw-ID model writers, Numbers-only batch apply, 15 obsolete direct tests, and
the legacy example. Shared attached-table APIs, lower physical machinery,
Pages/Keynote owners, builders, and fixture-only adapters remain. Numbers
passes 237 library tests with 4 ignored and 15/15 public cell tests; boundary
regressions pass 196/196. The neutral private rich-text wire edge
`litchi-numbers -> litchi-iwa-text-wire` leaves the current inventory at 64
workspace packages, 238 internal declarations, 14 `litchi-iwa` declarations,
and 14 ordered debts.

## 2026-08-12 current-status amendment: Keynote existing-slide deletion

`litchi-keynote` now contains the focused existing-slide deletion owner at
`slide::delete`, exposed only through the canonical nested
`Edit`/`Patch`/`Commit`/`Diagnostics`/`Error`/`LimitKind`/`Path` vocabulary and
`Package::{edit_slide_deletion, apply_slide_deletion}`. Exact navigator names
and checked semantic positions are the public identities. Native object IDs,
component locators, PackageMetadata identifiers, and wire values remain
private.

The changed path supports one exact flat Document -> Show/SlideTree ->
SlideNode -> Slide ownership chain. It proves aggregate and field-specific
reference agreement when optional field attribution is present, unique
package-wide inbound ownership, single selected messages, no merge/base/diff
state, unique current PackageMetadata components, exact Node/Slide UUID
bindings, exactly one supported object-specific or component-level
Node-to-Slide external edge, and exact selected data-reference owner/count
records. Unsupported
hierarchy or legacy alternate slide roots, ambiguous ownership, versioned or
contradictory registry state, and malformed metadata refuse before output.

Publication removes one Show slide-reference record, the Node and Slide
objects, two UUID bindings, any exact object-specific external-reference
record, and the selected data-owner records. A component-level edge remains.
It does not remove an IWA component. Co-located objects, component
registrations, the PackageMetadata last-object identifier, global data-catalog
records, and all data payloads remain. A component
data-reference record remains with surviving owners or is removed when none
survive. Exact root previews are invalidated; near-name previews and unrelated
ZIP/IWA state remain exact. The candidate is reassembled and reopened once,
and the exact patch inverse restores the accepted source.

This owner performs no media garbage collection. `Data/` members and shared,
uncertain, or newly unreachable media are preserved; reclamation remains a
separate future reachability transaction. Slide creation, duplication,
layouts, drawable graphs, and media or soundtrack-item CRUD also remain
outside this focused owner.

The host no longer contains `KeynoteEditor::remove_slide`, its
`slide_delete` module/source, its direct example, or its direct deletion tests.
A retained generated-presentation regression is creation-only and does not
claim its backlink topology is deletable; focused deletion refuses the
surviving child-to-parent-slide reference as `AmbiguousOwnership`. No public
bridge alias replaces the retired host method. Debt 014
(`litchi-iwa -> litchi-keynote`) nevertheless remains, and no manifest edge
is removed. The boundary suite passes 204/204; focused and retired-surface
audits each report zero findings, and the full checker reports only the 14
established unrelated findings. PackageMetadata
generated evidence is five files / 145,681 bytes / zero repeated views / SHA-256
`ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31`.
Native Save As/reopen evidence is frozen in ADR 0008. The final topology is 64
workspace packages, 238 internal dependency declarations, 14 `litchi-iwa`
dependency declarations, and 14 explicit ordered debts. Keynote passes 235/235
all-features tests and 9/9 doctests; the retained host library passes
1,418/1,418. The permanent generated-child-backlink regression proves typed
`AmbiguousOwnership` refusal and byte-exact source preservation. The focused
existing-slide deletion cut is green; broader host ownership and debt 014
remain.

## 2026-08-12 current-status amendment: private Numbers formula-cache foundation

The current tree's bounded private Numbers cell-cache planner preserves an
unrelated cycle marker byte-for-byte, refuses when an impacted marked formula
survives the final same-batch overlay, and succeeds when that overlay removes
it. Graph work has exact max-minus-one refusal coverage; scratch and allocation
remain bounded by the planner limits.

There is still no focused public formula-authoring surface. Production host
formula setters and raw formula vocabulary remain, so the crate graph,
manifest edges, and ordered debts are unchanged. No formula-native or
formula-authoring performance gate is claimed.

## 2026-08-13 current-status amendment: Pages section-background bounds

The focused Pages field-30 transaction reuses the bounded section transaction
profile for source discovery, strict wire work, ownership validation, rewrite,
reassembly, and candidate reopen. Its dedicated 4,096-to-8,192 object gate
keeps each observed size-sensitive counter at or below 2.20x, while each
successful changed operation performs one output allocation and one candidate
reopen. An instrumented required-minus-one `TransactionWork` ceiling refuses
before publication.

These are deterministic bounded-work gates, not claims about latency,
throughput, RSS, allocator events, peak scratch, or complete end-to-end
locality accounting. Exact package/member locality and inverse behavior are
separately checked by the focused transaction tests. Apple Pages 14.4.1
accepted, saved, closed, and exact-path reopened both candidates without repair
or conversion, retaining dark-red `Color Fill` and `No Fill` respectively.
The Pages-resaved ZIPs pass integrity, but their independent rewriting is not
used as a locality, allocation, scratch, or byte-preservation measurement.

## 2026-08-13 current-status amendment: Keynote reader cutover

The previous status entries that left `KeynoteDocument::show` and its eager
Prost graph as open debt are superseded. `litchi_keynote::Document::{open,
open_with_options}` is now the canonical semantic Keynote reader for complete
ZIPs and frozen app-authored package directories. It captures `PreparedSource`,
eagerly completes bounded decoding, and returns an archive-free full `Show`,
rooted text, source-derived metadata, and source statistics. Source-backed
metadata combines semantic Show values with narrowly decoded canonical
properties scalars when that diagnostic is present; `Some` does not prove
sidecar presence. `litchi_keynote::Package` remains the exact complete
regular-file artifact owner; the cross-format coordinator can delegate semantic
reads to the focused owner. The host
`keynote/document.rs` file, module, reader type, stats type, and re-export are
gone, removing 933 lines and the duplicate
`Bundle`/`ObjectIndex`/semantic-cache pipeline.

The focused reader is bounded, not Prost-free. Six generated-message decodes
remain behind strict wire preflight during semantic traversal; they are not a
second public reader.

The focused surfaces retain the complete supported read capability set.
`Document` owns semantic path ingress, cheap snapshots, rooted text, slides,
metadata, show, validation, and source statistics; `Package` owns exact ZIP
path/byte ingress, semantic projection, cheap shared `semantic_snapshot`,
writing, and editing. That package-derived semantic snapshot is intentionally
diagnostic-free: `metadata()` and `stats()` are `None`.
The old archive-bytes constructor was only an alias for byte ingress, and its
stats application field was a constant rather than semantic state. Direct
`Package::open` refuses directories so an `Index.zip` fragment cannot
masquerade as a complete artifact with write/edit provenance. Archive-free
semantic reads do not promise preservation of other sidecars, `Data/`,
previews, or exact package bytes.

Focused semantics intentionally differ where the duplicate reader was too
broad or lossy: unreachable theme/template storage is excluded from text,
rich storage fragments are retained, and metadata/validation are stricter.
Metadata lookup accepts only canonical logical `Metadata/Properties.plist` and
ignores unrelated basename matches. Its centralized 64 KiB hard admission
ceiling is independent of broader entry limits, and decoding is restricted to
the scalar fields projected into public metadata.
The generated roundtrip, host compile/lint/doctest, focused path, native
fixture, and boundary gates cover the retired surface and replacement paths.
The unchanged
500,058-byte native read fixture has SHA-256
`3a3d07476b45b6e543bcfba75fe38a245434176dcb3565e34570b817708b9f42`.
The live retired-reader audit is clean, and the full checker continues to
distinguish its dependency-policy baseline findings. Permanent path regressions
prove packaged/directory Keynote semantic parity through both focused
`Document` and the coordinator, plus directory/focused-ZIP snapshot parity.
Frozen ingress and semantic gates pass archive-directory 16/16, detection
18/18, focused Keynote native 7/7, coordinator `iwork_path` 7/7, and metadata
scalar/64 KiB-cap unit coverage 1/1.
Keynote 14.4 opened an isolated fixture copy without repair/recovery/conversion
and showed the one expected slide and its three visible text sentinels. The
separate non-UI focused fixture gate reports one slide/959 objects. Native
autosave normalization changed only the disposable copy; the
checked-in source remained exact.

The host still owns `KeynoteEditor` and `KeynoteDocumentBuilder`; debt 014 and
the manifest edge remain. This status closes only the duplicate reader cell.

## 2026-08-13 current-status amendment: Pages reader cutover

`litchi_pages::Document` is now the canonical archive-free Pages reader.
It captures complete ZIP files and checked app-authored package directories on
supported path-ingress platforms, or borrowed ZIP bytes and shared ZIP bytes
on every supported platform, through one prepared source. It eagerly validates
the semantic projection and retains cheap shared sections plus
optional source metadata and statistics. Semantic
constructors and a package's borrowed `semantic_document()` retain no source
diagnostics. `litchi_pages::Package` remains the exact regular-file/byte owner
for source bytes, package diagnostics, physical validation, and
editing; direct package open continues to reject directories.

The focused structural projection supersedes the retired reader's narrower
model: empty root means zero sections; rooted section names and UTF-16
boundaries are authoritative; and duplicate exact names remain a typed
selector ambiguity. The source-backed metadata
handoff freezes only the exact Properties, BuildVersionHistory, and
DocumentIdentifier authorities, retaining at most 64 KiB from each. It does not
retain arbitrary directory metadata, `Data/`, previews, media, or unknown
sidecars.

The legacy Pages reader source/module/export and its three reader/state/stats
types are gone. Its public eager-Prost path is not replaced by a second focused
reader. The focused path remains eager and not Prost-free: strict raw and
private Buffa projections qualify root and section-boundary reads. Fallback
candidates pass full known-field raw storage validation before their Buffa text
projection, while one bounded rooted StorageArchive decode remains
Prost-backed.
Rootless fallback now reproduces the retired object-level trigger and
aggregation contract under strict raw/Buffa validation and the semantic text
budget.

Directory capture enforces the 64 KiB ceiling before allocating each selected
sidecar. Packaged ZIP path, borrowed-byte, and shared-byte capture preflights
the three exact raw logical authorities' declared sizes and compression methods
before any package entry payload is materialized. Selection strips only the
chosen legacy outer-package prefix before raw-byte comparison, excludes raw
near-names, and requires local/central ZIP names and methods to agree. Public
semantic open maps lower layers into content-free `ReadError` categories and
numeric bounds.
Semantic ZIP ingress can still expand unrelated supported entries under the
generic source limits before discarding them; selected-sidecar preflight does
not make it a filtered catalog. Windows Pages file and directory path ingress
fails closed until stable, reparse-safe identity can be pinned, while borrowed
and shared byte ingress remains available there. Those are retained scope and
platform qualifications, not capabilities supplied by deleting the host
facade.

Final verification passes 153/153 focused Pages tests, 93/93 archive tests,
32/32 detector tests, and the 1/1 host generated-roundtrip gate. The boundary
suite passes 227/227; both the live retirement audit and the focused public-API
audit report zero findings.

The host still owns `PagesEditor`, `PagesDocumentBuilder`, creation, and broad
chart, table, media, and formatting workflows. Its direct focused dependency
is still used by production code, so ordered debt 017 and the manifest edge do
not close. The native read oracle and its silent-normalization caveat are
frozen in ADR 0008.

## 2026-08-13 current-status amendment: Numbers reader cutover

`litchi_numbers::Document` is now the canonical archive-free rooted Numbers
reader. It accepts checked packaged-file and app-authored-directory paths,
borrowed bytes, and caller-owned shared bytes, eagerly completes bounded
semantic projection, and retains one cheaply shared workbook state. Its public
surface is semantic and selector-first; native bundles, object indexes,
component/member names, protobuf/Buffa values, raw identifiers, and exact
package bytes do not cross the boundary.

Source-backed documents may retain content-free statistics and narrow metadata
from exactly `Metadata/Properties.plist`,
`Metadata/BuildVersionHistory.plist`, and
`Metadata/DocumentIdentifier`. Metadata capture is conditional on the frozen
source being classified as Numbers. Semantic constructors and package-derived
documents have no source metadata or statistics. The focused `DocumentStats`
fields are `source_record_count`, `sheet_count`, and `table_count`; the former raw
`total_objects` vocabulary and constant application tag are absent.
Unix file and directory path ingress uses pinned, no-follow capture. Other
non-Windows targets use version-checked path capture. Windows path ingress
fails closed until pinned, reparse-safe handle traversal exists;
borrowed/shared byte ingress remains portable.
Each selected metadata authority is capped at 64 KiB before packaged payload
materialization or directory-sidecar allocation. A narrow `plist::stream`
event projection then enforces fixed event, nesting, history-entry, scalar, and
retained-property budgets instead of constructing a general scalar DTO or
plist value tree.

Rooted `Document::plain_text` is available independently of source
diagnostics. It emits deterministic rooted workbook semantics and fixes the
legacy reader's inability to construct the checked-in native workbook before
its public `text()` could run. Empty rendered Text/Formula values are excluded.
Recovered private storage output matches `Package::text` on two frozen
fixtures only; `Package::text` remains a separate physical/storage diagnostic
with no unqualified legacy-parity claim.

The 460-line host `numbers/document.rs`, its module/export and
reader/state/statistics types, and the 142-line reader-only `NumbersSheet`
adapter are gone: 602 host reader lines in total. This removes the duplicate
bundle, object index, root/sheet/table traversal, semantic cache, and public
raw getters. `NumbersTable`, `TableDataExtractor`, `NumbersEditor`, and
`NumbersDocumentBuilder` remain host owners. Their continued production use
keeps debt 015 and the `litchi-iwa -> litchi-numbers` manifest edge open.

Focused Numbers still eagerly decodes substantial generated Prost graph state.
Private strict raw/Buffa projections selectively qualify root sheet order,
sheet references/names, TableInfo, table model/list/segment/tile/rich-text, and
formula-category boundaries; focused `Document` does not retain comments. The
remaining table/formula graph is not fully Buffa-lazy. This cut deletes the
duplicate public reader; it does not claim every `Package` path is hardened, a
codec completion, latency/RSS result, or monolith deletion.

Frozen gates pass 16/16 focused reader cases, with a seventeenth
Windows-configured case; 240 Numbers library cases pass and four are ignored;
compatibility and name gates pass 5/5 and 10/10. Archive coverage passes 127
cases (125 unit plus two integration), detector coverage passes 40/40, and the
host library passes 1,397/1,397, while generated-roundtrip and doctest gates
pass 1/1 and nine passed with three ignored. Host all-target check and no-run,
strict scoped host Clippy, focused
all-target Clippy, strict focused rustdoc, formatting, and diff checks pass. The
boundary units pass 237/237 and both live retirement/API audits report zero
findings. The 15-file host-scoped cut contains 329 insertions and 888 deletions,
net -559. Broad host all-target Clippy remains blocked by unrelated existing
lints; the global boundary policy still reports 14 unrelated
`soapberry-zip`/`xml-minifier` debt findings.

The native evidence and its silent-normalization caveat are frozen in ADR
0008. The tracked 136,357-byte source remains exact at SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`;
only an isolated copy was opened in Numbers 14.4 build 7043.0.93.

## 2026-08-13 current-status amendment: Numbers dimension sizes

`litchi_numbers::Package` is the Numbers owner for one
selector-first row-height or column-width override. Its public seam is
`table_dimension_size`, `edit_table_dimension_size`, and
`apply_table_dimension_size`, with transaction types under
`table::dimension::transaction`. The semantic leaf keeps `Dimension`, checked
positive finite `Points`, and the lossless distinction between `Size::Default`
and `Size::Points`.

The focused implementation reuses the existing strict/raw plus private Buffa
header-storage codec. One changed request owns one selected header-bucket
rewrite, deletes three canonical previews, performs one full reopen, and
preserves every other header facet, component, ZIP record, and unknown field. It does
not change table row/column counts and is not a table resize, axis lifecycle,
automatic-fit, or bulk-sizing API.

The public Numbers host exit changes five paths with four formatting
insertions and 312 deletions, net -308. It removes six methods (65 file lines,
63 method bodies), 200 test-section lines, the 41-line example, and the public
legacy re-export names. Private
header-bucket helpers remain in `litchi-iwa` because Pages and Keynote attached
tables call them. That cross-format helper, `NumbersEditor`, broader table
editing, debt 015, and the `litchi-iwa -> litchi-numbers` manifest edge remain.

Focused dimension tests pass 13/13, codec tests pass 11/11, protos passes
194/194 plus doctests, and Numbers passes 241 library tests with four ignored,
91 integration tests, and five doctests with one ignored. Archive passes
130/130 plus doctests. Focused strict checks/Clippy pass; boundary units pass
243/243 with zero live audit findings. Host retained-axis, Pages-layout,
generated-roundtrip, scoped boundary, check/no-run, strict library Clippy,
formatting, and diff gates pass. Broad host all-target Clippy retains nine
unrelated existing lints. ADR 0008 records the accepted native oracle.

## 2026-08-14 Current Numbers formula cut

Selector-first semantic formula-cell authoring is production-owned by
`litchi-numbers`. The focused API exposes bounded semantic expressions, opaque
same-source table handles, typed caches, and cell-edit constructors; it exposes
no raw formula key, table identity, UUID, protobuf/Buffa value, IWA object, or
wire payload. Local and distinct-owner cell/range/whole-axis authoring,
supported scalar/function evaluation, complete survivor overlays, replacement,
clear, downstream cache refresh, reversible patches, and conflict detection are
covered by the accepted gates in ADR 0008.

Physical publication remains private and raw-authoritative. It coordinates the
formula list, BNC tiles, inline and type-4009 tiled dependency mirrors,
CalculationEngine tracker, PackageMetadata, ArchiveInfo references, string and
rich storage, caches, headers, and one package rewrite. Output-free logical
plans feed a single aggregate execution barrier; strict generated views are
agreement checks only. Native Numbers recalculation and save/reopen evidence is
recorded in ADR 0008.

The legacy host still contains production readers, setters, examples, pivot
compatibility vocabulary, tests, and a manifest dependency. This cut removes
the focused facade's dependency on that public vocabulary; it does not remove
the host, its edge, or its ordered debt.

## 2026-08-23 current-topology amendment

The authoritative current inventory is 64 workspace packages and 239 internal
dependency declarations. The ordered migration-debt ledger has 13 entries:
001, 002, 004, 005, 008, 009, 010, 012, 013, 014, 015, 016, and 017. Earlier
inventory and debt entries remain historical records and are not rewritten by
this current-state amendment.

## 2026-08-23 superseding amendment: Keynote soundtrack-reference order ownership

This current-state amendment supersedes the earlier soundtrack-order wording;
those dated passages remain historical records. The Keynote package now owns
the bounded `soundtrack::order::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` transaction through
`Package::{edit_soundtrack_order, apply_soundtrack_order}`. It selects the
existing rooted Document -> Show -> Soundtrack chain, validates the field-3
reference closure and ownership, and changes only the order of existing media
references.

Media payloads/assets, data-reference metadata, unknown fields, and unrelated
components remain preserved. The order transaction does not create, delete,
replace, allocate, reclaim, or otherwise CRUD soundtrack media, and it does not
change playback settings. The focused `soundtrack::{Mode, Settings}` playback
transaction remains the package owner for mode and volume.

`litchi-iwa` remains the owner for the remaining soundtrack media/settings CRUD
and host compatibility, including soundtrack item/asset creation,
add/insert/replace/remove/lifecycle operations, and retained host readers and
editors outside the focused package seams. The current inventory remains 64
workspace packages, 239 internal declarations, and 13 ordered debts; debt 014
and the `litchi-iwa -> litchi-keynote` edge remain open.

This is bounded package-level ownership evidence only. No native Keynote
open/save gate is admitted, no host-edge retirement is claimed, and no
monolith-exit or host-deletion claim follows.

## 2026-08-25 current-topology amendment: Wave71 Pages artifact and footnote cut

Commit `580a5343a2c75a1c1b185a5cc8aff4a87e2a5c11` updates the focused Pages
surface without changing the workspace topology. `Package::source_bytes` is
now crate-private and `Package::from_archive_bytes` is removed; exact retained
ZIP output is public only through `Package::write_to` and root `WriteError`.
The writer streams exact bytes with partial-write/`Interrupted` handling and
typed redacted offset errors, but it does not flush, sync, rename, or provide
atomic or durable filesystem publication.

`PagesEditor::set_body_footnote_text` is removed. The existing-root
selector-first `litchi_pages::Package::edit_body_footnote_text` transaction
owns replacement, while the host retains footnote reads and insert/remove
graph lifecycle. This is one focused API/operation migration step; it does not
move the remaining Pages graph or native publication responsibilities.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. The
`litchi-iwa -> litchi-pages` edge remains, as do generated schemas, normal
Prost/Buffa owners, the migration host, and the monolith. Wave71 retires no
crate, edge, debt, or host-exit gate.

## 2026-08-25 amendment: Wave77 current-topology Pages body-footnote lifecycle cut

Implementation commit `5dc2ab5337cb61b72f83e369d818829c355d7141` makes
`litchi-pages` the current owner of bounded body-footnote insertion, selected
edit/removal, exact patch application, and ordered semantic reads. The focused
package coordinates the body anchor, table/reference/marker/storage graph,
strict raw-preserving codec, collision-free identifiers, Metadata UUID and
external ownership, root watermark, selected save tokens, previews, ZIP
locality, candidate reopen, and exact inverse artifacts.

The compatibility examples and host route now delegate through the focused
selector-first transaction rather than retaining a second raw graph writer.
Wave71 existing-root footnote-text replacement remains available. Ordinary
body-text cleanup that cannot prove the complete footnote graph and metadata
ownership remains fail closed, as do malformed, shared, aliased, ambiguous,
resource-owned, and otherwise unproven lifecycle graphs.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 017 and the
`litchi-iwa -> litchi-pages` edge remain open because other Pages package
editing, compatibility examples/tests, and host responsibilities remain. No
crate, dependency edge, debt item, generated-schema owner, normal Prost/Buffa
owner, migration host, or monolith is removed by this cut.

## 2026-08-25 amendment: Wave78 current-topology Pages header/footer text cut

Implementation commit `1596d5106ee42fe9238e8d63cc55d39c494d59c3` makes
`litchi-pages` the current owner of existing-root header/footer text reads,
`Some -> Some` replacement, and clear-to-empty edits through the
selector-first `HeaderFooterSelector` surface. The package resolves the
rooted body/section/template/storage chain, proves exact aliases and
aggregate/`FieldInfo` ownership, applies the selected metadata token
transition, preserves root field 1 and raw unselected/versioned data, and
publishes only after exact patch, preview, candidate-reopen, and locality
verification.

The raw header/footer APIs and public raw `TextStorageId`/
`PagesHeaderFooterInfo` route are retired. Header/footer number attachments
now route through `Package`; template/slot creation and removal, inheritance
and first/even/odd settings, section lifecycle, media, annotations, table
and remaining text graphs remain at their recorded compatibility owners.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 017 and the
`litchi-iwa -> litchi-pages` edge remain open. No crate, dependency edge,
debt item, migration host, generated-schema owner, normal Prost/Buffa owner,
or monolith is removed by this cut.

## 2026-08-25 amendment: Wave80 current-topology Numbers table-appearance cut

Implementation commit `bf01576c090cb508ac0596a5faac01951ef502d0`
makes `litchi-numbers` the current owner of selector-first, existing-root
table-appearance reads, copy-on-write replacement, exact patch application,
and inverse artifacts. The package owns strict model/style/stylesheet and
metadata validation, bounded style inheritance, new variation and registry
publication, UUID/watermark/save-token updates, aggregate resource accounting,
candidate reopen, preview invalidation, and locality verification.

The mutating Numbers editor wrapper is retired. A read-only compatibility
fallback remains in `NumbersEditor::tables()` for older/source-built graphs,
and the shared root appearance implementation remains for Pages and Keynote.
Preset/network lifecycle, cross-component stylesheet graphs, table topology
and content, appearance reset/cull, and remaining Numbers graph operations
stay at their recorded owners.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain open, as do debt 017, the Pages
edge, the migration hosts, generated-schema and normal Prost/Buffa owners,
and the IWA monolith. No crate, dependency edge, debt item, production
manifest dependency, or monolith is removed by this cut.

## 2026-08-25 amendment: Wave79 current-topology Pages section-text host-retirement cut

Implementation commit `507193d3c2ea7c6f6939f47189be5a7b425661c0`
makes `litchi-pages` the current owner of existing rooted Pages section-text
reads, checked UTF-16 set/clear/span edits, exact patch application, and
inverse artifacts. The package owns rooted body/section selection, strict
raw-preserving text-wire preparation/execution, section-boundary shifting,
archive and ZIP publication, candidate reopen, topology and locality checks,
and exact preservation of Metadata, previews, unrelated members, and
unselected section text.

The former raw-ID `PagesEditor` section-text methods and fallback writer are
retired. The exact aggregate-only drawables-z-order reference emitted by Pages
is retained as a non-owning ordering edge; other shared, aliased, data,
`FieldInfo`, marker, rootless, nested, or otherwise unproven routes fail
closed. Section lifecycle, whole-body edits across structural boundaries,
header/footer lifecycle, annotations, tables, media, and other Pages graph
responsibilities remain at their recorded owners.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 017 and the
`litchi-iwa -> litchi-pages` edge remain open. No crate, dependency edge,
debt item, migration host, generated-schema owner, normal Prost/Buffa owner,
or monolith is removed by this cut.

## 2026-08-25 amendment: Wave81 current-topology Numbers table-cell Pop-Up Menu cut

Implementation commit `4ea040b4f9cf97f61ae9a690c6ae592b1b7ac567`
makes `litchi-numbers` the current owner of selector-first reads, set/clear/
reset lifecycle, copy-on-write or rooted reuse, final unreferenced-model cull,
exact patch application, and inverse artifacts for admitted same-component
Numbers Pop-Up Menu graphs. The package owns rooted cell/table-list/model
selection, BNC refcount census, strict codec/storage validation, metadata
UUID/watermark/save-token transitions, aggregate resource accounting,
candidate reopen, and object/member locality.

The dedicated Numbers editor Pop-Up Menu methods are retired. The generic
Numbers data-format compatibility surface delegates its Pop-Up Menu branch to
the focused package; shared private adapters and Pages/Keynote table-control
hosts remain. Cross-component format/control graphs, other control formats,
general data-format and table topology, and unsupported alias/resource graphs
remain at their recorded owners and fail closed at this boundary.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain open, as do debt 017, the Pages
edge, migration hosts, generated-schema and normal Prost/Buffa owners, and the
IWA monolith. No crate, dependency edge, debt item, production manifest
dependency, or monolith is removed by this cut.

## 2026-08-25 amendment: Wave82 current-topology Keynote movie-playback cut

Implementation commit `666ee3ec3be5d7574ebb9324154b550fde83a5f1`
makes `litchi-keynote` the current owner of selector-first playback reads,
replacement, exact patch application, and inverse artifacts for admitted
existing rooted file-backed slide movies. The package owns strict slide/movie
selection, same-component identity and owner proof, playback-codec execution,
aggregate resource accounting, prepared reassembly, candidate validation and
semantic readback, and exact object/member locality. Metadata and previews are
preserved rather than rewritten.

The former raw-ID `KeynoteEditor::{slide_movie_playback_settings,
set_slide_movie_playback_settings}` methods are retired. The tracked Keynote
movie example and compatibility-host regression use the focused package. The
shared private media-playback implementation, movie graph/media creation and
removal, geometry, title/caption, builds, and Pages/Numbers media hosts retain
their recorded owners.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 014 and the
`litchi-iwa -> litchi-keynote` edge remain open, as do debts 015 and 017,
migration hosts, generated-schema and normal Prost/Buffa owners, and the IWA
monolith. No crate, dependency edge, debt item, production manifest
dependency, or monolith is removed by this cut.

## 2026-08-25 amendment: Wave83 current-topology Numbers table-sort cut

Implementation commit `20aca0ede7817e7fbb630338dbb4bcc852d7d500`
makes `litchi-numbers` the current owner of selector-first persisted sort-order
reads, set/clear transactions, exact patch application, and inverse artifacts
for admitted rooted Numbers tables. The package owns exact table/model
selection, strict field-44 codec execution, operation-local resource
accounting, candidate reopen, semantic readback, and exact model/member/
preview locality. Field 45 and Metadata remain opaque and exact.

The persisted `NumbersEditor` read/set/clear methods are retired. The physical
row-sort executors remain in `litchi-iwa`, consume the focused configuration,
and retain table-storage, UID, formula, comment, border, and other physical
movement responsibilities. Pages and Keynote adapters also remain at their
recorded hosts.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain open, as do debts 014 and 017,
migration hosts, generated-schema and normal Prost/Buffa owners, and the IWA
monolith. No crate, dependency edge, debt item, production manifest
dependency, or monolith is removed by this cut.

## 2026-08-26 amendment: Wave84 current-topology Pages body-table sort cut

Implementation commit `31d5081ca6cd56256e463bee0019e4c7241d6df1` makes
`litchi-pages` the current owner of selector-first persisted body-table
sort-order reads, exact configuration edits, patch application, and inverse
artifacts for admitted rooted Pages tables. The package owns
`BodyTableSelector` resolution, strict field-44 codec execution,
operation-local resource accounting, candidate reopen, semantic readback, and
exact model/member locality. Field 45 and Metadata remain opaque and exact.

The physical PagesEditor sort/apply/reorder executor remains in
`litchi-iwa`, along with row movement, table storage, UID, formula, comment,
cell, and other physical graph responsibilities. Pages and Keynote sort
adapters remain at their recorded hosts. Locked, aliased, malformed, or
otherwise unsupported table graphs fail closed.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 017 and the
`litchi-iwa -> litchi-pages` edge remain open, as do migration hosts,
generated-schema and normal Prost/Buffa owners, and the IWA monolith. No
crate, dependency edge, debt item, production manifest dependency, or
monolith is removed by this cut.

## 2026-08-26 amendment: Wave85 current-topology Numbers unified cell-control cut

Implementation commit `8f804fdc65f5d99a61d8b73503353901edf87488` makes
`litchi-numbers` the current selector-first owner of Checkbox, Star Rating,
Slider, Stepper, and compatibility Pop-Up Menu reads, set/clear transactions,
exact patch application, and inverse artifacts for admitted same-component
Numbers table graphs. The package owns rooted cell/model/tile/list selection,
strict mixed CellSpec and display-format codec execution, BNC/list refcount
census, control copy-on-write/cull, metadata UUID/save-token transitions,
operation-local budgeting, prepared publication, semantic readback, and exact
object/member locality.

The dedicated Numbers editor control method families are retired. The generic
data-format bridge consumes the focused owner for file-backed Numbers
controls. Private source-built compatibility readers/resets, Pages and Keynote
control adapters, cross-component format/control graphs, general scalar data
formats, and table topology remain at their recorded owners. Unsupported
graphs fail closed.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain open, as do debts 014 and 017,
migration hosts, generated-schema and normal Prost/Buffa owners, and the IWA
monolith. No crate, dependency edge, debt item, production manifest
dependency, or monolith is removed by this cut.

## 2026-08-26 amendment: Wave86 current-topology Numbers unified cell-control split-read cut

Implementation commit `3dfe506f4febe7f389db4f60b2988430bbd6038e` makes the
current topology claim for unified Numbers controls deliberately bounded:
strict reads and byte-exact no-ops may traverse selected split-component
format/control graphs, while all changed split-component routes fail closed
before native/ZIP candidate publication. The package proves current/effective
locator and external-edge ownership for the selected edge, rejects versioned,
conflicting, effective-locator-disagreeing, and physical-alias targets, and
uses one cached `RegistryFacts`/physical census per logical read.

Selected `TableModel` sidecar refs require one aggregate occurrence; explicit
`FieldInfo` is accepted only when unique and `ObjectReference` typed, with
producer-omitted `FieldInfo` accepted and no exact path claim. The root
Document/TableInfo-to-CalculationEngine/TableModel metadata edge is not owned
or proven. Same-component graphs do not require external-edge inspection;
opaque inbound refs are accepted for read/no-op but reject changed routes.
No cross-component COW, UUID/save-token, locality, candidate reopen, inverse,
or successful split write is part of the topology cut.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain open, as do debt 017 and the
`litchi-iwa -> litchi-pages` edge, migration hosts, generated-schema and
normal Prost/Buffa owners, and the IWA monolith. No crate, dependency edge,
debt item, production manifest dependency, or monolith is removed by Wave86.

## 2026-08-26 amendment: Wave87 current-topology Keynote movie-geometry cut

Implementation commit `e11a4cc993cf29e5524745f2fa51dd3dd7d20b3e` makes
`litchi-keynote` the current selector-first owner of admitted existing
file-backed slide-movie position and displayed-size reads, exact edits,
patch application, and inverse artifacts. The package owns strict rooted
slide/movie selection, same-component identity and parent proof, geometry
codec execution, preview invalidation, operation-local resource accounting,
prepared reassembly, candidate validation/readback, and exact object/member
locality. The public value remains archive-free `MovieGeometry`.

The physical `litchi-iwa` media graph, creation/removal, media/poster
replacement, captions/titles, playback, builds, metadata, and broader
drawable responsibilities remain at their recorded owners. Legacy angle and
flags compatibility plus flip/original-size-restore compatibility remain
retained at the host/adapter boundary; this topology cut does not claim those
operations or fields as a separate native graph owner. Cross-component,
non-file, aliased, malformed, or otherwise unproven movie graphs fail closed.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 014 and the
`litchi-iwa -> litchi-keynote` edge remain open, as do debts 015 and 017,
the Pages edge, migration hosts, generated-schema and normal Prost/Buffa
owners, and the IWA monolith. No crate, dependency edge, debt item,
production manifest dependency, or monolith is removed by Wave87.

## 2026-08-26 amendment: Wave88 current-topology Numbers split cell-control write cut

Implementation commit `d1e11c7f218583becc85be7fa426650aa0113de4`
makes `litchi-numbers` the current owner of admitted split-component Checkbox,
Star Rating, Slider, and Stepper reads, exact edits, patch application, reset/
clear, final cull, and inverse artifacts through the existing unified facade.
The package now owns strict current/effective component and metadata authority,
split model/tile/list transitions, BNC/list refcount census, operation-local
resource accounting, prepared publication, candidate validation/readback,
preview invalidation, and exact object/member locality.

Cross-component Pop-Up Menu writes remain unsupported; popup reads and the
previously admitted same-component popup lifecycle retain their existing
owners. General data formats, table topology, Pages and Keynote control graphs,
and arbitrary graph repair remain at their recorded boundaries. No additional
host method or adapter is retired by Wave88.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain open, as do debts 014 and 017,
the Pages edge, migration hosts, generated-schema and normal Prost/Buffa
owners, and the IWA monolith. No crate, dependency edge, debt item, production
manifest dependency, or monolith is removed by Wave88.

## 2026-08-26 amendment: Wave90 current-topology Numbers split Pop-Up Menu lifecycle cut

Implementation base `e867003d80fb9cc9373f2f659e2b51c3264de06b` makes
`litchi-numbers` the current selector-first owner of the admitted
split-component Pop-Up Menu lifecycle: read, exact no-op, replacement,
create/reuse, clear/reset, final cull, patch, inverse, and candidate semantic
verification. Private native and metadata siblings own member enumeration,
aggregate merge/split projection, exact-byte changed-member filtering,
current/effective locator and external-edge proof, UUID/save-token batching,
and BNC/list refcount and copy-on-write transitions.

The split topology accepts the narrow aggregate-only metadata producer shape
under those ownership invariants without claiming an exact `FieldInfo` path.
Prepared reassembly, candidate reopen, semantic readback, and object/member
locality remain package-owned. Malformed, segmented, versioned, aliased,
duplicate, conflicting, opaque-inbound, unsupported, or ambiguous graphs
remain fail-closed. The native acceptance record confirms semantic
persistence for one admitted fixture; it does not establish native/Rust byte
parity or general native graph compatibility.

The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain open, as do debts 014 and 017,
the Pages edge, migration hosts, generated-schema and normal Prost/Buffa
owners, and the IWA monolith. No host method, crate, dependency edge, debt
item, production manifest dependency, or monolith is removed by Wave90.

## 2026-08-26 amendment: Wave91 current-topology Numbers comment-reply read cut

Implementation base `e8749ca9e5320c5bf99c17967bf850f7c6e14d6c` makes
`litchi-numbers` the current selector-first owner of ID-free direct
table-cell reply reads. The private path owns strict comment-storage decoding,
recognized comment-graph census, direct source-order projection, alias/cycle/
external/nested refusal, and authored-text debug redaction.

`litchi-iwa` retains the deprecated identity-bearing `NumbersEditor` reply
read plus reply creation, replacement, removal, graph cleanup, authors, and
native object identities. This compatibility surface is not a duplicate
semantic owner: it exposes information and mutations intentionally absent
from the focused read facade.

The inventory remains 64 workspace packages, 239 internal dependency
declarations, and 13 ordered migration debts. Debt 015 and
`litchi-iwa -> litchi-numbers` remain, together with other migration hosts,
generated-schema/Prost/Buffa owners, and the IWA monolith. No host, edge, debt,
manifest dependency, crate, or monolith owner is retired.

## 2026-08-26 amendment: Wave92 current-topology comment-reply rewrite primitive cut

Implementation commit `65bdac3ece028a997c73dc82c4d49ef7839ff002`
makes the hidden comment-storage codec the owner of one prepared,
source-preserving reply-reference mutation primitive. The codec owns strict
known-field/reference validation, admitted unknown raw framing, checked reply
ordinal and identity transitions, candidate wire verification, and
codec-local execution requirements. Its fuzz target and owner-independent
boundary ratchet are now part of the current source topology.

`litchi-numbers` remains the read-only semantic reply owner from Wave91.
Reply graph mutation, comment-list and BNC refcounts, author and metadata
lifecycles, package archives, patches/inverses, candidate publication, and the
deprecated native-ID compatibility API remain outside this cut. Native direct
reply acceptance is withheld.

The inventory remains 64 workspace packages, 239 internal dependency
declarations, and 13 ordered migration debts. Debt 015 and
`litchi-iwa -> litchi-numbers` remain, along with other hosts,
generated-schema/normal Prost/Buffa owners, and the IWA monolith. No host,
edge, debt, manifest dependency, crate, or monolith owner is retired.

## 2026-08-26 amendment: Wave93 current-topology Numbers comment-reply lifecycle cut

Implementation commit `21004a78ec4c6c7d8436424de270ad6ed6051eb0`
makes `litchi-numbers` the current selector-first owner of the admitted
existing-root, same-member, direct-leaf comment-reply lifecycle. The package
owns ordinal reads, append/set/remove, shared-root COW, exact list/BNC
refcounts, strict author/ArchiveInfo/Metadata authority, private identity and
save-token transitions, staged publication, candidate reopen/readback,
locality, patches, and inverses. Its codec and package fuzz/boundary contracts
are part of the current source topology.

`litchi-iwa` Numbers mutation entry points delegate admitted sources to that
owner while retaining deprecated compatibility behavior for unsupported
native-ID graphs. Pages/Keynote comments, root-comment creation, shared reply
leaves, segmented or cross-member reply graphs, author creation, and broader
comment graph repair remain at their recorded owners or fail closed. Native
direct-reply acceptance remains withheld.

The inventory remains 64 workspace packages, 239 internal dependency
declarations, and 13 ordered migration debts. Debt 015 and the
`litchi-iwa -> litchi-numbers` edge remain, together with other migration
hosts, generated-schema/normal Prost/Buffa owners, and the IWA monolith. No
crate, edge, debt, manifest dependency, compatibility host, or monolith owner
is retired by Wave93.

## 2026-08-26 amendment: Numbers cell-comment clear host cut

Commit `4f7f2c16c` retires the public raw-ID
`NumbersEditor::clear_cell_comment` route and its migration-host fallback.
The supported operation now appears only on the selector-first
`litchi_numbers::Package` facade and in the migrated package-based example;
the old native helper is test-only for preservation of unsupported legacy
fixture behavior. A boundary ratchet enforces that topology.

This does not retire all Numbers comment compatibility: raw-ID comment read
and root replacement, reply-native identity fallback, root creation, and
unsupported graph cleanup remain migration-host responsibilities. The
inventory remains 64 packages, 239 internal dependency declarations, and 13
ordered migration debts. Debt 015, the `litchi-iwa -> litchi-numbers` edge,
other hosts, generated-schema/normal Prost/Buffa owners, and the IWA monolith
remain.

## 2026-08-27 amendment: Wave98 current-topology Keynote slide-table title cut

`litchi-keynote` is now the selector-first owner of lossless title visibility
and outline settings for an existing canonical type-6001 slide table. The
public boundary is `SlideSelector` plus an archive-free `TableSelector` and
the shared semantic table-title `Settings`; exact edit, patch, inverse, and
apply types remain Keynote-prefixed. The private resolver may follow the
slide, TableInfo, model, and title-style graph across physical members, while
a changed transaction rewrites only the selected model member.

Strict generated-free title and TableInfo projections, canonical known-field
validation, raw-preserving scalar replacement, one conservative operation
budget, prepared ZIP publication, preview invalidation when previews exist,
candidate reopen, and object/member locality are part of the owner boundary.
Locked tables, legacy type-6000 models, ambiguous ownership, invalid style
prerequisites, and unsupported compatibility graphs remain fail-closed or in
the legacy host. Metadata and save tokens remain byte-identical because this
in-place scalar slice allocates no native objects.

The inventory remains 64 workspace packages, 239 internal dependency
declarations, and 13 ordered migration debts. The raw Keynote compatibility
host, debt 014, debt 016, the `litchi-iwa -> litchi-keynote` edge, other
migration hosts, generated-schema/normal Prost/Buffa owners, and the IWA
monolith remain. No crate, edge, debt, manifest dependency, compatibility
adapter, or monolith owner is retired by Wave98.

## 2026-08-27 amendment: Wave99 current-topology Keynote slide-table persisted sort configuration (not a monolith-exit gate)

Wave99 defines a narrow `litchi-keynote` package boundary for persisted sort
configuration on an existing canonical `TST.TableModelArchive` (field 44).
The public operation is selector-first: `SlideSelector` identifies the slide
and the checked position-only `TableSelector` counts only table drawables in
that slide's z-order. The public semantic value is the archive-free common
`table::sort` model (`ColumnIndex`, `Direction`, `Order`, `Rule`, `Scope`, and
`RowRange`); `Scope::SelectedRows` is persisted configuration, while
`RowRange` is not accepted by the package configuration edit. Transaction
types remain Keynote-prefixed, and patch application is distinct from the
legacy physical `Sort Now` executor.

The focused owner is limited to a uniquely rooted canonical type-6001 table
and a source-preserving field-44 rewrite. Field 45 and every other model
field, object, member, metadata record, and preview remain byte-authoritative
unless the narrow transaction explicitly changes field 44. Legacy type-6000,
malformed, duplicate, aliased, locked, unsupported multi-owner, and ambiguous
graph routes fail closed. No native identifier, archive, ZIP, wire, generated,
Prost, or Buffa type is part of the public facade.

Admission requires the strict neutral `numbers_table_sort_order_codec` route,
its prepared report/requirements/execute contract, bounded operation-local
resource preflight, prepared reassembly, candidate reopen and semantic
verification, exact patch/inverse behavior, and object/member locality. The
sort slice preserves previews exactly (`deleted_previews == 0`); it does not
claim row movement, selected-row execution, table storage or cell mutation.
Those physical responsibilities remain in the legacy Keynote host.

The Wave99 boundary gate passes all 573 Python policy tests; both checker
modules compile with `py_compile`, and the scoped diff-check is clean. These
are source-topology and documentation gates only and do not constitute a
Cargo, native-application, or native byte-parity result.

A bounded Keynote 14.4 evidence record used a fresh table-only source at
`/private/tmp/wave99-native-table-source.key` (499,854 bytes, SHA-256
`e5ddda5583d4312f67501312f2681859e0969bc8c49cdcd8160f9b93fe4c21ef`).
The package owner published one entire-table ascending rule on column zero,
reopened the 499,861-byte candidate with exact semantic readback, and produced
a byte-identical inverse. The 86-member ZIP set was preserved and only
`Index/CalculationEngine.iwa` changed (3,183 to 3,190 member bytes). Keynote
reopened the exact candidate without repair and rendered the five-by-four
table. Its formatter exposed no persisted-sort control, so this proves native
no-repair readability, not native semantic UI acceptance, normalization,
physical row sorting, or native post-save byte parity.

The migration host, generated-schema owners, `litchi-iwa -> litchi-keynote`
edge, debt 014, debt 016, every other current migration debt, and the IWA
monolith remain unchanged. No crate, dependency edge, host adapter, debt item,
or monolith is retired by Wave99.

## 2026-08-27 amendment: Wave100 current-topology Keynote slide-table headers (conservative scalar owner)

Wave100 adds a narrow selector-first `litchi-keynote` owner for the persisted
header/footer/freeze/repeat settings of an existing, uniquely rooted canonical
`TST.TableModelArchive` type-6001 slide table. `SlideSelector` selects the
slide and the checked position-only `TableSelector` selects its table in
z-order. The archive-free semantic surface is
`slide::table::headers::{Count, Error, Settings}`; transaction types remain
Keynote-prefixed and are re-exported through the nested transaction namespace.
Native object identifiers, archives, ZIP entries, generated models, wire
views, Prost, and Buffa values do not cross the public facade.

The strict neutral `numbers_table_header_settings_codec` prepared codec
admits exactly the following seven persisted header/footer/freeze/repeat
settings: `header_rows`, `header_columns`,
`footer_rows`, `header_rows_frozen`, `header_columns_frozen`,
`repeating_header_rows_enabled`, and `repeating_header_columns_enabled`.
Known-field duplicates, wrong wire types, noncanonical encodings, malformed
framing, ambiguous ownership, locked tables, unsupported dependencies, and
legacy type-6000 models fail closed. Unknown fields and their raw framing are
preserved. Admission charges bounded input/output, field, work, nesting,
reference, allocation, retained, scratch, and transaction resources before a
candidate allocation; prepared requirements are checked again at exact
execution. Publication is one canonical model rewrite followed by candidate
reopen, semantic readback, exact inverse, and object/member-locality checks.

This is a conservative persisted-scalar boundary only. It performs no row
movement, cell or tile mutation, formula/storage rewrite, metadata/UUID/save
token publication, preview deletion, dimension/topology/appearance change, or
sort execution. The host retirement ratchet is correspondingly limited to
production `litchi-iwa` Keynote branches and Keynote branches of shared
examples; compatibility tests and unrelated Numbers/Pages branches remain
outside this retirement.

The native Keynote 14.4 acceptance record is deliberately narrower than the
seven-field synthetic matrix. The real table carries `HauntedOwner` and a
rooted `HeaderNameMgr`, so the owner correctly refuses count changes and the
accepted edit preserves the 2/1/1 header-row/header-column/footer-row counts
while toggling only the four freeze/repeat flags. The 500,128-byte source
(`47cf0d…563b`) produced a 500,134-byte candidate (`687e6e…157`) and an inverse
that is byte-identical to the source. Exactly one payload member changed,
`Index/CalculationEngine.iwa` (3,480 to 3,486 bytes); no member was added or
removed and every unrelated payload remained exact. Keynote opened the
candidate without repair, displayed the 5-by-4 table with the preserved
2/1/1 counts, and reopened a separately saved normalized copy without repair.
The normalized copy is native evidence only, not a byte-parity oracle.

Debt 014, debt 016, the `litchi-iwa -> litchi-keynote` edge, all other
migration debts, the migration host, generated-schema and normal
Prost/Buffa owners, and the IWA monolith deletion gate remain unchanged. No
crate, dependency edge, manifest dependency, compatibility host, debt item,
generated-schema owner, or monolith owner is retired by Wave100.

## 2026-08-27 amendment: Wave101 current-topology Keynote slide-table bounded borrowed discovery (not a monolith-exit gate)

The legacy `KeynoteEditor::slide_tables` listing path and the
`add_slide_table` template lookup now share one private, bounded
`KeynoteObjectCatalog` per operation. It retains compact object slots and
message descriptors while borrowing archive payloads only inside bounded
callbacks; listing reuses one decoded slide context rather than rebuilding a
cloned package graph for each table.

Strict TableInfo and borrowed table-model discovery projections admit table
identity, name, and dimensions only after canonical/wire/duplicate/role checks.
Canonical model type 6001 is authoritative; strict type-6000 legacy is
considered only when no type-6001 model exists. Simultaneous 6000/6001
candidates reject, malformed canonical input does not fall back, and
TableInfo-shaped aliases fail closed. Catalog axes cover archive
reads, objects, messages, payload/reference totals, retained descriptors, and
semantic decodes. The separate projection has finite input/field/work/text/
nesting limits. The catalog has no retained parsed payloads and is not a
package-wide budget.

This is a discovery/listing cut only. Full generated TableInfo/model values,
geometry, appearance, storage, cells, formulas, tiles, comments, metadata,
native mutation, and all existing `ObjectGraph` compatibility callers remain
at their current owners. No public native-ID surface is removed or expanded,
and no zero-copy, full generated-free Keynote, performance/RSS, or monolith
exit claim is made. The focused evidence is codec 5/5, catalog 9/9,
slide-table tests 30/30, and boundary tests 591/591; the focused Wave101 live
audit returned no findings.
An external public `KeynoteEditor` driver read the Wave99 source
`/private/tmp/wave99-native-table-source.key` (499,854 bytes, SHA-256
`e5ddda5583d4312f67501312f2681859e0969bc8c49cdcd8160f9b93fe4c21ef`) and the
Wave100 source (500,128 bytes, SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`). Each
read one slide and one 5-by-4 table named `Table 1`; fresh reopen parity was
true and post-read bytes and hashes were unchanged. Computer Use also opened
the Wave100 source in Keynote 14.4 without repair, recovery, or conversion UI
and observed `Table 1` with 5 rows and 4 columns; the source bytes and hash
remained exact afterward. No Keynote save, normalization, or mutation run is
claimed, nor any performance/RSS result or public catalog statistics.

The migration host, `litchi-iwa -> litchi-keynote`, all 13 ordered migration
debts (including 014, 016, and 017), generated-schema/Prost/Buffa owners, and
the IWA monolith remain unchanged. Wave101 retires no crate, edge, debt,
manifest dependency, host adapter, or generated owner.

## 2026-08-27 amendment: Wave102 current-topology Keynote slide-table appearance listing (not a monolith-exit gate)

The Keynote slide-table listing now adds a bounded, borrowed appearance
projection on top of the Wave101 `KeynoteObjectCatalog`. One
operation-scoped catalog and decoded slide context are reused. Strict
`table_appearance_codec` views resolve the model
style or preset, preset-to-network, network-to-style, and bounded full parent
inheritance; a direct nonzero style keeps legacy precedence. Appearance
payloads are borrowed for the projection and are not retained in the catalog.

Missing, malformed, duplicate, role-alias, cyclic, and over-depth facts on
the projected model/style/preset/network and traversed parent routes fail
closed. Stylesheet-registry ownership and unprojected style-property fields
remain compatibility-owned. This remains listing-only: existing public APIs
and writers, legacy mutation, and the selected generated TableInfo geometry
path are unchanged. It adds no metadata, UUID, save-token, global-ownership,
or package-wide transaction behavior and exposes no native identifiers,
archives, ZIP entries, wire values, generated models, Prost, or Buffa values.
The bounded limits apply to this discovery operation only; no package-wide
budget, zero-copy, allocation-free, RSS, or wall-clock performance claim is
made.

Wave102 evidence is 14/14 for the appearance codec, 50/50 for the focused
Keynote slide-table tests, 598/598 for the boundary suite, and an empty live
audit. Protos and IWA library checks, strict scoped Clippy (with existing
Pages warnings allowed for IWA), the isolated appearance fuzz check, and
format/diff checks passed.

The external read-only driver replayed the Wave99 and Wave100 sources
(`e5ddda5583d4312f67501312f2681859e0969bc8c49cdcd8160f9b93fe4c21ef` and
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`). Each
contained one 5-by-4 `Table 1` and reported disabled row banding, fixed row
sizing, and all five gridline classes visible. Fresh replay was exact across
all 86/86 ZIP members: no member changed, was added, or was removed.
Computer Use opened an immutable disposable Wave100 copy in Keynote 14.4
without repair, recovery, or conversion UI and rendered the selected 5-by-4
table with alternating row color and resize-to-fit off, all five gridline
toggles on, and 1/2/1 header-column/header-row/footer counts. Its original
500,128-byte hash remained exact after close. A separate writable disposable
copy auto-persisted view state on close; no native appearance mutation,
save, normalization, or post-save acceptance is claimed.

The migration host, `litchi-iwa -> litchi-keynote` edge, all 13 ordered
migration debts, generated-schema/normal Prost/Buffa owners, and the IWA
monolith deletion gate remain unchanged. Wave102 retires no crate, edge,
debt, manifest dependency, host adapter, generated owner, or monolith owner.

## 2026-08-27 amendment: Wave103 current-topology Keynote selector-first slide-table appearance owner (not a monolith-exit gate)

`litchi-keynote` now owns selector-first slide-table appearance reads and
edits through `Package::slide_table_appearance`,
`edit_slide_table_appearance`, and `apply_slide_table_appearance`. The
changed direct nonzero style path uses same-component copy-on-write and a
prepared table-appearance codec plus strict metadata, UUID, external-edge,
ArchiveInfo, candidate-reopen, inverse, and locality validation. Preset and
network routes are readable but not mutable; preset-only writes fail closed.

Canonical rooted graph admission rejects missing, malformed, duplicate,
role-aliased, wrong-wire, cyclic, over-depth, ambiguous, and otherwise
unproven style routes. Existing generated TableInfo geometry, public
compatibility surfaces, and legacy/native mutation remain at their owners.
No row/cell/formula/storage/tile or unrelated table-graph rewrite is part of
this boundary; changed commits delete stale previews instead of rewriting
them. Its resource ledger is conservative operation-local
logical accounting, not allocator, cache, RSS, or package-wide budget
telemetry; no full generated-free, zero-copy, allocation-free, wall-clock, or
broader Keynote graph claim is made.

ArchiveInfo admission resolves every distinct referenced object and current
data identifier while preserving native duplicate occurrences in unrelated
producer FieldInfo lists; selected slide, model, and style routes remain
exact-one checks.

The focused appearance integration passed 19/19; the Keynote library check
and strict library/test Clippy passed. Scoped host bridge, fuzz, and boundary
coverage remains included in the owner slice, without a full-workspace-green
claim.

The package driver used
`/private/tmp/wave100-native-table-headers-source.key` (500,128 bytes,
SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`) and
produced a 466,448-byte changed candidate (SHA-256
`eb51188fefe44f13001bb2db24fdc4fde4f4bdc208dbc98f6e8b2e4a4a0b1b28`), an
exact inverse, and an exact no-op. Candidate semantic reread matched banding
enabled, fit-cell-content row sizing, and all gridlines hidden; diagnostics
were changed=true, touched_components=3, deleted_previews=3, and
full_reparse=true, with 83 members after deleting three previews.

Computer Use opened a disposable candidate copy in Keynote without repair,
recovery, or conversion UI and showed alternating rows on, resize-to-fit on,
and all five gridline checkboxes off. Keynote normalized only that disposable
copy on close, changing the 466,448-byte copy from the candidate SHA-256 to
496,491 bytes with SHA-256
`280c57e3e997ef2380edd7759fc2ff17f7df9ee09fbdcbec2978b4b4ce751f57`;
the canonical source, candidate, and inverse stayed untouched. No native
mutation/save/reopen acceptance or byte-exact UI-save claim follows.

The migration host, `litchi-iwa -> litchi-keynote` edge, all 13 ordered
migration debts (especially 014, 015, and 017), generated/native
compatibility owners, and the IWA monolith deletion gate remain unchanged.
Wave103 retires no crate, edge, debt, manifest dependency, host adapter,
generated owner, or monolith owner.

## 2026-08-27 amendment: Wave104 current-topology Pages body-table appearance owner (not a monolith-exit gate)

`litchi-pages` now owns selector-first body-table appearance reads and the
supported exact-source package transaction through
`Package::body_table_appearance`, `edit_body_table_appearance`, and
`apply_body_table_appearance`. Callers use `BodyTableSelector` and the
archive-free `table::appearance::Appearance`; native IDs, archives, ZIP/wire
values, and generated models do not cross this boundary. The focused owner
admits canonical model/style routes, resolves the tested preset/network/default
read paths, allows exact no-ops on locked sources, and rejects unproven
ArchiveInfo, opaque metadata, role, and global style-inbound facts before
publication. Legacy Pages physical table/content ownership remains in
`litchi-iwa`.

The focused Package owner is strict for all of its reads and changed writes.
Separately, `litchi-iwa` legacy table listing retains one private read-only
appearance helper for compatibility; it does not attempt a focused Package
read, perform mutation, or bypass a focused Package write failure. Raw
`PagesEditor` mutation APIs remain retired, and the appearance example uses
the focused Package owner.

Wave104 verification is 21/21 for the focused Pages integration, 18/18 for
the appearance codec, 614/614 for the full boundary suite, and clean
`py_compile`; strict Pages and protos Clippy, the `litchi-iwa` library and
appearance example checks, and the eight-seed Pages fuzz-target check passed.
The HOST/FACADE/RESOURCE live audits returned no findings. These counts are
scoped to the Wave104 boundary and do not imply full-workspace health.

The native source `/private/tmp/wave104-pages-native-source.pages` (108,776
bytes, SHA-256
`997509fda639f5dcdabd4546c392b9ebdc3d8c7e9c1d967f35b9f2a0aca38359`, 43
members) failed strict selector read with `InvalidSource` at
`Table { table: 0 }`; it remained byte-exact and produced no candidate or
inverse. No native Pages mutation, save, reopen, or UI acceptance claim is
made. The operation budget is logical rather than Package-cache,
decompressed-Archive, allocator, or RSS telemetry.

The `litchi-iwa -> litchi-pages` edge, all 13 ordered migration debts
(including 017), generated/Prost/Buffa owners, and the monolith deletion gate
remain unchanged. Wave104 retires no crate, edge, debt, manifest dependency,
host adapter, generated owner, or monolith owner.

## 2026-08-28 amendment: Wave105 current-topology Keynote slide-table persisted lock owner (not a monolith-exit gate)

`litchi-keynote::Package` is now the selector-first owner for persisted
slide-table lock reads and lock-state read/edit/apply transactions. Its public
`State`, patch, edit, commit, diagnostics, error, limit, path, and selector
values are archive-free and do not expose native IDs, Archive/ZIP/member data,
wire views, or generated/Prost/Buffa types. Strict `table_info_codec` parsing
and prepared lock rewriting preserve unrelated fields and fail closed on
malformed, ambiguous, or unsupported sources while admitting exact lock and
unlock transitions.

The legacy Keynote editor no longer owns persisted title, sort, or lock
configuration methods/calls. `litchi-iwa` retains the physical `Sort Now` row
executor only; it may run after focused Package persisted `Order` and lock
admission, and it has no lock fallback. Every focused lock error propagates.
Physical row/cell/storage/formula/tile and broader table compatibility remain
in the legacy host.

Wave105 gates are 17/17 for the lock codec, 9/9 for focused Keynote lock
integration, 30/30 for the IWA slide-table suite, four migrated Keynote
example checks, and 625/625 boundary tests with passing `py_compile`. Strict
Keynote/protos checks and Clippy passed, as did both isolated fuzz checks (15
codec seeds and eight lifecycle command seeds). Scoped title/sort/lock audits
are empty; the live full checker reports only three unrelated untracked Pages
table-lock violations. No full-workspace result is implied.

The only lock driver attempt used the temporary-only compiled driver against
`/private/tmp/wave100-native-table-headers-source.key` (500,128 bytes,
SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`, 86
members). Strict Package read at slide 0/table 0 returned `InvalidSource`
before mutation; source bytes/hash stayed exact, with no candidate, inverse,
or UI run. Native mutation/save/reopen acceptance is withheld.

This is a topology and scoped-gate record, not allocator/RSS telemetry: the
Package/SourceCatalog cache, decompressed Archives, process allocator, and
codec-internal allocator behavior are not observable, and no zero-copy or RSS
claim follows. No crate, manifest dependency, dependency edge, or migration
debt is closed. The `litchi-iwa -> litchi-keynote` edge, all 13 ordered debts
(including 014, 015, 016, and 017), generated/Prost/Buffa owners, and the IWA
monolith deletion gate remain unchanged.

## 2026-08-28 amendment: Wave108 current-topology Keynote slide-table dimension owner (not a monolith-exit gate)

Wave108 moves selector-first slide-table dimension discovery and editing into
`litchi_keynote::Package` through
`slide_table_dimension_size`, `edit_slide_table_dimension_size`, and
`apply_slide_table_dimension_size`. The public semantic surface exposes only
archive-free `slide::table::dimension::{Dimension, Points, Size}` and typed
transaction values. Native identifiers, Archive/ZIP/member data, raw wire
views, and generated/Prost/Buffa values remain private.

The owner performs one atomic storage-plus-geometry transaction: the selected
`HeaderStorageBucket` dimension and the matching `TableInfo` drawable
geometry are rewritten together, with strict role, ArchiveInfo/FieldInfo,
metadata, global-inbound, and persisted-lock checks before publication.
One conservative logical budget spans source admission, candidate reopen and
semantic reread, locality checks, exact inverse, and stale/conflicting patch
rejection. `litchi-iwa` retains physical resize/geometry, rows/cells,
storage, formulas, tiles, and other legacy compatibility duties.

The global physical census enforces UUID-pair uniqueness and rejects all-zero
identifiers before publication. This is scoped authority validation for the
dimension route; unrelated metadata assignments remain outside this slice.

Raw persisted-dimension host methods and wrappers are retired, as is the
obsolete Keynote appearance bridge. Two remaining legacy geometry/remove and
archive-name lookups use `KeynoteObjectCatalog`; physical helpers and legacy
graph compatibility remain in the migration host.

The scoped evidence is a passing Keynote owner library check and strict
library Clippy; 12/12 focused dimension tests with strict test Clippy; 30/30
IWA slide-table host tests; passing `create_keynote_table` and
`list_keynote_tables` example checks; and a passing lifecycle fuzz-target
check plus strict target Clippy with 10 command-only seeds. Focused dimension
boundary checks passed 5/5, the full boundary unit suite passed 653/653 with
passing `py_compile`, and live facade/host audits were empty. The top-level
checker remains blocked only by three unrelated pre-existing untracked Pages
table-lock findings; these are scoped results, not full-workspace health.

Native positive Keynote dimension evidence is withheld. The one authorized
read used `/private/tmp/wave100-native-table-headers-source.key` (500,128
bytes, SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`, 86
members) and strict Row selection returned `UnsupportedDependency`. Its
source bytes remained exact; no candidate or inverse was produced, no UI run
occurred, and no repository files changed.

Accounting remains a conservative logical operation envelope. Package and
SourceCatalog caches, decompressed Archives, process allocator/RSS, and
codec-internal allocation telemetry are outside this owner's observations;
there is no zero-copy, allocator/RSS, or package-wide performance claim. The
authoritative topology remains 64 workspace packages, 239 internal
dependency declarations, one migration host, and 13 ordered debts. The
`litchi-iwa -> litchi-keynote` edge, generated/Prost/Buffa owners, and IWA
monolith deletion gate remain unchanged; Wave108 closes no package, edge,
debt, or monolith gate.

## 2026-08-28 amendment: Wave106 current-topology Numbers FormulaArchive extractor projection (not a monolith-exit gate)

Wave106 changes one remaining legacy Numbers extractor route. The
`litchi-iwa::numbers::table_extractor::TableDataExtractor` formula sidecar
retains strict bounded owned wire bytes and renders through the neutral
`numbers_formula_codec` scalar and compatibility visitors. Production
generated `tsce::FormulaArchive` decoding and generated AST construction are
removed from this extractor; the old generated renderer remains only as a
`cfg(test)` differential oracle. The sidecar owns bytes for preservation and
is not a zero-copy view.

Canonical but incomplete postfix programs deliberately retain legacy output:
operand-less negation becomes `=FORMULA()` and surplus expressions select the
final expression. Malformed wire, duplicated known fields, and noncanonical
known values remain rejected.

`FormulaOwnerDependencies`, category maps, and name maps remain permissive
generated best-effort compatibility products. Formula authoring, cache
refresh, physical mutation, cloning, merge handling, and dependency shifting
remain generated compatibility responsibilities. The focused extractor cut
therefore does not make the Numbers formula graph or the legacy host
generated-free, and it does not move all formula ownership out of
`litchi-iwa`.

One `ProjectionBudget` covers reference-map census, sidecar admission,
repeated renders, and table extraction; formula decode-report fields, work,
and text usage are merged into it. This is conservative logical accounting,
not Package/cache, decompressed-Archive, allocator, RSS, or codec-internal
telemetry. No zero-copy or package-wide performance claim follows.

The scoped evidence is 38/38 focused extractor tests; green `litchi-iwa`
library check, no-run, and strict Clippy gates; a passing formula fuzz binary
check and strict Clippy gate with a 14-seed smoke run completed for 100 runs;
the dedicated FormulaArchive extractor boundary audit returned no findings;
and the full boundary unit suite passed 643/643 with `py_compile`. These are
scoped verification gates, not a full-workspace result.

Native read evidence is limited to the Numbers-created
`/private/tmp/wave106-numbers-formula-source.numbers` (136,591 bytes,
SHA-256
`81fe99b6647e370d1b1663c703500f1fac239111ae7b4bea862f74803c94208c`, 43
members). It contains one 22-by-7 `Table 1`; the migrated extractor read six
materialized cells and found `=(B3+C3)` at zero-based `(1,3)` and
`=SUM(B4:C4)` at `(2,3)`. No mutation, candidate, inverse, save/reopen, or UI
acceptance evidence is claimed.

The authoritative topology remains 64 workspace packages, 239 internal
dependency declarations, one migration host, and 13 ordered debts. No
current edge is removed; debts 010, 015, and 016 remain open, as do every
other ordered debt. The generated/Prost/Buffa owners and IWA monolith
deletion gate remain unchanged.

## 2026-08-28 amendment: Wave107 current-topology Pages body-table name owner (not a monolith-exit gate)

Wave107 places selector-first Pages body-table name reads and rename
transactions in `litchi_pages::Package` through
`Package::{body_table_name, edit_body_table_name, apply_body_table_name}`.
The exposed `table::name::Name` and transaction values are archive-free;
native object identifiers, Archive/ZIP/member data, wire views, and
generated/Prost/Buffa values remain private. The strict
`table_model_discovery_codec` field-8 borrowed projection and prepared
rewrite preserve unknown fields/groups and enforce bounded malformed,
duplicate-known, noncanonical, and wrong-wire rejection.

The raw `PagesEditor` rename mutation route is retired and the example uses
the focused package API. `PagesEditor::tables()` intentionally remains a
legacy generated read-only name-listing compatibility route outside focused
owner admission; there is no mutation fallback. The legacy host retains
physical Pages table/content, storage, formula, and related compatibility
responsibilities.

Wave107 verification is 22/22 for the focused Pages body-table-name
integration, 38/38 for host Pages table tests, 9/9 for the focused codec, and
565/565 for the full `litchi-iwa-protos` library suite. Strict Pages owner and
test Clippy passed; the `edit_pages_table` example check passed; lifecycle
fuzz check and strict target Clippy passed with 10 command seeds. Boundary
unit tests passed 648/648 and live name HOST/FACADE audits were empty. The
full checker is blocked only by an unrelated pre-existing untracked Pages
table-lock file; these remain scoped results.

Metadata admission proves the selected model's unique current component and
locator and rejects unknown, external, data, ambiguous, and root-map routes.
Other UUID bits and component assignments are opaque-preserved and are not
independently verifiable in this owner.

Native positive Pages name/rename evidence is withheld. The only available
body-table source,
`/private/tmp/wave104-pages-native-source.pages` (108,776 bytes, SHA-256
`997509fda639f5dcdabd4546c392b9ebdc3d8c7e9c1d967f35b9f2a0aca38359`, 43
members), returned `InvalidSource { path: Table { table: 0 } }` on strict
selector read. It remained byte-exact; no candidate, inverse, or UI run was
produced.

Accounting is a conservative logical operation envelope, not telemetry for
Package/SourceCatalog caches, decompressed Archives, the process allocator,
RSS, or codec-internal allocation behavior. No zero-copy or package-wide
performance claim follows. The `litchi-iwa -> litchi-pages` edge, debt 017
and all 13 ordered debts, generated/Prost/Buffa ownership, and the IWA
monolith deletion gate remain unchanged.

## 2026-08-28 amendment: Wave109 Keynote slide-table name owner (topology remains unchanged)

Wave109 records selector-first litchi_keynote::Package ownership of the
archive-free slide::table::name::Name value and exact name transactions
through Package::{slide_table_name, edit_slide_table_name,
apply_slide_table_name}. The strict table_model_discovery_codec field-8
prepared rewrite preserves unknown canonical fields/groups and rejects
malformed, duplicate-known, noncanonical, or wrong-wire input before
publication. Native identifiers, Archive/ZIP/member data, wire views, and
generated/Prost/Buffa values remain private.

The admitted route is rooted and unique: slide -> table-info -> model ->
storage/name. Strict storage-route, role, ArchiveInfo, metadata/current-
component, UUID, and global-inbound authority checks run before staging;
unsupported, ambiguous, aliased, locked, malformed, or otherwise unproven
routes fail closed. Focused transaction tests verify source and unrelated-byte
preservation, candidate semantic readback/locality, and exact inverse
restoration including canonical preview state. The raw Keynote rename mutation
path is retired with no fallback. Physical rows/cells, storage, formulas,
tiles, and broader table compatibility remain in litchi-iwa.

The scoped gates are 18/18 focused slide-table-name tests; strict
litchi-keynote library and test Clippy PASS; 30/30 IWA host slide-table tests
and litchi-iwa library check PASS; slide-table-name fuzz-target check and
strict target Clippy PASS with 10 command-only hex seeds; 662/662 boundary
unit tests; py_compile PASS; full checker PASS; and live
FACADE=[]; RESOURCE=[]; HOST=[] audits.

No positive native Keynote name/rename acceptance is claimed. The strict read
of /private/tmp/wave100-native-table-headers-source.key (500,128 bytes,
SHA-256 47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b,
86 members) returned InvalidSource. The source remained exact; no candidate
or inverse was produced and no UI run occurred, so native mutation, save,
reopen, and UI acceptance evidence is withheld.

One conservative logical operation ledger is recorded. Package cache,
decompressed-Archive, OS/process allocator, and RSS telemetry are
unobservable; no zero-copy, allocator/RSS, or package-wide performance claim
is made. This amendment closes no crate, manifest, dependency edge, debt,
host owner, generated owner, or monolith gate. The authoritative topology
remains 64 packages, 239 internal dependency declarations, one migration host,
and 13 ordered debts; litchi-iwa -> litchi-keynote/debt 014, all other
edges/debts, generated/Prost/Buffa ownership, and the IWA monolith deletion
gate remain unchanged.

## 2026-08-28 amendment: Wave110 Numbers dimension bucket wire ownership

The private `litchi-iwa` Numbers dimension-storage adapter now delegates
header-bucket wire reads and edits to the generated-free
`table_dimension_codec`. It uses a finite streaming projection and the
prepared plan -> requirements -> single execute path, preserving unknown raw
fields and ordering while rejecting malformed, duplicate, out-of-range,
non-finite, negative, noncanonical-zero, hash-mismatched, or role-aliased
bucket sources before publication. External storage references, incorrect row
bucket counts, row/column storage aliasing, and cross-slot row records also fail
closed. The codec's internal header-size rescan now inherits the caller's
finite options instead of constructing unlimited ceilings.

This is intentionally narrower than a focused owner migration. The outer
legacy path still uses generated `TableModelArchive` selection and does not
claim package-wide shared-bucket authority, selector transaction artifacts,
candidate reopen/locality, or raw-ID retirement. Public crate topology and
dependency direction do not change.

Verification passed 58/58 focused codec tests, protos check/strict Clippy,
five Wave110 IWA regressions, IWA library check/strict scoped Clippy, 667/667
boundary unit tests, `py_compile`, and an empty live storage-codec audit. The
full checker reported only the three unrelated findings from the pre-existing
untracked Pages table-lock file. No native/UI evidence is claimed.

The finite limits are logical envelopes rather than package-cache,
decompressed-Archive, ZIP/Snappy, allocator, RSS, or zero-copy telemetry. The
authoritative topology remains 64 packages, 239 internal dependency
declarations, one migration host, and 13 ordered debts. No crate, manifest,
edge, debt, public owner, generated owner, or monolith gate closes here; the
remaining generated/Prost/Buffa ledger stays open.

## 2026-08-28 amendment: Wave111 current-topology Keynote movie-transform extension

`litchi-keynote` remains the selector-first owner for admitted existing
file-backed slide-movie geometry, now including the supported native transform
controls. `MovieGeometry` still contains only archive-free position and
displayed size; archive-free `MovieTransform` contains finite rotation
degrees and the supported reflection state, and `MovieFlipAxis` describes the
two typed Arrange operations. The existing geometry package transaction reads,
edits, applies, and inverses geometry and transform together, with strict
source identity, prepared codec execution, candidate validation/reread,
preview invalidation, and locality checks. Unknown wire spans and unrelated
native flag bits remain preserved and private.

The four raw Keynote host geometry/restore/flip entry points and their fallback
are retired. Typed host bridges delegate to the focused package and propagate
unsupported, malformed, locked, limit, and conflict outcomes; private physical
movie helpers remain for unrelated creation/removal, media, offset, property,
and graph responsibilities. This does not make those broader responsibilities
part of the focused owner.

The frozen scoped report contains 27/27 focused package tests and 10/10
focused transform-codec tests. Both the low-level
`keynote_movie_geometry_codec` and package-level
`keynote_slide_movie_geometry` fuzz targets passed isolated cargo checks and
strict target Clippy, and the bounded 32-run smoke passed. The package-level
corpus contains nine command-only seeds, with eight new command seeds across
the package and low-level targets. The final completion ratchet passed its
focused 8/8 and full 670/670 boundary checks; `py_compile` passed and the live
audits were empty. The full checker reported only the three unrelated findings
from the pre-existing untracked Pages table-lock file. A broader
`litchi-keynote` library sweep was 152/153, with the sole `soundtrack_order`
failure unrelated to Wave111. No native Wave111 source, candidate, inverse, or
UI acceptance exists; the historical Wave87 geometry record, including its
strict normalized reread `InvalidSource`, is preserved without being
reclassified.

Wave111 changes no public crate topology or dependency direction. The
64-package inventory, 239 internal dependency declarations, 13 ordered debts,
the `litchi-iwa -> litchi-keynote` edge/debt 014, all other edges and debts,
generated-schema/Prost/Buffa ownership, migration hosts, and the IWA monolith
gate remain open. Resource reporting remains a conservative logical operation
envelope, not package-cache, decompressed-Archive, ZIP/Snappy, allocator, RSS,
or zero-copy telemetry.

## 2026-08-30 amendment: Wave112 current-topology Keynote chart-axis-title owner (not a monolith-exit gate)

Wave112 gives `litchi-keynote` selector-first ownership of existing slide-chart
axis-title reads and transactions. `Package::{slide_chart_axis_title,
edit_slide_chart_axis_title, apply_slide_chart_axis_title}` exposes semantic
category/value titles through `SlideSelector`, `ChartSelector`, and `Axis`,
along with typed `ChartAxisTitleEdit`, `ChartAxisTitleCommit`,
`ChartAxisTitlePatch`, diagnostics, limits, and errors. Set and clear edits are
atomic and exact-source checked: no-op and conflict handling, reversible
patches, candidate reopen/readback, package-locality checks, and stale root
preview invalidation remain part of the focused transaction. Native object
identifiers, archive names, generated values, and wire views remain private;
only the selected primary axis payload is rewritten, while the chart graph,
opposite and secondary axes, unrelated members, and unknown wire spans remain
preservation boundaries.

The admitted route requires a rooted, uniquely selected chart graph with its
title stand-in, chart non-style, unlocked drawable, primary axis role, and
selected metadata and inbound-reference authority. Both source-built
same-component graphs and Keynote-normalized `DocumentStylesheet` graphs are
supported. Cross-component objects require matching physical stylesheet
registration plus current, non-weak PackageMetadata external references in the
proven owner direction; unrelated data references remain valid, while numeric
data/object collisions, foreign ownership, malformed, duplicate, wrong-wire,
noncanonical, role-aliased, locked, ambiguous, or otherwise unproven sources
fail closed before staging. The `litchi-iwa` Keynote host now delegates
axis-title read/set/remove operations through the semantic selector bridge; the
former raw-ID axis-title CRUD route has no fallback. Other native chart,
storage, formula, tile, and graph duties remain in `litchi-iwa`.

`litchi-iwa-protos::keynote_chart_axis_title_codec` owns only the four selected
fields (13/14 visibility and 15/16 text) from the generated
`TSCH.Generated.ChartAxisNonStyleArchive` extension. Its private Buffa
lazy-view projection is cross-checked by strict borrowed preflight and uses a
presence-preserving wire-local rewrite with finite limits and the prepared
plan -> requirements -> single execute path. The outer axis envelope and raw
source bytes remain caller-owned; generated/Prost/Buffa values do not cross the
semantic package boundary.

Recorded scoped evidence is 23/23 codec cases, 16/16 focused
`litchi-keynote` chart-axis-title cases, 10/10 chart-title cases, 3/3 focused
`litchi-iwa` host regressions, 678/678 boundary unit cases, strict Clippy for
the codec and package targets, and fixed-corpus fuzz smoke over 20 codec and 8
package seeds. The full boundary command retains three unrelated pre-existing
untracked Pages table-lock findings, and the broader Keynote suite retains its
pre-existing soundtrack-order failures; these results are scoped and do not
establish a full-workspace-green claim.

Real Keynote verification opened the source-built `Wave112 Revenue` candidate,
observed its category/value axes through the accessibility tree, saved it
natively, closed it, and reopened it without repair. The focused owner then
read the native-normalized `DocumentStylesheet` graph, proved an exact no-op,
changed the value title to `Native Roundtrip`, restored the exact native bytes
through the inverse patch, and Keynote opened that changed artifact with the
new value-axis title. The source-built and native exact-inverse SHA-256 values
are respectively
`74a1876ab0b286a7ebc610e53b452e3a4c8cf8e779c1ec781aaa8f0e25803b31`
and
`87db4dc036ece68c27144ab0303d54517b02dc433b839337b598fdd677788a8c`;
the disposable files are not checked-in fixtures.

Finite resource reports remain conservative logical operation envelopes, not
Package/SourceCatalog cache, decompressed-Archive, ZIP/Snappy, process
allocator, RSS, or zero-copy telemetry. Workspace package and dependency
topology do not change: the authoritative inventory remains 64 packages, 239
internal dependency declarations, one migration host, and 13 ordered debts.
The `litchi-iwa -> litchi-keynote` edge/debt 014, all other edges and debts,
the generated-schema/Prost/Buffa ownership ledger, and the IWA monolith
deletion gate remain open. Wave112 closes no workspace crate, production
manifest dependency, edge, debt, or monolith gate.

## 2026-08-30 amendment: Wave113 current-topology Numbers table-title host cleanup

Wave113 completes the retirement of the dead Numbers table-title seam in
`litchi-iwa`. The private `numbers::editor::table_title` module and its wire
submodule are removed, together with the host's table-title helper exports and
the `numbers::editor::Settings` alias for
`litchi_numbers::table::title::Settings`. No raw-ID table-title reader,
writer, compatibility alias, or host fallback remains. The canonical
selector-first `litchi-numbers` table-title package owner and its archive-free
`table::title::Settings` value remain; this cleanup does not remove the
cross-format table-title paths used by Pages or Keynote.

Table-title fuzzing is now deliberately two-layered. The low-level
`litchi-iwa-protos::numbers_table_title_codec` target exercises bounded
projection, all proto2 presence states, reference and IEEE-754 scalar reads,
unknown spans, malformed-input rejection, scalar/report agreement, and exact
typed decode limits. The package-level `litchi` target exercises semantic
selector reads and the table-title no-op, changed, inverse, conflict, preview
locality, candidate reopen, and bounded-ingress paths against package inputs.
Both layers keep native identifiers, generated/Prost values, Buffa views, and
source artifacts inside their respective private boundaries. Fixed-corpus
smoke passes 25/25 codec recipes and 5/5 package command recipes; this is not a
claim of sanitizer-backed fuzzing, exhaustive coverage, or performance/RSS
telemetry.

Scoped evidence also includes 9/9 codec unit cases, 2/2 focused package-owner
unit cases, 6/6 table-title integration cases, the `litchi-iwa` all-target
check, 679/679 boundary unit cases, strict Clippy for the codec, focused
Numbers package/test, and both fuzz targets, and the leaf/root public-API and
Numbers dependency audits. The live boundary checker has no Wave113 finding;
its only three findings are from the pre-existing untracked Pages table-lock
file.

Computer Use supplied the previously missing explicit-outline evidence. The
source package changed from `visible=Some(true), outlined=None` to
`visible=Some(true), outlined=Some(true)`, opened in Numbers without repair,
and exposed `Title` and `Outline Table Title` checkboxes both at value `1`.
Numbers saved a native copy, closed it, reopened it without repair, and again
reported both controls at value `1`. The focused reader then observed the
native copy as `Some(true)/Some(true)`, proved an exact no-op, changed the
outline to absence, and restored the exact native bytes through the inverse.
The source/no-op/source-inverse SHA-256 was
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`;
the focused outlined candidate was
`13b812ec056d6358ae44772c9b0db957f23c57d033991b39f9f002c71331558e`;
and the native-save/no-op/native-inverse SHA-256 was
`af9f6138949bc7ba2c752c2b2500998e1307a3e247fe1ae56aaf010ea165daf1`.
These disposable UI artifacts are evidence, not checked-in fixtures.

This is a focused host-module and alias cleanup, not a crate-topology change.
The authoritative inventory remains 64 workspace packages, 239 internal
dependency declarations, and 13 ordered migration debts. No workspace crate,
production dependency edge, ordered debt, format owner, generated-schema/
Prost/Buffa owner, or monolith gate closes in Wave113.

## 2026-08-30 amendment: Wave114 current-topology Keynote value-axis owner

Wave114 moves the Keynote primary value-axis aggregate (bounds, steps, and
scale) into a selector-first `litchi-keynote` package owner. The owner exposes
typed settings and an atomic edit/patch/commit path without raw identifiers or
low-level objects, and shares chart-axis graph authority with the existing
axis-title owner. The former six raw-ID host methods and the private
`axis_bounds`, `axis_steps`, and `axis_scale` modules are retired; no host
fallback remains.

The focused codec is private Buffa with strict raw-wire preflight and a bounded
eager borrowed-view exception for generated non-style fields 5, 6, 8, 17, and
18. Field 4 (decades) is strictly validated and preserved. Source spans remain
the rewrite authority, with exact no-op and inverse behavior, selected-message
and ZIP/member locality checks, and finite transaction resource accounting.

Scoped evidence is 13/13 common-axis, 18/18 codec, 17/17 production-guard,
20/20 value-axis, 16/16 axis-title, 10/10 chart-title, and 1/1 typed-API
tests; 687/687 boundary unit cases; low sanitizer smoke over 28 files and 29
executions; and high fuzz smoke over 10 seeds and 11 runs. The live boundary
checker reports only the three known unrelated untracked Pages table-lock
findings. These results are scoped and do not claim a full-workspace-green
build.

Computer Use verified the native Keynote artifact after save, close, and
reopen without repair: the value axis showed `Logarithmic`, Min `1`, Max
`120`, and Decades `2`. The source, focused edited, and native-save SHA-256
values are respectively
`74a1876ab0b286a7ebc610e53b452e3a4c8cf8e779c1ec781aaa8f0e25803b31`,
`10d215489857fff5c3bc93dfd68e931d8282aa069b2cefacd03fe0b687f573d6`, and
`783c1f750d2012c25186b4a4379fba0a4cdf13c75e976139f6ab37b43c4acf12`.
Native no-op output retained the same SHA and reported `changed=false`.

This remains a focused owner and host-seam migration, not a topology or debt
closure. The authoritative inventory remains 64 workspace packages, 239
internal dependency edges, 13 ordered debts, and one monolithic IWA host.
The `litchi-iwa -> litchi-keynote` edge and debt 014/open monolith-exit edge
remain open; no topology, debt, or monolith gate closes in Wave114.

## 2026-08-31 amendment: Wave115 current-topology Keynote table-title compatibility-alias cleanup

Wave115 removes only the `litchi-iwa` Keynote table-title compatibility alias
module and its `KeynoteTableTitleSettings` re-exports. The selector-first
`litchi-keynote` table-title owner and its private Buffa/lazy codec remain
unchanged. Focused tests and the table-creation example now use the focused
`Settings` type directly from `litchi_keynote::slide::table::title`.

This narrow alias cleanup records no dependency, debt, or workspace-topology
change. Native evidence is deliberately scoped: Computer Use opened the known
5-by-4 Keynote table fixture, created
`/private/tmp/litchi-wave115-title-alias-native.key` with Keynote's Save As,
enabled both `Title` and `Outline Table Title`, saved, closed, and reopened it
without a repair or recovery prompt. The reopened controls remained enabled;
the native file was 502,679 bytes with SHA-256
`57f8b172b7dbf741b73c12a6c66123e7343b82961270bac2a62e613af1f5f608`.
This is native UI persistence evidence, not a library-emitted-output claim: the
legacy table-creation example compiles but still returns
`UnsupportedDependency` before writing a file.

## 2026-08-31 amendment: Wave116 current-topology package-store edge retirement

Wave116 removes the direct normal `litchi-iwa -> litchi-iwa-package` edge and
ordered debt 009. The host's remaining entry, patch, change-kind, and store-
error uses now resolve through doc-hidden renamed re-exports owned by
`litchi-iwa-archive::package`. They preserve the exact neutral package types;
there are no wrappers, conversions, allocations, cache changes, or runtime
behavior changes. The physical archive `Entry` remains distinct from the
neutral `PackageEntry` alias.

`litchi-iwa-package` remains a workspace leaf with canonical inbound edges
from `litchi-iwa-archive` and `litchi-iwa-detect`; the archive edge is now the
host's composition boundary. The deprecated raw package facade and its
transactions remain in `litchi-iwa`, as do the other concrete-format and
monolith-exit responsibilities. Buffa/lazy codecs and source-byte authority
are unchanged by this dependency-only routing slice.

The authoritative current topology is 64 workspace packages, 238 internal
dependency declarations (167 required normal, 60 optional normal, and 11
development), 226 canonical edges, 12 ordered migration debts, and one
migration host. Remaining debt orders are
`[1, 2, 4, 5, 8, 10, 12, 13, 14, 15, 16, 17]`. Scoped verification passes
1 alias-identity, 10 package-state, 9 package-state atomicity, 44 host package,
1 host-only raw roundtrip, and 690 boundary cases. Cargo metadata plus source
and AST audits find no direct host reference to `litchi-iwa-package`.

As a preservation-only native gate, the host no-op replay/inverse emitted the
5-by-4 Keynote table fixture byte-exact at 500,128 bytes and SHA-256
`47cf0d9648ed9e189f03d5f6e66047d3fa94340e2b0ba89d312e683455bb563b`.
Keynote opened it without repair and displayed the table. A native Save As
copy was 500,006 bytes with SHA-256
`18aaa4042124fd5cc5d94fb37e6b8b2f8b4f4f7fe3d5bd44dc50e2487d0cdfdd`
and closed/reopened without repair. This does not prove the Cargo edge itself
and is not a full-workspace-green or performance claim.

## 2026-08-31 amendment: Wave117 current-topology Keynote chart-legend visibility owner (not a monolith-exit gate)

Wave117 gives `litchi-keynote::Package` selector-first ownership of the
effective visibility of an existing slide-chart legend through
`Package::{slide_chart_legend_visible, edit_slide_chart_legend,
apply_slide_chart_legend}`. The archive-free transaction values are
`ChartLegendVisibility{Edit,Patch,Commit,Diagnostics,Error,LimitKind}`;
callers use semantic `SlideSelector` and `ChartSelector` values. An absent
native field has effective visibility `false`, while its presence distinction
remains source-private. Native chart/non-style/component/object identifiers,
archives, ZIP members, generated messages, and Buffa views do not cross the
supported boundary.

The hidden `litchi-iwa-protos::keynote_chart_legend_codec` projects only the
generated chart non-style legend field (field 20). Strict wire preflight runs
before the private Buffa lazy view, and the prepared rewrite preserves all
unselected source spans, unknown fields, groups, and ordering while executing
under finite limits. The package owner retains rooted graph selection, exact
no-op/change and inverse behavior, source-checked conflicts, candidate
semantic readback, root-preview invalidation, and package-wide
locality/atomic-publication checks.

The raw-ID `KeynoteEditor` legend-visibility methods are retired with no
fallback. Keynote legend fill, frame, font, shadow, stroke, chart creation,
duplication/removal, and broader graph work remain in the migration host; the
shared private chart-options helper remains used by the Pages and Numbers
legend paths. This is a focused owner boundary, not a replacement for those
physical compatibility responsibilities.

Scoped verification records 13/13 codec tests, 10/10 focused package tests,
the allocation-mapping unit, strict library/test Clippy, all-target checks,
696/696 boundary-policy tests, and a 1,000-run nightly fuzz smoke over 12
checked-in corpus seeds. Computer Use opened the 46,022-byte focused candidate
without repair, observed the chart Legend checkbox off, saved and reopened a
154,754-byte native copy with Legend still off, and the focused API reverse-read
and exactly reproduced that native copy. The candidate and native-copy
SHA-256 values are respectively
`2ecf1327971c206e4017168cfd5fac86b0954ae32a0f2f0ba894f3b6b84c5acd`
and `b82af597b4469b056b559cde678db508fca13d7ec725633bdd82749fab560adf`.
This amendment changes no package or dependency direction. The
authoritative topology remains 64 workspace packages and 238 internal
dependency declarations (167 required normal, 60 optional normal, and 11
development), with 226 canonical edges, 12 ordered debts, and one migration
host. Debt 014 and the `litchi-iwa -> litchi-keynote` edge remain open, as do
the IWA monolith deletion gate and all broader semantic parity work.

## 2026-08-31 amendment: Wave118 current-topology durable publication, focused saves, and playback guard

Wave118 adds durable filesystem publication to the physical archive owner at
`litchi-iwa-archive::publication::replace_with`. It stages beside the
destination, flushes and synchronizes the staged file before replacement,
then applies existing ordinary regular-file permissions through the published
file descriptor and synchronizes it before attempting parent-directory
synchronization where the platform supports it. Unix staging and new
destinations remain `0600`; inherited set-user-ID/set-group-ID bits are
cleared. `litchi-pages::Package`,
`litchi-numbers::Package`, and `litchi-keynote::Package` now own focused typed
`save` APIs and format-specific `SaveError` values; their writers call through
to this archive publication boundary without duplicating physical replacement.

The legacy `litchi-iwa::IWorkPackage::save` route remains as a compatibility
path but is a thin call-through to the same archive publication owner. Its
legacy error mapping and committed-state inspection remain available; it owns
no second staging or publication implementation. Publication rejects missing-
filename, symbolic-link, and non-regular destinations, redacts path and
content details from default errors, and reports when replacement committed
before a post-replacement permission or published-file synchronization failure,
or a parent-directory synchronization failure other than Unix `InvalidInput`/
`Unsupported`; those two kinds are treated as an unavailable capability.
Reparse-point destinations are rejected on Windows. Ordinary local-filesystem
rename semantics are assumed;
network, userspace, unusual-filesystem, and platform directory-sync behavior
may provide weaker durability.

The private Buffa playback projection is fully generation-ratcheted. The
production codec guard keeps its strict raw-wire preflight ahead of the private
lazy view, excludes Prost and owned generated values from production ingress,
and leaves source bytes and spans as the preservation and rewrite authority.
The projection remains format-private and does not publish native identifiers
or generated values.

Computer Use opened the Rust-produced Pages, Numbers, and Keynote artifacts
without repair, saved, closed, and reopened their semantic markers. Their
native-normalized size/SHA-256 pairs are 96,413/
`93b904b95251c8160c71fb3e34cb169f66aff91ac85a70274c07b7942567273c`,
136,023/`8072f6c00c2e530867581104510e2f0eb08821a8a11b2a1158b540182bedfba1`,
and 500,021/`9c8dd8e80ce843d8376ffa90a9904a15f041f71fe436752700a0a7fd3b76c99f`;
each focused API republished its native-normalized artifact byte-for-byte.
This is scoped save interoperability evidence, not broader migration
completion. The authoritative topology remains 64 workspace packages, 238
internal dependency declarations, 226 canonical
edges, 12 ordered migration debts with IDs `[1, 2, 4, 5, 8, 10, 12, 13, 14,
15, 16, 17]`, and one migration host. Wave118 closes no dependency edge or
debt and does not close the IWA monolith gate.

## 2026-08-31 amendment: Wave119 debt-005 archive routing

Wave119 retires ordered migration debt 005. The archive owns a doc-hidden,
explicit, exact-type `iwa` route for the compatibility host. The route exposes
the core boundary by exact type identity; it introduces no wrapper, conversion,
copied value, compatibility implementation, or second owner. `litchi-iwa` no
longer declares or imports `litchi-iwa-core` directly; its existing host paths
consume that archive-owned route.

This is dependency routing only, not an object-level semantic migration.
Host/edge debt 002 and the migration host remain, as do the remaining host
editor/compatibility logic and the other ordered debts. No Buffa or
generated-schema implementation, projection, or budget changes are included.
Computer Use opened the canonical Pages, Numbers, and Keynote fixtures,
created native copies, closed them, and reopened the copies without repair or
recovery UI while retaining their text/date markers and Numbers value `42`.
The format-owned `Package::save` APIs then republished the native-normalized
copies byte-for-byte: Pages 96,407 bytes/
`d321bde90824664eb6122690eacd30441aa5d4d329c655b808e1b420a78e6bb5`,
Numbers 135,985 bytes/
`1e23a5b36e3f11bc0de2b11de37488c4c981f13f7335ef415e63eccb2adedc18`,
and Keynote 499,981 bytes/
`a720f3a1dbe32070a1c72bc710621747b9c305261b86c4ade1879b6a3eadaf02`.
These disposable artifacts establish preservation only, not object-level
semantic migration.

The authoritative post-Wave119 inventory is 64 workspace packages, 237
internal dependency declarations, 226 canonical edges, and 11 ordered
migration debts with orders `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, with
one migration host. Debt 005 is removed; debt 002, the relevant host/edge, and
the remaining host logic stay open.

## 2026-08-31 amendment: Keynote slide-table persisted sort transaction hardening (not a monolith-exit gate)

The existing selector-first Keynote persisted field-44 sort transaction now
reuses the shared `slide_table_core` authority for canonical admission, its
bounded operation-local budget, and its locality checks. Ambiguity in archive
roles, references, types, or wire framing fails closed before publication.
Exact aggregate-only producer metadata and current, unversioned in-package
cross-component edges remain admissible; partial route metadata, dangling or
duplicate edges, and foreign inbound ownership remain rejected. The
source-preserving rewrite continues to preserve previews, field 45,
unknown fields, and untouched archive entries, while retaining exact no-op,
inverse, and conflict behavior.

This remains a focused persisted-configuration hardening only. Physical
`Sort Now` and `RowRange`-based row execution, cells, and table storage remain
owned by the legacy `litchi-iwa` host. No native semantic acceptance,
performance, fuzz-exhaustiveness, or broader Keynote authoring claim follows
from this amendment.

The authoritative topology remains 64 workspace packages, 237 internal
dependency declarations, 226 canonical edges, and 11 ordered migration debts
with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, with one migration host.
No migration debt is retired, no host exits, and no IWA monolith-deletion gate
closes in this amendment.

## 2026-09-01 amendment: Keynote selector-first physical-sort host seam

Keynote physical `Sort Now` remains in `litchi-iwa`, but its normal entry points
are now selector-first: `execute_slide_table_sort_order` and
`execute_slide_table_sort_order_to_rows` resolve the slide and table before
entering the private native-ID executor. Deprecated raw-ID `apply_*`
declarations remain for source compatibility; production calls are rejected by
the topology ratchet. Dormant host persisted-sort set/clear routes and their
wire writers are removed because the focused `litchi-keynote::Package` already
owns that transaction.

The row-order algorithm is now an archive-free hidden primitive in
`litchi-iwa-common`. The compatibility adapter supplies decoded keys through
borrowed BNC views and applies finite dimension/key budgets before staging the
full native transaction. Table models, tiles, headers, UIDs, formulas, comments,
stroke layers, hidden axes, and generated Prost mutation remain in the host;
this cut does not create a focused physical-table package or a Buffa mutation
owner.

The authoritative topology remains 64 workspace packages, 237 internal
dependency declarations, 226 canonical edges, and 11 ordered migration debts
with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, with one migration host.
No package, dependency edge, debt, host, or monolith-deletion gate is removed
by this seam hardening.

## 2026-09-01 amendment: Numbers table-relocation owner (topology unchanged)

Physical relocation of one existing Numbers table between existing sheets now
has a focused owner at
`litchi_numbers::table::relocation::transaction::{Commit, Diagnostics, Edit,
Error, LimitKind, Patch, Path}`. The package surface is
`litchi_numbers::Package::{edit_table_relocation, move_table,
apply_table_relocation}` (`apply_table_move` remains an alias): callers provide a
source `SheetSelector`, a source-sheet `TableSelector`, and a destination
`SheetSelector`; native IDs and raw archive/protobuf values do not cross the
boundary.

The focused transaction owns exact-source patch construction and application,
`Patch::inverse()` restoration, conflict/stale/foreign-source refusal, and
locality proof. A same-sheet selection is an exact no-op and replays without a
write. For a changed move, the native fixture evidence is limited to
`Index/Document.iwa` and `Index/Tables.iwa`; table payload/content/unknowns and
unrelated members are preserved. The legacy
`litchi_iwa::NumbersEditor::move_table` remains a compatibility delegate for
ordinary graphs. Historical host-built storage outside the focused cell
projection uses a doc-hidden selector-first admission function in
`litchi-numbers` that calls the same rewrite/verification engine; the host only
performs legacy candidate readback. Unsupported graphs fail closed, and this
does not add a second physical owner.

This focused seam does not change the authoritative inventory: 64 workspace
packages, 237 internal dependency declarations, 226 canonical edges, 11
ordered migration debts with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`,
and one migration host remain. No package, edge, debt, host, or monolith-
deletion gate is removed by this ownership transfer.

## 2026-09-01 amendment: Keynote chart Arrange focused owner (topology unchanged)

Existing-chart Arrange interaction state now has a focused semantic owner in
`litchi-keynote`. The owner exposes `locked` and `constrain_proportions`
through selector-first `Package` read/edit/apply operations using
`SlideSelector` and `ChartSelector`; native drawable IDs and physical graph
values remain private. The source-bound transaction performs exact-source
patching, inverse/conflict checks, candidate reopen, semantic readback, and
locality-preserving publication. Arrange flags are non-rendering, so previews
are preserved and no preview invalidation route is added.

The legacy `litchi-iwa` chart Arrange entry point remains a compatibility
route. The focused owner has passed the integrated native Keynote gate recorded
in ADR 0008. This slice is
limited to existing charts and does not take ownership of chart legend
visibility, persisted sort configuration, physical table `Sort Now`, chart
data, geometry, or broader chart graph mutation. It adds no normal dependency
edge, retires no migration debt, and removes no migration host or monolith
gate. The authoritative topology remains unchanged.

## 2026-09-01 amendment: Numbers existing-cell Number-format owner (topology unchanged)

The existing-cell decimal Number-format seam now has a focused semantic owner
in `litchi-numbers`, using the existing `cell::data_format::Number` leaf and a
selector-first package transaction. The private adapter retains BNC storage,
format-table/refcount records, native identifiers, Buffa views, archive
members, and publication. The legacy `litchi-iwa` Numbers editor remains a
compatibility route that delegates admitted exact graphs; exact structural
admission failures do not fall through to the generic implementation. Its
deprecated raw-ID setter also fails closed for exact-source cross-family
rejections, while source-built compatibility packages retain the generic
`DataFormat` fallback and their historical replacement behavior. It is not a
second semantic owner.

This is an ownership seam within the current package graph. It adds no normal
workspace package or dependency declaration, retires no ordered migration
debt, removes no migration host, and does not change the monolith deletion
gate. The current authoritative inventory remains 64 workspace packages, 237
internal dependency declarations, 226 canonical edges, 11 ordered migration
debts with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration
host unless a root metadata audit proves otherwise. The focused seam does not
claim ownership of other Number-family formats, cell values/formulas, styles,
table topology, or generated-schema/normal Prost retirement.

## 2026-09-01 amendment: Numbers existing-cell Percentage-format owner (topology unchanged)

The existing-cell Percentage-format seam now has a focused semantic owner in
`litchi-numbers`. It uses the archive-free `cell::data_format::Percentage`
value and selector-first package read/edit/apply transactions. The private
adapter resolves native type 258 through the referenced format-list payload,
shares the audited decimal graph/COW owner with Number, and retains BNC storage,
list keys, refcounts, metadata, Buffa views, archive members, and publication
below the public boundary. Nominal protocol and transaction types prevent
Number/Percentage substitution at compile time.

The deprecated `litchi-iwa` Percentage methods delegate admitted exact graphs
to the focused owner. Every exact-source structural or wrong-family rejection
fails closed; only source-built compatibility packages may use the generic
legacy format implementation. This ownership transfer changes no workspace
package, dependency declaration, canonical edge, ordered migration debt,
migration host, or monolith-deletion gate. The authoritative inventory remains
64 workspace packages, 237 internal dependency declarations, 226 canonical
edges, 11 ordered migration debts with IDs
`[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host. The slice
does not claim ownership of the other display/control families, table topology,
cell values/formulas, or remaining generated-Prost host paths.

## 2026-09-01 amendment: Numbers existing-cell Currency-format owner (topology unchanged)

The existing-cell Currency-format seam now has a focused semantic owner in
`litchi-numbers`, with archive-free Currency values and selector-first package
read/edit/apply transactions. The private adapter resolves native type `257`
through the referenced format-list payload, handles the alternate-number BNC
record and optional secondary generic Number reference, and keeps list keys,
refcounts, native IDs, Buffa views, archive members, and publication below the
public boundary. Exact-source patches, inverses, strict preflight, candidate
reopen/readback, and locality checks remain owner requirements.

The deprecated `litchi-iwa` Currency methods are compatibility delegates for
admitted exact graphs; exact structural, budget, lock, or wrong-family
rejections fail closed, while source-built compatibility packages may retain
the generic route. This ownership seam adds no workspace package, dependency
declaration, canonical edge, or migration-debt retirement, and removes no
migration host or monolith-deletion gate. The authoritative inventory remains
64 workspace packages, 237 internal dependency declarations, 226 canonical
edges, 11 ordered migration debts with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15,
16, 17]`, and one migration host. Operation-specific native Currency E3/E4
evidence is recorded in ADR 0008; it closes no topology or deletion gate.

## 2026-09-01 amendment: Numbers existing-cell Scientific-format owner (topology unchanged)

The current worktree adds a focused Scientific-format owner in
`litchi-numbers`: an archive-free `Scientific` value, selector-first
existing-cell read/edit/apply transactions, a private type-259 strict codec,
and a private package adapter for BNC/format-list/refcount rewrites. Native
identifiers, generated messages, Buffa views, archive members, and raw bytes
remain below the public boundary. The deprecated `litchi-iwa` Scientific
methods are compatibility delegates; source-built packages may retain the
generic route, while exact-source owner failures remain fail-closed.

This is an ownership seam in the existing package graph only. It adds no
workspace package, dependency declaration, canonical edge, or ordered-debt
retirement; it removes no migration host and changes no monolith-deletion
gate. The authoritative inventory remains 64 workspace packages, 237 internal
dependency declarations, 226 canonical edges, 11 ordered migration debts with
IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host.
Scientific build/test/fuzz and native-validation evidence is recorded in ADR
0008; this topology record makes no package-wide or deletion-gate claim. No
topology fact is inferred from source visibility.

## 2026-09-02 amendment: Numbers existing-cell Fraction-format owner (topology unchanged)

The current worktree adds a focused Fraction-format owner in `litchi-numbers`:
an archive-free `Fraction` value with all nine `FractionAccuracy` strategies,
selector-first existing-cell read/edit/apply transactions, and a private strict
native type-262 codec. Handwritten field validation precedes the borrowed lazy
Buffa view. Field 20 (`requires_fraction_replacement`) is preserved when absent
or canonically encoded as `false` (absence remains absent); canonical `true` is
rejected because replacement semantics are not implemented. Native identifiers,
generated messages, Buffa views, archive members, and raw bytes remain below the
public boundary.

The owner binds exact-source patches and inverses to semantic sheet/table
selectors and checked cell positions, uses copy-on-write/refcount closure,
candidate reopen/readback and locality checks, and keeps the deprecated
`litchi-iwa` methods as fail-closed compatibility delegates for admitted exact
graphs. Focused source/build/test/fuzz evidence and operation-specific
native-app E3/E4 evidence for a representative Eighths-to-Hundredths edit are
recorded in ADR 0008. That record does not claim native UI acceptance for all
nine accuracy variants, arbitrary-producer parity, native byte parity after
Numbers normalization, or package-wide performance.

This is an ownership seam in the existing package graph only. It adds no
workspace package, dependency declaration, canonical edge, or ordered-debt
retirement; it removes no migration host and changes no monolith-deletion gate.
The authoritative inventory remains 64 workspace packages, 237 internal
dependency declarations, 226 canonical edges, 11 ordered migration debts with
IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host. The
Fraction slice closes no debt or gate.

## 2026-09-02 amendment: Keynote physical table-sort owner

`litchi-keynote` now contains the focused physical `Sort Now` owner for an
existing slide table. The selector-first transaction consumes persisted sort
order, exposes no public row/value reader, and admits only the canonical
type-6001 table-model route plus explicitly proven tile, data-list, header, UID,
and empty pre-BNC sentinel shapes. Scalar text/number/boolean/date/duration
keys use Rust lexical text ordering and `f64::total_cmp`-based deterministic
ordering for numeric-like values; mixed kinds and unproven row-affine
dependencies fail closed. It moves admitted tile-row envelopes, sparse row
headers, and UID mappings within a body-relative range; strict wire preflight
precedes private lazy Buffa inspection, and exact source patches/inverses,
candidate reopen/readback, locality, and preview invalidation remain the
transaction boundary.

Its one normal `litchi-keynote -> litchi-numbers-wire` edge is the private BNC
scalar adapter used by the owner; no BNC type is re-exported through the
semantic facade. The edge does not transfer package ownership to the shared
wire crate, retire a migration host, or close a debt. A disposable Computer Use
probe opened a pre-hardening candidate, but the current strict owner rejects
the app-authored source because model field 39 identifies an unowned
conditional-style CalculationEngine dependency graph. That run is external
exploratory evidence only. Native Keynote E3/E4
acceptance and artifact hashes for the current owner are intentionally pending,
and earlier host-owned physical-sort evidence is not evidence for this focused
owner.

The current inventory is 64 workspace packages, 238 internal dependency
declarations, 227 canonical edges, 11 development-only edges, 11 ordered
migration debts with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one
migration host. No debt, host, or monolith-deletion gate closes.

## 2026-09-03 amendment: current ownership after raw-facade retirement

The dated soundtrack and Numbers relocation/format paragraphs above remain
historical snapshots. Rooted soundtrack-item read/add/insert/replace/remove is
owned by `litchi-keynote::Package`; the old `litchi-iwa` item implementation
and public item API are absent. Soundtrack creation, broad media/slide-media
work, and remaining compatibility stay in the migration host, so debt 014 and
its host edge remain open. The dedicated public `NumbersEditor` Number,
Percentage, Currency, Scientific, and Fraction convenience routes are absent,
as is public `NumbersEditor::move_table`; generic source-built/cross-format
`DataFormat` mutation and a crate-private populated-sheet relocation bridge
remain. Debt 015 and its host edge therefore remain open.

The deprecated public `litchi_iwa::raw::{bundle, package}` facade, eight raw-only
inspection examples, and its raw-package no-op integration test are removed.
This does not remove the private host bundle/package implementation or the
remaining broad editors. The recursively hardened manifest ratchet still finds
the explicitly admitted nested `crates/litchi-iwa/fuzz -> litchi-iwa` edge;
that edge, hundreds of host-owned operations, host-referenced generation
guards, and debts `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]` continue to block
the deletion gate.

Focused Keynote ownership expands by one E1-only existing-build playback-order
transaction. It has no new dependency edge and does not transfer build effects,
timing, lifecycle, or native acceptance from the host. Keynote background
unknown/image fills now cross the public boundary only as
`Background::Unsupported`; raw preservation remains private. Package-scale
indexing/fallible-allocation hardening in the host and focused format owners
likewise changes no ownership edge.

The current inventory is 64 workspace packages, 238 internal declarations,
227 canonical edges, 11 development-only edges, debts
`[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host. The root
`litchi::iwork` facade remains host-free and bounded/read-only. No dependency,
debt, migration host, or monolith-deletion gate closes.

## 2026-09-03 amendment: reconciled iWork ownership and compatibility

The dated soundtrack snapshot above is superseded for current ownership:
rooted soundtrack-item CRUD is owned by `litchi-keynote::Package`, and the old
`litchi-iwa` item implementation and public item API are absent. Soundtrack
creation and broad media/slide-media compatibility remain host-owned. Public
`NumbersEditor::move_table` is retired while a crate-private duplication bridge
supports the remaining populated-sheet compatibility graphs. Public raw-ID
Number, Percentage, Currency, Scientific, and Fraction convenience routes are
retired; generic source-built/cross-format `DataFormat` compatibility remains.

Host build-order `reorder_slide_builds`, `move_slide_build`, and their wire
helper are removed, with add/update/timing/remove compatibility retained. The
focused Keynote build projection reads the nested modern effect at `[4, 18, 2]`
and its native Apple aliases before legacy fallbacks. This is an E1 semantic
correction only and supplies no native build-order proof. Number/Percentage
staging hardening is limited to allocation and budget safety and changes no
ownership edge.

The current topology is 64 workspace packages, 238 internal declarations, 227
canonical edges, 11 development-only edges, 11 ordered debts `[1, 2, 4, 8,
10, 12, 13, 14, 15, 16, 17]`, and one migration host. The nested host-fuzz
manifest removal closes ADR 0028 deletion gate 1; topology, debt, and host
counts remain unchanged. The separate 2026-09-03 amendment in
[`ADR 0028`](0028-iwa-monolith-exit.md) records the gate accounting and
remaining deletion work.

## 2026-09-04 amendment: Numbers existing-cell Text-format owner

`litchi-numbers` now owns the bounded semantic Text-format transaction for one
existing Numbers cell. The selector-first package API exposes an archive-free
`Text` marker and exact-source read/set/clear/reset/inverse behavior. Its
private type-260 adapter owns native format-list/refcount admission, strict
wire validation, lazy Buffa inspection, copy-on-write publication, candidate
reopen/readback, and physical locality checks; native IDs, generated messages,
archive members, and raw wire values remain private.

Canonical explicit Text attachment uses marker `0x80`. An unchanged admitted
converted-Text source with marker `0x81` and retained Number provenance remains
source-authoritative; numeric-to-Text conversion is not implemented. Focused
tests provide E1 evidence and the checked-in Apple fixture provides E2
read/no-op evidence only. Native E3/E4 acceptance is not claimed.

The dedicated raw-ID Numbers convenience routes for Number, Percentage,
Currency, Scientific, and Fraction remain retired. Generic source-built and
cross-format `DataFormat::Text` compatibility remains host-only, and a focused
Text refusal is not a fallback trigger. This owner adds no package, dependency
edge, debt retirement, migration-host removal, or monolith-deletion-gate
closure; the current topology remains 64 workspace packages, 238 internal
declarations, 227 canonical edges, 11 development-only edges, 11 ordered
debts `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host.

## 2026-09-04 amendment: Numbers Date & Time raw-ID route retirement

The production `NumbersEditor` raw-ID Date & Time convenience routes are now
retired. This removes a dedicated host API surface only. Generic source-built
or cross-format `DataFormat::DateTime`, the broad `TextDateTimeField`
smart-field lifecycle, and attached Pages/Keynote table compatibility remain
host-owned, so the migration host and its manifest/debt topology are
unchanged. No package, dependency edge, ordered debt, or deletion gate is
closed by this amendment.

The authoritative current inventory remains 64 workspace packages, 238
internal dependency declarations, 227 canonical edges, 11 development-only
edges, 11 ordered migration debts with IDs
`[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host.

## 2026-09-04 amendment: Numbers existing-cell Custom-format owner

`litchi-numbers::Package` now owns the bounded selector-first Custom-format
transaction for one existing rooted Numbers cell. The public API is
`Package::{table_cell_custom_format, edit_table_cell_custom_format,
apply_table_cell_custom_format}` with a `SheetSelector`, sheet-scoped
`TableSelector`, and checked `CellPosition`; the semantic value is the
archive-free `Custom` enum with Number, Text, and Date & Time variants.

The private document registry is rooted by `TN.DocumentArchive` field 9 and
message type 222. Its custom archive format discriminators are 270 (Number),
271 (Text), and 272 (Date & Time). Handwritten wire preflight runs before the
private lazy Buffa view, and deterministic source-built exact-source fixtures
provide the original E1 evidence for this amendment. At that historical point
there was no Apple-authored fixture or E2/E3/E4 evidence; the 2026-09-06
follow-up recorded in ADR 0008 and ADR 0028 supersedes that interim status.

Custom edits are exact-source transactions with copy-on-write publication,
candidate reopen/readback, exact inverse patches, physical-locality checks,
unknown/unselected field and member preservation, format-list/refcount
closure, and private UUID handling. Semantically equal registry entries are
reused; shared references remain live, replacement entries receive UUIDs, and
unused entries are culled only after their last reference is cleared. The
focused package suite passes 19/19, the strict custom-format codec passes 8/8,
and both fuzz targets complete 100-run AddressSanitizer smokes. This historical
amendment claimed no native Numbers acceptance or native save/resave evidence,
no generic Custom-format authoring, and no retirement of the legacy
`NumbersEditor` Custom route. The current follow-up admits both registry routes
(`TN.DocumentArchive` field 9/message type 222 and `TN.super` field 8 →
`TSA.custom_format_list` field 12), records operation-specific native
replacement/clear evidence, retires the three dedicated raw-ID conveniences,
and passes 37 focused integration tests with 927 boundary policy tests.
Generic source-built/cross-format `DataFormat::Custom` compatibility and
private attached Pages/Keynote table adapters remain in the migration host.

No workspace package, dependency edge, ordered migration debt, migration host,
or ADR 0028 deletion gate changes by this owner addition. The authoritative
topology remains 64 workspace packages, 238 internal dependency declarations,
227 canonical edges, 11 development-only edges, 11 ordered migration debts
with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host.

## 2026-09-04 amendment: Numbers Duration format groundwork (no owner)

The landed Duration work is limited to a strict native type-268
`FormatStructArchive` codec and narrow Buffa projection. It admits fields 1,
7, 15, 16, and 40; styles `0/1/2` (Colon/Abbreviated/FullNames); and unit
bits `1/2/4/8/16/32` (Weeks/Days/Hours/Minutes/Seconds/Milliseconds). BNC
shape handling records marker `0x0004` for an explicit primary-only Duration
reference and `0x0005` when the shared generic Number secondary is retained.
Wire and source-built host tests cover native round-trips, scalar/formula
preservation, compatibility conversion, format reuse/reset, and both marker
shapes; the codec fuzz target completes a 100-run AddressSanitizer smoke. A
disposable Numbers 14.4 probe supplied the marker/secondary
provenance; it was not an app acceptance or save/resave cycle.

This groundwork does not add a focused `litchi-numbers` package owner or
selector API. Duration remains unsupported in the public Numbers owner matrix,
and the legacy `NumbersEditor`/generic host compatibility route is not retired.
No workspace package, dependency edge, ordered debt, migration host, or ADR
0028 deletion gate changes. The authoritative topology remains 64 workspace
packages, 238 internal dependency declarations, 227 canonical edges, 11
development-only edges, 11 ordered migration debts with IDs `[1, 2, 4, 8, 10,
12, 13, 14, 15, 16, 17]`, and one migration host.

## 2026-09-04 amendment: Numbers existing-cell Duration owner and raw-ID route retirement

The current Numbers topology now includes a focused selector-first owner for
one existing rooted cell's Duration display metadata. The archive-free
`litchi_numbers::Package` API is
`Package::{table_cell_duration_format, edit_table_cell_duration_format,
apply_table_cell_duration_format}`; it keeps native type-268 format records,
BNC kind-4 metadata, list keys, and IWA object identifiers private. The
admitted native payload uses fields 1, 7, 15, 16, and 40, with styles `0/1/2`
and unit bits `1/2/4/8/16/32`. BNC markers `0x0004` (primary-only) and
`0x0005` (retained generic Number secondary) are distinguished; marker-zero
inherited tuples are refused by this owner. The transaction is
metadata-only and preserves values/formulas/caches, unknown and unselected
bytes, copy-on-write/refcounts, locality, and exact inverse/source checks.

The deterministic package/codec fixtures provide E1 synthetic/self-round-trip
evidence. No checked-in Apple-authored fixture provides E2 parse/no-op
evidence. A disposable Numbers 14.4 probe opened, saved, closed, and
reopened both Litchi-mutated candidates without error: the primary-only
candidate `c62fb9ca6e1b86e8a60abde6ebbcdf31edaefd88a09cbc2a65684e4be27372e2`
became native `c8a009d99a6d079f6feff38c73f658e502098c105db3027246f79801fcd1f43e`
with marker `0x0004`, Duration ID 5, type 268, style 1, custom units 1
through 32, and scalar `316310400`; the retained-secondary candidate
`c3a797dc63eb99926c88130318211511e43c6ba979626f77d89f0c1a7765c48e` became
native `04208a942693f19ead2020d1ac1449dce00e9de7df865d88cc8e46847b02714b`
with marker `0x0005`, Duration ID 3, generic Number ID 1, and scalar `86400`.
Strict post-native rereads reported no-op and exact inverse results for both;
the Duration Tile/DataList members remained byte-identical and only unrelated
members normalized across the 43-member packages. This is operation-specific
E3/E4 evidence for the two existing-cell marker shapes only, not broad native
acceptance or package parity. The Apple-authored starting packages are
disposable provenance, not checked-in E2 fixtures and not GUI/Computer Use
evidence. The dedicated raw-ID `NumbersEditor` Duration read/set/reset routes
are retired, but generic source-built or cross-format `DataFormat::Duration`
and attached Pages/Keynote compatibility remain in the migration host;
focused refusals are terminal.

This is an ownership/API-boundary change within the existing graph. It adds no
workspace package, manifest dependency, canonical edge, ordered debt
retirement, migration-host removal, or deletion-gate closure. The authoritative
topology remains 64 workspace packages, 238 internal dependency declarations,
227 canonical edges, 11 development-only edges, 11 ordered migration debts
with IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host.

## 2026-09-05 amendment: Pages body-table hidden-axis owner (topology remains unchanged)

This amendment records a focused ownership/API boundary inside the existing
Pages graph. `litchi_pages::Package::{body_table_hidden_axes,
edit_body_table_hidden_axes, apply_body_table_hidden_axes}` is now the
selector-first owner for one rooted body table, using a checked position or
exact visible name and archive-free `table::hidden_axes::{AxisIndex,
HiddenAxes}` values. `set`, `clear`, `reset`, exact-source commit/apply,
inverse, and semantic no-op behavior are package transactions; native IDs,
archives, protobuf values, and wire payloads do not cross the public value
boundary.

The route first proves the complete body/storage attachment and drawable
closure. It admits one role-qualified type-6000/current or explicitly
qualified type-6003/legacy `TableInfoArchive`, and one role-qualified
type-6001/current or explicitly qualified type-6000/legacy
`TableModelArchive`; the type role is part of the qualification. The model's
base column/row UID reference (field 46) is required. The table-info view UID
reference (field 6) is optional, but when present it must agree with field 46.
The UID map is canonical type 6267, with type 6200 accepted only through its
explicitly qualified legacy archive metadata; the nearby type 6005 message is
not a map owner. Bounded UID permutations must match model dimensions and
owner/reference cardinality must be unique.

The complete type-4008 -> type-6204/type-6220 dependency closure is admitted
and validated for an existing hidden-state owner. An ownerless table reads as
empty and may take an exact empty no-op. For a qualified exact `Indexed`
source, a changed request may now create a bounded owner/dependency closure
after physical, UUID, registry, metadata, and wire-limit preflight; this path
is unavailable to the `NativeVisible` profile and unsupported producer
shapes.

The `TableInfoArchive` active UUID is a read selector for one uniquely
matching stored view. A changed edit refuses multiple stored views, even when
the active UUID selects one unambiguously, because inactive-view preservation
is not proven. The strict source-preserving codec projects only user-hidden
positions; its narrow Buffa sidecar handles singular fields while a bounded
wire walker handles repeated states. Unknown fields/groups, unselected
members, and admitted filtered/pivot markers stay source-owned.

Malformed, duplicate, stale, dangling, wrong-wire, ambiguous, out-of-bounds,
finite-limit, pivot, and unsupported-dependency graphs refuse publication.
For an exact, unlocked, qualified `Indexed` source, a nonempty absent-owner
request now prepares four helper objects for column/row formula and filter
records, adds selected model references through the core canonical
object-reference `FieldInfo` primitive, and updates the strict
`Index/Metadata.iwa` registry/save-token route when present. A valid
metadata-free source is also supported. Physical object/reference/UUID census,
registry collision/misroute checks, and finite retained/output/wire budgets
run before allocation and publication. `NativeVisible` and unsupported
producer shapes remain `UnsupportedDependency`; a non-exact source reports
`UnsupportedSource` before that profile check. `TableLocked` takes precedence
for a changed edit on a locked selected table, while exact no-ops remain
permitted subject to read/graph admission. Existing-owner changes are
copy-on-write, selected-component-local, preview-invalidating,
candidate-reopened/read back, and exactly invertible.

Focused graph/codec, identity/COW, and concurrency test files provide E1
source/self-round-trip coverage. A fresh locked all-features rerun passed 460
Pages library/integration tests across 25 binaries and 802 protocol
library/integration tests (764 library and 38 integration); 170 generated
protocol doctests were ignored. These are focused-package results only. Clippy,
boundary, migration-host, sibling, and sanitizer status remain separate gates
and are not inferred here. Fuzz verification remains tracked with ADR 0008.

Scoped all-target strict Clippy for `litchi-pages` and `litchi-iwa-protos` is
green. The workspace/all-features lint now passes under its unchanged strict
policy after legacy compatibility accesses were confined to explicit host
boundaries and unused helpers were removed. This is workspace lint evidence;
focused Pages/protobuf and native-admission results remain separately scoped.

The checked-in [`body-table-visible.pages`](../../test-data/iwork/pages/body-table-visible.pages)
fixture is a native Pages 14.4 baseline with a visible 5-by-4 body table and a
body marker. Disposable copies were saved, closed, and reopened in the UI
without repair; the `Package` exact no-op check succeeds, and the native visible
profile reads empty and permits an exact empty no-op. A changed hidden-axis
request is refused as `UnsupportedDependency` before publication. The fixture
contains no user-hidden axes, so it is native visible-table/read/no-op evidence
and does not provide positive hidden-axis E2 or E3/E4 mutation evidence. An older exploratory Pages 14.4
generated-candidate attempt logged an NSCocoa MissingObject/TSPersistence Import
document error, timed out on save/close, and never reopened; AppleScript table
creation stalled without GUI inspection. That attempt remains historical
negative exploratory evidence only.

The native hidden-state envelope is ownerful with one state, but both its row
and column state lists are empty. The focused route therefore has no hidden
positions to project from this profile; changed requests remain an
`UnsupportedDependency` result rather than evidence of native mutation support.

A fresh Computer Use duplicate/save/close/reopen check showed the same body
marker and visible table without repair UI. The checked-in fixture was restored
at its recorded SHA-256 after the disposable UI checks, and the focused package
save path produced identical bytes for the visible-table/no-op operation. This
is native baseline evidence for empty hidden-axis parsing and an exact no-op
only; it does not establish parsing of nonempty hidden axes or a changed native
visibility mutation.

The Pages raw-ID `PagesEditor::{table_hidden_axes, set_table_hidden_axes}`
route, its tests, and the mixed example remain migration-host compatibility.
Retirement remains deferred because `NativeVisible` changed edits still return
`UnsupportedDependency`, native changed-edit parity is unproven, and the
migration host still carries compatibility behavior. Qualified `Indexed`
absent-owner creation is now owned by the focused route; it does not authorize
raw-ID retirement or broaden native scope. The shared hidden-axis helper also
preserves Numbers/Keynote compatibility, Numbers sort restoration, and
row/column-deletion cleanup. Focused refusals remain terminal; supported format
facades never fall back to the host. Existing functionality is retained until
ADR 0028's parity and native gates permit removal.

The host cleanup also removed an unused private package identity-regeneration
helper and its dead tests. No focused `regenerate_document_identity` API was
found or retired; UUID generation remains for source-built document
identities. This is dead-code cleanup and does not close a migration or
deletion gate.

The legacy Keynote table listing now uses the bounded Buffa table-model
discovery projection for the name and dimensions it actually consumes. Its
one-valid-candidate rule and historical role-alias rejection remain intact;
the native `table-discovery.key` sample supplies producer evidence only and
does not change package topology or migration-gate status.

No workspace package, manifest dependency, canonical edge, ordered debt, host
item, or deletion gate is removed or closed. The authoritative topology
remains 64 workspace packages, 238 internal dependency declarations, 227
canonical edges, 11 development-only edges, 11 ordered migration debts with
IDs `[1, 2, 4, 8, 10, 12, 13, 14, 15, 16, 17]`, and one migration host.

## 2026-09-05 follow-up: Indexed hidden-state owner creation

The focused Pages owner now has a qualified exact `Indexed` owner-creation
transaction for an absent hidden-state owner. It supports a valid
metadata-free source and a current `Index/Metadata.iwa` source, adds four
helper objects for column/row formula and filter records, and uses the core
canonical object-reference `FieldInfo` primitive for the selected model edges.
Physical object/reference/UUID census, current registry collision and misroute
checks, metadata save-token planning, finite resource budgets, candidate
reopen, exact source/inverse checks, and component/object locality remain the
publication boundary. Empty ownerless reads and exact empty no-ops remain
supported.

`NativeVisible` changed edits still refuse `UnsupportedDependency` before
candidate allocation or publication, and native changed-edit parity remains
open. The raw-ID host route remains compatibility; Indexed creation does not
retire it or broaden the native profile. Admitted Numbers numeric-family and
control-to-numeric routes remain focused-owner work, while the scoped native
Currency `$42.00` save/close/reopen observation does not alter topology or
native-parity gates. The 34-test hidden-axis integration suite, 87-seed corpus
replay, and 100-run AddressSanitizer smoke pass for this follow-up.

No package, dependency edge, ordered debt, migration-host item, or ADR 0028
deletion gate is removed or closed. The authoritative inventory above is
unchanged.


## 2026-09-06 amendment: Pages appearance listing ownership

The earlier `legacy_table_appearance_compatibility_read` exception and its
documentation-only Pages module are deleted together with the shared host
appearance reader. A focused Pages catalog read seam now owns style-edge
projection, preset/network traversal, inheritance, defaults, and bounded
Buffa decoding. The host supplies a payload-free index and borrows selected
payloads from its existing parsed-archive cache. Listing does not serialize
or reopen the whole package, and source-built defaults do not depend on
mutation metadata admission.

This retires a read-side compatibility implementation. Public selector-first
mutation admission, native changed-edit parity, and the remaining monolith
exit gates are unchanged. See the [ADR 0028 follow-up](0028-iwa-monolith-exit.md#2026-09-06-follow-up-focused-catalog-appearance-reads-and-native-datetime).

## 2026-09-09 Pages name-owner native admission follow-up

The Wave107 name owner now supports the retained native and source-built
body-table fixtures. Selected-model authority remains strict; unrelated known
metadata and data-asset references no longer require a physical-object UUID
bijection. Unique current effective locators and selected binding ownership
are verified before publication. A private Buffa-based dependency module
guards changed renames and patch application against active volatile name
cells, pivot owners, and malformed rooted calculation-engine routes. Reads
and exact no-ops retain their separate behavior.

Pages 14.4 displayed a focused rename, saved it, closed, and reopened the exact
file with its title, grid, and body marker preserved. The fixtures and receipts
are under test-data/iwork/pages/body-table-name-*. ADR0028's matching amendment
records the proof and remaining name-only publication scope. Legacy
PagesEditor::tables and physical table/cell responsibilities remain in the
monolith; no package/dependency/debt deletion is claimed.

## 2026-09-09 Pages semantic table catalog follow-up

`litchi-pages` now owns `Package::body_tables()`, a bounded read-only catalog
of rooted body-table positions, names, and declared dimensions. It shares
one discovery walk and cumulative budget, uses the borrowed Buffa model
projection, and exposes no native identities. Default entry selectors use
snapshot positions; duplicate exact names remain explicit lookup ambiguity.
The retained Pages 14.4 two-table fixture and its receipt qualify native
order/name/dimension readback after save, close, and reopen. The corresponding
ADR0028 amendment records the implementation and verification scope.

The legacy stylesheet-closure test uses this focused catalog for semantic
facts. Physical cell/graph witnesses remain host-owned, as do the raw-ID
consumers preventing removal of PagesEditor::tables. This adds the catalog
prerequisite without claiming table-cell support or a retired dependency edge.

## 2026-09-09 Borrowed table-row storage follow-up

The monolith and focused Numbers extractors now share the bounded `CellSpans`
view in numbers_table_cell_storage_codec. Their duplicate allocating offset
parsers are removed; decoded Buffa row snapshots select the active storage
pair and validated iterators borrow cell ranges. Cumulative projection budgets
include validation and traversal work before payload interpretation.

The populated native Pages fixture records save/close/reopen evidence for
nonempty table cells and formulas while preserving the existing host reader.
This shared storage primitive supports the next Pages/Keynote cell migration;
it is not a focused Pages cell API, full table-reader retirement, or a removed
crate dependency. ADR0028's matching amendment records the boundaries.

## 2026-09-09 Legacy cell storage follow-up

`litchi-numbers-wire` now owns the borrowed pre-BNC cell view alongside modern
BNC storage. Both the migration host and focused Numbers table reader consume
that view and have removed their local legacy header/field/scalar decoders.
The view retains caller-owned source bytes and opaque suffixes, with a fixed
21-field validation bound and finite typed scalars. Package topology,
semantic values, sidecar resolution, and operation budgets remain with their
existing owners. No format-peer edge or supported raw-ID API is introduced;
complete Pages table/cell ownership and monolith removal remain outstanding.

## 2026-09-09 Neutral sparse table model

The archive-free cell value, coordinate, and sparse table implementations now
belong to `litchi-iwa-common::table::{cell::value, coordinate, model}`. Numbers
preserves its facade with explicit value/coordinate re-exports and thin
Table/Builder wrappers. The host PagesTable stores the shared core directly.
Format selectors, native graph resolution, sidecars, and transactions retain
their concrete owners. This establishes shared semantic storage for future
focused Pages reads without introducing a Pages-to-Numbers dependency.

## 2026-09-09 Shared native formula event rendering

`litchi-numbers-wire` is extended from binary cell storage to the shared
native formula event-to-text adapter. Its new dependency on
`litchi-iwa-protos` consumes the existing bounded, generated-free formula
events; it does not introduce a dependency on a concrete Numbers, Pages, or
Keynote facade. The common expression arena remains archive/schema-free.

The migration host and focused Numbers supply their own reference resolvers,
typed errors, and operation budgets. Native object discovery, retained-wire
admission, decoding, and aggregate report charging remain concrete adapter
responsibilities. Sharing the event renderer is a prerequisite for focused
Pages/Keynote table reads; it does not by itself retire their host readers.

## 2026-09-10 Shared formula-envelope validation

The native formula envelope schema and bounded preflight move into
`litchi-numbers-wire::formula_envelope`, alongside the shared native event
renderer. The focused Numbers reader and migration host use one validator
for known wire types, canonical scalar encodings, required and duplicate
fields, UTF-8, scalar-render eligibility, and lazy traversal counts.

Concrete readers retain the owned formula bytes, package budget mutation,
format-specific error mapping, and decode-at-use policy. The shared scan
returns costs for successful and failed attempts; this boundary does not
retire the remaining Pages/Keynote host table-reader dependencies.

## 2026-09-10 Focused merged-cell readback

Archive-free merge geometry and its topology algebra now belong to
`litchi-iwa-common::table::merge`; Numbers retains its existing public paths
through compatibility reexports. Pages and Keynote expose the same checked
geometry through their format-owned semantic table APIs.

`litchi-numbers-wire::table_merges` owns bounded native merge-store readback.
The format adapters prove table ownership and supply their remaining read
budgets. The host merge read path also uses the shared decoder, while its
existing mutation code remains until focused write parity is complete.
The broader host table-value/comment reader dependencies remain tracked.

## 2026-09-10 Shared cell-value interpretation

The native value envelope now has one interpreter in
`litchi-numbers-wire::cell_value`. It projects BNC and pre-BNC payloads into
finite scalars or unresolved typed value sources without allocating on
successful reads. Both the focused Numbers extractor and migration host
delegate version dispatch, defaults, formula precedence, and comment
references to this implementation.

Concrete adapters retain sidecar lookup, text retention, formula rendering,
comment materialization, and their existing error categories. The shared
projection is a building block for selected-table readers; format graph
selection and table traversal remain owner-local. The Pages and Keynote host
table readers are still migration debt.

## 2026-09-10 Shared table-data-list coordination

`litchi-numbers-wire::table_data_list` now coordinates strict root and segment
selection, resolution, range validation, key merging, and sorted publication
for both Numbers readers. Borrowed message iterators leave archive ownership
and object lookup in the format adapter. Rejected and duplicate candidates
still receive complete wire validation without retaining semantic values.

The adapters retain their codec visitors, error mapping, and resource budgets.
In particular, the host reserves callback work after its no-callback preflight;
the focused reader preserves its existing single callback pass. An explicit
range-overflow policy preserves the readers' historical error precedence.
This shared coordinator prepares selected-table readback without moving
format graph traversal or concrete package state into the wire crate.

## 2026-09-10 Focused table values and comments

Pages `body_table_cells` and Keynote `slide_table_cells` resolve semantic
table selectors into an archive-free `litchi-iwa-common::table::read::TableRead`.
The result combines the shared sparse table model with positioned comments,
finite timestamps, author display metadata, and resolved replies. Owned
handoff constructors let bounded readers transfer retained comment storage
without rebuilding it.

`litchi-numbers-wire::table_cells` owns borrowed tile and row traversal and
delegates cell classification to `cell_value`. `table_sidecars` owns the five
native list kinds, strict formula envelopes, formula rendering, and comment
storage projection. `formula_names` projects formula-owner identity edges
and category labels without constructing a generated archive or formula AST.
The list coordinator accepts a resolver sharing the decoder's existing
ledger, so nested segment reads do not create independent budgets.

Concrete packages retain table ownership proofs, object lookup, reference
resolution, cumulative limits, and public format-specific errors. These
focused APIs add no dependency on the migration host. The host full-table
callers remain until their complete compatibility and mutation surfaces can
be retired; the monolithic crate is not yet removed.

## 2026-09-10 Host table-reader retirement

The host Numbers `table_extractor`, `formula_renderer`, and `table` modules
are deleted. Pages and Keynote host table-read entrypoints resolve their
legacy model selectors through the existing format catalog, then delegate
to focused semantic cell and merge readers with effective package limits.
Their result storage is the neutral common `TableRead`; no replacement host
extractor or raw-identity bridge is introduced.

Read-result comments now use resolved text, timestamps, authors, and replies
instead of the obsolete host raw-ID comment representation. Existing comment
mutation APIs remain separate migration debt. The Pages adapter currently copies
package source bytes; both adapters perform separate cell and merge scans.
A shared source and combined read API can remove these temporary costs. Format mutation
and other host modules remain, so the monolith deletion gate is still open.

## 2026-09-10 Retire unused host reference traversal

The remaining host object index is a physical location catalog for document
statistics and chart metadata lookup. Neither caller traverses reference
edges. Its MessageInfo edge ingestion, payload-reference fallback, and private
graph query adapters are removed. Physical identity, span, source-position,
duplicate-object, allocation, and deterministic-order checks remain.

Generic graph behavior remains owned by `litchi-iwa-index` and
`litchi-iwa-graph`. Format-specific reference proofs remain in concrete package
owners and their bounded wire codecs; the host no longer scans unrelated
payloads merely to construct a graph with no consumer. Document text, chart
metadata decoding, and the remaining mutation paths still require migration.

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

### 2026-08-08 Keynote ordering continuation

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
vertical capability. The current boundary ledger therefore still contains all
17 ordered `litchi-iwa` debts. Remaining editor, Prost graph, example, test,
fuzz, durable-patch, and atomic-save ownership prevents host deletion.

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
identities. The current checked topology is 64 workspace packages, 227
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

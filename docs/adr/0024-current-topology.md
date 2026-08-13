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

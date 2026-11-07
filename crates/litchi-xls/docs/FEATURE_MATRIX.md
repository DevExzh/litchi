# Excel BIFF8 (.xls) Feature Matrix

This document tracks the public feature families implemented by litchi-xls. It describes
typed library support, not Excel's rendering, recalculation, trust, or complete conformance
with every BIFF revision.

The audit uses the authoritative Microsoft [MS-XLS] Front Matter and ToC, especially the
compound-file and stream grammar in [MS-XLS] 2.1, the workbook and worksheet substream
grammars in [MS-XLS] 2.1.7.20, the cell/formula/chart/pivot/style model in [MS-XLS] 2.2,
and the examples for conditional formatting, defined names, tables, filters, external
references, charts, formatting, workbooks, and PivotTables in [MS-XLS] 3.1-3.10.

## Status model

| Mark | Meaning |
|------|---------|
| ✅ | Supported for the scope described in the Notes cell |
| 🟡 | Bounded, partial, metadata-only, pass-through, or inert support |
| ❌ | No public typed support currently available |
| N/A | The concept does not apply to this format or direction |

Read and Write are independent. A supported row does not imply recalculation,
rendering, external I/O, macro execution, or certificate trust.

## Package, workbook, and core BIFF

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| OLE Compound File container | ✅ | ✅ | ✅ | Opens and writes the CFB storages and streams used by .xls; package edits are bounded and transactional at the owning writer boundary. |
| BIFF8 workbook and worksheet record streams | ✅ | ✅ | ✅ | Typed workbook, worksheet, chart-sheet, macro-sheet, and dialog-sheet record sequences are decoded and emitted through the XLS reader/writer. The complete reader propagates worksheet failures instead of silently publishing an incomplete catalog. Native Microsoft Office XLS/XLT worksheet catalogs are checked by `microsoft_office_native.rs`. |
| Generic BIFF record access | ✅ | ✅ | ✅ | Public record utilities expose record headers and payloads. Native frame/offset/continuation round-trips and resource boundaries are covered. `BofRecord::substream_type` reads `BOF.dt`; it replaces the erroneous public `is_1904_date_system` field, which read the build identifier. Use `raw::Reader::date_system()` or the workbook trait for date mode from `Date1904`; the typed model does not claim a semantic implementation of every record in the [MS-XLS] enumeration. |
| Unknown and future BIFF records | 🟡 | ✅ | 🟡 | Selected Continue, FRT, extension, and opaque payload families are retained or replayed. Native metadata scalar editing preserves every unowned frame, including BookExt/XFExt/StyleExt/Theme records; unknown records are not guaranteed to survive every typed mutation or fresh authoring path. |
| Workbook and sheet directory | ✅ | ✅ | ✅ | Sheet names, order, visibility, stable metadata, sheet kinds, code names, and sheet lookup are modeled. |
| Cell values and cached results | ✅ | ✅ | ✅ | Numbers, strings, shared-string references, booleans, errors, blanks, serial dates, styles, and cached formula results are supported. Opened workbooks expose `cell_values::{Snapshot, Edit, Commit, Patch, SemanticPatch}` for fixed-width `Number`, standalone/packed `RK`, `BoolErr`, `Blank`, `LabelSst`, and Formula-cache fields, plus checked scalar-cell insertion/removal, bounded plain/rich continued SST interning, ordinary Formula authoring, and validated 20-byte XF authoring. Edits are bounded, failure-atomic, deterministic-join/conflict aware, semantically durable, reversible, three-way/transfer preflighted, stale-checked, and fully reopen the CFB/XLS candidate. Commit diagnostics expose `changed()`, while artifact and semantic patches expose `is_noop()`. Structural publication regenerates the affected row-block/`INDEX`/`DBCELL`/`DIMENSIONS` closure and shifts every later BoundSheet position. String-valued Formula caches own and rewrite their following `String` record; formula-cache transfer fingerprints the unchanged Formula token/metadata dependency. Existing `ExtSST` accelerators are updated exactly while their canonical bucket size remains stable, including reversible bucket-tail growth. |
| Shared strings and rich-string properties | ✅ | ✅ | ✅ | BIFF SST/ExtSST lookup plus bounded string-property metadata is available. The opened-workbook edit can append plain strings and absent-text rich strings with validated formatting runs across correctly framed `Continue` records, updating an adjacent canonical `ExtSST` unless the operation crosses a bucket-size boundary. An existing rich entry is reused only when its text occurs once and its non-phonetic formatting identity matches exactly; phonetic-property synthesis and a second differently formatted SST entry with the same text remain unsupported. |
| Formula tokens and cached formula values | 🟡 | ✅ | 🟡 | BIFF formula token rendering, cached values, and typed `FormulaMetadata` (`fAlwaysCalc`, `fFill`, `fShrFmla`, `fClearErrors`, and `chn`) are available. Existing Formula caches can be changed among checked numeric, Boolean, error, empty, and bounded string encodings without changing tokens or calculation metadata; durable replay and transfer require an exact fingerprint of that unchanged Formula dependency. Opened-workbook edits can also insert an absent ordinary Formula from the writer's bounded tokenizer, with an empty cache and `fAlwaysCalc`, and support deterministic join, durable replay, reopen, and exact inverse. They do not synthesize shared, array, or table-formula ownership. Canonical workbook opening and the low-level codec ignore an `fShrFmla` flag on a Formula that carries its own tokens (not part of a shared formula, so no `PtgExp` is required; Apache POI documents Excel writing these) and record `WorksheetDefect::SharedFormulaFlag` (record 0316); the writer still refuses `fShrFmla` without `PtgExp`. Callers may explicitly select the versioned `SharedFormulaFlagWithoutPtgExpV1` workbook profile through `OpenOptions::new().compatibility_profile(...)` for the real `ConditionalFormattingSamples.xls` fixture mirrored in the Apache POI corpus; the original producer is unknown, and every accepted record is reported with its checked BIFF8 cell and selected profile. `OpenOptions` is opaque and builder-based so future options remain source-compatible. `Builder::write_shared_formula` authorizes checked BIFF8 ranges, anchors, participants, `PtgExp`, `RefU`, `cUse`, and the Formula→ShrFmla sequence. Tokens remain inert and are not a complete Excel parser or evaluator; cached values are not recalculated. |
| Array, shared, and data-table formulas | 🟡 | ✅ | 🟡 | Strict typed BIFF8 `Array` owners parse and validate the complete rectangle and its `PtgExp` Formula binding, preserve bounded `Array`/`RgbExtra` payloads, and write the required Formula→Array layout. Conservative safe textual authoring is bounded to the supported formula subset. Existing-workbook array resize/add edits and a complete Excel formula compiler are not provided; cached values remain inert and are never evaluated or executed. Shared-formula authoring remains bounded to the typed `ShrFmla` owner path, while table metadata and Formula/Table records remain supported where implemented. |
| Defined names and built-in names | ✅ | ✅ | ✅ | Workbook and sheet-scoped names, print areas, print titles, filter-database names, and name metadata are typed and writable. Name formulas remain inert token expressions. Native coverage checks workbook-scoped FixtureAmounts and sheet-scoped print-area/title names, including exact token bytes and rendered references. |
| Workbook calculation properties | ✅ | ✅ | ✅ | Calculation mode, reference mode, iteration settings, precision, multithreading flags, recalculation markers, and force-full-calculation metadata are supported. |
| Workbook views and window state | ✅ | ✅ | ✅ | Workbook windows, selected/active tabs, stable tab identifiers, sheet views, panes, selections, zoom, and related view metadata are supported. |
| Core, summary, and custom properties | ✅ | ✅ | ✅ | CFB property sets are exposed for summary information, document summary information, and user-defined properties. Native MacExcel metadata CRUD preserves Workbook/VBA streams, standard property values, reserved hyperlink blobs and the root CLSID; Unicode custom properties upgrade their own section independently. |
| Theme, fonts, palette, number formats, and XF styles | ✅ | ✅ | ✅ | BIFF fonts, palette colors, number formats, cell/style XFs, XF extensions, differential formats, table styles, and theme metadata are typed. Opened-workbook edits can append exact validated 20-byte XF bodies when their font, format, protection, kind, and parent-style dependency prefix matches an effective XF. Appending/removing XFs updates an existing XFCRC count and MS-OSHARED checksum while preserving XFExt records; a stale source checksum refuses that edit, and the 4050-XF limit is enforced. Fresh XFCRC emission also computes the actual payload checksum. `set_cell_format` retains its palette-based authoring contract: the new XF clears `fHasXFExt` instead of cloning the source theme overrides. |
| Workbook protection | ✅ | ✅ | ✅ | Workbook protection flags and legacy password records are modeled; the legacy hash is not cryptographic protection. Changed worksheet renames share the structure-protection refusal used by tab insertion/removal; exact same-name renames remain no-ops. Native passwordless structure protection is covered, and worksheet-only protection leaves tab names editable. |
| BIFF password encryption | ✅ | ✅ | ✅ | Supported BIFF8 password-to-open profiles are decrypted and emitted through the XLS encryption integration; unsupported profiles fail rather than being treated as plaintext. RC4 writing uses `Builder::set_password`; legacy XOR is decode-only by default and writing it requires `WeakEncryptionPolicy::allow_xor_obfuscation()` plus `Builder::set_xor_obfuscation_password`. |
| Legacy digital signatures | ✅ | ✅ | ✅ | CFB signature reports can be verified and signature graphs can be edited through the signature integration. Verification is integrity/signature verification only and does not establish certificate trust or revocation. |
| VBA project and module storage | 🟡 | 🟡 | ✅ | VBA project topology, metadata, and bounded module/source payloads can be inspected or attached/replaced. VBA is never executed, resolved, or trusted; replacing project content invalidates stale signatures. |
| Custom XML data store | 🟡 | 🟡 | 🟡 | The workbook exposes bounded custom XML storage through the CFB/package layer, but it is preserved as payload-oriented data rather than a general XML-mapping authoring model. |
| Legacy macrosheet and dialog-sheet behavior | 🟡 | ✅ | 🟡 | Sheet kinds and selected records are recognized as metadata. Macro formulas, dialog controls, and host automation are not evaluated or executed. |

## Worksheet content and interaction metadata

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Cell and worksheet CRUD | ✅ | ✅ | ✅ | New worksheets can be authored with typed cells, dimensions, shared strings, and sheet metadata. The opened-workbook edit can insert absent standalone number/Boolean/error/blank/plain-text/rich-text/ordinary-Formula cells, physically remove standalone non-formula cells and any member of `MulRk` (splitting and exactly re-merging an interior deletion), assign, duplicate, or author validated XF resources, rename sheets, and insert dependency-safe row/column coordinates. It regenerates row blocks, `INDEX`/`DBCELL`, row extents, dimensions, and BoundSheet offsets together. Reversible shifts move ordinary Formula owners whose bounded tokens contain constants and same-sheet `PtgRef`/`PtgArea`, preserving relative/absolute flags, and also transform selections, merged ranges, external-target hyperlinks/tooltips, defined names and 3-D references naming the sheet, shared (`ShrFmla`), array (`Array`) and data-table (`Table`) owners that move whole, conditional formats (`CondFmt`/`CF`, `CondFmt12`/`CF12`, `CFEx`) and data validations whose relative references stay on one side of the shift point (records 0308, 0316), and OfficeArt anchors: `AnchorBehavior::MoveAndResize` shifts like cells, `MoveWithoutResize` moves whole, and an anchor that keeps its place (`Absolute`, or the far end of a size-locked one) keeps its cell coordinates when every row or column it spans from the shift point keeps its size. Endpoint-erasing deletions, 3-D references spanning sheets, a shift point inside an owner range, internal hyperlink locations, AutoFilter, Sort and Pane dependencies, formulas with `RgbExtra`, non-comment `Obj` records, destructive deletion with content, phonetic SST synthesis, and duplicate rich-text identities fail closed. |
| Row and column dimensions | ✅ | ✅ | ✅ | Row heights, hidden/collapsed state, outline levels, column widths, best-fit flags, formatting, and phonetic-guide flags are supported. |
| Cell formatting lookup | ✅ | ✅ | ✅ | Effective cell formats, number formats, borders, fills, alignment, fonts, colors, and date/time interpretation are available. |
| Merged cells | ✅ | ✅ | ✅ | MERGECELLS ranges are read and emitted with checked range boundaries. Direct `Worksheet::add_merged_cells` is create-only and refuses source-bound worksheets; native package edits preserve merges and bounded coordinate shifts can move them, but there is no general opened-workbook merge/unmerge API. |
| Hyperlinks | ✅ | ✅ | ✅ | URL, file, item, relative/absolute, display, tooltip, frame, and location metadata are typed; following a link is never performed. |
| Classic comments and notes | ✅ | ✅ | ✅ | BIFF NOTE comments, authors, visibility, anchors, and writer support are available. |
| Worksheet protection | ✅ | ✅ | ✅ | PROTECT, object/scenario protection, password records, and sheet protection flags are modeled; protected content is not a cryptographic boundary. |
| Data validation | ✅ | ✅ | ✅ | BIFF data-validation settings, ranges, operators, formulas, prompts, and error messages are typed and writable within the supported record forms. Native tests cover seven rule kinds. Best-effort formula text resolves PtgRefN/PtgAreaN offsets at the top-left of the first stored target range while ordinary reference coordinates remain literal; rules are never evaluated. |
| Conditional formatting | ✅ | ✅ | ✅ | Classic conditional-format rules and the implemented BIFF12-style extension metadata are available; rule formulas are stored/rendered, not evaluated by this crate. |
| AutoFilter and sort state | ✅ | ✅ | ✅ | AutoFilter columns, criteria, filter mode, sort metadata, and extended filter records are typed and writable. |
| Structured tables and List12 objects | ✅ | ✅ | ✅ | List objects, ranges, columns, totals metadata, styles, AutoFilter state, and selected external/web/XML source metadata are serialized. Source queries and calculated-column formulas remain inert. |
| Scenarios and what-if data tables | 🟡 | ✅ | 🟡 | Scenario manager and TABLE records have typed metadata and bounded writing; Excel's what-if calculation and scenario application are not performed. |
| Query tables and external-data metadata | 🟡 | ✅ | 🟡 | Query-table, connection, web-table, and external-table records are typed where implemented. Commands, credentials, and refresh behavior are never used. |
| Worksheet views and panes | ✅ | ✅ | ✅ | Window, zoom, pane/freeze, selection, and related worksheet view records are supported. |
| Page setup and print layout | ✅ | ✅ | ✅ | Margins, paper, scale/fit, orientation, print order, headers/footers, page breaks, gridlines, comments/errors display, and print flags are typed. Printer-driver bytes are inert. |
| Background and header/footer pictures | 🟡 | ✅ | 🟡 | Bounded record codecs retain BIFF background and header/footer picture payloads. Native tests join 21 background continuation records, check the bitmap header and embedded PNG boundary, and round-trip both opaque HFPicture payloads. Unrelated property publication preserves the complete Workbook stream; typed picture placement edits and rendering are not established. |
| Cell watches and formula-error features | 🟡 | ✅ | 🟡 | CellWatch has a detached record codec; formula-error features have a typed worksheet getter and detached payload writer. Native tests pin the A7 watch, A4:A5 calculation-error range and stored error/cache values. Unrelated property publication preserves their Workbook stream; this does not establish package CRUD, watch-window monitoring or diagnostic execution. |
| Phonetic information | 🟡 | ✅ | 🟡 | Phonetic string format and visible-range metadata are typed; phonetic layout and language conversion are not implemented. |
| Data consolidation | 🟡 | ✅ | 🟡 | Consolidation directories and source metadata are available as inert typed records; consolidation is not executed. |
| Web publishing | 🟡 | ✅ | 🟡 | Web publication destinations, ranges, source types, and page metadata are modeled; publishing and network I/O are not performed. All 17 native XLS/XLT fixtures lack WebPub records; ordinary URLs are not equivalent evidence. |
| Real-time data and RTD topics | 🟡 | ✅ | 🟡 | Contextual `real_time_data::{Snapshot, Transaction, Commit, Patch}` edits bounded BIFF8 RTD topics, shared-prefix framing, cells, and cached values; all ProgIDs, servers, and topics remain inert and no refresh or execution occurs. |
| Revision log and shared-workbook collaboration | 🟡 | ✅ | 🟡 | `revision_records::{Snapshot, Edit, Commit, Patch}` provides source-checked BIFF8 Revision Log edits for accepted/undo flags and selected header metadata while preserving unknown records; conflict resolution, shared locking, formula replay, and collaboration behavior are not implemented. |
| User Names stream and shared-workbook user log | 🟡 | ✅ | 🟡 | `user_names::{Snapshot, Edit, Commit, Patch}` and `user_names::publication` provide bounded stream and package metadata edits, exact no-ops, reversible source-checked patches, and `UsrInfo.guid` validation against the required Revision Log. Publication preserves unrelated CFB streams and directory metadata. Shared locking, revision replay, collaboration, and native-producer interoperability are not established; positive package CRUD coverage uses synthetic BIFF/CFB fixtures; the native legacy specimen has a noncanonical `UsrChk` version and invalid historical tab maps, documented below. |
| Custom views and workbook environment options | 🟡 | ✅ | 🟡 | Custom views, link-update modes, object-display modes, access/write strings, backup flags, and related behavior options are exposed as metadata. Native custom views interpret `fHiddenRw=1` as hidden rows present and `fDspZerosSv=1` as zero values displayed, matching native XLS/XLSX evidence and the Excel97–2007 BIFF definitions. Newer MS-XLS prose reverses these labels and contradicts its own hidden-row-name requirement. These snapshots do not apply a view or cause host actions. |
| Office Toolbars stream (`XCB`) | 🟡 | ✅ | ✅ | `raw::Reader::toolbar` reads and `Builder::set_toolbar` emits typed `CTBWRAPPER`/`CTBS`/`CTB` metadata through the root `XCB` stream, preserving reserved fields, optional visual bytes, `TBCCmd`, and bounded shared `TBCGeneralInfo`/`TBCExtraInfo` payloads for non-ActiveX controls. Ambiguous or malformed variable boundaries are rejected, and macros/UI/ActiveX are never executed. All 17 native XLS/XLT fixtures lack the XCB stream; its native qualification remains open. |

## Pivots, charts, drawings, and external references

| Feature | Status | Read | Write | Notes |
|---------|--------|------|-------|-------|
| Pivot caches | ✅ | ✅ | ✅ | Typed PivotCache definitions include shared items, value types, grouping, worksheet/consolidation sources, OLAP hierarchies/tuples, calculated items/members, and supported extension records. |
| PivotTable views | ✅ | ✅ | ✅ | Pivot fields/items, axes, page selections, data items, layouts, functions, view extensions, and the constrained pivot-view editor are available. `SXVI` item types use the BIFF codes 0–14, distinct from aggregation-function codes; native data items and automatic subtotals retain their correct meaning. Unknown item codes remain opaque. The crate does not refresh or calculate a PivotTable. |
| Pivot charts | 🟡 | ✅ | 🟡 | Pivot-chart links and related SXView/PivotChart records are recognized; chart rendering and complete pivot-chart authoring remain bounded by the chart graph support. |
| BIFF charts and chart sheets | 🟡 | ✅ | 🟡 | litchi-xls::chart provides chart discovery, strict BOF/EOF framing, a typed fixed-point chart-area snapshot/transaction, bounded series formula/cache edits, byte-preserved unknown records, and exact-offset replay. The chart-area edit never resizes an embedded host object. Full chart grammar certification, fresh chart authoring, rendering, and unsafe graph edits are refused. |
| OfficeArt drawings and shape groups | 🟡 | ✅ | 🟡 | Shape extraction exposes typed shape/group geometry, XLS `SheetAnchor` cell-relative endpoints, anchors, text, and selected drawing metadata. The complete OfficeArt drawing-group graph, renderer, and arbitrary shape mutation are not claimed. |
| OLE objects and form controls | 🟡 | ✅ | 🟡 | Typed OLE object/form-control records and bounded inert payloads are exposed through `ole_object::Snapshot`, failure-atomic `Transaction`, source-checked reversible `Patch`, and `ObjectMetadataEdit` facades. Safe `FtCmo`/`FtCf`/`FtPioGrbit` metadata edits preserve unknown BIFF subrecords and reject unsafe DDE storage transitions; payloads are never activated, instantiated, rendered, or executed. Workbook CFB `encryption`, BIFF `FILEPASS`/active protection, and stale PIDDSI `DigitalSignature` inputs fail closed before publication. |
| Embedded packages and external payloads | 🟡 | ✅ | 🟡 | `ole_object::{EmbeddedPayload, EmbeddedObjectDraft}` and the source-checked transaction facade support bounded inert creation, complete-storage replacement, and identity-checked removal for storage-backed `MBDxxxxxxxx` picture objects (`cmo.ot=8`, `fDde=0`, `fPrstm=0`). Supplied payload CFB streams are retained opaquely; unknown workbook/object streams survive, selected replacement owns the complete selected storage, and no OLE server, macro, control, link, or native payload is activated. DDE/link, control-stream, arbitrary Obj-subrecord graph, and payload-internal semantic authoring remain unsupported. Protected/signed/encrypted/DRM CFB inputs are refused by the shared mutation boundary. |
| Supporting books and external links | 🟡 | ✅ | 🟡 | Supporting-book records, external names, source paths, external references, link flags, and cached values are typed through `external_link::edit::{Snapshot, Transaction, Commit, Patch}` where implemented. Paths and formulas are stored verbatim and never opened, fetched, refreshed, or evaluated; unknown BIFF records and `Continue` payloads remain intact. |
| DDE and OLE link behavior | 🟡 | ✅ | 🟡 | DDE/OLE link metadata is retained as inert records. No conversation, COM activation, data refresh, or external program launch occurs. |
| MDX and cell/value metadata | 🟡 | ✅ | 🟡 | BIFF metadata, MDX records, and selected metadata blocks are exposed as bounded metadata. Cube connections and MDX expressions are not executed. |

## Explicit gaps

| Feature family exposed by [MS-XLS] | Status | Read | Write | Notes |
|-------------------------------------|--------|------|-------|-------|
| Threaded comments, persons, mentions, and modern comment threads | ❌ | ❌ | ❌ | The .xls model supports classic NOTE comments only; modern threaded-comment parts are not a BIFF8 typed feature. |
| Slicers and timelines | ❌ | ❌ | ❌ | No typed slicer/timeline cache or view model is provided for the BIFF8 crate. |
| Rich data types and dynamic-array spill semantics | ❌ | ❌ | ❌ | The [MS-XLS] cell/formula records do not become a modern rich-value or spill-calculation engine in this crate. |
| XML maps and mapped XML import/export | 🟡 | ✅ | 🟡 | `[MS-XLS]` XML-map metadata has a typed `MapInfo`/schema/map/data-binding owner, bounded compact XML codec, workbook `XML`-stream loading and authoring, and atomic list-column dependency checks; schema resolution, binding refresh, and mapped-cell import/export remain inert. |
| Complete Excel formula evaluation | ❌ | ❌ | ❌ | Formula tokens and cached results can be read, but there is no Excel-compatible recalculation engine, dependency graph, external-function execution, or volatile refresh. |
| Chart rendering and complete chart authoring | ❌ | ❌ | ❌ | The bounded chart model is intentionally not a renderer or a complete fresh chart writer; unsafe graph edits fail rather than producing a corrupt workbook. |
| Macro, control, DDE, RTD, database, and external-link execution | ❌ | ❌ | ❌ | These protocol families are represented only as inert metadata or cached payloads. |
| Certificate trust, revocation, and signed-content policy | ❌ | ❌ | ❌ | Signature verification establishes cryptographic integrity/signature validity only; it does not establish trust. |

The native encrypted Excel regression in `tests/microsoft_office_native.rs`
checks absent/wrong/correct passwords, the CryptoAPI RC4 128-bit source profile,
Sheet1 markers and cached values 42/84, ordinary-publication refusal, and exact
ciphertext retention through `to_encrypted_bytes`. It does not rekey, refresh
formulas or reopen the result in Microsoft Excel.

## Native fixture qualification

`tests/microsoft_office_native_metadata.rs` checks native PivotCache rows,
PivotTable item types/axes/aggregation metadata, CSE/shared/data-table owners,
dynamic-array cache retention, Scenario Manager metadata, classic notes,
hyperlinks, conditional formats, print setup, custom views and 1904 layout.
Legacy XLS LET conversion is tested as its stored `#VALUE!` formula tokens and
prior numeric cache; it is not presented as a reconstructed LET expression.

Scalar create/update/remove covers six native XLS/XLT variants, including macro
and embedded-object containers, with non-Workbook streams and formula-owner
records preserved. Pivot-caption, row/column visibility and page-break edits
check typed readback and exact inverse publication. The protected notes fixture
checks exact no-op behavior and atomic Protected refusal. Public operations
queries and two-sheet edits run on both Auto and Eager owners.

The public facade adapts the leaf BIFF coordinates to the generic worksheet
trait: cells and dimensions are one-based, while row iteration stays zero-based.
Native XLS/XLSX/XLSB tests verify Auto/Eager agreement and grid refusals.

The assertion index is `tools/native_office_tests.json`. These are bounded
feature assertions, not full feature qualification. Opaque formula tokens and
inert macro/OLE streams remain distinct from typed expression or payload
semantics. No fresh Excel reopen, rendering, calculation, refresh, hyperlink
navigation or macro/control execution was performed.

Native `DBCELL` indexes may append zero deltas for empty rows after the offsets
of populated rows. The reader accepts that padding only when the total slot
count equals the observed `ROW` count, retains the original payload, and still
checks every populated-row pointer against the actual cell record boundary.
Padding does not create extra cell pointers. Layout edits may regenerate a
compact `INDEX`/`DBCELL` pair; tests validate the rebuilt index and compare all
records outside that related index and the requested row/column change.

Native form-control tests additionally inventory all eleven BIFF8 controls,
including the button and the list/dropdown records whose terminal `FtLbsData`
uses `cbFContinued` rather than an ordinary exact subrecord length. Native
framing, range-formula tokens, selections, radio groups and control states
round-trip exactly. Appending checkbox OBJ metadata preserves all eleven
original records and inverts after save/reopen; it does not author OfficeArt
widget placement. Continued fragments retain their header and partial state;
full continuation-chain aggregation remains outside this qualification.
Boolean/error records and formula caches expose the standard Excel error
spellings, including the mixed checkbox's stored `#N/A` cell.

Native drawing assertions additionally distinguish picture id 1025 from
rectangle id 1026, read `CF_SHAPE_005`, and verify A12:B17/F16:H20 cell anchors
with their stored move/resize behavior. This is a shape projection; payload
resolution and full OfficeArt geometry or authoring are separate scopes.

The native external-workbook regression reads the encoded supporting-book path,
three sheet names and the exact D6/B4 cached values (49.75 and 14) while retaining
the host Sheet1 catalog. Public `Workbook::external_links` provides this read
view. The internal external-link transaction module is not exposed as a public
Package editor; this regression does not claim target/cache mutation support.


## Native shared-workbook revision qualification

`tests/microsoft_office_native_revisions.rs` reads the native legacy revision
fixture at the record layer. `RRDInfo` history flags/interval follow the complete
four-byte version at offsets 46/48. Earlier tab maps belong to earlier saved
sets; current BoundSheet8 count checks apply to the latest map. Fixed-size old
cell values retain wire-length checks, while variable string/formula size
descriptors are not physical offsets into their encoded/continued tails.
Tests pin five headers, six cell/sheet changes and exact source value bytes.

The XLS compatibility export is not a conforming full revision/user-log example.
Two historical RRTabId arrays contain duplicate zero IDs, and its closed User
Names stream stores UsrChk version `0006` instead of a documented version such
as `0600`. Strict errors and byte-exact preservation are asserted; body values
B2=25 and cached B3=75 still read and save unchanged. The User Names row remains
a reviewed fixture gap, not positive user-log coverage. No bytes are corrected
to manufacture a passing artifact, and no history is replayed or merged.

The complete-Excel-evaluation row is an explicit compatibility boundary, separate
from the optional bounded `litchi-ops` evaluator. Its native calculation suite
checks scalar arithmetic/aggregates, shared `ROW` formulas, cached error kinds,
1904 dates, cache preference, file reopen and read-only byte preservation.
Explicit evaluation refuses external references and syntax/context outside the
profile; an unregistered table reference cannot silently aggregate to zero.
Date functions use the workbook date system, including holiday/offset inputs.
These checks do not establish a complete Excel dependency/spill engine, external
refresh, cache publication or fresh Excel acceptance. See
`tools/native_office_evidence_reviews.json` and
`crates/litchi-ops/tests/microsoft_office_native_calc.rs`.

The runtime boundaries `xls-194`
are explicitly reviewed in `tools/native_office_evidence_reviews.json`.
The operations test `microsoft_office_native_runtime_boundaries` checks all 139
Office files through Auto/file and Eager/byte ingress: HTML, SVG, screenshot and
PDF requests return the advertised refusal through string, JSON and sink APIs,
without writing to the sink. Locked files first require a password; supported
unlocks still refuse visual output, and encrypted XLSB retains its separate
unsupported-decryption boundary. Reads preserve publication bytes and all input
files. Cited owner tests qualify saved data and bounded edits, not the excluded
application rendering, layout or execution.

`microsoft_office_native_signature_boundary` authors disposable XML signatures
over seven native Office parents. These are Litchi derivatives, not native
Office signatures. Mathematical validity, resource coverage and signature
presence remain distinct from `Trust::NotChecked`. Workbook cases also check
outer/inner signed-handle edit refusals, explicit unsign, complete graph/data
equivalence to an ordinary scalar edit, and invalid-payload detection. The
shared verifier's older Microsoft-signed certificate fixtures explicitly retain
unchecked trust too. These tests do not claim native application acceptance,
legacy CryptoAPI signature authoring, trust-chain validation or revocation.
Native signed-fixture handoffs remain open despite this bounded operation
coverage; historical authoring statuses are unchanged.

Native calculation and refresh tests retain physical tab positions across a
leading chart tab, include the trailing worksheet, and verify cross-sheet
results. `File::refresh` reports stale caches by default and writes only
supported results when `write_cache(true)` is explicit. Workbook refresh keeps
external dependencies unsupported and unchanged; it does not refresh data
providers. XLS's inline FormulaValue 03 is a saved blank string, distinct from
a missing following STRING record, and no longer triggers a spurious write.

The native shared-trait regressions in `microsoft_office_native_coordinates.rs`
qualify one-based `Worksheet::cell`, `cell_value`, inclusive dimensions and
`Cell` trait positions, including absent cells and format-grid refusals.
The leaf adapters now satisfy the same contract as the public facade; the
facade no longer applies a second coordinate conversion. Direct generic
formula-evaluator reads agree with native B2 values and Auto/Eager facade grids.
Native `get_cell`, cell constructors, concrete `Cell::row`/`column`, wire/editor
coordinates and positional source reads retain their zero-based contracts.
Callers that used zero-based arguments for shared-trait reads must add one or
use the explicit native stored-cell API. Row iteration remains zero-based.

The explicit scope reviews for xls-189, xls-190 and xls-191 distinguish native
classic notes, pivot caches and converted array/formula storage from modern threads,
slicer/timeline models and rich-value/spill engines. Public native XLS threaded-
comment and slicer paths retain their precise declared errors; supported bounded
calculation is not excluded by these reviews.

Native workbook and macro-template parents now qualify shared Custom XML store CRUD
and public CFB publication, followed by ordinary native cell edits. XML/Properties
streams, schema URI strings and opaque storage metadata remain inert and preserved;
complete CFB inverse replay restores input bytes. xls-152 gains bounded payload-
operation coverage, not a workbook XML-mapping model or a native-authored store
specimen.

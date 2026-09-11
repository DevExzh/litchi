# Specification Gap Audit

Date: 2026-09-08

Historical working audit: the findings below reflect the inspection described
on that date, not current support certification. Subsequent changes must be
checked against each crate's feature matrix and executable evidence. In
particular, the September 11 ODG batch adds bounded transition and drawing
metadata coverage; its current scope is recorded in
[`litchi-odg`'s feature matrix](../../crates/litchi-odg/docs/FEATURE_MATRIX.md).
Source searches alone do not establish specification completeness or native
application compatibility.

This report audits the current branch against every specification checked in under
`3rdparty/specs/`: [MS-DOC], [MS-DOCX], [MS-XLS], [MS-XLSB], [MS-XLSX], [MS-PPT],
[MS-PPTX], [MS-OGRAPH], [MS-OWEXML], [MS-OXRTFCP], RTF 1.9.1, OpenDocument v1.4 OS,
ECMA-376, and the shared specifications ([MS-CFB], [MS-OFFCRYPTO], [MS-OLEPS],
[MS-OLEDS], [MS-OVBA], [MS-VBAL], [MS-ODRAW], [MS-ODRAWXML], [MS-OSHARED], [MS-DTYP],
[MS-LCID], [MS-OAUT], [MS-UCODEREF], [MS-OE376], [MS-OI29500], [MS-WMF], [MS-EMF],
[MS-EMFPLUS]).

The audit granularity is "feature", not rendering fidelity. Each format family was
checked by reading the crate's `docs/FEATURE_MATRIX.md` (or the crate README and public
API where no matrix exists), cross-referencing the specification table of contents, and
spot-verifying claims with source-level greps for missing or refusing code paths.

## Overall assessment

The per-crate feature matrices are honest and largely complete as self-disclosures. The
vast majority of ❌ rows are **deliberate scope boundaries** (no macro/external-link
execution, no rendering, no formula evaluation, no certificate trust chains). The
genuinely actionable findings fall into three categories:

1. **Substantive feature gaps** — the spec defines the feature, no implementation
   exists, and it is not a stated exclusion.
2. **Structural blind spots** — spec structures that no code parses at all.
3. **Matrix documentation gaps** — features implemented in source but missing rows in
   the matrix, or spec areas with no declared disposition.

---

## 1. litchi-doc (Word binary .doc) — [MS-DOC], [MS-OSHARED], [MS-OLEPS]

Structural coverage is very high; genuine spec-level gaps are few. Most ❌ rows are
deliberate (rendering, evaluation, execution, external resolution).

### Declared ❌ rows (deliberate refusals)

- **Field evaluation and generated content** (read ❌/write ❌): TOC/TOA/INDEX/SEQ/
  STYLEREF/FORMULA/PAGE/NUMPAGES fields are never calculated; only instruction text and
  cached results are preserved (spec §2.9.88 Fld, §2.4 field semantics).
- **External resolution** (read ❌/write ❌): RD/INCLUDETEXT/LINK/DDE, mail-merge data
  sources (SQL/connection strings), and subdocument reference files are never opened or
  refreshed (spec §2.1 ObjectPool/PlcfWKB, §2.9 Pms/ODSO).
- **Macro/control execution** (read ❌/write ❌): VBA source, FFData macro names, OLE
  objects, and ActiveX controls are passive data (spec §2.1.9 Macros, [MS-OVBA]).
- **IRM license evaluation** (read ❌/write ❌): no decryption, no rights granting
  (spec §2.1.12–2.1.13 IRM DataSpace / Protected Content Stream, [MS-OFFCRYPTO]).
- **Rendering and pagination** (read ❌/write ❌): no layout, no pagination, no
  OfficeArt rasterization.
- **Word 6/95 and earlier binary profiles** (read ❌/write ❌): the reader refuses
  nFib < Word 97 files (spec §2.5.14; marked intentional in the matrix).
- **Implicit write-side boundary**: `body_text` transactions refuse structural table
  edits, field nesting/separator changes, destructive revision disposition, auxiliary
  story length changes, mixed character formatting, and changes to internally modeled CP
  boundaries. Write ✅ is bounded text/format/metadata transactions only.

### Spec areas with no code and no matrix row (real structural gaps)

All are FIB pointer tables in the table stream; none affect body-text semantics:

- **RgDofr/Dofr record group** (§2.9.55–2.9.63, fcRgDofr — frameset and list-style
  support records): no parsing at all.
- **PGPArray/PGPInfo/PGPOptions** (§2.9.187–2.9.189, fcPlcfPgp — paragraph-group
  borders/margins for e-mail bodies): only the PGPInfo identifier references in
  sprmPIpgp/sprmTIpgp are modeled; the table itself is unparsed.
- **Selsf** (§2.9.244, fcWss — saved selection state): unparsed.
- **Print environment**: PrDrvr/PrEnvPort/PrEnvLand (§2.9.211–2.9.213): unparsed.
- **VBA digital signature storage** ([MS-OSHARED] §2.3.2 DigSigInfoSerialized/
  DigSigBlob/WordSigBlob): implemented in `litchi-ole-common/src/vba_signature/` but not
  exposed through the litchi-doc public API; a gap by the matrix's own accounting rule.

### Implemented in source but missing matrix rows

AutoText (PlcfGlsy/SttbfGlsy/SttbGlsyStyle/LEGOXTR_V11), Word 2003 XML schema references
(Hplxsdr/XSDR/TIQ), structured document tags (SttbfBkmkSdt/PlcfBkf/BklSdt/SDTI), format
consistency checker bookmarks (SttbfBkmkFcc/DPCID), text services framework
(Plcfuim/PlfguidUim/UIM), grammar checker cookies (Plcfcookie), language auto-detection
(Plcflad), table character cache (PlcfTch), revision threading (RmdThreading), bookmark
repair (SttbfBkmkBPRepairs), annotation bookmarks (SttbfAtnBkmk), embedded fonts
(SttbTtmbd).

### Key partial (🟡) boundaries

IRM DataSpaces (read ✅/write ❌, metadata graph only); OfficeArt write-out (bounded
shape kinds in main/header stories only); floating shapes/text boxes (no full geometry,
WordArt, rotation, or z-order model); embedded font programs (table/licensing metadata
only); master-document subdocuments (semantic edit yes, referenced files never opened);
mail merge ODSO (typed metadata, no data-source connection or merge execution);
ActiveX/OCX (inert metadata only); document/range protection (settings readable and
writable, edit policy not enforced); command bars/key bindings/menus/toolbars (inert
metadata, no macro or UI execution); routing slip (lifecycle changes refused when
protection policy is not Off); unknown streams/storages (topology preserved best-effort,
no semantic model); document variables/attached templates/web export metadata (values
editable, templates never loaded).

---

## 2. litchi-docx (WordprocessingML .docx) — [MS-DOCX], [MS-OWEXML]

Every [MS-DOCX] extension section (§2.1–2.13) has a matrix row; there is no whole-chapter
blind spot in the extension layer. Real gaps are (1) the declared execution/rendering ❌
set and (2) base ISO WordprocessingML feature families that lack dedicated matrix rows.

### Substantive gaps

- **Ink / inkML** (read ❌/write ❌): no ink model anywhere in the crate; ISO 29500 ink
  part and `w:ink` contexts unsupported.
- **Revision UTC timestamps** (read ❌/write ❌, [MS-DOCX] §2.12): `w16du:dateUtc`
  (Word 2023 namespace) is unhandled for tracked changes; only modern comments carry
  dateUtc.
- **[MS-OWEXML] custom-function payloads** (read ❌/write ❌): §2.2.11
  `CT_ContainsCustomFunctions`, §2.2.12 `CT_BackgroundAppData`, §2.2.13
  `CT_CustomFunctionList` have no typed models.
- **Revision authoring and accept/reject** (write gap): `w:ins`/`w:del`/`w:moveFrom`/
  `w:moveTo`/`rPrChange`/`pPrChange` have typed read modules (`src/revision.rs`), but
  general revision authoring and accept/reject are unavailable.

### Declared ❌ rows (deliberate)

Layout/pagination/rendering; chart/SmartArt/DrawingML/VML/embedded-workbook rendering;
field calculation and refresh (ISO 29500 §17.16); mail-merge execution; AltChunk import
and foreign-content conversion (ISO §17.17.2.1); macro/ActiveX/form-control/add-in
execution; external relationship target fetching.

### Spec areas without dedicated matrix rows (auditing blind spots)

Revision model as a whole, field model (`src/field/` exists), OMML math
(`src/math.rs`, 642 lines), bibliography sources, smart tags, bookmarks, text boxes,
custom XML data storage (`src/custom_xml.rs`), and the operational APIs (streaming,
redact, sanitize, statistics, validation). [MS-OWEXML] is summarized in a single row
without per-element coverage of webextension/taskpanes/webextensionref.

### Key partial (🟡) boundaries

Main-document transaction model (checked leaf operations only); SDT extension controls
(`dataBinding`/`appearance`/`color` typed; checkbox/entityPicker/repeatingSection
contextual semantics inert); SDT web-extension links (OnOff metadata only); W14 conflict
revision markers (authoring limited to non-nested inline/range conflicts and table-row
property conflicts; no accept/reject); stylesWithEffects part (topology only); extended
number formats (common typed enum, not the full Microsoft extension enum); Word 2023
dateUtc (modern comments only); VBA projects (metadata and source payload, never
compiled); IRM/protected content (bounded codec, no rights evaluation); web
extensions/Office add-ins (task-pane graph editing, callbacks/commands inert);
protection enforcement (settings writable, policy not enforced); signature trust
(integrity verification only); Markup Compatibility (bounded MCE branch selection);
secondary-story edits (bounded selectors for footnotes/endnotes/comments/glossary only).

---

## 3. litchi-xls (Excel BIFF .xls) — [MS-XLS]

Nearly all ❌ rows are deliberate scope exclusions. The only single-direction ❌ is
embedded payload authoring (read 🟡/write ❌).

### Declared ❌ rows

Threaded comments/persons/mentions (XLSX-era parts; no BIFF8 records exist — effectively
N/A); slicers/timelines (no typed cache or view model); full formula evaluation engine;
rich data types and dynamic-array spill semantics; chart rendering and complete fresh
chart authoring (spec §2.2.3); macro/control/DDE/RTD/database/external-link execution
(spec §2.2.7–2.2.8); certificate trust and revocation; embedded payload creation (spec
§2.1.7.5 — the only single-direction ❌).

### Spec areas with no code and no matrix row

- **IRM/DRM** (§2.1.7.3 Data Spaces storage, §2.1.7.13 Protected Content stream
  `009DRMContent`, §2.1.7.19 DRM Viewer Content): only an existence warning in
  `validation.rs`; no typed model.
- **CFB-level streams** (§2.1.7): Component Object stream (`001CompObj`, §2.1.7.1),
  Control stream (`Ctls`, §2.1.7.2), Link Storage (§2.1.7.7), XML Signatures storage
  (§2.1.7.21) — unhandled.
- **List Data stream** (§2.1.7.8, SharePoint list-sync XML): no matches in source.
- **Shared features / smart tags** (§2.2.12, FeatHdr/Feat/Feat11/Feat12 with factoid
  smart tags): only display flags exist; no typed smart-tag model.
- **User Names stream** (§2.1.7.17, shared-workbook user log): partially handled via
  revision records, but the stream itself is not disclosed in the matrix.

### Key partial (🟡) boundaries

Formula tokens (writer tokenizer is a restricted subset: constants, Ref/Area/Area3d,
basic operators, fixed-arity Func, Paren/MissArg — no names, 3-D references,
Attr/Choose/If, intersection/union, Err, or mem tokens; never evaluated); array/shared/
data-table formulas (shared formulas only via ShrFmla owner path; no array resize/add);
BIFF charts (chart-area snapshot/transactions and bounded series edits only); OfficeArt
(typed geometry/anchors/text extraction, no full drawing-group graph operations); OLE
objects/form controls (metadata edits only, never activated); VBA (topology inspectable
and replaceable, never executed; replacement invalidates signatures); query tables/
external connections (typed records, no commands/credentials/refresh); pivot caches and
pivot tables (typed read/write, no refresh or calculation); XML maps (typed MapInfo/
schema/binding, no schema resolution or binding refresh); revision log/shared workbook
(bounded edits, no conflict resolution); RTD (bounded edits, server never refreshed);
conditional formatting (classic + BIFF12 metadata, rule formulas never evaluated);
unknown/FRT records (selected families preserved, not guaranteed to survive every typed
mutation).

---

## 4. litchi-xlsb (Excel binary OOXML .xlsb) — [MS-XLSB]

All explicit ❌ rows are the deliberate execution/evaluation/rendering/encryption/trust
exclusions. The functional blind spots are in the unlisted parts below.

### Declared ❌ rows

Full formula evaluation and recalculation; pivot refresh/calculation; rich values and
modern data types (§2.1.7.34, §2.2.4); complete chart grammar and rendering; mapped-XML
import/export, schema/XPath evaluation, custom XML data storage processing; ActiveX/form
control execution; macro execution and external link/connection refresh; password
encryption packages (§2.2.11 + [MS-OFFCRYPTO]); certificate trust and revocation.

### Spec parts with no code and no matrix row (largest blind spots)

1. **§2.1.7.35 Model (spreadsheet data model / PowerPivot)** — `BrtModelTable`,
   `BrtModelRelationship`, `BrtModelTimeGroupingCalcCol`, DAX measures: the entire Data
   Model family is absent. Only a connection-source enum mentions the data model.
2. **§2.1.7.31–33 Macro Sheet / International Macro Sheet / Macro Sheet Binary Index and
   §2.1.7.20 Dialog Sheet** — XLM macro sheet content is unparsed; the package layer
   recognizes the content types and refuses them on edit paths.
3. **§2.1.7.60 + §2.2.13 Volatile Dependencies part** — no `Volatile` matches in source.
4. **§2.1.7.4 Calculation Chain** — appears only as a refuse-on-escape target during
   chart transfer.
5. **§2.1.7.63 Worksheet Binary Index** — the writer emits a fixed 29-byte empty
   constant; the reader and matrix have no row.
6. **§2.1.7.3 Attached Toolbars** — zero source matches.
7. **§2.1.7.10/11 Custom Data + Custom Data Properties** — zero source matches.
8. **§2.1.7.16–19 Diagram Colors/Data/Layout/Styles (SmartArt)** — the drawing row only
   says non-chart graphic frames are a typed refusal; the diagram part family is never
   named.
9. **§2.1.7.49 Sort Map** (shared-workbook sort mapping) — zero source matches.
10. **§2.1.7.52 Theme part** — the writer generates a template `theme1.xml`, but
    reading/replacing the theme part has no matrix row.
11. **§2.1.7.9 Control Properties (ActiveX property bags)** — gap rows cover control
    "execution" only; property metadata reading is unaddressed.

### Key partial (🟡) boundaries

Formula tokens and cached values (never evaluated); array/shared/dynamic-array formulas
(records preserved, no spill calculation); AutoFilter/sort state (metadata only);
chart family (bounded resource discovery, byte-level transfer; no arbitrary chart
editing); slicers/timelines (bounded CRUD; no OLAP item authoring, filter execution, or
refresh); external links (unsupported link forms refused; formulas never re-resolved);
VBA projects (bounded rewriting, old signatures dropped on replacement); worksheet
protection (typed flags, not a cryptographic boundary); shared-workbook revision records
(lazy metadata, no collaboration); MDX/cell/value metadata (raw inspection, MDX never
evaluated); core/package properties (generic OPC editing only); OLE/embedded packages
(inventory + byte-level carry, no creation/activation).

---

## 5. litchi-xlsx (SpreadsheetML .xlsx) — [MS-XLSX], [MS-OREACTXML]

### Declared ❌ rows

Full formula evaluation (no dependency graph, volatile recalculation, dynamic-array
spill, or external functions — [MS-XLSX] §2.2.2–2.2.3); PivotTable refresh/calculation/
write-back (§2.3.1, §2.3.4); consolidation and what-if calculation; Python in Excel and
external code services (§2.1.23, §2.3.10, `python`/`externalCodeService` elements —
zero source matches); external link/connection/query/DDE/OLE/add-in execution; macro and
trusted ActiveX execution; chart/drawing/slicer/timeline/sparkline rendering and full
extension grammars (ChartEx unsupported); certificate trust chains and revocation;
**Surveys semantic authoring (read ✅/write ❌)** — the only directional ❌ outside the
explicit-gap set (§2.1.9).

### Spec areas with no code and no matrix row

1. **Comment reactions ([MS-OREACTXML])** — `reactions`/`commentsExtensible` (Like
   reactions, §2.1 + §5.1 schema): zero source matches; the threaded-comments matrix row
   lists threads/persons/mentions/replies but not reactions. Read ❌/write ❌.
2. **Newer pivot extension elements** (all zero matches): pivotTableServerFormats
   (§2.4.2), pivotTableData (§2.4.63), cachedUniqueNames (§2.4.61),
   implicitMeasureSupport (§2.4.91), aggregationInfo (§2.4.104), featureSupportInfo
   (§2.4.105), autoRefresh (§2.4.106), and the 2025 pivotDataSource family
   (pivotAreaReferenceSubtotals §2.4.111, pivotCacheDataSource §2.4.112,
   pivotFieldSubtotalLineItems §2.4.113, pivotFieldSubtotals §2.4.114).
3. **Data-model/connection/external-link minor extensions**: modelTimeGroupings
   (§2.4.71), refreshIntervals (§2.4.92), alternateUrls (§2.4.97),
   showDataTypeIcons/showDataTypeIconsCustomSheetView (§2.4.108/109) — all zero matches.
4. **formControlPr extension (§2.4.34)** — zero matches; only a generic ActiveX/form
   control metadata 🟡 row exists.
5. **_xlfn extension-function catalog (§2.2.3)** — formula text passes through
   transparently; worth an explicit matrix row.

Note: [MS-OXCDATA] is an Exchange mailbox ROP data-structures specification and does not
apply to .xlsx; it is N/A for this scope.

### Implemented in source but missing matrix rows

Data Model part (§2.1.6, `src/workbook/data_model/` + `src/package/xldm/`) and Custom
Data / Custom Data Properties parts (§2.1.2/2.1.3, `src/custom_data/`).

### Key partial (🟡) boundaries

Formulas (text, shared/array ownership, bounded edits; no evaluation); conditional
formatting (core direct collections on existing normal sheets only; x14/MCE owners and
invalid DXF references refused); hyperlinks (existing-sheet metadata edits only;
relationship create/delete/redirect refused); charts (closed relationship-graph CRUD;
ChartEx and nested user-shape dependencies refused); drawings/anchors (cross-sheet copy
translates selected anchors only); pivot caches/views (bounded metadata; no refresh,
cube, or unknown extensions); slicers/timelines (full graph CRUD, but filtering/
refresh/rendering inert); rich values (rvData/rvStructures/arrayData typed; rich styles
and supporting bags opaque); feature property bags/checkboxes (lazy typed + validation);
workbook properties (custom properties typed; core/extended properties via
`into_plain_opc()` raw boundary); defined names (lazy catalog + structure-safe remapping,
no fine-grained formula construction); revisions (users/headers/lazy log CRUD; no
real-time collaboration); sparklines (bounded groups/relationships; no rendering or
source calculation); data consolidation (metadata only); VBA/ActiveX/form controls
(inert metadata); digital signatures (inspect/edit, no trust chain); web extensions/
task panes (CRUD, no add-in activation); unknown package parts (preserved as opaque, no
semantic rewrite guarantees).

---

## 6. litchi-ppt (PowerPoint binary .ppt) — [MS-PPT], [MS-OGRAPH]

### Substantive gaps

- **Complete native chart authoring (write ❌, read 🟡)** — the largest gap.
  `PptWriter::add_chart` returns `litchi_ograph::Error::UnsupportedAuthoring` for every
  structurally valid request. Fresh encoding of all 107 [MS-OGRAPH] §2.4 record families
  (Chart/Series/Axis/Legend/DataFormat/Pie/Bar/Scatter/Radar/Surf/Chart3d/Dat) is
  unimplemented, and arbitrary semantic rewriting of parsed charts is unsupported.
  Excel-hosted charts (`Excel.Chart`) have a read inventory only, with no host write-back
  staging.
- **Native diagram / SmartArt authoring** (read ❌/write ❌): read-only inventory plus
  fixed-width build-metadata edits only ([MS-PPT] §2.8.13–2.8.14, §2.13.7).
- **Modern comments** (read ❌/write ❌): effectively N/A — a PresentationML feature;
  binary Comment 2000 is ✅.
- **Picture bullets** (§2.9.72–2.9.73 `BlipCollection9Container`/`BlipEntityAtom`):
  parse-only, no `to_record` — a small real write gap, with no matrix row.

### Declared N/A exclusions

Media playback/rendering/external activation, slideshow rendering, external resource
resolution, OLE/macro/script execution — deliberate inert-data policy.

### Implemented in source but missing matrix rows

Photo Album (§2.4.9 `PhotoAlbumInfo10Atom` + §2.13.19–20 — 29 implementation sites in
`document_properties.rs`); Presentation Advisor (§2.4.6 `PresAdvisorFlags9Atom`); sound
collection records (§2.4.16, §2.11.29 `SoundDataBlob`); RecolorInfo family
(§2.7.9–2.7.13); MetafileBlob (§2.11.6 — inert only); Graph workbook non-chart substreams
(MS-OGRAPH datasheet substream §2.4.14 `BOFDatasheet` and window/view state records
§2.4.64/104–106); custom table styles round-trip (§2.11.13
`RoundTripCustomTableStyles12Atom` — exported but no matrix row).

### Key partial (🟡) boundaries

Native Graph/Excel charts (read: typed lazy inventory + borrowed semantic view via
litchi-ograph; write: replacement of existing standalone Graph chart substreams only);
diagram build records (fixed-width fields within owning-slide envelopes only);
title/notes/handout masters (bounded SlideNameAtom/notes `txStyles` writing); action/
interaction settings (validated canonical records only; macro/program/active-OLE actions
refused); audio/video external objects (metadata validated; lazy path/flag edits only);
embedded fonts (record-level read/write + EOT 1.0 validation + fsType licensing
decisions; deletion/reorder refused when remapping is unproven); document-compare
metadata (no diff generation or accept/reject workflow); broadcast/HTML publishing/
routing/envelope/privacy (fully typed, no network/browser/mail workflows); color schemes
and PP12 themes (bounded round-trip, not a full OOXML theme model); OfficeArt unknown
records (inert preservation); transitions (spec table type/direction/speed writable;
record insertion unsupported); text/shape editing (length changes limited to unstyled or
single-paragraph/single-run closures).

---

## 7. litchi-pptx (PresentationML .pptx) — [MS-PPTX]

The matrix has a single declared ❌/N/A row (rendering, layout, playback, action
execution — deliberate). The real gaps below are matrix blind spots, verified as zero
implementation by source grep.

### Spec features with zero implementation (read ❌/write ❌)

| Feature | Spec location | Notes |
|---|---|---|
| **Morph transition** (`p159:morph`, PowerPoint 2016) | §2.6 / §2.2.1 | Most significant gap; common in modern .pptx files |
| **Preset transitions** (`p15:prstTrans`, 2012) | §2.4 / §2.2.1 | Zero matches |
| **Read-only recommended** (`p1710:readonlyRecommended`, 2017/10) | §2.14 | Zero matches |
| **Placeholder type extension** (`p232:phTypeExt`, 2023/02, CT_PlaceholderTypeExtension/ACB) | §2.22 | Newest schema, lowest impact |
| **2012 collaboration state** (`p15:presenceInfo` / `threadingInfo`) | §2.4 / §2.2.10 | Matrix modern-comments row covers only the 2018/8 model |
| **Media black-and-white playback** (`p14:bwMode`, 2010) | §2.3.2.2 / §2.2.4 | Trim/fade/bookmark are ✅; only this attribute is missing |

### Key partial (🟡) boundaries

PowerPoint 2010 extended transition effects (only `p14:ripple` is typed; the other 17 —
conveyor/doors/ferris/flash/flip/flythrough/gallery/glitter/honeycomb/pan/prism/reveal/
shred/switch/vortex/warp/wheelReverse/window, §2.3.1 — are raw-XML preservation with fade
fallback); zoom objects (typed targets and property CRUD; no rendering/layout); math
extension a14:m (presentation-level `brkBin`/`brkBinSub` only); 3D models (lazy GLB/
preview resources; no animation semantics or rendering); Designer family (typed codec +
transactions, but no public package-level facade for new owners); media tracks TracksInfo
(caption identity/language/display position and `isNarration` only); revision/change
information (source-validated metadata CRUD, not a collaboration engine); cross-slide
source-backed copies (narrow set of direct `p:pic` leaves and plain chart frames;
ChartEx, embedded workbooks, externalData refused); ActiveX controls (metadata/binary
replace/detach, never instantiated); VBA macros (parseable and authorable, never
executed); web extensions/Office add-ins (lazy CRUD, never loaded); action/interaction
settings (inert); presentation size/view settings (typed, no host-UI driving).

---

## 8. litchi-rtf — RTF 1.9.1, [MS-OXRTFCP]

**Important premise: this crate has no `docs/FEATURE_MATRIX.md`.** The top-level matrix
classifies litchi-rtf under "Conversion and interchange". This audit therefore used the
crate README and public API against the RTF 1.9.1 table of contents, verifying each
chapter against the lexer dispatch table (1,394 control-word dispatches). Unknown control
words and `{\*...}` destinations are preserved byte-for-byte as `opaque::Node` and can be
written back, so ❌ below means "no semantic API", not data loss.

### Spec chapters with no semantic read/write API (opaque preservation only)

| Topic | Direction | Spec page | Evidence |
|---|---|---|---|
| Positioned objects and paragraph frames (`\posx/\posy/\posxc…\absw/\absh/\phcol/\phmrg/\phpg/\pvmrg/\pvpara/\pvpg/\dxfrtext/\dfrmtxtx(y)/\frmtxlrtb…/\wraparound/\wrapthrough/\wraptight/\abslock/\absnoovrlp/\nowrap`) | read ❌/write ❌ | p.91–93 | Only `\dropcapt` exists in dispatch; the frame family is absent |
| SmartTag data (`\factoidname` and factoid bookmarks) | read ❌/write ❌ | p.153 | No `factoid`/`smarttag` matches anywhere |
| Move bookmarks (`\mvfm/\mvlt/\mvtb/\mvte/\mvfml/\mvtl`) | read ❌/write ❌ | p.146–148 | Only plain `Bookmark/BookmarkTable` |
| New-style protection hash `\passwordhash` (SHA-1 spin-count) | read ❌/write ❌ | p.40–41 | Only legacy `\password` hex kept inert |
| Word 6J-era East Asian legacy control words (`\jsksu/\jsku`, `\horzvert`, `\gcwN`, `\twoinoneN`, `\nosectexpand`, `\sectexpandN`, `\jclisttab`) | read ❌/write ❌ | p.204–206 | Modern equivalents (`\fchars/\lchars` kinsoku, document grid, character expansion) are covered |

All other spec chapters have corresponding types or tests. [MS-OXRTFCP] compressed RTF is
fully implemented in both directions (including the 256 MiB extended limit) — no gap.

### Missing documentation (the matrix itself is the blind spot)

The most important undocumented areas: paragraph frames, SmartTags, move bookmarks, and
the `\passwordhash` half-gap above; Quick Styles / Table Styles / style-and-formatting
restrictions (implemented: `\sqformat/\spersonal/\scompose/\sreply`, table-style
conditional formatting, `DocumentStyleRestrictions`); read-only password protection
(legacy `\password` inert only); heuristic chapters (ShiftJIS font inference without
`\cpgN`, composite/associated fonts, Word 6J/2000 East Asian control words); Macintosh
Edition Manager Publisher objects (implemented as `ObjectKind::Publisher`).

### Key partial (🟡) boundaries

Fresh authoring is limited to `streaming::StreamingRtfWriter`/`write::Writer`
(paragraph/run/basic `Format` plus hyperlink/note/revision helpers) and `tail_append`;
tables, lists, shapes, math, and fields have no typed from-scratch builders (full
model-internal write-back of parsed documents is ✅). `Document::edit()` bounded
transactions cover disjoint body UTF-8 spans, paragraph alignment, bold/italic ranges,
paragraph insertion, table-cell text, header/footer text, comment bodies, footnote/
endnote stories, and lazy root shape text boxes; everything else is read-only. Canonical
edits fail closed on snapshots containing unknown syntax; only exact body splices pass
through unknown grammar. `\password` is kept inert, never interpreted. OLE objects/
external links/fields/macro payloads are inert per the project-wide boundary.

---

## 9. ODF text family (litchi-odt, litchi-oth, litchi-odm) — OpenDocument v1.4

Chapter numbers refer to OpenDocument-v1.4-os part3-schema.

### Declared ❌ rows (litchi-odt)

- **Field evaluation and refresh** (read ❌/write ❌, §7): no field is ever recalculated;
  only cached values.
- **Index/TOC generation and refresh** (read ❌/write ❌, §8.3–8.9): structure and cached
  bodies are readable/writable (🟡) but entries are never generated and page numbers
  never computed.
- **Document-statistics recomputation** (read ❌/write ❌, §4.3.2.18, §7.5.18).
- **Change-tracking accept/reject and merge engine** (read ❌/write ❌, §5.5):
  declarations and change marks have CRUD ✅, but no semantic accept/reject.
- **Form runtime behavior** (read ❌/write ❌, §13): no control display, validation,
  submission, or active-state maintenance. OOXML/ActiveX controls read 🟡/write ❌.
- **Macro/script/event execution** (read ❌/write ❌, §3.12–3.13, §7.7.9–7.7.10, §14.5).
- **Embedded object activation/conversion** (read ❌/write ❌, §10.4).
- **Mail merge** (read 🟡/write ❌, §7.6 + §12): data sources never opened.
- **External link refresh** (read 🟡/write ❌, §5.4.2, §10.4).
- **Certificate trust chains/revocation/identity policy** (read 🟡/write ❌).
- **Flat text templates (.fott)** (read ❌/write ❌): `FlatDocument` refuses the template
  MIME classification; an API-scope refusal rather than a format impossibility.
- **Pagination, line layout, font shaping, visual rendering; table auto-layout; chart
  calculation and rendering** (all ❌).

### litchi-oth gaps

- **Encryption and signatures as a whole** (read ❌/write ❌): password opening,
  verification, signing, and changed-source publication are all refused (inert inventory
  only).
- **Undeclared body blind spots** (confirmed absent from the `Element` enum): tables
  (§9), sections `text:section` (§5.4), footnotes/endnotes (§6.3), annotations (§14.1),
  change tracking (§5.5), indexes/TOCs (§8), frames/text boxes (§10.4), ruby (§6.4).
  Writer/Web documents legitimately carry these structures (tables especially), yet OTH
  preserves them only verbatim with no typed projection, and the matrix declares neither
  ❌ nor N/A.

### litchi-odm gaps

- **Encryption/signature writing** (write ❌): signature verification/signing/re-signing
  and encrypted output are absent; encrypted packages can be opened with a password for
  inert inspection only (read 🟡), and any change is refused.
- **Declaration containers** (§5.7 variable/sequence declarations): read-only.

### Spec areas with no code and no matrix row

1. **In-content RDFa metadata (§4.2.1) and the `text:meta` element (§6.1.9)** — the
   matrix "RDF metadata graphs ✅" row covers only manifest-declared package-level `.rdf`
   graphs; `xhtml:about/property/content` on bookmarks/paragraphs and the `text:meta`
   inline container have no typed model. The ✅ mark overstates coverage.
2. **`xforms:model` (§13.4)** — the forms chapter never mentions XForms; source only
   detects its presence (`has_xforms` flag) with no model/instance/bind/submission
   model.
3. **`text:number` (§6.1.10, ODF 1.3+ in-paragraph list-number element)** — no element
   model anywhere in the repository.
4. **3D graphics (§10.5, dr3d) and custom-shape enhanced geometry (§10.6)** — the shape
   row enumerates rectangle/ellipse/path/connector/custom-shape but never dr3d
   scene/cube/sphere or enhanced-geometry equations/handles.
5. **Property-level coverage (§19–§20)** — the matrix marks "property sets" ✅ as a
   whole, but many individual §20 attributes (`text:animation-*`, `draw:allow-overlap`,
   `style:margin-gutter`, `table:tab-color`, `draw:decorative`) are not typed; actual
   coverage is below what the matrix suggests.
6. **`text:soft-page-break` (§5.6)** — recognized only inside field-cache/ruby
   whitelists; no standalone read or CRUD API.

### Key partial (🟡) boundaries

ODT indexes/TOCs (structure/templates/cached bodies; no generation); ODT field cached
values (never evaluation results); ODT table formulas (attributes and cached cells
preserved, no recalculation); ODT repeated rows/columns (expanded semantic access,
bounded write-back, no repetition-policy layout semantics); ODT footnotes/endnotes
(identity/reference/plain-text/bounded rich-text CRUD; nested fields/links/scripts
inert); ODT geometric shapes (preservation + bounded properties, no full CRUD); ODT
embedded charts (subdocuments/series/axes modeled via `odf-common::chart::authoring`; no
data calculation); ODT style registry (declarations and relationships; cascade
resolution not guaranteed); ODT database fields/data sources (metadata CRUD, no
connection/query/import); ODM encryption/signatures (inert inventory + no-change
preservation only); ODM paragraph/heading inline content (deliberately open as mixed
pass-through); OTH styles (family/parent + small RGB/bold/italic subset only); OTH
fields (common families with stored values only; uncommon fields refuse rewriting).

---

## 10. ODF spreadsheet (litchi-ods + litchi-odf-formula) — OpenDocument v1.4

### Declared ❌ rows

- **Formula evaluation and recalculation** (read ❌/write N/A, Part 4 in full): formulas
  and cached values are fully lazy; §9.4.1–9.4.3 covers only calculation-settings CRUD.
- **Full OpenFormula semantics** (read 🟡/write ❌, Part 4 ch. 5–8): `codec::formula` is
  a parser, not an evaluator; external workbook references, dynamic arrays/array
  expressions (§3.3, §5.13, §7), volatile functions, multiple operations, and
  host-defined functions are unparsed. The function catalog is a static whitelist of
  ~150 common functions (`src/codec/formula.rs`); Part 4 ch. 6 families (matrix 6.5,
  bitwise 6.6, complex 6.8, database 6.9, external access 6.11, base conversion 6.19)
  are largely absent.
- **Scenarios / consolidation / label ranges / audit detective** (one matrix row:
  🟡/🟡/❌): scenarios (§9.2.7) are inspect-only with no apply;
  `model::{consolidation,label_range,detective}` have fragment-level codecs but no
  facade entry points (zero facade references) and no public CRUD — write ❌.
- **DDE** (§9.8): read ✅ (bounded inspection) / write ❌ — no sessions, no refresh, no
  public mutation path.
- **CSV export** (read ❌/write ❌): not a spec feature; the crate does not provide it.
- **Rendering/pagination/recalculation engines** (❌/❌).
- **OOXML-specific families** (matrix "explicit gaps", ❌/❌ — no ODF counterparts):
  slicers/timelines, PivotTable caches, OLAP/data-model extensions, query tables,
  connections, external links, web/rich values, external code services (incl. Python),
  threaded comments/mentions/property bags.
- **Digital-signature signing/re-signing** (❌): change publication refuses signed
  inputs by default or requires explicit stripping; protected-structure unlocking and
  transactional re-encryption of encrypted sources are ❌.
- **litchi-odf-formula**: OpenFormula spreadsheet grammar read ❌/write ❌ (by design,
  belongs to litchi-ods); StarMath preserved as opaque annotation only.

### Spec areas with no code and no matrix row

- **§9.1.13/9.1.14 `<table:title>`/`<table:desc>`** (sheet accessibility title/
  description): zero modeling.
- **§16.20 `<table:table-template>` + banding-style attributes (19.741–19.746)**:
  `styles::table_template` has a typed semantic model and codec but is not wired into
  the facade or unified transactions — implemented but unregistered.
- **Flat single-XML ODS documents** (Part 3 §3.1.2, Part 2): `src/flat.rs` implements
  `FlatSpreadsheet` (Snapshot/Transaction with sheets/dde/scenarios accessors), but the
  matrix scope statement covers only "packaged ODS/OTS" — an implemented-but-
  unregistered blind spot.
- **§9.3.1 `<table:cell-range-source>`** (cell-level external range source): modeled in
  `model/source.rs`/`model/cell.rs`, but no explicit matrix row.
- **Wider in-sheet drawing families (ch. 10)**: `draw:custom-shape`/enhanced geometry
  (10.6), dr3d 3D shapes (10.5), client-side image maps (10.4.13), caption/measure/
  regular-polygon/glue-point, applet/plugin/floating-frame — zero typed support.
- **Data-style vocabulary gaps (§16.29, 19.34x)**: `number:fraction`,
  `number:scientific-number`, `embedded-text`, and `transliteration-*` attributes are
  unsupported.

### Key partial (🟡) boundaries

Database ranges/filter/sort/subtotals (§9.4.14–9.5): fragment-level parse/write codecs
and semantic types exist, but no package-level facade CRUD (the end-to-end test in
`model/database_range/tests.rs` is disabled via `#[cfg(any())]`); no query execution.
DataPilot (§9.6): metadata CRUD ✅, but pivot calculation, refresh, rendering, and
external-source execution are inert. Protection: flags and verifier metadata editable,
passwords never verified. Data styles: only a closed family of automatic styles
(number/date/time/currency/percentage/boolean/text) can be added/removed/replaced;
page/master/layout families are preserved only. Conditional styles `style:map`/calcext
conditional formats/sparklines: lazily modeled or whole-catalog replacement; conditions
never evaluated. Content validation (§9.4.4–9.4.8): conditions/help/error messages
modeled, never enforced. Cell rich text: bounded paragraphs/spans/whitespace/safe
hyperlinks writable; arbitrary fields and out-of-line inline markup not structurally
editable. In-sheet graphics: only rectangle/ellipse/line/connector/polygon with limited
geometry; arbitrary paths and linked-resource fetching unsupported. Positional cell
publication / cross-document sheet transfer: existing plain scalar cells only; repeated
rows, formulas, merges, protection, encryption, and signature changes are refused.
Encryption/signatures: password reopen and terminal encryption available; change
publication requires explicit stripping of invalidated signatures; signing itself
unsupported. styles.xml/meta.xml/settings.xml: inspection and package-level preservation;
`meta.xml` has bounded public-field CRUD; `settings.xml` (incl. §3.10 config items and
cursor positions) is lazily preserved.

---

## 11. ODF presentation/drawing/image (litchi-odp, litchi-odg, litchi-odi) — OpenDocument v1.4

### Declared ❌ rows

**litchi-odp**

- Document/slide protection and editing restrictions (read ❌/write ❌): no protection
  model.
- Encrypted-package authoring (read ✅/write ❌): password opening only; transactions
  and patches refuse encrypted sources.
- Digital-signature verification/signing/removal (read 🟡/write ❌): uniformly
  `CryptoCapability::Refused`.
- Macro/script execution, hyperlink/external-media/DDE/database fetching (read ❌/write
  ❌, deliberate).
- Rendering/layout/animation playback/media playback/chart rendering (read ❌/write ❌,
  out of scope).
- Non-ODF rows (PPTX sections/zoom objects, Morph and vendor transitions, threaded
  comments/collaboration, media bookmarks/clips, guides/UI state): no ODP counterparts;
  listed as a gap checklist only.

**litchi-odg**

- Password change/re-encryption of existing encrypted packages (write ❌, finally
  unsupported).
- Transactional re-signing of existing packages (write ❌, finally unsupported; fresh
  package signing is supported).

**litchi-odi**

- Fresh encryption, signing/re-signing, trust verification, change publication of
  protected snapshots (read ✅/write ❌): rewrites are refused rather than degraded.
- Direct mutation of forms/producer extensions (read ✅/write ❌, §13, §3.17): inventory
  only.
- Active content (script/event/macro/DDE) mutation (read ✅/write ❌, §3.12, §10.3.19):
  inert inventory only.
- Rendering and image conversion (read ❌/write ❌, out of scope).

### Spec areas with no code and no matrix row

- **Contours (§10.4.11, `draw:contour-polygon`/`draw:contour-path`)** — zero matches
  across all three crates; a genuine blind spot.
- **Client-side image maps (§10.4.13, `draw:image-map`)** — covered in ODI ✅, but zero
  matches in ODP and ODG; the spec allows image maps on presentation/drawing pages, and
  neither matrix has a row.
- **Glue points (§10.3.16, `draw:glue-point`)** — ODP preserves them lazily at the XML
  validation layer only; ODG does not recognize them at all. Connectors (§10.3.10) are
  covered, but connector adhesion semantics are missing.
- **ODP layer declarations (§10.2.2/10.2.3, `draw:layer-set`/`draw:layer`)** — ODG
  covers layers explicitly; ODP handles only the `draw:layer` attribute on shapes, with
  no typed layer-set declaration CRUD.
- **ODG page transitions/animations (§19.396, §20.237–20.240, `presentation:transition-*`
  and SMIL on `draw:page`)** — Draw documents support per-page transition effects; ODG
  has neither coverage nor a matrix row.
- **ODG 3D shapes (§10.5, `dr3d:*`)** — ODP recognizes dr3d:scene/light/cube/sphere/
  extrude/rotate as typed kinds (🟡); ODG's `ShapeKind` has no 3D variants, so 3D shapes
  fall into "unknown markup preservation".
- **Custom-shape enhanced geometry (§10.6.2/10.6.3, `draw:enhanced-geometry`/
  `draw:handle`) in ODG** — ODP has an inert typed model (🟡); ODG has only a Custom
  kind + generic geometry (Fontwork survives only via byte-level provenance).
- **Standalone `draw:equation` (§10.2.5) and `draw:page-thumbnail` (§10.3.14)** —
  page-thumbnails are grouped under "advanced shapes" 🟡 in ODP/ODG; standalone
  equations are handled only inside ODP's enhanced-geometry children, with no page-level
  equation row.
- **ODI Relax NG conformance validation (spec Part 3 schema proper)** — the matrix
  itself states "Full Relax NG validation remains absent" — a structural-validation gap.

### Key partial (🟡) boundaries

ODP advanced drawing shapes (polygon/path/caption/measure/custom-shape/3D): kinds
recognized and preserved; geometry properties inert and unmodeled. ODP charts: typed
definitions can be added/removed/replaced with series/cached-cell CRUD and
cross-document migration; no recalculation or rendering. ODP common styles/data styles/
background/named drawing resources (gradients, fill images, hatch, markers): bounded
shared-style values and source-level named resources only; no full style-graph resolver.
ODP embedded objects/OLE/applet/plugin/floating-frame: inert inventory; payloads never
opened or activated. ODP encryption/signatures: open/recognize/enforce-read-only on the
read side; typed refusal on the write side. ODG package semantic model (write): unknown
markup gets exact byte provenance only; FODG flat-format writing is limited to
standalone text patch chains. ODG active content: inventory and preservation only;
preserve-or-refuse write policy. ODI raw content.xml: namespace structural validation +
the ODF 1.4 single-frame/single-image contract, without full schema validation. ODI
styles: frame graphic/text style references editable with cross-package migration and
closure merging, but general style CRUD and closure over gradients/master pages/fonts/
markers are absent.

---

## 12. ODF chart/database/formula documents (litchi-odc, litchi-odb, litchi-odf-formula) — OpenDocument v1.4

### litchi-odc (chart, part3 ch. 11) — declared ❌

- **Rendering and layout** (read ❌/write ❌): no chart renderer, layout engine, or style
  resolver.
- **Write-back of opened generic trees** (write ❌): the namespace-aware chart tree is
  read/preserve-only; writing requires flat exact-span edits or packaged transactions.
- **Writing encrypted/signed packages** (write ❌): signature metadata and manifest
  encryption entries are inventory-only; change publication refuses signed or encrypted
  packages; no password opening/writing and no signature creation/verification.
- **OTC chart templates** (read ❌): only the exact chart MIME is accepted.

### litchi-odb (database front end, part3 ch. 12)

No matrix row is ❌. Live database execution (drivers, network, credentials, query
execution, refresh) is a declared permanent boundary by design. Explicit typed refusals:
SQL analysis refuses CTEs, subqueries, set operations, table functions, multiple
statements, and external/unqualified relations.

### litchi-odf-formula (formula document, §14.6 + part4) — declared ❌

- **OpenFormula spreadsheet formula grammar (part4 in full)** (read ❌/write ❌): cell
  references, operators, the function catalog (all 20 families of part4 §6), array
  expressions, named expressions, the type system (§4), the evaluation model (§3), and
  non-portable features (§8) are unimplemented. A separate ODS-scoped OpenFormula (1.2)
  parser exists at `litchi-ods/src/codec/formula.rs`, but it belongs to litchi-ods and
  covers only the common ODS subset.

### Spec areas with no code and no matrix row

**litchi-odc**

- **§11.6 3D plot areas (dr3d scenes)**: `dr3d:light`, `dr3d:rotate`, `dr3d:shade`, and
  `chart:three-dimensional` — zero matches; preserved only as unknown nodes.
- **§20.1–20.73 + 20.436/20.437 chart:* formatting attributes (70+ items)**:
  error-category, logarithmic, regression-*, spline-*, symbol-*, treat-empty-cells,
  tick-marks — the entire chart style-attribute chapter has raw preservation only, with
  no property-level read/write API. The matrix summarizes this as one "Styles remain
  raw" line.
- **Draw shapes inside `chart:chart` and local cached tables**: RNG allows arbitrary
  shapes (draw:rect/line/caption) overlaid on charts; never mentioned. `chart:symbol-
  image` (§19.916.4 data-point symbol images) also absent.
- **Chart-document table-decls prelude** (office-chart-content-prelude:
  table:named-expressions, database-ranges): zero matches.
- **Deprecated attributes** (§19.16/19.17/19.26 chart:column-mapping,
  data-source-has-labels, row-mapping): presumably generic preservation (low impact).

**litchi-odb**

- **§12.8 db:login** (user-name, is-password-required, use-system-user, login-timeout,
  host-name/port/local-socket): zero matches; an extremely common connection-data child
  in real Base files. The typed whitelist in `catalog.rs` excludes it, so it falls into
  `Element::Other` generic preservation.
- **§12.9–12.12 db:driver-settings / db:auto-increment / db:delimiter /
  db:character-set**: the driver-settings family is unmodeled.
- **§12.13–12.14 db:table-settings / db:table-setting**: unmodeled.
- **§12.15 db:application-connection-settings + §12.16–12.21 table-filter family
  (db:table-filter, table-include/exclude-filter, table-filter-pattern,
  table-type-filter, table-type) + §12.22–12.24 db:data-source-settings**: the whole
  application-connection/data-source settings structure is unmodeled and undisclosed.
- **§13 form control model typing**: form:* controls inside db:component are preserved/
  transferred byte-for-byte with active-content inventory only; the matrix does not
  state that the §13 control model is unparsed.

**litchi-odf-formula**

No additional blind spots: the part4 gap is disclosed as a single ❌ row. The MathML
version matches the ODF 1.4 reference ([MathML] = MathML 2.0, 2003); all MathML 2
presentation elements (including maction, mglyph, maligngroup/mark, mlabeledtr) are
present in the source element tables.

### Key partial (🟡) boundaries

litchi-odc typed views (read 🟡): borrowed views cover chart, first plot-area/legend,
axes, grids, series, domains, and data points; the full typed-definition projection
opens only for canonical packages generated by this crate (byte-level reserialization
must prove losslessness), and all other opened files degrade to the generic preservation
tree. litchi-odc new definitions (write 🟡): the Builder covers the 12 §19.15 chart
classes plus title/plot/axis/series/label/trend/error/stock/wall/floor/cached-table;
3D scenes, embedded shapes, and named expressions cannot be created. litchi-odc
formulas/ranges: range-list syntax and cached-formula prefix/separator/reference
structure validated; never evaluated or refreshed. litchi-odc styles: styles.xml
wholesale preserve/validate/replace; no cascade or rendering resolver; `office:scripts`
always refused. litchi-odb package snapshots: strictly single db:data-source (per RNG)
with bounded structure; unconventional but legal packages may be refused. litchi-odb
new packages: only an inert data-source shell, not a full Base document. litchi-odb
encryption/signatures: encrypted packages openable but credentials unexposed; signatures
inventoried and cryptographically verified (no PKI trust claims); mutating encrypted
packages refused, mutating signed packages refused by default (explicit signature
removal opt-in). litchi-odb edits/patches: transactions, inverse patches, history, and
three-way conflict planning bounded by a 256 MiB / 65,536-operation budget; advanced SQL
shapes are typed refusals. litchi-odf-formula: positioned as a bounded checked MathML
tree, not a full MathML validator; `annotation-xml` payloads, free-text attributes, and
foreign-namespace extensions stay lazy, and some attribute value domains lack
schema-level validation.

---

## 13. Infrastructure and shared specifications

Scope: litchi-cfb ([MS-CFB]), litchi-crypto ([MS-OFFCRYPTO]), litchi-ole-common
([MS-OLEPS]/[MS-OLEDS]), litchi-vba ([MS-OVBA]/[MS-VBAL]), litchi-opc (ECMA-376 Part 2),
litchi-drawingml ([MS-ODRAWXML]), litchi-imgconv ([MS-WMF]/[MS-EMF]/[MS-EMFPLUS]), plus
[MS-DTYP]/[MS-LCID]/[MS-OAUT]/[MS-UCODEREF]/[MS-OE376]/[MS-OI29500]. None of these
crates has a FEATURE_MATRIX; findings come from READMEs, public APIs, and spec ToCs.

### litchi-cfb ([MS-CFB]) — essentially complete

- Directory-entry metadata: state bits and creation/modification times are parsed but
  not exposed (`DirectoryEntry` has no fields; `OleWriter` has no setters) — read 🟡/
  write 🟡 (non-zero values are lost on write).
- Range Lock Sector (§2.8): no dedicated handling; transactional CFB does not occur in
  practice (minimal impact).

### litchi-crypto ([MS-OFFCRYPTO])

- **Standard Encryption non-default configurations**: RC4 and AES-192/256 Standard
  EncryptionInfo are all `Unsupported` (§2.3.4.5) — read ❌/write ❌ (only AES-128/SHA-1
  is emitted).
- **Agile non-default configurations**: only the Office publishing default
  (AES-128/CBC/SHA-1 + dataIntegrity) is accepted; other cipherAlgorithm/hashAlgorithm
  combinations (SHA-256/512, AES-192/256) and legal profiles without dataIntegrity are
  `Unsupported` (§2.3.4.10) — read ❌/write ❌.
- **Legacy binary RC4 encryption** (§2.3.6, non-CryptoAPI): not implemented in the
  shared crate; XLS/DOC each implement XOR obfuscation (§2.3.7), but the shared crate
  lacks §2.3.6.
- **IRM content decryption** (§2.2.10–2.2.11 Protected/Viewer Content Streams, XrML
  licenses): deliberately inert ciphertext envelopes; licenses are never acquired and
  content never decrypted (a deliberate security boundary).
- Binary-document CryptoAPI signatures (§2.5.1 `_signatures`) and xmldsig (§2.5.2–2.5.3)
  are covered by the separate `litchi-sign` crate.
- VBA project signatures (DigSigBlob/WordSigBlob, [MS-OSHARED]):
  `litchi-ole-common/vba_signature` does inert parsing and rewrite preservation only; no
  cryptographic verification — read 🟡.

### litchi-ole-common ([MS-OLEPS]/[MS-OLEDS])

- **[MS-OLEDS] OLE1.0 structures** (§2.2 PresentationObjectHeader/EmbeddedObject/
  LinkedObject): no parsing — read ❌/write ❌.
- **[MS-OLEDS] OLEPresentationStream/TOCENTRY/OLENativeStream** (§2.3.4–2.3.6): no typed
  parsing; preserved as raw streams only — read ❌ (byte-level) / write ❌. `OLEStream`
  (§2.3.3) and `CompObjStream` (§2.3.8) are covered.
- **[MS-OLEPS] PropertyBag property sets** (§2.25.2, ActiveX control persistence):
  generic property-set codec is complete (incl. VT_CF, VT_ARRAY, VT_VERSIONED_STREAM,
  §2.23 bindings), but the PropertyBag control-stream binding shape is not modeled — 🟡.

### litchi-vba ([MS-OVBA]/[MS-VBAL])

- **PROJECT stream typed records** (§2.3.1, all 20 record kinds: ProjectProperties,
  reference lines, CMG/DPB/GC protection state, HostExtenders, ProjectWorkspace): the
  read side exposes only raw decoded text; no record-level model — read 🟡 (raw
  available, no semantics) / write 🟡 (writer synthesizes minimal content + pseudo
  CMG/DPB).
- **dir stream PROJECTREFERENCES** (§2.3.4.2.2: REFERENCE/REFERENCENAME/
  REFERENCECONTROL/REFERENCEORIGINAL/REFERENCEREGISTERED/REFERENCEPROJECT): entirely
  absent from the dir.rs record-ID table — read ❌/write ❌. Type-library references
  cannot round-trip through the project model.
- **PROJECTlk stream LICENSEINFO** (§2.3.2, ActiveX licensing): read ❌/write ❌.
- **Designer Storages and VBFrame streams** (§2.2.10–2.2.11): unparsed; the designer
  module classifies them only as `DocumentClassOrDesigner` — read ❌.
- **SRP streams (§2.2.6, performance cache/p-code) and module-stream p-code prefixes
  (§2.3.4.3)**: deliberately ignored (the writer also refuses `__SRP_` names); the spec
  permits readers to ignore them.
- **[MS-VBAL] the entire language specification** (lexer/grammar/runtime/standard
  library): the crate "never compiles, interprets, or executes"; no lexer/parser —
  entirely unimplemented (design decision; a large gap only if macro-semantic extraction
  becomes a goal).
- PROJECTwm stream: write ✅/read ❌ (unneeded for reading; module names come from dir).

### litchi-opc (ECMA-376 Part 2)

Core areas (parts, content types, relationships, PackURI, core-properties recognition,
digital signatures via `sign.rs` + litchi-sign) are implemented for both read and write.
No ❌-level gaps found; interleaving/streaming-consumption hints are not validated
(minor).

### litchi-drawingml ([MS-ODRAWXML])

The crate covers core DrawingML (color/fill/geom/text/transform), the `c:` chart
namespace, chartStyle/colorStyle (§2.8), diagram, model3d (§2.31), and a16:creationId
(part of §2.25). The remaining extension namespaces largely have no typed
implementation:

- **Ink**: `ink/2010/main` (§2.11), `drawing/2016/ink` (§2.30), `powerpoint/2014/
  inkAction` (§2.21) — no ink module anywhere; read ❌/write ❌.
- **SVG blip**: `drawing/2016/SVG/main` (§2.26, `asvg:svgBlip`) — read ❌/write ❌ (zero
  svgBlip matches repo-wide).
- **Chart extensions**: `drawing/2012/chart` (§2.6), `2014/chart` (§2.22),
  `2014/chart/ac` (§2.23), `2015/06/chart` (§2.44) — ❌; chartex (§2.24) has partial
  models in litchi-pptx/litchi-xlsx but not in the shared crate.
- **Theme extensions** `thememl/2012/main` (§2.4, themeFamily etc.) — ❌.
- **2010/main miscellaneous** (§2.3: cameraTool, compatExt, contentPart,
  hiddenEffects/Fill/Line/Scene3d/Sp3d, imgProps, isCanvas, shadowObscured, useLocalDpi;
  a14:m math) — mostly ❌ (a14:m is implemented inside litchi-pptx
  presentation_properties).
- **Word drawing canvas/group/shape extensions** (§2.13–2.19) and **Excel 2010
  spreadsheetDrawing** (§2.20) — ❌ in the shared crate (host crates may absorb anchor
  portions individually).
- **Miscellaneous small namespaces**: decorative (§2.34), hyperlinkcolor (§2.35),
  animation and animation/model3d (§2.36–2.37), sketchyshapes (§2.38),
  classificationShape (§2.39), oembed (§2.40), scriptlink (§2.41), livefeed (§2.42),
  imageformula (§2.43), signatureLine (§2.2.9.2) — all ❌ (most survive as lazy XML via
  MCE, but have no semantic models).
- legacySpreadsheetColorIndex (§2.3.2.1) and legacy groups/camera tool (§2.2.6,
  §2.2.9.1) — ❌.

### litchi-imgconv ([MS-WMF]/[MS-EMF]/[MS-EMFPLUS])

- **Write direction entirely ❌**: only EMF/WMF/PICT → SVG/PNG/JPEG/WebP conversion
  exists; no metafile authoring API at all.
- Seven EMF records are unconditionally `Unsupported` on read: `EMR_WIDENPATH`,
  `EMR_EXTFLOODFILL`, `EMR_CREATEMONOBRUSH`, `EMR_CREATEDIBPATTERNBRUSHPT`,
  `EMR_DRAWESCAPE`, `EMR_EXTESCAPE`, `EMR_NAMEDESCAPE` — read ❌.
- EMF+ `GetDC` dual-stream interleaving: refused wholesale — read ❌ (with clear
  diagnostics).
- Conditional refusals: glyph-index text, CMYK bitmap decoding, device-dependent
  Bitmap16, target/mode-dependent ROPs, unrepresentable clipping combinations — read ❌
  (triggered by per-record flags).

### Small specifications

- [MS-VBAL]: no counterpart documentation at all (out of design, but should be recorded
  explicitly at the audit level).
- [MS-ODRAWXML] §2.1.4 Ink Content Part and the ink namespace family: neither
  implementation nor documented refusal path.
- [MS-OLEDS] §2.2 (whole OLE1.0 chapter) and §2.3.4–2.3.6 presentation streams: no
  counterpart.
- [MS-OFFCRYPTO] §2.3.6 legacy binary RC4: no shared-layer counterpart.
- [MS-OFFCRYPTO] §2.4 four write-protection methods: no shared-crate counterpart
  (implemented per-format on the DOC/DOCX/XLS side, but ownership is undeclared).
- [MS-LCID]: no dedicated implementation or statement; LCIDs pass through as opaque
  u16/u32 values (e.g. litchi-rtf `LanguageId`) with no spec-table validation —
  acceptable, but undocumented.

### Key partial (🟡) boundaries

litchi-crypto Agile/Standard (legal but non-default encryption profiles are uniformly
`Unsupported`; only Office default products are covered); litchi-crypto IRM (structural
graph and envelope fully validated; content decryption and license processing
deliberately absent); litchi-cfb directory entries (CLSID visible; state bits/timestamps
lost on read and write); litchi-vba PROJECT stream (text-level round-trip; no
record-level semantics for references or protection state); litchi-vba write-out (can
generate complete inert projects — dir/PROJECTwm/module compression — without
references, licenses, designers, or p-code); ole-common vba_signature (blob structure
parsing + rewrite preservation; no signature verification); imgconv EMF+ (complex
brushes/effects/curve approximations are "approximate + diagnostics", not refusals;
convenience entry points refuse with diagnostic results); [MS-OLEPS] PropertyBag
control-persistence shape (generic property sets can read/write the data, but the
control-stream binding form is unmodeled).

Overall: OPC, OLEPS, and CFB are nearly complete; crypto gaps concentrate in
"non-default encryption profiles"; VBA lacks references/designers/the VBAL language
layer; drawingml's main gap is the long tail of [MS-ODRAWXML] extension namespaces
(ink, svgBlip, chart extensions, theme extensions); imgconv has full read coverage with
a documented semantic-unrepresentable refusal set, and the write direction is entirely
absent.

---

## Cross-cutting deliberate exclusions (declared ❌ across formats)

These are part of the project's security model, not oversights: macro/VBA/ActiveX
execution; external-link and data-source refresh; mail-merge execution; field/formula
evaluation; chart/slide rendering; pagination and layout; IRM license evaluation and
protected-content decryption; certificate trust chains and revocation checking; Word
6/95 and earlier formats.

## Priority recommendations (by practical impact)

1. **PPTX Morph and prstTrans transitions** — frequent in modern files.
2. **XLSB Data Model part family** (§2.1.7.35) — the largest single blind spot.
3. **XLSX commentsExtensible reactions** ([MS-OREACTXML]).
4. **ODT in-content RDFa / `text:meta` / XForms model**.
5. **OTH body-structure projection** (tables, annotations, indexes, change tracking).
6. **Crypto non-default encryption profiles** (Standard RC4/AES-192/256, Agile
   SHA-256/512).
7. **drawingml ink / svgBlip extension families**.
8. **Matrix back-filling** (low cost, improves auditability): implemented-but-unlisted
   features in litchi-xlsx (Data Model, Custom Data), litchi-doc (12 parts modules),
   litchi-ppt (photo album, advisor, sound collection, RecolorInfo, custom table
   styles), litchi-ods (flat XML, table-template), litchi-docx (math/bibliography/
   smart_tag/custom_xml/streaming), and the entire missing litchi-rtf matrix.

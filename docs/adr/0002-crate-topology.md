# ADR 0002: Crate topology and dependency direction

- Status: Accepted
- Date: 2026-07-31

## Decision

The target workspace uses small single-responsibility crates and rejects peer
format dependencies in CI. In the diagram, `A -> B` means that `B` may depend on
the more foundational `A`.

```text
litchi-core
├── litchi-detect
├── litchi-word
├── litchi-slide
└── litchi-sheet

litchi-opc -> litchi-ooxml-common -> litchi-drawingml
                                      ├── litchi-docx
                                      ├── litchi-pptx
                                      ├── litchi-xlsx
                                      └── litchi-xlsb

litchi-cfb -> litchi-ole-common
litchi-ole-common -> litchi-doc
litchi-ole-common -> litchi-ppt
litchi-ole-common -> litchi-xls
litchi-odraw -> litchi-doc
litchi-odraw -> litchi-ppt
litchi-odraw -> litchi-xls

litchi-cfb -> litchi-sign -> litchi-opc
litchi-cfb -> litchi-ograph
litchi-odraw -> litchi-imgconv

litchi-codepage
├── litchi-cfb
├── litchi-ole-common
├── litchi-rtf
└── litchi-vba
```

The diagram shows the main direction, not every foundation edge. In particular,
concrete Word, presentation, and spreadsheet crates also depend on their neutral
vocabulary crate. `litchi-drawingml` may depend on `litchi-sheet` for neutral
chart data references; no concrete spreadsheet crate may depend on another.

`litchi-odf-common` owns ODF-neutral package, manifest, namespace, and safe
archive-path vocabulary. `litchi-odf` owns document-family orchestration and
format-specific codecs, while its semantic owners are layered beneath their
contextual module paths. ODF consumers use canonical names such as
`metadata::Metadata` and `media::Image`; the concrete crate does not recreate
common archive-path or namespace logic and does not retain prefix-expanded
compatibility aliases.

ADR 0023 records the target ODF family split: independent `litchi-odt`,
`litchi-ods`, `litchi-odp`, `litchi-odg`, `litchi-odc`, `litchi-odi`,
`litchi-odm`, `litchi-oth`, `litchi-odb`, and `litchi-odf-formula` owners depend on
`litchi-odf-common`, while `litchi-odf` becomes only detection and optional
facade wiring. No family crate depends on the umbrella or on another concrete
family crate.

The IWA subtree follows the same downward-only rule. `litchi-iwa-common` is
the foundational, dependency-neutral layer for bounded varint and protobuf
wire primitives plus neutral table and color vocabulary; `litchi-iwa` and future
`litchi-pages`, `litchi-numbers`, and `litchi-keynote` owners may depend on it.
The common crate must not depend on an archive, graph, facade, or concrete
iWork format crate, and concrete format owners retain their own object-model
and package-topology semantics.

The shared table vocabulary begins at
`litchi-iwa-common::table::cell::BorderSide`. It is a compact, four-variant
cell-edge selector with no stroke, appearance, archive, or protobuf knowledge.
`litchi-iwa` retains `numbers::editor::table::cell::Borders` because that
aggregate contains the facade-owned `ShapeStroke`; native stroke sidecars
convert the neutral selector at the concrete boundary. The old Numbers-owned
selector is removed rather than retained as a compatibility alias, and the
Numbers, Pages, and Keynote border APIs now take this canonical selector
directly.

The physical IWA substrate is layered beneath the application crate:
`litchi-iwa-protos` owns the generated raw schemas, and `litchi-iwa-core`
depends on it for bounded archive framing and checksum-free Snappy encoding.
`litchi-iwa` consumes the core's typed, slice-based codecs directly; its former
633-line duplicate Snappy implementation and 172-line varint kernel are gone.
The core layer does not open packages, resolve application message IDs, or own
document topology, while the facade retains those application-level
responsibilities. The common wire crate is also the sole owner of parsed
`WireField` values and bounded scalar/repeated mutation; the facade's private
`wire.rs` is only a temporary callback/error adapter and does not copy parsed
fields or maintain a second wire representation.

The first extracted semantic value layer is `litchi-iwa-text`, which owns only
the allocation-bearing rich-text values shared by the format leaves. It has no
archive, protobuf, or application dependency. `litchi-pages` owns the concise
`Section`/`SectionType` vocabulary, and `litchi-keynote` owns `Slide`, `Show`,
build, and transition values; both depend downward on `litchi-iwa-text` only.
The shared text leaf now also owns the strict `font::{Font, Name}` vocabulary
and its typed `NameError`; the IWA facade keeps only a thin error conversion and
native archive adapters. `Name` stores one boxed UTF-8 identifier, validates
before allocating borrowed input, and consumes owned `String` input directly.
The leaf therefore remains archive-free while Pages, Numbers, and Keynote use
one canonical font model instead of maintaining format-local copies.
The common color leaf now owns `color::{RgbColorSpace, Rgba}` and its typed
`color::Error`; native protobuf conversion remains in the IWA shape adapter.
`Rgba` is a fixed-size, copyable value that validates all four finite channels
before construction, so format owners do not allocate or import archive error
state merely to exchange a color.
The leaf's `transition::Effect` owns the lossless native transition-effect
identifier vocabulary, including canonical known variants and lossless unknown
identifiers; IWA retains transition archive decoding, wire patching, and
transactional validation at the format boundary.
The existing `litchi-iwa` package reader temporarily consumes these leaf values
through private migration adapters. The direct edges are present in the
canonical boundary graph because the adapters are already dependency-safe;
their removal is a staged ownership exit, not a public compatibility layer.
The Numbers migration continues with table, formula, and sheet ownership, with
no peer dependency between the three concrete crates.

The Numbers migration now begins with dependency-free `litchi-numbers::cell`,
whose concise `Value`, `Type`, and `Update` vocabulary is shared by the
Numbers reader and the structured facade through a private adapter. The first
table/sheet semantic slice now also lives in `litchi-numbers`: `table` owns
compact checked coordinates and dimensions, half-open ranges, sparse cells,
budgeted grid views, and the fallible builder-to-immutable-table transition;
`sheet` owns the immutable table collection and duplicate-name validation.
Neither module depends on archives, protobufs, comments, or application
topology. `NumbersDocument::semantic_sheets` now provides the consuming IWA
reader seam into the immutable `litchi_numbers::Sheet` model; it transfers
finished sparse tables without rebuilding cell maps and intentionally leaves
comments/native sidecars on the archive adapter. The dependency-free formula
vocabulary now follows the same boundary: `litchi-numbers::formula` owns
formula caches, references, operators, and expression construction, while
`litchi-iwa` retains protobuf compilation and calculation-engine mutation.
The former `litchi-iwa::numbers::formula` module is crate-private; the facade's
root re-exports are deliberate ergonomic aliases, not a compatibility layer.
The shared formula types retain their `Formula*` prefix as a cross-format
vocabulary exception so Pages, Keynote, and Numbers call sites remain
unambiguous when the types are imported without a module qualifier. Their
constructors are allocation-conscious, while archive-boundary compilation
enforces bounded depth, node count, function arguments, and precedents.
Package owners continue the same downward-only extraction pattern.

The Pages and Keynote table readers now consume the same leaf `Table` through
an ownership-preserving adapter seam. Their public table facades borrow the
canonical sparse cells directly while retaining format-owned comments and
merge regions as separate sidecars; read-only comment sidecars are compact
sorted boxed pairs, and the former tuple-keyed cell maps are no longer rebuilt
in either reader. The generic structured extractor remains the last current
`NumbersTable::into_parts` consumer and is staged separately.

The first Numbers wire seam is now `litchi-numbers::cell::wire`. It owns the
dependency-free, byte-preserving BNC codec, stored-value and cached-scalar
views, data-format identifiers, and decimal128 codec; it preserves unknown
trailing bytes for round trips. `litchi-iwa` retains archive traversal,
protobuf integration, and package mutation, exposing the wire module only
through a private migration adapter and converting its local error at that
boundary. This is an ownership move, not a compatibility surface. The IWA
reader now uses a mutable archive-boundary adapter around the leaf table
builder while it carries format-owned comments and converts native archive
values. It also
retains the finite ingress profile: table rows, columns, addressable cells,
and materialized sparse cells are bounded; tile keys and local/global
coordinates are checked against those dimensions; and a tile reference must
resolve to exactly one typed `6002` payload. A native `6000` TableInfoArchive is
metadata only; cell extraction consumes the typed `6001` TableModelArchive.
Sparse offset ranges are decoded into one fallibly reserved vector, with
count, slot, storage, and monotonicity checks performed before allocation.
These limits belong temporarily to the adapter and are not a dense-grid
compatibility promise.

`litchi-drawingml::chart` owns the host-neutral classic-chart model and bounded
XML codec. Its contextual modules are `model`, `data`, `axis`, `series`,
`plot_area`, `reader`, and `writer`; the public codec verbs are the short
`reader::read` and `writer::write`. `writer::write_with_rels` is the focused
low-level seam for relationship identifiers allocated by a concrete package.
`litchi-drawingml::diagram` likewise owns the SmartArt semantic tree plus the
data, definition, and generated-part grammar. DOCX, PPTX, XLSX, and XLSB retain
only host anchoring, relationship allocation, and concrete package topology.
Neither shared module depends on a concrete format or the OOXML migration host,
and malformed input returns the crate-local `Error` rather than a host error.

`litchi-ooxml-common::custom` owns the package-level custom-document-property
grammar and graph service shared by every OOXML host. Its complete facade is
`custom::{Props, Value}`: fallible `insert`, case-insensitive `get`, `contains`,
and `remove`, plus `names`, `iter`, `clear`, `read`, and `write`. Property names
use canonical Unicode caseless identity while retaining producer spelling.
`Value::{Empty, Text, I32, I64, F32, F64, Bool, Time}` preserves the supported
wire type without an external type-code enum. The Office producer profile,
PID and format-ID rules, RFC3339 `vt:filetime` lexical form, namespace and
cardinality checks, and bounded resource budgets are enforced at this owner.
Missing relationships mean absence; ambiguous, external, orphaned, malformed,
or wrong-content-type graphs are errors. Empty writes remove both part and
relationship, and only actual mutations invalidate signatures.

`litchi-ooxml-common::custom_xml` similarly owns inert Custom XML Data Storage
grammar and topology. Its contextual vocabulary is `Conformance`, `Props`,
`Item`, `NewProps`, and `NewItem`; its verbs are `read_props`, `write_props`,
`discover`, `add`, and focused validation helpers. `NewProps` groups the
properties part, relationship, and value so a partially-specified properties
request is unrepresentable. Loaded `Item` state is read-only behind short
accessors. Payloads share the OPC part's immutable allocation and `xml()` lends
a slice, preventing relationship multiplicity from copying large XML parts.
Creation consumes owned bytes, validates every fallible graph and XML step
before mutation, rolls back defensive failures, and invalidates signatures only
after commit. Neither service resolves schemas, executes XPath, or depends on a
concrete document format. The migration host contains no compatibility module
or alias for either former owner.

`litchi-ooxml-common::embedded` owns inert discovery of normative Embedded
Object and Embedded Package relationship occurrences. Its complete vocabulary
is `embedded::{Entry, Kind, Limits, Payload, Target}` and its verbs are `scan`
and `scan_with`. Entries lend their source, relationship ID, target metadata,
and payload bytes from the OPC package; discovery never copies, sniffs, opens,
activates, or recursively parses an embedded payload. Safe defaults bound both
the occurrence inventory and aggregate relationships on uniquely validated
payload parts. Duplicate references reuse that validation, strict and
transitional relationship families are accepted, and output order is stable.
The source policy includes the ISO OOXML host parts, every Word main-part
content-type variant, and the additional binary SpreadsheetML sources defined
by `[MS-XLSB]` sections 2.1.7.36 and 2.1.7.37. External targets remain inert and
are never fetched. DOCX, PPTX, XLSX, and XLSB expose the same short `embedded`
facade while retaining responsibility for host anchors and mutations; the
migration host owns no duplicate module or type alias.

`litchi-word`, `litchi-slide`, and `litchi-sheet` depend only on `litchi-core`.
They contain selectors, queries, events, detached builders, and semantic values,
not container parsing or concrete document handles. Concrete imported objects
remain canonical in their format crate.

`litchi-odraw` owns only the OfficeArt record grammar, property tables, shape
containers, bounded traversal, and deterministic record writing defined by
`[MS-ODRAW]`. The `OfficeArtClientData` and `OfficeArtClientTextbox` payloads
are explicitly host-application records in `[MS-ODRAW]` section 2.2.14, so DOC,
PPT, and XLS decode those payloads in their concrete crates. Shared shapes
expose the borrowed host payload records without interpreting them. Canonical
types use their module context (`record::Record`, `prop::Props`,
`shape::Shape`) instead of repeating an `Escher` or `OfficeArt` prefix.

`litchi-ole-common::object` owns bounded, inert discovery of DOC/XLS object
storage topology and transactional CFB stream/storage rewrites. It exposes
contextual names such as `object::{Object, Objects, Editor, Limits}`. Semantic
lookup (`Objects::get`) is the primary selector, while checked discovery-order
lookup (`Objects::at`) remains available; neither selector panics. Concrete
host metadata is not modeled in the common crate. Common objects retain those
bytes opaquely, and the owning format crate provides the typed interpretation,
such as `doc::embedded_object::Info` for `[MS-DOC]` `ObjInfo` flags.

Additional focused crates are permitted where the responsibility is real:

- `litchi-codepage` owns exact legacy code-page selection plus bounded text
  encoding and decoding. Its short contextual vocabulary is `Page`, `Mbcs`,
  `Ansi`, and `Error`; all three capabilities occupy one byte. `Mbcs` excludes
  UTF-16 from byte-terminated record paths, while `Ansi` admits only the exact
  `[MS-OSHARED]` ANSI set. Checked construction rejects unsupported identifiers
  instead of silently substituting a superficially similar encoding. Strict
  decoding is the default, decoding recovery is explicitly named, and concrete
  formats retain responsibility for terminators and other record-level text
  rules. Generic hexadecimal decoding remains `litchi-core::hex` and does not
  pull a legacy text codec into the neutral vocabulary crate.
- `litchi-math` replaces the current equation-focused `litchi-formula` name.
- `litchi-calc` owns spreadsheet formula parsing, dependency graphs, and pure
  calculation; it has no network or async-runtime dependency.
- `litchi-crypto`, `litchi-sign`, and `litchi-vba` own shared inert security
  capabilities rather than creating OPC/OLE cross-dependencies.
- `litchi-ograph` owns the neutral `[MS-OGRAPH]` chart model, record grammar,
  and standalone compound-package codec. XLS owns workbook tab/Obj integration
  and PPT owns presentation frames and embedded-object integration; PPT never
  depends on the concrete XLS crate.
- Runtime adapters such as `litchi-tokio` are separate optional crates.

`litchi-sign` owns the bounded, trust-neutral signature engine rather than a
format facade. Its root vocabulary is compact (`Signer`, `Policy`, `Coverage`,
`Report`, `Status`, and `Trust`), `xml` owns XMLDSig canonicalization and
verification, and `cfb` owns the compound-file storage adapter. `litchi-opc`
depends downward on that engine and owns only OPC graph selection, relationship
and content-type maintenance, and package-level transaction staging. This
direction prevents a signing/OPC cycle and lets DOC, PPT, and XLS use the same
neutral engine without depending on OOXML. Strict policy accepts only complete
package coverage; compatibility policy may report a typed partial coverage for
real producer signatures that intentionally select a subset. Neither policy
turns partial coverage into an unqualified success.

`litchi-odraw::image` owns the OfficeArt BLIP, FBSE/BStore, delayed-storage,
digest, and bounded writer grammar. Image decoding and conversion remain in
`litchi-imgconv`, which consumes the grammar instead of redefining it. Host
crates retain their native topology: in particular, PPT resolves a picture ID
through the drawing-group FBSE table to a delayed Pictures-stream BLIP instead
of treating that headerless stream as a second BStore.

The optional umbrella image facade depends on both layers directly. It exposes
the grammar as `images::art` and codecs as `images::codec`, so the codec crate
does not become a compatibility tunnel for types it does not own. File helpers
use short contextual names (`images::doc`, `images::ppt`, `images::escher`, and
`images::store`) and return borrowed views whenever the input lifetime permits.

`litchi-ograph` owns only neutral chart records, bounded chart-substream
discovery, borrowed `chart::Ref` and move-owned `chart::Stream`/`chart::Book`
capabilities, the semantic `chart::Chart` model, deterministic record encoding,
and the standalone compound-package codec. Strict `[MS-OGRAPH]` packages have a
globals-plus-one-Graph-chart Workbook; the separate host-neutral scanner also
accepts Excel chart BOFs nested in arbitrary Workbook streams without claiming
that the surrounding workbook is a standalone Graph package. Context-specific
`chart::Link` variants keep Graph's fixed datasheet coordinate and Excel's
variable parsed formula from being conflated.

An untouched parsed semantic chart consumes back into its exact source
allocation. Mutation of a parsed chart is refused until every opaque record and
reserved byte has a proven placement. Fresh semantic authoring is likewise
refused until the complete mandatory chart-sheet, format, series, axis-parent,
and cache grammar is modeled; a partial self-roundtrip is not treated as an
Office-compatible artifact. XLS owns workbook tabs, BIFF objects, and chart-host
mutation; PPT owns frames and embedded-object integration. The neutral crate
does not depend on either host, expose a runtime lock wrapper, or imply
rendering, formula evaluation, activation, or current fresh-authoring support.

`litchi-crypto` owns bounded `[MS-OFFCRYPTO]` structures and transformations,
including compound-file DataSpaces metadata and password-derived cipher
contexts. It may depend downward on `litchi-cfb` and `litchi-ole-common`, but
not on either migration host or any concrete document format. Its namespaces
provide short typed names such as `rc4::{Flags, Header, Context, Error}`;
format crates remain responsible for locating native records and mapping
crypto failures into their own error vocabulary. Secret-bearing contexts keep
their material private and zeroizing, and the crate has no async-runtime edge.
`ooxml::{Kind, Mode, Limits, Password, Opened, Error}` owns the supported
Standard and Agile encrypted-package profiles plus the
StrongEncryptionDataSpace CFB adapter. Password-free `inspect`, move-consuming
`open`/`encrypt`/`rekey`, and runtime-neutral `load` are the complete envelope
service; ordinary unencrypted input returns the same allocation, while explicit
`_with` variants apply caller-selected resource ceilings. The migration host
re-exports this vocabulary under its contextual `encryption` module. DOCX,
PPTX, and XLSX retain the detected mode, refuse an implicit plaintext save,
name plaintext output explicitly, and atomically replace path destinations.
They depend only on this service and never import a CFB parser, cipher
primitive, or encrypted-container implementation directly.

`litchi-vba` owns the inert, bounded `[MS-OVBA]` codec and project model. It
depends downward only on `litchi-cfb` and `litchi-codepage`; it does not own
DOC, PPT, XLS, OPC, or OOXML package integration and never compiles,
interprets, or executes source. Its contextual namespaces keep the public
vocabulary short:
`codec::{encode, decode}`, `dir::{Dir, Module, Kind}`,
`project::{Project, Module, Text}`, and
`build::{Project, Module, Id, Platform, Kind}`. A serialized `Payload` is a
validated, move-first capability rather than an arbitrary byte alias. Callers
can obtain one only by validating an existing compound payload or by consuming
a checked builder; host packages consume it directly instead of accepting an
untyped `Vec<u8>`. This preserves a concise high-level boundary without hiding
the lower-level directory and compression codecs needed by focused tooling.
The crate has no async-runtime edge, public lock wrapper, compatibility facade,
or public type carrying a redundant `Vba` prefix.

`litchi-docx::font` owns the WordprocessingML font-table model, bounded
Strict/Transitional XML codec, and font-part relationship graph. Its public
vocabulary is contextual (`Table`, `Font`, `Conformance`, `Family`, `Pitch`,
`Charset`, `Signature`, `Embed`, `Style`, `Resource`, and `License`, with
extension markup isolated as `font::raw::Attr`) rather than repeating `Docx`,
`Wordprocessing`, `FontTable`, or `EmbeddedFont` in every name. Package writes
consume the table or owned payload being installed;
reads lend or share package-owned bytes. The capability validates names,
licensing flags, resource ceilings, relationship topology, and orphan removal.
One normalized Unicode-caseless identity is used by lookup and every CRUD
operation, so spelling normalization cannot make selectors disagree. The
package host exposes symmetric `fonts`, `put_fonts`, and `remove_fonts` entry
points, but never discovers, loads, renders, or executes a font program.

`litchi-docx::numbering` owns the package-neutral numbering collection,
definitions, instances, levels, overrides, picture bullets, closed numbering
domains, and bounded WordprocessingML codec. Its contextual facade uses names
such as `Collection`, `Definition`, `Instance`, `Level`, `Format`, `Restart`,
and `Suffix`; it does not carry a redundant `Docx` or `Numbering` prefix. The
OOXML migration host only resolves the numbering relationship, preprocesses
the part with markup compatibility, maps errors, and returns the owner
collection. It does not define a second model or retain prefix-expanded
compatibility aliases.

`litchi-docx::alt` owns WordprocessingML alternative-format anchors and opaque
payload typing. Its short vocabulary is `Chunk`, `Conformance`, `Data`,
`Import`, `Kind`, `Part`, and `Target`; cheap low-level identifiers are checked
`Rel` and `Uri` values. `Data` and `Import` are deliberately move-only, package
insertion transfers their payload allocation into OPC storage, and borrowed
`Part` access never parses or copies foreign bytes. Checked-in `[MS-OI29500]`
section 2.1.527 and `[MS-OE376]` section 2.1.558 define the ten supported Word
media families and case-sensitive Transitional `aFChunk` relationship. The
host exposes ordered `add_alt`, `insert_alt`, `replace_alt`, `remove_alt`, and
`move_alt`; public writer CRUD does not accept raw relationship IDs. External
targets remain inert. Markup-compatibility selection retains original source
coordinates, so read and mutable selectors agree on the active Choice/Fallback
branch; full-document parsing also preserves inherited Strict and Transitional
namespace aliases. Payload, XML, nesting, and anchor limits are enforced before
unbounded package or parser work.

`litchi-docx::web` owns the bounded WordprocessingML web-settings grammar,
recursive frameset/division model, deterministic producer bytes, and optional
OPC graph. Its contextual vocabulary is `Settings`, `Conformance`, `Key`,
`Id`, `Twips`, `Div`, `Borders`, `Frameset`, and `Frame`; the shared theme-color
vocabulary is `litchi-docx::color::Theme`. Nonzero producer-visible numeric
division IDs are the primary selector and checked source positions are the
repair fallback. `Div` carries all four schema-required margins as typed signed
twips, so ordinary construction cannot omit them. Package
`load`, consuming `put`, and `remove` validate dialect, ownership, frame edges,
content type, and resource bounds before commit. Exact and semantic no-ops
retain source bytes and signatures. The migration host exposes only `web`,
`put_web`, and `remove_web` while the wider DOCX package remains there.
Schema-valid `OnOff` lexical forms remain readable, but division-role markers
write explicit numeric values because the native Word gate rejects empty true
`bodyDiv` and `blockQuote` elements.

`litchi-docx::glossary` owns the bounded WordprocessingML glossary-document
grammar, semantic building-block catalog, and auxiliary OPC graph. Its ordinary
vocabulary is contextual (`Catalog`, `Entry`, `Props`, `Name`, `Category`,
`Gallery`, `Id`, `Kind`, `Insert`, and `Conformance`); physical `Graph`, `Part`,
and `Rel` values are isolated under `glossary::raw`. Canonical Unicode-caseless
names are the primary selectors and checked source positions are the repair
fallback for lookup, replacement, rename, removal, and reorder. A private name
index plus checked per-entry and catalog size totals keep repeated CRUD
proportional to the selected entry rather than the entire catalog. Fresh entries
require the properties and name needed by Word 2007,
while the reader retains valid empty or less-constrained producer catalogs.
Entry payloads move across semantic mutations, while low-level graph publication
borrows its recovery copy. Package `load`, consuming semantic `put`, and `remove`
validate dialect, role-derived relationship permissions and target modes, every
target, content type, graph-wide bounds, reserved part names, and package-wide
inbound ownership before publication. Internal hyperlinks remain references
rather than owned dependencies. Producer duplicate names remain readable and
make semantic lookup ambiguous; new conflicts are rejected. Unchanged bound
catalogs and canonical/exact raw no-ops retain producer paths, bytes, and
signatures. Relationship-bearing semantic catalogs are privately bound to their
validated physical resources, and every referencing entry/background carries
per-value lineage, preventing cross-package `r:id` rebinding; a real update
stages all fallible work before commit. Untouched direct producer entries
retain bounded serialized inactive/ignorable MCE content and its relationship
references across unrelated CRUD. Shared namespace scopes avoid per-descendant
copies, while aggregate projection/snapshot and DOM-allocation budgets prevent
cross-entry amplification. Fresh semantic authoring allocates canonical-first
free names for glossary-local styles, settings, font-table, and web-settings
resources; `Package::new_template()` selects the DOTX container used by native
AutoText. The migration host exposes short document/package adapters plus the
canonical owner module as a contextual re-export, and owns no duplicate glossary
model or legacy type alias.

`litchi-pptx::transition` owns the PresentationML transition model and bounded
XML codec. Each `Kind` variant carries only the direction/orientation value
valid for that effect, so invalid effect-option pairs are not representable.
Checked duration, delay, and wheel-spoke values reject invalid input before
serialization. Unknown source effects and extension children are retained as
bounded inert markup. A semantic sound or effect variant is exposed only when
both read and write preserve it; the API does not keep constructor-only or
writer-rejected compatibility variants.

`litchi-pptx::shape` owns the canonical semantic index over PresentationML
shape trees. `Scene` builds one bounded, namespace-aware owner index and
exposes a non-exhaustive data-bearing `Shape` enum with contextual variants
such as `Auto`, `Picture`, `Table`, `Chart`, `Diagram`, `Ole`, `Group`, and
`Connector`; callers never compare a separate native type discriminator.
Scenes preserve depth-first source order while `Group::shapes` exposes direct
children, so nested groups remain both searchable and hierarchical. Exact
producer-visible names are the primary selector and checked pre-order positions
are the repair/import selector. Ordinary lookup represents a missing name as
`None`; strict lookup, ambiguous names, and out-of-range positions have typed
errors, and neither path uses indexing panics. MCE-free owners stay borrowed. When
Choice/Fallback processing is required, the scene owns one bounded processed
owner buffer, and every shape XML view remains a checked span into that shared
owner rather than a copied subtree. The concrete PPTX crate retains shape
classification and host semantics; `litchi-drawingml` remains responsible only
for host-neutral DrawingML vocabularies.

`litchi-pptx::tag` owns the bounded PresentationML programmable-tag grammar,
low-level relationship inventory, and anchor-aware package mutation. Its
contextual vocabulary is `List`, `Tag`, `Key`, `Source`, `Conformance`, and
`tag::raw::Attr`. Semantic name lookup is the primary selector inside a list,
while checked numeric positions support source-order repair without exposing
relationship IDs or part names through the ordinary facade. Litchi chooses one
deterministic NFD/default-case-fold/NFD identity for lookup and every detached
add, insert, replace, set, remove, and reorder operation. Direct presentation
and common-slide-data anchors use singleton `load`, `put`, and `remove`; the
migration facade exposes short slide-scoped `tags`, `put_tags`, and
`remove_tags` operations selected first by producer-visible slide name and
second by checked position; an already-resolved `Slide` reads its attachment
directly without rescanning unrelated slide parts. Direct-owner reads and
mutations select the same active MCE branch, then map the semantic insertion,
container, and anchor back to checked raw-source coordinates; inactive branches
never become mutation targets, while every preserved raw anchor participates in
shared-edge retention. Shape-owned lists remain distinct objects and are never
flattened into the slide result.
`tag::shape::{load, put, remove}` is the focused package layer for those
anchors and reuses canonical `shape::Key`: exact producer-visible names remain
the ordinary selector and checked depth-first positions remain available for
repair. The editor resolves five schema shape families plus nested groups,
maps the semantic selection back to the active raw-source MCE branch, and never
requires a relationship ID in the public selector. The migration facade adds
short package `shape_tags`, `put_shape_tags`, and `remove_shape_tags` methods,
selecting the slide first by producer-visible name or checked position and the
shape by the same semantic key. An already-resolved `Slide::shape_tags` reads
without a presentation-wide rescan. Checked-in
`[MS-OE376]` section 2.1.1170(c) requires case-insensitive uniqueness but does
not prescribe this normalization algorithm.
Malformed producer duplicates remain inspectable by numeric position and make
semantic selection explicitly ambiguous. Values and retained extension markup
stay inert. Private escaped-wire counters make aggregate size preflight O(1)
after scanning only the incoming value, so every successful checked mutation
remains serializable under the 8 MiB part ceiling. Strict and Transitional
relationship discovery and anchored mutation reject external,
wrong-content-type, duplicate-target, and relationship-bearing tag parts.
Unanchored relationships remain visible only through the explicitly low-level
inventory. Candidate operations complete bounded validation before commit,
change the XML anchor, relationship, and target part as one transaction,
preserve byte-identical signed no-ops, fork shared targets on replacement, and
remove a target only after a package-wide inbound-edge scan proves it orphaned.
A dirty legacy presentation writer is rejected because a later materialization
could overwrite the edited slide markup and relationships.

`litchi-pptx::notes` owns the bounded PresentationML speaker-notes graph,
Strict/Transitional XML validation, plain-text notes producer, deterministic
notes-master asset, and transactional OPC mutation. Its contextual vocabulary
is `Conformance`, `Theme`, `Master`, `Slide`, and `Graph`; physical relationship
and part identities remain private. `load` returns a lifetime-free editable
graph and copies each validated payload once, while focused `slide` copies only
the selected notes payload and metadata-only deletion copies none. Consuming
`put` validates and stages every graph and relationship change before commit,
then moves the owned XML buffers into OPC parts. Exact no-ops preserve
signatures. The migration host retains only semantic slide selection and dirty-
writer guards around `notes`, `put_notes`, `remove_notes`, and `clear_notes`;
the former host owner and forwarding aliases are deleted.

`litchi-pptx::table::style` owns the bounded DrawingML table-style catalog,
deterministic producer bytes, and optional presentation graph. Its concise
vocabulary is `Conformance`, allocation-free `Id`, compact `Parts`, `Def`, and
`List`. Stable GUID identity is the primary selector, `at` is the checked raw-
order fallback, and `named` returns every match because display names may be
empty or duplicated. Definitions borrow checked ranges from one list-owned XML
allocation; unchanged stores move that allocation back to OPC, while rename
preserves opaque formatting content. Package `load`, consuming `put`, and
`remove` validate all six main-document profiles, graph ownership, dialect,
content type, schema order, and resource ceilings before mutation. The
migration host exposes only `styles`, `put_styles`, and `remove_styles`.

`litchi-pptx::font` owns the bounded PresentationML embedded-font grammar,
typed semantic values, and optional package graph. Its concise vocabulary is
`Fonts`, `Font`, `Face`, `Data`, `Style`, `Format`, `PitchFamily`, and `Key`;
physical relationship IDs, part names, and content-type strings remain private.
Typeface-first Unicode-caseless lookup is backed by one cached library-defined
identity, while checked positions remain available for malformed-producer
repair. Font programs use shared immutable allocations, and aggregate limits
count unique resources rather than face references. Package `load`, consuming
`put`, and `remove` validate both conformance families and all six main-part
profiles, preserve exact signed no-ops, and publish real changes atomically.
PowerPoint-compatible authoring validates an Embedded OpenType container for
`application/x-fontdata`; the standards-only raw `x-font-ttf` profile is
explicit, and Word-only obfuscation is absent. The migration host exposes `fonts`,
`put_fonts`, and `remove_fonts` and owns no duplicate embedded-font model.
Shared automatic discovery keys its roaring-backed, scalar-only `Glyphs` by a
typed family-and-face `Request`; concrete adapters map that neutral four-style
enum into their own font owner. `litchi-opc::FontEmbedding` owns the closed
None/Full/Subset save policy. DOCX alone owns the typed 16-byte `FontKey` and
its XML-boundary lexical codec.

`litchi-xlsx::chain` owns SpreadsheetML calculation-chain grammar, its typed
ordered model, and the single-part workbook relationship service. Short types
`Sheet`, `Step`, `Flags`, `Cell`, and `Chain` encode the native sheet-ID range,
mutually exclusive dependency roles, packed orthogonal markers, checked grid
addresses, and nonempty ordering. Semantic sheet/address CRUD is primary;
checked numeric order remains available for repair, and malformed duplicate
keys are inspectable but make semantic selection ambiguous. `load`, `put`, and
`remove` validate the complete Strict or Transitional OPC graph, preserve
bounded extension markup, retain signatures on exact no-ops, and never evaluate
formulas. The migration host caches this canonical model only until the XLSX
package owner itself moves out of the monolith.

`litchi-xlsb::raw` owns the BIFF12 record wire kernel: `Kind`, `Header`,
borrowed `Record`/`Records`, bounded `Cursor`, and `Writer`, with constants
under `raw::kind`. Following `[MS-XLSB]` section 2.1.4, record kinds use exactly
one or two bytes and remain below 16,384, while record lengths use at most four
bytes. Clean end-of-stream is distinct from a truncated header or payload,
payload and string budgets are explicit, and strict UTF-16 decoding is
separate from byte preservation. `Header` and borrowed `Record` keep their
validated fields private and expose short accessors. Following `[MS-XLSB]`
section 2.5.123, RK reads preserve the signed 30-bit/floating and divide-by-100
flags; RK writes refuse values that cannot be represented bit-exactly instead
of silently rounding them. The kernel has no OPC, DrawingML, XLSX, runtime, or
concrete peer dependency; XLSB semantic records remain in the concrete owner
and migrate onto this substrate incrementally.

`litchi-xlsb::calc` owns the canonical 26-byte `BrtCalcProp` semantic record
and streams it through the canonical raw `Cursor` and `Writer`. Reads also
accept the exact 25-byte form emitted by an early Microsoft Excel 12 producer,
zero-extending its one-byte option tail without allocating or copying; writes
always emit the canonical 26-byte form. Every other length remains a typed
error. Its short public vocabulary is `Props`, `Mode`, `Opts`, `Delta`, and
`Threads`. Private fields, checked setters, and consuming `with_*` builders
make every `Props` value directly writable. `Opts` packs the nine switches into
one `u16`; unknown bits are rejected. Checked-in `[MS-XLSB]` section 2.4.318
fixes the mode enumeration, reserved bits, and `1..=1024` thread-count domain,
while section 2.5.172 makes NaN, infinity, subnormal values, and negative zero
invalid `Delta` states. The migration host exposes concise `calc`, `calc_mut`,
and move-accepting `put_calc` entry points instead of retaining the former long
compatibility types.

`litchi-eval` remains runtime-neutral when `web_functions` is enabled. External
retrieval is an explicit caller capability: `FormulaEvaluator::with_fetch`
borrows an implementation of `Fetch`, whose boxed future can be driven by any
executor. With no provider, evaluation performs no network I/O and
`WEBSERVICE` returns a connection cell error; supplied responses are bounded,
strictly decoded as UTF-8, and checked against the cell text limit. The
evaluator's method-scoped `At` context carries the current cell while borrowing
both the evaluator and a private circular-reference session. Concurrent
top-level calls therefore cannot mistake one another for a cycle, and RAII
removes a visit marker on every exit. No runtime lock wrapper enters the public
API. Tokio remains test-only, and neither Tokio nor Reqwest is a normal
dependency of the crate.

The `litchi-ole` monolith is removed after DOC, PPT, and XLS migrate into their
concrete crates. It does not remain as a compatibility crate, feature, or
module. The current `litchi-ooxml` monolith is likewise removed after its
contents migrate. The umbrella `litchi` contains no format implementation logic
and re-exports canonical types without creating aliases with redundant
prefixes. Legacy Word, PowerPoint, and Excel are independently gated as `doc`,
`ppt`, and `xls`, with concise low-level facades at `litchi::{doc,ppt,xls}`.

## Enforcement

- A checked-in dependency allowlist rejects concrete peer edges, including dev
  and optional dependencies.
- The allowlist inventories every direct `crates/*/Cargo.toml` workspace member.
  Every internal edge is either a canonical downward ceiling or an ordered,
  stale-checked migration-debt entry with a reason and exit condition. Migration
  hosts have no canonical edges: adding an unclassified edge fails, and removing
  a debt edge also fails until its ledger entry is deleted.
- `litchi-core` owns only format-neutral sources, blobs, budgets, execution,
  scalars, selectors, diagnostics, patch envelopes, and content events. It owns
  no ZIP, XML, CFB, format feature, Tokio, Reqwest, or Rayon dependency.
- Runtime-neutral policy evaluates normal Cargo dependency edges, including
  optional normal edges. Development-only runtimes may support tests without
  weakening or masking the production dependency check.
- Container/common crates do not depend on concrete formats.
- Default `litchi` enables DOCX, PPTX, and XLSX. XLSB, legacy formats, crypto,
  signing, VBA parsing, calculation, rendering, and runtime adapters are opt-in.
  Enabling a feature adds capability and never changes existing semantics.

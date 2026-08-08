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
The core also depends downward on `litchi-iwa-common` for the sole shared
source-bound wire-tree preflight. That edge does not move archive or Snappy
ownership into common: common validates raw spans under aggregate byte, field,
and nesting budgets, while core supplies the archive-header schema policy and
projects the result into its physical metadata values.
`litchi-iwa` consumes the core's typed, slice-based codecs directly; its former
633-line duplicate Snappy implementation and 172-line varint kernel are gone.
The raw schema build deliberately omits prost's runtime type-name metadata:
the workspace has no type-name consumer, so generated `Name` implementations
would add code and static strings without improving archive decoding. Schema
identity remains explicit at the IWA application boundary where diagnostics
need it.
The core layer does not open packages, resolve application message IDs, or own
document topology, while the facade retains those application-level
responsibilities. The common wire crate is also the sole owner of parsed wire
representations and bounded scalar/repeated mutation. New strict readers use
the source-bound `WireView<'a>`/`WireFieldView<'a>` pair, which retains one
borrowed source plus compact private spans instead of per-field byte owners;
the older `WireField` mutation representation remains only while its callers
are migrated. The facade's private `wire.rs` is a callback/error adapter and
does not copy parsed fields or maintain a second wire representation.

Archive-header protobuf interpretation is the first production Buffa seam.
`litchi-iwa-protos` keeps Buffa 0.9.1 generated eager/lazy types private behind
an internal codec and continues to expose the existing generated compatibility
types during migration. `litchi-iwa-core` performs schema-directed wire-tree
preflight before invoking that codec; neither Buffa views nor protobuf values
enter Pages, Numbers, or Keynote semantic APIs. The original encoded header
remains core-owned preservation state, so a lazy semantic projection never
becomes the byte-authoritative save representation.

The archive-neutral package-entry substrate now lives in
`litchi-iwa-package`. It owns only ordered, uniquely named entry storage, its
fallible name index, and source-checked reversible entry patches; it has no
ZIP, Snappy, protobuf, graph, or application dependency. `EntryStore` and
patch clones are copy-on-write handles, so applying a patch shares payload
allocations rather than materializing a second package. `litchi-iwa` depends
downward on this leaf while retaining `IWorkPackage`'s ZIP ingress, IWA
decoding, resource policy, and transactional snapshot validation. This staged
boundary lets the eventual Pages, Numbers, and Keynote package owners share
entry storage and raw package transactions without making the package leaf
depend on a concrete format or leaking application message IDs upward.

Pages package-root and body-storage protobuf interpretation remains in the IWA
format adapter: `litchi-pages` is now archive-free and owns only its bounded
semantic document, body, and section values. The IWA facade retains ZIP
ingress, object lookup, native protobuf decoding, and generic text fallback;
semantic construction failures cross that boundary through the typed
`litchi_pages::Error`. Keynote presentation settings follow the same boundary
at `litchi-keynote::show::{Mode, Settings}`: the semantic crate owns validated
dimensions and playback values, while `litchi-iwa` retains only the native
`KN.ShowArchive` wire adapter and transactional publication.

Physical organization follows the same ownership rule inside format crates:
the Numbers text-box API is isolated in the private
`numbers::editor::text_box_api` module, leaving the editor root focused on
shared orchestration and the remaining migration seams. This is an internal
layout step toward the independent `litchi-numbers` crate, not a new facade
layer or a compatibility surface.

The first extracted semantic value layer is `litchi-iwa-text`, which owns only
the allocation-bearing rich-text values shared by the format leaves. It has no
archive, protobuf, or application dependency. `litchi-pages` owns the concise
`section::{Section, SectionType}` vocabulary and the archive-free
`document::{Root, Body, Document}` snapshot model. Native root/body decoding,
object lookup, and protobuf adaptation stay in `litchi-iwa`; the semantic crate
never imports an `Archive` or generated schema. `litchi-keynote` owns `Slide`,
`Show`, build, and transition values; both depend downward on
`litchi-iwa-text` only.
The Pages `header_footer::{Template, Kind}` role vocabulary is likewise
owned by `litchi-pages`: it is fixed-size and archive-free, while IWA retains
header/footer object discovery, native identifiers, text-storage resolution,
and package mutation.
Pages formatter values now follow the same boundary in
`litchi-pages::{section, page_layout, document_options, footnote}`. Section
pagination, validated page geometry/orientation, document formatter flags, and
footnote formatter values retain only compact semantic state and lossless
native discriminants; the IWA adapter retains document/package identifiers,
protobuf codecs, discovery aggregates, opaque background payloads, and
transactional mutation.
The shared text leaf now also owns the strict `font::{Font, Name}` vocabulary
and its typed `NameError`; the IWA facade keeps only a thin error conversion and
native archive adapters. `Name` stores one boxed UTF-8 identifier, validates
before allocating borrowed input, and consumes owned `String` input directly.
The leaf therefore remains archive-free while Pages, Numbers, and Keynote use
one canonical font model instead of maintaining format-local copies.
The shared text-frame column vocabulary now follows the same boundary at
`litchi-iwa-text::columns::{Columns, Count, Gap, Width, Equal, Following,
Variable}`. The leaf owns only the bounded, archive-free equal/variable layout
and its typed validation error; variable layouts use one boxed following-column
slice. `litchi-iwa` retains `ColumnsArchive` decoding, native presence checks,
protobuf construction, and format-specific error mapping in its private text
adapter. The former flat `TextColumn*` definitions and facade reexports are
removed rather than retained as compatibility aliases.
The Numbers table-cell display-format vocabulary now follows the same boundary
at `litchi-numbers::cell::data_format`. `DataFormat` and its focused child
modules own checked number, currency, percentage, scientific, fraction,
numeral-system, date/time, duration, control, pop-up, text, and custom values;
the semantic crate contains no protobuf, registry identifier, archive, or
package state. `litchi-iwa` retains the native format-table/control codecs,
custom-format UUID registry, BNC scalar coordination, unknown-field
preservation, and transactional package publication. Pages, Numbers, and
Keynote table APIs consume the Numbers leaf directly, and the former flat
`litchi-iwa::table_cell_data_format` and `table_cell_number_format` owners are
deleted rather than retained as compatibility aliases.
The rich-text storage vocabulary now follows the same boundary at
`litchi-iwa-text::storage::{Storage, Run, Fragment}`. `Storage` owns only UTF-8
text and validated byte ranges in one text allocation plus one boxed run slice;
native object IDs, style-table IDs, protobuf messages, and archive terminology
do not enter the leaf. The IWA adapter retains decoded storage-message
selection, native identifiers used for lookup and diagnostics, and all raw
unknown wire content. Keynote, Pages, and structured aggregation consume the
short semantic types directly; invalid ranges cross the leaf boundary as its
typed storage error.
The archive-free object-index foundation now lives in `litchi-iwa-index`. It
owns only typed fragment identities, checked byte spans, immutable object
records, and deterministic reference queries over `litchi-iwa-graph`; it does
not depend on ZIP, Snappy, protobuf, package, or concrete iWork crates. Native
payloads, unknown fields, archive traversal, and the future private IWA index
adapter remain below this leaf.
Keynote build semantics use the focused `litchi-keynote::build` leaf. Its
bounded unknown identifiers, finite effect parameters, typed actions/emphasis,
and boxed motion-path values contain no object or archive identity; native
build CRUD and conversion remain an adapter migration seam.
The common color leaf now owns `color::{RgbColorSpace, Rgba}` and its typed
`color::Error`; native protobuf conversion remains in the IWA shape adapter.
`Rgba` is a fixed-size, copyable value that validates all four finite channels
before construction, so format owners do not allocate or import archive error
state merely to exchange a color.
The shared table-appearance value now follows the same boundary at
`litchi-iwa-common::table::appearance::{Appearance, Banding, RowSizing,
GridlineVisibility, Gridlines}`. These compact, fixed-size values contain no
style inheritance, protobuf, or package state. `litchi-iwa` retains the native
bool conversion, bounded style-inheritance walk, wire decoder, and
copy-on-write style mutation as the concrete archive adapter; its contextual
`Table*` names are only migration-facing facade aliases.
The archive-free table-cell text layout now follows the same ownership rule at
`litchi-iwa-common::table::cell::layout::{TextWrap, VerticalAlignment, Inset,
Insets, Layout}`. The common leaf owns compact fixed-size values and its
allocation-free inset validation error; the IWA Numbers adapter retains native
alignment identifiers, padding archives, cell-style inheritance, and
transactional package mutation. All concrete table editors consume the common
module directly; the former public `litchi-iwa::table_cell_layout` facade and
its contextual aliases are removed rather than retained as compatibility
paths.
The archive-free hidden-axis value now follows the same boundary at
`litchi-iwa-common::table::axis::{AxisIndex, HiddenAxes}`. `HiddenAxes` stores a
sorted, duplicate-free boxed slice and reports duplicate positions through its
typed module error; the IWA adapter retains hidden-state UUID ownership,
protobuf fields, archive traversal, bounds checks, and package mutation.
Numbers, Pages, and Keynote consume the common types directly, and the former
flat IWA semantic definitions and contextual aliases are removed.
Shape and ordinary text-box frame layout is a distinct, table-independent value
module at `litchi-iwa-common::text::layout::{VerticalAlignment, AutoSize, Inset,
Insets, Layout}`. These values are heap-free, fixed-size, and archive-free; the
common module owns finite non-negative inset validation, while `litchi-iwa`
retains protobuf conversion, style inheritance, shared-style ownership, and
transactional package mutation. The old `ShapeText*` value family is removed,
so Pages, Numbers, Keynote, and their creation examples use the contextual
`text::layout` module directly without a second facade model.
Media classification is likewise owned by the compact, archive-free
`litchi-iwa-common::media::Type`. It classifies extensions and bounded
signatures without allocating, preserves the conservative unknown-`ftyp`-as-
video rule, and makes `Unknown` explicit. `litchi-iwa` retains `MediaAsset`,
catalog traversal, package/protobuf metadata, filesystem I/O, resource limits,
and transactional replacement; all consumers import the common type directly
and the old facade-owned enum is removed.
Media playback values now follow the same boundary at
`litchi-iwa-common::media::playback::{MediaVolume, MediaLoopMode,
MediaPlaybackSettings}`. The common module owns the compact duration, volume,
loop-mode, and builder/validation vocabulary, including lossless unknown loop
values. `litchi-iwa` retains movie protobuf decoding, legacy/modern loop-field
reconciliation, wire-preserving replacement, package transactions, and IWA
error mapping; Pages, Numbers, Keynote, and their creation examples import the
canonical common types directly, with no facade compatibility aliases.
Keynote slide-audio creation uses the adjacent archive-free
`litchi_keynote::slide::audio::Options` value. It owns only a common geometry
point and one canonical native `f32` duration, validating finite coordinates
and a positive representable duration before an IWA adapter can mutate a
package. `litchi-iwa` retains slide and drawable identifiers, `TSD.MovieArchive`
decoding, object-graph discovery, zero-size control geometry, media insertion,
raw-wire updates, build construction, and transaction-scoped readback. The
audio info and removal result remain IWA-owned because they carry native IDs,
drawable properties, playback presence, and package-GC results.
The shape-path value slice follows the same ownership boundary at
`litchi-iwa-common::shape::path::{Preset, CornerRadius, PolygonSides,
StarPoints, InnerRadiusRatio}`. These compact, copyable controls and the
source-buildable preset enum contain no archive, protobuf, or package state.
`litchi-iwa` retains structural `ShapePathKind`, native path classification,
geometry-dependent validation, protobuf conversion, and wire-preserving path
patching. Its path-source adapter preserves the envelope's known metadata,
unknown fields, and family-field position while replacing the owned family
payload. The three concrete format owners consume the common `Preset` directly;
the former redundant `Shape*` value names are removed rather than retained as
compatibility aliases.
The dependency-free shape-geometry leaf likewise owns only the compact
`shape::geometry::{Point, Size, FlipAxis}` values. `litchi-iwa` retains the
aggregate `DrawableGeometry` adapter because optional native field presence,
reflection flags, rotation conventions, and wire-preserving patching are
format-specific; the common crate never imports those archive details.
Shape gradients follow the same boundary at
`litchi-iwa-common::shape::fill::{Kind, Angle, StopPosition, StopMidpoint,
Opacity, Stop, Gradient}`. The common value layer stores validated colors as
`color::Rgba`, keeps scalar controls fixed-size, and owns gradient stops in a
boxed slice without protobuf or package dependencies. `litchi-iwa` retains
`ShapeFill` because image fills and native data references are facade-owned;
its fill adapter alone decodes and writes `GradientArchive` while preserving
strict validation. The former `ShapeGradient*` owners are removed rather than
retained as compatibility aliases.

Native straight-line decorations follow the same boundary at
`litchi-iwa-common::shape::line::{Endpoint, Endpoints}`. The two-byte value
stores directed start and end decorations without native line-end archives;
`litchi-iwa` retains endpoint inheritance, style variation, and wire updates.
The concrete Pages, Numbers, and Keynote editors consume the common value
directly, and the former `LineEndpoint*` names are removed.

Chart kinds follow the same boundary at `litchi-iwa-common::chart::kind::Kind`.
This fixed-size value preserves every native integer, exposes only archive-free
capability predicates, and keeps protobuf conversion in `litchi-iwa`. The old
protobuf-coupled `ChartKind` owner is removed rather than retained as an alias.
Chart-axis selectors and tick-mark values now follow the same boundary at
`litchi-iwa-common::chart::axis::{Axis, TickMarkLocation}`. `Axis` is a
one-byte category/value selector, while `TickMarkLocation` is a compact
copyable value that keeps an unrecognized native integer explicit. The common
module owns no archive, protobuf message, or package state; IWA retains axis
object lookup, field mapping, and lossless wire mutation. Pages, Numbers,
Keynote, and chart examples consume the short common names directly, and the
former `ChartAxis*` facade names are removed rather than kept as aliases.
The remaining archive-free axis values now use focused child modules under the
same owner: `axis::bounds::{Bound, Bounds}`, `axis::label_angle::LabelAngle`,
`axis::label_position_3d::LabelPosition3d`, `axis::scale::Scale`, and
`axis::steps::{MajorStepCount, MinorStepCount, Steps}`. IWA retains only the
native field numbers, protobuf conversion, chart-kind capability checks,
archive lookup, and lossless wire mutation; no axis value module depends on
the facade or protobuf crates.
Chart label number-format values now follow the same ownership boundary at
`litchi-iwa-common::chart::number_format`: `FixedDecimalPlaces`,
`DecimalPlaces`, `NegativeStyle`, `NumberFormat`, and `LabelAffixes` are
archive-free semantic values with concise names. The scalar format is packed
into one byte, while affixes share one bounded allocation. IWA retains
`DualNumberFormatFields`, native field identifiers, protobuf decoding, strict
legacy/current reconciliation, and lossless wire patching. Axis and series
defaults remain explicit because their native thousands-separator defaults
differ; the former long `Chart*` number-format names are removed rather than
kept as compatibility aliases.
Chart series orientation now follows the same boundary at
`litchi-iwa-common::chart::Direction`. It is a small, copyable value with
enum-style `Rows` and `Columns` constants plus a `DirectionKind` projection;
unknown native integers remain lossless. The common crate owns only this
archive-free semantic vocabulary and its compact native representation. IWA
retains protobuf field mapping, archive lookup, and mutation validation; Pages,
Numbers, Keynote, and chart readers consume `Direction` directly. The former
`ChartSeriesDirection` facade type and protobuf-dependent implementation are
removed.
`litchi-keynote::transition` owns Keynote's archive-free transition semantics:
the focused `Settings`, `AnimationParameters`, and `CustomParameters` values,
their memory-conscious opaque owned semantic payload containers, and the
existing `Effect` and scalar values
(`Direction`, `MosaicType`, `Acceleration`, and `TextDelivery`). The semantic
constructors enforce bounded ownership plus finite-number, NUL-free text, and
canonical-value validation before a value is published. No raw native IDs or
archives leak into this crate or its public API. `litchi-iwa` retains only
native/protobuf decoding, payload structural validation, wire patching, graph
lookup, and transactions; opaque payload decoding remains at that IWA
boundary.
Pie and donut label semantics follow the same focused boundary at
`litchi-iwa-common::chart::pie::{LabelVisibility, LeaderLineVisibility}`.
`LabelVisibility` is a one-byte bitset for data-point names and values, while
`LeaderLineVisibility` is a four-byte transparent native integer that preserves
future states losslessly. IWA retains the pie field identifiers, strict
varint validation, series graph, stylesheet/object-container ownership, and
transactional package mutation; the former `ChartPie*Visibility` names are
removed rather than retained as aliases.
Chart category-label semantics follow the focused
`litchi-iwa-common::chart::category_labels::{Interval, Frequency, Layout}`
module. The common values are heap-free and archive-free: `Interval` admits
only explicit values from 2 through the native signed maximum, while
`Frequency` retains canonical automatic/all modes and unknown signed native
values losslessly. IWA retains interval field numbers, strict int32/boolean
wire validation, axis visibility, style-slot ownership, and transactional
package mutation; the former `ChartCategoryLabel*` owners are removed.
Chart reference lines follow the same ownership boundary at
`litchi-iwa-common::chart::reference_line::{Value, Kind, Line}`. `Value` is a
finite transparent scalar; `Line` stores a bounded optional label and packs its
two visibility flags; and `Kind::Unsupported` can only be created through a
checked lossless constructor. The public IWA path is the focused
`charts::reference_line` module. IWA retains generated protobuf schemas,
extension framing, graph/object ownership, and package transactions, including
pre-decode graph budgets and wire-preserving nested custom-value patches. Typed
graph updates use an occurrence-aware raw-wire merge that preserves unknown
fields inside graph, axis, item, style, sparse-reference, reference, and UUID
messages instead of rebuilding those messages through Prost.
The former flat `ChartReferenceLine*` model is removed rather than kept as an
alias.

The Keynote soundtrack leaf follows the same boundary at
`litchi_keynote::soundtrack::{Mode, Settings}`. `Mode` is a compact semantic
enum that round-trips unknown native discriminants, while `Settings` validates
finite playback volume and canonical known modes without importing protobuf,
archive, graph, package-ID, media-reference, or transaction state. The IWA
adapter retains native `KN.Soundtrack` decoding, optional-field presence,
unknown-field preservation, soundtrack media references and their metadata,
and atomic package edits; no long `KeynoteSoundtrack*` semantic aliases remain.
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
reader seam into the immutable `litchi_numbers::Sheet` model through a lazily
cached `Arc<[Sheet]>`; it transfers finished sparse tables without rebuilding
cell maps and intentionally leaves comments/native sidecars on the opaque
archive adapter. The adapter's conversion is private and the semantic leaf
remains the only archive-free Sheet owner. The dependency-free formula
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

Numbers table axis sizing follows the same boundary at
`litchi-numbers::table::dimension::{Dimension, Points, Size}`. The leaf owns
only the archive-free row/column selector, finite positive point validation,
and the distinction between native-default and explicit sizing. `litchi-iwa`
retains header-bucket discovery, the native zero sentinel, archive bounds,
wire-preserving mutation, and transactional reparse verification. The former
IWA semantic definitions are removed; only crate-private re-exports remain
where untouched format modules still resolve the shared type during their
later migrations.

Numbers table section settings follow the same ownership boundary at
`litchi-numbers::table::headers::{Count, Settings}`. `Count` is a compact
`NonZeroU8` value for the native `1..=5` domain, so an optional count retains
presence without a second byte; `Settings` preserves optional count and
Boolean field presence while exposing only archive-free effective-value
helpers. Pages and Keynote consume these canonical Numbers table values
directly. IWA retains native model conversion, wire-presence and framing
validation, header/body/footer capacity checks, object lookup, unknown-field
preservation, and staged transactional publication. The former
`NumbersTableHeader*`, `PagesTableHeader*`, and `KeynoteTableHeader*` facade
names are removed rather than retained as compatibility aliases. Counts above
the native range are rejected as a typed malformed-document/value error; they
are not silently widened or normalized.

The Pages and Keynote table readers now consume the same leaf `Table` through
an ownership-preserving adapter seam. Their public table facades borrow the
canonical sparse cells directly while retaining format-owned comments and
merge regions as separate sidecars; read-only comment sidecars are compact
sorted boxed pairs, and the former tuple-keyed cell maps are no longer rebuilt
in either reader. The generic structured extractor remains the last current
`TableDataExtractor` consumer and is staged separately.

Shared iWork merged-cell geometry now follows the same ownership boundary at
`litchi-numbers::table::merge::{Region, Axis, Deletion, AnchorRelocation}`.
`Region` is a checked, 16-byte archive-free rectangle backed by compact `u32`
coordinates and non-zero spans; the leaf also owns the pure axis insertion,
deletion, and surviving-anchor transformations. Numbers, Pages, and Keynote
table facades accept and borrow this one semantic type rather than publishing
an `IWorkTableCellRegion` or format-prefixed duplicate. `litchi-iwa` retains
only merge-formula parsing, native table-bound checks, unknown-wire
preservation, formula-anchor relocation, and transactional package mutation;
the conversion from physical `usize` indices to the bounded leaf domain is
checked at that adapter boundary.

The archive-free result aggregation is now isolated in
`litchi-iwa-structured`. Its `StructuredData` value depends only on the
semantic `litchi-keynote`, `litchi-numbers`, and `litchi-pages` leaves; it has
no protobuf, ZIP, package, graph, or facade dependency. `litchi-iwa` retains
the private application-specific archive traversal and constructs this value
at the adapter boundary. This keeps cross-format result composition below the
physical reader and gives the eventual concrete format crates a reusable
semantic handoff without making the common vocabulary crate depend upward.

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

The Numbers display-format seam now follows the same rule at
`litchi-numbers::cell::data_format`. Checked number, currency, percentage,
scientific, fraction, numeral-system, date/time, duration, custom, and
interactive-control values are archive-free leaf types with bounded text,
finite numeric ranges, and typed construction errors. Native identifiers,
protobuf fields, custom-format registries, and package publication remain in
`litchi-iwa`'s private adapter; legacy IWA format structs are no longer used
as the semantic API. The leaf intentionally uses compact scalar wrappers and
boxed variable-length values so ordinary formats do not carry archive state.

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

Pages body-footnote semantics now follow the same downward boundary in
`litchi-pages::footnote::body::{Footnote, Position, Selector}`. The leaf owns
only bounded text, custom-marker, and UTF-16-position values; selector-based
CRUD in `litchi-iwa` uses source order or body position and never publishes a
package/runtime object identifier. Reference, footnote-storage, and marker
objects remain private to the IWA graph adapter, including cleanup after
ordinary body edits.

## iWork split and migration host

The current iWork owners are concrete rather than prospective. Shared physical
ZIP/IWA framing lives in `litchi-iwa-archive`, `litchi-iwa-core`,
`litchi-iwa-detect`, and their focused leaves. `litchi-pages`,
`litchi-numbers`, and `litchi-keynote` each own application package ingress and
publish archive-free semantic documents. In particular,
`litchi-pages::package` owns native Pages ZIP, IWA, and protobuf traversal;
`litchi-pages::{document,section}` remain semantic and archive-free.

`litchi-keynote::package` owns source-backed Keynote playback-state
transactions. It resolves slides by exact navigator name or checked semantic
position, privately maps that selection to the required slide-node payload,
and publishes only after bounded ZIP/IWA reassembly plus full semantic
readback. Raw component names, protobuf values, and native object identifiers
remain inside the package adapter. Shared ZIP, Snappy, archive framing, and raw
wire patching remain owned by the focused infrastructure crates; the Keynote
semantic slide exposes only `is_skipped`.

The same concrete package owner now resolves the strict document/show/slide
graph for semantic text. It consumes the focused `litchi-iwa-text-wire` Buffa
projection privately after exact message-type and graph-ownership checks; no
generated Buffa or protobuf value crosses the supported Keynote boundary.
`litchi-iwa-archive::SourceCatalog` binds the authoritative immutable package
bytes, physical/logical entries, explicit exact-versus-legacy provenance, and
stable component ordinals. Keynote therefore classifies, indexes, reads
metadata, and prepares its focused edit from one snapshot instead of reopening
the ZIP. Pages likewise reads metadata and semantic components from one
catalog. The format adapters can retain compact sorted object locators without
depending on the monolith's object index or repeating component scans. The
dependency remains downward: neither the wire projection nor the archive
catalog depends on a concrete format.

`litchi-numbers-wire` owns the shared low-level Binary Numbers Cell codec. It
is a versioned physical adapter so the migration host can depend on it without
re-exporting a format implementation. `litchi-numbers` uses it privately, and
the supported `litchi-numbers` and root `litchi::numbers` APIs expose neither a
`cell::wire` module nor BNC flags, native format indices, or raw object IDs.
Direct use of the wire crate is an explicit opt-in to unstable native storage
details.

Direct `litchi-numbers::Package` ingress now depends on the focused
`litchi-iwa-detect` leaf to prove that the root payload belongs to Numbers
before any TN schema is decoded. A schema-directed wire-tree preflight bounds
the one canonical detector payload under the selected physical profile;
arbitrary root siblings are not application candidates. The same application
proof protects the rooted-independent compatibility-table byte entry point.
This remains a downward format-owner-to-infrastructure edge, matching the
existing Keynote package boundary; it does not route through `litchi-iwa` or
expose detector, protobuf, component, or object identifiers in the supported
Numbers API.

The private protobuf leaf also owns a derived Buffa projection for the Numbers
GroupNode category path. Generated code contains only an empty node envelope,
UUID, and four scalar wrappers. The bounded adapter streams recursive children
and CellValue routing from preflighted source bytes, so it retains neither an
input-width child index nor unknown fields. Generated Buffa types remain
private, traversal memory is proportional to depth, and source IWA bytes remain
authoritative.

`litchi-iwa` is now the sole declared iWork migration host, not a canonical
dependency layer. Every one of its internal workspace edges is an ordered
migration-debt item with a reason and a deletion condition. The host may be
removed only after every remaining editor, example, fuzz target, generated
schema path, and compatibility test has moved to a concrete format or focused
substrate. No compatibility crate, feature, module, alias, or facade re-export
may retain it after that gate.

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

## 2026-08-08 amendment: Keynote show and slide-order ownership

`litchi-keynote` now owns the application-shaped projection of
`KN.ShowArchive` and its embedded `KN.SlideTreeArchive`. A narrow generated
Buffa lazy view is derived from the canonical schema and remains private to the
protobuf infrastructure crate. Keynote performs schema-directed wire preflight
before constructing that view, streams the ordered slide references instead of
materializing an attacker-sized generated repeated-message index, and forces
every deferred value that affects semantic publication. Generated Buffa and
Prost values, native object identifiers, component names, and raw slide-tree
types do not enter the supported format API. The accepted source bytes and
complete validated raw field records, including each encoded key, encoded
length, and payload, remain the preservation authority.

The concrete package also owns selector-first slide ordering through the direct
`Package::edit_slide_order() -> SlideOrderEdit<'_>` entry point,
`SlideOrderCommit`, and `SlideOrderPatch`, with format-owned
`SlideOrderDiagnostics`, `SlideOrderError`, and `SlideOrderLimitKind`. The
source is selected by exact navigator name or checked semantic position and the
destination is a checked final position in the base slide list. This
transaction family is deliberately separate from the existing skip/include
`Edit`, `Commit`, and `Patch` types, whose Boolean accessors remain source
compatible. Commit and patch application publish only after bounded package
reassembly, complete reopening under the retained read options, and semantic
order readback.

The migration host no longer owns `KeynoteEditor::move_slide`, its focused
example, or its move-specific compatibility assertions. Slide creation,
duplication, deletion, show-setting mutation, and the larger slide, build,
shape, note, table, chart, and media editors remain host migration work. This
vertical move does not remove a dependency edge: all 17 ordered
`litchi-iwa` debt entries remain until their complete exit conditions hold.

## 2026-08-08 amendment: focused Keynote show settings and graph-edge retirement

The preceding show-setting and dependency counts describe the slide-order
slice and are superseded by this amendment. `litchi-keynote::Package` now owns
`show_settings`, `edit_show_settings`, exact-source patch application, and the
format-owned `ShowSettingsEdit`, `ShowSettingsCommit`, `ShowSettingsPatch`,
`ShowSettingsDiagnostics`, `ShowSettingsError`, and `ShowSettingsLimitKind`
types. The focused reader validates the complete known `KN.ShowArchive` and
embedded `KN.SlideTreeArchive` envelope, including the slide-reference ceiling,
but projects only the presentation size and scalar settings through the private
Buffa view. It neither initializes the full semantic slide cache nor retains a
slide-node identifier list. Generated Buffa and Prost values, native object
identifiers, component names, and raw wire objects remain outside the supported
API. Buffa does not retain unknown content; validated raw source field records
remain the preservation authority.

A null root show reference reads as `Settings::default()`. Because there is no
owning Show component to rewrite, only an exact no-op edit is publishable for
that state. For a present Show, changed exact-package sources rewrite one
owning IWA component and publish only after complete reopening under the
retained `ReadOptions` and focused semantic readback. Changed legacy nested
`Index.zip` sources remain a typed `UnsupportedSource`: their historical
normalizing mutation stays in the migration host together with its compatibility
method, example, and assertions. This is therefore focused exact-source
ownership, not complete retirement of the host's show-settings compatibility
path.

The migration host also no longer depends directly on `litchi-iwa-graph`.
Authoritative `MessageInfo` references and schema fallback references are
inserted directly into `litchi-iwa-index::IndexBuilder`; graph identities and
the immutable graph snapshot are consumed through the index owner's existing
reexports. Ordered debt 007 is removed without renumbering later debt
identities. The boundary inventory is now 63 workspace packages, 223 internal
dependency declarations, and 16 ordered migration debts. The canonical
`litchi-iwa-index -> litchi-iwa-graph` edge remains, and this direct-edge
retirement does not claim that the host's remaining graph-backed editors have
migrated.

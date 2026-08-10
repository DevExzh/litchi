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
archives leak into its public API. The concrete `Package` owner also provides
selector-first, exact-source modern slide-transition transactions. A private
Buffa lazy view projects the known native fields after strict bounded wire
preflight; validated raw records remain authoritative for wire-preserving
rewrites, and a full retained-options reopen verifies each candidate. No
archive, protobuf value, component name, or native object ID is exposed. The
legacy `litchi-iwa` Keynote reader/editor compatibility surface remains
available; this owner transaction does not claim that surface was deleted.
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

## 2026-08-08 amendment: focused Pages section-name ownership

`litchi-pages::Package` now owns the first Pages-specific exact-source
transaction: selector-first section-name replacement through
`edit_section_name`, `apply_section_name`, and the format-owned
`SectionNameEdit`, `SectionNameCommit`, `SectionNamePatch`, diagnostics, error,
and limit types. The supported surface exposes an exact section name or a
checked semantic `Position`; native object identifiers, IWA member names,
generated messages, and raw wire values remain private.

The concrete owner rewrites only field 26 of the selected native section
message, preserves the complete object header through the shared bounded core
helper, and publishes only after exact-package reassembly, a complete reopen
under the retained limits, and semantic readback. Raw validated field records
remain the unknown-content authority. This mutation does not require a Buffa
projection: decoding into a generated value and re-encoding it would weaken
the exact-record preservation contract. Buffa lazy views remain the policy for
schema projections that need semantic decoding, not a replacement for the raw
preserving mutation primitive.

The migration host's raw-identifier section-renaming example is removed and
its README directs callers to the focused Pages example. The host's
normalizing `PagesEditor::set_section_name` compatibility path remains for
changed legacy nested-`Index.zip` sources, which the exact-source transaction
deliberately refuses. No dependency edge is retired in this slice; the ledger
remains at 63 workspace packages, 223 internal declarations, and 16 ordered
migration debts.

## 2026-08-08 amendment: focused Pages section-text ownership

`litchi-pages::Package` now owns selector-first section-body text reads and
rooted exact-source text transactions through `section_text`, `edit_section_text`,
the single-section `edit_body_text` convenience, `apply_section_text`, and the
format-owned `SectionTextEdit`, `SectionTextCommit`, `SectionTextPatch`,
diagnostics, error, and limit types. `litchi-pages` re-exports the shared
archive-free `TextPosition` and insertion-capable `TextSpan`; callers select a
section by exact semantic name or checked position and never exchange a body
storage identifier, object number, component name, protobuf value, or package
member.

The byte-preserving `TSWP.StorageArchive` splice kernel is owned below the
concrete format in `litchi-iwa-text-wire`. It bounds fields, fragments,
positional-table entries, nested messages, references, text, output, and total
rewrite work; adjusts every recognized position-bearing table according to its
schema policy; retains untouched raw records; and reports reference deltas
without deciding application ownership. Pages retains the section-boundary,
dependent-content, exact-package, and publication policy. The migration host
also consumes this downward kernel without making `litchi-pages` depend on
`litchi-iwa`.

Private Buffa lazy views validate the known document-to-body references,
section-boundary references, and text-storage projection only after strict
bounded raw preflight. Generated Buffa and Prost values remain private and are
not used to publish a rewritten message. The accepted raw records remain the
preservation authority for unknown and untouched content. On a rooted exact
source with one unambiguous native body storage, a changed candidate replaces
one body payload while preserving its IWA object header, reassembles one
component, reopens the complete package under the retained limits, and
performs semantic and native-topology readback before publication.

The focused transaction refuses edits that would consume native section
breaks, footnote anchors, inline-object markers, or their owned reference
metadata; graph deletion remains a separate Pages capability. Exact no-ops,
including legacy nested-`Index.zip` sources, share the original immutable
source. Changed legacy sources remain `UnsupportedSource`, and the host's
normalizing raw-ID methods remain compatibility-only while that provenance
case has no preservation-safe owner. Changed no-root/fallback bodies are also
unsupported until their physical ownership has an explicit mutation boundary.
This slice transfers a substantial Pages mutation vertical but removes no
manifest edge: the inventory remains 63
workspace packages, 223 internal dependency declarations, and 16 ordered
migration debts.

## 2026-08-08 amendment: physical cache state and focused projection boundary

The preceding 16-debt ledger is historical and is superseded by this
amendment. Cache-backed `PackageState` has moved from the `litchi-iwa`
migration host to the physical `litchi-iwa-archive` owner. The archive now owns
the bounded physical parsed-component state; `litchi-iwa-cache` remains a
dependency-free cache leaf. The host retains format selection and error-policy
decisions, so the move does not transfer application semantics or make the
cache responsible for archive policy. The direct `litchi-iwa ->
litchi-iwa-cache` debt, identity 003, is retired without renumbering later
identities. The current boundary inventory is 63 workspace packages, 223
internal dependency declarations, and 15 ordered migration debts.

For Numbers, this is a deliberately narrow decoder boundary: focused
`TableInfo.tableModel` reads use a strict,
small private Buffa projection. A bounded raw-wire preflight precedes Buffa,
the selected table-model reference must be nonzero, and the projection neither
encodes, retains unknown content, nor stores repeated fields. Accepted raw
source remains authoritative for preservation. This does not claim migration of
the broader table model or the wider Numbers graph.

## 2026-08-08 amendment: focused Keynote speaker-notes ownership

`litchi-keynote::Package` now owns selector-first reads and exact-source text
transactions for an existing slide's existing speaker-notes graph. The public
surface exposes `slide_notes`, `edit_slide_notes`, `apply_slide_notes`, the
format-owned `SlideNotesEdit`, `SlideNotesCommit`, `SlideNotesPatch`,
diagnostics, error, and limit types, plus the shared archive-free
`TextPosition` and `TextSpan`. Callers select by exact navigator name or
checked semantic position and never exchange native object identifiers,
component names, generated messages, or raw wire values.

A strict private Buffa projection is limited to the selected
slide-to-note-to-text-storage ownership references. Its generated closure is
exactly five files and 151,735 bytes under a 160 KiB build cap, contains no
repeated view, and has no production encoding path. Bounded schema-directed
raw preflight proves canonical selected fields, required envelopes, nonzero
references, field/byte/work/nesting ceilings, and forces every lazy value that
affects publication. Buffa does not retain unknown content and is not the
preservation representation; accepted caller-owned raw records and exact IWA
object headers remain authoritative.

The format owner independently scans every slide and note payload to prove a
unique slide-to-note and note-to-storage edge, counts every metadata reference
occurrence, requires exact owner/type association, and refuses dependent or
unknown note shapes. A changed transaction performs checked UTF-16
set/clear/insert/delete/replace, rejects surrogate splits and reserved native
markers, rewrites one component, validates all outer IWA header-length
prefixes, completely reopens under the retained limits, and performs semantic
and topology readback before publication. Exact no-ops retain the original
source bytes; exact inverse application restores them.

The migration host's raw-ID notes example has moved to the semantic Keynote
example. The host's broader notes compatibility API and graph creation or
deletion behavior remain migration work, as do legacy provenance, durable
patch serialization, and atomic file publication. No ordered host debt is
retired by this vertical move.

The boundary checker now distinguishes normal and development-only use for
every internal declaration. Exclusively development-only edges must have an
exact stale-checked policy annotation, and promotion, mixed normal/dev use,
missing annotations, and stale annotations fail closed. Numbers and Pages test
fixture writers now use the archive owner's package serializer, removing their
two direct development-only `soapberry-zip` edges. The current inventory is 63
workspace packages, 221 internal dependency declarations, and 15 ordered
migration debts; this is dependency hygiene, not debt retirement.

## 2026-08-08 amendment: structured read-seam retirement

The historical construction of neutral structured data inside `litchi-iwa`
is superseded. The root facade now asks each concrete format owner for its
semantic projection and constructs `litchi-iwa-structured::StructuredData`
directly. Numbers owns its deliberately allocating package-global
compatibility projection, including valid detached table models; Pages and
Keynote retain their focused rooted document projections. The neutral crate
continues to own only the format-independent aggregate model and aggregate
budgets.

The host's `structured` module, `StructuredData` re-export,
`Document::extract_structured_data` method, adapter tests, support-only
Numbers hooks, and manifest edge to `litchi-iwa-structured` are removed. Debt
identity 011 is deleted without renumbering later debts. The checked inventory
is now 63 workspace packages, 220 internal dependency declarations, and 14
ordered migration debts.

This is the retirement of one read seam, not the monolith. The migration host
still owns unmigrated editing and compatibility behavior, the wider Numbers
payload graph still contains bounded eager Prost paths, and root preparation
can transiently materialize unrelated ZIP members. Those ownership, Buffa,
and aggregate peak-memory boundaries remain open.

## 2026-08-09 amendment: focused Keynote title/body ownership

`litchi-keynote::Package` now owns selector-first reads and exact-source text
transactions for an existing slide's existing semantic title and body
placeholders. The role-aware surface consists of `slide_text`,
`edit_slide_text`, and `apply_slide_text`, with `slide_title`, `slide_body`,
`edit_slide_title`, and `edit_slide_body` conveniences. Its format-owned
`SlideTextRole`, edit, commit, patch, diagnostics, error, and limit types use
the shared archive-free `TextPosition` and `TextSpan`. Callers select a slide
by exact navigator name or checked semantic position; placeholder, storage,
component, and protobuf identities remain private.

The slide-text Buffa seam has two format-ownership projections. The existing
speaker-notes codec now also projects `KN.SlideArchive` fields 5 and 6, the
optional title- and body-placeholder references; its required style,
transition, and in-document fields and its name and note fields remain part of
the same bounded slide-owner snapshot. The new placeholder codec projects the
selected `KN.PlaceholderArchive` kind and the singular required inheritance
chain `PlaceholderArchive -> TSWP.ShapeInfoArchive -> TSD.ShapeArchive ->
TSD.DrawableArchive`, ending at `ShapeInfoArchive.owned_storage` field 4.
The selected read forces the slide-owner view. Package-wide ownership proof
first raw-scans fields 5 and 6 of every slide candidate and forces that Buffa
view only for a candidate that references the selected placeholder. It
similarly raw-scans every placeholder candidate and forces the placeholder
view only when its modern or deprecated storage edge can reference the
selected storage. The same bounded raw scanner audits
`ShapeInfoArchive.deprecated_storage` field 2, `ShapeInfoArchive.text_flow`
field 3, standalone shape-info references, embedded `TSP.Reference` metadata,
and `NoteArchive.containedStorage` field 1. Alias discovery does not force the
Buffa `NoteArchive` view. The existing
`litchi-iwa-text-wire` storage codec and raw text rewrite remain the text-value
seam; this amendment does not make the whole operation a two-message Buffa
decode.

Schema-directed raw preflight precedes each Buffa access, and the lazy fields
used for authorization are forced. The two owner projections have no
production encoding path or unknown-field retention. Accepted raw records,
exact IWA object headers, and the caller-owned package source remain the
preservation authority, subject to an explicit rendered-cache exception. A
changed edit commit proves role-correct exclusive ownership, performs one
checked UTF-16 set, clear, insert, delete, or replacement, and rewrites the
selected storage plus the selected `KN.SlideNodeArchive` preview state. Those
objects occupy one or two distinct IWA components, which are the value reported
by `SlideTextDiagnostics::touched_components`. The slide-node rewrite removes
its thumbnail references and rendered thumbnail fields, marks thumbnails
dirty, prunes the referenced preview object IDs and preview-owned aggregate or
field data-reference occurrences, retains proven unrelated data references,
and rejects ambiguous aggregate-only ownership. A selected
slide that already carries the separate cached title/body strings in
`KN.SlideArchive` fields 37 or 38 is rejected until that cache has a proven
rewrite rule.

The physical `litchi-iwa-archive` owner now supplies bounded, exact-name,
deletion-aware flat-package reassembly. The Keynote transaction uses it to
delete any root `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg`
members while retaining the raw physical records of every other unedited ZIP
member. Storage, slide-node, and preview mutations publish atomically as one
candidate. These deletions are not IWA components and are therefore not added to
`touched_components`. Keynote and package preview consumers can otherwise keep
rendering pixels produced before the semantic text edit.

The changed candidate fully reopens under the retained limits and verifies the
selected semantic text, invalidated slide node, absent root previews, remaining
object graph, and unselected semantic slide state before publication. Applying
a changed patch does not reassemble the package; it reopens and verifies the
exact target bytes already stored in the patch and reports the originating
edit's one- or two-component count. An exact edit no-op relies on the immutable
selected snapshot established when editing began, while an exact patch no-op
checks artifact identity. Both leave all caches and previews byte-identical and
return a snapshot that shares its source allocation without whole-source
validation or a candidate reparse. Applying a changed
inverse patch likewise reopens and verifies the exact original artifact,
including its former preview members and cache state.

The migration host no longer provides its nine raw-index
`set`/`replace`/`clear_slide_{title,body,notes}` methods or their private
storage-resolution helpers. This is a breaking replacement, not a
compatibility shim: title and body callers must use a semantic selector and
role, checked UTF-16 spans, and the immutable commit/package flow, while notes
callers use the previously accepted `SlideNotesEdit` transaction. The focused
owner also rejects ambiguous or shared graphs that the old raw-index editor
could reach. Changed-output compatibility is semantic rather than byte-local to
the storage: callers and differential tests must allow the declared selected
slide-node invalidation and root-preview deletion, while exact no-ops remain
byte-identical. The host's slide-creation, placeholder visibility/layout,
arbitrary text-box, generic text-storage, and other unmigrated editor surfaces
remain separate compatibility work.

This vertical does not create or delete title/body placeholder graphs and does
not cover arbitrary text boxes. It also makes no claim of durable patch
serialization, atomic filesystem publication, whole-Keynote Buffa conversion,
or deletion of the migration host. No manifest edge is removed: the checked
metadata/policy inventory is 64 workspace packages, 235 internal dependency
declarations, and 14 ordered migration debts.

## 2026-08-10 amendment: focused Numbers table-lock ownership

`litchi-numbers` now re-exports the canonical archive-free
`litchi-iwa-common::table::lock::State` through `table::lock` and owns the
selector-first package surface for one attached table's effective interactive
lock state. `Package::table_lock`, `edit_table_lock`, and `apply_table_lock`
accept semantic sheet and table selectors. The public edit, commit, reversible
patch, diagnostics, error, and limit types expose neither native drawable or
model identifiers nor component names, generated messages, or native wire
field identifiers and values.

The physical adapter resolves a selected semantic sheet/table position back
through the rooted document, sheet drawable order, and exactly one canonical
or legacy `TableInfo` payload. `litchi-iwa-protos::table_info_codec` remains the
narrow generated-free boundary. Its bounded, canonical raw-wire preflight
reads the required `TableInfo.super`, the presence-preserving optional
`TSD.DrawableArchive.locked` Boolean, and the required nonzero table-model
reference. The private Buffa lazy views are then forced for both the drawable
`super.locked` value and the table-model ownership reference, and the complete
presence-preserving snapshot is checked against preflight. Buffa is neither
the encoder nor the preservation representation. The projection closure
contains only the drawable envelope, lock scalar, and model reference, has no
production encoding path or repeated view, and leaves all unselected bytes
caller-owned.

This supersedes the 2026-08-08 two-message, opaque-super, 64 KiB TableInfo
scope. The current private projection contains exactly three messages
(`TableInfoArchive`, `DrawableArchive`, and `TableModelReference`), forces and
cross-checks both `super.locked` and the model reference, and generates five
files totaling 83,529 bytes under an 84 KiB cap.

Changed publication raw-patches only drawable field 5 in the selected
table-info message, rewrites one IWA component, reassembles the exact flat
package through `litchi-iwa-archive`, and completely reopens it under the
retained read options before checking the requested state. Changed publication
rejects competing rooted sheet ownership, contradictory selected-sheet or
selected-TableInfo reference metadata, noncanonical outer object-length
prefixes, and merge/diff metadata on the selected owner instead of normalizing
them. Detached or unrooted pseudo-sheet and view-state dependent references
are not competing owners; they remain opaque and preserved. Unknown protobuf
fields, sibling messages, unselected object-header metadata,
unrelated IWA components, and unrelated ZIP members remain
preservation-owned. An absent lock field and an explicit `false` both read as
`Unlocked`, but exact no-ops retain their distinct source encodings.
Exact-source reversible patches retain the complete before and after artifacts;
changed application reopens the stored target rather than reassembling it,
while inverse application restores the original bytes.
Legacy nested `Index.zip` packages remain readable and support exact no-ops,
but changed publication fails closed because that source cannot be preserved
by the flat-package reassembler.

The Numbers-specific host read and mutation seam is retired rather than
retained as raw-ID aliases. Removed surface includes the direct
`table_lock_state` and `set_table_lock_state` methods, private
`table_lock_context`, `NumbersTableInfo.lock_state` and the field-population
branch inside `tables()`, the shared codec's model-specific read/write helpers
`table_lock_state_for_model` and `set_table_lock_state_for_model`, and the
Numbers-only model-ID matching branch. Numbers callers now read through the
focused `Package::table_lock` API as well as mutate through its transaction.
The boundary checker ratchets exactly five functions under the scopes where
collisions are meaningful: three under the host Numbers tree
(`table_lock_state`, `set_table_lock_state`, and `table_lock_context`) plus
both model-specific helpers in the shared codec. It separately rejects a
public `NumbersTableInfo.lock_state` field. The generic shared
getter/setter and wire codec remain for Pages and Keynote, so this vertical
neither deletes that common host path nor removes a manifest edge.

The focused source inventory contains two semantic lock-state unit tests, nine
TableInfo codec tests, and 15 package transaction tests covering selector
resolution, absent/explicit-false/true states, exact no-ops, preservation,
inverse replay, conflicts, legacy refusal, limits, concurrent reads, and a
checked-in native fixture. The focused transaction suite passed 15/15,
including the rooted `FormBasedSheet` drawable field path `[1, 2]` and
preservation of detached/unrooted reference metadata. It also covers changed
flat legacy type-6003 TableInfo publication with exact inverse and partial-sink
write accounting.
The bounded `numbers_table_lock` fuzz target compiles, and all 57 boundary
policy regressions pass. The full policy command still reports the 14
pre-existing soapberry-zip/xml-minifier annotations. A Numbers-only fuzz
package and a sustained sanitizer campaign remain open.

The current writer passed Apple Numbers 14.4 (7043.0.93) open, native Save As,
close, and reopen without warning. Source, Rust-locked, and native-resaved
SHA-256 values are
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`,
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`,
and `8aa87a3afcb145b66c5c6f4e10645cd1cf658f4b65f0976612ac6d62d4652995`.
Numbers showed the table locked with disabled cells and retained the B2/B3
semantics; focused reread remained locked, the equal-state edit was byte-exact,
and inverse application restored the exact source.

This ownership move does not claim an aggregate transaction peak-memory or
total-work contract, a complete proof that all transitive allocations are
fallible, durable patch serialization, or a library-owned atomic durable
filesystem save. The focused example's sibling-temporary no-clobber write is
workflow evidence rather than that library contract. `Package::write_to`
reports the exact accepted byte count on a sink failure. The process-local patch
also lacks a versioned semantic operation envelope, read/write sets,
composition, three-way merge, and bounded history.
Resource and allocation errors do not yet carry the selected semantic table
path, and exact source-byte access still sits on the ordinary `Package`
surface instead of an explicit advanced/raw boundary.
The flattened `TableLock*` transaction names also remain migration debt
against the focused-module short-name rule.
The archive-free `Table` snapshot does not yet carry lock state, and remaining
host table/cell mutations do not enforce that state by default; read-model
convergence and protection enforcement remain host-migration debt. The private
Numbers locator also remains specialized instead of converging on the neutral
IWA index owner.

## 2026-08-10 amendment: focused Pages page-layout ownership

`litchi-pages::Package` now owns the document-wide, presence-preserving page
layout read and exact-source transaction. `page_layout`, `edit_page_layout`,
and `apply_page_layout` operate on the existing archive-free `page_layout::Layout`;
the edit, commit, reversible patch, diagnostics, error, and limit types expose
no object identifiers, component names, message types, generated messages, or
wire fields.

The private adapter resolves exactly one object 1 with exactly one type-10000
`TP.DocumentArchive` payload across the component catalog. Its bounded raw
preflight checks the required opaque `super` envelope and the
presence-preserving layout scalars at fields 30 through 39 and 42, including
canonical keys and varints, singularity, wire types, and Boolean values. It
then forces the corresponding scalar fields on the private Buffa lazy view and
requires the complete projected `Layout` to equal the raw result. The existing
Pages document-body projection is reused: it has no production encoder or
repeated view, leaves `super` opaque, and now generates five files totaling
122,114 bytes under a 124 KiB cap. Raw source records remain the rewrite and
unknown-content authority.

A changed edit raw-patches only the selected layout scalars. Cache ownership is
rooted, not package-wide: the adapter rejects the deprecated document fields
11 and 12, follows `TP.DocumentArchive.super` field 15 to
`TSA.DocumentArchive.view_state` field 5, resolves the unique referenced
type-210 shared view-state object, follows its field 1 root reference, and
resolves the unique referenced type-10147 `TP.ViewStateRootArchive`. Both
followed local edges require one aggregate metadata occurrence and optional
unique field metadata at `[15, 5]` and `[1]`. When that
root has layout-state field 1, the transaction removes that field and the
selected layout-state identifier from its aggregate reference metadata and,
when present, its unique field metadata at path `[1]`. It preserves UI-state
field 2, unknown view-state fields, the detached opaque layout-state object,
unrelated reference metadata, and detached or unrooted type-10147 candidates.
Missing, duplicate, or contradictory objects along the rooted chain, a
layout/UI alias, selected merge/diff state, or noncanonical object-length
framing fail closed. The document and rooted view-state root may share one
component or occupy two, so changed diagnostics report one or two touched IWA
components; the intermediate type-210 bridge is not rewritten.

The same atomic flat-package reassembly deletes any root `preview.jpg`,
`preview-micro.jpg`, and `preview-web.jpg`; those ZIP deletions are reported
separately and are not components. Every other retained ZIP record, unselected
component, object header, message, and unknown field remains preservation-owned.
Bounded canonical unknown protobuf groups are readable and exact no-ops retain
them, but changed layout splicing currently fails closed on a group-bearing
document payload rather than attempting group-aware rewriting.
The changed candidate fully reopens under the retained limits and verifies the
requested layout, absent layout-state edge and previews, unchanged package
statistics, and unchanged section semantics before publication.

An exact semantic no-op preserves layout presence, view-state caches, previews,
and every source byte, reports zero components and zero deleted previews,
shares the source allocation, and skips reassembly and candidate reopen.
Applying a changed patch reopens and verifies the exact target artifact stored
in the patch instead of reassembling it; applying the inverse restores the
complete original artifact, including the former layout-state edge and preview
members. Legacy nested `Index.zip` inputs remain readable and admit exact
no-ops, but changed publication fails closed because the source is not an exact
flat reassembly authority.

The migration host's eager-Prost `PagesEditor::page_layout` and
`set_page_layout`, their private `editor::page_layout` module and source, and
the old host example are removed. The replacement example opens a focused
package, stages a validated layout, publishes without clobbering through a
sibling temporary file, and can emit an exact inverse artifact. The boundary
policy ratchets the two retired host method declarations, module/source
topology, and physical vocabulary in the new facade. This is one Pages read
and mutation owner transfer, not a manifest-edge retirement or deletion of the
remaining Pages host editor.
The checked inventory remains 64 workspace packages, 235 internal dependency
declarations, and 14 ordered migration debts.

The full `litchi-pages` test/doctest gate passes 92 cases; the focused
page-layout transaction suite passes 10/10 and the private codec suite passes
6/6. The Pages package check, no-dependency library Clippy with warnings denied,
all 63 boundary-policy unit tests, and the live Pages-specific audits pass. The
focused fuzz binary compiles and completed 32 generated smoke inputs plus a
fixed changed-layout corpus. A sanitizer-backed `cargo fuzz` campaign did not
start because the active stable toolchain rejects `-Zsanitizer=address` and no
nightly toolchain is installed, so sustained sanitizer fuzzing remains open.

The checked-in native Pages fixture transaction changed US Letter portrait to
792 by 612 point landscape, touched two components, deleted all three root
previews, and retained semantic text. Its source and Rust-candidate SHA-256
values are
`21107bc9323fba6f1589152454c0b0b0cc8e239313c6a369bc4a891116601b42`
and `79e00545ef6e2e30e366e3160b7d9126bf06cffac5fbbd5551e3d3789cc298e4`;
inverse application restored the exact source hash, and an equal-layout run on
the Rust candidate retained its exact hash.

Apple Pages 14.4 (7043.0.93) opened that candidate without a warning, repair,
recovery, or conversion. The inspector showed Any Printer, US Letter,
Landscape, 11.00 by 8.50 inches, and Document Body; all three fixture body
lines rendered exactly. Native Save As, close, and exact-file reopen repeated
those results and regenerated all three root previews. The native-resaved
SHA-256 is
`8228e7518bb080bd8e5ec134d0abc7484c8825ad3cde3d16cabf76c5dbd8ef82`;
a focused equal-layout transaction reported zero components and preview
deletions and reproduced that hash exactly.

Open work remains deliberately bounded. This vertical does not own the opaque
layout-state object, other Pages settings or render caches, a whole-graph
Buffa conversion, a durable serialized patch format, or a library-level atomic
durable filesystem replacement. It also makes no aggregate transaction
peak-memory or total-work claim across retained artifacts, rewrite buffers,
hashing, reassembly, and full candidate reopen, and a complete transitive
fallible-allocation proof remains open. Exact source bytes remain on the
ordinary `Package` surface, and the flattened `PageLayout*` transaction names
remain migration debt against the focused-module short-name rule.

## 2026-08-10 amendment: combined Pages document-settings ownership

`litchi-pages` now owns the combined document-formatter and footnote-formatter
state that Pages stores in one `TP.SettingsArchive`. The archive-free
`document_settings::Settings` composes the existing `document_options::Options`
and `footnote::Settings`. Its canonical focused module re-exports the short
`Edit`, `Commit`, `Patch`, `Diagnostics`, `Error`, and `LimitKind` names;
The new `Package::{document_settings, edit_document_settings,
apply_document_settings}` method and focused type signatures expose no native
identifiers, component names, message types, generated types, raw fields, or
source artifacts.

The private rooted owner is `TP.DocumentArchive.settings` field 7, not a field
inside `TP.SettingsArchive`. The adapter resolves the unique object-1,
type-10000 document payload, requires a nonzero local field-7 reference with
one aggregate metadata occurrence and optional unique field metadata at path
`[7]`, then resolves exactly one referenced type-10012 `TP.SettingsArchive`.
Strict raw preflight and forced Buffa lazy views cross-check the root reference
and the presence-preserving settings fields: body 1, headers 2, footers 3,
hyphenation 9, ligatures 10, footnote kind 30, format 31, numbering 32, gap 33,
and facing pages 34. Booleans, signed `int32` values, framing, singularity, and
the aggregate byte/field/work/nesting budget are canonical before Buffa can
authorize the result. Future enum integers remain visible through canonical
archive-free `Unknown` variants; noncanonical `Unknown` wrappers that shadow a
known value are rejected by the semantic model.

This expands and supersedes the page-layout-only size record for the shared
Pages body projection. The read-only body/layout/settings closure now generates
five files totaling 174,682 bytes under 176 KiB, with deterministic aggregate
SHA-256
`7618a60db84b87e28eea67a8acd85ce8eb19513cf4cee7654c1c4e78f405f824`.
Build ratchets reject repeated views and production encoding. The projection
still leaves document `super` opaque and never owns preservation; exact caller
bytes and header-preserving raw splices remain authoritative.

A changed commit patches only the ten selected settings scalars, rewrites the
settings-owner component, and applies the already documented rooted layout-
cache invalidation plus deletion of root `preview.jpg`, `preview-micro.jpg`,
and `preview-web.jpg`. The settings owner and cache root may share one component
or occupy two; diagnostics report that one-or-two component count and the ZIP
deletions separately. Canonical unknown scalar fields are retained exactly.
Canonical unknown groups are readable and exact on no-op paths, but changed
settings splicing fails closed on a group-bearing payload. Selected merge/diff
metadata, noncanonical object framing, ambiguous ownership, or stale cache
metadata also fail closed.

An exact semantic no-op returns before cache ownership inspection, shares the
source allocation, preserves optional presence, caches, previews, and bytes,
and performs no reassembly or reopen. Changed publication fully reopens and
verifies the stored settings, absent cache edge and previews, and unchanged
section semantics. Patch application requires the exact complete source;
changed apply reopens the retained target, conflicts on replay/tamper/competing
targets, and inverse application restores the complete source artifact.
Legacy nested `Index.zip` sources remain readable and admit exact no-ops, but
changed publication now returns `UnsupportedSource` instead of preserving the
old host's normalization behavior.

The migration host's `document_options`, `set_document_options`,
`footnote_settings`, and `set_footnote_settings` methods are deleted with three
implementation files: `document_options.rs`, its `wire.rs`, and
`footnote_settings.rs`. The two host examples and duplicate host CRUD tests are
replaced by one focused combined-settings example with immutable chaining,
no-clobber sibling-temporary publication, and optional exact inverse output.
The boundary ratchet covers all four methods, both module declarations, all
three sources, and native/protobuf/source-byte leakage from the focused facade.

The final gate passes all 108 Pages tests and doctests, including 14/14 focused
transaction cases, 4/4 strict codec cases, and 6/6 facade cases. Package check,
strict no-dependency Clippy, strict documentation, and all 70 boundary-policy
regressions pass; the live repository boundary command still reports only the
14 unrelated pre-existing soapberry-zip/xml-minifier diagnostics. The focused
fuzz binary compiles and its exact no-op and changed smoke inputs pass. A
sanitizer campaign remains open because the installed stable toolchain rejects
cargo-fuzz sanitizer flags and nightly is unavailable.

Apple Pages 14.4 (7043.0.93) passed the current writer gate on a fresh
app-authored source containing a real footnote. Source, Rust-candidate, and
native-resaved SHA-256 values are
`9da01e2805459e05450551827140069eefe8049aeeacc7625d3c62d7e00ffeab`,
`3d052e7f1ec86e57ea0553e46f628de1d9fa5bdda615ded9410fca29c93f0995`,
and `803167e2479c459f9a33c8ecfc4d713f596fdc5d5d337090ab3c90e467a0cba6`.
The Rust edit reported changed, touched two components, deleted three previews,
and inverted to the exact source. Pages opened, saved as, closed, and reopened
without warning; it showed body/header/footer and facing pages enabled,
hyphenation and ligatures disabled, Roman footnotes restarting each page with
an 18-point gap, and retained all three body markers plus the note text exactly.
Native Save As regenerated all three previews. Focused same-settings readback
on the native artifact was a byte-exact zero-component/zero-deletion no-op, and
its inverse retained the same native hash.

This moves one combined read/mutation vertical without retiring a manifest
edge; the inventory remains 64 packages, 235 internal declarations, and 14
ordered debts. Remaining shared debt includes aggregate transaction peak-memory
and total-work accounting, the infallible retained-`ArchiveInfo` clone in the
archive encoder, complete fallible-allocation proof, canonical group-aware
changed splicing, exact streaming/partial-output accounting, library-owned
atomic durable filesystem replacement, and a versioned deterministic patch
envelope with read/write sets, composition, merge, and bounded history. Exact
source bytes also remain on the ordinary `Package` surface. The opaque cache
object and other Pages settings/render state remain outside this vertical.

## 2026-08-10 amendment: hardened Keynote show-settings ownership

This amendment supersedes the 2026-08-08 show-settings topology and its
temporary flattened transaction names. `litchi-keynote::show` now owns the
archive-free `Settings` plus canonical short `Edit`, `Patch`, `Commit`,
`Diagnostics`, `Error`, and `LimitKind` types. `Package::{show_settings,
edit_show_settings, apply_show_settings}` is the focused read/edit/apply
surface. `Edit::set` consumes and returns the edit, so immutable chaining is
explicit. Those method and type signatures expose no native identifier,
component/member name, generated message, raw field, or source byte slice.
Callers publish the returned immutable `Package` through bounded
`Package::write_to`; exact transaction artifacts remain private.

The physical owner is resolved from the unique root `Document.iwa`, object 1,
and its selected `KN.DocumentArchive` message. Required show reference field 2
must be local, occur exactly once in the selected message's aggregate
references when nonzero, and, when field metadata exists, have one matching
path `[2]` entry and no competing path. A nonzero reference must resolve in
exactly one component to one object containing exactly one selected
`KN.ShowArchive` message. A zero reference reads as `Settings::default()` and
admits only an exact no-op because the transaction does not create/register a
new native owner.

Both ownership hops use strict raw preflight followed by forced private Buffa
lazy views and a complete equality cross-check. The root projection forces the
full three-field show reference after validating the required opaque document
base. Its five generated files total 58,630 bytes under 60 KiB with aggregate
SHA-256
`7918aad2578cf3bd07eb0be36f2e31d11f93391584308c1e4adc1fd86ed065fd`.
The show projection validates the complete known Show/SlideTree envelope and
slide-reference ceiling, routes the repeated slide tree by hand, and forces
theme, size, stylesheet, optional references, and all eight optional scalar
settings without retaining slide identifiers. Its five generated files total
138,661 bytes under 140 KiB with aggregate SHA-256
`747fe9f99dc5bb1855aae1bfcb16065a5fe6305bdbf8730a21ef24bb75e915ee`.
Both build ratchets forbid generated repeated views and production encoding;
raw source records retain rewrite and unknown-field authority.

Reads accept the bounded canonical ownership projection without imposing
mutation-only publication rules. A changed edit additionally requires
canonical object-length framing for every selected component and rejects
`should_merge`, base-message, diff/merge-version, diff-field-path,
fields-to-remove, and diff-read-version state on the selected root and show
messages. It raw-splices only size field 4 and scalar fields 6, 8-11, 15-16,
and 18 in the unique Show message, rewrites exactly one IWA component, fully
reopens under retained limits, and verifies settings, ownership, framing,
package structure, and unchanged content.

Cache invalidation follows semantics rather than treating every settings edit
as rendering. A size or slide-number-visibility change deletes the existing
root `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg` entries (zero to
three), while leaving all slide components and slide-node thumbnail/playback
caches exact. Playback-only settings changes preserve those root previews and
all slide caches byte-for-byte. Diagnostics report one touched component and
the actual root-preview deletion count separately. An exact semantic no-op
shares the source allocation and skips cache inspection, reassembly, and
reopen. Changed patch application authorizes exact source bytes and reopens its
retained target; inverse application restores the entire exact source artifact.

Legacy nested `Index.zip` remains readable and supports an exact no-op, but a
changed edit now returns `show::Error::UnsupportedSource`. This intentional
Preserve-policy break deletes the old normalizing compatibility writer rather
than silently changing physical provenance. The host
`KeynoteEditor::{show_settings, set_show_settings}`, its `mod show_settings`,
`keynote/editor/show_settings.rs`, `examples/edit_keynote_show.rs`, and direct
editor mutation tests are removed. The focused `edit_show_settings` example
uses the consuming semantic edit, optional exact inverse, and `write_to` with
no-clobber temporary publication.

This is direct editor-mutation retirement, not all Show ownership. The host's
read-only `KeynoteDocument::show` still decodes a Prost `KN.ShowArchive`, and
other Keynote creation, slide, transition, media, soundtrack, and graph paths
still use host/generated representations. No manifest edge or ordered debt is
retired by this amendment.

The current deterministic evidence includes 19/19 focused show-settings
transactions, 106/106 full codec tests, 49/49 focused Keynote codec tests,
Keynote all-target checking, `litchi-iwa` library checking, umbrella Keynote
facade compilation, strict rustdoc, and 80/80 boundary regressions. The focused
retirement and public-leak live audits are empty; the general repository
boundary run still has the 14 unrelated pre-existing `soapberry-zip` and
`xml-minifier` diagnostics. The show-settings fuzz target passes `cargo check`;
its stable-built executable completed 32 bounded cases with the expected
missing-sanitizer-symbol warnings. A cargo-fuzz sanitizer run still requires
nightly, which was unavailable.

Apple Keynote 14.4 (7043.0.93) completed two native gates from source SHA-256
`f3adcde9315b6df580805bcb63c995cc1e1ef569a4befa06a102485e13c883b2`.
The pristine Rust slide-number candidate was
`6d28d461c1203f00384fe6a758df1f903c7555b90ff02d2dc32d856aa9056c13`;
native Save As, close, and exact-path reopen produced
`031a701040ed1ea9a5111fe3e298bcddcf33d498891f827b703d01328ba17224`.
The pristine Rust 1280-by-720 candidate was
`67e9ff0557683af105dfe57f999acabcde23f121f7aebb06102c93e03121c027`;
its native resave was
`a3a2f6e072db4bd952f2c02e528f25c3656dba5810fbff75e93b5a699aac0eda`.
Each Rust inverse restored the exact source. Both candidates opened without
repair, recovery, or conversion, auto-played, and retained Self-Playing, Loop,
Play on Open, five-second transitions, and two-second builds; the inspectors
showed Widescreen 1920-by-1080 and Custom 1280-by-720 respectively. Exact-path
reopen preserved those values and auto-played again. All four
`Index/Slide*.iwa` hashes remained exact from each pristine Rust candidate
through native resave. Rust deleted all three root previews and Keynote
regenerated them on resave.

Keynote normalized `slide_numbers_visible = Some(true)` to absence during the
first native resave. Restaging absent is an exact no-op at the `031a7010...`
hash, while restaging true is a change. The 1280-by-720 native artifact's
same-settings no-op and inverse were both exact at `a3a2f6e0...`. This is
native evidence for slide-cache preservation and conservative root-preview
invalidation, not evidence that Keynote persists the slide-number scalar.

Remaining debt includes the host Prost `KeynoteDocument::show` reader and
other Show/graph consumers, aggregate transaction peak-memory and total-work
accounting, a complete transitive fallible-allocation proof, canonical
group-aware changed splicing, stable versioned patch serialization with
semantic operations/read-write sets/composition/merge/history, and
library-owned atomic durable filesystem replacement. `write_to` provides
bounded exact output and sink-offset errors but does not flush, sync, rename,
or make publication durable. A full sanitizer-backed fuzz campaign remains a
verification gate rather than an architectural claim.

## 2026-08-10 amendment: Numbers names ownership

`litchi-numbers` now owns atomic sheet/table naming through the canonical
nested `names::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` family
(with semantic `Path` and `InvalidReason`). `Package::edit_names` creates an
infallible, allocation-free empty batch; consuming `rename_sheet` and
`rename_table` stages operations against the immutable base snapshot, and
`Package::apply_names` applies the exact-source reversible patch. The root
`litchi::numbers::names` facade reexports that focused owner without flat
aliases. Its signatures expose neither native identities nor generated/wire
types or source slices. `Package::source_bytes` is crate-private and callers
stream exact output with `Package::write_to`.

The changed owner graph starts at `Index/Document.iwa`, object 1, whose field-1
local references select the rooted sheet sequence. Each selected object must
contain exactly one `TN.SheetArchive` or `TN.FormBasedSheetArchive`; ordinary
sheet name field 1 or nested form `super.name` is the sheet owner. A table
follows the selected sheet drawable at `[2]` or form path `[1, 2]` to one
canonical/legacy TableInfo, then its required local field-2 reference to one
canonical/legacy TableModel, whose required field 8 owns the display name and
field 1 supplies the stable table identity. Every followed local reference
must occur exactly once in aggregate metadata and, when field metadata is
present, exactly once at the expected path. Selected table models must have
one competing-root-free rooted TableInfo owner; detached/unselected objects
remain outside this vertical.

Names are decoded by strict raw preflight and forced private Buffa lazy views.
The projection borrows the required sheet name, forces nested form `super`, and
cross-checks both required TableModel identity/name strings without allocating
owned text. Its five generated files total 82,641 bytes with deterministic
aggregate SHA-256
`944b7637fd6bf0eb895174b1e9229aa9eb9c393e05c666a86dd2843792eefe3e`.
Raw field records retain preservation and rewrite authority.

Batch semantics are final-state atomic: sheet names are unique workbook-wide,
table names are unique within each sheet, swaps and collision-away batches are
valid, and duplicate targets or final collisions fail before publication.
Changed table renames refuse an interactively locked selected table, any
rooted pivot owner in the workbook, and rooted nonempty volatile
sheet/table-name dependencies. Sheet-only renames remain allowed when an
unselected table is locked. Changed planning is conservatively pre-bounded,
including the native quadratic table dependency scan, before native work.
Every touched IWA component is rewritten once and the complete candidate is
reopened under retained limits.

All accepted non-name fields, messages, objects, components, ZIP records, and
`Index/ViewState.iwa` remain exact. Every changed batch deletes the existing
zero-to-three root `preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg`
entries, with component and preview counts reported separately. Exact semantic
no-ops share the source, preserve previews and ViewState, and skip changed-only
framing, cache, lock/dependency, reassembly, and reopen work. Changed patch
application reopens its stored target after exact artifact/state checks; the
inverse restores the complete source, including previews.

Canonical and form sheets plus the accepted canonical/legacy TableInfo and
TableModel message variants remain supported when their rooted graph is
unambiguous. Legacy nested physical packages remain readable and exact for
no-ops, but a changed rename returns `names::Error::UnsupportedSource` rather
than normalizing provenance.

The host `NumbersEditor::{rename_sheet, rename_table}`, their direct tests,
and `litchi-iwa/examples/rename_numbers_items.rs` are deleted. The focused
`litchi-numbers/examples/edit_names.rs` owns selector-based batch staging,
optional inverse verification, and synced no-clobber publication through
`write_to`. The private `rename_attached_table_in_package` helper remains for
Numbers sheet duplication, and its `rename_table_in_package` wrapper remains
for Pages and Keynote table workflows; this retires the Numbers editor surface,
not those shared internal mutation primitives. No crate edge is removed, so
ordered debt 015 (`litchi-iwa -> litchi-numbers`) remains. The inventory stays
64 packages, 235 internal dependency declarations, and 14 ordered debts.

Verification passes 10/10 focused names tests, 105/105 `litchi-numbers`
library tests, the 1/1 root facade test with `--features numbers`, 89/89
boundary regressions, both live Numbers names/host audits,
`litchi-numbers --all-targets` checking, `litchi-iwa --lib` checking, and
strict rustdoc. Host `litchi-iwa --all-targets` is not claimed because
unrelated examples remain red. The fuzz target builds on stable and its
control-flow executable completed eight bounded runs with expected
missing-sanitizer-symbol warnings; this is not an ASan campaign.

Apple Numbers 14.4 (7043.0.93) opened the Unicode candidate without warning,
repair, recovery, or conversion. Source SHA-256 was
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`;
the Rust candidate was
`22f8bc21223317318ec23ec764b8998af77a2c7800c68cbe88351abdb26b6e56`,
and its public inverse restored the exact source. Numbers displayed sheet
`Líneas 你好 🧪`, table `表 Café №42`, the exact B2 marker, and B3 value 42.
Save As, close, and exact-path reopen succeeded and produced native SHA-256
`e1803b0568454a345f7962c5b4c72e8cb3d78adb2c87d5db1e6c58288a9413c4`,
with all three previews regenerated. Equal restaging, no-op application, and
its inverse were all byte-exact at that native hash.

The ordinary Unicode source was unlocked: its table was selectable/editable
and renamed successfully. A separate locked protection oracle, SHA-256
`eb2e29c97c415c1b61ed1f8fe766e7211ed386c825c32dec056b72c9398d3e09`,
reported `Locked` and `Locked items cannot be edited` through Numbers
accessibility state; cells were disabled, Unlock was enabled, and the Edit
table-title action made no change. The focused API likewise rejects a locked
table rename while permitting a sheet-only rename.

Remaining debt includes the conservative native Θ(T²) pivot preflight despite
its up-front work bound, aggregate transaction peak-memory and total-work
accounting, a complete fallible-allocation proof, stable versioned patch
serialization with semantic operations/read-write sets/composition/merge and
history, and library-owned atomic durable filesystem replacement. Patches
retain complete source/target artifacts process-locally; `write_to` does not
flush, sync, rename, or make publication durable. A sanitizer-backed fuzz
campaign also remains open.

## 2026-08-10 amendment: Keynote transition host retirement

The earlier slide-transition ownership record is superseded by the canonical
nested `transition::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}`
transaction family alongside the existing archive-free transition semantics.
`Package::{slide_transition, edit_slide_transition,
apply_slide_transition}` remains selector-first and exposes no native identity,
component, generated message, raw field, or source bytes. Exact package output
uses `Package::write_to`.

The changed owner is no longer authorized by semantic position alone. The
private package proves the rooted Show/SlideTree reference at path `[3, 2]` to
the selected SlideNode and its required local field-2 reference to the selected
SlideArchive. Every followed nonzero local edge must be unique in the component
catalog, occur exactly once in aggregate reference metadata, and have only
optional unique matching field-path evidence. The selected node and slide each
contain one expected typed message and must agree between the strict raw
projection, semantic record, and transition/node-marker views.

The rooted-owner audit walks the Show's slide-node list once and resolves each
node through the package's sorted, globally unique object index. Its lookup
cost is `O(slides log objects)`, and aggregate node-message plus local-reference
payload bytes are charged to `LimitKind::WireWork` rather than receiving a
fresh per-node allowance.

The selected transition and node marker are decoded only after strict raw
preflight, then cross-checked through five private Buffa lazy-view messages.
The 2,347-byte derived schema is provenance-checked against the canonical KN
fields, contains no repeated projection, and has no production encoder. Its
generated closure is five files/208,052 bytes under the 224 KiB ceiling;
validated caller-owned raw records remain the preservation and splice
authority. One field counter and one strict-plus-Buffa work counter are shared
across the selected SlideArchive, transition, attributes, and animation
envelopes, so nesting cannot reset either budget.

Changed publication additionally rejects selected object/message merge, base,
and diff state and requires canonical framing for every rewritten component.
Only the selected SlideArchive transition envelope at field 4 and, when effect
presence changes, the selected SlideNode `hasTransition` field 7 enter the
mutation closure. Node and slide may share one component or occupy two; each
touched component is decompressed and rewritten once. Full reopen verifies the
requested semantic transition, marker agreement, rooted ownership, and exact
locality. Unselected messages/objects/components, unknown fields, ZIP records,
the three root previews, `Index/ViewState.iwa`, and slide/node playback caches
remain exact. Transition mutation is playback-only and therefore does not use
the rendering transaction's root-preview deletion policy.

`Edit::clear` stages Keynote's modern no-effect representation while retaining
the supported timing/seed semantics. Clearing a slide whose transition is
already absent is idempotent and publishes an exact no-op rather than failing
or synthesizing an owner. All semantic no-ops share the source and skip
reassembly/reopen. Changed patch application authorizes exact bytes and reopens
its retained target; inverse application restores the complete source.
Legacy nested packages retain reads and exact no-ops, but changed transition
publication returns `transition::Error::UnsupportedSource`.

The host deletes `KeynoteEditor::slide_transition`,
`set_slide_transition`, and `clear_slide_transition`; the
`transition_lifecycle` module and
`keynote/editor/transition_lifecycle.rs`; and the three direct mutation
examples `clear_keynote_transition.rs`, `edit_keynote_transition.rs`, and
`set_keynote_transition_effect.rs`, together with their direct lifecycle/CRUD
tests. The exact cut is three methods, one module/source, three examples, and
five whole mutation tests; the host transition scope changes by +120/-998
lines, net -878. The focused `edit_slide_transition` example is the supported
mutation workflow.

This retires direct editor mutation, not all host transition vocabulary.
`KeynoteSlideInfo.transition` and slide-reading helpers remain;
`transition_wire.rs` specifically remains for `KeynoteEditor::slides()`
aggregate decoding and no-op validation, while creation uses the separate
`creation.rs::transition()` helper and retained `create_keynote_transition`
workflow. No manifest edge is removed, so ordered debt 014
(`litchi-iwa -> litchi-keynote`) remains. The current topology is 64 packages,
235 internal dependency declarations, 14 `litchi-iwa` dependency declarations,
and 14 ordered debts.

The final gate passes 8/8 focused `slide_transition` integration tests, 79/79
Keynote library tests, 6/6 Keynote doctests under warning-denied rustdoc, 7/7
root-facade tests with `--features keynote`, 6/6 transition-codec tests, and
the host transition conversion and reader suites at 3/3 and 7/7. The common
exact-artifact/batch substrate passes its focused 10/10 and full 140/140 tests
plus strict library Clippy; the archive exact-artifact gate reports 79 unit
and 2 integration tests. `cargo check -p litchi-keynote --all-targets` and
`cargo check -p litchi-iwa --lib`, the host no-run gate, formatting, and
focused diff checks pass. The boundary regression suite passes 101/101. All
fuzz bins check; the generated no-op,
fixed clear, and fixed set executables completed six bounded runs each. Those
stable-built smokes emitted the documented missing-sanitizer-symbol warnings
and are not an ASan campaign.

Apple Keynote 14.4 (7043.0.93) exercised disposable copies of source SHA-256
`ab186d8d59c858e1b3c2596fd45463cec75ddd92e9fda9032da656a940e68dca`.
The reproducible pristine Magic Move and clear candidates were respectively
`d5d24386cb544374f4c26da4349f7be961be34180a4536578616886a56af8c1a`
and `5235a3d03dbabced6d06a03b4873826da8602d97f478c61f6467b35d732a08e5`;
each public inverse restored the exact source. Both opened without warning,
repair, recovery, or conversion. Before Save As and after close/exact-path
reopen, the first showed Magic Move, 2 seconds, Automatic, and a 2.25-second
delay; the second showed No Transition Effect while retaining Automatic and
the 2.25-second delay.

Native Save As produced Magic SHA-256
`dda5049cf431b5c88ea0a9fb209c67edc0d7f0764c23a17eb4e9fdf947d786f6`
and clear SHA-256
`784069ca8bd2729829bcf204cccdced93f7fbea2b5f8c6b3e4965b47ef423e94`.
Focused equal restaging on each native artifact reported `changed=false` and
`touched_components=0`; output, comparison, and the no-op inverse were exact
at the respective native hash. Remaining debt is the shared aggregate
peak-memory/total-work and complete fallible-allocation proof, process-local
complete-artifact patches without stable semantic serialization/read-write
sets/composition/merge/history, library-owned durable atomic publication, and
a sanitizer-backed fuzz campaign. `write_to` itself does not flush, sync,
rename, or make publication durable.

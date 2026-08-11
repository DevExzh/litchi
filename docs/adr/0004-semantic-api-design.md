# ADR 0004: Semantic API design

- Status: Accepted
- Date: 2026-07-31

## Names and types

Presentation-level Keynote settings use the focused
`litchi-keynote::show::{Mode, Settings}` module. `Settings` stores validated,
private dimensions and exposes checked setters for playback delays and mode;
the IWA adapter cannot publish a non-finite size or delay through the ordinary
semantic API. Unknown native mode values remain lossless, while values already
assigned to a named mode are rejected as non-canonical rather than silently
rewritten.

Public types use short names inside focused modules. Prefer
`chart3d::BarShape` to `Chart3dBarShape`. Type tags plus unrelated optional
fields are replaced by non-exhaustive data-bearing enums:

```rust,ignore
pub enum Shape<'a> {
    Auto(Auto<'a>),
    Picture(Picture<'a>),
    Table(Table<'a>),
    Chart(Chart<'a>),
    Diagram(Diagram<'a>),
    Ole(Ole<'a>),
    Group(Group<'a>),
    Connector(Connector<'a>),
    Unknown(Unknown<'a>),
}
```

The same rule applies to fills, transitions, plots, axes, fields, records, and
other sum types. `Unknown` retains bounded lossless content. Common read methods
live on the enum or narrow static-dispatch traits; the facade does not use boxed
trait objects.

Numbers table header and footer configuration uses the focused
`litchi_numbers::table::headers::{Count, Settings}` module. `Count` is a
one-byte, `NonZeroU8`-backed value with the checked native `1..=5` domain;
`Option<Count>` therefore preserves presence without an extra storage byte.
`Settings` contains only archive-free optional counts and Boolean fields plus
effective-value helpers. Its non-exhaustive typed error rejects zero, overflow,
and out-of-range counts. Native protobuf presence, canonical wire framing,
unknown-field preservation, table-bound validation, and transaction readback
remain private to the IWA adapter. Pages and Keynote consume these canonical
short names directly, with no format-prefixed compatibility aliases.

Pages header/footer roles follow the focused-module rule at
`litchi_pages::header_footer::{Template, Kind}`. These one-byte enums contain
only semantic page-template and region roles; IWA keeps native object lookup,
storage metadata, protobuf decoding, and package mutation at the boundary.

Pages formatter values use focused modules and short names:
`litchi_pages::section::{Start, PageNumbering, PageNumber}`,
`litchi_pages::page_layout::{Layout, Orientation}`,
`litchi_pages::document_options::Options`, and
`litchi_pages::footnote::{Kind, Format, Numbering, Gap, Settings}`. Unknown
native discriminants remain lossless but cannot shadow named values, page
geometry validates finite positive dimensions and non-negative margins before
construction, and layout/options presence is packed into compact values. The
IWA side retains only native field mapping, protobuf validation, opaque fill
payloads, discovery/package identifiers, and transactional publication; the
former flat `Pages*` formatter aliases are removed.

Pages body-footnote values use the focused
`litchi_pages::footnote::body::{Footnote, Position, Selector}` module. A
`Footnote` contains its checked UTF-16 body position, bounded text, and
optional bounded custom marker; its native reference, contained storage, and
marker identifiers are deliberately absent. Body-footnote reads and edits use
`Selector::At` or `Selector::Index`, resolve ambiguity or absence as adapter
errors, and publish only after staged wire and graph validation.

Keynote build values use the focused `litchi_keynote::build` module. `Settings`
and `Effect` expose typed start relationships, bounded unknown text, finite
effect parameters, and boxed path/node collections; failed setters validate
the candidate before mutation. Native object identifiers, archive fields, and
raw direction/delivery integers are not part of this leaf's ordinary API.

The dependency-free Numbers formula vocabulary is an intentional naming
exception. `litchi-numbers::formula` is consumed by Numbers, Pages, and Keynote,
so `FormulaExpression`, `FormulaCellReference`, and the other `Formula*` names
remain explicit when re-exported from format facades; this prevents collisions
between otherwise identical cross-format concepts without restoring a flat
monolith. The compact reference and UUID structs are copyable value inputs, and
the archive compiler validates their table bounds and formula resource budgets
before emitting wire nodes. The formula compiler also validates arity for the
known fixed-arity functions; recognized functions without arity metadata and
unknown functions fail closed as typed parse errors.

Numbers table axis sizing uses the focused
`litchi-numbers::table::dimension::{Dimension, Points, Size}` API. `Points`
rejects zero, negative, infinite, and NaN values before construction, while
`Size::Default` remains distinct from `Size::Points(_)` so native default
sentinel semantics are not conflated with an explicit override. Archive
bounds, header-bucket storage, wire preservation, and transactional
publication remain in the IWA adapter; the former flat semantic owners are
removed rather than duplicated.

The same focused-module rule applies to the dependency-free iWork text leaf:
`litchi_iwa_text::font::{Font, Name}` owns the shared font identity, while
`NameError` reports its bounded validation failures. Format crates may expose
contextual aliases at their archive boundary, but they do not duplicate the
allocation-bearing model or publish a flat `TextFontName` implementation in
each application owner. `Name` validates before allocating borrowed input and
stores exactly one boxed identifier; no unchecked font-name constructor exists.

Text-frame columns use the adjacent focused
`litchi_iwa_text::columns::{Columns, Count, Gap, Width, Equal, Following,
Variable}` module. `Count` rejects zero and values above the explicit 256-column
budget, gaps reject non-finite and negative (including negative-zero) values,
and widths reject non-finite or non-positive values. `Variable` requires at
least two columns and stores the following widths and gaps in one bounded boxed
slice. Native `ColumnsArchive` presence, protobuf conversion, and package
mutation remain in `litchi-iwa`; no archive type or facade-wide error enters
the semantic leaf, and the former `TextColumn*` names are removed rather than
aliased.

Numbers cell display formats use the focused
`litchi_numbers::cell::data_format` API. `DataFormat` is the typed sum over
checked number, currency, percentage, scientific, fraction, numeral-system,
date/time, duration, checkbox, star-rating, slider, stepper, pop-up, text, and
custom values. Child modules validate finite/range/bounded text inputs before
allocation and use boxed slices only where the semantic value is inherently
variable-sized. The IWA adapter alone maps native format-table identifiers,
control-cell metadata, custom UUID registries, BNC scalar state, protobuf
fields, and transactional package changes. The old `TableCell*` semantic
owners and facade aliases are removed; Pages and Keynote use the Numbers leaf
types directly.
Rich-text storage uses the focused `litchi_iwa_text::storage` module. `Storage`
contains only owned UTF-8 text and validated semantic byte ranges; `Run` exposes
range geometry rather than a native style or object identifier, and `Fragment`
borrows text without allocating. Empty runs remain in the validated run slice
for lossless semantic retention but do not produce fragments. Out-of-bounds,
overflowing, or non-UTF-8-boundary ranges return typed errors before
publication. Protobuf decoding, UTF-16/native boundary conversion, archive
lookup, and unsupported wire-field preservation stay in the IWA adapter.

Pages document state follows the same raw/semantic split:
`litchi_pages::document::{Root, Body, Document}` owns an immutable, bounded
semantic snapshot, while `litchi-iwa` decodes native root and body payloads
before constructing it. The semantic model exposes borrowed section views and
exact-name or typed `litchi_core::Position` selection; duplicate names are a
typed ambiguity instead of an arbitrary first match. Every section can produce
an unambiguous snapshot-local position selector, while a producer-visible name
is preferred when one is actually present. The native Pages adapter does not
synthesize section names from headings when the current schema projection has
no name, and native object identifiers and protobuf messages do not appear in
ordinary signatures. Snapshot cloning shares the semantic section allocation
and never reparses or mutates the source document.

Keynote uses the same selector contract through
`litchi_keynote::SlideSelector::{Name, Position}`. A slide's developer-facing
navigator name is distinct from its visible title; duplicate exact names are a
typed error and a checked source position remains the deterministic fallback.
The first production Buffa archive-header adapter is below both facades:
generated lazy views, decode contexts, and codec errors remain private, while
the existing neutral compatibility structs and semantic values form the crate
boundary.

The same focused-module rule applies to neutral visual values:
`litchi_iwa_common::color::{RgbColorSpace, Rgba}` owns the validated RGBA
model and `color::Error`. The value has no protobuf or archive dependency, stores
only fixed-size channel data, and rejects non-finite or out-of-range channels
before a caller can publish it. Concrete IWA modules retain only native color
conversion and map the leaf error at their archive boundary.

The same focused-module rule applies to table appearance:
`litchi_iwa_common::table::appearance::{Appearance, Banding, RowSizing,
GridlineVisibility, Gridlines}` owns the compact semantic value. The common
module uses short names in its table context and stores no archive state;
native style inheritance and protobuf conversion remain in the IWA adapter.
Contextual `Table*` aliases are migration adapters for the concrete facade,
not duplicate value owners.

Table-cell text layout uses the same focused module:
`litchi_iwa_common::table::cell::layout::{TextWrap, VerticalAlignment, Inset,
Insets, Layout}`. `Inset` is a four-byte transparent value constructed only
through finite, non-negative validation, and `Layout` is a fixed-size
composable value. Native alignment and padding conversion remain outside this
leaf; the facade only adapts its typed error at the archive boundary.

Table hidden-axis semantics use the focused
`litchi_iwa_common::table::axis::{AxisIndex, HiddenAxes}` module. `AxisIndex`
is a zero-based row-or-column value, while `HiddenAxes` validates duplicates
and stores canonical row-then-column ordering in one boxed slice. Its typed
duplicate error is independent of archive state. Native hidden-state UUIDs,
protobuf field mapping, graph traversal, bounds validation, and transactional
package mutation remain in the IWA adapter, and concrete Numbers, Pages, and
Keynote APIs consume the common values without contextual compatibility aliases.

Shape and ordinary text-box frame layout uses a separate focused module:
`litchi_iwa_common::text::layout::{VerticalAlignment, AutoSize, Inset, Insets,
Layout}`. It deliberately does not reuse table-cell `TextWrap` semantics: shape
autosizing and four-way frame alignment are independent values. All five values
are copyable and heap-free; `Inset::from_points` returns a typed, allocation-free
error for non-finite or negative input. Native protobuf conversion, bounded
style inheritance, and package transactions remain in `litchi-iwa`, and the
facade exposes the common module through `litchi_iwa::text::layout`.

Media classification uses the focused `litchi_iwa_common::media::Type`
module. The five variants are a one-byte, copyable value with no archive or
filesystem state; extension matching is ASCII-case-insensitive and signature
sniffing borrows only the supplied byte prefix. Unknown bytes stay `Unknown`,
while an unrecognized ISO-BMFF brand remains conservatively classified as
`Video` to match iWork replacement semantics. Archive discovery, asset
metadata, limits, and replacement validation remain in the IWA adapter.

Media movie/audio playback uses the focused archive-free
`litchi_iwa_common::media::playback` module. `MediaVolume` is a compact,
validated linear multiplier; `MediaLoopMode` maps named native values while
retaining genuinely unknown discriminants; and `MediaPlaybackSettings` uses
consuming builders plus checked trim canonicalization. The semantic module has
no archive, protobuf, graph, package, or IWA-error dependency. Native duration
decoding, legacy/modern loop reconciliation, unknown-field-preserving wire
patches, and transactional publication remain in `litchi-iwa`, while all three
iWork consumers use the common owners directly and the old facade owners are
deleted rather than aliased.

Keynote slide movie classification uses the focused
`litchi_keynote::slide::media::MovieKind` value. Its compact, non-exhaustive
variants (`File`, `Audio`, `Placeholder`, and `LiveVideo`) contain no
archive, package, or native media identifiers. Movie creation uses the adjacent
`litchi_keynote::slide::movie::Options` value, which validates finite placement,
strictly positive displayed and natural dimensions, and a positive duration in
the native finite `f32`-seconds domain. `litchi-iwa` retains
`KeynoteSlideMovieInfo`, graph-aware CRUD, native media identifiers, and the
mapping from native movie flags to this product value.

Keynote slide-audio creation uses the focused
`litchi_keynote::slide::audio::Options` value. The fields are private, and
construction validates finite placement plus a positive duration that fits
the native finite `f32`-seconds domain; accessors expose the canonical point
and duration without archive state. The IWA adapter retains audio graph
discovery, native identifiers, zero-size geometry, wire-preserving mutations,
package transactions, and the IWA-owned info/removal values. Shared
`MediaPlaybackSettings` optional fields, loop discriminants, and volume
validation remain format-neutral IWA playback semantics until their own
cross-format extraction.

Shape paths use the focused `litchi_iwa_common::shape::path` module. Its
`Preset`, `CornerRadius`, `PolygonSides`, `StarPoints`, and `InnerRadiusRatio`
names are concise in their path context, fixed-size where scalar, and
validated before they can enter a public preset. The star control stores an
inner-to-outer radius ratio in `[0, 1)`; it is not an archive-owned absolute
radius. Structural `ShapePathKind`, native path-family decoding, natural-size
constraints, and protobuf/wire mutation remain in `litchi-iwa`, so the common
leaf stays dependency-free and allocation-free.

Drawable geometry uses the adjacent focused
`litchi_iwa_common::shape::geometry::{Point, Size, FlipAxis}` module for its
fixed-size neutral values. The IWA-only `DrawableGeometry` aggregate retains
optional native field presence, reflection flags, rotation conventions, and
wire conversion; those protobuf details do not leak into the common leaf.

Shape fills keep their neutral gradient vocabulary in the focused
`litchi_iwa_common::shape::fill` module. `Kind`, `Angle`, `StopPosition`,
`StopMidpoint`, `Opacity`, `Stop`, and `Gradient` use typed validation, fixed
scalar storage, `color::Rgba`, and boxed stops without archive or protobuf
state. `ShapeFill` remains the IWA boundary aggregate because image fills and
native data references are format-specific; the former `ShapeGradient*`
semantic owners are removed rather than retained as aliases.

Shape line endpoints and chart kinds use the same focused common-value rule:
`litchi_iwa_common::shape::line::{Endpoint, Endpoints}` and
`litchi_iwa_common::chart::kind::Kind` are compact, lossless, archive-free
inputs. Native endpoint inheritance, field numbers, protobuf conversion, and
wire-preserving mutation remain private to `litchi-iwa`.

Chart axis controls use the focused `litchi_iwa_common::chart::axis` module.
`Axis::{Category, Value}` is the compact semantic selector shared by all three
iWork owners, and `TickMarkLocation` models the exclusive formatter choices
with an explicit `Unsupported(i32)` case for future native values. Native
integer conversion, archive lookup, shared-object ownership, and protobuf
patching remain in `litchi-iwa`; the common values stay copyable and free of
package state.

The child modules keep the axis vocabulary contextual and short:
`axis::bounds::{Bound, Bounds}`, `axis::label_angle::LabelAngle`,
`axis::label_position_3d::LabelPosition3d`, `axis::scale::Scale`, and
`axis::steps::{MajorStepCount, MinorStepCount, Steps}`. Their constructors
return module-owned typed errors for finite/range validation; unknown native
integer values remain explicit in the enum variants rather than being silently
mapped to defaults.

Chart number formatting uses the focused
`litchi_iwa_common::chart::number_format` vocabulary:
`FixedDecimalPlaces`, `DecimalPlaces`, `NegativeStyle`, `NumberFormat`, and
`LabelAffixes`. Fixed decimal places are checked at construction, the packed
`NumberFormat` occupies one byte, and `LabelAffixes::new` returns a typed error
before accepting more than its bounded UTF-8 budget. Affixes expose borrowed
prefix/suffix views over one allocation. There is no ambiguous generic format
default: callers select the explicit axis or series native default because
those defaults differ on thousands separators. Native field IDs, protobuf
decoding, dual-field conflict checks, and wire-preserving patching remain
outside the semantic module.

Chart series orientation uses the focused `litchi_iwa_common::chart` module's
four-byte `Direction` value. `Rows` and `Columns` are ergonomic named
constants, while `DirectionKind` projects recognized values and every other
native integer is preserved losslessly. The common value owns only this
lossless integer conversion; protobuf field mapping, archive lookup, and
mutation remain in the IWA adapter. The long `ChartSeriesDirection` name is
removed rather than retained as a facade alias.

Keynote transition semantics use the focused
`litchi_keynote::transition::{Settings, AnimationParameters, CustomParameters}`
API. The module owns memory-conscious opaque semantic payload containers and
retains `Effect` plus the existing `Direction`, `MosaicType`, `Acceleration`,
and `TextDelivery` scalar values; its constructors enforce bounded ownership,
finite numbers, NUL-free text, and canonical semantic values. No raw native
IDs or archives leak upward. `litchi_keynote::Package` owns selector-first
read/set/clear transactions for existing modern slide-transition envelopes and
their exact-source-checked reversible patches. A private Buffa lazy view
projects known native fields after strict bounded wire preflight. Validated raw
records remain authoritative for lossless patching; graph lookup, opaque
payload validation, and retained-options candidate verification stay private
to the package boundary. The legacy `litchi-iwa` Keynote compatibility APIs
remain available and are not claimed to be removed.

Chart series value-label selectors use the focused
`litchi_iwa_common::chart::series_labels::{Visibility, Index}` module.
`Visibility` is a compact two-state value with boolean conversion, while
`Index` is a copyable zero-based series-position value. Visibility does not
provide an ambiguous common default: native defaults are chart-family-specific
(pie is visible; other supported families are hidden). The IWA adapter retains
chart-kind field selection, generated extension decoding, canonical boolean
validation, sparse default insertion/removal, unknown-field preservation, and
package mutation. Unsupported chart kinds remain typed failures, and the former
`ChartSeriesValueLabelVisibility` and `ChartSeriesIndex` owners are removed
rather than retained as aliases.

Pie and donut label settings use the focused
`litchi_iwa_common::chart::pie` vocabulary. `LabelVisibility` packs the two
independent label toggles into one byte with explicit native defaults;
`LeaderLineVisibility` retains the signed native integer so unknown future
states can be read, compared, and written without information loss. The IWA
adapter owns field 31/44 and field 102 decoding, canonical varint checks,
lossless unknown-field patching, and the styled versus geometry-only series
allocation boundary. Concrete Numbers, Pages, and Keynote APIs consume these
short semantic values directly.

Chart category-label settings use the focused
`litchi_iwa_common::chart::category_labels::{Interval, Frequency, Layout}`
module. `Interval` validates the explicit native range before construction;
`Frequency` distinguishes hidden, automatic, all, and custom labels while
retaining unknown signed native intervals losslessly; and `Layout` composes
that frequency with final-category visibility. The IWA adapter owns native
field mapping, strict int32/boolean validation, unknown-field-preserving wire
patches, axis visibility, and package transactions. The old long semantic
owners are removed rather than retained as aliases.

Chart reference-line settings use the focused
`litchi_iwa_common::chart::reference_line` module. `Value` rejects non-finite
custom positions, `Kind` keeps known calculations distinct from checked
future native kinds, and `Line` uses an optional bounded name plus packed
visibility state so default labels do not allocate. The focused IWA facade
module exposes `Line`, `Kind`, and `Value`; raw extension messages and graph
identifiers are not part of ordinary CRUD signatures.

PresentationML implements this rule as `litchi-pptx::shape::{Scene, Shape}`.
`Scene` is a bounded semantic index over one slide-like owner, not a vector of
detached XML allocations. Shapes are visited in depth-first pre-order, while a
`Group` lends only its direct children so hierarchy is never inferred from raw
non-visual IDs. The ordinary selector is an exact producer-visible name:

```rust,ignore
let Some(title) = scene.get("Title")? else {
    return Ok(());
};
let fourth = scene.at(3)?;
```

The numeric form is a checked secondary path for source-order repair and import
algorithms. Ordinary lookup represents a missing name as `None`, while strict
`shape` lookup provides a typed not-found failure. Duplicate exact names and
out-of-range positions are typed errors in either applicable path; none of
these operations implements `Index` or panics. Native non-visual IDs remain
diagnostic/reference metadata rather than the primary facade selector.

An MCE-free scene borrows the caller's owner bytes. If bounded
Choice/Fallback processing must rewrite the owner, the scene owns that one
processed buffer. In either case each shape lends a checked byte span from the
shared owner, so indexing does not allocate one XML buffer per shape. Decoded
names and text may use a bounded compact arena; “borrowed” therefore describes
the source payload and shape views, not a claim of allocation-free parsing.

Shape-owned programmable tags reuse this same selector through
`tag::shape::{load, put, remove}`. The focused package layer accepts the
containing owner plus a `shape::Key`; it does not make callers rediscover a
native shape ID, relationship ID, or tag-part name. Name selection therefore
has the same strict not-found/ambiguity behavior as `Scene::shape`, and checked
depth-first positions remain the deliberate repair path. The lower layer keeps
package topology explicit without leaking that topology into the semantic
selector.

The migration facade composes the two semantic catalogs directly:

```rust,ignore
let current = package.shape_tags("Overview", "Status badge")?;
package.put_shape_tags("Overview", "Status badge", replacement)?;
package.remove_shape_tags(0_usize, 3_usize)?;
```

An already-resolved slide shortens the read path to
`slide.shape_tags("Status badge")`. These conveniences preserve the same typed
selection and transaction rules; they do not introduce a second ID-based API.

Validated semantic scalars include `Length` (canonical EMU storage), `Percent`,
`Angle`, `Row`, and `Column`. Internal const-generic bounded integers support
domain aliases. Constructors return `Option` or typed `Result` and are usable in
constant evaluation. There are no unchecked constructors.

Use `Default` and struct update syntax only where every resulting combination is
valid. Otherwise use a consuming builder or data-bearing enum. Typestate is
reserved for high-value safety boundaries and must not spread generic noise
through normal CRUD.

Names with an Office-defined domain use focused checked types such as
`xlsx::sheet::Name`, while short verbs continue to accept borrowed strings and
validate them internally. Owned strings and prevalidated names move into edit
plans without another payload copy; borrowed or owned checked names are also
ordinary lookup selectors. A case-preserving name type uses the
document format's identity semantics for equality and hashing; spreadsheet
sheet names therefore use canonical, locale-independent Unicode caseless
identity rather than process locale or ASCII-only comparison.

Human-readable semantic selectors are the primary facade. Spreadsheet columns,
for example, accept A1 labels such as `"B"`; reusable checked `Column` values and
raw zero-based indexes remain concise secondary forms for import and numeric
algorithms. Invalid syntax and out-of-grid coordinates are errors, while a
missing catalog object is `None`; selectors return `Result` or
`Result<Option<_>>`, never indexing panics. The same policy applies to cells,
ranges, rows, sheets, and other developer-facing collections; native
relationship IDs, part names, and physical style indexes stay below the facade.

Properties with constrained wire domains use short types in focused modules.
Bubble charts are the concrete DrawingML precedent: `chart::bubble::Size` is an
enum for the exclusive `area`/`w` wire domain, while
`chart::bubble::Scale` is a `repr(transparent)` `u16` newtype whose checked
constructors admit only the inclusive `0..=300` schema range. `Scale`'s inner
value and `BubbleTypeGroup`'s scale and size fields remain private, so safe
client code cannot construct an invalid bubble scale or size-representation
state, and the writer does not repeat late string or range validation. The
group exposes the concise
`scale`/`set_scale`/`with_scale` and `size`/`set_size`/`with_size` families.
Reading validates both domains before constructing the group: numeric XML is
checked before conversion to `Scale`, and size tokens are matched exactly by
`Size::from_xml`. Writing maps `Size` directly to a borrowed static token, so
domain-to-wire conversion performs no allocation. Unit tests cover scale
boundaries and exact size-token conversion; the public integration tests cover
typed builder access, writer/reader round trips, and rejection of an
out-of-range scale or unknown size token.

SpreadsheetML borders follow the same rule through the focused
`xlsx::styles::border` module. `Line` replaces the open-ended style string;
`Rgb` is an exact four-byte ARGB value; `Tint` admits only finite values in
`-1.0..=1.0`; and `Color` distinguishes default, RGB, theme, indexed, and
explicit automatic values. `Side` composes one visible line with that typed
color. `Diagonal::new(Side, Dir)` makes a complete authored diagonal, while a
private partial representation retains side-only or direction-only
schema-valid producer states. `Border` is the single model shared by the style
parser, worksheet facade, `CellFormat`, and writer, including physical or
Strict logical edges, inside edges, and the outline setting. Absence is
`Option<Side>` rather than a contradictory
`Line::None`; there are no `BorderStyle`, `CellBorder`, or `CellBorderSide`
compatibility aliases. Exact line tokens map to borrowed static strings,
unknown tokens fail parsing, incompatible edge conventions fail writing, and
full-value resource equality plus resolved cell-format keys prevent hash
collisions from aliasing distinct borders.

SpreadsheetML cell alignment follows the same rule in the sibling
`xlsx::styles::alignment` module. `Horizontal` and `Vertical` are closed enums;
`Rotation`, `Reading`, and `Indent` are compact checked scalars; and
`Alignment` exposes public typed fields so ordinary authoring remains concise
with `..Alignment::new()`. The complete modeled alignment value participates in
shared-XF equality and therefore cannot disappear during resource
deduplication. Microsoft's context-dependent rotation value is represented
explicitly and rejected when writing Strict SpreadsheetML rather than leaking
an unexplained integer or string through the facade.

PresentationML modern-comment completion uses `Progress`, a private
`NonZeroU32` offset representation for the inclusive Office range
`0..=100_000` thousandths of one percent. `Progress::new` accepts an ordinary
whole percentage, `from_thousandths` is the precise lower-level constructor,
and `Option<Progress>` remains four bytes. Parsing accepts only the specified
percentage lexical forms and Office's numeric form; writing emits canonical
numeric thousandths without allocating a temporary string.

WordprocessingML section options use separate types where two visually similar
wire domains are not interchangeable. `ChapterSep` closes the page-number
separator domain; `Footnotes` and `Endnotes` carry distinct `FootnotePos` and
`EndnotePos` enums so an endnote cannot be assigned a footnote-only position;
and `BorderColor::{Auto, Rgb([u8; 3])}` replaces hexadecimal strings without a
heap allocation. Page-border artwork is a closed `PageBorderArt` value carried
by `PageBorderStyle::Art`, not an arbitrary token. Copyable option structs use
public fields and `Default` only where struct-update syntax cannot create an
invalid domain value.

DrawingML preset geometry is one shared closed vocabulary in
`litchi-drawingml::geom`. `Preset` contains all 187 `ST_ShapeType` values and
`TextPreset` contains all 41 `ST_TextShapeType` values from the checked-in
ECMA-376 Strict and Transitional schema archives. Both are one-byte enums,
convert exact tokens to borrowed static strings, and return a compact typed
error for an unknown token. DOCX, XLSX, and XLSB consume these same types;
there is no format-local partial enum, `Custom(String)` escape hatch, or
allocation in preset-to-wire conversion. Whether an object owns a text-box
story is separate from its preset geometry because `textBox` is not an
`ST_ShapeType` value.

Worksheet shapes use `xlsx::Geometry::{Preset, Custom}` from both parser and
writer. This makes competing `a:prstGeom` and `a:custGeom` states
unrepresentable after parsing, while the parser rejects duplicate or competing
elements. The large, comparatively cold custom-geometry payload is boxed so
every ordinary preset shape does not carry its size; `From` conversions and
the semantic shape constructors hide that storage choice. Borrowing and
move-out accessors avoid copying the custom payload.

PresentationML universal time offsets use
`litchi-pptx::time::{Offset, Unit}` rather than lexical strings. `Offset`
implements the complete `[MS-PPTX]` decimal grammar with bounded input,
retains values exactly as canonical decimal milliseconds, and defines
equality, hashing, and ordering by represented duration. Consequently `1s`
and `1000ms` are one semantic bookmark time rather than two distinct strings.
Short integral constructors and checked decimal parsing cover ordinary
authoring; exact conversion to `std::time::Duration` is available only when
the value has nanosecond precision and fits that type.

SpreadsheetML page setup lives in the focused `xlsx::page_setup` module.
`Orientation`, `Order`, `Comments`, `ErrorMode`, and `Unit` close the token
domains; `Paper`, `Scale`, `Fit`, `FirstPage`, `Copies`, and `Dpi` are compact
checked numbers with Office-specific reserved ranges; `Measure` retains exact
positive-universal-measure decimals; and `Setup` has public typed fields for
concise struct-update authoring. Options keep absent attributes distinct from
explicit defaults. The mutable facade uses `set_page`, `page`, and
move-returning `remove_page`; `set_fit` atomically sets both dimensions and the
independent fit-to-page policy. The immutable worksheet also exposes one
`page` view backed by the complete typed parser. Printer-settings relationship
IDs are deliberately absent from public `Setup`: the dedicated
`xlsx::printer_settings` graph API owns and validates them, so ordinary page
authoring cannot create a dangling package relationship. The earlier raw
numeric/boolean read model and string setter are deleted rather than retained
as compatibility paths.

SpreadsheetML text and conditional formatting do not expose schema tokens as
strings. Cell fonts use `styles::{Underline, Scheme, Script}`; worksheet state
uses `writer::Visibility`; and `conditional_formatting` owns the compact
`Kind`, `Operator`, `ValueKind`, `Period`, `Direction`, `Axis`, `ColorRole`,
`IconSet`, and `IconSet14` vocabularies. Core and Office 2010 icon-set types are
separate, preventing an extension-only icon set from entering a core writer.
The reader, conditional-format writer, sort model, and workbook facade share
those types. Spreadsheet colors reuse the checked four-byte `styles::Rgb`
value, including tab colors, instead of accepting arbitrary hex strings.
Unknown values in closed domains fail parsing. The old worksheet validation
and conditional-formatting facade that cloned typed values back into strings
is removed; `Worksheet::data_validations` and
`Worksheet::conditional_formattings` expose the complete typed models
directly.

WordprocessingML follows the same rule. `numbering::NumberFormat` contains the
complete fixed number-format vocabulary, `MultiLevelType` models numbering
structure, `settings::NotePosition` models footnote/endnote placement, and
`settings::CompatFlag` models the complete Transitional compatibility-flag
set while identifying the Strict subset. These enums are compact, copyable,
and have exact codecs; invalid fixed tokens are errors. Free-form numbering
text, style identifiers, compatibility-setting extension triples, and custom
number-format strings remain strings because their value spaces are not
closed.

Presentation media separates presence, time, geometry, and payload ownership.
Trim/fade values remain optional typed `time::Offset`s so absence is not
collapsed into zero, and a slide-show seek event is `Seek { at: Offset }`,
making the required time part of the variant. Media offsets use checked
`drawingml::coord::Coordinate`; extents use the integer-only inclusive
`coord::Extent` domain from `ST_PositiveCoordinate` (whose lower bound is
zero). `MediaData` hides shared immutable storage behind slice access and a
move-first recovery path, so cloning a resource does not copy its bytes.
Bounded canonical `p:extLst` content is retained as inert XML rather than
discarded or interpreted as executable markup.

Keynote soundtrack playback uses the focused archive-free
`litchi_keynote::soundtrack::{Mode, Settings}` API. `Mode` preserves unknown
native discriminants losslessly and rejects known values disguised as
`Unknown`; `Settings` validates finite volume in the native `0.0..=1.0`
domain. The current native soundtrack schema has no string or time fields, so
the semantic owner does not invent filename or duration state. Protobuf
presence, media references, package IDs, unknown bytes, graph lookup, and
transactional edits remain in `litchi-iwa`, where changing playback settings
cannot reorder or rebuild the media collection.

The DrawingML diagram data model uses `Id::{Number(i32), Guid([u8; 16])}` for
the complete `ST_ModelId` union and the concise `Point`, `PointType`,
`Connection`, and `ConnectionType` names inside `diagram::data`. Transition and
presentation metadata lives in the enum variant that requires it. Semantic
CRUD rejects duplicate identifiers and parent conflicts and cascades removal
of dependent transitions. Publication validates the graph, XML characters,
and aggregate output budget before touching a caller's sink. This writer is an
explicit fresh/canonical modeled-subset API: a parsed part containing unmodeled
rich XML must not be serialized through it under a lossless-edit claim.

The standard permits `bubble3D` directly under `bubbleChart`, but desktop
Microsoft Excel rejects that placement as documented by MS-OE376 section
2.1.1458(b). The reader therefore accepts the standard form and projects its
semantic value onto the typed series state, while the writer emits only the
Office-compatible series-level element. This is a deliberate canonicalization
at a measured native-application boundary, not an unchecked string workaround.

The XLSX grid-property surface therefore uses `column::{Width, Props, State}`
and `row::{Height, Props, State}` with one shared checked `Outline` type. Widths
admit only finite Office widths, heights admit only finite Excel point heights,
and outline levels cannot exceed the supported depth. `Props` exposes
independent read-only facets, while `State` distinguishes an implicit grid
property from a stored record. Shared format identity is a lineage-checked
resource handle, not a public numeric ID. Mutation uses paired short verbs
(`hide`/`show`, `best_fit`/`fixed`, `collapse`/`expand`) plus checked set/reset
operations. Independently prepared edits may join when they touch different
facets of one row or column; two writes to the same facet conflict rather than
acquiring a public lock or choosing a last writer.

Row and column formatting uses the same opaque `Style` resource as cell
formatting through the short `style`/`reset_style` verbs. These operations name
the grid-default layer: an explicit local cell style remains a separate,
higher-precedence layer rather than being silently rewritten. A row style
derives its required custom-format marker. A new column style must share a
transaction with an explicit width, because Excel interprets the resulting
style-only column record as zero-width; the safe facade returns a typed block
instead of collapsing an implicit column. Existing column records retain their
effective width while their style is retargeted.

Worksheet-wide grid defaults live in the focused `xlsx::layout` module rather
than adding ambiguous long names to the crate root. `layout::{Height, Width,
Descent, Defaults}` is deliberately distinct from explicit `row::Height` and
`column::Width`: the wire domains and inheritance layers differ. A sheet
returns `Result<Option<&layout::Defaults>>`; absence remains observable because
the correct row height and column width can depend on fonts and producer state
and must not be guessed. The stored value exposes effective behavior, including
Microsoft's rule that an `x14ac:dyDescent` makes `custom_height` true even when
the core marker is false. The checked numeric wrappers are niche-encoded so an
optional descent occupies one machine word.

The paired `SheetEdit::defaults()` editor uses short named verbs for each
orthogonal facet and works unchanged on a transaction-local new sheet.
Materializing an absent `sheetFormatPr` requires `height` in the same
transaction, enforced by a typed error before bytes change; reset-only edits on
an absent record remain no-ops. Whole-record `remove` is explicit. Row-specific
descent uses the existing `row` selector and the same checked scalar. Compact
`layout::Fields` bitflags describe overlap, allowing independently prepared
default edits to join when their facets are disjoint while retaining
deterministic conflicts for the same facet.

Merged cells are a structural view over the sparse grid, not synthetic cell
records. `Sheet::cell` returns `cell::View::{Missing, Covered(Rect), Stored}`:
the merge anchor remains an ordinary stored-or-missing cell, while every other
coordinate in the range is `Covered` even if a producer left a physical
follower record behind. `Sheet::merges()` lazily borrows checked `Rect` values;
neither lookup nor traversal expands a range into its constituent addresses.
This extra enum is deliberate: returning `Option<&Cell>` cannot distinguish a
missing cell from a covered coordinate without encouraging callers to inspect
native IDs or XML.

The ordinary mutation verbs are `merge(area)` and `unmerge(at)`. The primary
inputs are A1 areas and lookup coordinates, with reusable checked ranges and
raw checked numeric forms remaining concise secondary inputs. `unmerge`
selects the containing range, so callers do not need to rediscover its exact
boundaries. A one-cell merge, overlap, protected sheet, multi-cell formula,
unknown compatibility owner, or unmodeled merge-container payload is a typed
error. Creating a merge that would hide follower content is also rejected; a
caller must explicitly clear, remove, or relocate that content in the same
transaction. Transaction ordering is structural removal, ordinary grid edits,
then structural creation, allowing both safe unmerge-and-edit and
clear-and-merge workflows without publishing an intermediate state.

Patch membership transitions use `merge::Change::{Add, Remove}` rather than two
independent booleans, so an impossible no-op transition cannot be constructed.
Merge changes participate in reversible patches and disjoint edit joins. Two
merge intents conflict only where their checked rectangles intersect; a merge
creation also conflicts with an independently planned follower content write.
Follower clear/remove and style-only effects remain composable because the
three-phase writer applies them before creating the merge.
Disjoint structural ranges move into one edit without public locks or wrapper
types. The low-level writer preserves untouched merge records, namespace
spelling, unknown attributes, and schema order; a safe facade never exposes
relationship IDs or physical merge record positions.

The iWork table adapters use the same semantic ownership rule. Their shared
`litchi-numbers::table::merge::Region` is a compact checked rectangle, and its
pure axis rebase algebra is independent of protobuf or package state. Native
IWA merge formulas, formula-store ordering, anchor payload movement, and
transactional publication remain private to `litchi-iwa`; Pages and Keynote
retain only their format-owned comment and merge sidecars.

The supported Numbers API does not expose BNC parsing, storage flags, native
format indices, generated protobuf values, or IWA object identifiers. Those
details live in `litchi-numbers-wire` and private package adapters. Semantic
cells, coordinates, selectors, tables, sheets, and documents are the ordinary
application vocabulary.

Pages section identity comes from the native section graph, never from a
heading. `TP.DocumentArchive.section` supplies an optional initial boundary at
UTF-16 position zero; ordered later boundaries come from
`TSWP.StorageArchive.table_section`. Each reference must resolve to exactly one
type-10011 `TP.SectionArchive`, whose field 26 is the producer-visible name.
The body is split at strictly increasing UTF-16 scalar boundaries. Every later
boundary requires a preceding U+0004 marker, and that marker is omitted from
semantic text. Names are copied verbatim, including the distinction between
absent and empty. Duplicate names are valid authored data, while exact lookup
returns typed ambiguity. Repeated references, malformed boundaries, missing
objects, and wrong or duplicate payload types reject ingress.

Keynote navigator identity is likewise distinct from visible title text, and
read selectors use an exact name or checked position without publishing a
native ID. Writing `KN.SlideArchive.name` is not yet a supported concrete
package transaction. A field-10-only prototype passed internal parse/readback
tests but caused real Keynote to render layout placeholder text and ignore the
requested label; the legacy editor exhibited the same failure. That API was
removed rather than publishing a transaction whose internal model disagreed
with the native application. A future name edit must identify the additional
producer metadata or graph mutation and pass native content-preservation gates.

Slide playback omission is a separate, supported semantic property.
`Slide::is_skipped` projects the required singular Boolean field 4 of the one
type-4 slide-node payload. Ingress rejects a missing, duplicate, wrong-wire,
noncanonical, or non-Boolean occurrence. Mutation starts from
`Package::edit()` and accepts an exact navigator-name or typed-position
selector through `skip_slide`, `include_slide`, or `set_slide_skipped`; it never
accepts a native node identifier. Because canonical `false` and `true` are both
one byte, a real edit changes exactly that payload byte while preserving the
object header, every other decompressed IWA byte, and every unrelated package
member. Full reopen and semantic readback precede publication. Native Keynote
open, save-as, close, and reopen verification remains a required gate for this
and future Keynote mutations.

Keynote text extraction follows the same strict ownership rule. The adapter
walks only the document-referenced show, ordered slide-tree nodes, slides,
placeholders, shapes, and speaker notes. A drawable must contain exactly one
recognized placeholder or shape owner, and its referenced storage must contain
exactly one schema-proven type-2001 storage payload. Type 2022 is not guessed
to be storage because native fixtures use that identity for incompatible
payloads. Valid protobuf bytes in an unrelated message are never guessed to be
text. Title and speaker-note slots
remain plain strings because that is their semantic model; body and other
drawable storages retain archive-free `Storage` fragment ranges. `Package::text`
then emits the title, visible body/drawable content, and speaker notes in the
same presentation order as `Show` and `Slide`, without exposing or sorting on
native identities.

`ReadOptions` combines the existing checked physical archive profile with a
checked `SemanticLimits` profile. Objects, slides, traversed graph references,
decoded text storages, retained fragment ranges, and aggregate semantic UTF-8
bytes have independent non-zero ceilings bounded by format-wide maxima. Focused
show/slide/build preflights enforce slide counts, used build/drawable reference
counts, and retained name/effect identifiers before those specific generated
vectors or semantic owners are materialized. They also require the document,
show, slide-node, slide, build, placeholder, shape, and note envelope fields
consumed by this adapter. Exceeded semantic and native-payload ceilings carry a
content-free semantic path plus observed and maximum counts. Duplicate object
identities, duplicate typed payloads, wrong wire kinds, invalid UTF-8,
ambiguous text owners, missing references, and malformed known payloads fail
before a semantic snapshot is published. Physical/object-index limits apply at
package construction; the remaining semantic profile applies lazily on first
semantic access or explicit validation. Complete allocation-envelope
preflights for every ignored nested generated field remain migration debt until
the larger Keynote graph moves to focused bounded projections.

Numbers deliberately exposes two table projections with different ownership
contracts. `Package::document()` and `Package::sheets()` are the ordinary
semantic view: they follow the canonical document sheet sequence and each
sheet's drawable sequence, reject duplicate ownership, and never publish a
detached table model. `Package::extract_structured_tables()` is an explicitly
allocating compatibility view for replacing the migration host's historical
structured extractor. It classifies objects by their first native message,
emits canonical type-6001 table models before compatible type-6000 models in
object-identity order, deduplicates candidates, and retains valid detached
models. The method name and documentation make that archive-wide behavior
visible without adding native IDs, generated messages, or low-level objects to
the supported signature. Detached compatibility tables do not acquire a fake
sheet or leak into selector-based ordinary APIs.

Bitflags represent small orthogonal settings, Roaring bitmaps represent large
sparse integer sets, enums represent exclusive states, and inheritance uses an
explicit tri-state. The facade exposes named operations rather than bit math.

## Views and formatting

Local formatting and resolved/effective formatting are separate views. Editing
always names a layer: local, named style, layout/master, or theme. Shared styles
are first-class immutable resources. Editing one reports its fan-out; forking a
style and retargeting a selection is a distinct operation.

Theme colors and fonts retain references and ordered transforms rather than
collapsing to RGB or resolved family names. Fill, stroke, effects, placement,
and chart types contain only valid settings. Coordinate spaces are typed where
mixing them would be unsafe, while the facade uses concise constructors.

## Errors and diagnostics

Each low-level crate owns a non-exhaustive typed error. The facade wraps sources
in a small stable kind taxonomy plus structured context frames: format,
part/stream, semantic object, record/XML location, and byte offset. Expected
failures do not become `Other(String)`.

Reusable queries are typed serializable ASTs; closure predicates remain a local
convenience. Traversal is lazy and borrowed. Durable bulk/thread selections are
lineage-checked `Selection<T>` handles.

## Legacy PPT reader-shape mutation migration

During the `0.x` API series, legacy PPT reader-shape mutators intentionally
migrate from unconditional in-memory mutation to fallible semantic mutation.
Methods such as text, fill, line, geometry, formatting, placeholder, picture,
and group setters return `Result<_, litchi_ppt::shapes::MutationError>`.
Detached values remain editable. Values decoded from an opened presentation
carry private source lineage and return `MutationError::SourceBound` before
changing state when no faithful source-checked transaction can publish the
operation.

This is an intentional pre-1.0 breaking correction: mutable fields that could
create invalid or silently unpublishable reader state become private and gain
immutable accessors. The public `shapes::Shape` trait remains externally
implementable and contains no package-lineage hook. Source binding and parser
hydration are crate-private companion infrastructure for Litchi's built-in
shape variants; third-party implementations are not required or permitted to
model package provenance through the semantic trait.

## 2026-08-08 amendment: Keynote slide-order semantics

Presentation order is a semantic property of a Keynote show, not a public
`SlideTree` or list of native references. Callers use
`Package::edit_slide_order()` and `SlideOrderEdit::move_slide`; the source is a
`SlideSelector` by exact navigator name or checked position, and the
destination is a typed `Position`. The destination denotes the moved slide's
final zero-based position in the base list. It is valid only when it is less
than the current slide count; moving to the current source position is a
byte-exact no-op rather than an error.

Resolution never falls back from navigator name to visible title text, never
publishes a node/object identifier, and never asks the caller for a component
or package member. A successful move preserves the number and identity of
slides and all attached semantic content. Internally, the transaction reorders
complete raw slide-reference field records in the one validated show envelope,
including each encoded key, encoded length, and nested reference payload.
Unknown and deprecated reference scalars remain attached, and all other show
fields and slide components remain untouched. The package is fully reopened and
the requested order is read back before the new immutable snapshot is
published.

This capability authorizes ordering only. It does not imply slide insertion,
duplication, deletion, navigator-name mutation, layout reassignment, or an
ordinary public raw-reference collection. Those operations retain their own
dependency-closure and native-application acceptance gates.

## 2026-08-08 amendment: focused Keynote presentation settings

Presentation settings are exposed as the singleton semantic value returned by
`litchi_keynote::Package::show_settings()`. Callers stage changes with
`edit_show_settings()` and the existing archive-free
`litchi_keynote::show::Settings`; they never select a native object, component,
package member, or generated message. The value covers checked presentation
size, optional slide-number visibility, looping, presentation mode, autoplay
transition and build delays, idle-timer activation and delay, and automatic
play-on-open. `Size`, `Seconds`, and `Mode` continue to enforce finite/domain
rules, preserve optional presence, and retain unknown future mode values
without permitting a named discriminant to masquerade as `Unknown`.

The direct reader validates the full known Show/SlideTree envelope and its
resource ceilings but does not initialize the full semantic slide cache or
retain slide-node identifiers. A null root show has the semantic default
settings. It cannot acquire a synthesized physical Show through this API, so
only an exact no-op edit is valid until creation owns identifier allocation and
component registration. Format-owned show-settings errors and limit kinds
remain content-free, and the public patch vocabulary contains semantic values
rather than raw identities. Changed legacy package normalization remains an
explicit host compatibility capability, not implicit behavior of the ordinary
focused transaction.

## 2026-08-08 amendment: Pages section-name semantics

A Pages section name is edited through
`Package::edit_section_name(SectionSelector)` rather than a native object ID or
low-level section table. `SectionSelector::name` is an exact producer-name
match and `SectionSelector::index` resolves to the checked public `Position`
stored by the transaction. Selection completes against the base snapshot
before mutation; missing names, missing positions, and ambiguous exact names
are distinct typed failures.

The semantic value is `Option<&str>`: absence and an explicitly present empty
name are observably different and round-trip independently. NUL is the only
format-level string invariant imposed by this slice. Assigning a duplicate
name is valid, while selecting that duplicate name later is ambiguous. Errors,
limits, patches, diagnostics, and `Debug` output omit authored names, native
identifiers, package members, raw bytes, and lower-layer diagnostic strings.

This surface authorizes replacement or removal of one existing section name
only. It does not imply section creation, deletion, ordering, body mutation,
identifier allocation, legacy package normalization, or a public collection of
raw references. Those capabilities retain separate dependency-closure and
native-application gates.

## 2026-08-08 amendment: Pages section-pagination semantics

Section pagination is a lossless semantic value selected through
`SectionSelector`, never through a section object identifier, component name,
or protobuf object. `Pagination` retains independent optional presence for
`Start`, `PageNumbering`, and `PageNumber`; therefore absent native fields stay
distinct from explicitly encoded defaults. `PageNumber` excludes zero, while
the two lossless enums preserve future native values without allowing a known
discriminant to be constructed through `Unknown`.

`Package::section_pagination` reads the value for an exact name or checked
position. `Package::edit_section_pagination` resolves the selector immediately
and retains only the public semantic `Position`; the staged editor can replace
the complete value, change one setting, or clear all three fields. Exact no-ops
share the immutable source allocation. Changed edits publish only after a full
retained-limit package reopen and semantic readback. Reversible patches keep
exact authorization artifacts private and expose only the position, semantic
before/after values, compact fingerprints, and content-free diagnostics.

This capability owns `TP.SectionArchive` fields 20--22 only. Header/footer
inheritance and first-page flags, section background, section names, template
references, section creation/deletion/order, and legacy package normalization
remain separate capabilities with independent preservation and native gates.

## 2026-08-08 amendment: Pages section-text semantics

Pages section text is selected through `SectionSelector` and returned by
`Package::section_text` as the exact semantic text owned by that section. The
value excludes the native U+0004 delimiter before a following section. Exact
names remain case-sensitive and ambiguity is a typed error; an index resolves
immediately to the checked semantic `Position` retained by the edit and patch.
The public model contains no global body-storage coordinate, raw object ID, or
native section-table entry.

`TextPosition` is a UTF-16 code-unit boundary. `TextSpan` is an ordered,
half-open pair of positions that may be empty, so the same type describes both
a replacement selection and an insertion point. Construction rejects indexes
outside the compact native domain and reversed endpoints; the Pages adapter
then rejects endpoints beyond the selected section or between a scalar's
surrogate pair. Byte indexes and native absolute body offsets remain private.

`Package::edit_section_text` stages exactly one unambiguous splice. `replace`
is the primitive; `insert`, `delete`, `set`, and `clear` are typed conveniences.
The source text stays borrowed until staging requires an owned replacement.
Reserved U+0004 section breaks, U+000E footnote anchors, and U+FFFC inline
objects cannot be synthesized or consumed by this capability. If an edit would
remove dependent reference metadata it fails with `DependentContent` rather
than silently deleting another semantic graph. `edit_body_text` is only a
single-section convenience and therefore cannot flatten a multi-section body.

A successful patch exposes its semantic section position, original span, and
complete before/after section text. Exact source and target artifacts stay
private; fingerprints are diagnostics and exact bytes authorize application.
An inverse swaps the artifacts, semantic precondition, and replacement span so
it can restore the original package byte-for-byte. Semantic no-ops share the
source allocation and report zero touched components. Changed edits publish
only after bounded full-package reopening plus section text, neighboring
section, object-count, and root/section-reference topology verification.

This surface authorizes one existing section-body splice only. It does not
create, remove, or reorder sections; delete footnote or attachment graphs;
edit headers, footers, floating text, or text boxes; normalize legacy nested
packages; change a no-root/fallback body whose physical ownership is not
rooted; serialize durable patches; or publish files atomically. Those remain
separate capabilities and migration gates.

## 2026-08-10 amendment: hardened Keynote show-settings semantic surface

This amendment supersedes the 2026-08-08 show-settings naming and compatibility
claims. The supported entry points are
`litchi_keynote::Package::{show_settings, edit_show_settings,
apply_show_settings}`, and the canonical focused family is
`litchi_keynote::show::{Settings, Edit, Commit, Patch, Diagnostics, Error,
LimitKind}`. Flat `ShowSettings*` transaction names are not re-exported from the
crate root. The public method/type signatures expose no native identity,
component/member name, generated message, raw field, source bytes, or retained
artifact accessor. The consuming `Edit::set` keeps immutable chaining explicit;
`Package::write_to` is the bounded exact-output seam, not patch-byte exposure.

The semantic value remains the singleton checked size and optional presentation
settings already described. A null root show reads as `Settings::default()`
but cannot be changed by this API because it does not allocate or register a
Show owner. Under the explicit Preserve policy, a physical legacy nested
`Index.zip` source remains readable and supports an exact no-op, but a changed
edit returns `show::Error::UnsupportedSource`. The former host normalization is
not compatibility behavior for this focused surface.

Accordingly, `KeynoteEditor::{show_settings, set_show_settings}`, the private
`editor::show_settings` module/source, the host `edit_keynote_show` example,
and direct editor mutation tests are deleted rather than retained behind a
shim. This retires the direct editor mutation API, not all Show reads: the
host's read-only `KeynoteDocument::show` still returns a Prost-backed
`KN.ShowArchive`, and other creation and graph consumers remain migration work.

`show::Patch` is an exact-source, reversible, process-local value that privately
retains complete source and target artifacts. It is not a compact or durable
patch encoding. ADR 0003's versioned deterministic serialization, semantic
operations and read/write sets, composition, three-way merge, and bounded
history remain deferred rather than implied by `inverse` or exact patch
application.

## 2026-08-10 amendment: focused Numbers names semantic surface

This amendment supersedes direct host rename methods and raw-ID examples. The
canonical focused vocabulary is
`litchi_numbers::names::{Edit, Patch, Commit, Diagnostics, Error, InvalidReason,
LimitKind, Path}`. Root aliases, glob re-exports, and prefixed transaction names
such as `Name*`, `Names*`, `SheetName*`, and `TableName*` are not part of the
surface. `Package::{edit_names, apply_names}` are the package entry points;
existing `Package::document()` projection supplies readback rather than a
second names getter.

`edit_names()` is infallible and `O(1)`. Consuming
`Edit::{rename_sheet, rename_table}` methods accept semantic sheet/table
selectors and exact UTF-8 names; they expose no native object identifier,
component/member name, generated message, raw field, source bytes, or retained
artifact accessor. Names must be nonempty and NUL-free. `Path` is a
content-free checked position path. Errors, diagnostics, patch `Debug`, and
public patch accessors redact authored names and lower-layer details.

All selectors resolve against the immutable base snapshot, and each selected
semantic owner may occur only once. Commit validates the simultaneous final
batch under one workbook-wide sheet namespace and one table namespace per
owning sheet, so a swap or collision-away batch is valid without introducing
order-dependent selection. Invalid names, duplicate targets, final-name
collisions, source ambiguity, selected table locks, rooted volatile
sheet/table-name dependencies, rooted pivot ownership, limits, verification
failures, and exact-source patch conflicts remain typed semantic errors.

The focused surface uses Preserve policy. It supports unambiguous canonical
and alternate legacy flat ownership encodings without promotion; a changed
physical legacy nested-`Index.zip` source is `UnsupportedSource`, while reads
and exact no-ops remain exact. Changed publication removes existing root
previews but preserves `Index`/`ViewState` and unrelated package content. The
patch is an exact-source reversible process-local value that privately holds
two full artifacts, not a serialized or durable editing protocol; callers
publish through the existing bounded package writer.

`NumbersEditor::{rename_sheet, rename_table}`, their direct mutation tests, and
`examples/rename_numbers_items.rs` are retired without aliases or shims. The
private `rename_table_in_package` helper remains solely for cross-format Pages
and Keynote table creation/edit flows. That internal dependency does not
authorize public raw-ID naming entry points or weaken this deletion gate.

## 2026-08-10 amendment: canonical Keynote transition transaction surface

This amendment supersedes the earlier transition compatibility paragraph. The
canonical public family is
`litchi_keynote::transition::{Settings, Edit, Patch, Commit, Diagnostics,
Error, LimitKind}` alongside the existing transition semantic value types.
`Package::{slide_transition, edit_slide_transition, apply_slide_transition}`
are the selector-first entry points. Flat `SlideTransition*` transaction
aliases, transition transaction root aliases or globs, and the root `Effect`
alias are removed; callers use the contextual `transition` module.

Exact navigator-name or checked-position selectors replace host numeric slide
indices. Public signatures expose no native identity, component/member name,
generated message, raw field, source bytes, or retained artifact accessor.
`Edit::settings` borrows the selected optional value; consuming `Edit::set`
replaces one existing modern envelope, while consuming `Edit::clear` stages the
modern no-effect value. An absent transition is readable as `None`, cannot be
synthesized by `set`, and makes `clear` an idempotent exact no-op. Transaction
errors and `Debug` output remain content-redacted; invalid semantic settings
retain their typed archive-free cause.

Changed admission is deliberately selected and focused rather than a general
eager parse of unrelated slides. It proves the rooted Show/SlideTree,
SlideNode, and SlideArchive chain, exact reference metadata, strict semantic
and marker agreement, canonical framing, and absence of selected merge/base/
diff state. Rooted ownership uses indexed `O(slides log objects)` lookups under
an aggregate `LimitKind::WireWork` charge, while one shared nested codec budget
governs fields and work across the complete transition projection. Only the
selected transition subtree and a conditionally changed node marker may differ;
one or two selected components are rewritten once and then reopened with exact
locality verification. No-op, exact apply, conflict, and inverse behavior
follows ADR 0003's immutable two-artifact contract.

The focused API writes existing unambiguous modern envelopes only. Legacy
database-field transition state remains readable and no-op-preservable but is
not promoted through a changed transaction. Physical legacy nested packages
remain readable and exact on no-op paths; changed publication returns
`transition::Error::UnsupportedSource`. `Package::write_to` remains the bounded
exact-output seam and does not turn the process-local patch into durable or
atomic publication.

`KeynoteEditor::{slide_transition, set_slide_transition,
clear_slide_transition}`, the lifecycle module/source, their five whole direct
mutation tests, and the three clear/edit/set-effect host examples are retired
without aliases or shims. The host's `KeynoteSlideInfo.transition` snapshot
field, slide readers, and `transition_wire.rs` remain for
`KeynoteEditor::slides()` aggregate decoding and no-op validation. Creation
separately retains `creation.rs::transition()` and the creation example. This
is therefore public legacy editor retirement, not deletion of all host
transition vocabulary or creation ownership.

## 2026-08-10 amendment: focused Numbers table-header transaction API

The existing semantic family remains
`litchi_numbers::table::headers::{Count, Settings}`. The focused transaction
family is nested separately as
`litchi_numbers::table::headers::transaction::{Edit, Patch, Commit,
Diagnostics, Error, LimitKind, Path, InvalidReason}`. Flat `HeaderSettings*`,
`TableHeader*`, `TableHeaders*`, or `TableHeaderSettings*` transaction aliases,
crate-root aliases, and glob re-exports are not part of the surface.

The package entry points are `Package::{table_header_settings,
edit_table_headers, apply_table_headers}`. Read/edit methods take semantic sheet
and table selectors, resolve exact names or checked zero-based positions against
the immutable base snapshot, and expose no native object identifier,
component/member name, generated/wire type, raw field, source bytes, or
retained artifact accessor. `Edit::settings` returns the compact
presence-sensitive staged value; infallible consuming
`Edit::set(self, Settings) -> Self` replaces it as one unit, and consuming
`Edit::commit` returns the immutable verified package, patch, and content-free
diagnostics.

`Path` identifies a checked semantic sheet/table position without names or
native identifiers. `InvalidReason` carries only the numeric row-section or
header-column capacity facts needed to explain a bound failure. The transaction
`Error` keeps selector failure, invalid source, unsupported physical source or
dependency, selected table lock, invalid settings, resource ceiling,
allocation, verification, and exact-source patch conflict typed and otherwise
content-redacted. `LimitKind` names the finite input/output, package-entry,
payload, reference, wire byte/output/field/nesting/work, and aggregate
transaction-work ceilings; allocation is a separate typed error.

Presence is semantic state: `None` differs from an explicitly encoded false,
while present counts are nonzero and at most five. Header rows plus footer rows
must fit declared rows and header columns must fit declared columns. A locked
table remains readable and admits an exact no-op, but a changed edit is a typed
`TableLocked` refusal. Changed publication rewrites one selected component,
deletes existing root previews, reopens under retained limits, and verifies
exact locality; no-op shares the source and skips changed-only work.

Ordinary and FormBasedSheet ownership paths accept one unambiguous
role-specific modern or legacy TableInfo/TableModel message and retain its
physical type. Mixed or duplicate role candidates are invalid. Under Preserve,
a physical legacy nested-`Index.zip` source remains readable and byte-exact for
an equal edit, while a changed edit is `Error::UnsupportedSource`.

`Error::UnsupportedDependency` is deliberately conservative. A valid
TableModel field-85 pivot owner blocks every changed edit. Header-row/column
count changes block on present fields 81/84/86, nonempty field 83, rooted
HeaderNameMgr state, selected TableInfo fields 4/5/7/8/15/17, or a true
TableInfo field 16. Footer changes block on nonempty field 83, active grouping
decoded through fields 81/86, selected TableInfo fields 5/15/17, or a true
TableInfo field 16; dependency references are exact, local, non-aliased
ownership proofs. Repeating-header changes block on deprecated sheet field 4.
Malformed,
duplicate, contradictory, or role-aliased dependency state is `InvalidSource`;
neither family is normalized or left stale.

Patches are exact-source, reversible, process-local values that privately hold
the complete source/target artifacts and exact selected source/target payloads
for a change. Changed apply verifies exact source settings and payload, charges
the source topology and distinct retained target bytes before target reopen,
then verifies the retained target payload and exact locality. Apply never
restages or merges a semantic edit; inverse swaps both artifact and payload
preconditions. `Package::write_to` remains the bounded exact-output seam and
does not make the patch serialized, compact, atomic, or durable.

The native count oracle changed
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`
to
`5c2323b509e5ea9a975b5f254bbd46cf42657aa1c3858d2c7e98f30f07e4b40c`
and demonstrated HeaderNameMgr/tile/CalcEngine work beyond TableModel. The
Boolean-only freeze-off save
`015568e6b922e80fbfb760491dc49994ccc2218356ed197131beb46c1bd75850`
and same-state native control
`df44ed7d0b12c1d372dad7ad7361ed1140d41967921ee42b71a4072b78615721`
preserved B2/B3 and counts and regenerated semantically equivalent ViewState
with different allocated references. That supports raw ViewState preservation
by the focused edit, not byte-stable native Save As output.

The retired public boundary is
`NumbersEditor::{table_header_settings, set_table_header_settings}`, their
direct mutation tests, the duplicate host count test, and the old host example.
Shared private attached-table primitives and the private Pages/Keynote package
bridges remain; no flat compatibility alias or shim replaces the deleted host
methods.

## 2026-08-10 amendment: focused Keynote placeholder visibility API

The canonical family is
`litchi_keynote::slide::placeholder::{Kind, State, Edit, Patch, Commit,
Diagnostics, Error, LimitKind}`. `Package::{slide_placeholder_visibility,
edit_slide_placeholder_visibility, apply_slide_placeholder_visibility}` are the
selector-first entry points. Flat `Placeholder*`, `PlaceholderVisibility*`,
`SlidePlaceholder*`, `SlidePlaceholderVisibility*`, or
`SlideTextPlaceholder*` transaction aliases, crate-root aliases, package-owner
exports, and glob re-exports are not part of the surface.

`Kind` contains `Title`, `Body`, and `SlideNumber`; `State` distinguishes
`Visible` from `Hidden`. Read returns `Result<Option<State>, Error>` so a
missing role remains distinct from an existing hidden placeholder. Edit returns
`Error::PlaceholderNotFound { kind }` when the selected role is absent;
infallible consuming `Edit::{set, show, hide}` can change only an existing role
and consuming `Edit::commit` returns the immutable verified package, patch,
and content-free diagnostics.

Public signatures expose no slide/node/placeholder object identifier,
component/member name, native reference list, field number, generated/wire
type, source bytes, or retained artifact accessor. Exact navigator-name or
checked-position selectors replace host numeric-index mutation, while edit and
patch `Debug` output remains redacted. Errors keep selector ambiguity/missing
position/name, missing placeholder, invalid or unsupported source, resource or
allocation failure, verification, and exact-source conflict typed without
leaking document content. `LimitKind` bounds package bytes/entries, slides,
references, and wire bytes/fields/nesting/work.

Visibility owns only direct per-slide participation of an existing selected
placeholder. Hidden state preserves its reference, graph, text, and storage;
visible state requires one occurrence in each selected ownership list. For
title/body, showing appends the exact role-reference bytes at the end of both
lists, as independent native title and body oracles established. Title/body
changes also delete existing root previews and invalidate the selected
SlideNode's rendered thumbnails before exact candidate reopen; no-op preserves
all caches. The distinct slide-number ordering and cache contract is specified
below.

The direction-aware node exact-delta verifier uses linear payload
occurrence/kind and metadata declaration indexes instead of repeated per-ID
rescans; 4,096 to 8,192 distinct references remain within 2.3 times measured
work. Work includes each `MessageInfo`, each `FieldInfo`, and each field-path
element even when metadata records are empty. A 4,096-empty-`FieldInfo`
regression proves atomic refusal under zero-work and payload-only allowances.
Full precharge includes selected and nonselected payload bytes; each
`MessageInfo` scalar/base record, version/diff vector, and diff/removal path;
each `FieldInfo` record, path, version vector, and feature identifier; and both
Work and References for every aggregate or `FieldInfo` reference occurrence.
Conditional invalidation charges this structure before broad validation and
also charges `header_length` before core replacement; exact verification
charges the same complete source and candidate structure before equality. A
256 KiB sibling plus 2,048-element reference/vector metadata regression proves
atomic low-allowance refusal for exact verification and mutation.
Conditional invalidation reuses one payload scan for any rewrite, while exact
verification neither clones the node nor reruns invalidation. Both consume the
transaction's remaining work/reference allowance and merge exact reports into
the same budget. Before allocating its output, the slide router charges exact
`6 * source.len() + output_len + 2 * parsed_fields` work.

This API does not select a layout placeholder, create or delete placeholder
graphs, edit slide styles, or absorb layout ownership.
Changed admission refuses style-level title/body visibility, a selected build
dependency, or selected direct-cache state rather than treating any of them as
authority for a wider mutation. Layout APIs, snapshot reads, creation, and the
shared private placeholder-ownership machinery remain outside this handoff;
the separate amendment below transfers only direct per-slide slide-number
visibility.

Patches are exact-source, reversible, process-local values that privately hold
the complete source and target artifacts plus compact semantic/owner/cache
preconditions. Apply reopens only the retained exact target; inverse restores
the complete source caches and previews. `Package::write_to` remains the
bounded exact-output seam and does not make patches serialized, compact,
atomic, or durable. Physical legacy nested sources remain readable and exact
for no-op paths but refuse changed publication under Preserve.

The completed retirement boundary is the three direct public
`KeynoteEditor` visibility setters, public `KeynoteSlideTextPlaceholder`, the
direct `placeholder_visibility` module/source and mutation tests, and
`set_keynote_placeholder_visibility.rs`. `KeynoteSlideInfo` title/body
visibility snapshot fields, `set_slide_layout`, and their retained private
helpers are not aliases or shims for this transaction. The separate amendment
below records retirement of `set_slide_number_visible` without retiring the
snapshot, layout, or creation surfaces.
Mixed layout reads migrated to `slide::placeholder::Kind`, and the focused
`SlideTextRole` title/body discriminator was removed rather than retained as a
parallel enum. Host `KeynoteSlideTextRole` remains the aggregate discriminator
for title, body, text-box, and shape reads; it is not an alias for this
transaction. The implementation, native, compatibility, and host-removal gates
enumerated in ADR 0003 passed; the title/body totals included 94/94 focused
library tests, 18/18 preview tests, and the unchanged 5/5 visibility tests.

## 2026-08-11 amendment: focused per-slide slide-number visibility API

`slide::placeholder::Kind::SlideNumber` is the canonical extension of the
existing visibility transaction. It deliberately reuses
`slide::placeholder::{State, Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` and the existing `Package` read/edit/apply methods; no flat
`SlideNumber*` transaction family, root alias, raw identifier selector, or
package-owner export is added. Missing field-20 ownership reads as `None` and
editing it returns `Error::PlaceholderNotFound`; the API operates only on an
existing layout-provided role.

The selector remains an exact navigator name or checked slide position. For
`Kind::SlideNumber`, `State::Hidden` requires absent/false SlideNode field 18
and no selected membership in SlideArchive fields 7/42; `State::Visible`
requires true field 18 and one selected membership in each. The stable role
reference remains SlideArchive field 20. Showing inserts its exact reference
payload after the last existing member of each native list; hiding removes only
that payload. The title/body append-at-message-end convention does not apply to
this role.

Per-slide visibility is not the show-wide preference:
`show::Settings::slide_numbers_visible` owns KeynoteShow field 6 and remains
byte-exact across this transaction. Nor does `Kind::SlideNumber` expose text;
the slide-text API returns
`SlideTextError::UnsupportedKind { kind: Kind::SlideNumber }`. This explicit
guard lets one semantic `Kind` select related operations without implying that
the numeric attachment character is ordinary writable slide text.

The changed-source contract proves the rooted node/slide, unique local type-7
kind-1 placeholder, parent/lock state, exact reference metadata, optional
type-2001 storage and type-2043 attachment closure, and all local style,
attribute-table, presentation, and node-reference dependencies. Strict raw
decoders and private Buffa views agree on node field 18, storage fields 1/9/10,
and attachment super fields 1/2; the repeated attachment table stays borrowed
raw bytes behind a bounded handwritten decoder. The generated five-file
closure is 112,101 bytes under 116 KiB with SHA-256
`eacce4103b5c9f9f32fd98639b81249ae1d15fcd63da6fe636569e0a2a324c30`,
and generated production code exposes neither repeated views nor encoding.

Changed commits rewrite one or two selected components, delete existing root
previews, and reopen the candidate. The Node verifier permits only canonical
field-18 insertion/replacement plus message length and preserves every
thumbnail/cache byte and metadata record; the Slide verifier permits only the
field-7/42 membership splice plus message length. No-op shares the source;
exact-source apply reopens only its retained target; `inverse` restores the
complete source artifact and previews. Resource reports for strict/Buffa
decoding, scalar editing, exact delta verification, closure scans, and the
common `6 * source.len() + output_len + 2 * parsed_fields` slide-router plan
merge into one finite transaction budget before allocation/publication.

Native and Rust gates establish the semantic contract recorded in ADR 0003:
the Rust candidate opens warning-free with Slide Number checked and canvas
number 1, preserves title/body/date, survives Save As and reopen, and inverses
exactly. Apple-only show/hide controls change node field 18 and the placement
of the exact field-20 payload in fields 7/42 while preserving node caches,
KeynoteShow field 6, and the complete slide-number closure.

The retired host surface is exactly
`KeynoteEditor::set_slide_number_visible`, the complete 172-line direct module,
two whole direct mutation tests plus four exclusive constants, and the 23-line
`set_keynote_slide_number.rs` example. The cut also removes their
`test_package_with_slide_number` fixture helper. The source-free creation
example now uses the focused handoff. Snapshot read
`KeynoteSlideInfo::is_slide_number_visible`, creation-time
`KeynoteDocumentBuilder::slide_number_visible`, private creation/materialization,
layout mutation, and the shared ownership helper remain and are not
compatibility aliases for the deleted setter. Patches remain process-local and
non-durable; changed nested legacy sources refuse under Preserve. The frozen
handoff passed 8/8 codec, 98/98 Keynote library, 7/7 focused visibility, 9/9
facade, 22/22 preview, and 7/7 doctest cases; ADR 0003 records the nonnumeric
strict, fuzz, patch, and native gates. The boundary unit suite passed 138/138;
live slide-number host, placeholder host, and focused audits were empty, and
the full checker retained only the unchanged 14 dependency-policy baselines.

## 2026-08-11 amendment: focused Keynote soundtrack-settings API

The canonical surface is the direct namespace
`litchi_keynote::soundtrack::{Mode, Settings, Edit, Patch, Commit, Diagnostics,
Error, LimitKind}` plus
`Package::{soundtrack_settings, edit_soundtrack_settings,
apply_soundtrack_settings}`. It supersedes this ADR's earlier assignment of
playback-settings graph lookup and mutation to `litchi-iwa`; media-resource
ownership and ordered item CRUD remain there. No nested `settings` module,
flat `SoundtrackSettings*` transaction aliases, root aliases, raw identifiers,
component paths, reference payloads, or package source bytes enter the public
surface.

`Package::soundtrack_settings` returns `Result<Option<Settings>, Error>`:
`None` means there is no rooted soundtrack, while `Some(Settings::default())`
means a soundtrack exists with both optional playback fields absent.
`Package::edit_soundtrack_settings` returns a fallible edit of that singleton
or `Error::SoundtrackNotFound`. `Edit::settings` observes the staged value;
consuming `Edit::set(self, Settings) -> Self` is infallible because `Settings`
is already validated, and consuming `commit` returns the immutable verified
package, exact patch, and diagnostics. Exact application is a method on the
immutable source `Package`, not a mutable editor facade.

`Settings` owns only independently optional volume and mode. `None` preserves
field absence rather than spelling zero or a default. Finite `0.0..=1.0`
volume and named known `Mode` variants are checked semantic values;
`Mode::Unknown(i32)` preserves genuinely future values but cannot disguise a
known discriminant. Media filenames, durations, bytes, ordered entries,
package identifiers, and data-reference metadata remain deliberately absent.
The transaction error and `LimitKind` expose only redacted categories and
bounded measurements.

The physical owner is the unique nonexternal Show field-17 reference reached
from rooted Document object 1 field 2; it resolves to one type-21 Soundtrack.
Only optional Soundtrack field 1 (volume) and field 2 (mode) form the semantic
write set. Field 3 and its exact raw reference payloads, order, multiplicity,
aggregate and any field-local metadata, package-metadata ownership counts,
data records, and package members form a retained validation closure, not a
second API feature.
The transaction cannot create a missing soundtrack or add, reorder, replace,
remove, or garbage-collect a soundtrack item.

A semantic no-op shares the exact source and skips changed-only guards and
reopen. A change rewrites only the selected Soundtrack component, preserves
all media and previews, and reopens the full candidate. Exact-source apply
rejects a different source or prior semantic value; inverse restores the
complete original package. `Patch` is a process-local two-artifact capability,
not a serialized operation log or durable save protocol. Changed physical
nested `Index.zip` sources return `UnsupportedSource` under Preserve, while
reads and no-ops retain their exact representation.

The private generated projection contains only scalar fields 1/2. The strict
codec validates them and streams field-3 identifiers under the same bounded
decode, then cross-checks the Buffa snapshot. The five-file closure is 27,753
bytes under 32 KiB and has aggregate SHA-256
`458206e0b57d8ec5ae4c3fc706bf793ccd385ab867b7e92ac30d66ab1858b4d3`;
production generated code contains no repeated view or encoder. Typed limits
and aggregate transaction budgeting are part of the API contract. The frozen
performance review found no P0/P1 issue. A test-only real-streaming-path gate
compares realistic 4,096- and 8,192-record metadata/media states: reference
count doubles exactly and fields, work, and references remain within a 2.3
ratio, without a wall-clock assertion or production-path change.

The populated Rust/native evidence in ADR 0003 proves exact inverse, single-
component locality, preview and media preservation, warning-free Keynote
open/playback, native Save As/reopen, and an exact post-native no-op at the
normalized volume. It does not establish a settings right to mutate field 3.

The completed retirement removes both direct `KeynoteEditor` settings methods,
the whole 68-line settings module, settings-only shared-wire code and module
declaration for an exact production diff of two insertions and 91 deletions,
157 lines of direct host settings tests, and the complete 29-line legacy
example. The structure inspector and README now use the focused `Package` API.
No host settings alias or shim remains.

The `KeynoteEditor` item CRUD, `KeynoteSoundtrackItemInfo`, media example,
creation, and shared soundtrack wire/media substrate remain intentionally. The
accepted frozen gates are 5/5 codec, 1/1 focused scaling unit, 4/4 focused
settings integration, 99/99 Keynote library, 10/10 facade, and 8/8 doctest
cases, plus formatting, strict library, all-target, example, live-host, and
diff checks. The boundary unit suite passed 152/152; live host and focused
audits were empty, and the full checker retained only the unchanged 14
baselines: six development-only annotations and eight edge classifications.

## 2026-08-11 amendment: focused Numbers sheet-order API

The canonical surface is
`litchi_numbers::sheet::order::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind}` with `Package::{edit_sheet_order, apply_sheet_order}`. Existing
semantic Document/Package sheet iteration is the read surface. No
`sheet_order()` reader, flat `SheetOrder*` alias, package-root transaction
export, raw object ID, component path, archive/protobuf value, or source byte
slice is public.

`edit_sheet_order(&self) -> Edit<'_>` is an infallible borrow. Consuming
`move_sheet(self, selector, destination) -> Result<Self, Error>` accepts one
exact-name or checked-index selector and one existing final zero-based
destination after removal. A second move and an empty commit are typed errors.
Selection is against the immutable base snapshot; the edit is not a mutable
cursor and cannot accumulate a batch. A same-position move is an exact shared-
artifact no-op with zero changed diagnostics, preview deletion, or reopen.

The semantic write set is the relative order of existing rooted sheets. The
physical write set is deliberately dual: raw `TN.DocumentArchive` field-1
reference records and raw type-205 sidebar-root `TSK.TreeNode` field-2 child
records, together with their selected subsequences in the two MessageInfo
aggregate object-reference lists. Required Document field 5 resolves the
sidebar root, and each child field 3 must associate positionally with the
matching sheet. Every existing reference record, identifier, child/descendant
node, sheet, table, and data subgraph is retained exactly. Selected order IDs
appearing in any FieldInfo are `UnsupportedSource`, not a license to rewrite
field attribution. Optional non-order field declarations are permitted only
when their path and reference are exact.

Changed support is intentionally narrower than semantic read support: rooted
owners, children, descendants, and ordinary `TN.Sheet` objects must share
`Index/Document.iwa`. A `FormBasedSheet`, split owner, external/zero/duplicate
or role-aliased reference, mismatched sidebar order, merge/diff state, or
noncanonical selected framing fails closed. `FormBasedSheet` reorder remains
native-unproven P2 debt, not an inferred compatibility claim.

A change is admitted only when the source has exactly one each of
`preview.jpg`, `preview-micro.jpg`, and `preview-web.jpg`; a missing or
non-unique canonical preview refuses the changed edit. It raw-splices the two
existing reference sequences, reorders only the matching aggregate
subsequences, rewrites one component, deletes all three previews, and reopens
the complete candidate. `Diagnostics` exposes changed, touched-component,
deleted-preview, and full-reparse facts. All unknown fields,
ViewState and CalculationEngine state, child/sheet/table content, unrelated
messages/objects/members, and ZIP metadata apart from derived sizes/checksums
and offsets remain exact. Forward apply requires the retained exact source and
prior moved-sheet position and validates the three-to-zero preview transition;
inverse swaps directional artifacts/proofs, validates zero-to-three, and
restores the complete original package including all three previews. Patches
are process-local two-artifact capabilities, not semantic logs, merge formats,
or durable save protocols; publication remains `Package::write_to`.

The sole generated schema is the frozen scalar-reference projection
`TNNumbersSheetReferenceArchive.proto`. Handwritten two-pass routers own
Document fields 1/5 and TreeNode fields 2/3, force each selected reference
through borrowed Buffa parity, and merge exact byte/field/work/reference
reports into one transaction budget. Generated production code has zero
repeated views and no encoder. Its deterministic five-file closure is 32,579
bytes under 33 KiB with SHA-256
`2a0850fd82cfbf337ed48e582d4a998bd27e5046eb63c61f6939fa5ff1a09854`.

Budgeting covers selected wire and metadata structure, indexed native lookup,
raw/aggregate reorder, component decode/encode, package reassembly, complete
candidate reopen, and locality verification. The real production path's
4,096-to-8,192 regression keeps work, references, and payload within 2.3 times
plus a fixed 32-unit allowance and rejects a max-minus-one reference budget in
preflight; it is deterministic resource evidence, not a wall-clock benchmark.
Performance review is P0/P1 clean. Bounded P2 costs are the roughly
four-per-sheet reference snapshots
retained by a changed patch under the 4,096-sheet cap, transient Vec-to-Arc
target duplication, and one possible full-byte authorization comparison for a
separately allocated but equal source. The identity path and patch inverse are
constant time.

ADR 0003 records the complete source/control/native/Rust/resave and preview
hashes. Together they prove warning-free Numbers 14.4 open, visible tab and
cell-marker correctness, Save As/close/exact-path reopen, exact Rust inverse,
exact post-native positional no-op, and preview regeneration equal to an Apple
reorder. Apple-only fresh tree/ViewState/revision identity, metadata/physical
ordering, and CalculationEngine cache culling are normalization evidence, not
an expansion of this API's minimal write set.

The accepted production gates include 7/7 codec, 132/132 proto, 109/109
library, 4/4 private, and 1/1 public focused cases, plus all-target,
all-feature, formatting, and diff checks. There is no sheet-order strict-Clippy
finding; unrelated legacy baseline diagnostics remain elsewhere. Package,
feature, dependency, declaration, and ordered-debt topology is unchanged at
64 packages, 235 internal declarations, 14 host declarations, and 14 debts.

The host-retirement boundary removes only direct
`NumbersEditor::move_sheet`, exclusive `selectors::sheet_index`, its legacy
example, move-specific direct/mixed tests, and README call. The exact cut is 58
production deletions, +2/-43 host-test lines, and the 23-line move example; the
retained remove-sheet example's semantic-selector migration is +2/-6. Sheet
reads, add/duplicate/remove CRUD, table moves, drawable ordering, creation, and
shared `update_numbers_document` writing remain outside this handoff. The
boundary suite passes 165/165; host and focused audits each report zero, and
the full checker reports only the unchanged 14 baselines: six development-only
annotations and eight edge classifications. The focused boundary inventory is
the private error/resolve/rewrite helper tuple and five source files.

## 2026-08-11 amendment: focused Numbers table-title API

The canonical public namespace is
`litchi_numbers::table::title::{Settings, Edit, Patch, Commit, Diagnostics,
Error, LimitKind, Path}`. It has no flat `TableTitle*` aliases or package-root
transaction re-exports. The immutable `Package` owns exactly these methods:

- `table_title_settings(sheet, table) -> Result<Settings, Error>`;
- `edit_table_title(sheet, table) -> Result<Edit<'_>, Error>`; and
- `apply_table_title(&Patch) -> Result<Commit, Error>`.

Both read and edit select an exact semantic sheet by name or checked position,
then a table by name or checked position within that sheet. No native object
identifier, archive/component locator, protobuf/generated value, source byte
slice, or mutation-capable package handle crosses the public boundary. Source
bytes stay crate-private and publication is through `Package::write_to`.

`Settings::new(visible: Option<bool>, outlined: Option<bool>)` is the complete
lossless semantic value. The two accessors return presence exactly;
`is_visible` and `is_outlined` are convenience effective-value checks that are
true only for `Some(true)`. `None` and `Some(false)` are never aliases. The API
supports all nine combinations as Rust transaction values, but native evidence
currently proves only visibility field 22 changing from `Some(true)` to
absent. Neither explicit false nor outline field 37 is claimed native-proven.

`Edit::path` reports only checked semantic positions, `Edit::settings` returns
the staged complete value, consuming `Edit::set(self, Settings) -> Self` is
infallible, and consuming `Edit::commit` performs source admission and returns
one verified immutable package, exact patch, and content-redacted diagnostics.
Selection and initial settings are resolved against the immutable base
snapshot. `Patch::{before, after, path, inverse}` exposes semantic and reversal
information without exposing retained artifacts; diagnostic fingerprints do
not authorize application.

The error vocabulary distinguishes missing sheet/table selectors, a selected
table lock, missing visible-title rendering dependencies, unsupported physical
provenance, invalid rooted source, finite resource exhaustion, allocation,
verification, and exact-source patch conflict. `LimitKind` covers package and
entry bytes, entries and aggregate bytes, payload bytes/objects/messages/items/
references, strict wire input/output/fields/nesting/work, and aggregate
transaction work. `Path` is content-free (`Package` or checked sheet/table
positions); errors and `Debug` output do not reveal sheet names, table names,
member paths, native IDs, or document content.

Changed publication is deliberately narrower than reading. It accepts the
shared, uniquely rooted Document -> Sheet or FormBasedSheet -> TableInfo ->
TableModel ownership profile only when selected reference metadata and message
framing are exact, the table is unlocked, and the physical source is exact.
Canonical and unambiguous legacy TableInfo/TableModel message variants remain
readable; changed nested/non-exact storage returns `UnsupportedSource` under
Preserve. An effectively visible requested title additionally requires the
existing finite nonnegative title height and distinct exact local paragraph-
and shape-style closure. The semantic API neither exposes those prerequisites
nor grants permission to repair them. Before any changed commit, the owner
scans all messages in `Index/ViewState.iwa` and returns `UnsupportedSource` if
any native type-6284 transient table-title selection message is present. This
is a conservative package-wide refusal, not an ownership inference. Reads and
exact no-ops skip the changed-only guard; an accepted changed source has no
type-6284 message, and every other ViewState byte remains exact.

An equal staged value is an exact shared-artifact no-op with zero component or
preview changes and no candidate reopen. A change preserves the independent
field-22/37 presence requested by `Settings`, touches one selected
CalculationEngine component, deletes every existing canonical root preview,
and completely reopens and verifies the candidate. Exact apply conflicts on a
different source or prior semantic state; inverse restores the full source
artifact and preview set. The patch is a process-local exact two-artifact
capability, not a durable semantic operation, serialized patch format, merge
contract, or save protocol.

The codec boundary is scalar-only: private
`TSTTableTitleSettingsArchive.proto` covers fields 22/33/37, while strict raw
routing forces fields 30/36 through the already private scalar Reference view.
Generated code has no encoder or repeated-field view. Its frozen five-file
closure is 32,332 bytes under 33 KiB with SHA-256
`56cfd70666ffa6079175bdab0a63a4ddd055099edf3c771ed3ad8b3051596ee1`.
This projection is validation evidence, never the preservation authority.

Final performance review reports no P0 or P1 issue. The real rooted `Package`
path's 4,096-to-8,192 structural gate measures fields 53,307 to 108,363
(2.0326 times), wire work 315,936 to 636,752 (2.0155 times), references 16,386
to 32,770 with the exact `2 + 4N` formula (1.9999 times), and transaction work
9,084,384 to 18,298,157 (2.0142 times). All remain within 2.3 times, and a
maximum-minus-one allowance rejects before output. The test has no wall-clock
threshold. Bounded linear selector temporary vectors and redundant changed-
path decodes are accepted P2 debt; the API admits neither quadratic work nor
unbounded retention.

The matched Numbers 14.4 control and hidden-title artifacts are respectively
136,204 bytes/SHA-256
`25c9fc858ca4fb4f1fedeafb944e96afb81af03a082a41be297ecf6f2542dbdb`
and 136,273 bytes/SHA-256
`ac8a7117ad6256b0da2e6d191b9e64f721b689d71696a89ac0f78bc6aa513a28`.
The matched oracle establishes field 22 true-to-absent only. The final Rust
gate starts from the 136,357-byte source SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`,
produces the 136,351-byte hidden artifact SHA-256
`4c7f6340b6f2675240577c5b59d5c154de24c8a7e763a31257c56a9899a8e40c`,
and inverses exactly to the source. Numbers 14.4 opened the Rust artifact
without warning, showed the title checkbox off while retaining the 22 by 7
table, `B2` marker, and `B3 = 42`, and preserved those semantics across native
resave and exact-path reopen. That 136,353-byte output has SHA-256
`5b162f8431f45333f0ae9a8654dfa724794f2ec2b391ea11f6a5eee7822cbb10`.
These oracles do not extend the semantic write set to explicit false, outline,
or ViewState.

The direct `NumbersEditor` read/set methods, their four whole direct tests, and
the legacy raw-ID example are retired. The exact cut is 32 production lines,
245 test lines, and the 39-line example. The private cross-format package/wire
helpers and common `Settings` remain for Pages and Keynote, whose table-title
APIs and CRUD are outside this handoff.

The final gates pass 9/9 codec, 141/141 protobuf, 111/111 Numbers library, 2/2
private transaction, and 5/5 public focused integration tests. Boundary
regressions pass 173/173; live host and focused audits are empty, and the full
checker retains only the unchanged 14 baselines: six development-only
annotations and eight edge classifications. Topology remains 64 packages/237
internal declarations/14 host dependency declarations/14 ordered debts, and
debt 015 remains.

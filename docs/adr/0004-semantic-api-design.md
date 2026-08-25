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
the public `litchi_pages::{Root, Body, Document}` values form the immutable,
bounded semantic snapshot. The focused Pages owner admits and projects native
sources before publishing that state; `litchi-iwa` does not decode, construct,
or re-export this reader model. The semantic model exposes borrowed section
views and exact-name or typed `litchi_core::Position` selection; duplicate
names are a typed ambiguity instead of an arbitrary first match. Every section
can produce an unambiguous snapshot-local position selector, while a
producer-visible name is preferred when one is actually present. The focused
adapter does not synthesize section names from headings when the current schema
projection has no name, and native object identifiers and protobuf messages do
not appear in ordinary signatures. Snapshot cloning shares the semantic
section allocation and never reparses or mutates the source document.

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

## 2026-08-11 amendment: aggregate Pages section-settings API

The canonical aggregate value is `litchi_pages::section::Settings`. It is not a
protobuf message, property bag, raw-ID map, or lossy collection of effective
values. It owns at most one exact-size section-name allocation and packs native
presence/value for four optional Booleans; `Start`, `PageNumbering`, and
`PageNumber` retain the three optional pagination fields. `Settings::new()` and
`default()` mean that all eight native fields are absent, not that Pages
defaults were serialized explicitly.

Callers read and mutate one existing section through:

```rust,ignore
Package::section_settings(SectionSelector) -> Result<Settings, Error>
Package::edit_section_settings(SectionSelector) -> Result<Edit<'_>, Error>
Package::apply_section_settings(&Patch) -> Result<Commit, Error>
```

The focused `section::settings` module owns `Edit`, `Patch`, `Commit`,
`Diagnostics`, `Error`, `LimitKind`, `Path`, and `DependencyKind`. `Path`
contains only `Package` or a checked semantic section position. `Edit::set` is
consuming and returns the edit after validating the complete value; staging a
failure leaves the source package unchanged. Public errors and `Debug` output
omit authored names and values, native identities, member names, raw bytes,
retained artifacts, and lower-layer diagnostic strings.

The semantic distinctions are lossless:

- every Boolean distinguishes absent, explicitly false, and explicitly true;
- the name distinguishes absent, present empty, and present nonempty, rejects
  NUL, and permits duplicate destination names;
- exact-name selection is case-sensitive and reports ambiguity for duplicate
  current names;
- section start and numbering preserve future native `u32` discriminants but
  reject an `Unknown` wrapper that aliases a known variant; and
- a starting page number is absent or a nonzero checked `PageNumber`.

`DependencyKind` reports only a missing otherwise-valid
`PreviousSectionTemplates`, `FirstTemplate`, `EvenTemplate`, or `OddTemplate`.
Malformed, contradictory, ambiguous, aliased, external, or wrongly typed
relationships are `InvalidSource`, not unsupported caller values. A missing
prerequisite is checked only when the target setting needs it. Exact no-ops
remain valid without forcing dormant dependencies.

This surface owns replacement of fields 17--22, 26, and 28 on one existing
section. It does not create, delete, or reorder sections; allocate identifiers;
edit body/header/footer text; create or mutate page templates; edit background
field 30; change guide or hyperlink state; normalize legacy packages; publish
files atomically; or expose a reference collection. Layout/cache state and root
previews are neither semantic `Settings` members nor mutation consequences;
matched native evidence requires their exact preservation.

`section_name` and `section_pagination` remain narrow ergonomic projections.
They select and expose only their established semantic subset, then delegate
publication to the aggregate owner. Their continued public presence therefore
does not authorize multiple physical writers. This amendment supersedes the
older statement that fields 20--22, names, and header/footer flags have
independent current mutation owners; it does not invalidate their prior
semantic or native evidence.

The implemented API is covered by 7/7 focused integration tests and four
private production/security gates for budget observation, object scaling,
alias-metadata refusal, and repeated-reference scaling/max-minus-one refusal.
Its real-package 4,096-to-8,192-object scaling keeps selected fields, strict
wire work, and references constant at 77, 564, and 4;
aggregate transaction work scales 292,154 to 587,222 (2.0100x), with one
output allocation and reopen at either size. A maximum-minus-one work budget
returns a typed error before either operation. These counters establish bounded
linear transaction behavior, not latency or RSS performance. The full Pages
library/integration gate passes 118/118, boundary regressions pass 181/181,
focused facade and host audits are empty, and the live checker retains only its
14 established dependency-policy baselines.

## 2026-08-11 amendment: Numbers table-cell read semantics

`litchi_numbers::table::cells` now defines the read-only semantic result:
`State` pairs a checked `CellPosition` with `Storage`, and `Storage` preserves
the difference between `Missing` and `Stored(Value)`, including
`Stored(Value::Empty)`. `Error`, `LimitKind`, and `Path` are typed and
content-free. `Package::table_cell` reads one checked coordinate;
`Package::table_cells` returns a fallibly allocated dense row-major `Vec<State>`
for a checked half-open `CellRange`. Exact name selectors reject duplicates,
index selectors are checked, and bounds, retained-element, and owned-text
limits reject before a partial result can escape.

The implementation projects from the package's already-eager semantic
`Table`, built by `litchi-numbers::package::extractor` through its existing
BNC/protobuf path. It does not call the newly committed strict table-cell
storage or dependency Buffa codecs. Those projections prepare a later
physical owner and are not a second public model, a current read dependency,
or an encoder. This amendment narrowly supersedes the earlier claim that the
monolith alone owns BNC-backed semantic decoding. It leaves `litchi-iwa` in
charge of existing cell mutation, formula compilation and AST wire handling,
calculation-engine mutation, and formula-cache changes. No edit, patch, cache,
preview, source-byte, or output API is implied by this read surface.

For range area `A`, materialized cells encountered across the selected row
span `K`, selected owned-string bytes `B`, selected owned strings `T`, and
table materialized-cell count `M`, the non-empty path has
`A + 2K + 2*O(log M)` size-sensitive work, one fallible `A`-state allocation,
and `T` fallible string allocations totaling `B` bytes. Empty ranges scan and
allocate nothing. Applying these formulas to paired 4,096/8,192 shapes keeps
size-sensitive terms at or below 2.0x and the result-allocation count at one;
element and text over-limit cases reject before that allocation. These
governed counters establish the shape of the algorithm, not elapsed-time or
resident-memory performance.

The native `basic.numbers` read oracle is 136,357 bytes with SHA-256
`f225d5b1cd59e9da454f91a96fe8f81154bc31037c10029230e75d49b45fb693`.
It establishes a 22x7 Sheet 1/Table 1, stored text at B2, stored number 42 at
B3, a missing A1, dense row-major A1:C3 behavior, and an out-of-bounds A23.
The 140,498-byte formula/rich-text oracle, SHA-256
`80deb7b87df27f58b26e6f247acee9d1fc6dcd3d268e85046c3efc16070b2edf`,
adds formula and rich-text-backed read coverage. Stored-empty presence is
synthetic-test evidence only. The gates pass 114/114 Numbers library, 4/4
public read integration, 13/13 strict codec, 163/163 full protobuf, and
187/187 boundary tests, plus strict library/test Clippy and warning-denied
rustdoc; the full checker has only the unchanged 14 baselines.

## 2026-08-12 amendment: Numbers table-cell mutation semantics

The read-only model above is extended, not replaced. The canonical namespace
now contains `Input`, `Change`, `Edit`, `Patch`, `Commit`, `Diagnostics`, and
`DependencyKind` beside `State`, `Storage`, `Error`, `LimitKind`, and `Path`.
`Input` admits only finite text, number, Boolean, Apple-epoch date, and duration
scalars. `Change` is one checked set or clear. `Package::edit_table_cells`
selects a sheet and table before staging; the consuming edit supports single
and iterator batches without exposing object identifiers, archive bytes,
protobuf/Buffa values, BNC records, component paths, or native references.
`Package::apply_table_cells` borrows the public exact patch while the private
retained `PackagePair` remains an implementation detail.

The semantic surface does not imply arbitrary cell editing. Accepted changed
sources cover scalar writes and direct/unsegmented string-list ownership with
exact refcounts, sparse missing-tile growth including a 513-row synthetic
boundary for finite non-text scalars, in-place authored-text replacement in
uniquely owned rich backing,
and strict downstream formula-cache refresh evaluated from the complete final
batch overlay. The rich path retains key/storage identity and releases exact
style references. It does not construct formulas, edit formula or
error cells, format cells, alter controls, change comments, or edit merge,
pivot, category, spill, hidden, or conditional-style state. Header cells with
rooted HeaderNameMgr ownership are rejected: CalculationEngine field 14 is
projected and root-validated, but the referenced manager payload/update
semantics are not. Sparse text-to-missing-tile changes refuse as
`SharedString`. Segmented string lists, aliased rich text, nonempty rich
FieldInfo reference transitions, noncanonical/ambiguous FieldInfo rich
ownership, existing formula or error cells, and modeled unsupported reachable
formula forms, cycles, ranges, deletion, and sparse/cache combinations fail
with `UnsupportedDependency`. Canonical payload field-1-to-storage and storage
field-2-to-style FieldInfo metadata may remain present and exact when no
field-specific transition is needed. Impacted active merge, pivot,
category, spill, hidden, or conditional-style states refuse by their matching
dependency kind; unrelated/inert state remains exact. A modeled missing
storage prerequisite fails as `UnsupportedDependency { CellStorage }`,
malformed storage fails as `InvalidSource`, and an unmodeled stored BNC
value/source kind fails as `UnsupportedSource`. Locked tables reject changed
edits.

The storage and dependency Buffa projections are strict borrowed validators,
not public semantic owners or encoders. Their generated ratchets are
respectively five files/465,932 bytes/SHA-256
`1a894fd5d22b004db664bc7c348d9591a4608ab9263a8122c726c8a1ecb0c3b3`
and five files/544,538 bytes/SHA-256
`2fba7c22aef58ed3cfe6eba1f77e5eaf79d2597dd79966e05d20e50c0e2b33b3`,
with zero repeated generated views in either projection. Caller-owned raw
records remain authoritative for exact preservation and splice. The strict
dependency-only formula projection remains five files/201,539 bytes with
SHA-256
`ccd972b3dcd76b6142342d36435f2f76a305c029265853ced04d64c1e2bf1752`;
its focused codec passes 7/7 and the full protobuf suite passes 178/178.
The PackageMetadata projection is five files/145,681 bytes with no repeated
generated view and SHA-256
`ee49927f75c6b632c83055f9b7e647920b389be41bec10e25871a6ef7b56ab31`;
its focused gate passes 7/7.

Changed publication verifies exact directional message and aggregate/FieldInfo
reference deltas and exact source-to-target preview membership. The private
source/target `PackagePair` lets exact apply and inverse reuse already reopened
snapshots under the same read profile; it does not make `Patch` durable. Rooted
4,096-to-8,192 transaction-work ratios are 1.1899x numeric, 1.2245x unique
text, 1.1396x same-tile batch, and 1.8021x formula closure, with no governed
subterm above 2.0x. Required-minus-one formula and sparse budgets refuse before
component, reassembly, output, reopen, or locality work. No wall-clock/RSS
claim follows.

Reads and exact no-ops retain broad compatibility. A changed package without
an exact physical `SourceCatalog`, including nested legacy layout, fails as
`UnsupportedSource` before mutation.

Native Numbers 14.4 accepts the exact numeric B3=43 scalar and unique-rich
authored-text candidates; the rich gate proves no-impact preservation of its
independent formula/cache,
not impacted formula refresh. The completed cut removes the three direct
NumbersEditor cell mutators, Numbers-only raw-ID writers and batch apply, 15
obsolete tests, and the legacy example. Shared attached-table APIs and
Pages/Keynote ownership remain. The new private rich-text wire implementation
introduces `litchi-numbers -> litchi-iwa-text-wire`, raising internal
declarations from 237 to 238 while leaving 64 packages, 14 host declarations,
and 14 ordered debts unchanged.

## 2026-08-12 amendment: Keynote existing-slide deletion semantics

The canonical API is
`litchi_keynote::slide::delete::{Edit, Patch, Commit, Diagnostics, Error,
LimitKind, Path}` with
`Package::{edit_slide_deletion, apply_slide_deletion}`. A caller stages
`Edit::remove_slide` with an exact `SlideSelector` name or a checked semantic
position (`Position`, including its established `usize` conversion). Native
SlideNode or Slide IDs, component and package-member names, UUIDs, media IDs,
protobuf records, and raw bytes never enter the public vocabulary.
`KeynoteEditor::remove_slide` is retired without a compatibility method or
public bridge alias.

Selection uses the immutable base snapshot. Names mean developer-facing
navigator names, not visible title text; duplicate exact names are ambiguous.
Deletion of the last slide is a distinct refusal. `Commit` returns the new
immutable package, a reversible exact-source patch, and content-free
diagnostics. Forward diagnostics report one slide removed and none restored;
inverse application reports none removed and one restored. Both report that
the operation changed, the number of distinct touched components, and whether
a full reparse occurred; they do not reveal selected content or physical
identity. Public patch `Debug` likewise omits names, members, fingerprints, and
bytes.

This surface authorizes removal of one existing slide only. It does not imply
slide creation, insertion, duplication, bulk deletion, ordering, layout
assignment, placeholder creation, drawable/media CRUD, or a public native
graph. It rejects unsupported hierarchy and uncertain incoming ownership
rather than asking a caller to select cascade, detach, or retarget behavior.
Slide ordering remains the distinct `slide_order` transaction family.

Nor does `remove_slide` mean media reclamation. The physical transaction
removes exact data-reference owner/count records for the deleted Node and
Slide. A component data-reference record remains with surviving owners or is
removed when none survive, while the global data-catalog record and every
media/data payload remain preserved. Shared or uncertain media can remain
after the semantic slide is gone. Reclaiming unreachable data is a separate
future `gc` API and proof.

The focused `remove_slide` example accepts `index:N` or `name:NAME`, requires a
distinct new output, optionally materializes the exact inverse at another
distinct path, and uses synchronized sibling-temporary no-clobber publication.
It demonstrates safe command-line publication without claiming that
`Package::write_to` itself flushes, syncs, renames, or atomically replaces a
filesystem destination.

## 2026-08-12 amendment: no formula-authoring semantic facade yet

The bounded internal Numbers cache planner preserves an unrelated cycle marker
byte-for-byte, refuses when an impacted marked formula survives the final
same-batch overlay, and succeeds when the overlay removes it. Graph work has
exact max-minus-one refusal coverage; scratch and allocation remain bounded.
This foundation exposes no supported semantic formula constructor, formula
edit input, or focused formula mutation method. The production host formula
API remains in place. No native application or formula-authoring performance
evidence is accepted by this amendment.

## 2026-08-13 amendment: Pages section-background semantics

`litchi_pages::section::Background` models the direct fill on one existing
section as `None`, `Solid(Rgba)`, or `Unsupported`. `Solid` accepts only finite,
in-range sRGB or Display-P3 components. `Unsupported` is an observable
preservation state for gradients, images, future fills, and unsupported color
models; it intentionally carries no raw payload and a changed edit from that
state is refused. Consequently callers cannot inject native protobuf bytes or
choose a destructive conversion policy.

`Package::section_background`, `edit_section_background`, and
`apply_section_background` use `SectionSelector`, resolve it immediately, and
retain only `Path::Section { position }`. `Edit::background`, `set_solid`,
`clear`, and `commit` form the complete mutation vocabulary. Typed,
content-free errors distinguish selector ambiguity/missing position, invalid
or unsupported source, bounded resource refusal, verification, allocation,
and patch conflict. `Debug` omits content, artifacts, names, locators, and
bytes.

This surface owns only TP.SectionArchive field 30. It neither creates or
deletes sections, changes fields 17--29 or 31, touches templates/text/cache or
preview state, nor implements gradient/image/media editing. Those capabilities
remain separate ownership and migration work.

Native acceptance covers only the two supported semantic states: Pages 14.4.1
opened replacement as a dark-red `Color Fill` and clear as `No Fill`, without a
repair or conversion prompt, and preserved each state through Save, close, and
exact-path reopen. It does not expand `Unsupported` into a mutable variant or
establish semantic ownership of any other fill family.

## 2026-08-13 amendment: canonical Keynote reader API

Callers read Keynote semantics through `litchi_keynote::Document::{open,
open_with_options}`. Both complete ZIP artifacts and app-authored package
directories are frozen through `PreparedSource`, fully decoded under checked
limits, and returned as archive-free snapshots with full `Show`, rooted text,
source-derived metadata, and source statistics. Metadata begins with semantic
Show fields and incorporates narrowly decoded canonical-properties scalar
diagnostics when that sidecar exists; `Some` does not establish its presence.
The cross-format `litchi::iwork::Document` coordinator can delegate to this focused
semantic owner. Exact complete regular-file artifacts remain on
`litchi_keynote::Package`, including byte ingress, semantic projection,
cheap shared `semantic_snapshot`, exact `write_to`, and edit provenance. A
package-derived semantic snapshot is intentionally diagnostic-free, with
`metadata()` and `stats()` both `None`. The retired
`litchi_iwa::keynote::KeynoteDocument` surface is not aliased or bridged.

Every supported legacy capability has a focused equivalent distributed by
provenance: `Document` owns semantic path opening, cheap snapshots, text,
slides, metadata, show, validation, and source stats; `Package` owns exact byte
opening and physical-artifact operations. The sole legacy-only
constructor name, `from_archive_bytes`, performed exactly the same work as
`from_bytes`, so `Package::from_bytes` is canonical. The retired stats type's
`application` member was always `Keynote`; format identity follows from the
concrete reader rather than a repeated runtime field.

The ingress split prevents a false provenance claim. A directory's
`Index.zip` or loose `Index/` can supply semantic components to `Document`, but
is not the complete regular-file artifact needed for `Package::write_to` or
exact-source edits. `Package::open` therefore refuses directory backing.
`Document` deliberately does not retain an archive and promises no preservation
of other sidecars, `Data/`, previews, or exact package bytes.

The semantic contract is focused rather than historically bug-compatible.
`text` is rooted and excludes unreachable or template-only storages. Slides
retain rich storage fragments instead of flattening body/date text into a
legacy vector. Focused metadata can expose plist revision and content-status
facts, and validation rejects malformed graph or resource state under retained
bounds. No raw `Bundle`, `ObjectIndex`, native identifier, or Prost message
enters the public API. The focused semantic traversal still performs bounded
generated Prost decoding after wire preflight; `Document` eagerly completes
that projection before publication, while `Package` performs it lazily.

Focused metadata accepts only the canonical logical
`Metadata/Properties.plist`. It ignores unrelated entries that merely share the
basename. Legacy nested-ZIP wrapper prefixes are normalized by the package
catalog; arbitrary flat wrapper prefixes receive no such authorization. The
canonical diagnostic has a centralized, non-configurable 64 KiB admission
ceiling independent of broader entry limits, and decoding is limited to the
scalar fields projected into public metadata.

This amendment retires only the read facade. `KeynoteEditor`,
`KeynoteDocumentBuilder`, creation, and mutation surfaces remain in
`litchi-iwa`.

## 2026-08-13 amendment: canonical Pages reader API

The supported Pages semantic read API is `litchi_pages::Document::{open,
open_with_options, from_bytes, from_bytes_with_options, from_shared_bytes,
from_shared_bytes_with_options}`. `DocumentReadOptions`
combines a checked `DocumentSourceLimits` profile with checked
`SemanticLimits`; the latter admits nonzero ceilings no greater than 4,096
sections and 64 MiB of retained section-name and text bytes. The resulting
`Document` exposes constant-time snapshots, semantic sections and selectors,
plain text, source metadata and statistics, and allocation-free semantic
validation. Wrong-family iWork input is the typed `ReadError::NotPages`.
`ReadError` exposes only closed I/O, wrong-family, limit, invalid-source,
invalid-format, and allocation categories. Semantic admission failures flatten
into those public categories, with resource failures reported through
Pages-owned `ReadLimitKind` values and numeric bounds. Paths, package member
names, native identifiers, native content, and lower-layer diagnostic strings
do not cross the archive-free boundary.

The Pages-owned `DocumentSourceLimits` profile names five source resources:
encoded input bytes, independently addressed source items, decoded bytes from
one source item, aggregate decoded source bytes, and bytes from one document
payload component. Their defaults are respectively 1 GiB, 100,000, 512 MiB,
2 GiB, and 512 MiB; callers may tighten but not relax those hard ceilings.
Rootless/body fallback retains at most 4,096 semantic storages.
The semantic text ceiling charges retained section names and storage text;
render-only section separators are not retained input bytes.

The API split is provenance, not two names for the same owner. `Package`
accepts exact regular-file and byte artifacts and owns `source_bytes`, package
metadata/statistics, physical validation, and edits. `Document` is
archive-free, accepts packaged Pages paths, checked directory sources, borrowed
package bytes, or shared package bytes, and owns no exact artifact.
`PagesDocumentStats::application` was always the constant Pages; the focused
statistics retain only object and native section counts.

Packaged and directory paths are available only where Pages can pin stable
source identity. Windows path ingress fails closed until it can do so safely,
while borrowed and shared byte ingress remains available. For packaged-path,
borrowed-byte, and shared-byte routes, the three exact canonical metadata
members are checked from physical package headers for supported compression and
the 64 KiB declared-size ceiling before package entries are materialized. This
selection compares exact raw logical-name bytes after stripping only the
selected legacy outer-package prefix; raw near-names do not match. The
selected-member preflight does not prevent unrelated supported members from
being expanded under the broader source limits before semantic projection
discards them.

Capability parity includes deliberate semantic corrections already accepted
for the focused Pages model. An empty native root has zero sections rather
than a synthetic unnamed body section. Rootless bodies use the bounded native
storage fallback: the same 14 message types trigger an object, all registry
storage messages in that object are fully raw-validated and Buffa-projected,
and all fragments are newline-joined in source order under the aggregate text
budget. Rooted bodies honor the native initial-section/table graph, UTF-16 boundaries,
section-break removal, exact name presence, and typed duplicate-name
ambiguity. The checked-in native fixture therefore has one Body section named
`Blank`, not the legacy unnamed synthetic section. These are format-owned
semantics, not object-for-object compatibility with the retired reader.

Source diagnostics bind only the three exact canonical authorities
`Metadata/Properties.plist`, `Metadata/BuildVersionHistory.plist`, and
`Metadata/DocumentIdentifier`; no more than 64 KiB from any one selected
authority is retained in the semantic handoff. Hostile basename matches and
unrelated sidecars are inert. Revision falls back to the latest admitted
build-history entry, and application defaults
to Pages when the properties file omits it. All present selected diagnostics
must decode before any source-backed `Document` is published.

The public API and behavior gates are frozen: 15/15 document-reader tests and
153/153 total focused Pages tests pass. Supporting archive and detector suites
pass 93/93 and 32/32, the host generated-roundtrip passes 1/1, and the boundary
suite passes 227/227 with zero live retirement or focused-public-API findings.

## 2026-08-13 amendment: canonical Numbers reader API

The supported archive-free Numbers read entry points are
`litchi_numbers::Document::{open, open_with_options, from_bytes,
from_bytes_with_options, from_shared_bytes,
from_shared_bytes_with_options}`. `DocumentReadOptions` combines a checked
`DocumentSourceLimits` profile with checked semantic `DocumentLimits`.
`DocumentLimits::new` returns `Result<DocumentLimits, DocumentLimitsError>`,
preserves zero as an exact tightening, and rejects values above the hard
semantic maxima. `DocumentLimitKind` identifies the rejected resource;
`DocumentReadOptions::new` remains infallible once both profiles exist.
Construction is eager and failure-atomic: a document is published only after
format ownership, graph structure, sheet/table/cell counts, semantic text,
source metadata, and source statistics have been accepted. Selected
projections use physical preflight where stated below; not every limit is
claimed to precede every decode or intermediate allocation. Focused `Document`
projection borrows packaged source bytes and releases them before publication;
selected components, sidecars, and semantic values are owned. Attacker-scaled
`Vec`, `String`, and collection growth is explicitly reserved fallibly where
the focused adapters own it. The standard library provides no fallible `Arc`
allocation, so bounded reference-count control blocks and final immutable
publication allocations retain an allocator-abort caveat. These guarantees
apply to focused `Document` constructors, not the preserve/edit `Package` path.

The returned API is selector-first and archive-free: `sheets`, `sheet`,
`sheet_count`, `snapshot`, `shared_sheets`, `plain_text`, `text_len`,
`metadata`, `stats`, and `validate`. None exposes a bundle, object index,
component name, member name, native identifier, protobuf/Buffa value, or exact
package bytes. Public source opens use a content-free `DocumentReadError`
taxonomy; its display, debug, and error-source chain retain only stable
categories and numeric bounds, never caller paths, authored content,
member/component names, native identifiers, or lower-layer diagnostic strings.

Rooted `plain_text` is a deliberate semantic correction, not compatibility
flattening. It visits sheets and their tables in source order and materialized
cells in row-major order. Each non-empty sheet name, table name, and non-empty
cell display value occupies one newline-separated line; missing cells, headers,
and empty rendered Text/Formula strings are not independently emitted.
Non-empty Text/Formula displays retain their strings, finite
number/date/duration displays use shortest round-trippable decimal text,
Booleans use `true`/`false`, and error displays use the `ERROR: ` prefix. This
rooted projection intentionally differs from storage-oriented text. The legacy
public `NumbersDocument::text` was unreachable on the native fixture because
legacy document construction failed first; the independently recovered private
storage projection matched `Package::text` on the basic and formula-rich
fixtures only. `Package::text` therefore remains the separate physical/storage
diagnostic, without an unqualified legacy-parity claim.

Source diagnostics are optional by construction. A source-backed document
retains content-free
`DocumentStats { source_record_count, sheet_count, table_count }`
and narrowly projected metadata from the three canonical Numbers authorities.
Semantic constructors and package-derived documents expose neither metadata
nor stats. This distinction prevents a package's semantic view from claiming
the acquisition provenance of a source-backed document. `application` is not
stored because the concrete type already proves Numbers format identity.
Each canonical metadata authority has a fixed 64 KiB physical ceiling before
ZIP payload materialization or directory-sidecar allocation. The private
`plist::stream` event projector additionally admits at most 1,024 events,
nesting depth 16, 128 build-history entries, 16 KiB per selected scalar, and
64 KiB of retained selected properties; it does not deserialize a general
scalar DTO or plist value tree. A build-history dictionary rejects duplicate
exact `Version` or `Build` keys; when distinct values are both present,
`Version` is authoritative. These diagnostic ceilings are fixed rather than
caller-extensible.

Exact complete artifacts remain on `litchi_numbers::Package`, including
regular-file and byte ingress, write/edit provenance, package text and object
diagnostics, and transactions. `Package::open` rejects directories. The
retired `from_archive_bytes` and limit variants were redundant byte aliases;
callers choose semantic `Document::from_bytes` or exact `Package::from_bytes`
instead of retaining a second name.

The focused reader is not Prost-free or wholly Buffa-lazy. Selective strict
raw/Buffa projections qualify root sheet order, sheet references and names,
TableInfo ownership, and table model/list/segment/tile/rich-text boundaries;
the focused `Document` projection skips comments. Generated Prost remains in
the wider admitted table and formula-owner graph. Deleting the second public
reader neither hardens every `Package` path nor authorizes a whole-graph codec
rewrite.

On Unix, file and directory path ingress uses pinned, no-follow capture. Other
non-Windows targets use version-checked path capture rather than descriptor
pinning. Windows path ingress is deliberately unavailable until stable,
reparse-safe handle traversal can freeze the source. Borrowed and shared byte
constructors are portable.

The public contract is frozen by 16/16 focused reader cases, with a seventeenth
Windows-configured case; 240 Numbers library cases pass and four are ignored;
compatibility and name gates pass 5/5 and 10/10. Archive coverage passes 127
cases (125 unit plus two integration), detector coverage passes 40/40, and the
host library passes 1,397/1,397, while generated-roundtrip and doctest gates
pass 1/1 and nine passed with three ignored. Host all-target check and no-run,
strict scoped host Clippy, focused
all-target Clippy, strict focused rustdoc, formatting, and diff checks pass. The
boundary units pass 237/237 and both live retirement/API audits report zero
findings. Broad host all-target Clippy remains blocked by unrelated existing
lints; the global boundary policy still reports 14 unrelated
`soapberry-zip`/`xml-minifier` debt findings.

## 2026-08-13 amendment: Numbers dimension-size API

The focused public read/edit surface is
`Package::{table_dimension_size, edit_table_dimension_size,
apply_table_dimension_size}`. Reads and edits select a sheet and table through
the existing selector contracts and then select one zero-based
`Dimension::Row` or `Dimension::Column`. The transaction module exports
`Edit`, `Patch`, `Commit`, `Diagnostics`, `Path`, `LimitKind`, and
`TransactionError`. `Path` reports semantic sheet/table/axis positions only;
native IDs and archive/component vocabulary stay private.

`Points::new` accepts only a strictly positive finite `f32`. `Size::Default`
is the semantic default-override state, representing native zero or physical
absence and resolving through the table model's row-height or column-width
default. `Size::Points(points)` is an explicit override. It must not collapse
to `Default` merely because `points.value()` equals the current model default,
and `Default` must not be presented as an explicit zero-point size.

The edit is deliberately singular: one selected dimension and one requested
`Size`. It validates the selector, axis index, unique table/header ownership,
stored header shape, strict/Buffa agreement, finite value, transaction limits,
wire rewrite, exact reopen, and locality before publication. The existing
strict/raw and private Buffa header-storage codec remains the read oracle; a
new generated-Prost-only compatibility decoder is not part of the public cut.

Despite the word “dimension,” this API controls display size only. It does not
change row count, column count, cell addressability, table area, or structure.
`resize_table`, axis insertion/deletion, automatic fit, equal distribution,
style-default authoring, and bulk multi-axis sizing are outside this cut.

The API is frozen by 13/13 focused dimension tests and 11/11 strict/Buffa codec
tests. The complete protos suite passes 194/194 plus doctests; Numbers passes
241 library tests with four ignored, 91 integration tests, and five doctests
with one ignored. Archive passes 130/130 plus doctests. Focused Numbers/archive
all-target check and strict all-target Clippy, formatting, and diff checks
pass. Boundary units pass 243/243 and both live retirement/API audits report
zero findings. Host all-target/all-feature check and no-run, retained-axis 2/2,
Pages-layout 1/1, generated-roundtrip 1/1, scoped boundary 6/6, and strict
library Clippy pass; broad host all-target Clippy retains nine unrelated lints.

## 2026-08-14 Selector-first formula-cell semantics

The public authoring surface is
`litchi_numbers::formula::{Expression, CachedValue, CellReference,
AxisReference, BinaryOperator, Table, Error, LimitKind}` together with
`table::cells::Input::{formula, formula_cached}`,
`Change::{set_formula, set_formula_a1, set_formula_cached,
set_formula_cached_a1}`, and the corresponding `Edit` methods. `Edit` also
provides `formula_table` for an opaque handle bound to the same immutable
source. A1 strings select destination cells only; formula bodies are semantic
values, never unparsed formula text.

Expressions support finite number, bounded text, and Boolean literals; mixed
absolute/relative cell endpoints; local and distinct-owner cell, rectangular
range, whole-row, and whole-column references; unary negate and percent; all
arithmetic, concatenation, and comparison `BinaryOperator` variants; and the
checked functions `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `AND`,
`OR`, `IF`, `IFERROR`, `NOT`, `ABS`, and `ROUND`. Function arity is checked.
One expression is limited to 1 MiB owned UTF-8, 65,536 nodes, depth 64, and 256
arguments per call. Non-finite numbers and reversed or out-of-bounds ranges are
rejected before package planning.

Typed caches support finite Number, bounded Text, Boolean, finite Apple-epoch
Date, and finite Duration. A supplied cache is only a claim: the independent
strict evaluator recompiles the authored AST, exact-compares dependency facts,
rejects cycles and poisoned direct scalar reads, applies the complete final
scalar overlay, and requires the typed result to agree. Aggregate reads ignore
Text and Boolean where Numbers semantics require it; direct Text remains a
typed formula error. Arrays, spills, deferred or volatile state, pivots, raw
formula identities, ambiguous owners, and unsupported native host shapes fail
closed rather than entering the public model.

## 2026-08-23 superseding amendment: Keynote soundtrack-reference order ownership

This amendment supersedes the earlier current-state wording that assigned
soundtrack item ordering to the migration host; those earlier dated passages
remain historical. `litchi-keynote::soundtrack::order` now owns the bounded
`Edit`/`Patch`/`Commit` transaction exposed through
`Package::{edit_soundtrack_order, apply_soundtrack_order}`.

The semantic operation is one checked move of an existing soundtrack media
reference. It validates the rooted Document field 2 -> Show field 17 -> type-21
Soundtrack chain, streams the bounded field-3 reference sequence, and preserves
the referenced assets, data-reference metadata, unknown fields, and all
unrelated package bytes. No media payload or asset is created, removed,
replaced, allocated, reclaimed, or otherwise edited by this seam. The existing
focused `soundtrack::{Mode, Settings}` API remains the separate owner for
playback mode and volume only.

The remaining soundtrack media/settings CRUD and legacy host compatibility stay
with `litchi-iwa`, including item/asset creation and add/insert/replace/remove
operations, lifecycle/reclamation, and host-facing readers/editors outside the
two focused package seams. The order transaction is therefore a bounded
semantic ownership slice, not a general soundtrack facade or compatibility
alias.

Its package-level bounded cases do not constitute a native Keynote open/save
gate. No `litchi-iwa` host edge or ordered debt is retired, and this amendment
makes no host-exit or monolith-exit claim.

## 2026-08-23 amendment: Keynote exact-file ingress and opaque validation authority

`litchi_keynote::Package::open` and its limit/options variants now admit only
a descriptor-pinned regular file. Unix opens the final component with
`O_NOFOLLOW`, `O_NONBLOCK`, and `O_CLOEXEC`; Windows rejects device namespaces
and non-disk handles, opens the final reparse point, rejects reparse
descriptors, permits concurrent readers while excluding
writers/renames/deletes, and retains descriptor identity. Other platforms fail
closed. The reader compares descriptor metadata before and after capture,
requires the observed byte count to equal the admitted length, retries
interrupted reads, begins with at most 64 KiB of source capacity, and grows
geometrically without exceeding the physical input ceiling. Byte constructors,
archive-free `Document` ingress, directory preparation, and output publication
are unchanged; this is not a claim of descriptor safety for every Keynote API.

Opaque transition color and path bytes remain public semantic values only as
bounded source-owned payloads. Their strict wire authority now resides in the
doc-hidden `litchi-iwa-protos::keynote_slide_transition_codec`, beside the
private Buffa transition projection. Known fields retain canonical framing,
wire-kind, singularity, required-field, UTF-8, Boolean, signed-int, and finite
float checks. Unknown fields remain byte-owned by the source and may retain
noncanonical framing, while groups remain forbidden. Aggregate selected-message
bytes, fields, and nesting preserve the predecessor `WireLimits` contract; the
codec's two scans are bounded by twice that unchanged aggregate byte ceiling
rather than introducing the rewrite-work limit as a new semantic restriction.

Graph selection, package mutation, raw-byte preservation, error mapping, and
patch/inverse proof remain private to `litchi-keynote`. The adapter-only public
`litchi_keynote::transition::validate_opaque_transition_settings` export is
removed intentionally from the unpublished `0.0.1` API; semantic callers use
`Settings`, checked payload setters, and package transactions rather than a
wire validator. Generated Buffa types and codec values do not cross the
focused public facade.

## 2026-08-24 amendment: focused Keynote slide-background ownership

Commit `0b12df5e1` adds the selector-first slide-background seam to the
focused Keynote package. `Package::{slide_background,
slide_background_override, edit_slide_background, apply_slide_background}`
select a slide by the existing semantic selector contract. The edit exposes
semantic `set`, `set_solid`, `set_gradient`, `set_opaque`, `clear`, and `reset`
operations, while `Background::{None, Solid, Gradient, Opaque}` is the public
value vocabulary. Effective inherited fill and direct override remain
distinct: resetting a variation restores its parent effective state, and a
shared collapsible variation redirects only the selected slide while retaining
the other owner's style.

The doc-hidden
`litchi_iwa_protos::keynote_slide_background_codec` owns the bounded wire seam.
It uses the private Buffa projection and borrowed snapshots for solid,
gradient, image, opaque, and explicit-none payloads. Selected known fields are
validated for canonical wire shape, required values, finite/ranged scalars,
and supported semantic models; unknown source bytes remain source-owned and
are preserved when the typed rewrite is safe. Deprecated protobuf groups are
rejected. Typed gradient rewrites fail closed when unknown fields occur inside
selected old stops, colors, or angle structures rather than silently dropping
data.

`litchi-keynote` retains graph selection, source/target style identity,
copy-on-write style creation, stylesheet and package-metadata UUID/external
reference maintenance, exact raw-byte publication, and reversible patch
application. The package invalidates affected previews and publishes edits
atomically; native identifiers and protobuf/wire types remain private. The
`litchi-iwa` slide-background adapter remains a compatibility consumer of this
semantic seam. This is a focused ownership slice and does not claim removal of
the migration-host facade or compatibility behavior outside this operation.

## 2026-08-24 amendment: Numbers name-publication save-token semantics

Commit `35c2ae281d40bc16c659c30dd3690a702a750ea0` gives the semantic Numbers
name-edit transaction an explicit PackageMetadata publication boundary. The
follow-up accounting hardening is commit
`ba57356166e82ff37ee1cd5223b68acd484aac42`; it removes detached, unmetered
candidate-size `Budget` work and makes candidate sizing and selector
comparisons share the reported budget.

A changed name commit requires exactly one `Index/Metadata.iwa` entry containing
the type-11006 PackageMetadata archive. It advances the root
`PackageMetadata.save_token` once, then writes that new root value only to the
selected current `ComponentInfo` records. Selection is by component identifier
and effective locator (explicit locator when present, otherwise preferred
locator); versioned records and unselected current components remain raw and
unchanged. `last_object_identifier` is left untouched. Missing, duplicate, or
ambiguous metadata ownership fails closed before publication.

The no-op path is byte-exact and bypasses metadata rewriting. A successful
change stores exact source and target package artifacts, so applying the
inverse restores the original bytes rather than decrementing a token in place.
Diagnostics continue to count semantic native components in
`touched_components`; the Metadata sidecar is not counted as an additional
native component. No raw object-ID API was added to the semantic facade.

The doc-hidden
`litchi_iwa_protos::package_metadata_codec` owns the borrowed save-token wire
seam, with private Buffa projection parity. Known root field 8 and component
field 12 require canonical singular keys, canonical values, and the expected
wire kind; malformed, duplicate, overflow, or selector-mismatch inputs are
rejected. Unknown source bytes, including accepted unknown groups and
noncanonical unknown scalar framing, remain source-owned and are retained.
The codec performs bounded sizing/preflight and candidate validation before
publication. A test-only aggregate proves every successful `Budget` charge
equals `RewriteReport.work_bytes`, and a max-minus-one limit fails before
candidate allocation. The package caller precharges decompressed type-11006
payload bytes multiplied by the changed semantic-operation upper bound for
visitor locator matching; compressed Snappy bytes are charged separately for
publication. The package still owns semantic selection, native name edits,
metadata locality, exact artifact capture, and patch/inverse behavior; wire
and generated types do not cross the public semantic surface.

## 2026-08-24 amendment: Numbers root cell-comment clear ownership

The Wave60 implementation series, culminating in commit
`895ef17848516cf201e717239c44c119eae5da87`, extends the selector-first
Numbers comment transaction with a deliberately narrow changed-clear seam.
`Package::{clear_table_cell_comment, clear_table_cell_comment_a1}` and
`Edit::clear` use the existing `SheetSelector`, `TableSelector`, and
`CellPosition` vocabulary. `TableCellCommentPatch` retains exact source and
target artifacts, and `Package::apply_table_cell_comment` applies the same
semantic patch/inverse contract as replacement. Clearing an already-empty
cell is an exact no-op.

A changed clear is admitted only for one existing root-list comment entry with
`refcount == 1`, one global cell occurrence, one list occurrence, one storage
occurrence, no segment ownership, and no replies. The storage UUID and all
archive aggregate/field reference metadata must be exact and unambiguous. The
selected BNC cell reference and root list entry are removed, the unshared
storage object is deleted, affected message metadata is pruned, and the three
root previews are invalidated. Shared, segmented, reply-bearing, aliased,
metadata-owned, unknown-owner, or otherwise ambiguous graphs fail closed.

Changed publication also requires the exact type-11006
`Index/Metadata.iwa` sidecar. The focused owner proves that the deleted storage
is absent from UUID, external-reference, data-owner, ambiguous-owner, and root
data-metadata-map ownership before deleting it. It advances the root save
token once and writes that value only to the current components changed by the
native mutation; `last_object_identifier`, unselected/versioned components,
and unknown source fields remain unchanged. The package reopens and verifies
the candidate before returning a commit, while inverse application restores
the original package bytes.

The deprecated raw-ID `litchi-iwa` clear remains a migration-host adapter. It
maps the native table owner to scoped sheet/table selectors and delegates
metadata-bearing sources to the focused owner, then reopens and checks legacy
readback before assignment. A package with no Metadata sidecar retains the
pre-existing compatibility implementation. Once Metadata is present, focused
parse, selector, ownership, limit, and verification errors are hard failures;
there is no cell-only fallback that could bypass the sidecar contract. Native
IDs, archive routes, protobuf messages, and wire types remain outside the
focused public API.

## 2026-08-24 amendment: Pages body-table lock canonical semantic vocabulary

The Wave61 implementation series is commits
`a27a22c9cf95d548da7634167d6bac7190b8321a`,
`887c4a9c87a21012b1bebfc3f7e93a1577522e21`,
`11c9550f9003da5e9c19f7976dbca132d3e21272`, and
`d6988d5ae279bb10a7d43e47746bf1a91bf5baf5`. The selector-first Pages owner
uses `Package::{body_table_lock, edit_body_table_lock,
apply_body_table_lock}`, `BodyTableSelector`, and the transaction types
`BodyTableLock{Edit,Patch,Commit,Diagnostics,Error,LimitKind}`. Semantic state
lives under `table::lock` as `BodyTableLockState`; no native table or object
identifier crosses this public surface.

The former flat `TableLock{Edit,Patch,Commit,Diagnostics,Error,LimitKind}`,
`TableLockState`, and `TableSelector` aliases were removed intentionally from
the unpublished `0.0.1` facade. They did not identify a distinct semantic
operation and made the application-owned body-table scope ambiguous. Callers
now use only the canonical body-table vocabulary above.

The package owner resolves the selected body attachment and table model,
cross-checks the rooted ownership graph, rewrites the selected lock field,
reopens the candidate, and returns exact source/target patch artifacts. Native
Pages packages commonly describe these known edges only in aggregate
`MessageInfo.object_references`; field-local reference metadata is optional.
When field-local declarations are present, they must be exact, unique, and
consistent with the payload and aggregate owner. Contradictory, duplicate,
wrong-path, or unknown ownership continues to fail closed. Patch inverse and
no-op application retain the exact-byte transaction contract.

Commit `4742e20107f29a1990d6a1886d8046a9333133b5` ratchets that vocabulary over
the complete `litchi-pages` source tree and the retained Pages migration host,
including alternate aliases, wildcard exports, relocated modules, renamed
host methods, and every host example. Individual `cfg(test)` items are masked
without hiding later production declarations. The `litchi-iwa` host remains a
compatibility concern; this semantic amendment does not claim its deletion.

## 2026-08-24 amendment: Pages body-table title semantic ownership

Commits `48f203aae56e43133fd931accfa6558661594ba0`,
`a92f8f11a50c709877b8d1f0a158da72114dc4ab`,
`a7088be4dd9fde9b2e843473b093839e7b7629b3`, and
`a1c1e83a3edad808bacefa648fe0c1bdd53308f5` establish the selector-first
Pages body-table title owner. The public surface is
`Package::{body_table_title_settings, edit_body_table_title,
apply_body_table_title}`, `BodyTableSelector`, the presence-preserving
`table::title::Settings`, and the
`BodyTableTitle{Edit,Patch,Commit,Diagnostics,Error,LimitKind}` transaction
types. `Settings` distinguishes absent optional Booleans from explicitly
stored `false`; no native object identifier, archive route, wire field, or
generated message crosses this semantic API.

The package resolves the selected rooted body table, cross-checks its native
model, and delegates strict selected-field projection and rewriting to the
doc-hidden `litchi_iwa_protos::numbers_table_title_codec`. Enabling a visible
title additionally requires distinct, non-external title paragraph and shape
references, exact aggregate ownership, exact field-local ownership whenever
that optional metadata is present, and unique one-message style objects of
the expected native types. Unknown or contradictory ownership, resource
references, duplicate objects, hostile message metadata, malformed selected
fields, and non-finite or negative title height fail closed.

The transaction preserves untouched source bytes, deletes the three root
previews only for a changed publication, reopens the complete candidate, and
returns exact source/target artifacts whose inverse restores the original
package. The tracked raw-`u64` Pages host title methods were removed
intentionally from the unpublished `0.0.1` compatibility surface, and the
retained host examples now use the semantic package owner. The Keynote title
host and other `litchi-iwa` operations are unrelated and remain.

## 2026-08-24 amendment: Pages body-table header semantic ownership

The Wave63 body-table header owner is finalized by commit
`7bc68903ef8f107b4473843945c908856988215f` after the shared semantic,
codec, package, host, boundary, and dependency-neutral hardening commits
`6f35a3d88`, `dd9e7f25c`, `58f04cd90`, `2bf7598d7`, `c93598d22`,
`7261dd1c4`, and `30207fa47`.

The selector-first surface is
`Package::{body_table_header_settings, edit_body_table_header_settings,
apply_body_table_header_settings}`, `BodyTableSelector`, the
presence-preserving `table::headers::{Count, Settings}` values, and the
`BodyTableHeaderSettings*` transaction types. Header/footer counts, freeze,
and repeating-header settings remain semantic values; native identifiers,
archive routes, wire fields, and generated messages remain private.

The package proves the rooted body-table/model graph and accepts aggregate
reference metadata emitted by Pages only when optional field-local
declarations are exact and consistent. Canonical nonzero local references,
external/data-reference rejection, duplicate/contradictory ownership checks,
malformed dependency scalars, and deprecated protobuf groups at package
ingress fail closed. The doc-hidden
`litchi_iwa_protos::table_header_settings_codec` owns borrowed
strict snapshots and raw-preserving rewrites behind private Buffa types.
Changed transactions preserve unknown source bytes, invalidate only changed
root previews, reopen the candidate, and retain exact source/target inverse
artifacts; no-ops remain exact.

The tracked raw-identifier Pages header methods were retired in favor of this
owner. Remaining migration-host compatibility operations are not removed.

## 2026-08-24 amendment: focused Keynote chart-caption semantic ownership

Commit `514b82bdf658d78ea0154f4fc075b1b85488f31c` establishes the
selector-first Keynote chart-caption text owner. The public package surface is
`Package::{slide_chart_caption, edit_slide_chart_caption,
apply_slide_chart_caption}`, `SlideSelector`, `ChartSelector`, and the
`ChartCaption{Edit,Patch,Commit,Diagnostics,Error,LimitKind}` transaction
types. `None` denotes an absent caption graph, while `Some("")` denotes an
existing native caption storage containing empty text. No native object ID,
component route, archive object, protobuf message, or wire payload crosses
this public surface.

The focused transaction intentionally owns replacement only. It resolves the
selected chart, validates the canonical chart-to-caption edge, checks the
caption-info parent, storage, placement, and style graph, proves that the text
storage is exclusively owned, and delegates the selected storage rewrite to
the strict internal Keynote text boundary. Creation of an absent graph and
changed removal both remain `UnsupportedDependency` because they require
object allocation, UUID/Metadata registration, or native stand-in policy.
Clearing an already absent caption and replacing text with the same value are
exact no-ops.

Changed publication preserves unknown selected-storage fields, invalidates
only the root previews, reopens the complete candidate, and returns exact
source/target artifacts whose inverse restores the original package bytes.
Shared storage, contradictory parentage, malformed selected wire fields,
dependent text markers, hostile message-diff metadata, and unknown ownership
fail closed. The `litchi-iwa` migration host now exposes selector methods and
delegates existing-caption reads and replacements to this owner; native graph
creation and stand-in removal remain compatibility-host operations. The old
raw-ID methods remain deprecated rather than deleted in this slice.

## 2026-08-24 amendment: full Keynote chart-caption graph semantic ownership

Commit `f0bbe079b094b3652751f6ab7c89b2dc64fac6a9` completes the admitted
Keynote chart-caption graph transition in the selector-first package owner.
The public surface remains
`Package::{slide_chart_caption, edit_slide_chart_caption,
apply_slide_chart_caption}`, `SlideSelector`, `ChartSelector`, and the
`ChartCaption{Edit,Patch,Commit,Diagnostics,Error,LimitKind}` transaction
types. `None` denotes the canonical empty stand-in, while `Some(text)`
denotes the inline CaptionInfo graph and its text storage; no native object
identifier, component route, archive object, protobuf message, or wire
payload crosses this semantic facade.

The owner now admits the complete canonical lifecycle: an exclusive canonical
stand-in can become an inline caption graph, existing text can be replaced,
and an active graph can be replaced by a fresh canonical stand-in. Creation
allocates the CaptionInfo, text storage, placement, and style graph and
registers every new object UUID in the exact Metadata sidecar. Removal keeps
the old graph for exact history/inverse semantics while retargeting the chart
to a fresh stand-in and registering that stand-in. Existing storage rewrites
retain the strict text-owner guarantees. Empty-to-empty edits remain exact
no-ops; changed transitions invalidate the root previews, reopen the complete
candidate, and return exact source/target artifacts whose inverse restores
the source bytes.

Chart-edge, CaptionInfo parentage, placement/style, storage ownership,
co-location, source-preserved wire fields, UUID allocation, Metadata
component selection, save-token advancement, last-object-identifier, and
candidate graph counts are checked together. Cross-component, shared,
ambiguous, malformed, hostile-diff, noncanonical, or otherwise unproven
graphs fail closed; the supported lifecycle is not a claim to normalize
arbitrary future caption graphs.

The former raw-ID `litchi-iwa` chart-caption methods were removed from the
production migration host and its example now uses selector-first methods.
This is an intentional unpublished `0.0.1` breaking removal of an
adapter-only surface; the host's remaining Keynote compatibility operations
are unrelated and remain. The focused facade and boundary ratchets keep
archive, metadata, wire, generated, and physical types private.

## 2026-08-24 amendment: Keynote chart-caption source-authority hardening

Commit `cd76c394e2cd6e360cd01e8bd61cb7594d837701` hardens the admitted
chart-caption lifecycle without changing its selector-first public surface.
`Package::{slide_chart_caption, edit_slide_chart_caption,
apply_slide_chart_caption}`, `SlideSelector`, `ChartSelector`, and the
semantic chart-caption transaction types remain the only public seam; no
archive route, component or object identifier, Metadata record, protobuf
message, or wire value is exposed.

Existing-caption text replacement now requires one exact current Metadata
component selector and advances the root and selected current component save
tokens in the same private publication. Missing or ambiguous Metadata,
selected-token overflow, malformed selected token fields, and unrelated
Metadata changes fail closed. Root last-object-identifier and unselected or
versioned component records remain source-preserved for this no-allocation
graph path.

Graph creation and stand-in removal now admit proven cross-component
stylesheet and paragraph-style dependencies only when the selected current
slide has the exact strong PackageMetadata external-reference tuples to the
physical owner components and objects. Missing, duplicate, weak, versioned,
or contradictory dependencies reject. The chart edge is rewritten with a
source-authoritative MessageInfo reference transition: aggregate and
field-local reference lists, retained raw ArchiveInfo header bytes, selected
payload, and candidate readback are authorized together. Known chart,
reference, theme, preset, geometry, and width fields require canonical
singular wire forms, and every parsed IWA object requires canonical outer
framing. Unknown retained header bytes remain exact rather than being
normalized.

The owner also rejects unrelated aggregate or FieldInfo owners of a selected
caption graph and aliased caption stand-ins. These checks expand the
fail-closed acceptance boundary; they do not claim support for arbitrary
future Keynote caption graphs or normalize producer extensions.

## 2026-08-24 amendment: Keynote chart-title raw-ID host retirement

Commit `e62b6fdb1` retires the remaining raw-identifier Keynote chart-title
host operation. The selector-first semantic owner remains
`Package::{slide_chart_catalog, slide_chart_title, edit_slide_chart_title,
apply_slide_chart_title}`, with `SlideSelector`, `ChartSelector`, and the
`ChartTitle{Edit,Patch,Commit,Diagnostics,Error,LimitKind}` transaction types.
The package owner and its focused behavior are unchanged; chart-title reads,
edits, no-ops, exact patches, inverse application, and candidate verification
continue to use semantic selectors and do not expose native identifiers,
archive objects, component routes, protobuf messages, or wire payloads.

Wave68 removes the deprecated `litchi-iwa::KeynoteEditor` raw-ID chart-title
method and setter/remover wrappers, together with their private
identifier-to-position bridge and fallback path. Nineteen examples and the
host chart-title tests were migrated to selector-first methods using a
positional slide and `ChartSelector`. This is an intentional unpublished
`0.0.1` breaking removal of an adapter-only surface; no compatibility shim or
public raw-ID replacement is added. Other Keynote migration-host operations,
graph ownership, and the remaining `litchi-iwa` compatibility edge remain in
scope for later slices.

This retires one host operation rather than declaring the host, dependency
edge, generated schema owners, or monolith complete. The semantic package
facade remains the authority for chart-title behavior, while any retained
host adapter is selector-based and cannot widen the public API with native
IDs.

## 2026-08-24 amendment: shared strict chart-caption edge

Commit `cd382f95dad52fecdfdb17a7fbb1565c489e6de1` moves the shared
`TSCH.ChartDrawableArchive.super` to `TSD.DrawableArchive.caption` to
`TSP.Reference.identifier` edge behind the format-neutral hidden
`litchi_iwa_protos::chart_caption_codec` seam. The existing
`keynote_chart_caption_codec` spelling remains as a compatibility name for
the focused Keynote package, but generated Buffa/protobuf types and wire
records do not cross the shared helper or any public facade.

Pages `PagesEditor::{body_chart_caption, set_body_chart_caption,
remove_body_chart_caption}` and Numbers
`NumbersEditor::{sheet_chart_caption, set_sheet_chart_caption,
remove_sheet_chart_caption}` retain their existing semantic behavior and
signatures. Their caption-local read and retarget paths now use one bounded
strict helper instead of decoding `IWorkChartArchive` locally or performing a
lossy nested-field patch. Broader chart graph and absent-caption creation
logic still owns its theme, storage, placement, UUID, Metadata, and component
dependencies in the compatibility host.

The selected reference requires canonical singular known fields, including
the retained deprecated type and external flags. Unrelated source fields,
unknown groups, and unknown scalar framing are preserved, while malformed,
duplicate, wrong-wire, noncanonical, nonfinite-limit, or failed candidate
readback states reject. This is a private shared codec boundary, not a new
public raw-ID API or a focused Pages/Numbers package owner.

## 2026-08-24 amendment: Pages drawable-order strict internal ownership

Commit `9b44e20ac0d3696aeb28f94d923612157446a83c` moves the complete
`TP.DrawablesZOrderArchive` repeated `TSP.Reference` projection behind the
hidden `litchi_iwa_protos::pages_drawable_order_codec` seam. The retained
Pages compatibility host at `litchi-iwa/src/pages/editor/drawable_order.rs`
no longer decodes the generated `DrawablesZOrderArchive` or imports Prost for
this operation. No `litchi-pages` focused semantic owner was introduced: a
document-wide drawable order has no meaningful public selector in the current
facade, and no raw-ID facade was added or widened by this slice.

The codec owns the complete repeated-reference path while preserving the
source-authoritative records. Root unknown fields may be interleaved with the
selected repeated field; complete nested references, unknown balanced groups,
and overlong unknown scalar framing remain exact. Known keys, length prefixes,
and identifier values are canonical and singular, identifiers are required,
nonzero, and unique, and requested edits must be an exact permutation. The
host retains semantic validation and transactional candidate/readback/reopen
ownership; native identifiers remain inside that compatibility boundary.

This is an internal strict-codec ownership step, not a public Pages API,
facade migration, or claim that the remaining Pages drawable graph has moved
out of `litchi-iwa`.

## 2026-08-25 amendment: Wave71 Pages exact output and footnote API

Commit `580a5343a2c75a1c1b185a5cc8aff4a87e2a5c11` makes the Pages package's
exact-output boundary explicit. `Package::from_bytes` and
`Package::from_bytes_with_limits` are the supported byte-ingress names;
`Package::from_archive_bytes` has been removed. `Package::source_bytes` is
crate-private. Public callers publish the immutable exact artifact through
`Package::write_to<W: Write + ?Sized>`, with `WriteError` reexported by the
focused crate root.

`write_to` emits the retained ZIP bytes exactly, including unsupported members
and unmodeled protobuf fields. It handles partial writes and
`Interrupted`, detects zero-length and sink over-reporting, and exposes only
typed, redacted failure information together with `bytes_written`; package
bytes and sink error text do not enter `Display` or `Debug`. It does not flush,
sync, rename, or atomically or durably publish a filesystem destination.

The same commit removes the raw-ID host setter
`PagesEditor::set_body_footnote_text`. Existing-root text replacement is now
owned by the selector-first `Package::edit_body_footnote_text` transaction.
The host still owns footnote reads and insert/remove graph lifecycle paths.
This does not add a raw-ID replacement API, broaden accepted graph shapes, or
move native identifiers, archive objects, generated messages, or wire values
across the focused facade.

## 2026-08-25 amendment: Wave72 shared chart-caption source hardening

Commit `997e0bf55e8a1381296a4d446f83dd5217a15385` hardens the hidden shared
chart-caption codec without changing the public semantic surfaces. Every
internal projection, sizing, rewrite, and candidate-verification traversal is
metered. Unknown overlong scalar values and complete source groups remain
source-retained, while unknown keys and length framing, selected known fields,
and exact group-depth accounting remain strict. Typed output and allocation
failures stay on the existing bounded error seam.

Pages and Numbers caption retargeting now preserves raw `ArchiveInfo` headers
and the exact aggregate and `FieldInfo` reference sets. Shared or
misattributed `CaptionInfo`/storage ownership rejects atomically rather than
publishing a partial graph. The Keynote owner folds the residual codec report
and candidate-reopen precharge into its existing private `CaptionBudget`.

This is an internal source-authority and accounting refinement. It does not
create a focused Pages or Numbers chart-caption package owner: graph/theme
creation and the remaining chart lifecycle stay in the compatibility host.
No raw-ID, archive, generated, Metadata, or wire type crosses a focused public
facade, and the P2 follow-ups for duplicate known `MessageInfo` scalar
canonicality and arbitrary unmodeled payload/reference parity remain open.

## 2026-08-25 amendment: focused Keynote movie-caption ownership

Commit `6ed7ee4e277c78a99f1b779a2b5530d7267beddf` adds a selector-first
Keynote package owner for captions on existing file-backed slide movies.
`MovieSelector` resolves checked movie source order within a semantic
`SlideSelector`; no native identifier crosses the public facade. The focused
surface is `Package::{slide_movie_caption, edit_slide_movie_caption,
apply_slide_movie_caption}` with exact `SlideMovieCaptionEdit`, patch, commit,
diagnostics, error, and limit types.

The hidden `litchi_iwa_protos::keynote_movie_caption_codec` owns the strict
`MovieArchive.super -> DrawableArchive.{title,caption} -> Reference.identifier`
projection. Selected known fields are singular, canonical, and source checked;
unrelated raw fields, balanced unknown groups, and overlong unknown scalar
values remain source authoritative. The package owner additionally proves a
file movie, exact slide parent and one global slide owner, distinct title and
caption edges, co-located and exclusively owned `CaptionInfo`/storage/style/
placement objects, exact Metadata current-component selectors, every touched
component save token, preview invalidation, candidate reopen, and exact patch
inverse/locality.

This phase admits reads, exact no-ops, and existing-caption `Some -> Some`
replacement only. Caption graph creation/removal and movie-title mutation
remain explicit focused `UnsupportedDependency` operations. The compatibility
host routes active-caption reads/replacements through the focused package but
retains title CRUD and legacy caption create/remove graph lifecycle. This is a
bounded semantic owner, not a raw-ID replacement API or full movie lifecycle
retirement.

## 2026-08-25 amendment: full Keynote movie-caption lifecycle ownership

Implementation commit `40c3d0b217e4b17304efca24c6a81b9d1240246b`
completes the canonical focused movie-caption lifecycle. The selector-first
surface remains
`Package::{slide_movie_caption, edit_slide_movie_caption,
apply_slide_movie_caption}` with `SlideSelector` and `MovieSelector`;
`SlideMovieCaptionEdit::{set, clear}` now admits exact no-ops, active-caption
replacement, canonical stand-in-to-caption creation, and active-caption-to-new-
stand-in removal. `SlideMovieCaptionPatch` retains exact source and target
artifacts, supports exact inverse application, and publishes no native object
identifier.

Creation is deliberately narrow: a canonical, co-located, exclusively owned
stand-in is replaced by a four-object caption graph (style, info, storage, and
placement). Removal redirects the selected movie to a fresh stand-in while
retaining the old caption graph and its UUID registrations. Both routes update
the exact current Metadata component selector, UUID registry, root object-ID
watermark, root save token, and selected component save token in the same
publication transaction. Malformed, shared, ambiguously owned, resource-owned,
cross-component, or otherwise unproven graphs fail closed.

The compatibility host now exposes only selector-typed movie-caption methods
and delegates reads, replacement, creation, and removal to this package owner.
Its raw movie-caption ID methods and legacy caption graph fallback are retired.
Movie-title mutation remains a separate compatibility-host responsibility and
retains its legacy raw-ID methods; this amendment does not claim a full movie
or Keynote-host retirement. The selected MovieArchive caption edge continues
through the hidden strict handwritten codec; generated schemas, Buffa
projections, and Prost types do not cross the public Keynote facade.

## 2026-08-25 amendment: full Keynote movie-title lifecycle ownership

Implementation commit `ad7fc362036fa81c70c3e8ac373d7901e7689179` moves the
canonical file-movie title lifecycle behind the selector-first Keynote package
owner. The public surface is
`Package::{slide_movie_title, edit_slide_movie_title, apply_slide_movie_title}`
with `SlideSelector`, `MovieSelector`, and the semantic
`SlideMovieTitle{Edit,Patch,Commit,Diagnostics,Error,LimitKind}` transaction
types. Title `None` is the canonical empty stand-in and `Some(text)` is the
inline title graph; the Wave74 movie-caption surface remains the owner for
caption values.

The title owner admits exact no-ops, existing-title replacement,
stand-in-to-four-object title creation, and active-title-to-fresh-stand-in
removal. Creation allocates the style, info, storage, and placement graph;
removal retains the old graph and UUID registrations while retargeting the
movie to the fresh stand-in. Both transitions update the exact current
Metadata selector, UUID registry, root object-ID watermark, root save token,
and selected component save token, then verify previews, locality, candidate
reopen, and exact inverse artifacts. Shared, aliased, cross-component,
malformed, noncanonical, or otherwise unproven graphs fail closed.

The selected MovieArchive title edge uses the shared hidden
`keynote_movie_caption_codec` Title slot, and the graph writer is typed as
`CaptionGraphKind::Title` with `CaptionGraphProfile::Inline`. Those physical
and generated values remain private; no native ID or wire record crosses the
public facade. The compatibility host now delegates title reads, replacement,
creation, and removal through typed selectors, retiring its raw-ID title
mutators and legacy title graph fallback. Broader movie/media/build/theme
compatibility remains a separate host responsibility.

## 2026-08-25 amendment: Pages body-table dimension semantic ownership

Implementation commit `e624ebab09a891eb9d5921f6ebf8c0e601bf5dff` establishes
the selector-first Pages body-table dimension owner. The public surface is
`Package::{body_table_dimension_size, edit_body_table_dimension_size,
apply_body_table_dimension_size}`, `BodyTableSelector`, the archive-free
`table::dimension::{Dimension, Points, Size}` values, and the corresponding
typed transaction, diagnostics, error, and limit types. `Dimension` selects
one checked row or column position. `Size::Default` retains the native absent
or zero override state, while `Size::Points` requires a strictly positive,
finite value; the two states remain semantically distinct.

This owner changes display row height or column width only. It does not resize
the table, change row or column counts, alter cell addressability, insert or
delete axes, perform automatic fitting or equal distribution, author style
defaults, or perform bulk multi-axis sizing. The hidden neutral
`table_dimension_codec` owns strict borrowed `HeaderStorageBucket` selected
`Header.size` planning and execution with raw-preserving output. The package
owns rooted table/model/bucket selection, cross-component ownership,
lock/dependency proof, exact patch/inverse/apply, preview invalidation,
candidate reopen, and locality verification. Unknown raw fields, balanced
unknown groups, and overlong unknown scalar values remain source-authoritative;
malformed, duplicate, ambiguous, shared, resource-owned, or otherwise
unproven graphs fail closed.

No native identifier, archive or component route, wire value, generated
message, Buffa view, or Prost type crosses the semantic facade. The six raw-ID
Pages dimension methods and their tracked layout owner are retired by the
compatibility cut; other Pages table and document responsibilities remain at
their recorded owners.

## 2026-08-25 amendment: Pages body-footnote insertion/removal semantic ownership

Implementation commit `5dc2ab5337cb61b72f83e369d818829c355d7141`
establishes the selector-first Pages body-footnote lifecycle owner. The public
surface is `Package::{body_footnotes, insert_body_footnote,
edit_body_footnote, apply_body_footnote}`, the archive-free
`footnote::body::{Footnote, Position, Selector}` values, and the typed
`BodyFootnote{Edit,Patch,Commit,Diagnostics,Error,LimitKind}` transaction
types. Insertion accepts one checked UTF-16 body position, text, and an
optional custom mark. An existing edit may replace text or the custom mark, or
clear the complete selected graph. Exact no-ops preserve the source artifact.

The hidden strict `pages_footnote_graph_codec` owns raw-preserving body anchor,
footnote-table, reference-attachment, marker, and storage graph projections
and rewrites. The package owns rooted body/storage selection, ordered semantic
projection, collision-free identifier allocation, global aggregate and
`FieldInfo` ownership census, Metadata UUID/external-reference/watermark and
save-token transitions, preview invalidation, exact package locality,
candidate reopen, patch conflict detection, and exact inverse artifacts.
Malformed known wire, duplicate or missing required values, shared or aliased
owners, ambiguous metadata, resource ownership, cross-component divergence,
and otherwise unproven graphs fail closed. Unknown fields and supported opaque
framing remain source-authoritative.

Insertion registers the new storage object's UUID while the reference and
marker remain archive-owned. When the current `ViewState` component is
present, the transaction adds the exact weak ViewState-to-Document/storage
external edge and advances both selected current component tokens; without
that component, only the exact Document selector is updated. Removal deletes
the selected graph, its storage UUID, and the authorized current external
tuple, advances the selected tokens once, and deliberately retains the root
last-object-identifier watermark. All metadata changes and the native graph
rewrite publish atomically.

No native identifier, archive/member name, component selector, wire value,
generated schema type, Prost message, or Buffa view crosses the public facade.
This owner does not authorize ordinary-body replacement to infer or reclaim an
unattributed footnote graph, and it does not move unrelated Pages text,
annotation, section, table, or media responsibilities.

## 2026-08-25 amendment: Pages header/footer text semantic ownership

Implementation commit `1596d5106ee42fe9238e8d63cc55d39c494d59c3` establishes
the selector-first existing-root Pages header/footer text owner. The public
surface is `Package::{header_footers, edit_header_footer_text,
apply_header_footer_text}` and the archive-free
`HeaderFooterSelector`/`HeaderFooter` vocabulary.
`HeaderFooterSelector` contains a `SectionSelector`, `Template` (`First`,
`Even`, or `Odd`), `Kind` (`Header` or `Footer`), and typed zero-based slot
`Position`; no section, template, storage, component, or archive identifier
is exposed.

The semantic operation reads reachable existing slots, replaces text in one
selected storage, or clears it to an explicitly empty value. Absence and an
empty storage remain distinct. `Some -> Some` and clear-to-empty are the only
changed states admitted by this slice. UTF-16 boundaries and reserved native
markers are validated before staging; exact no-ops retain the source
artifact. Exact patches retain private source/target authorization artifacts,
support conflict-checked application and inverse restoration, and expose
only semantic values and redacted diagnostics.

The package owns rooted `Document -> body -> section -> template -> storage`
selection, exact alias handling, aggregate and `FieldInfo` ownership proof,
Metadata root plus selected current-component save-token updates, preview
invalidation, candidate reopen, and locality verification. Weak `ViewState`
references are preserved. Strong or unspecified external/data references,
ambiguous ownership, and conflicting root data-map ownership fail closed.
The hidden strict `pages_header_footer_codec` validates canonical known
framing and preserves unknown fields and balanced unknown groups; the
text-wire rewrite preserves all unrelated storage bytes. The package does
not create/remove slots or templates, allocate/cull native graph objects,
change inheritance or first/even/odd settings, edit number attachments,
modify margins or section topology, or mutate body, annotation, media, or
text-box graphs.

No native ID, raw storage type, archive/member name, wire field, generated
schema type, Prost message, or Buffa view crosses the semantic facade.

## 2026-08-25 amendment: Pages section-text host-retirement semantic ownership

Implementation commit `507193d3c2ea7c6f6939f47189be5a7b425661c0`
completes the focused host-retirement cut for existing rooted Pages section
text. `SectionSelector` chooses by checked position or semantic name;
`TextPosition` and `TextSpan` express checked section-relative UTF-16
boundaries. `Package` exposes semantic read, whole-section set/clear, staged
span replacement/insertion/deletion, exact patch application, and inverse
restoration without exposing native IDs, object or component routes, ZIP/IWA
artifacts, wire values, generated messages, Prost values, or Buffa views.

The package resolves the unique `Document -> body storage -> ordered section`
chain and authorizes a changed write only in the canonical rooted member. The
text-wire owner strictly validates the source and performs a raw-preserving
splice: known framing and reference tables remain canonical, balanced unknown
groups and overlong unknown scalar values remain byte-exact, section boundary
indexes shift by checked UTF-16 deltas, and unrelated storage fields retain
their original order and framing. Exact ArchiveInfo/MessageInfo metadata,
unknown object headers, unrelated objects/members, Metadata, previews, and
unselected section text are preserved.

The ownership census accepts only the rooted document edge and the exact
aggregate-only type-10015 drawables-z-order consumer emitted by Pages; the
latter orders the same stable storage identity and is preserved. Other
aggregate, `FieldInfo`, data-reference, shared, divergent-alias, dependent-
marker, rootless, or nested ownership fails closed. Because this operation
neither allocates nor removes an object and does not invalidate layout,
Metadata identifiers/save tokens and previews do not change. Section graph
lifecycle, whole-body replacement across structural boundaries, and other
text, table, annotation, or media graphs are outside this owner.

## 2026-08-25 amendment: Numbers table-appearance semantic ownership

Implementation commit `bf01576c090cb508ac0596a5faac01951ef502d0`
adds selector-first `Package::{table_appearance, edit_table_appearance,
apply_table_appearance}` ownership. `Appearance`, `Banding`, `RowSizing`,
`GridlineVisibility`, and `Gridlines` are archive-free semantic values;
`Edit`, `Commit`, `Patch`, `Diagnostics`, `Path`, and typed errors describe an
atomic package transaction without exposing native identifiers or physical
IWA details.

The hidden strict table-appearance codec validates the selected
`TableModelArchive` direct-style edge, bounded `TableStyleArchive`
inheritance, and current plus versioned `StylesheetArchive` registries. It
preserves unknown fields, balanced unknown groups, overlong unknown scalar
values, unselected and versioned registry bytes, and all unrelated object
metadata. Known fields, references, registry entries, inheritance, and
resource limits are canonical and fail closed. A changed edit creates one
copy-on-write style variation, replaces only the selected model edge, appends
the exact stylesheet registry relationship, updates UUID/watermark and root
plus selected current-component save tokens, invalidates canonical previews,
reopens the candidate, verifies locality, and retains an exact inverse.

This owner accepts an existing rooted direct field-3 style route. A nonzero
field-48 preset is opaque and byte-preserved while the direct edge remains
authoritative; preset-only and preset/network lifecycle are unsupported.
The selected style and stylesheet must have one proven current owner and be
co-located; cross-component stylesheet graphs, aliases, shared ownership,
missing or conflicting metadata, locked/dependent tables, malformed
MessageInfo/FieldInfo edges, and ambiguous/versioned/data/root-map ownership
fail closed. Native producer cases where a component-root stylesheet is not
repeated in its UUID map, or where current field-local metadata is absent, are
accepted only when the component identity, aggregate references, strict
registry, and all remaining ownership facts prove the route exactly.

This is appearance replacement, not table/style/preset/network creation or
culling, reset semantics, table topology or content editing, arbitrary
producer-graph repair, or a public native graph API.

## 2026-08-25 amendment: Numbers table-cell Pop-Up Menu semantic ownership

Implementation commit `4ea040b4f9cf97f61ae9a690c6ae592b1b7ac567`
adds selector-first `Package::{table_cell_pop_up_menu_format,
edit_table_cell_pop_up_menu_format,
apply_table_cell_pop_up_menu_format}` ownership. Reads return an optional
archive-free `PopUpMenu`; an `Edit` can set, clear, or reset the value and
produces a conflict-checked `Commit`, exact `Patch`, inverse, redacted
`Diagnostics`, typed `Path`, and typed resource errors. `PopUpMenu`, `Item`,
and `InitialSelection` remain semantic values; no native identifier or physical
IWA representation crosses the facade.

The package resolves the selected sheet, rooted table, exact BNC cell, data
store, tile, format list, control list, cell specification, and type-6206
model. A changed transaction validates exact aggregate and per-key
`FieldInfo` ownership, censuses every selected-archive BNC format/control key,
and proves refcounts before and after the edit. It may reuse a rooted matching
model, copy on write when a sibling retains the old model, create a new model,
or reset and cull the final unreferenced model. Every mutated existing object
and reused model has exact current UUID ownership; new identifiers are reserved
against physical and current/versioned metadata registries, and root watermark
plus every touched current-component save token advance exactly once.

Known fields, references, NIL/string menu values, table-list entries, and
metadata ownership are strict. Supported unknown fields, overlong unknown
scalars, balanced unknown groups, unselected objects and members, and
versioned metadata remain byte-authoritative. Exact cross-component aliases,
opaque or unknown inbound ownership, segmented lists, malformed or dangling
format/control routes, refcount disagreement, missing metadata, and
locked/dependent graphs fail closed. This owner is limited to rooted
same-component Pop-Up Menu graphs; it is not a general Numbers data-format,
table topology, other control, cross-component graph-repair, or public native
graph API.

## 2026-08-25 amendment: Keynote movie-playback semantic ownership

Implementation commit `666ee3ec3be5d7574ebb9324154b550fde83a5f1`
adds selector-first `Package::{slide_movie_playback_settings,
edit_slide_movie_playback_settings, apply_slide_movie_playback_settings}`
ownership. Reads return an optional archive-free `MediaPlaybackSettings`;
an `Edit` replaces the value and produces a conflict-checked `Commit`, exact
`Patch`, inverse, redacted `Diagnostics`, and typed resource errors. Optional
start, poster, loop, and volume values retain the native omitted/present
distinction, while the playback end remains required and validated.

The package resolves an exact slide position and movie source position, admits
only a unique same-component type-3007 movie with exactly one strict
`movie_data` edge and a strict slide-parent/owner route, and refuses non-file,
duplicate, cross-component, or otherwise ambiguous graphs. A changed edit
rewrites only the selected playback projection, validates the full candidate,
rereads the selected semantic value, proves every unrelated object/member
byte-authoritative, and preserves all preview and Metadata members exactly.
Metadata is deliberately independent for this no-allocation scalar edit.

The hidden codec rejects malformed, duplicate, wrong-wire, non-canonical, and
non-finite known playback fields, resolves modern/legacy loop conflicts, and
preserves admitted unknown framing. Its codec boundary retains balanced
unknown groups and overlong unknown scalars; stricter package graph ingress
may fail closed before that selected edge. This owner does not create or
remove movie graphs, replace media/posters, edit geometry, title, caption, or
builds, repair unsupported producer graphs, or expose native graph values.

## 2026-08-25 amendment: Numbers table-sort semantic ownership

Implementation commit `20aca0ede7817e7fbb630338dbb4bcc852d7d500`
adds selector-first `Package::{table_sort_order, edit_table_sort_order,
apply_table_sort_order}` ownership. Reads return `Option<Order>`; an `Edit`
can set a checked `Order`, clear/reset its rules, and produce a conflict-checked
`Commit`, exact `Patch`, inverse, redacted `Diagnostics`, typed `Path`, and
typed resource errors. `Order`, `Rule`, `Scope`, `ColumnIndex`, and `Direction`
are archive-free values; no native identifier or physical representation
crosses the facade.

The owner resolves one exact rooted sheet/table/model, rejects changed edits
to locked tables, and preserves the selected model's unrelated fields,
message metadata, all unselected objects/members, Metadata, and previews.
Absent field 44 and an explicit empty sort marker both read as `None`. Clearing
an absent value is byte-exact; clearing an existing value retains its field-44
marker, scope, and admitted unknown framing while removing only the rules.
Field 45, including any reference tracker, remains opaque and byte-exact.

Known scope/rule fields, rule columns/directions, duplicate columns, bounds,
wire framing, and resources are strict. Admitted unknown overlong scalars and
balanced groups are source-preserved by the hidden codec. This owner edits
persisted configuration only: it does not reorder rows, apply “Sort Now”, own
table storage/UID/formula/comment movement, create/remove tables, repair
unsupported producer graphs, or expose a native graph API.

## 2026-08-26 amendment: Pages body-table sort semantic ownership

Implementation commit `31d5081ca6cd56256e463bee0019e4c7241d6df1` adds
selector-first `Package::{body_table_sort_order, edit_body_table_sort_order,
apply_body_table_sort_order}` ownership. Reads return the archive-free
persisted `Order`; an edit produces a conflict-checked commit, exact patch,
inverse, redacted diagnostics, and typed resource errors. The public semantic
vocabulary is `BodyTableSelector` plus
`table::sort::{Order, Rule, Scope, ColumnIndex, Direction, RowRange}`. No
native identifier, physical member, archive, wire, generated, Prost, or Buffa
value crosses the facade.

The owner resolves one rooted body table, refuses locked or otherwise
unproven table graphs, and changes only `TableModelArchive` field 44. Field 45
(`sort_rule_reference_tracker`) is strictly validated for source authority but
remains opaque and byte-exact. Exact no-op, patch application, inverse,
candidate reopen, semantic readback, and object/member locality are package
operations. The physical PagesEditor apply/reorder executor remains in
`litchi-iwa`; this slice does not reorder rows, mutate cells, alter row/column
UIDs, touch formulas/comments/storage, or create/remove tables.

Unknown framing and unrelated model fields remain source-preserved. Malformed,
duplicate, aliased, locked, or unsupported producer graphs fail closed before
publication. Metadata, UUIDs, save tokens, previews, and unrelated package
members are not mutation targets and remain exact. This is persisted sort
configuration ownership only, not a public physical row-sort or general table
graph API.

## 2026-08-26 amendment: Numbers unified cell-control semantic ownership

Implementation commit `8f804fdc65f5d99a61d8b73503353901edf87488` adds
selector-first `Package::{table_cell_control_format,
edit_table_cell_control_format, apply_table_cell_control_format}` ownership.
Reads return `Option<CellControl>`; an edit can set any of
`CellControl::{Checkbox, StarRating, Slider, Stepper, PopUpMenu}` or clear the
control and produces a conflict-checked commit, exact patch, inverse, redacted
diagnostics, typed path, and typed resource failures. The semantic namespace
reuses archive-free `Checkbox`, `StarRating`, `PopUpMenu`, and checked
`Range`/`DisplayFormat` values; no native identifier or physical value crosses
the facade.

The owner resolves one exact rooted sheet/table/cell, admits the selected
model, tile, format list, and control list only when their graph is unique and
same-component, and strictly validates every mixed control-list entry.
Checkbox, star, slider, and stepper transitions preserve scalar cell values,
use copy-on-write list entries, maintain exact BNC/list refcounts, and cull
unreferenced entries. Pop-Up Menu transitions reuse the existing strict popup
model/metadata lifecycle rather than introducing a second writer. Exact
aggregate/FieldInfo authority, current UUID ownership, metadata collisions,
lock/dependency policy, candidate semantic readback, previews, and unrelated
objects/members are verified before publication.

Known control fields, fixed64 ranges, display-format fields, references, list
keys/refcounts, and wire framing are strict. Admitted unknown overlong scalars
and balanced groups are preserved by the hidden codec. Malformed, segmented,
aliased, cross-component, ownership-ambiguous, or unsupported graphs fail
closed. This slice does not own arbitrary scalar formatting, table topology,
Pages/Keynote control graphs, cross-component repair, or a public native graph
API. A fresh Numbers 14.4 source demonstrated the cross-component boundary,
so native mutation acceptance is explicitly withheld rather than inferred.

## 2026-08-26 amendment: Wave86 Numbers unified cell-control split-read semantics

Implementation commit `3dfe506f4febe7f389db4f60b2988430bbd6038e` records a
narrower Wave86 scope for the existing selector-first
`Package::{table_cell_control_format, edit_table_cell_control_format,
apply_table_cell_control_format}` facade. A strict read may return the
archive-free `CellControl` for an admitted split-component graph, and an
exact no-op produces byte-identical package bytes. No native identifier,
component/member locator, archive, wire, generated, Prost, or Buffa value
crosses the facade.

For the selected split edge, the package checks the current/effective locator,
rejects versioned or conflicting ownership and physical aliases, and reuses
one cached `RegistryFacts`/physical census for the logical read. Selected
`TableModel` sidecar references must occur exactly once in aggregate metadata.
Any explicit `FieldInfo` occurrence must be unique and typed
`ObjectReference`; producer-omitted `FieldInfo` is accepted, and this slice
does not claim an exact `FieldInfo` path. The root Document/TableInfo to
CalculationEngine/TableModel metadata edge is not owned or proven. Same-
component graphs do not require external-edge inspection.

Opaque inbound references remain admissible for read/no-op. Changed
split-component operations, including popup-only changes, fail closed through
`reject_cross_component_write` before native/ZIP candidate publication and
preserve the source. This semantic slice does not own cross-component
copy-on-write, UUID or save-token updates, candidate/locality verification,
or inverse artifacts; no split write is successful or published. It also does
not expand ownership to table topology, scalar data formats, or Pages and
Keynote controls.

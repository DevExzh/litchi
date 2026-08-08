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
checked zero-based selection; native object identifiers and protobuf messages
do not appear in its ordinary signatures. Snapshot cloning shares the IWA
state and never reparses or mutates the source document.

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
IDs or archives leak upward. The IWA adapter is limited to native/protobuf
decoding, payload structural validation, wire patching, graph lookup, and
transactions, and opaque payload decoding remains at that boundary.

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

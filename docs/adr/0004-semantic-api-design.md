# ADR 0004: Semantic API design

- Status: Accepted
- Date: 2026-07-31

## Names and types

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

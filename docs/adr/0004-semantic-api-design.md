# ADR 0004: Semantic API design

- Status: Accepted
- Date: 2026-07-31

## Names and types

Public types use short names inside focused modules. Prefer
`chart3d::BarShape` to `Chart3dBarShape`. Type tags plus unrelated optional
fields are replaced by non-exhaustive data-bearing enums:

```rust,ignore
pub enum Shape {
    AutoShape(auto::Shape),
    Callout(callout::Shape),
    Canvas(Canvas),
    Picture(Picture),
    Chart(Chart),
    Unknown(OpaqueShape),
}
```

The same rule applies to fills, transitions, plots, axes, fields, records, and
other sum types. `Unknown` retains bounded lossless content. Common read methods
live on the enum or narrow static-dispatch traits; the facade does not use boxed
trait objects.

Validated semantic scalars include `Length` (canonical EMU storage), `Percent`,
`Angle`, `Row`, and `Column`. Internal const-generic bounded integers support
domain aliases. Constructors return `Option` or typed `Result` and are usable in
constant evaluation. There are no unchecked constructors.

Use `Default` and struct update syntax only where every resulting combination is
valid. Otherwise use a consuming builder or data-bearing enum. Typestate is
reserved for high-value safety boundaries and must not spread generic noise
through normal CRUD.

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

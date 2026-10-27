# 0375a — W1-E: the style-variation write engine and `litchi-iwa-style`

- Status: landing on branch `uo2/iwa-w1e` in merge-ready slices. Slice 1
  (series-style payload writes) merged on trunk at `0026f92bdb`; slice 2
  (the `litchi-iwa-style` crate and the series-override graph engine) below.
- Date: 2026-10-26.
- Parent: [0375](0375-iwa-w1e-shared-chart-shape-media-engine.md) (W1-E
  read side), plan [0363](0363-iwa-monolith-exit-plan.md) wave W1-E,
  harness and ledger [0369](0369-iwa-exit-parity-harness-and-ledger.md),
  owner decision D1 and the [macOS queue](0363-mac-queue.md).
- Coordinator decisions: the write side of the shared stroke/fill values and
  the style-variation graph engine of series and shape styles belong to
  W1-E; option (1): a new focused crate `litchi-iwa-style` (deps
  `litchi-iwa-archive`, `-core`, `-common`, `-protos`), consumed privately by
  the three leaves and never re-exported by a facade, hosting the graph
  engine and absorbing the per-leaf `drawing_native`/`media_native` copies.
  Media writes stay with the format waves.
- `performance_claim`: every engine route, source and candidate admission
  included, is cheaper than its host route (callgrind ABBA, §4).

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `56e1391dba` | `litchi-iwa-protos` | stroke and fill encoders, the schema-preserving message merge (`chart_style_codec::merge`), the series-style write rules (`property::{series_stroke, series_message, symbol_size, symbol_fill}`) |
| `a8e2415857` | `litchi-iwa-common` | stored forms of colors, strokes and fills (`shape::native::Stored*`) |
| `64fd3919e8` | leaves | `drawing_native` encoders (since folded into `litchi_iwa_style::drawing`) |
| `8d469445ca` | `litchi-iwa` (tests) | 24 Tier A series-style write cases (strokes, fills, symbol fills, symbol outlines) on three formats |
| `4722062a30` | new `litchi-iwa-style` | the crate; `drawing` and `media` absorb the per-leaf copies (removed from `litchi-pages`, `litchi-numbers`, `litchi-keynote`); ADR 0024 amendment; `tools/crate_boundaries.json` package, four edges and migration debt order 11 (`litchi-iwa -> litchi-iwa-style`, deleted with the host in W3); non-iWork gate 18/21; glossary profile `iwa-style` |
| `8ecfb4da11` | `litchi-iwa-protos` | `chart_style_codec::rewrite_series_non_styles` (the chart's `series_non_styles` sparse array as a protobuf encoder writes it) |
| `3e4588dc62` | `litchi-iwa-style` | `graph` (the copy-on-write workspace and registry transitions) and `series` (the series-override graph engine; connection lines, wedge explosions, label distances) |
| `25bf949c8b` | `litchi-iwa` (tests) | 21 graph-engine cases on three formats, cost probe, three mac-queue rows |
| `a2f2dc081b` | docs | this record (first version) |
| `75abec418d` | `litchi-pages` | after the trunk merge: the W1-B body-media lifecycle reads its inventory through `litchi_iwa_style::media` (edge `litchi-pages -> litchi-iwa-style`, ADR 0024 note) |
| `3d253aea5d` | `litchi-iwa-style` | `series::{StyleProperty, style_values, set_style_values}` (series-style writes with the exclusivity rule), `series::{set_symbols, symbols}`, `property::{SymbolMarks, SymbolSizes, Trendlines}` |
| `4e1c6f7de2`, `8882bd2cce` | `litchi-iwa` (tests) | 12 data-symbol and trendline cases, two mac-queue rows; probe extended |
| this commit | docs | this record |

## 2. The engine

### 2.1 `litchi_iwa_style::graph`

A `Workspace` stages one edit over the parsed IWA components an owner
already holds (`graph::Components`; `[(String, Archive)]` implements it). It
clones a component the first time the edit changes it, tracks where every
object lives, inserts and removes objects, creates members (with the source
member to insert before) and deletes members, and edits the
`TSP.PackageMetadata` registry. `finish()` returns `Changes`: the components
to re-encode, the positioned insertions and the deletions; the owner
re-encodes and reassembles its package once, so untouched members keep their
exact bytes.

Registry transitions follow the host byte for byte: the watermark is raised
through `package_metadata_codec::prepare_package_metadata_rewrite` and
lowered in place; object UUID bindings, component external references and
component registrations are appended after the component's (or root's) last
field; removals drop exactly the selected record (component removal through
`rewrite_package_metadata_component_removal`); the next identifier is one
above every live object and every identifier the registry names (the strict
registry visitor). A package without `Index/Metadata.iwa` skips registry
edits, as the host does. New object UUIDs are derived deterministically
(SplitMix64 over the object identifier and a per-edit counter,
collision-checked, RFC 4122 version 4 bits), record 0298 §2.1; the host
draws random ones.

### 2.2 `litchi_iwa_style::series`

`set_values(workspace, chart, property, expected)` is the host's
`set_chart_series_non_style_values` ported operation for operation, over the
workspace instead of `HostPackage`:

- the graph: the chart's `series_non_styles` (count, unique indices and
  objects, one `TSCH.ChartSeriesNonStyleArchive` payload each) and the
  document stylesheet of the first series theme style;
- per series whose value changes: patch in place when this chart alone owns
  the override (registering it with the stylesheet when a styled property
  first gives it a value); copy to a new override when other charts share it;
  allocate when the series had none — styled overrides in the chart
  component, unstyled (geometry-only) ones in a new `TSP.ObjectContainer`
  component with its registration (`document_*_version` 2.0.0, save token,
  `required_package_identifier`), UUIDs and the chart component's external
  references; remove an override the write empties (stylesheet
  unregistration, container member removal, container component deletion
  when it empties);
- then, in the host's order: payload updates, unregistrations, removals,
  insertions, the container, registrations, the chart's sparse array and
  message references (`rewrite_series_non_styles`), UUID unbinding and
  binding, the watermark, the identifier-suffix release, and a read-back.

`series::Property` maps a value to the override payload. `series::property`
holds the Format-panel properties ported from the host:
`ConnectionLines(ConnectionLineFamily)` (styled; the host's field-order
patch, scatter visibility included), `WedgeExplosions` and `LabelDistances`
(unstyled; the `chart_style_codec::property` write rules), `Trendlines`
(styled; the host's trendline patch over `common::chart::series::Trendline`)
and `SymbolMarks` (styled; symbol visibility and shape). Values are the
`litchi_iwa_common::chart` types.

`series::StyleProperty` is the second half: a property stored in the
effective series style (the `series_private_styles` override, else the theme
style), read through the `TSS.StyleArchive.parent` chain and written only to
a style that one chart alone references (the host's `ensure_exclusive`; a
shared style is refused, as the host refuses it). `property::SymbolSizes` is
one; `series::set_symbols` composes marks and sizes as the host's
`set_*_chart_series_symbols` does.

## 3. Parity

`crates/litchi-iwa/tests/parity/series_graph_cases.rs`: the host editors
(`set_{slide,sheet,body}_chart_series_connection_lines`,
`…_pie_wedge_explosions`, `…_pie_label_distances`) against the engine over a
workspace of the same input, on a chart the host added to the native-saved
`chart-data-native.{key,numbers,pages}` (optionally first written through
the host so the engine meets stored overrides). Every case also compares the
host's read of the property on both outputs and strictly reopens both.

| Path | Cases | Tier |
|---|---|---|
| in-place patch (scatter curved → straight) | Pages | A |
| removal with stylesheet unregistration and watermark release (curved → straight, scatter → hidden) | Keynote, Numbers, Pages | A |
| container member removal (one wedge reset), container deletion (all reset) | Keynote; Numbers, Pages | A |
| patch of an exploded wedge's override (label distance) | Numbers, Pages | A |
| data symbols: patch (shape, size; one hidden), removal | Keynote, Pages; Numbers | A |
| trendlines: patch (linear → polynomial → moving average), removal | Keynote, Numbers; Pages | A |
| data-symbol and trendline allocation | three formats each | B |
| styled allocation (line, scatter, radar; and a patch plus allocation) | three formats | B |
| unstyled allocation in a new object container (wedges, labels) | three formats | B |

Tier B differs only by the documented fresh-UUID policy: the new objects'
UUIDs (deterministic in the engine, random in the host), and the ZIP local
layout of the inserted container member and of the members after it (equal
content). 16 Tier A, 17 Tier B; all pass. The public setters stay
`host-only` in the ledger: they belong to the format families (W2), which
wrap this engine with their selectors; three mac-queue rows cover the
engine routes, two more for symbols and trendlines. Earlier slice 1: 24 Tier A series-style write cases
(`series_write_cases.rs`).

## 4. Performance

Callgrind, release with debug info, CPU 3 pinned, Ir, ABBA per format; the
engine route includes source admission (`leaf_reopen`) and candidate
admission (strict reopen), the host route its own reopen-and-verify.

| Probe | Host | Engine | Δ |
|---|---:|---:|---:|
| `series_style_write_cost_abba` (slice 1: strokes, fills, symbol fills) | 9.86 G | 1.03 G | −89.6% |
| `series_graph_write_cost_abba` (styled allocation, container allocation, container release) | 8,477,902,285 | 1,876,741,217 | −77.9% |
| `series_graph_write_cost_abba`, extended with data symbols and a trendline | 15,083,360,713 | 3,164,011,104 | −79.0% |

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind '--toggle-collect=*series_graph_cases::host_series_graph_route*' \
  target/release/deps/parity-<hash> --ignored --exact \
  series_graph_cases::series_graph_write_cost_abba --test-threads=1
# likewise *series_graph_cases::leaf_series_graph_route*
```

## 5. Still to land in 0375a

| Capability | Engine state | Next |
|---|---|---|
| error bars, value labels and their affixes/number formats, pie label visibility | same engine (`Base::Styled`) | port the per-property patches |
| shape-style variations (fill, stroke, shadow, effects, line end, text layout) | `shape_style_codec` readers and encoders | the shape-style variation graph |
| chart drawable codec and reference lines; chart font | 0375 §6 | coordinate with W1-C; values in `litchi_iwa_text` |

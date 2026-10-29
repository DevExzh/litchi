# 0400 — W2-B Pages chart-series symbols facade

- Status: implementation checkpoint on `uo2/iwa-w2b-series-symbols`, based on
  frozen `e0e297a82b0a534dfd5611ced3a9811851edbf76`; coordinator integration,
  final performance qualification and native verification remain pending.
- Scope: Pages body-chart series symbol visibility, shape and optional point
  size through the selector-first facade. Keynote, Numbers, chart fonts,
  chart text, series fills/strokes and `litchi-ops` dispatch are outside this
  slice.
- Ownership: `litchi-pages/src/package/body_chart_series.rs`, its registration
  in `package.rs` and `chart.rs`, plus the dedicated Pages parity and cost
  modules in `crates/litchi-iwa/tests/parity/`. The shared series engine and
  Numbers adapter are reference-only in this work package.

## Public contract

The supported Pages surface is:

```text
litchi_pages::Package::{body_chart_series,
                        edit_body_chart_series,
                        apply_body_chart_series}
litchi_pages::chart::series::{Property, Edit, Commit, Patch,
                              Diagnostics, Error, LimitKind, Refusal}
```

`Property::Symbols(Vec<Option<Symbol>>)` is the only mutation property.
`Edit::set` is the only semantic write verb; host single-series and all-series
aliases are not exposed. Chart selectors remain `BodyChartSelector`; native
object identifiers, archive names, ZIP paths and protobuf values stay private.

The adapter delegates graph work to
`litchi_iwa_style::series::{symbols,set_symbols}` and the common
`SymbolFamily`/`Symbol`/`SymbolSize` values. It admits the workspace census,
series vector and publication capacity before graph work, preserves the
transaction budget through candidate reopen/readback, and maps shared-style,
series-count, unsupported-kind, allocation and resource failures to typed
terminal results. Exact-source patches carry `inverse()`, exact no-op behavior,
source authorization, candidate verification and source-relative locality.

## Differential proof

`crates/litchi-iwa/tests/parity/pages_chart_series_symbol_cases.rs` uses
`test-data/iwork/pages/chart-data-native.pages` plus a host-created Line2d
chart. It contains:

- allocation/set (Tier B, deterministic leaf graph metadata policy);
- patch/rewrite and all-`None` hide/removal (Tier A where bytes permit);
- read-only host-reader versus independent leaf-reader comparison on the same
  host-authored source;
- exact inverse and untouched ZIP/object locality through the shared harness.

The existing W1-E `series_graph_cases.rs` allocation/patch vectors remain
engine differential evidence only; they do not qualify this full Pages facade.

The ignored `pages_chart_series_symbol_cost.rs` probe has separate
`#[inline(never)]` host and leaf wrappers for read, set and hide routes. Each
route measures open, selector resolution, operation, candidate reopen/readback
and serialization in host/leaf/leaf/host order. No performance result is
claimed until pinned release Callgrind capture on the final integrated source.

Native Apple open/save/close/reopen evidence is absent. The four Pages ledger
rows for `body_chart_series_{symbol,symbols}` and
`set_body_chart_series_{symbol,symbols}` are migrated on the host differential
only; their native routes remain pending in
`docs/unified-ops/0363-mac-queue.md`.

## Verification checkpoint

Scoped Pages compilation and the dedicated parity target compile successfully
with the frozen lockfile, four Cargo jobs, incremental disabled and debug info
disabled. The four Pages series-symbol parity cases are the required next
verification checkpoint before any release capture.

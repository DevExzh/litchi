# 0400 — W2-B Pages chart-series symbols facade

- Status: boundary-adoption checkpoint on `uo2/iwa-w2b-series-symbols`,
  synchronized with coordinator trunk `c0c10b86ce82151b730a7e6ee00615f7232d7f74`
  and shared admission/finalization helper tip
  `60d2f93efeb9dd3645f95348bc0970bec43c6e15`.
  Coordinator integration, final performance qualification and native
  verification remain pending.
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
`SymbolFamily`/`Symbol`/`SymbolSize` values. It derives the shared engine's
`Limits` from the residual Pages transaction ledger and charges the complete
`Report`, including references, registry entries and identifiers. Archive
encoding and Snappy bounds, reassembly preparation/execution, and candidate ZIP
ingress are admitted before their allocation-bearing phases; the same
transaction budget continues through candidate reopen/readback. Shared-style,
series-count, unsupported-kind, allocation, archive and resource failures
retain typed terminal results. Exact-source patches carry `inverse()`, exact
no-op behavior, source authorization, candidate verification and source-
relative locality. Publication locality consumes the engine's precise
`MutationScope` (selected members/objects, allowed message fields and planned
additions/removals), rather than deriving authorization from the observed
candidate diff. Plan grouping/copies are fallibly pre-admitted. Changed object
metadata, archive reference transitions and message unknowns are checked, and
unrelated object/member mutations refuse. Staged workspace finalization uses
`finish_requirements` plus residual `FinishLimits` and `finish_bounded`; the
successful `FinishReport` is charged once after publication vectors are
materialized.

## Differential proof

`crates/litchi-iwa/tests/parity/pages_chart_series_symbol_cases.rs` uses
`test-data/iwork/pages/chart-data-native.pages` plus a host-created Line2d
chart. It contains:

- allocation/set (Tier B, deterministic leaf graph metadata policy);
- patch/rewrite and all-`None` hide/removal (Tier A where bytes permit);
- read-only host-reader versus independent leaf-reader comparison on the same
  host-authored source;
- exact inverse and untouched ZIP/object locality through the shared harness.
- production locality controls rejecting an unplanned object mutation, a
  changed unknown chart field, an extra ZIP member, ZIP metadata drift and
  archive-object metadata drift.

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

Scoped Pages compilation and strict Clippy pass with the frozen lockfile, four
Cargo jobs, incremental disabled and debug info disabled; all four Pages
series-symbol parity cases and the five production locality-negative tests
pass. The shared helper's field-5 parent encoding, hard caps and registry-cache
accounting are included in the adopted F tip. Coordinator integration, final
performance capture and native evidence remain pending.

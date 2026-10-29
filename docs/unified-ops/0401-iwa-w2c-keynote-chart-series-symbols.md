# 0401 — W2-C Keynote chart-series symbols facade

## Scope and status

This bounded W2-C slice owns the four Keynote host routes in
`keynote/editor/slide_charts/series_symbol.rs`: all-series and one-series
symbol writes, plus the corresponding all-series and one-series reads. It is
implemented on `uo2/iwa-w2c-series-symbols` from the coordinator root and is
not an integration or W3 claim. Native mutation evidence remains pending in
the macOS queue.

The slice does not change the shared `litchi_iwa_style::series` engine, Numbers,
Pages, unified operations dispatch, or coordinator progress. Symbol fills,
symbol outlines, connection lines, trendlines and other chart families remain
outside this ownership boundary.

## Public contract

The leaf destination is:

```text
litchi_keynote::Package::{slide_chart_series_symbols,
                           edit_slide_chart_series,
                           apply_slide_chart_series}
litchi_keynote::chart::series::Property::Symbols(Vec<Option<Symbol>>)
```

Callers select a slide and chart by semantic position or exact visible name;
native drawable identifiers, IWA objects, component names, ZIP paths and
protobuf values stay private. `Edit::set` is the sole semantic mutation verb
for this family. `Commit`, `Patch`, `Diagnostics`, `Error`, `LimitKind` and
`Refusal` are scoped to the route. A commit reads the immutable source,
applies the shared series graph to a copy-on-write workspace, publishes one
exact package candidate, reopens it, and verifies the selected family, series
count and requested values. A source-relative locality pass requires every
untouched ZIP member to remain byte-identical (apart from the documented root
preview invalidation); exact-source patches support `inverse()` and reject a
different source. A set-to-current value is an exact byte no-op.

The adapter derives a bounded transaction envelope from the existing Keynote
budget, admits graph indexes and archive reassembly before their allocations,
maps graph/archive ceilings to typed `Limit`/`Allocation` failures, and does
not fall back to the legacy host route. The shared engine remains the
implementation authority for mark/shape and inherited size behavior.

## Differential proof

`crates/litchi-iwa/tests/parity/keynote_chart_series_cases.rs` uses the tracked
`test-data/iwork/keynote/chart-data-native.key` fixture and a host-added Line2d
chart. The fixture's host chart is selected by test-only host ID; the leaf uses
slide position 0 and chart selector index 1.

The cases are:

| Route | Case | Tier |
|---|---|---:|
| all-series set | `keynote_slide_chart_series_symbols_allocate_facade_parity` | B |
| one-series set | `keynote_slide_chart_series_symbol_allocate_facade_parity` | B |
| all-series read | `keynote_slide_chart_series_symbols_read_facade_parity` | A |
| one-series read | `keynote_slide_chart_series_symbol_read_facade_parity` | A |

The two read cases directly compare the independent host and leaf readers on
the same populated, host-authored bytes; the generic parity harness's
same-reader checks are supplementary. The control case covers exact no-op
bytes, inverse restoration, stale/out-of-range selectors and an empty staged
value. Set vectors retain the host differential, strict candidate reopen,
inverse patch and untouched-member checks. Tier-B differences are limited to
the existing deterministic leaf metadata policy and invalidated root previews.

## Native and performance boundary

The four route rows are appended to `docs/unified-ops/0363-mac-queue.md` with
`pending` results. No native open/save/close/reopen receipt is claimed, so no
E3/E4 evidence is asserted and the pending queue still blocks W3.

`crates/litchi-iwa/tests/parity/keynote_chart_series_cost.rs` contains four
ignored full-facade ABBA probes (all-series read/set and one-series read/set).
Setup and inverse/output checks are outside the measured wrappers; each route
executes host/leaf/leaf/host with one selected operation per call. The
coordinator must run the release Callgrind qualification on the final
integrated candidate. This slice makes no wall-clock or final integrated
performance claim.

## Scoped verification

The focused parity filter passed five tests (four vectors plus the no-op,
inverse and refusal control), and the Keynote package crate compiled offline.
The complete 53-row gate, final integrated requalification, native queue and
coordinator merge remain pending.


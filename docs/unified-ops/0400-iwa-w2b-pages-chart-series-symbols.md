# 0400 — W2-B Pages chart-series symbols facade

- Status: owner qualification handoff on `uo2/iwa-w2b-series-symbols`,
  synchronized with the shared admission/finalization/construction substrate.
  Runtime adoption, host differential parity and the final source-bound
  Callgrind qualification are complete on source
  `a06fd62966b6a8ec76e8073381ee95ca4bf8513d`; coordinator integration and
  native verification remain pending.
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
materialized. Workspace construction uses the shared source census and bounded
`new_bounded` index path; census/build work and graph-index reservations are
charged separately from the engine report.

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
  changed unknown chart field, an extra ZIP member, unplanned member removal,
  ZIP metadata drift, archive-object metadata drift and registry-object drift.

The existing W1-E `series_graph_cases.rs` allocation/patch vectors remain
engine differential evidence only; they do not qualify this full Pages facade.

The ignored `pages_chart_series_symbol_cost.rs` probe has separate
`#[inline(never)]` host and leaf wrappers for the four measured routes:
`all_symbols_read`, `all_symbols_set`, `single_symbol_read` and
`single_symbol_set`. Each route measures the full facade and executes its
host/leaf/leaf/host ABBA sequence, with two calls to each selected wrapper.
The source-bound CPU5 release capture and raw-call validator are recorded in
`docs/unified-ops/results/0400/integrated-callgrind.json`; all four leaf/host
instruction ratios are below the 1.05 admission bar. The raw captures and
threshold-100 annotations remain outside the repository at
`(deleted-coord-artifact)`.

### Final facade qualification

| route | host Ir | leaf Ir | leaf/host | calls/side | status |
| --- | ---: | ---: | ---: | ---: | --- |
| all symbols read | 1,350,514,913 | 31,932,968 | 0.023645 | 2 | meets |
| all symbols set | 2,795,747,961 | 181,301,099 | 0.064849 | 2 | meets |
| single symbol read | 1,350,775,602 | 31,927,621 | 0.023637 | 2 | meets |
| single symbol set | 3,321,533,012 | 179,613,478 | 0.054075 | 2 | meets |

The receipt binds the archived release ELF SHA-256
`991b304b3b258da6e76d9aa0205dda63929a36d14800ddd68020d6472f03614d`, source
head, lockfile hash, CPU5 plan, raw artifacts and the c063 validator receipt.
The executed capture filenames include the route-qualified
`all_symbols_read`/`all_symbols_set` names; the output-name-only correction
manifest is preserved at
`(deleted-coord-artifact)/capture-correction-v2.json`
(SHA-256
`14f62c8d418b5216c384507422695f1ad0b832b64753fca506a89cacf9a19c02`). It
binds the actual argv, environment, wrapper toggle, raw/log/status/annotation
paths and hashes without changing the approved source, ELF or CPU settings.
This is an instruction-count qualification only; it is not native application
acceptance.

Native Apple open/save/close/reopen evidence is absent. The four Pages ledger
rows for `body_chart_series_{symbol,symbols}` and
`set_body_chart_series_{symbol,symbols}` are migrated on the host differential
only; their native routes remain pending in
`docs/unified-ops/0363-mac-queue.md`.

## Verification checkpoint

Scoped Pages compilation and strict Clippy pass with the frozen lockfile, four
Cargo jobs, incremental disabled and debug info disabled; all four Pages
series-symbol parity cases and eight production locality/mapping tests pass,
including unplanned package-registry object mutation. The shared helper's
field-5 parent encoding, hard caps, registry-cache accounting and bounded
workspace construction are included in the adopted F tip. The shared core's
95 tests, ledger check, crate-boundary check, no-3rdparty check and Pages
glossary profile also pass. Coordinator integration and native evidence remain
pending; no source or runtime changes were made for the performance capture.

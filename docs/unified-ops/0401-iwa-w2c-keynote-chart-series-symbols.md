# 0401 — W2-C Keynote chart-series symbols facade

## Scope and status

This bounded W2-C slice owns the four Keynote host routes in
`keynote/editor/slide_charts/series_symbol.rs`: all-series and one-series
symbol writes, plus the corresponding all-series and one-series reads. It is
implemented on `uo2/iwa-w2c-series-symbols` from the coordinator root. The
source-scoped Callgrind qualification is complete for this leaf branch; native
mutation evidence remains pending in the macOS queue, so this is not a W3 or
native-acceptance claim.

The slice does not change the shared `litchi_iwa_style::series` engine, Numbers,
Pages, unified operations dispatch, or coordinator progress. Symbol fills,
symbol outlines, connection lines, trendlines and other chart families remain
outside this ownership boundary.

## Public contract

The leaf destination is:

```text
litchi_keynote::Package::{slide_chart_series,
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
budget, separately admits the borrowed leaf-component view and the shared
graph construction census/index, and uses `Workspace::new_bounded` for every
symbol read, write, and readback. Construction charges the external census
plus the builder's two internal passes exactly; the bounded workspace exposes
zero construction work to the engine report. Graph/archive ceilings map to
typed `Limit`/`Allocation` failures, and the route does not fall back to the
legacy host path. The shared engine remains the implementation authority for
mark/shape and inherited size behavior.

The package boundary reuses the private chart-font selection/publication seam:
archive encoding uses exact encoded lengths and maximum Snappy bounds, ZIP
publication uses prepared reassembly requirements, and the candidate verifier
checks raw ZIP headers/records, archive metadata and route-specific unknown
fields. Deferred selector-name ceilings remain typed as selector-byte limits;
the malformed control now requires the precise `InvalidSource` refusal rather
than accepting any generic refusal.

The shared series entry points accept the caller residual profile and return a
bounded report/scope. The Keynote facade charges that report once, then keeps
publication, candidate reopen, locality, and patch application on the same
transaction ledger. No default ceiling, fixed headroom, or legacy estimator is
introduced at this boundary.

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
same-reader checks are supplementary. The controls cover exact forward and
inverse patch application, stale `PatchConflict`, hide-all-series (`None`)
readback, nonempty count mismatch, malformed chart payloads, low semantic
budgets including the object-census ceiling, shared series-style dependency
refusal, and a host/leaf vector with
an injected unknown chart field plus explicit unrelated-member locality. Set
vectors retain the host differential, strict candidate reopen, inverse patch
and untouched-member checks. Tier-B differences are limited to
the existing deterministic leaf metadata policy and invalidated root previews.

## Native and performance boundary

The four route rows are appended to `docs/unified-ops/0363-mac-queue.md` with
`pending` results. No native open/save/close/reopen receipt is claimed, so no
E3/E4 evidence is asserted and the pending queue still blocks W3.

`crates/litchi-iwa/tests/parity/keynote_chart_series_cost.rs` contains four
ignored full-facade ABBA probes (all-series read/set and one-series read/set).
Setup and inverse/output checks are outside the measured wrappers; each route
executes host/leaf/leaf/host with one selected operation per call. The
one-series read wrapper calls the host's actual one-series reader and projects
the selected element from the leaf family read; it is not an alias of the
all-series route. The one-series set leaf wrapper performs a full-family read,
modifies the selected index, then performs the full-family read-modify-set.

The final source-scoped release Callgrind receipt is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2c-series-typed-errors-perf-babab049dfd5094f9bcc555e4d2beae7799a9fac/`.
It is bound to source `babab049dfd5094f9bcc555e4d2beae7799a9fac`; its immutable
ELF is 216,072,352 bytes, mode 0555, SHA-256
`b13f21e620dc53ce8a70fa01efcc9aecbe6a84395949837fc6f21636ba9f7757`; the
c063 selected-call validator passed all eight legs. Leaf/host inclusive Ir
ratios were: full read `0.122133149864`, full set `0.170939816892`, one-series
read `0.122080924722`, and one-series set `0.143480833832`. These are
source-scoped instruction diagnostics only: no wall-clock, native, or final
integrated performance claim is made.

## Scoped verification

The focused parity filter covers eight series cases (four route vectors,
patch/no-op and refusal controls, malformed/low-budget, unknown/locality, and
shared-style refusal); the Keynote package crate and shared style unit suite
compiled/tested offline. The four archived release smokes passed, and the
focused chart-font admission/locality regressions remain green. The complete
53-row gate, final integrated requalification, native queue and coordinator
merge remain pending.


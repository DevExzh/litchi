# 0403 — W2-A: Numbers chart-series caller budget

- Status: Numbers chart-series caller-ledger repair is isolated on
  `uo2/iwa-w2a-numbers-series-budget`; current source includes trunk merge
  `59a501492c` and F's reviewed bounded-symbol helper
  `7eaececb6e8307528a73c3da50ba42ca9fef5255`.
- Scope: `crates/litchi-numbers/src/package/sheet_chart_series.rs`, the
  Numbers `ChartBudget` residual series seam in `chart_arrangement.rs`, typed
  mappings required by that seam, and focused Numbers chart-series tests.
  Pages/Keynote callers, shared series engine changes, ledger rows, native
  queue work, and performance captures are out of scope.
- Native evidence: none. No E3/E4 claim is made.

## Caller ledger

The series facade now starts one `ChartBudget` before selector resolution and
threads it through chart selection, bounded chart-archive and metadata reads,
graph workspace preparation, shared `series::Limits` construction, engine
reports, candidate planning, ZIP/Snappy output admission, candidate ingress,
and semantic readback. The budget carries residual references, registry-entry
and identifier axes in addition to wire/work/output/allocation/retained/scratch
axes. No physical `+1024` allowance or fresh process-wide profile is used.

The symbol path adopts F's bounded `series::Limits`/`Report`/`MutationScope`
seam. The facade charges the report once and verifies the selected chart,
changed member names, changed object identifiers/message types, planned new
members/objects, and planned removals before publication. The candidate is
reopened through the existing exact-source publication path and the selected
property is read back through the same budget.

Generic series properties retain their existing semantic engine and style
facades, but now run after the same selector/workspace admission and use the
same candidate output/readback boundary. Existing selector-first APIs,
no-op behavior, inverse patches, exact source authorization and typed
refusals remain unchanged.

## Correctness controls

The existing chart-series regression still proves a same-value no-op and
series-count refusal. It now also changes an error-bar value, reopens and
reads the result, applies the inverse, and proves exact source bytes. The
reference-ceiling control opens the native chart fixture under a lowered
semantic reference profile and requires a typed limit/refusal before graph
publication. Symbol mutation uses the shared mutation-scope report and
readback/locality gate; shared-owner and unknown/reference preservation remain
covered by the accepted F engine tests.

The prior `0402` symbol-only slice remains historical context; it explicitly
called out the missing reusable Numbers caller ledger. This slice closes that
Numbers-side gap but does not qualify the full Numbers facade or performance.

## Verification

- `cargo check --offline -p litchi-numbers --lib` — pass.
- `cargo clippy --offline -p litchi-numbers --lib --all-features -- -D warnings` — pass (vendor warnings only).
- `cargo test --offline -p litchi-numbers --test sheet_drawables -- --test-threads=1` — 14 passed.
- `cargo fmt --all -- --check` — pass.

The coordinator owns the integrated workspace gate, affected fuzz locks,
final facade performance qualification, native queue execution, and any
post-merge review of F's shared helper adoption. No binaries or raw evidence
are committed by this record.

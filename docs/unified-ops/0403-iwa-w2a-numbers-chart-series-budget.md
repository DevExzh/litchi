# 0403 — W2-A: Numbers chart-series caller budget

- Status: Numbers chart-series caller-ledger repair is isolated on
  `uo2/iwa-w2a-numbers-series-budget`; final handoff source is the current
  branch tip `07008d1269` after canonical construction/index admission
  `38a0b13df3`, canonical bounded finish `60d2f93efeb9`, the accepted
  registry-reuse merge `73b7bf0027`, retained-memory policy repair
  `d4fcd3af7f`, and final caller-gate closure.
  It contains implementation
  checkpoint `7573d88827`, shared styled-parent/residual hard-limit fixes
  (`055b609ac1`, `701c2c3f37`), and the malformed-source regression
  (`0bf1083940`).
- Scope: `crates/litchi-numbers/src/package/sheet_chart_series.rs`, the
  Numbers `ChartBudget` residual series seam in `chart_arrangement.rs`, the
  canonical shared `Workspace::construction_requirements`/`new_bounded` and
  `finish_requirements`/`finish_bounded` seams,
  typed mappings, and focused Numbers chart-series/fixture tests.
  Pages/Keynote callers, ledger rows, native queue work, and performance
  captures are out of scope.
- Native evidence: none. No E3/E4 claim is made.

## Caller ledger

The series facade now starts one `ChartBudget` before selector resolution and
threads it through chart selection, bounded chart-archive and metadata reads,
graph workspace preparation, shared `series::Limits` construction, engine
reports, candidate planning, ZIP/Snappy output admission, candidate ingress,
and semantic readback. The budget carries residual references, registry-entry
and identifier axes in addition to wire/work/output/allocation/retained/scratch
axes. The caller derives its aggregate ceilings from the selected physical
archive profile; it does not multiply source size or create a fresh
process-wide profile. Codec residuals reject exhausted allocation, retained,
and scratch axes before codec work. Retained/scratch graph memory follows the
checked archive header-memory policy, so a small input ceiling does not make
valid selector admission fail before the intended candidate limit. Final
graph index construction and final change-vector construction use canonical
shared requirements. Census, recensus, and population are charged from the
canonical explicit `total_work`/`builder_work` requirements before execution;
finalization uses residual `FinishLimits`, fallible
`finish_bounded`, and post-success `FinishReport` charging. No legacy
infallible finish path is used.

The symbol path adopts F's bounded `series::Limits`/`Report`/`MutationScope`
seam. The facade charges the report once and verifies the selected chart,
changed member names, changed object identifiers/message types, generated
fields, object/message metadata references, planned new/removed members and
objects, nested raw ArchiveInfo unknown fields in order, and raw field transitions before
publication. Changed ZIP members preserve static local/central-directory
records; untouched members are compared by raw names, local records,
compressed bytes and physical metadata after candidate reopen. The candidate is reopened through the
prepared exact-source publication path and the selected property is read back
through the same budget.

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
publication. A source-built line-chart fixture exercises the actual Numbers
`Property::Symbols` facade: same-value no-op, changed value, exact inverse,
candidate readback, untouched-member/unknown-byte preservation, normalized
central-directory locality, shared-owner preservation, malformed-parent
refusal, and a source-ceiling refusal before publication. The shared engine's
owner-copy and low-axis controls remain additional evidence; no native Apple
acceptance is claimed. Generic non-symbol series-style paths retain their
legacy writer and are explicitly not claimed as part of this exact
symbol-publication proof.

The prior `0402` symbol-only slice remains historical context; it explicitly
called out the missing reusable Numbers caller ledger. This slice closes that
Numbers-side gap but does not qualify the full Numbers facade or performance.

## Verification

- `cargo check --offline -p litchi-numbers --lib` — pass on `e2d2866d40`.
- `cargo clippy --offline -p litchi-numbers --lib --all-features -- -D warnings` — pass on the final source (vendor warnings only).
- `cargo doc --offline -p litchi-numbers --no-deps` — pass.
- `cargo test --offline -p litchi-numbers --test sheet_drawables -- --test-threads=1` — 15 passed.
- `cargo test --offline -p litchi-numbers --test sheet_chart_font -- --test-threads=1` — 4 passed, including the expanded-candidate typed refusal.
- `cargo test --offline -p litchi-iwa-style graph::tests --lib` — 13 passed, including registry-inspection cache reuse, bounded construction, and bounded finish admission.
- `cargo test --offline -p litchi-iwa-style symbol_tests --lib` — 8 passed.
- `cargo fmt --all -- --check` — pass.
- `python3 tools/iwa_exit_ledger.py --check` — pass.
- `python3 tools/check_crate_boundaries.py` — pass (70 packages, 314 dependency declarations).
- No chart-series-specific fuzz target exists; workspace affected-fuzz locks and
  the integrated full gate remain coordinator-owned.

The coordinator owns the integrated workspace gate, affected fuzz locks,
final facade performance qualification, native queue execution, and any
post-merge review of F's shared helper adoption. No binaries or raw evidence
are committed by this record.

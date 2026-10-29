# 0403 — W2-A: Numbers chart-series caller budget

- Status: Numbers chart-series caller-ledger repair is isolated on
  `uo2/iwa-w2a-numbers-series-budget`. The implementation and probe source
  used for scoped qualification is clean commit
  `c768523ebd91a2c7affe7dc9432f92443f9ac1bd`; the probe-only tip isolates
  correctness/setup helpers from the measured `#[inline(never)]` wrappers.
  The final source also includes the reviewed publication admission and
  typed candidate-error mapping repairs.
  Earlier construction/index admission (`38a0b13df3`), bounded finish
  (`60d2f93efeb9`), registry reuse (`73b7bf0027`), retained-memory policy,
  and final caller-gate repairs remain in its ancestry.
- Scope: `crates/litchi-numbers/src/package/sheet_chart_series.rs`, the
  Numbers `ChartBudget` residual series seam in `chart_arrangement.rs`, the
  canonical shared `Workspace::construction_requirements`/`new_bounded` and
  `finish_requirements`/`finish_bounded` seams,
  typed mappings, and focused Numbers chart-series/fixture tests.
  Pages/Keynote callers, ledger rows, and native queue work are out of scope.
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
Numbers-side gap and adds source-bound `Property::Symbols` facade instruction
qualification; it does not qualify the full Numbers facade or the integrated
trunk.

## Source-bound performance qualification

The four whole-facade `Property::Symbols` routes below were measured on the
clean source `c768523ebd91a2c7affe7dc9432f92443f9ac1bd`, with CPU 7 and two
selected-wrapper calls per host/leaf side. The comparison is inclusive
Callgrind Ir, not wall time; all rows meet the 1.05 regression bar.

| Route | Host Ir | Leaf Ir | Leaf/host |
| --- | ---: | ---: | ---: |
| all chart symbols read | 87,332,917 | 35,314,392 | 0.404365194855 |
| all chart symbols set | 453,463,339 | 192,214,082 | 0.423880092322 |
| single chart symbol read | 87,332,941 | 35,328,295 | 0.404524279103 |
| single chart symbol set | 485,814,985 | 191,137,810 | 0.393437452326 |

The immutable release-debug1 ELF is external at
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-symbols-facade-c768523ebd/bin/parity-8ccb274458659a4c`
(mode 0555, 217,637,848 bytes, SHA-256
`fd5f8d37bf78ab0b9521dc771cd2e93b0580a7b33a39ba46a01a9305e96eb8e1`).
The clean-source build receipt is external at
`.../build-receipt.json` (SHA-256
`54c55dbae68c2b71cd8e18a8edb169937f06f2658bafa8e17a29e310b5578dd0`).
The exact v3 plan is external at
`.../callgrind-plan-v3.json` (SHA-256
`d3a344e0a8ea721215c1780e6fb672e5122839106769a81c4ae6a49b533a9046`),
and the c063 raw-call receipt is
`.../capture-v3/c063-validator-receipt-v3.json` (SHA-256
`80c9699b6be20def9cd9cab7ce2034f90d34b843b1a2dd1ef367a6e29f4378ff`).
The reduced four-row registry input is
`docs/unified-ops/results/0403/integrated-callgrind.json`; raw profiles,
annotations, logs and the binary are not committed. The c063 validator
(`validate_call_audit.py`, SHA-256
`c0637e4304f5fdbfae6fca6b0e9a9c17f7d10cb9a56c7df6846951fa4bdf6b8e`)
verified exact two-call selection, raw summary/function Ir, PROGRAM TOTALS,
and threshold-100 inclusive annotations for all eight legs.

The earlier c3f8 capture set is retained externally as superseded historical
evidence; it is not used for these rows. This is source-bound Numbers facade evidence only. Merging the slice into the
coordinator's current trunk invalidates this qualification; the coordinator
must rerun the final integrated matrix. No wall-time, native acceptance, or
E3/E4 claim is made.

## Verification

- `cargo check --offline -p litchi-numbers --lib` — pass on the final c768 source.
- `cargo clippy --offline -p litchi-numbers --lib --all-features -- -D warnings` — pass on the final source (vendor warnings only).
- `cargo doc --offline -p litchi-numbers --no-deps` — pass.
- `cargo test --offline -p litchi-numbers --test sheet_drawables -- --test-threads=1` — 15 passed.
- `cargo test --offline -p litchi-numbers --test sheet_chart_font -- --test-threads=1` — 4 passed, including the expanded-candidate typed refusal.
- `cargo test --offline -p litchi-iwa-style graph::tests --lib` — 13 passed, including registry-inspection cache reuse, bounded construction, and bounded finish admission.
- `cargo test --offline -p litchi-iwa-style symbol_tests --lib` — 8 passed.
- `cargo fmt --all -- --check` — pass.
- `cargo test --release --locked --offline -p litchi-iwa --features parity-host-oracle --test parity --no-run` — pass for the final c768 source; all four route smokes exited 0.
- `python3 tools/unified_bench_matrix.py --check` — pass; the four source-bound 0403 rows are registered in the Numbers matrix.
- `python3 tools/iwa_exit_ledger.py --check` — pass.
- `python3 tools/check_crate_boundaries.py` — pass (70 packages, 314 dependency declarations).
- No chart-series-specific fuzz target exists; workspace affected-fuzz locks and
  the integrated full gate remain coordinator-owned.

The coordinator owns the integrated workspace gate, affected fuzz locks,
final facade performance qualification, native queue execution, and any
post-merge review of F's shared helper adoption. No binaries or raw evidence
are committed by this record.

# 0408 — Numbers chart-series style gate repair

- Status: bounded runtime repair on `uo2/iwa-numbers-series-gate-repair`, cut
  from coordinator source `3ce2a06774`, synchronized with report-only trunk
  `73944b984b`, and finally merged with the coordinator's non-overlapping font
  fix at integrated tip `a5f00ac8f0f54ab340e6fd7797d1d91eb68e73a6`.
- Scope: the four Numbers chart-series style families only — strokes, fills,
  symbol fills, and symbol outlines. The owning implementation is
  `crates/litchi-numbers/src/package/sheet_chart_series.rs`,
  `sheet_chart_series_style.rs`, and the shared private publication seam in
  `sheet_drawable_graph.rs`. No Pages/Keynote facade, ledger, queue, or
  central-progress change is included.
- Native evidence: none. No E3/E4 or native-acceptance claim is made.

## Diagnosis and repair

All four failing parity routes entered `admit_style_candidate`. That helper
observed 67,195,910 bytes of total work while using
`series_limits().max_output_bytes()` (67,108,864 bytes) as candidate-reopen
work against the transaction's 16,000,000 work ceiling. The refusal therefore
occurred before style publication and was not a real archive-size requirement.

The style writer now supplies one transaction-owned admission object to the
shared drawable publisher. It uses the canonical prepared Archive encoding and
Snappy compression requirements, admits each preparation chunk and execution
requirement before allocation, admits physical reassembly before the reader is
constructed, charges reassembly execution from the exact prepared output
requirements, and preflights the candidate ZIP-directory/member, Snappy, and
IWA archive/header ingress phases before typed candidate reopen. Candidate
reopen maps archive, detection, common, semantic, limit, allocation, and
arithmetic failures through the chart-series error contract. The graph
publisher reserves its staging vectors, charges their retained/work/scratch
envelopes, and uses the core archive clone-inventory visitor before cloning an
edited component; clone arithmetic is reported as a typed overflow. Legacy
drawable callers retain their established encoder through a no-op admission
implementation; they do not inherit the chart-series style route.

No ceiling was raised, source-size multiplier was added, or fallback to the
legacy host was introduced. Existing selector, no-op, inverse, candidate
reopen, and locality behavior remains in the surrounding transaction; the
existing `publish_changes` path performs the style readback under the same
budget after the shared publisher returns. The style reader/rewriter's
pre-existing `Limits::DEFAULT` codec calls remain outside this narrow
publication-admission repair and are not claimed as a complete style-parser
budget migration.

## Differential proof

The host-backed parity target was run against
`iwork/numbers/chart-data-native.numbers` with the `parity-host-oracle` feature.
The focused filter selected all four affected cases; each case exercises the
existing read/set/reset vectors and host-vs-leaf candidate readback. All five
filtered tests passed (one read case plus the four style-family cases), and
each family reported its changed `Index/CalculationEngine.iwa` member and
matching host/leaf object counts. This is Tier A differential evidence for the
fixture only; it is not native application acceptance evidence.

## Source-bound performance qualification

The four complete-facade style routes were measured on the clean pre-merge
source `3dd6a6fee751253f3e061eb6229f9c02233637dc` using the immutable
release-debug1 parity ELF, CPU 7, and two selected wrapper calls per host/leaf
side. The Callgrind total is inclusive Ir; each row is a source-built fixture
route, not an engine-only measurement. All rows are below the 1.05 regression
bar.

| Route | Fixture | Host Ir | Leaf Ir | Leaf/host |
| --- | --- | ---: | ---: | ---: |
| chart series fills | Column2d chart fixture | 266,012,588 | 58,927,439 | 0.221521242446 |
| chart series strokes | source-built Line2d chart | 254,696,939 | 58,755,837 | 0.230689215311 |
| chart series symbol fills | source-built Line2d chart | 266,043,359 | 58,818,332 | 0.221085511103 |
| chart series symbol outlines | source-built Line2d chart | 255,539,602 | 58,734,393 | 0.229844581976 |

The archived ELF is external at
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-series-style-perf-3dd6a6fee7/bin/parity-3dd6a6fee7-release-debug1`
(0555, 219,623,840 bytes, SHA-256
`f0fd836369ebf3346183dcb94ae29159a03852d810eff748cd86989a7d65b043`). The
fresh capture manifest is external at
`.../capture-manifest-v1.json` (SHA-256
`fe7040fd9c03f71c3ebd7c9b9a495e842981ea81189c8617a3f93cafc2295906`), with
the compact plan at `.../callgrind-plan-v1.json` (SHA-256
`ace539ce39cfdfce33ad07f97dfc646996e8c39114bc530c14efc2b4399a8b4e`). Raw
Callgrind files, stdout/stderr, exit statuses, threshold-100 annotations,
build logs, and artifact hashes remain outside the repository under
`.../capture-v1/`; all eight legs exited 0. The c063 validator used for the
final audit is external at
`/home/zhuhe/code/litchi-worktrees/.coord/resume-20261028/perf/iwork-probe-isolation-final/validate_call_audit.py`
(SHA-256
`c0637e4304f5fdbfae6fca6b0e9a9c17f7d10cb9a56c7df6846951fa4bdf6b8e`).
The reduced four-row summary is committed at
`docs/unified-ops/results/0408/integrated-callgrind.json`; raw profiles and
the ELF are not committed.

These captures predate the coordinator's non-overlapping font merge, so they
are scoped evidence for this WP rather than final integrated qualification.
The coordinator must rerun the matrix on the final integrated trunk.

## Verification

Commands used with the capped offline environment (`CARGO_BUILD_JOBS=4`,
`CARGO_INCREMENTAL=0`, dev/test debug info disabled, and the shared
`gate-trunk/target` cache):

- `cargo check --offline -p litchi-numbers --lib` — pass.
- `cargo clippy --offline -p litchi-numbers --lib --all-features -- -D warnings`
  — pass on the final physical-ingress source (vendor warnings only).
- `cargo test --offline -p litchi-iwa --features parity-host-oracle --test parity numbers_chart_series_style_cases::numbers_sheet_chart_series_ -- --nocapture` — 5 passed after the final clone-inventory tightening.
- `cargo test --release --locked --offline -p litchi-iwa --features parity-host-oracle --test parity --no-run` — pass for the source-bound performance ELF.
- Four archived-ELF ignored ABBA smokes (`chart_series_fills`,
  `chart_series_strokes`, `chart_series_symbol_fills`,
  `chart_series_symbol_outlines`) — 4 passed.
- Direct `rustfmt --edition 2024 --check` on the three owned Rust files — pass.

The full workspace gate, affected fuzz locks, final integrated performance,
native queue execution, and `PROGRESS.md` remain coordinator-owned. Full
SourceCatalog storage census and legacy style-codec aggregate admission remain
pre-existing debt outside this fixed-cap publication slice; this record does
not claim a complete style-parser budget migration. No binary or large
evidence file is committed.

## Coordinator acceptance

The final clean handoff is `1c997abf13fd6382cad25a9a35202e26e3aead45`,
including the integrated font repair. Additional tests exercise the real
`ChartBudget` and `StylePublicationAdmission` with exhausted allocation and
work budgets. Both return the corresponding typed limit before invoking the
edit callback, with exact source bytes preserved. They passed with strict
Clippy and formatting. This is allocation/work refusal coverage; it does not
claim separate output-limit or full catalog-storage qualification.

Post-font-merge check, strict Clippy, rustdoc, ledger, no-3rdparty and five
parity tests passed. External receipts are under
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w2a-series-style-perf-3dd6a6fee7/post-merge-gates-54b0e44f79/`;
`low-budget-controls/tests.log` has SHA-256
`b76d5f0dbdd2305b2d553b5cf452576caecc527115e4e85b977a7d727fadd4cb`.

After capture, the coordinator independently audited all eight raw legs and
ran the pinned c063 validator against separately derived input metadata. The
raw audit is `coordinator-raw-audit-v1.json` (SHA-256
`a62d7bf39d095a35382b47f25b948f62e5a93fac19442684b76ab63b0fc3a91c`);
the actual machine-generated receipt is `audit-v2/derived-v2.json` (SHA-256
`19de2a1ef67c937956cf68e1e6ee85347c2bc0f644155d595b383197f75ed76c`)
under the same external artifact root. All eight legs have exactly two
selected calls and full-function Ir equal to raw summary and PROGRAM TOTALS.
The source-scoped CPU 7 evidence remains separate from the final integrated
CPU 3 matrix, which adds these four set routes to the existing 119 routes.

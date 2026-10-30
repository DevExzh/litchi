# 0408 — Numbers chart-series style gate repair

- Status: bounded runtime repair on `uo2/iwa-numbers-series-gate-repair`, cut
  from coordinator source `3ce2a06774` and synchronized with the report-only
  trunk tip `73944b984b` before handoff.
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
requirements, and preflights the candidate ZIP-directory/member ingress plan
before typed candidate reopen. Candidate reopen maps archive, detection,
common, semantic, limit, allocation, and arithmetic failures through the
chart-series error contract. The graph publisher reserves its staging vectors,
charges their retained/work/scratch envelopes, and uses the core archive
clone-inventory visitor before cloning an edited component; clone arithmetic
is reported as a typed overflow. Legacy drawable callers retain their
established encoder through a no-op admission implementation; they do not
inherit the chart-series style route.

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

## Performance boundary

No performance qualification is claimed for this repair. The old route
refused before meaningful style publication work, so historical W2-A symbols
or Numbers facade measurements do not qualify these families. A future
source-bound instruction probe should cover the existing fills route and
separate strokes, symbol-fills, and symbol-outlines routes, with setup and
proof helpers excluded from measured wrappers. The coordinator must rerun
that matrix on the final integrated trunk.

## Verification

Commands used with the capped offline environment (`CARGO_BUILD_JOBS=4`,
`CARGO_INCREMENTAL=0`, dev/test debug info disabled, and the shared
`gate-trunk/target` cache):

- `cargo check --offline -p litchi-numbers --lib` — pass.
- `cargo clippy --offline -p litchi-numbers --lib --all-features -- -D warnings`
  — pass on the clone-inventory source before the subsequent typed-ingress and
  staging-charge follow-up; rerun on the final clean source.
- `cargo test --offline -p litchi-iwa --features parity-host-oracle --test parity numbers_chart_series_style_cases::numbers_sheet_chart_series_ -- --nocapture` — 5 passed after the final clone-inventory tightening.
- Direct `rustfmt --edition 2024 --check` on the three owned Rust files — pass.

The full workspace gate, affected fuzz locks, final integrated performance,
native queue execution, and `PROGRESS.md` remain coordinator-owned. No binary
or large evidence file is committed.

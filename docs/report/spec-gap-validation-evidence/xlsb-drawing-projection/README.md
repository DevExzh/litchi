# XLSB explicit drawing projection and CRUD measurement

This batch adds an explicit owned cell/catalog projection that skips typed
SpreadsheetDrawing decoding. Ordinary XLSB constructors remain eager. The
underlying OPC package still loads and retains raw drawing, chart, image, and
relationship content. `Workbook::drawings_parsed()` distinguishes a skipped
inventory from an eager workbook with no drawings.

Workbook publication preserves that choice across cell values, workbook
structure and in-memory patches, calculation chains, sparklines, cell watches,
and raw OPC edits. Scalar cell transfer treats donor drawings as opaque and
retains the target projection. Typed drawing authoring and transfer refuse a
skipped target.
Durable workbook patches retain their existing wire format: callers explicitly
choose `WorkbookPatch::from_bytes_without_drawing_parse` (or its external-link
limits variant) to decode opaque drawing content. Ordinary patch decoding stays
eager. The choice is not trusted from serialized input.

The new integration regressions exercise referenced and unrelated malformed
drawings, exact no-op paths, scalar changes, sheet rename, durable patch decoding,
and retained package content. The benchmark also checks package-level and
per-part relationships, inert VBA/opaque members, and save/reopen semantics.
These tests establish library-level preservation and validation behavior, not
native Office acceptance or complete drawing/chart support.

The opt-in `xlsb_crud` harness now distinguishes `owned`, `owned_direct`,
`owned_without_drawings`, and `source_backed`. Restricted case combinations
fail before corpus loading. Selected-cell coordinates/values and complete scan
contents are checked; full scan replay occurs outside timing. Source read and
part-cache counters must remain stable across samples. The fixture geometry
gate is separate from operation semantics and is not an allocation-level proof
against rectangular expansion.

## Validation and reproduction

Final checks passed on unchanged source hashes: 815 XLSB tests (one existing
ignored test), nine harness tests, strict Clippy for both scopes, XLSB rustdoc
with warnings denied, Rustfmt, and the diff whitespace check.

`gates/` contains the final commands, logs, source hashes, and workspace lockfile.
The harness uses the tracked `tools/perf-baseline/Cargo.lock`. To reproduce the
library gates from a fresh checkout, copy `gates/workspace-Cargo.lock` to the
workspace root as `Cargo.lock`, then run the commands in the gate receipt with
Rust 1.95.0. Use a disposable `CARGO_TARGET_DIR`, `CARGO_INCREMENTAL=0`, and
`CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0` to bound build artifacts.

`verification.json` records the independent raw-sample, gate, counter,
source-hash, and binary-hash checks, plus the observed local CPU environment.

`performance/` contains fresh release measurements and their reproduction
commands. Eager, skipped, and source-backed paths have different semantic and
validation scopes; their timings do not establish an equivalent-work speedup.
Cache retained bytes are point-in-time cache gauges, not total heap usage or
peak RSS. This batch makes no allocation reduction, native acceptance, or
program-wide 10x performance claim.

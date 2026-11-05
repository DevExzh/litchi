# 0360ee — XLSB merged cells and data validation raw policy

- Status: source and consumer migration complete on `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Base: `8d59abb26f` (the 0360ed named-ranges slice).
- Scope: `crates/litchi-xlsb` merged-cell and data-validation codecs, retained owner consumers, tests and this record.
- Policy: every raw entry point in this slice requires an explicit `raw::Limits`; no default or `*_with_limits` compatibility twin was added.

## Changes

Merged-cell wire operations now use `MergedCell::{parse,serialize}(.., raw::Limits)`. Fixed-width payload admission and range validation return the crate's typed wire error. Workbook package inspection, merged-cell edits, cell-value durable patches and worksheet writers pass the retained raw policy; the standalone durable patch wire uses the explicit default profile because it has no opened owner from which to inherit a policy. Unknown worksheet records and untouched parts remain byte-preserved; the integration suite covers insertion alongside an unknown record and unchanged sibling/package parts.

Data-validation classic and extension-14 parsers, collection settings, list-formula parsing, record writers and the durable edit encoder all take the required raw profile. Payload and UTF-16-unit ceilings are checked before semantic allocation and on output. `CellsReader` stores the owner raw profile and threads it through merged-cell and validation records. Workbook source/cursor, worksheet feature, edit and writer paths now pass that profile rather than silently using `raw::Limits::DEFAULT`. `raw::Encoder::limits()` exposes the already-admitted writer profile to worksheet serializers.

The strict tests cover exact and one-under payload limits for fixed-width merged cells and data-validation settings/list data, exact and over-limit UTF-16 strings, output payload/string limits, malformed validation settings, and merged-record unknown-byte preservation.

## Validation

Commands were run with `(deleted-coord-artifact)` and `CARGO_BUILD_JOBS=1` where applicable:

- `cargo fmt --all -- --check`: exit 0.
- `cargo test -p litchi-xlsb --lib writer::data_validation::tests::raw_limits_bound_writer_strings_and_payloads --offline`: 1 passed.
- `cargo test -p litchi-xlsb --tests --offline`: exit 0; 744 unit tests plus all integration targets passed.
- `cargo clippy -p litchi-xlsb --all-targets --all-features --offline -- -D warnings`: exit 0.

Full-test log (kept outside the checkout): `xlsb-merged-dval-final-tests.log`, 114149 bytes, SHA-256 `5b7a9fd751e0ece9e8bbdc90f430f383af24527743254e8fdaadacc2d1b820d6`. Clippy log: `xlsb-merged-dval-final-clippy.log`, 172 bytes, SHA-256 `4a944caea790f52354b9249877080f4d24a58dc35d2585701a07e032fe0f4677`. Both are under `(deleted-coord-artifact)`; no evidence or binary is committed.

## Current raw-audit disposition

Closed in prior or this slice, verified against current signatures:

- comments and calculation records: 0360ea;
- styles, including `StylesTable::from_bytes(bytes, raw::Limits)`, alignment and border parsers: 0360eb;
- auto-filter, page-break and cell-watch worksheet paths: 0360ec;
- named ranges: 0360ed;
- merged cells and data validation: this 0360ee slice.

Still open in the current source and intentionally left for later 0360e slices:

- `pivot_view::Part::from_bytes(Vec<u8>)` in `src/pivot_view/codec.rs:25`, plus its package and writer consumers;
- formula wire/parser entry points (`ParsedFormula`, `Parser`, array/shared formula records);
- hyperlinks public parse/serialize wrappers and their codec writer;
- conditional-format binary records/edit parser and writer headers;
- shared-workbook users/headers/log codecs;
- slicer and timeline cache/view codecs and package consumers;
- threaded-comments byte codecs;
- host record/part codecs for pivot tables, chartsheets, connections, tables, scenarios, sheet views, drawings, web-extension bindings, data-model and XML-map parts;
- remaining raw formula/complex-XML families and any owner-policy propagation that those paths expose.

This list is based on the signatures in the current tree, not the historical inventory: styles and data validation are removed from it, while pivot-view is explicitly retained. The separate 0360b redaction, source-backed metadata replay, OOXML/XLDM typed-origin and lexical-only OPC follow-ups remain unresolved coordinator work.

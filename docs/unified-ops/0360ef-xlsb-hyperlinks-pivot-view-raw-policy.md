# 0360ef — XLSB hyperlinks, pivot view and binary index raw policy

- Status: source and consumer migration complete on `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Base: `54055165f3` (the 0360ee merged-cell and data-validation slice).
- Scope: `crates/litchi-xlsb` hyperlink codec and model, pivot-view part parser, worksheet binary index, their retained-owner consumers, tests and this record.
- Policy: every raw entry point in this slice requires an explicit `raw::Limits`; no default or limit-less twin remains.

## Changes

- `Hyperlink::{parse, serialize}(.., raw::Limits)` is the one public `BrtHLink` codec. The historical infallible `serialize()` (no limits, no string checks) and the duplicate checked `try_serialize()` are deleted, not aliased. `parse` admits the payload against the caller's payload ceiling before decoding and decodes the wide strings under the caller's string-unit ceiling; `serialize` checks every output string and the payload.
- Consumers pass the retained owner policy: the sheet-features reader (`sheet/features.rs`), the hyperlink editor (`hyperlinks/edit.rs`, read and encode through `Staged::raw_limits`), `CellsReader` (`host/cells_reader/codec.rs`, the reader's retained profile), sheet-removal dependency checks (`cell_values/structure.rs`, the workbook's retained profile) and the worksheet writer (`Encoder::limits()`, the writer's admitted profile).
- `pivot_view::Part::from_bytes(Vec<u8>, raw::Limits)`: the hidden `PIVOT_VIEW_LIMITS` constant is gone. Record headers and the `BrtBeginSXView.irstName` string are admitted under the caller's profile; the MS-XLSB ceilings (`MAX_PIVOT_TABLE_PART_BYTES`, record count, 255-unit view name) remain format invariants in addition to it.
- `WorksheetBinaryIndex::from_parts_default` is deleted; `from_parts(index, worksheet, raw::Limits)` is the only binder (package and source consumers already passed the owner profile).

## Tests

New strict tests: `hyperlinks::tests::raw_limits_bound_hyperlink_input_and_output` (one-under payload refusal on input, string-unit refusal on output) and `pivot_view::tests::raw_limits_bound_view_payload_and_name` (one-under `BrtBeginSXView` payload refusal; one-under name-unit refusal). Existing round-trip, truncation and index-binding tests move to the explicit profile.

## Validation

Commands were run with `(deleted-coord-artifact)`:

- `cargo fmt --all -- --check`: exit 0.
- `cargo test -p litchi-xlsb --all-features --offline`: exit 0; 1,332 tests passed, 0 failed.
- `cargo clippy -p litchi-xlsb --all-targets --all-features --offline -- -D warnings`: exit 0.

Logs kept outside the checkout under `(deleted-coord-artifact)`: `xlsb-0360ef-tests.log` (124,574 bytes, SHA-256 `c728810cf693a9bcc61dfe835ce9fc664d37e39ee668f896c9938e96d8f0311a`) and `xlsb-0360ef-clippy.log` (172 bytes, SHA-256 `4beccec7aef354d0dd9beaddb607d9f1801fcffd74174247bcea7e8a2fe67952`).

## Performance

The hyperlink parser gains one payload-length comparison before decoding; the pivot-view parser swaps one constant profile for the caller's. Neither adds allocation. Native timing is pending with the rest of W4b (see the W4b qualification record).

## Remaining 0360e surfaces after this slice

Formula wire/parser entry points; conditional-format binary records and `conditional_formatting/edit.rs`; alignment/border parsers; slicer, timeline, shared-workbook and threaded-comment codecs; host part parsers (chartsheet, connections, pivot cache, table, scenarios, sheet view, web-extension bindings, drawing, data model); and the drawing-image limits carrier. Later 0360e slices close them.

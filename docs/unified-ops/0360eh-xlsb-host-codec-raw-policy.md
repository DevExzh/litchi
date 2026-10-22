# 0360eh — XLSB host part codecs raw policy

- Status: source and consumer migration complete on `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Base: `805ceee4df` (the 0360eg formula and conditional-format slice).
- Scope: `crates/litchi-xlsb` package-public host codecs (`package::{chartsheet, connections, pivot, scenarios, sheet_view, table, web_extension_bindings}`), the BIFF12 slicer and shared-workbook codecs, the shared record walker, their package loaders, editors and writers; tests and this record.
- Policy: every public raw entry in this slice requires an explicit `raw::Limits`; no default, alias or limit-less twin remains.

## Public API after this slice

- `package::chartsheet::parse_chart_sheet_part(data, name, state, raw::Limits)`.
- `package::connections::{parse_connections_part(data, raw::Limits), package::load(&OpcPackage, raw::Limits)}`. `package::validate_graph` (a second name for `load`) and `connections::validation::graph` (a third) are deleted; `load` validates the complete graph.
- `package::pivot::{parse_pivot_cache_definition, parse_pivot_cache_records}(.., raw::Limits)`.
- `package::scenarios::{parse_worksheet, replace_worksheet}(.., raw::Limits)`.
- `package::sheet_view::{parse_view, parse_pane, parse_selection, write_view_payload, write_pane_payload, write_selection_payload}(.., raw::Limits)`; inputs and outputs are admitted against the payload ceiling.
- `package::table::{parse_table_part, parse_table_part_rel_ids}(.., raw::Limits)`.
- `package::web_extension_bindings`: `Binding::{new, parse_payload, to_payload}`, `parse_xlsb_web_extension_bindings`, `write_xlsb_web_extension_bindings`, and `Snapshot::{read, read_with_external_sheet_indices, read_with_package_bindings}` take the profile. `Snapshot` retains it, and its `Draft` and `Patch` reuse it for validation, preview, apply and inverse. The aliases `Snapshot::parse`, `Snapshot::bytes`, `Draft::source`, `Draft::replace_at` and the free `read` / `apply` twins of `Snapshot::read` / `Patch::apply_bytes` are deleted (no caller outside the module).
- `slicer::{parse_cache, write_cache, parse_views, write_views}(.., raw::Limits)` and `slicer::package::{load_caches, load_views}(.., raw::Limits)`.
- `shared_workbook::codec::{parse_users, write_users, parse_headers, write_headers, parse_log, write_log}(.., raw::Limits)` and `shared_workbook::package::load(&OpcPackage, raw::Limits)`; the module ceilings (`MAX_PART_BYTES`, `MAX_STRING_UNITS`) are intersected with the caller's profile.
- `connections::Snapshot::read` and `shared_workbook::Snapshot::read` become `pub(crate)` and take the owner profile; `connections::Snapshot::load` (an alias of `read`) is deleted. The owner entries are `Package::{connections, edit_connections, shared_workbook, edit_shared_workbook}`.
- `xml_maps::SemanticLimits` gains `raw: raw::Limits` (default `raw::Limits::DEFAULT`), matching `data_model::SemanticLimits::raw`; the XML-maps reader passes it to the table and connection streams it reads.

Internally, `RecordWalker::{new, recording}` take the profile and expose `limits()`, so every walker-based parser (tables, connections, chart sheets, pivot caches and records) admits records and decodes payload strings under it instead of `raw::Limits::DEFAULT`. A crate-private `raw::Limits::{admit, admit_units}` pair replaces the hand-written admission helpers of 0360eg.

## Consumers

Owner paths pass their retained profile: the opening loader (`workbook/package.rs`, `self.raw_limits`), the source-backed workbook and its cell cursor (`inner.raw_limits`), `Workbook` accessors for slicers and scenarios, `CellsReader` (sheet views, web-extension bindings), the scenario, slicer, connections and shared-workbook editors (`Staged::raw_limits`, bound before each staging closure), the data-model reader (`data_model::SemanticLimits::raw`) and the XML-maps reader and transaction (`xml_maps::SemanticLimits::raw`). Writers pass `Encoder::limits()`. The `Builder` worksheet and connections writers have no opened owner and use the documented default authoring profile explicitly.

## Tests

New strict tests: `package::table::tests::raw_limits_bound_table_records_and_strings` (one-under payload and string-unit refusals on a table stream), `slicer::tests::raw_limits_bound_slicer_cache_strings_and_output` (exact round trip, one-under string refusal on parse and write), `web_extension_bindings::tests::raw_limits_bound_binding_payload_and_app_ref` (exact round trip, one-under payload and appRef-unit refusal on parse, one-under payload refusal on output) and `..::raw_limits_bound_sheet_view_payloads` (exact and one-under on parse and write). Existing tests move to the explicit profile.

## Validation

Commands were run with `/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb-raw-policy/env.sh`:

- `cargo fmt --all -- --check`: exit 0.
- `cargo test -p litchi-xlsb --all-features --offline`: exit 0; 1,338 passed, 0 failed.
- `cargo clippy -p litchi-xlsb --all-targets --all-features --offline -- -D warnings`: exit 0.
- `cargo check -p litchi-ops -p litchi -p litchi-cli --all-targets --all-features --offline`: exit 0.

Logs under `/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb-raw-policy/`: `xlsb-0360eh-tests.log` (121,831 bytes, SHA-256 `d87e57b1a0a0b02f535d93e4908d1c4699bad816ed9f34d0821f2cd22d9eabda`) and `xlsb-0360eh-clippy.log` (171 bytes, SHA-256 `d54cc455049d1b9d7ba8009f096dfd41efac50265e9d666c5ce63e8cde28a7dc`).

## Performance

The walker stores one extra `Copy` profile; standalone payload parsers gain one length comparison. No allocation is added. Native timing is pending with the rest of W4b.

# 0360eg — XLSB formula and conditional-format raw policy

- Status: source and consumer migration complete on `uo2/w4b-xlsb-raw-policy`; native performance qualification remains open.
- Base: `814be893ec` (the 0360ef hyperlink, pivot-view and binary-index slice).
- Scope: `crates/litchi-xlsb` formula codec (`formula/`), conditional-format binary records, writer and `conditional_formatting/edit.rs`, the cell-formula constructor, and every consumer of those APIs; tests and this record.
- Policy: every public raw entry in this slice requires an explicit `raw::Limits`; no default or limit-less twin exists.

## Public API after this slice

Formula (`litchi_xlsb::formula`):

- `Parser::{new, with_extra, with_base_cell, with_base_cell_and_extra}(.., raw::Limits)`. `parse` admits both the token and the ancillary stream against the payload ceiling before decoding, and checks every `PtgStr`, `SerStr` and `PtgExtraList`/`SxOs` string against the UTF-16-unit ceiling before decoding it.
- `ParsedFormula::{parse, to_bytes}(.., raw::Limits)`: both declared stream lengths are admitted before they are copied; the output, with its two length prefixes, is admitted.
- `Group::{parse_array, parse_shared, to_record_data}(.., raw::Limits)`, `Range::parse_binary(.., raw::Limits)` (fixed 16 bytes; the profile is validated) and `Token::to_extended_binary` / `TableReference::to_extended_binary(raw::Limits)` (table and column names are checked against the string ceiling, the `RgbExtra` output against the payload ceiling).
- `formula::Error::Wire(raw::Error)` carries the refusal; package and structural-edit error mappings translate it to `PackageError::Wire`.
- `cell_values::CellFormula::new(flags, tokens, ancillary, raw::Limits)`: the detached constructor no longer applies a hidden `raw::Limits::DEFAULT` encoded-size ceiling.

The MS-XLSB ceilings (`MAX_CELL_FORMULA_BYTES`, 255-unit `SerStr`, 1–255-unit table names) remain format invariants in addition to the caller's profile.

Conditional formatting (`litchi_xlsb::conditional_formatting`):

- `Value::{parse, parse_with_context, parse_extension14, parse_extension14_with_context, serialize_extension14}`, `Color::{parse, to_bytes, parse_extension14, serialize_extension14}`, `Bar14::{parse_header, serialize_header}`, `Icon::{parse, serialize}`, `IconSet14::{parse_header, serialize_header}`, `Rule::{parse, parse_with_context, parse_extension14, parse_extension14_with_context, serialize_extension14}`, `Formatting::{parse_extension14_header, parse_extension14_header_with_base, serialize_extension14_header}`, `parse_classic_header` and `parse_rule_extension_guid` all take a required `raw::Limits`. Fixed-width records validate and admit their fixed size; variable records admit the payload when their cursor is created, check every nullable string's declared unit count before allocating it, and admit every serialized payload. `serialize_rule_extension_guid([u8; 16]) -> [u8; 20]` is a constant-width constructor over a typed GUID, not byte ingress, and keeps its signature.
- `edit::DifferentialStyle::{parse, encode}(.., raw::Limits)`: payload admission on both sides; font names checked against the string ceiling.

## Consumers

Owner paths pass the retained profile: `CellsReader` (`self.raw_limits`) for cell formulas, formula groups, conditional-format records and cell construction (`Cell::new_formula_binary` / `new_grouped_formula` now take it); worksheet cell-value snapshots (`limits.raw` or the record decoder's raw profile); named-range reading (`entries::read` and `parse_with` limits); data-validation formula rendering (the parse call's profile); cell-value formula text and transfer (`Workbook::raw_limits_value` of the base, source or target); name rewriting in `cell_values/names.rs`; sheet-removal dependency checks; the conditional-format editor (`Staged::raw_limits` for its `Encoder`, `BrtDXF` reads and encoding, replacing a hidden `raw::Limits::DEFAULT`); the workbook and worksheet writers (`Encoder::limits()`).

Two paths have no opened owner and use the documented default authoring profile explicitly: the formula text compiler's `PtgList` encoding and the `writer::Worksheet` builder's grouped-formula validation. The web-extension binding formula check also passes `raw::Limits::DEFAULT` for now, matching the record walk it sits in; the host-codec slice replaces both with the owner profile.

## Tests

- `formula::codec::tests::raw_limits_bound_formula_streams_strings_and_output`: exact and one-under payload limits on the token and ancillary streams; one-under string limit on `PtgStr`; exact and one-under output limits for `ParsedFormula::to_bytes` and `Group::to_record_data`; one-under input limits for `ParsedFormula::parse` and `Group::parse_shared`.
- `conditional_formatting::codec::tests::binary::model_tests::raw_limits_bound_rule_payload_text_and_output`: exact and one-under `BrtBeginCFRule` payloads; exact and one-under nullable-string units; cursor payload admission; `BrtCFRuleExt` admission; `DifferentialStyle` round trip at exact limits and refusal one under on payload and font-name units, on both parse and encode.
- Existing codec, reader, writer and editor tests move to the explicit profile unchanged in their expectations.

## Validation

Commands were run with `/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb-raw-policy/env.sh`:

- `cargo fmt --all -- --check`: exit 0.
- `cargo test -p litchi-xlsb --all-features --offline`: exit 0; 1,334 passed, 0 failed.
- `cargo clippy -p litchi-xlsb --all-targets --all-features --offline -- -D warnings`: exit 0.
- `cargo check -p litchi-ops -p litchi -p litchi-cli --all-targets --all-features --offline`: exit 0 (no consumer outside the crate called the changed signatures).

Logs under `/home/zhuhe/code/litchi-worktrees/.coord/w4b-xlsb-raw-policy/`: `xlsb-0360eg-tests.log` (121,491 bytes, SHA-256 `7cc2da4d5b343cf8b6c73b401be65ab1534d8f0e626a48639cac8c31bb85491b`) and `xlsb-0360eg-clippy.log` (172 bytes, SHA-256 `0f72bba120dead49c5d9cb8b5cb7b72485f3d8effaa96a4d4d8d9fa9943dacb2`).

## Performance

The formula parser gains two length comparisons per parse and one per decoded string; the conditional-format cursors gain one comparison per record and one per string. No allocation is added. Native timing is pending with the rest of W4b.

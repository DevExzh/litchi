# 0360eb — XLSB styles and cell-value raw policy

Status: source and consumer migration complete; native performance
qualification remains open.

This slice continues `b04fcd5bc7` and gives the standalone BIFF12 styles
codec one required `raw::Limits` profile. `styles::read` and every public
font, fill, color, border, number-format, and cell-format parser now validate
the profile before decoding. Fixed-width payloads are subject to the same
payload ceiling as variable records. The package-facing alignment and border
parsers use the same required profile.

Opened eager and source-backed workbooks pass their retained raw profile into
styles loading and differential-style loading. Cell-value format creation,
style interning, style transfer, number-format/font dependencies, conditional
formatting, and source-backed publication carry the source and recipient
profiles explicitly. The comments editor's worksheet-layout and legacy-
drawing framing fallbacks likewise use the owning workbook profile.

The tests cover strict fixed-payload and style-string boundaries, preserve
unknown style records through collection append, and keep the existing
source-backed and conditional-formatting behavior. No raw style consumer
silently installs a default profile after owner admission; detached tests
select `raw::Limits::DEFAULT` explicitly.

Validation on the final source:

* `cargo fmt --all -- --check` passed.
* `cargo test -p litchi-xlsb --tests --offline` passed: 737 library tests
  and all integration tests passed.
* `cargo clippy -p litchi-xlsb --all-targets --all-features --offline --
  -D warnings` passed.

The complete test log is retained outside the checkout at
`(deleted-coord-artifact)/xlsb-styles-tests-final.log` (113557 bytes,
SHA-256 `1e6af3c7a6630efc610e56a6cbcb96bd80cab8ae5ab987e4b0d0a41ea4e111be`).
The Clippy log is retained at
`(deleted-coord-artifact)/xlsb-styles-clippy-final.log` (171 bytes,
SHA-256 `b497299dfa7cc194970efa6b857069a639714fc0cf54ac525cb4e0c8f75012fc`).
No native timing or binary evidence is claimed by this record.

# 0360ea — XLSB comments and calculation raw policy

Status: source and consumer migration complete; native performance qualification
remains open.

The standalone BIFF12 `comments::read` and `calc::read` APIs now require an
explicit `raw::Limits` profile. Calculation payload admission checks the
caller payload ceiling before decoding, while comments keeps the same profile
through record framing, wide strings, and rich-string phonetic text. Their
writers continue to receive the required profile through the bounded raw
`Encoder`.

`OpenOptions::raw_limits` retains the selected profile through `Package`,
`Workbook`, and `SourceBackedWorkbook`. Workbook reparses, cell-value
publication, comments edits, source-backed comments reads, and the workbook
calculation parser all pass that retained profile. No raw codec path in this
slice falls back to the raw default after owner admission; detached test
constructors explicitly select `raw::Limits::DEFAULT`.

The owner policy test verifies a non-default profile survives Workbook →
Package → reconstructed Workbook. Direct codec tests cover exact payload
limits and comment record payload limits while preserving legacy 25-byte
`BrtCalcProp` input and exact comment bytes.

Validation on the final source:

* `cargo fmt --all -- --check`
* `cargo test -p litchi-xlsb --lib --offline` — 733 passed.
* `cargo test -p litchi-xlsb --tests --offline` — all library and integration
  tests passed.
* `cargo clippy -p litchi-xlsb --all-targets --all-features --offline --
  -D warnings` passed.

The complete test log is retained outside the checkout at
`(deleted-coord-artifact)/xlsb-tests-final.log` (113131 bytes, SHA-256
`e652ccebc54ce5cdd7ac84b865a9bf892fde881736efb5a030e9ea487fe4d6ce`). The
Clippy log is retained outside the checkout at
`(deleted-coord-artifact)/xlsb-clippy-final.log` (3925 bytes, SHA-256
`bef1c1138397be5e5b497a77bc61854e4fc6860fd4b2d36b3a31c5ec19b033ea`). No
native timing or binary evidence is claimed by this record.

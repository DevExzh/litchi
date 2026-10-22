# 0360ed — XLSB named-ranges raw policy

Status: source and consumer migration complete; native performance
qualification remains open.

This slice gives the standalone `BrtName` codec and every named-range consumer
one required `raw::Limits` profile. `named_ranges::parse` and
`Definition::parse` now require the caller's profile, and the tolerant
workbook parser takes the same profile while preserving its existing repair
report. Payload admission is checked before semantic decoding; the family
ceilings for a name payload and wide strings remain in force by intersecting
the retained caller profile with the codec's fixed maxima.

The profile is retained through both eager and source-backed workbook
`defined_name_entries` projections. Formula-name transfer, sheet-removal
dependency checks, and the defined-name editor use the owning workbook's
profile for record framing, headers, `BrtName` parsing, and authored splices.
The raw host `NameRecord` boundary and all sheet-order/insert/remove test
consumers use an explicit profile as well. The workbook writer passes its
explicit build profile into named-range payload encoding instead of creating a
hidden default encoder in that family. No reopened-owner path in this slice
installs `raw::Limits::DEFAULT` after admission.

The codec tests cover complete-name fidelity and malformed records, exact
payload and UTF-16 ceilings, and one-below refusals with typed raw errors.
Existing writer round-trip and sheet-structure tests continue to exercise
unknown-record preservation and exact defined-name scope/formula behavior.

Validation on the final source:

* `cargo fmt --all -- --check` passed.
* `cargo test -p litchi-xlsb --lib --offline` passed: 741 tests.
* `cargo test -p litchi-xlsb --tests --offline` passed: all library and
  integration targets.
* `cargo clippy -p litchi-xlsb --all-targets --all-features --offline --
  -D warnings` passed.
* No native timing or binary evidence is claimed by this record.

The complete integration log is retained outside the checkout at
`.coord/w4b-xlsb-raw-policy/xlsb-named-ranges-tests.log` (113908 bytes,
SHA-256
`0bee7aad32576a76f4b28d1fe2e3ea3c41dbc7c1e3b94dac050b0a25098cd098`). The
Clippy log is retained outside the checkout at
`.coord/w4b-xlsb-raw-policy/xlsb-named-ranges-clippy.log` (172 bytes,
SHA-256
`637a70c8c5a14fdc047b254c8a6dd945e8464a47629f0bde28ebe128e1b82de8`).

The remaining 0360e raw-policy audit rows are outside named ranges: formula
and complex XML/record codecs whose standalone parsers and writers still need
required policy threading, plus the workbook host record helpers that belong
to those families. The earlier family-level gaps in 0360b remain open as
separate owner work: free-text edit refusals, the deferred content-message
redaction audit, source-backed metadata patch replay, ooxml-common/XLDM typed
limit origins, and lexical-only raw OPC edits. No item is waived by this
slice.

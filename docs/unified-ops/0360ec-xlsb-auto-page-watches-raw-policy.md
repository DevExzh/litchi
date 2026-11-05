# 0360ec — XLSB auto-filter, page-break and cell-watch raw policy

Status: source and consumer migration complete; native performance
qualification remains open.

This slice continues `94e5f39501` and gives the standalone worksheet feature
readers and writers one required `raw::Limits` profile. The public
`auto_filter::read` and `page_breaks::read` entry points now require that
profile. Their worksheet-layout scans, record readers, fixed-width payload
parsers and encoders carry it through; auto-filter range payloads and page
break records refuse before emitting a payload that exceeds the caller's
ceiling.

Eager and source-backed workbook feature projections pass their retained raw
profile into auto-filter and page-break reads. Their editors use the staged
owner profile for layout scans, reads, column preservation and publication.
No consumer installs `raw::Limits::DEFAULT` after owner admission.

The cell-watch worksheet and package readers now require a raw profile, and
`cell_watches::workbook::load` carries it over every worksheet. Snapshots retain
the profile; edit, replay validation and rendering reuse it for watch and
phonetic fixed payloads, BIFF12 framing, and retained opaque records. The
workbook facade passes its retained package profile into the projection. The
tests cover an explicit fixed-watch payload boundary, profile retention through
an edit, and byte-identical preservation of an unknown record.

The earlier `0360ea` record's validation prose was corrected so its external
test and Clippy receipts are described separately.

Validation on the final source:

* `cargo fmt --all -- --check` passed.
* `cargo test -p litchi-xlsb --tests --offline` passed: 740 library tests and
  all integration tests passed.
* `cargo clippy -p litchi-xlsb --all-targets --all-features --offline --
  -D warnings` passed.

The complete test log is retained outside the checkout at
`(deleted-coord-artifact)/xlsb-auto-page-watches-tests-final.log` (113825
bytes, SHA-256
`a3c60188eef35babce55b84cb994b67e901433d59998328bd3077c8ad011dd7f`). The
Clippy log is retained outside the checkout at
`(deleted-coord-artifact)/xlsb-auto-page-watches-clippy-final.log` (171
bytes, SHA-256
`1db3448a5429562e6cc3a4b7322193876b82d615961ba9950e6b1aa36c96c570`).

Named ranges remain the next standalone worksheet/workbook raw-policy slice.
Formula, calculation-adjacent, and other complex XML/record families remain
outside this record and still require their own required-limit consumer
migrations. No native timing or binary evidence is claimed by this record.

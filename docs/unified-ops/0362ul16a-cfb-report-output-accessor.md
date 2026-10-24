# CFB report output accessor canonicalization

The CFB report surface had two duplicate `bytes()` shorthands. The
`SequentialWriteReport::bytes()` method duplicated its canonical
`output_bytes()` accessor, and the overlay `PublishReport::bytes()` method used
the same ambiguous name for the complete artifact length. Record 0351 and ADR
0034 require one explicit output accessor and prohibit compatibility aliases.

This follow-up removes both methods, keeps `output_bytes()` as the canonical
name, and migrates the seven exact overlay test consumers at lines 403, 544,
836, 863, 967, 1002 and 1515. The consumer compile then exposed three more
`PublishReport` uses in legacy test code: `litchi-ppt/src/text_edit.rs:6018`
and `litchi-xls/src/cell_values/tests.rs:1469,1505`; those are migrated in the
same patch. No other `bytes()` method is changed; the CFB overlay
operation-shape and unrelated payload APIs retain their own semantics.
The glossary registers exact scoped bans for
`litchi_cfb::writer::SequentialWriteReport` and `litchi_cfb::PublishReport`, so
the generic `bytes` vocabulary remains available outside these report types.

The preceding CFB source change is `f9ee8613dff75df1058ee2df228ac9f84afedc8f`.
This follow-up is source-only API cleanup plus consumer migration; it does not
alter serialization, limits, source fidelity, publication or allocation
paths. The [validation receipt](results/0362ul16a/validation-receipt.json)
binds the final source hashes, successful logs, the two initially rejected
consumer-check logs, and external target sizes. Focused CFB tests, Clippy,
rustdoc glossary checks and legacy consumer compilation pass for this exact
change. No ABBA or performance acceptance is claimed.

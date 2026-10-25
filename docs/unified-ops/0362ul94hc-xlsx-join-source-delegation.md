# 0362ul94hc — XLSX join retained-source delegation

This source-only follow-up is composed after 0362ul94hb at
`c0bf8cd9f88a81efe2874c972c864f298aae9bce`.  It corrects the join fallback
boundary without adding a public constructor, test-only export, compatibility
alias, or dependency from the XLSX crate to `litchi-ops`.

## Classification boundary

`litchi_xlsx::Error::Join` now delegates every retained
`normalization_source()` to the existing XLSX classifier. Therefore retained
refusals preserve their `RefusalKind` code, `PatchConflict` remains
`invalid_input`, typed `Limit` remains `limit_exceeded`, and nested shared
`SourceChanged` follows the same `io_error` rule as the direct XLSX leaf. Only
a join with no retained normalization source keeps the historical
`corrupt_file` fallback.

The focused ops tests cover every current refusal kind, patch conflict, typed
limit, nested-versus-direct source change, both operation-boundary branches,
and the no-source fallback. The existing real cross-snapshot join conversion
and the XLSX source-chain tests remain in place.

## Validation status

No Cargo, build, rustdoc, runtime, or performance command was run.  The
coordinator must rerun the composed OOXML consumer and ops gates.

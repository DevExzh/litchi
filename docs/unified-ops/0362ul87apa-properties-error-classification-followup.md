# 0362ul87apa — finish OOXML metadata error classification

The follow-up closes the remaining §2.4 classification gaps in the shared
OOXML metadata transaction. Source application-properties scans remain package
corruption, authored XML lexical failures remain invalid values, and bounded
resource failures remain typed limits. Prepared core, application, and custom
outputs now classify generated-size, overlap, read-back, decode, allocation,
and other prepared-output invariant failures as content-free `Internal`
errors. Patch replay decodes use the same generated-output fence.

A core part whose tolerant read recorded repairs is now refused with
`RefusalKind::UnsupportedContent` when a semantic core rewrite would discard
opaque content. Application-only edits retain the repaired core bytes. The
three XLSB source-backed stale-source expectations use the typed
`SourceChanged` error while cancellation expectations retain their nested OPC
`Cancelled` shape. Unused field bindings were removed, and the current
glossary/0253 metadata description now names `PatchConflict` and typed
`Refused(Refusal)` errors.

Focused tests cover repaired-core refusal and opaque-byte preservation,
authored application-property lexical rejection, and the source-backed error
shapes. Only source edits, rustfmt, and diff checks are intended for this
repair; no Cargo, rustdoc, glossary, native, or performance command is run.

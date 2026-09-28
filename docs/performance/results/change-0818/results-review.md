# 0818 terminal results review

This review covers the fresh terminal-0 quality, artifact, admission, and
ZIP-preservation evidence for the 0818 repair. It does not authorize timing or
interpret this preservation correction as an optimization.

## Disposition

**Pass: preservation repair admitted.** The fresh independent artifact audit
exited zero with `ok: true` and no errors. The report is
`admission-0/artifact-audit.json`, schema
`litchi.performance.0818.artifact-audit.v1`, 2,603,942 bytes, SHA-256
`84174f8c8201257888220d3a8e1ed6cb9a77f2d7eed6fd2f129f3c81c156c799`.
The admission receipt records the same terminal result and binds the fresh
manifest at SHA-256
`31a1186d465fca5706e5d3b32faa902d32c90b1d6c34b1782b9977c196ad66ba`.

## Quality and artifact cardinality

All six quality gates exited zero: formatting, all-target checking, all-feature
tests, warnings-denied Clippy, warnings-denied rustdoc, and crate-boundary
checking. The fresh test summary records 1,938 passed, 32 ignored, zero
failed, 66 suite summaries, and both new preservation regressions included.
The separately retained baseline regression exited 101 before the repair, as
required: the unchanged-production exact-relationship assertion failed. That
historical failure remains evidence of the original defect and is not merged
into the fresh passing result.

The exporter produced six cases with five policy outputs per case, for 30
outputs. Each case's five output hashes are identical, each output reopens
successfully, and the independent artifact report is error-free. The six cases
are generated DOCX/XLSX/PPTX and real DOCX/XLSX/PPTX. Their output identities
remain bound to the source hashes in `artifacts/manifest.json`.

## DOCX preservation evidence

The real DOCX source is 23,503 bytes with SHA-256
`1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5`.
Every policy output is 23,535 bytes with SHA-256
`de9e163ac26e170ee3881c7d7e53ac29836efd6caea6f72db7ebd2cd31d6c774`.
The ZIP-preservation replay reports equal archive comments and member order,
with only `word/document.xml` changed; all 11 other members retain their
decoded and compressed payloads and ZIP metadata.

For the repaired member, `word/_rels/document.xml.rels` is 817 decoded bytes
in both source and output, with source and output SHA-256
`f4e76362d55b0a76ff8303c06dc7e1fa4c3447df4992c53af39b77725b47ccb0`. Its
compressed payload is 249 bytes with SHA-256
`24145287ebde286a32409a47e5ff4410036a5802809b65bdbdbceb209fe5baaf` on both
sides, and its ZIP metadata comparison is true. Because all five policy ZIPs
are byte-identical, this exact relationship-member result applies across the
default, full, file-only, no-sync, and stream outputs.

This directly resolves the 0817 admission failure. The immutable 0817 receipt
still records exit 1 and 22 errors, including the three real-DOCX reports for
an outside-closure relationship graph change, changed decoded relationship
bytes, and changed canonical relationship XML. The 0818 result preserves that
failed receipt while admitting the repaired output under the same strict
relationship-order and untouched-member contract.

The generated and real XLSX cases retain their declared calculation closure,
and the PPTX cases retain their declared slide-only edit closure. The fresh
audit also continues to distinguish logical generated counters from physical
ZIP inventory; no old 0817 rejection was overwritten or treated as passing.

## Interpretation boundary

The evidence establishes correctness and preservation for this ordinary-save
repair, including all five publication policies and the six-case independent
artifact oracle. It provides no latency, allocation, throughput, or speedup
result. Any performance claim requires the separate timing batch described by
the packet's next-step plan.

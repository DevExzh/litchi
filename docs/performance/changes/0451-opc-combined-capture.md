# 0451: Combine OPC decoded reads with source-authorized compressed capture

Add `PartView::data_and_authorize_precompressed`. A cold cache loader captures
and decodes a source member once, then returns the ordinary managed `PartData`
and an opaque OPC transfer token. Warm hits and waiters reuse decoded bytes and
verify a fresh compressed capture. Existing leaf, source, signature, encryption,
limit and publication checks remain shared.

The token pins decoded memory/object reservations after package/data handles
are dropped, alongside compressed and writer staging. Deterministic tests cover
ordinary/combined loader coordination, provisional rollback, exact budget
boundaries, short reads, source changes and native-input compressed transfer.

Eight independently opened Store/Deflate cases preserve exact decoded/compressed
bytes and whole published outputs. At 262,181 decoded bytes, Store source bytes
fall from 524,576 to 262,365 (22→13 calls), and Deflate from 524,810 to 262,482
(28→15 calls). See [all cases](../results/change-0451/measurements.md).
These are deterministic I/O assertions, not latency or memory-peak samples.

Validation passes 466 all-feature OPC tests, 59 targeted PPTX tests, 361 default
harness tests and 20 feature-gated binary tests. Strict OPC lint, warning-denied
rustdoc, the non-iWork workspace, formatting and boundaries pass. The bounded OPC
fuzz target exercises cold and warm combined capture before eager parsing; its
1,000-run ASan/coverage smoke and strict fuzz lint both pass. Portable evidence
is recorded in the results bundle.
[Validation notes](../results/change-0451/validation-notes.md) retain the initial
compile/lint failures and distinguish native-input transfer from an Office
application roundtrip.

Replay the sealed bundle with
`python3 -B docs/performance/results/change-0451/verify.py --sealed --cleanup`.
PPTX plans have not adopted this API. Reusable publication needs explicit writer
reservations; cloning the current token would undercount concurrent staging.
Matched complete timings and the broader non-iWork goal remain open.

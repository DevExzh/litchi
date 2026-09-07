# Bounded existing ODP tail publication

This batch adds a specialized source-backed ODP tail append path. It scans and
validates the source and candidate through bounded XML windows, then measures
and emits the changed ZIP member through deterministic replay. Untouched member
records and payloads remain source-backed. It does not construct an ordinary
owned Commit or reversible Patch.

The common dependencies are verified callback member readers, an XML scanner
with before-growth limits, a source-bound content insertion plan, and a ZIP
replay writer with size/CRC/SHA checks. Fixed Deflate input windows make output
independent of caller write boundaries. Typed failures retain the accepted
output count and the primary transport/source cause.

## Measurements

The control and candidate each retain 360 samples: two repeats, normal and
allocator binaries, and 64/4,096/8,192 source slides. Each lane has three warmups
and 30 measured operations, pinned to CPU 2 with one worker. Source corpus and
append semantics match; retained-result contracts differ. The final candidate
includes the retained-fragment capacity/lifetime correction found in review.
See [`comparison-final.json`](comparison-final.json),
[`control/README.md`](control/README.md), and
[`candidate-final/README.md`](candidate-final/README.md) for exact scope and
commands. The initial 360-sample candidate remains historical evidence in
`candidate/`, with its original comparison and source binding.

| Source slides | Owned control peak above entry | Source-tail peak above entry | Owned allocated bytes | Source-tail allocated bytes |
|---:|---:|---:|---:|---:|
| 64 | 781,342 | 620,381 | 10,613,383 | 2,607,922 |
| 4,096 | 18,027,568 | 620,385 | 110,226,105 | 36,593,937 |
| 8,192 | 35,958,388 | 620,385 | 211,442,207 | 71,164,177 |

These allocator values repeat exactly in both captures. Peak above entry is an
operation-scoped allocator measurement; it excludes source/fixture ownership
already live at entry and is not process RSS. Allocation volume still grows
with document size. The source-tail path makes four content passes and does
not imply constant CPU work.

Final normal-binary medium/large p50 deltas span -4.221% to +3.262%; tiny deltas
are -33.352% and -33.185%. These are numeric comparisons of different APIs,
not an ordinary Commit/Patch speedup. Tiny process-lifetime peak RSS is +8.301%
(R1 normal), +10.184% (R1 allocator), and +6.517% (R2 normal). No elapsed
quantile or GNU-time maximum-RSS positive delta exceeds 5%. The complete
quantiles, RSS scopes, and deterministic bootstrap intervals remain in the
final comparison and [`performance-review.md`](performance-review.md).

## Verification and limits

The ZIP, ODF-common, and ODP full release suites passed 1,348 tests with three
ignored; the harness passed 383 with one ignored, and OPC passed 497 with one
ignored. Thirteen focused insertion tests passed, including the retained
capacity/lease regressions; the full common and ODP suites were rerun after
that correction. Warning-denied Clippy, rustdoc, non-iWork formatting,
workspace feature checks, and crate boundaries passed. ZIP and XML ASAN fuzz
targets each completed 1,000 runs with seed 457. Exact source epochs and earlier
failures remain in `checks/` and are described in
[`integration-notes.md`](integration-notes.md).

Three final synthetic source/output archive pairs and ten public LibreOffice
fixture pairs passed the independent ZIP/XML oracle from the final build. Their
synthetic archive/XML identities match the initial output exactly. The native
probe also reopened
through the library and checked exact title/body text. No Office application
was launched. Malformed/over-limit XML, ambiguous topology, normalized text,
source changes, and unsupported package framing/security cases fail closed.

Initial-source whole-process profiles remain in `profiling/`, including their
retained oracle failures and explicit validation amendments. They identify
hashing, XML validation, memory comparison/movement, and namespace processing
as follow-up candidates; they are not final-source API phase attribution.

The sealed bundle passes [precleanup verification](precleanup.json) and
[portable-copy verification](portable-verification.json), including
[replay after staging cleanup](postcleanup-verification.json). Three
[altered-copy checks](negative-verification-r2.json) reject changed summary
statistics, a modified fuzz seed, and an unlisted fuzz input even after the
outer seal is regenerated. [Cleanup](cleanup.json) removed 284 inventoried
files (375,674,717 bytes) from the owned staging directory; shared build caches
remain. The full non-iWork goal remains open,
including ordinary reversible-patch integration and the wider CRUD/input/output
and bounded-worker performance matrix.

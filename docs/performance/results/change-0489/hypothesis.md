# 0489: reuse a prepared candidate XML audit

## Measured opportunity

0487 reduced authored-heavy replay sink-fence overhead while leaving all XML
audits intact. Source-heavy normal medians remain about 385–436 ms depending
on input; its instruction counts changed less than 0.2%. The 0484 lifecycle
and 0486/0487 profiles identify repeated candidate parsing in ZIP measurement
and emission after OPC preparation has already validated the complete candidate.
Expected-artifact publication adds a private preview that repeats those passes.

This experiment removes repeated candidate XML parsing using a private,
plan-scoped capability minted only after successful source and candidate audits.
It binds immutable source/replay/candidate proof scalars, target identity and
frozen XML limits. Preview clones only that same authorization. A different
plan or durable patch application must establish its own audits.

The standalone source audit remains mandatory: generic insertion can repair
malformed source XML, so candidate validity alone cannot prove source validity.
The initial candidate audit and final DOCX semantic reopen also remain.

## Invariants and tradeoffs

All later passes must use the existing verified decoded reader and splice
adapter, rechecking source/replay/candidate lengths and digests, replay EOF,
source freshness, cancellation, Work, ZIP verification and accepted output.
They may drain bounded adapter windows instead of consuming XML token slices.
That changes the exact private prefix reached on a Work-limit failure; typed
failure, cumulative byte charging, bounded cancellation opportunities, no
output after authorization loss and exact sink accounting remain required.
A changed provider may now fail byte authentication instead of XML grammar.
No unsuccessful pass may produce a successful publication authorization.

Keep fixed-fragment and replay XML workspace admission reservations unchanged
for this experiment, even when a later pass needs no parser allocation. This
conservative choice isolates parser elimination from budget-policy changes.
The proof capability contains scalar state only, with no document-sized buffer,
no global cache, no public API change and no new dependency.

## Evidence plan and ADR mapping

Retain the full 0487 matrix: 18 arms, normal and allocator builds, two reversed
process repeats, 30 samples after three warmups, before and after (144 children,
4,320 samples). Before uses the retained 0487 executables. Collect whole-child
perf and strace diagnostics separately. Retain every adverse latency/heap/RSS
row above 5%; two repeats do not justify a strong confidence interval.

| ADR | Obligation |
| --- | --- |
| 0001 | Correctness and preservation before measured speed; reject an unhelpful or unsound candidate. |
| 0002, 0010, 0011, 0024 | Keep the implementation private to the OPC physical owner. |
| 0003 | Bind immutable proof/limits to one prepared plan; retain preview, publication, patch and inverse authentication. |
| 0005 | Bounded windows, existing admission checks, per-pass Work and exact accepted-output accounting; measured normal/allocator evidence. |
| 0006 | Preserve independent source and initial candidate audits, source-first failures, malicious-input limits, EOF/hash checks and final reopen. |
| 0008 | Focused OPC/DOCX tests, feature combinations, lint/docs/boundaries, helper tests and ASan fuzz campaigns. |

All accepted ADR hashes match the previously read set, recorded separately.
The full non-iWork goal remains open; this is a pre-change hypothesis.

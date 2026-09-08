# DOCX fresh streaming creation measurement

The full goal requires very large row/paragraph/slide streams and bounded
memory proportional to an explicit window. The prior 0472 independent audit
finds public StreamingDocumentWriter correctness coverage but no dedicated
performance selector. Existing DOCX semantic small creation constructs an
owned Package and Vec, so its evidence cannot characterize this writer.

This batch adds only opt-in harness coverage. Production writer, codec,
container, allocator observer and default scenario matrix remain unchanged.
Fresh paragraph creation is distinct from logical append to an existing
structure, package Part addition, and modification followed by repackaging.

The public writer emits two static OPC members and streams word/document.xml
as the third/final member. It reserves a fixed 64-byte semantic escaping
capability. Its documentation explicitly excludes ZIP/Deflate and caller sink
state from that accounting. The operation allocator observer measures these
logical requested-heap allocations independently; no measured or modeled
scratch figure substitutes for process RSS or total memory.

The proposed matrix grows paragraph count from 64 through 8,192 to 131,072,
with a zero-retaining hashing sink during measured operations and a separately
materialized preflight artifact. Full paragraph/run readback, exact package
membership, semantic text digest, archive hash and byte counters authenticate
output. Preflight setup/reopen may raise whole-process RSS and remains outside
the operation allocation region. Normal and allocator-target timings must not
be compared as an optimization result.

All 30 previously read ADR files retain their hashes (adr-refresh.json).
ADR 0001/0004 constrain semantic API use, 0002/0010/0011/0024 retain crate
ownership, 0003/0006 separate fresh creation from preservation/editing,
0005 requires explicit memory scopes and reproducible measurements, and
0008 requires correctness and negative-oracle checks. No production API,
unsafe boundary, execution provider or dependency changes are planned.

No new Office application run or fuzz campaign is part of this harness-only
measurement. Existing writer contract tests cover typed refusals, cancellation
and output progress; scoped tests are rerun. Native breadth, all-feature
streaming, other append meanings, source variants and parallel scaling remain
separate open requirements of the full non-iWork goal.

The implementation is committed as `a1bd623b3dd0ab23259d2d6f692392ee275475e0`.
Root review removed an unused work-budget local, releases the preflight archive
and document XML before timed samples, and distinguishes scratch reservation
from sink output retention. The getter before finish excludes the known
closing suffix; its value is checked against the complete preflight part.
Every paragraph must have exactly one run with the expected full text.
The negative oracle rejects an extra member, changed text and a split run.

The first five focused tests passed with an unused-local warning. After cleanup,
the full harness run found only the stale selector count (439 versus 440).
Its failure is retained, and the corrected complete suite passes 363 tests
with one ignored. The default count remains 37. Formatting, strict Clippy and
warning-denied rustdoc pass. Independent initial source audit was completed;
follow-up review agents stopped at a service usage limit, so root completed
the final source and evidence review locally. No independent final approval
is claimed.

The source inventory deliberately covers all 7,033 tracked Rust/TOML/lock
files, including 39 historical probe/source files under docs. The prior 0472
inventory excluded those historical files. Exactly two active harness files
change in this batch; the new module accounts for the remaining count increase.

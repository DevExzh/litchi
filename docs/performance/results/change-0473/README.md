# DOCX fresh streaming creation and operation memory

This bundle measures the public StreamingDocumentWriter through the new opt-in
`docx_streaming_create` harness selector. The existing buffered DOCX creation
cases remain separate. Production authoring and the 37-case default matrix
are unchanged.

The frozen protocol grows one-run paragraphs through 64, 8,192 and 131,072
units. It runs normal and allocator executables in fresh processes on CPU 2,
with three warmups and thirty samples, twice in reversed order. The 64-byte
XML-escaping scratch capability, zero retained hashing-sink output, operation
allocator peak and process RSS are different quantities. No normal-versus-
allocator timing comparison or production speedup claim is made.

Untimed corpus construction materializes an artifact and reopens all
paragraphs/runs. It checks exact package membership/order, text digests and
finalized document XML identity. Each timed operation checks its archive hash,
accepted bytes and writer counters against that preflight. The timer and
allocator region include context construction, text generation, writer setup,
streaming calls, Deflate/ZIP finalization and writer/context destruction.
Sink construction and digest extraction, observer reads and semantic preflight
are outside that interval. Process RSS still includes preflight materialization.

Rebuild from the authenticated revision in `build.json` at its recorded clean
build path, populate the two bound compile-time fixtures, and use Rust 1.98.1.
`build.py` builds both targets together with the allocator feature; only the
allocator executable installs the existing observer. `capture.py LANE` runs
one item from the protocol. All heavy commands are serialized using
`/tmp/litchi-goal-0473/cpu.lock`. Capture into a new bundle rather than replacing
historical reports. Retained build/capture scripts bind the exact source
inventory, fixtures, executable hashes, commands, environment and chronology.

The allocator observer records callback-order logical requested heap, including
other process threads between the endpoints. It excludes allocator-internal
reallocation overlap, physical copies and RSS. ZIP/Deflate state and the fixed
three-member directory are included in the observed operation allocations,
although they are outside the writer's semantic scratch budget. Any stable
peak result is limited to these tested shapes and this scalar text subset.

Fresh creation does not prove logical append to an existing document, adding
a package Part, or arbitrary modification followed by repackaging. Native
Office, source variants, full feature breadth and worker scaling remain open.
The full non-iWork goal is not complete.

## Results

| Paragraphs | DOCX bytes | Normal p50 ms R1 / R2 | Operation peak above entry |
| ---: | ---: | ---: | ---: |
| 64 | 1,233 | 0.119590 / 0.118971 | 414,732 bytes |
| 8,192 | 24,571 | 11.010543 / 10.973853 | 414,732 bytes |
| 131,072 | 376,848 | 176.711590 / 176.334184 | 414,732 bytes |

All 180 allocator samples have a 414,732-byte peak above entry, zero live-byte
change at exit and no failed allocation calls. All 360 formal samples pass.
Normal mean/median/tail repeat changes stay within five percent; no registered
latency or general RSS claim is made. See the [change record](../../changes/0473-docx-streaming-operation-memory.md)
for allocation work, raw RSS, verification scope and limitations.

Compilation succeeded before temporary binary publication hit the /tmp quota.
The recorded sparse-checkout recovery retained all authenticated source files;
exact compile fixtures were restored after a capture precheck refused their
absence. Formal captures began only after restoration. No protocol or raw
formal capture was replaced. Root completed final evidence review after the
follow-up review agents reached their service usage limit.

`analyze.py` derives results and uses `report_checks.py` for strict shared
report/schema/statistic arithmetic, adapted from 0432. `verify.py --live`
authenticates live binaries and the source tree; `--sealed` verifies exact
regular-file coverage and requires temporary binaries/tree absent when not
using `--live`. A portable replay needs only this complete bundle and Python's
standard library: run `python3 -B verify.py --sealed --portable-check` from a
fresh copy. It does not need prior bundles, Git or Rust after cleanup.

Final sealed live verification passed before cleanup. Fresh-copy portable
verification and output-mutation rejection pass after both binaries and the
temporary source checkout are removed. The portable copy is also removed;
shared Cargo caches and both user-owned files are preserved.

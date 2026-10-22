# 0362ye — ODP deferred append structural facts

The returned candidate moved the append-only namespace and page parser into
every source open.  That made open do more work than the historical ODP
UTF-8/body admission pass, even though the parser was bounded and its facts
were small.  This repair restores the original open boundary and computes
append facts only when an append is prepared.

The first successful append scan records only the proof's immutable facts:
decoded member length, the insertion offset after the direct page prefix, and
the direct page count.  They are held in a shared `OnceLock` containing no XML
or parser events, so shares and subsequent uncancelled appends can reuse the
bounded result.  The common insertion plan still streams `content.xml` and
proves the source hash, exact splice, and final candidate.

The common `SourceContentInsertionPlan` still streams `content.xml` after the
authored page is constructed.  It continues to prove source freshness,
source and target SHA-256 values, the complete bounded splice, and the final
candidate XML.  The cached facts therefore replace only the second
format-owned structural traversal; they do not replace the byte-level source
proof.

## Admission and policy parity

Open continues to run the historical bounded UTF-8/body probe.  It does not
run the append namespace parser, so malformed XML, unusual depth, unsupported
body shapes, entity/DOCTYPE inputs, and their existing error timing remain at
the previous admission boundary.  A source containing a DOCTYPE remains
accepted by open and is refused by the append scanner as before.

After a successful default append scan, subsequent tail-append and replay
operations reuse the cached facts only when the operation has no execution
context or cancellation token.  Policy-bearing operations use the existing
scanner so input/work accounting, cancellation checkpoints, and error timing
remain unchanged.  Caller-selected content and slide limits are checked
before reuse; the common insertion plan rechecks the source revision before
and during its proof.  Shares reuse the small facts cell while deferred
owners keep the same source association.

Regression coverage pins malformed XML open fallback with semantic-read
failure, DOCTYPE open acceptance with append refusal, and the fact that open
leaves the append scan deferred until a successful prepare.  The source
repair is commit `b8a41e01a5` on branch `uo2/odp-heap-repair`; the earlier
structural-facts commits `126a2bb945`, `f58aa3b76c`, and `3178136432` remain
returned evidence.

## Performance status

The prior 0362yd matched evidence remains returned: its default append route
performed one strict-open content pass, one append structural pass, and the
common insertion pass, with approximately 113.9 KiB of open-phase allocation
attributed to the extra Deflate decoder/state.  This repair removes the
extra parser from open while retaining the append scan and common proof.  It
must still be rebuilt and rerun against the accepted e4a baseline with the
same allocator and route-local Callgrind lifecycles.  The 0362yd receipts are
not used as qualification for this repair.

## Validation to complete

After the coordinator releases the owned target, run one-job release builds
for the tracked 0362ye probe, the affected ODP unit/integration tests, strict
Clippy, allocator routes, and route-local Callgrind.  Record the source and
binary SHA-256 values, semantic output lengths/hashes, per-route allocation
and instruction rows, and the exact commands in the small 0362ye receipts.
Native timing remains pending a quiet host window.

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
leaves the append scan deferred until a successful prepare. The source
repair is on branch `uo2/odp-heap-repair`; the earlier
structural-facts commits `126a2bb945`, `f58aa3b76c`, and `3178136432` remain
returned evidence.

## Performance status

The prior 0362yd matched evidence remains returned: its default append route
performed one strict-open content pass, one append structural pass, and the
common insertion pass, with approximately 113.9 KiB of open-phase allocation
attributed to the extra Deflate decoder/state.  This repair removes the
extra parser from open while retaining the append scan and common proof.  It
must still be rebuilt and rerun against the accepted e4a baseline with the
same route-local Callgrind lifecycles.  The 0362yd receipts are not used as
qualification for this repair.

The open admission probe was then made allocation-free for its fixed
boundary state: UTF-8 validation uses a six-byte stack bridge for a scalar
split across fills, and the body/DOCTYPE window is a fixed stack array
scanned in bounded subchunks.  This removes the old 16 KiB UTF-8 scratch and
4 KiB sliding-window allocations without changing the probe's marker,
DOCTYPE, UTF-8, or body-missing decisions.  On the matched allocator probe,
the current candidate's cold open fell from 116,001 peak / 200,373 allocated
bytes to 83,233 peak / 168,181 allocated bytes; the append commit and
publication rows were unchanged.  The baseline `tail_append` route is still
not an admission-equivalent performance control because it does not read
`content.xml` at open; this row is recorded for attribution, not as a native
qualification.

The current candidate probe binary is outside the repository at
`/home/zhuhe/code/litchi-worktrees/uo2-odp-heap-repair/target/release/examples/odp_0362ye_probe`,
size 3,599,992 bytes, SHA-256
`f4b3e47bdcceb9e9e536a3c94b181bde040360ffd66c78ead73feb74562737f9`.
The accepted e4a baseline binary is at
`(deleted-coord-artifact)`,
size 3,096,104 bytes, SHA-256
`ecb034b0f57588c7d131ba9981d7004f5181ee214a88d14519f73660f35e0669`.
Both were run with the tracked route drivers and one build job with
incremental compilation disabled. Route-local Callgrind completed on the
same fixed baseline binary; the reduced rows are in
`results/0362ye/allocator-fixed-probe.md`. Native timing remains pending a
quiet host; no native performance acceptance is claimed here.

## Validation to complete

The focused checks pass on the fixed-buffer source.  The source unit suite
has 5 tests, the append unit suite has 14 tests, and the combined
`doctype_leniency`, `sequential_text`, and `source_backed` integration run
has 32 tests. Strict Clippy passes with the tracked `route_after.rs` probe
include. The current release probe preserves the baseline output
lengths and SipHashes for open, semantic slide/text reads, unchanged write,
append publication, and replay. Route-local Callgrind is recorded above;
native timing remains pending a quiet host window.

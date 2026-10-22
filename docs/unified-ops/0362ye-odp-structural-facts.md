# 0362ye — ODP strict-open structural facts

This candidate removes the duplicate format-owned `content.xml` structural
scan from the default ODP tail-append path.  The source-backed owner already
admits `content.xml` at open.  That same bounded XML pass now records only the
append proof's immutable facts: decoded member length, the insertion offset
after the direct page prefix, and the direct page count.  The source owner
does not retain the XML or parser events.

The common `SourceContentInsertionPlan` still streams `content.xml` after the
authored page is constructed.  It continues to prove source freshness,
source and target SHA-256 values, the complete bounded splice, and the final
candidate XML.  The cached facts therefore replace only the second
format-owned structural traversal; they do not replace the byte-level source
proof.

## Admission and policy parity

The structural pass is opportunistic.  If its XML parser reports a parser,
depth, token, event, resource, or allocation error, the owner retries the
historical bounded UTF-8/body probe.  I/O, committed-I/O, and source-change
errors remain primary.  This preserves the previous open admission for
malformed XML, unusual depth, unsupported body shape, and entity/DOCTYPE
inputs.  A source containing a DOCTYPE is never given cached append facts,
because the existing append scanner deliberately refuses that source.

Default tail-append and replay reuse cached facts only when the operation has
no execution context or cancellation token.  Those policies use the existing
scanner so input/work accounting, cancellation checkpoints, and error timing
remain unchanged.  Caller-selected content and slide limits are checked
before reuse; the common insertion plan rechecks the source revision before
and during its proof.  Shares copy the small facts value, while deferred
owners keep the same source association.

The source/test implementation starts at commit `126a2bb945` on branch
`uo2/odp-heap-repair`; policy and limit-precedence follow-ups are
`f58aa3b76c` and `3178136432`.  Regression coverage includes direct
bounded-fact collection, malformed XML open fallback with semantic-read
failure, and DOCTYPE open acceptance with append refusal.

## Performance status

The prior 0362yd matched evidence remains returned: its default append route
performed one strict-open content pass, one append structural pass, and the
common insertion pass, with approximately 113.9 KiB of open-phase allocation
attributed to the extra Deflate decoder/state.  This candidate removes the
append structural pass while retaining the strict open pass and common proof.
The new candidate must still be rebuilt and rerun against the accepted e4a
baseline with the same allocator and route-local Callgrind lifecycles.  The
0362yd receipts are not used as qualification for this repair.

## Validation to complete

After the coordinator releases the owned target, run one-job release builds
for the tracked 0362ye probe, the affected ODP unit/integration tests, strict
Clippy, allocator routes, and route-local Callgrind.  Record the source and
binary SHA-256 values, semantic output lengths/hashes, per-route allocation
and instruction rows, and the exact commands in the small 0362ye receipts.
Native timing remains pending a quiet host window.

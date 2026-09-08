# Next work after the 0475 attribution

Implement a private reusable raw Deflate compressor state and fixed output
scratch in the shared owned ZIP-entry writer, then run a frozen matched
before/after experiment. The repeated stack evidence assigns 6,240,505,472
requested bytes to backend initialization and 538,083,328 bytes to encoder
output buffers for the 16,421-member workload. The combined 99.543891% share
is allocation work within the run context; initialization is only about 5%
of writer CPU sample weight, so a proportional timing win is not established.

Follow `transport-audit.md`: preserve one independent raw Deflate stream per
member, StreamEnd and descriptor ordering, accepted-input CRC/counts, all
byte/name/directory limits, short-write and WriteZero behavior, poisoning,
cancellation, terminal-error handling and deterministic archive bytes. Reuse
state only after successful entry finalization and discard it after failure.
Use existing ZIP/OPC/PPTX differential, framing and failure-path tests, plus
tests that specifically exercise reuse across successful and failed entries.

Freeze candidate/control source and executable hashes before matched captures.
Use allocator samples across all three slide counts, bracketing/reversed normal
repeats and relevant shared-transport guards. Keep requested work, operation
heap, process RSS and normal latency separate. Test the actual claimed behavior;
do not infer a speedup from allocation savings.

The 0474 operation peak still grows with members. Compressor reuse cannot
remove persistent OPC folded/original-name and ancestor maps or ZIP directory
and name storage. A bounded-memory API needs an explicit metadata storage and
duplicate-name strategy, finite budgets and separate evidence. Do not rename
the active-slide XML limit or zero retained sink output a total-memory bound.

Untimed materialized preflight repeatedly reconstructs archive readers through
`OwnedPhysPkgReader::read_member`; it dominates whole-process profiles. A later
harness cleanup may reuse a validated reader while retaining exact membership,
every slide's text/geometry/layout ownership and the package digest gates.
That setup change must remain separate from production timing attribution.

Fresh creation, logical append, adding package Parts, arbitrary modification
and repackaging, native-producer breadth, source variants and bounded-worker
scaling remain distinct requirements. The full non-iWork goal is open.

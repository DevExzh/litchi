# Allocation and ownership review

The control is commit `7c7561160`, following retained candidate reuse in change
0418. The media-rich lifecycle's operation allocator counter reported about
36.8 billion requested bytes. This is cumulative requested allocation volume:
`record_reallocation(old_size, new_size)` counts the full new request, even if
the system allocator can extend in place. It is not physical copying or RSS.

`opened/cross_copy_plan.rs::BoundedVecWriter::write` reserves exactly the next
incoming chunk before every append. `build_candidate` serializes through this
writer, then calls `OpcPackage::from_vec`. The lifecycle builds once during
planning and once during fresh application-time verification. The baseline
heap trace tests the allocation attribution independently of this source
hypothesis; it includes the entire command, not just those two timed builds.

`OpcPackage::authorize_owned_source` retains `Arc<Vec<u8>>`, so spare capacity
in a generated archive persists. Geometric growth alone is therefore not an
acceptable memory-neutral assumption. A final fallible exact reservation and
copy can remove deliberately requested spare capacity, at the cost of one
copy and coexistence of the two buffers. Allocator rounding remains possible.
The archive byte limit is not an aggregate process-memory budget. Both the
old realloc path and such a compaction can have old/new buffers coexist;
measure the process high-water consequences explicitly.

The decoded-payload ownership issue is separate. Reopening constructs newly
decoded part buffers alongside the generated archive; it loses the staged
graph's shared payload allocations. A future storage-only adoption operation
would have to live in OPC, preserve exact-source authority without accepting
arbitrary source/graph pairs, and update all provenance holders. This batch
does not implement that broader ownership change.

ADR review: 0001/0003 require unchanged typed refusal and atomic publication;
0005 requires finite checked reservations and scoped reproducible evidence;
0006 requires exact bytes and failure before publication; 0010/0011/0024 keep
archive/OPC ownership below the format facade. A private writer-capacity change
does not alter public APIs, graph validation, source authorization, output
limits, custom `Part` behavior, save options, or dependency direction. Any
retained candidate must pass the corresponding existing cross-copy tests.

Independent source review: `profile_plan` checked allocator accounting and the
double candidate build; `zip64_review` checked retained ownership, fallible
compaction and near-limit arithmetic. Neither source review alone proves a
performance improvement.

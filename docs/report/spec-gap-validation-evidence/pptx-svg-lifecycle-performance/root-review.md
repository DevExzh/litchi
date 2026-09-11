# Root verification and measurement scope

The independent root verifier rehashed all 4,757 retained inputs, recomputed
report statistics, and checked 72 process receipts containing 1,440 samples.
All expected success/refusal outcomes and allocator equations passed. The
isolated build target and Python caches were removed.

The `clone_*` lanes include opening the editor, capturing the owner, and two
snapshot clones. They do not isolate clone latency; interpret them as complete
capture-and-clone workflows. Capture lanes likewise include editor open.
End-to-end attach/detach includes fallible closure validation at commit,
publication, and reopen checks. RSS covers the whole process, including fixture
setup; requested allocated bytes are cumulative allocations, not peak residency.

Successful samples return to their starting live-byte count. Refusal samples
retain 37, 39, or 54 diagnostic bytes at the observation boundary because the
returned error is still alive. The harness matches and drops that error after
taking the allocator snapshot. The verification does not mislabel these samples
as zero-net-live results.

Namespace-heavy input remains a measurable cost: the refusal lane requests about
81.6 MB cumulatively while its median incremental live peak is about 3.53 MB.
The 1,024-picture inventory cases request roughly 27–31 MB cumulatively and take
about 20–22 ms at the median on this machine. These measurements characterize
bounded work; they do not prove fully linear scaling, a before/after improvement,
native Office acceptance, or a general library-wide performance target.

The original profile sources, manifests, receipts, and generated report were
left unchanged. This note clarifies their interpretation.

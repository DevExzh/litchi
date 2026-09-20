# Independent source diagnosis

A read-only reviewer localized the rejected candidate's added work to index
publication and indexed replay. Publication linearly searches retained slots,
allocates another workbook-path vector, restores the worksheet-start hint,
walks allocation metadata to the target frame and retains a weak checkpoint.
The 0723 trace confirms one construction walk per owner/index publication;
prior 0723 stored-target allocator observations usually add 56 requested bytes, one
allocation/deallocation pair and 40 retained bytes.

Missing targets never call the target-checkpoint probe and have no replay cursor call.
Their regressions therefore cannot be explained by extra FAT traversal. The
candidate changes index layout/accounting and adds a replay selection branch.
The split between those source contrasts and compiler/timing effects remains
unproven before this experiment. Early stored targets avoid only 1–14 links,
while paying the same checkpoint-selection structure. Existing replay path-vector
allocation remains; the rejected change adds no new replay allocation call.

Potential next seams, subject to measured attribution and fresh qualification:

1. For an empty indexed slot range, skip unused replay setup while preserving
   the worksheet lookup, final cancellation and source-current checks, and existing
   error ordering.
2. Reuse the scan's existing borrowed workbook-path vector when constructing
   the checkpoint, removing the extra fallible 16-byte allocation.
3. Capture the first matching frame offset in transient TargetCell state to
   avoid searching every retained slot after scanning; retain duplicate order.
4. Consider an exact cursor checkpoint only if it can precede read-ahead without
   widening the public API or weakening index/source identity and fences.

A threshold alone does not remove the new 40-byte index field when it is None.
The existing physical layout must retain its fixed charge. Any side storage or
conditional accounting design requires separate lifecycle and budget evidence.
No proposed optimization is installed by this diagnostic packet. These are
source-level mechanisms and hypotheses, not isolated machine-instruction costs.

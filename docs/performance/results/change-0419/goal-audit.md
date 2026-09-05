# Change 0419 goal audit: remaining non-iWork work

`performance_claim: none`

`claim_authorized: false`

`disposition: OPEN`

This audit uses the [`docs/GOAL.md`](../../../GOAL.md) definition of done,
the [`CRUD_COVERAGE.md`](../../CRUD_COVERAGE.md) index contract, the
[0417 timing boundaries](../change-0417/timing-boundaries.md), and the
verified [0418 result table](../change-0418/result-table.md) and
[memory review](../change-0418/memory-review.md). iWork remains outside this
audit.

## What 0418 establishes

The 0418 owned PPTX cross-copy candidate-reuse path has a reproducible warm
ABBA result on generated media-rich and plain corpora. The primary
`pptx_cross_copy_media_rich_lifecycle` p50 falls from 1,175.449/1,172.463 ms
in the two control legs to 720.698/723.741 ms in the candidate legs, a
38.69%/38.27% paired reduction. The result table accepts p50, mean, p95, and
p99 for the primary and plain lifecycle guards. The capture retains 6,400
normal observations, 480 separate allocator observations, and 32 preflights;
the full [0418 protocol](../change-0418/protocol.json) records CPU 2, one
worker, warm generated sources, and the timer boundaries.

The resource evidence changes the next action. Primary media allocation calls
fall 7.319%, while cumulative requested allocation bytes fall only 0.135%.
Whole-process RSS rises in both normal pairs by 8.00% and 8.06%, and in both
allocator pairs by 8.01% and 8.15%, exceeding the declared 5% review
threshold. The memory review attributes the direction to retaining the
serialized archive and parsed graph; its live and high-water values are
process snapshots, not operation-local peaks. The whole-command PMU capture
is descriptive, and `cache-references:u` is an unvalidated zero alias, so it
does not supply phase-local cache or IPC evidence. The borrowed real-producer
fixture proves typed provenance refusal; it does not establish native Office
acceptance. These results support a narrowly scoped latency decision only.

The current non-iWork index remains representative: 15 categories and 30
representative selectors, with dynamic calculation unsupported. The default
matrix remains 36 cases and 198 rows; the 0418 lifecycle selectors are
opt-in. The 0417 rows include prepared-open and commit-only boundaries, and 28
of the 30 representative selectors still lack operation allocation
attribution. Existing source-backed correctness and locality evidence is
useful, but it is not an end-to-end resource baseline for the full matrix.

## Recommended order

1. **Resolve the 0418 memory tradeoff before broadening the optimization.**
   Attribute the candidate's per-chunk requests and retained bytes at
   `crates/litchi-pptx/src/opened/cross_copy_plan.rs` (including
   `BoundedVecWriter`) with matched control/candidate heaptrack or allocator
   evidence. Add an operation-local retained/peak observation and near-limit
   or low-memory owned media cases. Determine whether the temporary serialized
   archive or decoded duplicate can be released after the existing graph,
   fingerprint, patch, limit, and readback proofs. Preserve the existing
   fallback for dirty destinations, caller-defined `Part` implementations,
   non-default `SaveOptions`, and revoked authorization. Until this is done,
   retain only the scoped latency result and make no memory-improvement claim.

2. **Capture an end-to-end source-backed XLSX/OPC read and targeted-update
   baseline.** Use the existing selector families
   `xlsx_source_open`, `xlsx_source_list_sheets`, `xlsx_source_first_cell`, and
   `xlsx_source_narrow_column_range_scan`, paired with their eager controls
   where the corpus and oracle match. For writes, use the existing
   `xlsx_source_backed_cell_values_one_edit_save`,
   `xlsx_source_backed_cell_values_one_percent_edit_save`,
   `xlsx_source_backed_cell_values_batch_edit_save`, and
   `xlsx_source_backed_cell_values_multi_sheet_edit_save` families with the
   eager counterparts. The measured boundary should include the selected
   source ingress, semantic operation, commit, and sequential publication;
   retain independent reopen, exact-output, untouched-member, source-version,
   and refusal checks. Record operation allocations, source range requests,
   decompressed/recompressed and copied bytes, and RSS with fixed corpus and
   binary identities. Current prepared-open and commit-only rows, and the
   older source-backed correctness records, do not answer this question.
   Keep these selectors opt-in until the independent baseline satisfies the
   index's status contract; do not change the default matrix from this audit.

3. **Establish physical-access and workload generalization on the measured
   paths.** The existing `xlsx_range_source_*` and CFB positional selectors
   provide a starting point, but simulated range providers and tmpfs captures
   do not prove physical cold-cache or device behavior. Use a controlled
   block-backed cold run and a caller-supplied high-latency range source to
   retain request count/size distributions, bytes read, and selected-versus-
   mandatory work. Repeat the strongest media-rich and selective-edit cases on
   permitted native Office/LibreOffice fixtures; 0418's borrowed fixture is
   only a refusal oracle. Then measure explicit 1/2/4/8-worker execution with
   memory, I/O, cancellation, and task-size budgets, plus managed-cache
   contention. The current one-worker warm PMU run cannot establish any of
   these generalizations or scaling claims.

4. **Return to the remaining format slices after the above evidence is
   stable.** Prioritize measured ROI in CFB/DOC/XLS/PPT mixed FAT/MiniFAT and
   semantic publication cases, then richer ODF/RTF and OOXML dependency
   closures. Existing correctness and narrow phase evidence do not establish
   broad allocation, physical-I/O, cold, producer, failure-matrix, or
   end-to-end CRUD results. ZIP64 and preservation capability work also needs
   broader semantic, remote/cold, and failure evidence before it can count as
   a performance-program result.

No item above changes production behavior or promotes a selector. A future
**claimable** result needs the exact source revision, corpus and output
hashes, timer definition, uncertainty, preservation gates, and resource
review retained alongside its paired measurements. The non-iWork goal remains
open until the definition-of-done requirements and these evidence gaps are
closed.

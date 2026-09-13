# 0556 XLSX provenance merge candidate review

`status: independent read-only candidate review`

`performance_claim: none`

`production_change_by_reviewer: none`

This review audits the candidate copy at `candidate/cell.rs` against the
baseline source and the frozen patch. Root applied the candidate in the shared
working tree after the candidate freeze; the reviewer did not apply or edit
production. This review does not build or capture performance data. The
audited base revision is
`3210f07ac7a9daa2686e781645f6539a156a26ce`.

## Identity and patch binding

The source proof is recorded in
[`source-proof.md`](source-proof.md), SHA-256
`ccf77475d133b2833b823e93c4e162a122a3e9fcb208c6d1ca17660817d39dbe`.
The pre-freeze draft is preserved at
[`review-attempts/draft-01.md`](review-attempts/draft-01.md), SHA-256
`a83651a1f9143e2f696ed97ccea74472704750b261f4c5055186ca68b9fc63be`.

| artifact | SHA-256 at review time |
| --- | --- |
| baseline production `crates/litchi-xlsx/src/cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| applied production `crates/litchi-xlsx/src/cell.rs` | `fc3ede23ee3ff9b7c68b0f11d73bc95ff168105a44a8ad409f7116f92c4b196e` |
| `baseline-cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| `candidate/cell.rs` | `fc3ede23ee3ff9b7c68b0f11d73bc95ff168105a44a8ad409f7116f92c4b196e` |
| `candidate/candidate.patch` | `2b91feb2003f646741956f3510eddf046e34dd6801084cd4593d7fd2993872b8` |

The candidate source and regenerated patch now bind to the same
`fc3ede23...` source bytes, including the explicit `drop(parsed_cells)`
lifetime fix and the final test additions. The exact patch passed
`git apply --check` against the baseline source, and the applied production
source has the same SHA as the candidate source.

## Correctness review

`from_unsorted` still sorts and rejects duplicate cells and rows, then calls
the private `from_sorted` assembler. The only production parser constructor
is that path, and the only production merge caller consumes a parsed `Store`
created by it. The derived `Store::default` supplies empty cell and row
slices, so it is also trivially sorted; the malformed `Store` literals in the
candidate tests are test-only invariant bypasses.

The candidate therefore has the required precondition: parsed cells and
source cells are strict row-major, duplicate-free sequences, and
`omitted_entries` selects a sorted source subsequence after the checked
one-row range contract. Its two-way merge emits the lower address, refuses an
equal address, and checks `previous_address >= entry.address` before every
append. That final guard safely refuses test-only malformed cell sequences
instead of constructing an invalid binary-search store. Parsed rows are
already sorted and unique and are not changed by this merge, so moving them
to `from_sorted` is valid.

The candidate rebuilds `cell_rows`, all three derived cell bounds, and the
merge index from the final cell sequence. It keeps the parsed declared
dimension, rows, columns, defaults, and merge ranges authoritative. It copies
the parsed merge ranges and rebuilds the index; no merged-index move is part
of this candidate. Every `Stored` field is moved for parsed records and
cloned for selected source records, including style, formula provenance,
shared-string identity, rich-text marker, and metadata.

## Error, fallback, and resource review

The candidate retains the existing refusal order for empty or unordered
omissions, zero selected source records, and parsed cells inside an omitted
rectangle. It retains checked count addition, fallible reservations and their
labels, and the complete `from_sorted` error paths. The snapshot caller still
turns speculative `Err` or `None` into complete-output parsing; complete XML
validation, reduced-readback validation, execution fences, and publication
readback are outside the candidate and remain required.

The source now explicitly drops the exhausted parsed `Box<[Stored]>::IntoIter`
before allocating and rebuilding merge indexes (`candidate/cell.rs:879-881`).
Without that drop, the iterator could retain the reduced parser's backing
allocation while the merged cell vector and merge index were live, increasing
the operation peak. The drop is present in both the frozen patch and applied
source. During the interleaved walk the iterator necessarily remains live
alongside the merged vector until the merge completes; the post-loop drop only
removes overlap with later merge-range/index work. Peak memory and allocation
counts still require fresh operation-local evidence; source inspection makes
no such claim.

No semantic correctness blocker remains under the private constructor graph.
`from_sorted` intentionally trusts its row-order precondition, so any future
constructor or mutation that populates `Store` directly would need either to
retain that proof or restore validation. The current source search found no
such production path. The candidate's malformed-source and malformed-parsed
tests exercise the cell-stream refusal guard; production tests must still
cover the complete snapshot fallback boundary.

## Test coverage required before adoption

The current candidate source adds differential comparisons against a local
copy of the original append-then-`from_unsorted` `Store` merge path. Those
direct tests cover all `Stored` fields, rows, columns, defaults, merge
authority and lookup, cell-row traversal, extents, empty parsed stores,
adjacent and grid-boundary omissions, malformed source and parsed order,
collisions, and no-source refusal. They are not complete XML parser-oracle
tests. Existing `snapshot.rs` tests provide that oracle by comparing the
candidate result with a complete parse of the rewritten worksheet and cover
caller-level scalar values, implicit addresses, stale dimensions, unsupported
provenance, malformed output, source mismatch, exact no-op sharing, and
commit/readback behavior.

The candidate source tests were not executed in this read-only audit. Root
should finish the candidate unit and snapshot tests, workspace and feature
quality checks, and the bounded allocation/peak measurement protocol against
the exact frozen patch. The old and candidate paths must be compared on the
same fixtures, with fallback cases retained.

## Verdict

Conditional source review pass: the linear merge is correct under the proven
`Store` construction invariant and preserves the complete-parser fallback
boundary. The source and patch are now bound and applied; candidate tests and
fresh correctness, allocation, and performance evidence remain required
before adoption.

# 0556 XLSX provenance merge candidate review

`status: independent read-only candidate review`

`performance_claim: none`

`production_change: none`

This review audits the candidate copy at `candidate/cell.rs` against the
baseline source and the proposed patch. It does not apply the candidate,
build it, or capture performance data. The audited base revision is
`3210f07ac7a9daa2686e781645f6539a156a26ce`.

## Identity and patch binding

The source proof is recorded in
[`source-proof.md`](../source-proof.md), SHA-256
`ccf77475d133b2833b823e93c4e162a122a3e9fcb208c6d1ca17660817d39dbe`.

| artifact | SHA-256 at review time |
| --- | --- |
| production `crates/litchi-xlsx/src/cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| `baseline-cell.rs` | `7e807528568ebdd4a717382a3b1b249e178504b03e04d38147cb0159c5b567c7` |
| `candidate/cell.rs` | `fc3ede23ee3ff9b7c68b0f11d73bc95ff168105a44a8ad409f7116f92c4b196e` |
| `candidate/candidate.patch` | `1184f7a51e65f138cc42ffb3918c3ca2488401b2407a4276eadccb41a16c4fa3` |

The candidate source and patch are currently out of sync. The candidate
source contains the explicit `drop(parsed_cells)` lifetime fix and the newer
test additions, while the patch hash above predates those bytes. The patch
does pass `git apply --check` against the unchanged production source, but it
is not the reviewed candidate until it is regenerated from the current
candidate source and rehashed. This is a blocking packaging issue for
application, separate from the merge proof.

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
the operation peak. The explicit drop must be present in the regenerated patch
and the applied source. Peak memory and allocation counts still require fresh
operation-local evidence; source inspection makes no such claim.

No semantic correctness blocker remains under the private constructor graph.
`from_sorted` intentionally trusts its row-order precondition, so any future
constructor or mutation that populates `Store` directly would need either to
retain that proof or restore validation. The current source search found no
such production path. The candidate's malformed-source and malformed-parsed
tests exercise the cell-stream refusal guard; production tests must still
cover the complete snapshot fallback boundary.

## Test coverage required before adoption

The current candidate source adds complete-parser-oracle comparisons for all
`Stored` fields, rows, columns, defaults, merge authority and lookup,
cell-row traversal, extents, empty parsed stores, adjacent and grid-boundary
omissions, malformed source and parsed order, collisions, and no-source
refusal. Existing snapshot tests cover caller-level scalar values, implicit
addresses, stale dimensions, unsupported provenance, malformed output, source
mismatch, exact no-op sharing, and commit/readback behavior.

The candidate source tests were not executed in this read-only audit. Root
should run the candidate unit and snapshot tests, workspace and feature
quality checks, and the bounded allocation/peak measurement protocol only
after regenerating and hashing `candidate.patch`. The old and candidate paths
must be compared on the same fixtures, with fallback cases retained.

## Verdict

Conditional source review pass: the linear merge is correct under the proven
`Store` construction invariant and preserves the complete-parser fallback
boundary. Regenerate the stale patch first; then require candidate tests and
fresh correctness, allocation, and performance evidence before adoption.

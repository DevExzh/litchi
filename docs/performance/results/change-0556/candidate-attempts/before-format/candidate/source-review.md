# 0556 candidate source review: checked linear provenance merge

`status: source preparation only`

`performance_claim: none`

`adoption_allowed: not established`

This candidate is based on revision
`3210f07ac7a9daa2686e781645f6539a156a26ce`.  The accepted ADR continuity
manifest was rechecked before source preparation; all 30 recorded paths still
match `docs/performance/results/change-0555/adr-manifest.json`.

The only production candidate file is the full copy of
`crates/litchi-xlsx/src/cell.rs` in this directory.  `Store::from_unsorted`
keeps its existing cell sort, cell duplicate check, row sort, and row
duplicate check, then delegates index and extent assembly to the private
`from_sorted` helper.  `Store::merge_omitted_cells` merges the reduced parsed
cell sequence with the ordered source omission iterator by address.  Parsed
entries are moved; selected source entries are cloned exactly as before.  An
equal or out-of-order address returns `Ok(None)`, preserving the caller's
complete-parser fallback.  The existing fallible combined-cell and merge-range
reservations, merge-index construction, rows, columns, defaults, declared
extent, and extent/index assembly remain in the same ownership boundary.

The address guard checks the emitted sequence as it is built.  It therefore
catches duplicate or unsorted input from either sequence without adding a
second full validation pass.  Valid `Store` values already establish strict
cell order and unique rows through `from_unsorted`; the guard remains a
checked refusal at this private boundary if a future caller or an in-module
test violates that invariant.  The exhausted parsed `IntoIter` is explicitly
dropped before merge-range allocation and index rebuilding so its original
buffer is not retained through that work.  During the interleaved walk it is
necessarily live until the merge completes; this is a measurement limitation
for the fresh allocation and profile campaign.

The candidate tests in `cell.rs` include:

* a differential comparison with a local copy of the original append-then-
  `from_unsorted` constructor path;
* every `Stored` provenance field, all `Cell` variants, rows, columns,
  defaults, declared extent, merge ranges, cell-row index, and all computed
  extents;
* different source and parsed structures, including parsed merge coverage and
  sparse range traversal;
* an empty parsed store, adjacent omission spans, and first/last worksheet
  coordinates; and
* empty, reversed, colliding, and malformed-order refusal cases, including an
  out-of-order source sequence.

No build, test, benchmark, profile, or adoption decision was made while
preparing this artifact.  The parent campaign must apply the exact patch,
run the targeted quality checks, and perform fresh source-bound performance,
allocation, correctness, and preservation measurements before making any
claim.

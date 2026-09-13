# 0546 integrated bulk-count candidate

The final formatted patch is candidate.patch (identical to applied-candidate.patch).
Its source hashes are in candidate-source-hashes.json and the candidate freeze.
The original prepared patch, files, hashes and detailed provenance remain in
candidate-sources/. The independent source-review.md reviews that prepared
patch; format-adaptation.json records the three formatting-only changes made by
cargo fmt before any candidate build. Root inspected the exact line-wrap diff.

The algorithm and source/resource/error/publication behavior remain exactly the
reviewed design: sixteen sparse marker hits, then checked 64 KiB bulk counts for
< and &, plus the original full text-successor scan. The 0544 shared traversal,
validation-first fallback and post-EOF retry behavior are unchanged. No public
API, dependency, unsafe, event-limit or source-limit change is introduced.
Explicit reader settings, exact text cap and chunk pairs strengthen the oracle.

The passing isolated screen is not runtime admission. The original workflow,
allocation, refusal, cap and conditional gates, plus two new valid sparse
comment guards under the 5% envelope, all remain required. OLE2/OOXML remains
first and ODF waits until that optimization goal completes.

# 0832 — defer dense column-map storage until the second record

Baseline: `7eeaab48c0281b06f53527d4c4f4ea79050d27e7`.

The 0830 exact-owner allocation profile attributes two 1 MiB allocations per
real-file edit to `Assignments<Properties>`, once in the source worksheet parse
and once in the required post-write parse. The 0831 change removes the separate
unused validator map. The remaining parser maps are built for a worksheet with
one physical `<col>` record. Neither parse can simply be removed: both serve
the existing validation and preservation contract.

The candidate stores zero or one complete assigned record inline. A second
record triggers a fallible allocation of the original fixed-grid tree, followed
by replay of the first record and assignment of the second. Subsequent records
use the existing tree algorithm. There is no growing record list and no scan
proportional to each record's covered width. Promotion initializes a fixed
`2 * COLUMNS` vector once, costing `O(COLUMNS)`; after promotion each assignment
is `O(log COLUMNS)`. A first record and its lookups are constant time. The tree's
ordered traversal and adjacent-equal range compaction remain authoritative.
The inline `into_ranges` path reserves its single final range directly, so it
also avoids the dense traversal's intermediate raw-range vector. Its allocation
remains fallible with the existing `column ranges` resource label.

Last matching record means complete replacement, including properties omitted
by the later record. It does not mean property inheritance. Parser, validation
and column-writer callers propagate the now-fallible assignment. Construction
becomes infallible. Promotion must leave the inline record intact if allocation
fails, and no partial parser/edit result may be published.

Allocation error ordering intentionally changes: the map no longer reserves at
`<cols>` opening or at validator/writer construction. Dense allocation failure
can instead occur at the second valid assignment. Invalid records and existing
semantic refusals still run; a zero/single-record path can reach those refusals
without the previous allocation attempt. Empty `<cols>` remains malformed.

The experiment retains the ordinary edit and default durable save contracts,
the pinned real-file full-output oracle, synthetic one-cell/one-percent scale
guards, and the exact no-op control. Native elapsed measurements and allocator
observer measurements remain separate. The allocation benefit gate is at least
50% fewer requested bytes for the real edit; every latency/resource regression
flag must be assessed individually. This file records a hypothesis, not an
adoption decision or performance result.

Coverage limits to retain in the final result: the real corpus has one column
record; the fixed synthetic matrix does not establish nonempty column-action
latency or a general multi-producer distribution of column-record counts.
Promoted overlap semantics are covered by the grid oracle and existing parser,
writer and transaction tests. This does not make a broad speed claim for
worksheets with many column records.

The retained structural census inspected all 180 checked `test-data/**/*.xlsx`
files and counted 389 worksheet members: 43 have exactly one direct physical
column record, 159 have multiple records, and 187 have none. All 202 worksheets
with a direct `<cols>` element have at least one record. Thus the inline case
occurs beyond the pinned fixture, but most column-bearing worksheets in this
test corpus will promote. Zero-record worksheets without `<cols>` already avoid
map construction; this census does not imply a new benefit for those sheets.
It does not process MCE or establish production-parser admission, producer
market share, or latency. `census.py --check` replays every counted member and
checks the retained archive/member hashes and distribution.

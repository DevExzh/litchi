# Next OLE2/OOXML priority after the 0525 XLSX reconstruction candidate

`scope: bounded read-only queue review of the terminal 0525 candidate`

`performance_claim: none`

The 0525 native and profile admission gates passed for the measured
source-backed XLSX value-only edit/save route. The supplemental eager ABBA
confirmation also passed both median/mean guards with no new >5% flags.
The source-bound decision retains the candidate; original adverse metrics remain
visible. Cleanup and strict sealed evidence replay pass.
Keep OLE2 and OOXML ahead of ODF; the overall optimization goal remains open
and ODF stays deferred.

This review selects a source-inspection target. It authorizes no production
change, build, or new capture.

## Decision

The primary next source audit should inspect the current
[`scan_with_limit`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs#L200)
layout/provenance pass, including
[`Scanner::start_cell`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs#L952)
and
[`Scanner::cell_address`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs#L973).
This is the largest remaining owner in these XLSX commit profiles. The audit may select a
follow-up only if it finds a new, independently proven removable
representation or pass. Its instruction share alone does not establish that
any scan, address, namespace, limit, or provenance work is redundant.

The Store merge is a secondary bounded source audit:
[`Store::merge_omitted_cells`](../../../../crates/litchi-xlsx/src/cell.rs#L805),
its [`Store::from_unsorted`](../../../../crates/litchi-xlsx/src/cell.rs#L736)
rebuild, and the `Stored`/`Cell` cloning they perform. It is a current
candidate-owned phase, but its 5.97% commit-Ir share is a smaller ROI lead than
the scan/layout owner. The native timing rows put commit at roughly 35–36% of the
candidate's measured total p50; treating those shares as proportional would
give only a rough ~2% whole-route opportunity if the entire merge vanished.
That multiplication is a screening heuristic, not a measured or rigorous
wall-time bound.

The merge audit asks whether a proof-preserving operation can consume the
parsed and source-owned entries with less copying or sorting while still
rebuilding every required Store index and extent. It must first establish the
current parser's ordering, duplicate, formula, metadata, row, column, merge,
dimension, and error-precedence contracts. If either audit finds no concrete
removable work, close that target without a candidate. Do not requeue the 0522
scanner change under a different name, and do not pass writer events into the
semantic parser.

## Admission evidence and disjoint ownership

The native admission rows report the following candidate-versus-baseline p50
improvements from the four primary shape/repeat pairs:

| shape | repeat | total p50 | commit p50 |
| --- | ---: | ---: | ---: |
| dense-sparse | 1 | 13.21% | 30.09% |
| medium | 1 | 11.24% | 30.70% |
| dense-sparse | 2 | 14.90% | 30.77% |
| medium | 2 | 12.67% | 30.36% |

The profile admission also passed all four pairs. Across the four candidate
profiles, selected `MultiSourceEdit::commit` totals are 782,799,138 Ir versus
1,148,595,405 baseline Ir, a 31.85% reduction. Callgrind Ir is mechanism
evidence; it is not a latency, allocation, RSS, I/O, scaling, or native Office
producer claim.

The aggregate direct-child rows below are inclusive Ir, but the two named
edges are disjoint direct children of the selected `MultiSourceEdit::commit`
owner. Their sum is valid at that parent. Descendant rows in the next table
must not be added again.

| candidate direct child of `MultiSourceEdit::commit` | aggregate Ir | share of candidate commit |
| --- | ---: | ---: |
| `rewrite_value_only_with_provenance` | 444,163,798 | 56.74% |
| `Snapshot::from_rewritten_value_source` | 336,155,723 | 42.94% |
| other direct children plus owner self | 2,479,617 | 0.32% |
| selected commit total | 782,799,138 | 100.00% |

The rewrite branch is therefore the largest remaining direct owner, but its
output writer is a small part of that row. In the four candidate annotation
files, `scan_with_limit` accounts for 416,939,582 aggregate Ir (53.26% of
commit and 93.87% of the rewrite edge), while
`write_sheet_data_with_provenance` accounts for 6,432,005 Ir (0.82% of
commit). The scan currently owns layout discovery, implicit row/cell address
handling, namespace/tag spans, and provenance inputs; its share alone does
not show that any of that work is redundant.

## Validation, reconstruction, and Store costs

The candidate reconstruction edge has three materially different descendants:

| phase inside `from_rewritten_value_source` | aggregate path Ir | share of commit | treatment |
| --- | ---: | ---: | --- |
| complete candidate `validation::validate_xml` | 272,808,107 | 34.85% | required full-byte validation; no skip target |
| `Store::merge_omitted_cells` | 46,745,728 | 5.97% | secondary source audit |
| reduced `raw::worksheet::parse` | 15,737,919 | 2.01% | already reduced to retained XML; no fusion |

The complete validator row in the combined profile analysis is 273,461,573
Ir for the candidate, versus 273,491,342 for baseline. Its direct incoming
edges explain the difference between the combined diagnostic and the selected
commit path:

| direct caller of `validation::validate_xml` | aggregate Ir | role |
| --- | ---: | --- |
| `from_rewritten_value_source` | 272,808,107 | candidate worksheet readback in the selected commit |
| `with_invalidated_workbook` | 524,449 | workbook invalidation verification |
| `invalidated_workbook_xml` | 129,017 | workbook invalidation path |
| combined validator row | 273,461,573 | all four profiled processes and callers |

The validator scans the complete actual candidate bytes and preserves the
existing validation boundary and error order. Its inclusive share is not a
license to remove validation or fuse it with rewrite or parsing.

The merge row includes the full combined Store construction. Across the same
four profiles, `Store::from_unsorted` accounts for 26,196,915 Ir and
`Cell::clone` accounts for 13,045,089 Ir within that merge path. The helper
rebuilds sorted cells, cell-row indexes, merge indexes, and stored/content/
styled extents. A source audit may investigate a linear ownership-preserving
merge or a narrower clone boundary, but it must retain duplicate detection,
parser ordering behavior, all metadata and formula storage, and exact extents.
The measured clone/sort rows are not evidence that the full Store rebuild is
optional.

The reduced raw parser is only 15,737,919 aggregate Ir in the selected commit;
it parses the actual retained worksheet context and changed records. It is
already the part reduced by the current candidate path and should not be fused
with the writer or validator.

## Source audit boundary

Read the scanner implementation with the writer's layout consumers and the
provenance renderer. Establish, from source, whether any representation or
pass is duplicated after accounting for:

* implicit row and cell address propagation;
* namespace scope, element/tag spans, and markup boundaries;
* action validation, row/column/default effects, and all event/byte limits;
* dimension and formula dependency state; and
* exact output preservation and source-bound omission ownership.

The secondary merge audit should then read its raw-parser Store construction
and index consumers. Establish, from source, whether each of these operations
is required after combining parsed changed entries with imported source
entries:

* sorting and duplicate rejection for every address;
* row-start and cell-row index construction;
* merge-index reconstruction;
* declared, stored, content, and styled extents; and
* cloning of source-owned formula, value, style, and metadata fields.

The smallest defensible follow-up would preserve the complete candidate XML
validator, ordinary full-output writer, execution fences, limits, source
identity, staged semantic readback, calculation-chain invalidation, and
publication behavior, then measure only a concrete representation or pass
identified by one of these audits. A persistent overlay or Store
representation change is a separate design and is outside this queue item.

The scanner inspection must not repeat 0522's scanner candidate, skip full
validation, or combine scanner events with semantic parsing. Any actual
optimization would need a fresh owner-scoped profile and allocator/native
comparison after this source-only review.

## Evidence custody and remaining scope

This review is bound to the frozen plan SHA-256
`78d3fad228e4bd00476148044bbcffc8d48ec02f1e825b59c95385f92da40f97`, candidate
source manifest SHA-256
`0fab9bafc238611761659bcafdd2e7b5df2543aecb194ac5477c168a2d6e6096`, combined
profile-analysis SHA-256
`24ec2d392d2d6e6e35f14240c571ac62fd03d29414be1ecb9b6a916891cb7c12`, candidate
profile-analysis SHA-256
`81cc89e5557fbe538f1344b551284999b6131d3de2da5840533e9a6b658899b5`, and
comparison SHA-256
`4190c055763e835c5e5c951c57e9516736ff6943ef44d6baf1bec9fe1aaccc5a`.
The source-bound artifacts are [`profile-analysis.json`](profile-analysis.json),
[`candidate/profile-analysis.json`](candidate/profile-analysis.json), and
[`comparison.json`](comparison.json); their Callgrind descendants are the
four `candidate/profile-*.inclusive.txt` files.

The admission result applies to the synthetic source-backed XLSX route and
its declared guards. It does not close provider, physical-I/O, cold-cache,
native-producer, scaling, broader CRUD, fuzz, or OLE2 coverage. OLE2/OOXML
remains the active priority, ODF remains deferred, and no global optimization
goal is marked complete by this queue review.

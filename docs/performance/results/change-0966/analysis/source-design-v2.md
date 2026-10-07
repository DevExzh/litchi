# DOC SPRM reservation and reparse design note (0966)

**Status:** read-only source review, 2026-10-07. No production edit was
performed. The ranking below is a mechanism and actionability ranking for this
narrow question; it is not a global performance ranking. The later 0966
workload is bound by the separate capture and reduction packet.

The repository packet copy normalizes the historical performance-record links
below relative to this file. The sealed external v2 note is preserved at
`/home/zhuhe/litchi-performance-0966/doc-sprm-design-v2.md`.

## Finding

At the current source (`HEAD a499d34f51`),
`crates/litchi-doc/src/sprm.rs:172-186` reserves
`max(grpprl.len() / 4, 8)` `Sprm` elements before it knows whether the input
has any records. Therefore every input shorter than 32 bytes requests the
minimum eight-element capacity. The 0919 allocation report observed 19,781
384-byte allocation events among 21,213 `parse_sprms` invocations on
FloatingPictures/32, and 1,154 allocation events among 1,168 invocations on
NoHeadFoot/32
([`docs/performance/0919-doc-allocation-attribution.md:85-95`](../../../0919-doc-allocation-attribution.md)).
The numerators count allocation events, including operand allocations, within
the inclusive parser subtree; they are not parser-call counts. For scale only,
the arithmetic ratios are 93.2 allocation events per 100 recorded invocations
for FloatingPictures/32 and 98.8 per 100 for NoHeadFoot/32. Multiple events
can belong to one invocation, so these ratios are not call fractions and do
not estimate a saving. The counts also do not say whether the inputs were
empty, valid short sequences, or malformed tails.

The smallest semantically closed candidate is an empty-input fast path:

```rust
if grpprl.is_empty() {
    return Ok(Vec::new());
}
```

An empty group is already a successful empty SPRM sequence. This removes the
minimum reservation for exactly that case and leaves every non-empty parse,
including malformed-input error behavior, on the existing path. The parallel
0966 census should decide whether the empty bucket is large enough to justify
even this small change. A broader change that removes `.max(8)` for all short
inputs is plausible, but its reallocation cost and malformed-input mix are
unknown; it should remain a measured option.

## Current call graph and stale attribution

The three names that appear together in historical allocation or sample
stacks are not three SPRM consumers:

| Stack name | Current work | Consequence for a SPRM change |
| --- | --- | --- |
| `source_paragraphs` (`body_text.rs:4276-4311`) | Strictly decodes the main story, scans paragraph marks, and calls `is_in_table_at_cp`. The latter uses ordered CLX/PAPX containment and the cached `PapxRun.in_table` scalar (`package.rs:1665-1682`). | It does not call `parse_sprms`. The 0056 containment index and 7b89 table-membership cache are already current. Do not add a parser or re-open this accepted work. |
| `replace_plain_text` (`package.rs:1466-1543`) | Checks raw CHPX group equality, clones the editor, edits pieces, appends a raw `FcRun`, and calls `rewrite_chpx`. | It does not parse SPRMs itself. The edit guard before it calls `text_revision_intersects`, which calls the full `revisions()` scan; that ancestor is the parser cost. |
| `rewrite_chpx` (`package.rs:2886-2929`) | Validates run lengths, clones raw `grpprl` bytes into `ChpxFkpBuilder`, emits pages, and appends the CHPX table. | It does not parse SPRMs. A parsed-vector cache would not remove its raw copies or page/output allocations. |

The actual strict-parser consumers at this boundary are:

* `parse_chpx` validates each raw CHPX group (`codec.rs:453-483`), and
  `parse_papx` validates each PAPX body while deriving `in_table`
  (`codec.rs:486-525`).
* `RevisionEditor::open` then calls `revisions()` for the resource-limit check
  (`package.rs:1117-1119`); `revisions()` parses every CHPX and PAPX run again
  (`package.rs:2341-2406`).
* `Edit::replace_text` calls `text_revision_intersects`
  (`body_text.rs:4383-4398`), which runs that full revision scan before the
  raw `replace_plain_text` mutation. Other property and picture guards parse
  selected CHPX groups (`package.rs:1437-1463`, `1697-1722`, and
  `1956-1973`).

This makes a second, higher-cost candidate possible: an operation-local parse
handoff or compact summary from strict CHPX/PAPX validation to the revision
guard. It must be designed around the current raw-byte model. Retaining a
`Vec<Sprm>` in every `FcRun` would increase editor retention and clone cost,
and would require invalidation after every raw group mutation. The safer shape
is an ephemeral summary or one-shot parse result whose lifetime ends before a
mutation. It can only be considered after the census proves the same run is
parsed twice in the same operation and after the full revision output and
error ordering are accounted for.

## Ranked next actions

| Rank | Candidate | Cost that could be removed | Next action | Confidence / history |
| ---: | --- | --- | --- | --- |
| 1 | Empty `grpprl` fast path | One eight-`Sprm` capacity request per empty parse (384 bytes in the 0919 histogram) | Gate on the 0966 empty/non-empty census; if material, change only `parse_sprms_two_byte` before reservation and run the existing differential/error tests. | High semantic confidence; prevalence unmeasured. No prior empty-specific trial. |
| 2 | Replace the blanket minimum for short non-empty groups | Excess capacity for `1..31` byte inputs, balanced against possible `Vec` growth | Compare requested capacity, parsed-record count, success/error, and repeat behavior by length bucket; retain the current minimum if short sequences commonly contain several records. | Medium; 0919 cannot distinguish this bucket. No adopted trial. |
| 3 | Share strict parse results across open/edit guards | Repeated `Sprm` vector construction and scans: CHPX/PAPX validation followed by `revisions()`, and per-edit `text_revision_intersects` scans | First identify exact duplicate run identities and the required consumer fields. Prototype an operation-local handoff or compact summary only; do not cache full vectors in `FcRun`/`PapxRun` yet. | Medium-low until duplication is measured; this is a new design, not a revival of 0596. |

## Semantic constraints

1. `parse_sprms` must continue to consume the exact whole sequence and report
   the first malformed record. Do not skip a one-byte opcode tail, a truncated
   fixed operand, a zero-length variable record, or any other non-empty input
   merely because it is short. The empty fast path is valid only for
   `grpprl.len() == 0`.
2. PAPX still requires its two-byte style index before SPRM parsing
   (`codec.rs:504-508`). An empty PAPX body after that prefix is valid, while a
   zero- or one-byte PAPX group remains the existing `PAPX style index is
   missing` error.
3. `papx_in_table_state` must preserve the last duplicate
   `sprmPFInTable` record and literal-byte-`1` rule (`codec.rs:358-372`). Any
   PAPX transform must recompute the scalar only after strict success; 7b89's
   atomic `split_transform_papx` behavior is part of the contract.
4. Raw `grpprl` bytes, offsets, and record sizes remain authoritative for
   `retain_sprms`, wall restoration, revision metadata, and picture-location
   lookup. A parsed summary cannot replace those bytes or become stale after
   `split_transform_chpx`, `replace_plain_text`, or revision transforms.
5. `RevisionEditor::open` must retain the existing `MAX_REVISIONS` guard and
   `revisions()` must retain its complete CHPX/PAPX metadata, sorting, merging,
   and move inference. A specialized text-revision scanner cannot silently
   replace full validation unless it proves identical error precedence and
   resource-limit behavior.
6. The source-paragraph path must retain strict text decoding, CP-length
   validation, table-cell filtering, and cached table membership. Its current
   containment and scalar-cache work is already accepted history, not an
   unmeasured SPRM opportunity.

## Prior trials and disposition

The 0596 eager-open change already removed repeated parsing in the old
`parts/pap_bin_table` path. Its record requires reuse only when the concatenated
bytes are byte-identical and the existing consumption/error order is preserved
([`docs/performance/0596-doc-eager-open-terms.md:53-74,130-140`](../../../0596-doc-eager-open-terms.md)). That code path is not the current
`tracked_revision` parser, so reviving its per-entry design would be stale.

Likewise, 0056 already changed paragraph table checks to indexed containment
([`docs/performance/changes/0056-doc-papx-containment-index.md:18-45`](../../../changes/0056-doc-papx-containment-index.md)), and 7b89 added the cached PAPX table-membership scalar. Neither should be counted again as a missing source-paragraph optimization.

0942 reports `parse_sprms` at 5.7% of the FloatingPictures/32 marker subset,
but explicitly says those weights are not removable-work estimates
([`docs/performance/0942-doc-cfb-operation-profile.md:78-90`](../../../0942-doc-cfb-operation-profile.md)). The actionable decision here is therefore the empty/short reservation census first; the only currently source-closed production candidate is the empty-input fast path, with parse sharing held behind exact duplicate and semantic evidence.

## Coverage limits

This review covers the current DOC tracked-revision source and the 0596,
0056, 7b89, 0919, and 0942 records. It does not establish latency, peak-live
memory, or cross-format behavior, and it excludes iWork. The 0919 evidence is
allocation attribution over its retained fixtures; the 0942 evidence is a
sampled marker subset. No claim here generalizes beyond those cohorts until the
new 0966 empty/non-empty census supplies coverage.

# Next experiment: bounded DOCX body-scan fusion

Status: review-only recommendation. No production change or performance claim
is authorized by this note. The 0719 batch strengthens the producer-shaped XLSX
output oracle and uses this file to record the next work-elimination boundary.

The next substantive candidate is the duplicate structural walk in the DOCX
writer's managed body parser. The exact current entry point is
`crates/litchi-docx/src/writer/doc/package.rs:85`,
`ParsedDocumentBody::from_xml`. After removing a leading BOM, that function
currently does the following:

1. `scan(bytes)?` calls `crate::alt::scan`, which parses the complete XML for
   `altChunk` metadata and then runs the shared MCE active-offset selection.
2. The loop beginning at `package.rs:118` calls
   `active_block_ranges(bytes)?`, which performs another complete
   `scan_word_element_ranges` walk and then runs active-offset selection again
   for paragraph, table and `altChunk` starts.
3. The same function performs its own reader walk to capture the body prefix,
   direct children and suffix. That third walk is outside this candidate.

`crates/litchi-docx/src/source_backed.rs:1730` is a separate consumer:
`alt_chunk_target` calls `crate::alt::scan` and `body_block_ranges`. It should be
tracked as a later consumer of any shared helper, but it is outside the first
writer experiment and must not be folded into its result denominator.

## Existing evidence

Current source confirms that the duplicate walks remain. Historical profiles
re-analyzed in 0712 justify a bounded qualification, but are not a fresh profile
of 0719 source and do not predict a saving. Those profiles attribute the
structural owners as follows:

| Corpus | `active_block_ranges` inclusive Ir | selected structural scan Ir | MCE active-offset owner Ir |
| --- | ---: | ---: | ---: |
| generated medium | 3,449,750 | 2,294,438 | 1,068,977 |
| `NumberedList` | 2,663,144 | 415,959 | 2,245,011 |

The inclusive rows contain nested work and must not be added together. They do
show that both corpora pay a complete range walk in addition to the alt scan,
and that the remaining cost has a different shape on the real fixture. Change
0713's `memchr::memmem` search is already part of the current baseline; a new
experiment must not count its earlier scalar search as removable work.

The narrower 0711 event/resolver ownership candidate is not the proposed
change. It was rejected because one `NumberedList` edit p50 improved only
2.888465327%, below its frozen 3% floor, even though the generated edit and
the other real-fixture pair improved. Change 0715's section-publication
collection fusion is also rejected and has a different source boundary. The
0718 Deflate allocation attribution does not justify allocator policy or
cross-member pooling; preservation pooling was already rejected by 0618.

## Smallest safe candidate

The first candidate should fuse only the structural event walk. It should
collect the `altChunk` metadata and the paragraph/table/`altChunk` ranges in a
single source-order pass, then feed those collected offsets to the existing
selection and body-capture stages. It must retain **two separate MCE
active-offset calls**:

1. Run the existing `altChunk` selection first, with exactly the offsets that
   `crate::alt::scan` would have supplied, and filter the collected chunk map.
2. Run the existing block-range selection second, with exactly the offsets
   that `active_block_ranges` would have supplied, and filter the collected
   ranges.

Do not union the two offset vectors in this experiment. A union can change
`max_offsets`, marked-byte accounting, marker placement, and the first refusal
returned by `active_offsets`; 0712 explicitly identifies those as semantic
boundaries. Keeping both calls also makes the first result attributable to one
removed structural walk rather than to an altered MCE policy.

The combined scanner must preserve the stronger error ordering of the current
sequence. `alt::scan`'s XML and anchor errors, followed by its active-selection
error, must remain ahead of any range-only error. A range scanner error may be
recorded while collecting, but it must be returned only after the alt scan's
structural and active stages have succeeded, at the point where the current
`active_block_ranges` call would have run. The body capture walk remains after
both selections. No error, limit, or output may be made dependent on whether a
range happened to be needed by the selected edit.

The proof must cover both scanners' existing limits and rules: BOM-adjusted
offsets, `MAX_XML_BYTES`, XML depth, `MAX_SCAN_DEPTH`, `MAX_SCAN_NODES`,
`MAX_CHUNKS`, duplicate anchors, strict and transitional Word namespaces,
unknown direct body children, nested anchors, `mc:Choice`/`mc:Fallback`,
`MAX_VISIBILITY_OFFSETS`, `MAX_MARKED_XML_BYTES`, MCE output limits, malformed
XML, and typed relationship/anchor errors. The collected ranges and chunks must
be byte-identical to the two existing outputs, including source order and
duplicate/refusal behavior. Source freshness, cancellation, preservation,
atomic refusal and the final body capture remain unchanged.

## Measurement gate

Before editing production code, add a source-bound diagnostic that records the
current two structural pass counts, event/input-byte counts, the two MCE call
inputs and outputs, and the final range/chunk digests at the exact
`ParsedDocumentBody::from_xml` boundary. This establishes the current
denominator and catches a scanner that quietly changes its effective scope.

If that diagnostic confirms the duplicate work, freeze an A/B pilot on the
current source with generated-medium and the admitted `NumberedList` fixture,
plus BOM, marker-bearing, nested-anchor, malformed and typed-refusal controls.
Use separate native edit and lifecycle clocks, an operation allocator lane,
and exact output/source/resource oracles. The native gate should retain the
0711 discipline: paired edit p50 and mean improvements must be shown for both
corpora, lifecycle p50/mean must not regress beyond the declared bound, and
every greater-than-5% phase or tail movement remains visible. The diagnostic
must also prove that the MCE call count remains two and that both calls receive
the same per-call offset vectors as baseline.

No production implementation, cache, new public API, borrowed event lifetime,
or shared offset representation should be introduced until that plan passes
independent review. A successful pilot would remove one complete structural
walk while retaining the existing MCE proof and the final body parse; a failed
pilot should leave the current writer byte-for-byte unchanged.

This recommendation is a work-elimination experiment for the DOCX writer. It
does not reopen the rejected 0711 or 0715 candidates, does not infer a native
speedup from historical Callgrind percentages, and does not continue the 0718
fault-attribution lane.

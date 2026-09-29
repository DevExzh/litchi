# 0834 — PPTX cold replay proof design review

This is a read-only design review for the pending PPTX cold-classifier repair.
It authorizes no measurement or performance claim. The scope is the
filesystem harness replay evidence; the existing DOCX proof and production
PPTX reader remain outside this repair.

## Evidence and the current failure boundary

`PptxReplaySource::record` currently classifies the returned range
`offset..offset + count` against the compressed slide and media ranges. It
retains aggregate counters, merged coverage, and return sizes, while the
pending diagnostic extension adds raw request records. The source wrapper has
one stable `SourceVersion` and the payload locator maps slide names to
presentation positions before it permits a replay. See
[`filesystem.rs`](../../../../tools/perf-baseline/src/filesystem.rs#L3797),
[`filesystem.rs`](../../../../tools/perf-baseline/src/filesystem.rs#L3833), and
[`filesystem.rs`](../../../../tools/perf-baseline/src/filesystem.rs#L4109).

`replay_pptx_source` constructs the source-backed presentation first and only
then performs the selected semantic operation. Its current classifier combines
the constructor reads and semantic reads, so a metadata read at the end of an
aligned ZIP can appear to be an illegal media or unselected-slide read. The
`open_read_count` captured immediately after construction, together with the
chronological raw vector from the same immutable source, is an exact phase
boundary: records before that count are `open_catalog`, and records after it
are `semantic_query`. The boundary must be validated against the retained
vector and the source envelope. Repeating source id, mode, or phase on every
record is unnecessary. See
[`filesystem.rs`](../../../../tools/perf-baseline/src/filesystem.rs#L4160) and
the call site after the timer at
[`filesystem.rs`](../../../../tools/perf-baseline/src/filesystem.rs#L2837).

The ZIP locator explains why the aligned source needs separate treatment. It
first reads the final fixed 22 byte EOCD record and accepts it only for a
zero-comment, non-ZIP64 record. Otherwise it performs a bounded backwards
search in buffer-sized reads, carrying a signature split across chunk
boundaries. Those are structural metadata reads; any overlap with a compressed
payload is a coincidence of physical ranges and is not decompression. See
[`locator.rs`](../../../../crates/soapberry-zip/src/locator.rs#L687) and
[`locator.rs`](../../../../crates/soapberry-zip/src/locator.rs#L1466). Existing
tests already pin both paths: a zero-comment archive reads the fixed record
first, and a commented archive falls back to the backward-search chunk
([`locator.rs`](../../../../crates/soapberry-zip/src/locator.rs#L2337)).

The verified-cold copy changes only ZIP framing: `page_aligned_archive`
extends the EOCD comment and appends zero bytes while retaining the EOCD
position and the logical archive members. The verifier separately proves
page alignment, source identity, pre-operation residency, post-operation
identity, and a positive process `read_bytes` delta. It does not prove a
physical-media read. See
[`cold_verified.rs`](../../../../tools/perf-baseline/src/cold_verified.rs#L523)
and [`cold_verified.rs`](../../../../tools/perf-baseline/src/cold_verified.rs#L717).

## Minimal evidence contract

The repair should keep the existing aggregate counters and add only the state
needed to explain a failed classification:

* The immutable replay envelope carries one source id and revision, operation,
  child mode, source length, and source SHA-256. Raw records remain
  chronological and carry offset, requested length, returned length, and a
  checked returned range. A zero-byte/EOF return is retained as a record with
  an empty returned range. The records are bound to that one immutable source;
  there is no need to repeat the envelope tokens on every record.
* Capture `open_read_count` immediately after the source-backed constructor
  and before the semantic query. Validate that it is a count boundary within
  the chronological raw vector; project records before it to `open_catalog`
  and records after it to `semantic_query`. This is an explicit phase
  boundary, rather than an inference from offsets or payload overlap. The
  review does not require new constructor-error propagation or arbitrary
  failure diagnostics.
* Raw records remain chronological. Aggregate and coverage fields may remain
  derived, but the failure diagnostic must include enough information to
  recompute them from the raw returned ranges. Requested ranges and returned
  ranges must not be conflated.
* The diagnostic is bounded and content-free. It may retain range metadata and
  overlap labels, but never source bytes, XML, paths, or external-tool output.
  An exceeded diagnostic bound is itself a replay failure rather than a reason
  to discard records.

The existing source identity token is sufficient for this purpose: the replay
source creates one `SourceVersion` at construction and returns it from every
`version()` call. Exposing `id()` and `revision()` in the diagnostic does not
change the public Office API or claim that two independent processes share an
identity.

## Classification rules

The classifier should evaluate the constructor and semantic phases separately,
then derive the operation result from both. The following rules retain the
current semantic contracts:

| Operation | `open_catalog` phase | `semantic_query` phase |
| --- | --- | --- |
| `open`, slide count, and open+slide-count lifecycle | No semantic slide or media payload. Ordinary sources require zero payload overlap. An aligned source may have exactly the bounded EOCD-tail allowance below. | No query payload. |
| Selected slide and open+selected-slide lifecycle | Ordinary sources require zero payload overlap during open. An aligned source may have only the bounded EOCD-tail allowance. | The selected compressed slide range is completely covered; unselected slides and media have zero overlap. |
| List slides | Ordinary sources require zero payload overlap during open. An aligned source may have only the bounded EOCD-tail allowance. | Every slide range is completely covered; media has zero overlap. |

The aligned allowance is legal only for the `VerifiedPrime` and
`ColdVerified` source modes after the aligned source hash and size have been
bound to the prepared corpus copy. `Prime`, `Warm`, and advisory `Cold` use
the ordinary unaligned source contract even if an unrelated file happens to
have a page-aligned length. Mode text alone is not proof of alignment.

For the current aligned fixture, define
`tail = [aligned_length - RECOMMENDED_BUFFER_SIZE, aligned_length)`. The
open-phase proof must show one exact returned read of `tail` and must compare
the open-phase payload overlaps with the bytewise overlaps of that exact range
against each classified payload set. A fixed final 22 byte probe may also be
present, but it carries no allowance unless it is part of the same proven
bounded locator sequence. Any open-phase payload overlap outside the exact
locator reads, any second unaccounted tail read, or a short/missing read where
the expected exact tail is required fails classification. Incidental
zero-byte/EOF metadata records do not fail by themselves; they remain visible
and must not introduce extra payload overlap. This mirrors the existing DOCX
verifier’s exact-tail boundary without weakening it or making it a general ZIP
rule.

The semantic query must never inherit the open allowance. A selected-slide
query that reads an unselected slide or media fails even when the same raw
range would have been explainable as an EOCD probe during open. Conversely, an
open tail that overlaps media is retained as raw media overlap and accepted
only when the aligned geometry predicts that exact overlap. It must not be
subtracted from the counters or relabelled as semantic I/O.

The proof of the aligned geometry must remain independent of the read
classification:

1. Bind the aligned source hash and length to the source prepared by the
   parent, and bind the unaligned corpus hash and length to the pinned PPTX
   corpus.
2. Prove that the EOCD offset is unchanged, the aligned length is the
   unaligned length plus the computed padding, and the EOCD comment length is
   increased by exactly that padding.
3. Prove byte equality outside the comment-length field and the appended
   suffix, and prove that the suffix consists entirely of zero bytes.
4. Derive the expected per-class overlap of the exact tail from the parsed
   aligned archive. Do not accept a caller-supplied overlap count without
   deriving it from the source ranges.

This permits a source-specific metadata explanation while preserving the
existing refusal for extra unselected-slide or media reads. It also keeps
ordinary unaligned operations useful as a control: their payload overlap must
remain zero during open.

## Source-grounded tests

The smallest useful test set is structural and can run before any workload
capture:

1. Keep the existing ZIP locator tests for a zero-comment fixed probe and a
   nonzero-comment backward-search fallback. Add a replay fixture assertion
   that the fixed 22 byte attempt and the bounded tail search are both marked
   `open_catalog`, never `semantic_query`.
2. Exercise `pptx_payload_ranges` with a duplicate slide member and with one
   missing slide member. Both must fail before classification; neither may be
   converted into a zero-payload catalog result. The existing name-to-position
   slot checks are the authority for these cases.
3. Feed an aligned selected-slide replay with one exact open tail read whose
   predicted overlap is nonzero, followed by a complete selected-slide query.
   It must pass and retain the raw overlap in the open phase. Move that same
   tail read into `semantic_query`; it must fail.
4. Add an extra unselected-slide read and an extra media read to the selected
   query phase independently. Each must fail, even when the open phase has a
   valid tail read. Add the same extra reads outside the exact tail in the
   open phase; each must fail as well.
5. Add an ordinary unaligned selected-slide control with the same extra tail
   overlap. It must fail because the aligned allowance is not inferred from
   length or overlap alone.
6. Mutate the aligned copy in three independent ways: change a byte in the
   zero suffix, change the EOCD comment-length field, and move or truncate the
   EOCD tail. Each must fail the alignment proof before any cold result is
   admitted. A missing exact tail read and a segmented/short read where that
   exact tail is required must likewise fail closed and retain the diagnostic.
   Incidental zero-byte/EOF metadata records remain acceptable when they do not
   create extra payload overlap.
7. Assert that the diagnostic has one immutable envelope source id/revision,
   that `open_read_count` is a valid boundary in the chronological raw vector,
   and that every record has a valid checked returned range. Derive the phase
   from that boundary (or compare any serialized phase labels with the
   derivation). Recompute aggregate returned bytes and each payload overlap
   from the records and compare them with the serialized counters. This catches
   requested-versus-returned range drift.

The first four tests can be pure harness/source-range tests. They do not need
to run a cold filesystem operation or make a timing claim. The existing DOCX
aligned verifier should receive no changes in this batch; broad short-read
handling, range coalescing, and a new DOCX source matrix are separate work.

## Admission boundary

After the diagnostic and proof tests pass, run one fresh warm plus
`cold-verified` qualification for the PPTX lifecycle and inspect the retained
raw phase vectors. A successful subprocess is insufficient: the source hash,
alignment proof, semantic digest, no-extra-overlap rules, strict post-fincore
observation, and positive `read_bytes` gate must all pass. Only then should the
unchanged 0833 matrix be reconsidered. Failed 0833 requests remain failure
evidence and cannot be promoted into formal samples.

This design follows the accepted contracts in
[`ADR 0005`](../../../../docs/adr/0005-io-memory-and-performance.md),
[`ADR 0006`](../../../../docs/adr/0006-validation-security-and-compatibility.md),
and [`ADR 0008`](../../../../docs/adr/0008-migration-and-verification.md):
positional sources retain stable identity, preservation evidence remains raw,
validation fails closed, and measured claims require representative gated
evidence.

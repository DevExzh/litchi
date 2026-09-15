# Log paragraphs for change 0596

The coordinator merges these into the four shared logs. Nothing here edits those
files directly.

## HOTSPOTS.md

**DOC eager open, the three size-independent terms (items DOC-2, DOC-3 and DOC-4
of change 0587).** Change [0596](../../0596-doc-eager-open-terms.md) implemented
all three. `TextExtractor::new` no longer walks the decoded text four times: the
UTF-16 buffer is sized from the pieces' clipped extents, Windows-1252 decodes
through a 256-entry `u16` table instead of one `encode_utf16` plus
`extend_from_slice` per character, and `from_utf16_lossy` plus the
`encode_utf16().count()` plus the `char_indices` walk that built `cp_to_byte`
collapse into one `char::decode_utf16` pass. `PapBinTable::parse` parses each
entry's `grpprl` once instead of about three times, keeps a bounded
least-recently-used map of resolved style baselines instead of change 0051's
single adjacent-style slot, and reuses one `PrcData` visit set. The FIB reads
through an `Arc<Vec<u8>>` of the `WordDocument` stream at an offset instead of
owning a copy of the suffix, so the eager open no longer copies the stream a
second time and the attached glossary FIB no longer copies it a third.
**Measured** per-open instructions on the retained 0587 driver: 5,609,212 ->
4,481,165 (-20.11%) on `saved-by-table.doc`, 3,471,280 -> 3,080,182 (-11.27%) on
`FloatingPictures.doc`, 4,971,153 -> 4,058,410 (-18.36%) on the 1.6 MB
kwsymphony form; `parse_sprms` calls per open 1,931 -> 686, 965 -> 531 and
795 -> 467; `memcpy` calls 10,722 -> 8,955, 7,349 -> 6,615 and 8,754 -> 7,438.
What remains in `TextExtractor::new` is the `cp_to_byte` table itself, still
eight bytes per UTF-16 code unit and still built eagerly; the fused decode is
still the largest single term of the `saved-by-table.doc` open at 38.33%
inclusive and 1,336,150 instructions self. What remains in `PapBinTable::parse` is the
per-entry `[sprms, piece_modifier].concat()` (one allocation per entry, retained
in the run) and the `TapParser`'s own re-parse of the same `grpprl` for
table paragraphs. Neither is addressed here.

## GOAL_AUDIT.md

**A ninety-instruction edit moved one harness scenario by fifteen percentage
points.** Change 0596 removed 19.8% of the instructions of the 1.6 MB DOC open
and 21.2% of its native cycles — the two agree because the removed term is a
`memcpy` of a 697,827-byte stream. But its first candidate build was 13% to 15%
worse at p50 on `doc_semantic_paragraph_count`, reproducibly, in two independent
windows against a floor under 0.3%, on a path where the query's own instruction
count moved by +0.03% and its native cycles by -0.66%. 99.98% of that query is
`ParagraphExtractor::count_paragraphs_in_range`, source the change does not
touch, which the candidate build inlines differently. Hoisting one FIB slice
resolution out of `get_all_subdoc_ranges` — worth 299 to 289 instructions per
call — turned that scenario into -2.63% and -0.32% in the next two windows. The
lesson for this audit is that at these scenario sizes harness percentiles rank
candidates rather than measure them, and that the A/A floor must be read per
scenario and per window, not once per run: in one window of this batch the floor
was -10.26% at p50 on `doc_semantic_open/large` and -0.11% on
`doc_semantic_paragraph_count/large` at the same time, and in another it reached
-80% at p95.

## REPORT.md

- [`0596-doc-eager-open-terms.md`](0596-doc-eager-open-terms.md) — the three
  size-independent terms of the eager DOC open: the four-pass text decode, the
  triple PAPX resolution and the FIB's copy of the `WordDocument` suffix.
  Retained, `performance_claim: none`. Per-open instructions fall 20.11%,
  11.27% and 18.36% on three real fixtures; native cycles fall 13.3% and 21.2%
  on two of them and rise 1.97% on the third. Across two windows against the
  committed candidate, `doc_semantic_open` p50 falls 4.7% and 5.4% on the tiny
  corpus and `doc_semantic_one_edit_save` p50 1.0% to 3.6%. The record also
  carries a 13% to 15% regression on `doc_semantic_paragraph_count` that the
  first candidate caused, the counts that showed it was not added work, and the
  ninety-instruction edit that removed it. The oracle is a differential digest
  over all 57 `.doc` fixtures:
  byte-identical text, identical resolved paragraph, run and section property
  structures, identical paragraph counts, identical FIB bytes, and identical
  typed refusals on the 15 fixtures that refuse.

## ADR_COMPLIANCE.md

**ADR 0006 (lossless preservation) and ADR 0003 (bounded resources), DOC eager
open.** Change 0596 is value-identical by construction and checked by a
differential digest over every admitted `.doc` fixture: `Document::text()`,
every paragraph's resolved properties and runs, the section table, the
subdocument ranges and `FileInformationBlock::raw_data()` all hash identically,
and the 15 typed refusals keep their exact messages. The FIB's shared-stream
representation preserves `raw_data()` byte for byte, so the public
`smart_tags::Snapshot::fib_bytes`, its `finish()` output and its `fingerprint()`
are unchanged. Two bounded-resource notes: the new UTF-16 reservation is sized
from the bytes the stream actually holds rather than from the character counts
the piece table claims, so a hostile piece table cannot make it reserve more
than it will fill; and the style-baseline cache is capped at 64 entries with
least-recently-used eviction, so a stylesheet naming thousands of styles cannot
grow it without bound, while a run of same-style paragraphs still resolves its
baseline once — the property change 0051 established. Retention falls rather
than rises: the eager `Document` held the `WordDocument` stream twice (once in
the document, once inside the FIB) and now holds it once.

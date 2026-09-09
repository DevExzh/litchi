# DOCX bounded tail append benchmark contract

This benchmark compares the existing materialized plain-paragraph-copy tail
route with the bounded source-backed plain-text tail append route. It measures
the complete per-iteration lifecycle in isolated normal and allocator
processes, while all archive reopening, semantic checks, physical-member
checks, and independent byte oracles remain outside the timed region.

## Corpus and operation equivalence

The three frozen source sizes are 64, 8,192, and 131,072 body paragraphs.
Each source is generated deterministically from one XML template and one
opaque binary member. The first source paragraph is the explicit authored
tail text; every other paragraph has deterministic index text. The control
copies that first paragraph to the body tail through the existing materialized
`paragraph_copy` API. The bounded route receives the same UTF-8 text as an
explicit caller-owned plain-text argument borrowed from the corpus. The text
value and its storage are prepared outside timing; bounded fragment encoding,
admission, scanning, and publication remain inside the measured lifecycle.
This makes the resulting semantic paragraph sequence identical while keeping
the two authoring routes explicit. The control is never described as authoring
new text.

The initial comparison corpus has no final `w:sectPr`, so both routes exercise
the same insertion point. Section-property placement, strict namespace
prefixing, malformed-input refusal, and the bounded append limit matrix remain
separate correctness cases. A later section-enabled comparison must add a
materialized control with equivalent placement before it enters the measured
matrix.

The generated source archive contains the ordinary Word main part, its
relationships/content types, and one 32 KiB opaque member. The opaque member
is never read into the timed operation. The candidate oracle checks that it is
byte-identical in the source and both outputs, and that all non-main physical
members retain their raw local records, compressed payloads, central records,
and ordering. The only normalized field is each untouched central record's
four-byte local-header relocation offset, which necessarily moves when the
main member grows. The main member is checked by decoded XML and semantic paragraph
digests. The materialized route preserves the source paragraph's lexical
`<w:t>` form, while the bounded authoring route records its generated
`xmlns:w` plus `xml:space="preserve"` form; each route has an independent
raw-XML oracle and the route-level record reports that raw check. Their
compressed bytes may differ and are reported separately rather than treated as
a correctness failure.

## Measured lifecycle

For each iteration, both routes own source adapter creation, source-backed
package opening, admission/scan, edit preparation, commit/plan finalization,
fresh sequential publication to a scalar-only SHA-256 sink, sink finalization,
and destruction. The measured sink retains only byte count, write-call count,
write-size histogram, and digest. It accepts short writes at a fixed maximum
chunk so output code cannot rely on one particular writer chunking pattern.

The materialized route uses the existing
`plain_paragraph_copy_snapshot_with_limits` + `copy_plain_paragraph` + commit
and publication API. The bounded route uses the public DOCX tail-append plan
and publication API once that API is available. Corpus creation, independent
candidate construction, complete archive reopen, semantic verification,
untouched-member comparison, and route-independent expected output generation
are outside timing.

The harness reports one total-lifecycle sample series per route. Source adapters expose scalar
positional-read calls, requested bytes, returned bytes, and a fixed
request-size histogram. Allocator builds report callback calls, requested
bytes, reallocations, failed calls, and peak live bytes. Normal processes
remain the authority for elapsed time and RSS.

The bounded route uses one fixed parser profile at every source size:
`max_token_bytes=4096`, `max_depth=16`, and `max_workspace_bytes=2097152`.
The source, candidate, event, paragraph, fragment, and output ceilings scale
with the selected source size; this parser window does not. The generated
corpus has maximum direct depth five and lexical tokens far below 4 KiB, so
the fixed profile covers all three cases while keeping parser reservation
independent of paragraph count.

## Required proofs

Before formal samples the harness independently proves, for every count and
route:

* source archive bytes and source main XML digest match the generated corpus;
* exactly one paragraph is appended, with the caller text equal to the source
  template and the output order/text digest equal across routes;
* all untouched physical ZIP members are exact, including the opaque member;
* the output main XML is well-formed and the final semantic paragraph has the
  expected text;
* source bytes are unchanged after each operation;
* the source and candidate archive digests are captured from fresh decoded
  outputs, not inferred from timing sinks; and
* any route-specific compact proof (source/candidate lengths, hashes, paragraph
  and event counts, and insertion offset) is checked outside timing. The
  bounded proof is retained as scalar JSON fields; the materialized route has
  no bounded proof object.

The report records source-read counters and output sink digests for every
retained sample. A digest mismatch, source mutation, semantic mismatch,
untouched-member mismatch, or short publication error aborts the run instead
of producing a partial performance result.

## Route and comparison labels

The schema must identify `materialized_paragraph_copy` and
`bounded_plain_text_tail_append` explicitly. It must also record whether the
main compressed payloads are byte-equal and the decoded main XML digest for
each route. Route timing is a same-executable comparison: it does not compare
different source revisions, and it does not include DOCX package integration
outside the focused public operation.

The benchmark does not claim a complete RSS bound from these counters. Source
storage, package index/catalog memory, ZIP codec state, caller text storage,
allocator metadata, and process setup are reported or documented separately.
It also does not claim that the bounded route supports arbitrary DOCX markup;
the route's typed admission and refusal contract is part of the recorded
corpus shape.

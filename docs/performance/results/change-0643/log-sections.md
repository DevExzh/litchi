# Change 0643 log paragraphs

The coordinator merges these into the four repository logs. Written in the style
of each log's newest section.

## For `HOTSPOTS.md`

## 0643 — the DOCX sink read stops copying every event, and change 0592's open regression is named

**The DOCX main-document sink read stops copying every event (change 0643).** Change
0229 made `for_each_word_text_chunk` borrow its events out of the pinned
main-document slice; its sibling on the sink path,
`write_text_to_with_operation_check`, kept reading each event into a caller buffer
and then calling `event.into_owned()` on it — two copies and one heap allocation
per event, on a source that is a byte slice. Survey item **XML-6 is answered**.
`NsReader::read_event_into` and `NsReader::<&[u8]>::read_event` are one-line
wrappers over the same `read_event_impl` and the same `ReaderState`, and
`preflight_semantic_xml` in the same file has always used the borrowed form on the
same bytes, so the emission loop was the only pass that did not. Measured on
`taskset -c 18`, isolation pairs at three shapes: `write_text_to` loses
**88.84%** of its allocations (90,051 → 10,045 at 10,000 paragraphs; 1,851 → 245
at 200; 267 → 69 at 24) and **70.09%** of its allocated bytes, **9.67%** of its
instructions, and **15.57%** of its median wall time; the source-backed sink loses
the same allocations and −8.22% of its instructions. Change 0592's observation that `write_text_to` cost
"about 101.9 M instructions and about 90,000 allocations more than `text()`" now
reads **84.2 M and 9,994**. The cycles fall further than the instructions
(−15.6% against −9.7%) because what left was 80,006 allocate-and-copy round trips,
not arithmetic. **What remains on this path is one allocation per paragraph** —
the bounded `String` the parser accumulates, 10,045 at 10,000 paragraphs against
`text()`'s 51 — and it is the next reduction here, though it changes what the
parser retains between paragraphs and moves the `try_reserve` sites that report
`Error::Allocation`. **The same shape survives elsewhere in the crate and is
unmeasured**: `read_event_into` with a caller buffer is still how
`header_footer/codec.rs`, `modern_comments/codec.rs`, `validation.rs`,
`smartart.rs`, `font/codec.rs`, `chart/codec.rs` and
`source_backed/tail_append.rs` read, each over a slice; none of them was priced
here, and only the last is on a save path. **A hotspot is closed by measurement
rather than by code**:
change 0592's `docx_file_eager_paragraph_count` regression is not a hotspot at
all — see `GOAL_AUDIT.md`.

## For `GOAL_AUDIT.md`

## 0643 — survey item XML-6 closed by code, change 0592's open question closed by attribution

Change 0643 closes item **XML-6** of the 0587 survey and the one question change
0592 left open. XML-6 is closed by code: the sink emission loop borrows, and the
counts are in `HOTSPOTS.md`. The 0592 question is closed by **evidence**, and the
answer is that the regression is not work.

Change 0592 reported `docx_file_eager_paragraph_count` at **+5.49%** pooled
against a 0.78% floor and wrote, in as many words, "this record does not know what
it is". It reproduces on the current base at **+9.81%** with a **−0.24%** floor,
both orders agreeing, so the first thing this change establishes is that 0592 was
right to report it and right not to bury it in a mean. The second is what it is.
The timed region — `document()` followed by `paragraph_count()` — executes
**0.53% fewer instructions** on the after leg and takes the same page faults, so
no work was added. The per-symbol count says where the difference lives:
`scan_word_element_ranges::<ParagraphIndex::from_xml::{closure#0}>` runs **twice**
per harness sample on the before leg — once inside `PreparedDocx::eager`'s untimed
`document()?.text()?` preparation, once inside the timed call — and **once** on
the after leg, inside the timed call. The before leg pays 1,454,305 extra
instructions before the clock starts and gets a warm scan for them.

A layout control in change 0623's shape — a third binary that links the `OnceLock`
field and the `get_or_init` initializer exactly as the base build does but fills
the cell in `from_part`, so the lazy path is never taken — splits the 6.63% the
probe reproduces into **+0.85%** of code layout, constant across all three
preparations, and **+5.74%** of scan placement. Make the preparation symmetric by
adding one paragraph query to it and the placement term collapses to **+0.27%**;
remove the preparation entirely and it is **+0.36%**. It scales with what happens
between the preparation and the clock: +2.40% on an in-memory package, +5.74% on
the same operation over a 16.79 MB archive, +9.81% in the harness's own child.

**It is not fixed, and the record says why.** The only value-identical removal is
restoring the eager scan, which costs −39.04% on the same corpus's eager text read
and −30.24% on its source-backed one in the same window — the trade change 0592
made deliberately. The non-identical option, memoizing the index on the `Package`
so a second `document()` reuses the first view's, would help exactly this caller
shape and is **not proposed**: it extends a cache lifetime across view boundaries
and would move a `DocumentIndexAdmission` charge on the budget-managed
source-backed route, which is the contract line 0592 already declined to cross. It
needs a frozen design record. The optimization order in `docs/GOAL.md` is
respected throughout: XML-6 removes unnecessary copying and allocation, ahead of
layout, algorithms and parallelism; nothing was vectorized and no parallelism was
introduced. **The audit gap this change does not close** is that
`tools/perf-baseline` still has no DOCX text-sink selector, so the path XML-6
improves is covered by a retained probe and a 333-document differential rather
than by a harness case.

## For `REPORT.md`

## 0643 — a DOCX text export that stops copying, and a mystery from 0592 solved

`litchi-docx`: streaming a document's text into a caller's sink used to copy every
piece of XML twice on the way past — once into a scratch buffer, once into a fresh
heap allocation — even though the document was already sitting in memory as a plain
slice of bytes. It no longer does. On a 10,000-paragraph document the export makes
**10,045 heap allocations instead of 90,051**, moves **1.2 MB fewer bytes through
the allocator**, and finishes about **16% sooner**; on a real 200-paragraph file
with 16 MB of images it finishes about 12% sooner. The text it writes is identical,
byte for byte, and so is every refusal: 333 documents were exported through five
different sink policies on both the before and the after build — including 129
malformed-document refusals, 142 size-limit refusals and 74 sink failures, each
with the exact amount of output it had already produced — and the two signature
files match exactly.

The same change settles something an earlier one could not. Change 0592 made the
paragraph index lazy, which made reading a document's text about 40% faster, but it
also made one benchmark — counting paragraphs — about 5% *slower*, and that record
honestly said it did not know why. It is now known, and it is a measurement
artifact rather than a cost. The benchmark warms up by extracting the document's
text first, and on the old build that warm-up happened to build and throw away the
very index the timed step then rebuilt, so the timed step was running warm code;
on the new build the warm-up does not touch it, so the timed step runs it cold. The
timed step executes *fewer* instructions on the new build. Give the warm-up one
paragraph query, so both builds enter the stopwatch equally warm, and the gap
essentially disappears. It was not "fixed", because the only honest fix is to put
back the work that made text reading 40% faster.

## For `ADR_COMPLIANCE.md`

## 0643 — compliant; one contract question raised and answered by declining

Change 0643 (`litchi-docx`: the borrowing sink text parser, plus an attribution
that changed no source) is compliant and adds no ADR question. **ADR 0003
(snapshots and borrowed views):** "borrowed zero-copy inputs produce scoped views,
and conversion to owned storage is explicit" is the shape this restores — the
events now borrow the slice the caller already owns for the whole call (the part's
cached visible XML on the eager facade, the `Cow` materialized before the loop on
the source-backed one), and the only owned storage left in the emission loop is
the bounded paragraph `String` the parser must accumulate. **ADR 0005 (bounded
resources):** retained memory moves strictly downwards — one heap allocation per
event and one growable buffer are gone, nothing new is retained, and no limit is
relaxed. `MAX_SEMANTIC_TEXT_RAW_XML_BYTES`, the processed-XML ceiling,
`MAX_SEMANTIC_TEXT_EVENTS`, `MAX_SEMANTIC_TEXT_DEPTH`, the name, attribute,
namespace-binding, namespace-byte, reference, run, paragraph-count,
paragraph-byte and decoded-document-byte ceilings and every `try_reserve` are
untouched, and each trips on the same event because `semantic_event_bytes`
measures the same payload borrowed or copied. No new `unsafe`, no new dependency.
**ADR 0006 (validation and preservation):** no validation moved, none was added
and none was removed; `preflight_semantic_xml` still walks the whole part before
emission begins, and the source-backed route's `check_execution()` and
`source_version()` fences still run before the payload, before and after every
read, before each paragraph emission and around every sink write. Because the
events borrow an owned in-memory buffer rather than the source, no borrow outlives
a fence it did not outlive before. Error identity is proved two ways: two new
tests pin end of input in both of its shapes on both facades — a tail ending on a
token boundary and one ending inside a start tag, each with zero progress and an
empty sink — **and both tests pass on the unmodified base as well**, and the
333-document differential reports identical refusals, identical partial output and
identical progress. This is a read path; no package byte is written, so `Preserve`
is untouched. **Contract movement: none — and one proposal to move one was
declined.** Memoizing the DOCX paragraph index across `document()` calls on one
`Package` would remove the residual measured in part (a), but it extends a cache
lifetime across view boundaries and would move a `DocumentIndexAdmission` charge
on the budget-managed source-backed route; it is left to a frozen design record,
for the same reason change 0592 left the managed branch alone.

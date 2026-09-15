# Log sections for change 0592

The coordinator merges these into the four logs; this batch did not edit them.
Each is written in the style of that file's newest section.

## For `docs/performance/HOTSPOTS.md`

### 0592 — DOCX-2 implemented: the paragraph index is built by the first paragraph query

Rank 10 of the 0587 queue (DOCX-2, step 1, "build the paragraph index lazily")
is implemented and can be struck from the table.
`DocumentPart::from_part` and the unmanaged branch of the source-backed
`Package::document()` no longer run `scan_word_element_ranges` to build
`ParagraphIndex`; a `OnceLock` fills on the first `paragraph_count`,
`paragraphs`, `paragraph(i)` or `paragraph_text(i)`. The MCE visibility pass
stays. **Measured** by callgrind isolation pairs on a 10,000-paragraph document:
`document()` alone −69.30% Ir (64.96 M → 19.94 M), `document()` + `text()`
−34.88%, + `write_text_to` −19.47%, + `tables()` −40.99%; the removed scan is
45.01 M Ir and **70.5% of the text-extraction pass itself**, which reproduces the
survey's "about 75%" on a different fixture. `document()` allocates 342,735 fewer
bytes and 25 fewer blocks. Paragraph queries are unchanged in allocation and
0.6-0.7% cheaper in instructions: the scan is moved, not removed, for them.
Warm paired timing, order `A1 B1 B2 A2` on CPU 12: `docx_file_eager_full_text`
**−40.22%** at p50 against a 1.52% A/A floor and `docx_file_source_full_text`
**−32.37%** against 0.66%. Two regressions are reported rather than averaged
away: `docx_semantic_one_paragraph` and `docx_semantic_list_paragraphs` move by
×2.2 × 10⁴ and +1,836% because those selectors build the view before starting
the clock, so the scan they used to do untimed now lands inside the interval;
and `docx_file_eager_paragraph_count` regresses **+5.49%** pooled (floor 0.78%,
both orders agreeing) although the same operation pair is 0.69% cheaper in native
cycles and takes the same page faults — that one is recorded as an open
question, not explained. The budget-managed
source-backed route is deliberately unchanged, because its index carries an
execution-budget charge; deferring that is a contract change needing its own
record. `performance_claim: none`. [Change and limitations](0592-docx-lazy-paragraph-index.md);
[evidence](results/change-0592/README.md).

**Two entries for the queue, found while measuring, not acted on.** DOCX-3's
second pass is now bounded from above: after 0592, `write_text_to` costs 101.9 M
Ir and about 90,000 allocations more than `text()` on the same 10,000-paragraph
document — nine allocations per paragraph — but the difference contains the
streaming parse and the sink as well as `preflight_semantic_xml`, so it is a gap,
not an attribution. And the harness has no selector that times a DOCX
`document()`: three of the four `docx_semantic_*` read selectors construct the
view before starting the clock, which is why DOCX-2's gain is invisible to them
and why this batch needed a scratch probe.

## For `docs/performance/GOAL_AUDIT.md`

### 0592: the first item off the 0587 queue, and what the harness could not see

`docs/GOAL.md`'s optimization order puts "eliminate unnecessary work" first, and
[0592](0592-docx-lazy-paragraph-index.md) is that step applied to the smallest
aligned item the 0587 survey found: a cache built on every DOCX main-document
open for the benefit of three of its dozen readers. It is retained on
deterministic evidence — callgrind isolation pairs and allocation counts on two
document shapes, plus a 332-document differential whose signature files are
byte-identical between legs — and `performance_claim: none`.

**The batch confirms one of the survey's falsification conditions rather than
refuting it.** 0587 wrote that DOCX-2 would be "falsified if the full-text
harness timer starts after `document()`, in which case the gain is invisible to
the selector". It does, and it is: `docx_semantic_full_text` moves −1.83% and
+0.10% across the two orders against an A/A floor of −0.80%. The gain is real and
is measured by the file-backed selectors, whose timed region does contain
`document()` (−40.22% at p50 on `docx_file_eager_full_text`, A/A floor 1.52%),
and by the instruction counts. The lesson for the goal's evidence rules is the one 0587
already recorded in another form: **a selector's timer boundary is part of the
claim's scope**, and three of the four `docx_semantic_*` DOCX read selectors
exclude the open from the interval they time. Anything that changes `document()`
must be measured with a probe or with the file-backed lifecycle selectors, and
this record says so rather than quoting the semantic selector's flat result as
"no effect".

**The A/A floor in this window was not the host's usual one.** With eight agents
building and measuring concurrently, the same-binary floor at p50 reached
**17.34%** on `docx_semantic_open` and **8.54%** on
`docx_file_eager_open_full_text_lifecycle`, so neither of those two carries a
result here; sub-microsecond selectors (`docx_semantic_one_paragraph` times a
130 ns array lookup) are below the clock's usable resolution entirely. The
counts, not the timings, carry this record's result, and the record says which is
which.

## For `docs/performance/REPORT.md`

### 0592: one cache, built when it is used

The first implementation off the 0587 queue, and the smallest: `litchi-docx`
built a paragraph-offset index on every `document()`, and the three readers of
the eager view that use it — four on the source-backed view — are outnumbered by
the eight that do not. [0592](0592-docx-lazy-paragraph-index.md)
moves it into a `OnceLock` filled by the first paragraph query, leaving the MCE
visibility pass exactly where it was. On a 10,000-paragraph document the open of
a main-document view drops **69.30%** of its instructions and **99.26%** of its
allocated bytes; a full text extraction drops **34.88%**; a streaming text export
drops **19.47%**. Paragraph readers pay the same scan, 0.6-0.7% later and
0.6-0.7% cheaper. `performance_claim: none`.

Two things about it are worth more than the percentages. The first is that the
index never carried a numbered record at all — it arrived in a checkpoint commit
— and so no record had ever priced it on a read that does not use it; the survey
that found it was the first thing in the program to look. The second is that the
in-process harness cannot see the result: three of the four `docx_semantic_*`
DOCX read selectors build the document before starting the clock, so the saving
lands entirely in their untimed setup and two of them report a large *regression*
instead, because the scan they used to do untimed now happens inside the
interval. The record reports those regressions as measured, explains the
boundary, and rests its result on the file-backed selectors and on the
instruction counts. The differential that backs the safety argument is 332
documents — every DOCX fixture in `test-data/` and every DOCX the 0483 and 0495
fuzz campaigns retained — signed across ten query families on both facades, with
zero differences, including six malformed documents whose every structural query
refuses identically on both legs.

## For `docs/performance/ADR_COMPLIANCE.md`

### 0592: aligned, and the one line where it stops

[0592](0592-docx-lazy-paragraph-index.md) is the first of the ten items 0587's
matrix called **aligned**, and it is aligned for the reason the matrix gave: it
changes no byte of output and moves no refusal. ADR 0005 is not merely permissive
here, it is prescriptive — "semantic payloads load lazily into thread-safe
weighted caches […] cache behavior is semantically invisible" — and the eager
build was the part that did not match. The index's error was already swallowed
with `.ok()`, so deferring the same expression to the first consumer leaves the
same operations able to observe the same failure; the pre-existing malformed-XML
tests on both facades pin that and were not touched. No exception is requested.

**The compliance line this record found is a budget one, and it is the reason the
change stops short of the whole path.** On the budget-managed source-backed
route the index is not free of the ADR: `DocumentIndexAdmission` reserves memory
and objects for it and is retained for the view's lifetime, and the query-parser
admission reserves memory, objects and depth and *consumes* `Work` for the scan.
ADR 0005 requires every operation to charge its budget, so deferring the scan
while keeping those reservations would run it with fewer live reservations than
the route established, and re-admitting at the deferred site would charge `Work`
twice and could refuse a document that is accepted today. Either is a contract
change. That branch therefore keeps its eager scan, and the record says so in
*What was changed* rather than in a footnote. If it is ever wanted lazily, it
needs a frozen design record that states what the admission covers and when it is
reserved — which is the same shape of prerequisite 0587 attached to twenty-one of
its other items.

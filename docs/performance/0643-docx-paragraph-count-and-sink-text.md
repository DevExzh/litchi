# 0643: the DOCX sink text parser borrows its events, and change 0592's open regression is the harness's own preparation

Status: retained. `performance_claim: none` — the paired medians and the
deterministic counts below are reported as evidence and are not registered as
claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Two DOCX follow-ons, one of which changes code and one of which does not.

**(a)** Change [0592](0592-docx-lazy-paragraph-index.md) left one question open:
`docx_file_eager_paragraph_count` moved **+5.49%** pooled at p50 against a 0.78%
floor after the paragraph index became lazy, while the same operation pair
measured 0.69% *cheaper* in native cycles and 0.63% cheaper in instructions on a
direct probe. That record said, in as many words, "this record does not know what
it is". This one does: it reproduces at **+9.81%** on the current base, and it
decomposes into +0.85% of code layout — proved with a binary in which the lazy
mechanism is linked but never taken — and about +5.7% that exists **only** while
the harness's untimed preparation builds and discards a paragraph index the timed
call then rebuilds. No source changed for this part; the attribution is the
result.

**(b)** Survey item **XML-6** (DOCX-3 territory) of change
[0587](0587-remaining-opportunity-survey.md):
`write_text_to_with_operation_check` read events with
`read_event_into(&mut buffer)` and then called `event.into_owned()` although its
source is a byte slice, while the sibling `for_each_word_text_chunk` has borrowed
since change [0229](changes/0229-docx-text-binding-tracker.md). Removing the two
copies costs the sink read **88.8% of its allocations** and **9.7% of its
instructions** on a 10,000-paragraph document and **15.6%** of its median wall
time. Scope: `litchi-docx`.

## What was changed

One statement pair in `crates/litchi-docx/src/paragraph/codec/text.rs`. The
emission loop of `write_text_to_with_operation_check` replaces

```rust
let mut buffer = Vec::new();
…
buffer.clear();
let event = reader.read_event_into(&mut buffer)?;
…
let event = event.into_owned();
```

with `let event = reader.read_event()?;` and drops the buffer. Nothing else
moves: the two raw/processed byte ceilings, the preflight, every
`operation_check()` call site — the two around the preflight, the three in the
loop, and the one threaded into `consume` — `validate_semantic_attribute_names`,
the namespace resolution through `reader.resolver()`, `SemanticTextParser::consume`
and every limit inside `SemanticTextXmlBudget` run on the same events in the same
order. No public signature, no error type, no output byte, no limit, no `unsafe`,
no new dependency.

Both entry points into that parser get it: `Document::write_text_to` on the eager
facade (`document/package/model.rs`) and `Package::write_text_to` on the
source-backed one (`source_backed.rs`), which passes its own source-freshness and
execution checks in as the operation check. Neither call site changed.

Two tests were added to `crates/litchi-docx/tests/sequential_text.rs`. Part (a)
changed no source at all.

## Why it is sound

**The two readers are the same reader.** `NsReader::read_event_into` and
`NsReader::<&'i [u8]>::read_event` are both one-line wrappers over the same
`NsReader::read_event_impl`, which calls `self.pop()`, delegates to
`self.reader.read_event_impl(buf)` and then `self.process_event(…)`
(quick-xml 0.41.0, `src/reader/ns_reader.rs:64`, `:190`, `:518`). They differ only
in the `XmlSource<'i, B>` witness: `B = &mut Vec<u8>` copies each token into the
caller's buffer, `B = ()` hands back a slice of the input. The tokenizer, the
`ReaderState`, the namespace push/pop and the error stream are the same objects;
`into_owned()` then copied the token a second time, out of the buffer and into a
fresh allocation, and existed only so the next iteration could `buffer.clear()`.
`consume` takes `Event<'_>` by value with a free lifetime and never stores it, so
nothing needed the owned form.

**This path already read borrowed events.** `preflight_semantic_xml`, in the same
file and on the same bytes with the same `trim_text(false)` and
`check_end_names = true` configuration, has always used
`Reader::from_reader(xml_bytes)` and `read_event()`. Every document that reaches
the emission loop has therefore already been tokenized once by a borrowed reader
over the identical slice. This change does not introduce the borrowed form to the
path; it stops the second pass from using the other one.

**No limit is relaxed, and each trips on the same event.**
`semantic_event_bytes` measures `as_ref().len()` of the event payload, which is
the same byte count borrowed or copied, so `MAX_SEMANTIC_TEXT_EVENT_BYTES` is
reached on exactly the same event. `MAX_SEMANTIC_TEXT_EVENTS`,
`MAX_SEMANTIC_TEXT_DEPTH`, the name, attribute, namespace, reference, paragraph
and document ceilings and the per-append `try_reserve` are untouched and observe
the same sequence.

**Bounded resources move only downwards.** ADR 0005 asks for streaming reads that
retain one bounded object; this removes a per-event heap allocation and a
growable buffer and retains nothing new. The events borrow `xml_bytes`, a slice
the caller already owns for the whole call — on the eager facade the part's
cached visible XML, on the source-backed facade the `Cow` materialized before the
loop. ADR 0003's "borrowed zero-copy inputs produce scoped views, and conversion
to owned storage is explicit" is the shape this restores: the only owned storage
left in the loop is the bounded paragraph `String` the parser must accumulate.

**No fence moves.** The source-backed route's `check_execution()` and
`source_version()` still run before the payload, before and after every read,
before each paragraph emission, and inside `SourceCheckedTextSink` around every
underlying write. Because the events borrow an owned in-memory `Cow` rather than
the source, no borrow outlives a fence it did not outlive before. ADR 0006's
`Preserve` default is untouched: this is a read path that writes no package byte.

**Error identity.** The two ways the refusal could have moved are end of input and
event chunking, and both are pinned by new tests and by the corpus differential
below. A tail that ends on a token boundary still reaches `Eof` and is refused by
the preflight's own balance check; a tail that ends inside a start tag is still
refused by the reader with `syntax error: tag not closed: \`>\` not found before
end of input`; both still report zero progress and write no byte, because the
preflight walks the whole part before emission begins.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, ext4 root; rustc
1.95.0, valgrind 3.26.0, quick-xml 0.41.0. Every timed process was pinned to
**CPU 18** with `taskset`; the callgrind runs used CPU 19. Seven other agents
built and measured on the other cores throughout. Legs, all built
`--release --locked`:

| leg | tree |
| --- | --- |
| **rev0592** | base `c7326f680` with change 0592's two source files reverse-applied |
| **layoutctl** | base with `DocumentPart::from_part` filling the `OnceLock` (the layout control) |
| **base** | the shared read-only checkout of `c7326f680` |
| **after** | this branch |

### (a) The regression reproduces, and it is larger than 0592 saw

`tools/perf-baseline`, `--filesystem-cache warm`, 60 samples and 3 warmups per
leg per selector, one fresh child process per sample against a real 16.79 MB file
on ext4 with a 29,027-byte 200-paragraph main document, order `A1 B1 B2 A2` with
A = **rev0592** and B = **base**. `A2` against `A1` is the floor.

| selector | A1 | B1 | B2 | A2 | A/A | pooled B−A |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_file_eager_paragraph_count` | 127,591 | 140,021 | 139,856 | 127,291 | **−0.24%** | **+9.81%** |
| `docx_file_eager_full_text` | 211,276 | 129,925 | 129,605 | 214,471 | +1.51% | −39.04% |
| `docx_file_source_full_text` | 304,681 | 216,051 | 210,061 | 306,187 | +0.49% | −30.24% |

p50 nanoseconds; means, p95 and p99 are in
[`timing/repro-summary.txt`](results/change-0643/timing/repro-summary.txt). The
window was quiet — a −0.24% floor against 0592's 0.78% — and the regression is
**+9.81%**, both orders agreeing (+9.74%, +9.87%). The two full-text selectors
reproduce 0592's −40.22% and −32.37% within 2.2 points, so this is the same
effect on the same corpus, measured better.

### (a) The timed region executes *less* work on the after leg

`probe0643` (retained) reproduces the selector's timed region exactly:
`PreparedDocx::eager`'s preparation — `detect_format_smart_with_limits`,
`Package::from_opc_package`, then one untimed `document()?.text()?` — followed by
the `document()` + `paragraph_count()` pair the harness clocks. Isolation pairs at
**reps = 0 and reps = 1** price that one call *including* what it pays for being
the first in the process, which is exactly the term 0592's reps = 4 / reps = 20
pairs differenced away.

| operation | rev0592 | layoutctl | base | base − rev0592 |
| --- | ---: | ---: | ---: | ---: |
| one timed call, harness-shaped preparation | 2,173,875 | 2,171,314 | 2,162,311 | **−0.53%** |
| one timed call, no preparation | 2,172,370 | 2,171,096 | 2,161,084 | **−0.52%** |

Instructions, callgrind, deterministic. The call the selector reports as 9.81%
slower runs **0.53% fewer instructions**, reproducing 0592's −0.61% to −0.68%.
Page faults attributable to that call are 0 or ±1 on every operation and every
leg over 40 processes each
([`counts/first-call-perf.csv`](results/change-0643/counts/first-call-perf.csv)),
reproducing 0592's fault finding. Native cycles could not be scoped to the call
at all and are reported as **inconclusive**: the process median is 78–81 M cycles
with an interquartile range of 3.2–4.8 M against a signal near 0.4–1.0 M.

### (a) The per-symbol count names the mechanism

Inclusive Ir for the paragraph scan in the same two profiles:

| leg | reps = 0 (preparation only) | reps = 1 (preparation + the timed call) |
| --- | ---: | ---: |
| rev0592 | 48,910 | 97,820 |
| layoutctl | 48,910 | 97,820 |
| base | **symbol absent** | 42,635 |

`scan_word_element_ranges::<ParagraphIndex::from_xml::{closure#0}>` runs **twice**
per harness sample on the before leg — once in the untimed preparation, once
inside the timed call — and **once** on the after leg, inside the timed call. The
whole-process totals agree: the before leg's preparation costs 85,758,265 Ir
against the after leg's 84,303,960, a difference of 1,454,305 Ir, which is the
scan the preparation no longer runs. The timed call on the before leg is the
scan's *second* execution in the process; on the after leg it is its *first*.

### (a) The layout control separates the two terms

0623's method: **layoutctl** is the base build with the `OnceLock` field and the
`get_or_init` initializer linked exactly as on base, but with `from_part` filling
the cell, so the lazy path is never taken and the scan runs where it ran before
0592. `C − A` is therefore the field and the code layout alone; `B − C` is the
placement of the scan alone. Order `A1 C1 B1 B2 C2 A2`, 60 timed first calls per
leg (40 on the synthetic shape), one process per call.

Real 16.79 MB harness corpus file:

| untimed preparation | A/A | C/C | B/B | **C−A** (layout) | **B−C** (placement) | B−A |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `document().text()` — the harness's own | −0.29% | +0.20% | −0.17% | **+0.85%** | **+5.74%** | +6.63% |
| `text()` **and** `paragraph_count()` | +0.38% | +0.30% | +0.52% | +0.79% | **+0.27%** | +1.06% |
| none | −0.60% | −1.12% | −1.04% | +0.93% | +0.36% | +1.29% |

Synthetic in-memory 200-paragraph package (change 0592's fixture):

| untimed preparation | A/A | C/C | B/B | C−A | B−C | B−A |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `document().text()` | +0.13% | −0.14% | −0.16% | +0.22% | **+2.40%** | +2.63% |
| `text()` and `paragraph_count()` | −0.26% | −0.10% | −0.05% | +0.06% | +0.22% | +0.28% |
| none | +0.02% | −0.07% | +0.50% | +0.72% | −0.33% | +0.38% |

These three `B−A` figures were also taken independently, without the control
binary interleaved, before the control was built:
[`timing/firstcall-summary.txt`](results/change-0643/timing/firstcall-summary.txt)
reports +6.64%, +1.05% and +1.30% against floors of −0.11%, −0.22% and −0.02%,
agreeing with the table above to within 0.02 points on all three.

**What this says.** The layout term is a **constant +0.79% to +0.93%** on the real
file, present under all three preparations and therefore independent of them: it
is the `OnceLock` field and the codegen it produces, isolated by a binary that
links the lazy mechanism and never takes it. The placement term is **+5.74%**,
and it appears *only* when the preparation builds and discards an index — add one
paragraph query to the preparation, so that both legs enter the timed call with
the scan already executed once, and it collapses to **+0.27%**. It scales with
what happens between the preparation and the call: +2.40% on an in-memory package,
+5.74% on the same operation reading a 16.79 MB archive, and +9.81% in the
harness's own child, which additionally reads the file and snapshots process
metrics between the two.

The regression is therefore **not work**: the timed call runs 0.53% fewer
instructions and takes the same page faults. It is the loss of a warm-up the
harness's preparation used to supply for free, plus a small constant layout cost.
It is **not fixed here**, because the only value-identical fix is to restore the
eager scan, which is what 0592 removed for −39.04% on the same corpus's text
reads. A fix that is *not* value-identical — memoizing the index across
`document()` calls on one `Package` — is named under *Limitations*.

### (b) Allocations

Isolation pairs at reps = 4 and reps = 20 through the probe's counting global
allocator (`realloc` counted as one allocation of its new size). Deterministic;
identical across repeats.

| operation | shape | base allocs | after allocs | Δ | base bytes | after bytes | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `write_text_to` (eager) | 24 | 267 | 69 | **−74.16%** | 7,781 | 4,755 | −38.89% |
| | 200 | 1,851 | 245 | **−86.76%** | 37,349 | 13,555 | −63.71% |
| | 10,000 | 90,051 | 10,045 | **−88.84%** | 1,683,749 | 503,555 | −70.09% |
| `write_text_to` (source) | 24 | 268 | 70 | −73.88% | 7,877 | 4,851 | −38.42% |
| | 200 | 1,852 | 246 | −86.72% | 37,445 | 13,651 | −63.54% |
| | 10,000 | 90,052 | 10,046 | **−88.84%** | 1,683,845 | 503,651 | −70.09% |
| `document()` (control) | all | 26 | 26 | 0 | 2,558 | 2,558 | 0 |
| `text()` (control) | 24/200/10,000 | 42/45/51 | 42/45/51 | 0 | identical | identical | 0 |
| `extract_text()` (control) | 24/200/10,000 | 42/45/51 | 42/45/51 | 0 | identical | identical | 0 |

Change 0592 observed that on its 10,000-paragraph document `write_text_to` cost
about **90,000 allocations more** than `text()`. It now costs **9,994 more**
(10,045 against 51): 88.9% of that gap is the removed per-event `into_owned()`,
one allocation per event. What remains is one allocation per paragraph — the
bounded paragraph `String` the parser accumulates — and is named under
*Limitations*.

### (b) Instructions

Same isolation pairs under `valgrind --tool=callgrind --cache-sim=no
--branch-sim=no`, reps 4 → 20 at 24 and 200 paragraphs and 1 → 5 at 10,000.

| operation | shape | base | after | Δ |
| --- | ---: | ---: | ---: | ---: |
| `write_text_to` (eager) | 24 | 492,538 | 447,706 | **−9.10%** |
| | 200 | 3,762,724 | 3,403,197 | **−9.56%** |
| | 10,000 | 185,883,829 | 167,912,905 | **−9.67%** |
| `write_text_to` (source) | 24 | 562,001 | 518,205 | −7.79% |
| | 200 | 4,321,281 | 3,967,527 | −8.19% |
| | 10,000 | 213,526,942 | 195,977,858 | **−8.22%** |
| `document()` (control) | 24/200/10,000 | 72,302/422,890/19,944,486 | 72,304/422,895/19,944,655 | +0.00% |
| `text()` (control) | 24/200/10,000 | 236,793/1,709,027/83,647,782 | 236,736/1,709,349/83,666,762 | −0.02% to +0.02% |
| `extract_text()` (control) | 24/200/10,000 | 237,694/1,710,801/83,680,969 | 237,859/1,710,853/83,701,911 | +0.00% to +0.07% |

The base leg's 185,883,829 Ir for `document() + write_text_to` on the
10,000-paragraph fixture matches 0592's 185,842,549 to 0.02%, which is how these
two records are known to be measuring the same thing. 0592 priced the sink read at
**101.9 M Ir more** than `text()` on that document; it is now **84.2 M more**
(167,912,905 − 83,666,762), so 17.6% of the instruction gap closed against 88.9%
of the allocation gap — the copies were cheap in instructions and expensive in
allocator work, which is what the timings below confirm.

### (b) Paired timing

Order `A1 B1 B2 A2` with A = **base** and B = **after**, 60 timed calls per leg
per case in one process after one untimed preparation, pinned to CPU 18. `A2`
against `A1` is the floor. Every case that was run is listed.

| case | A/A | pooled B−A |
| --- | ---: | ---: |
| `write_text_to`, eager, 24 paragraphs | +0.51% | **−15.19%** |
| `write_text_to`, eager, 200 paragraphs | +0.57% | **−17.89%** |
| `write_text_to`, eager, 10,000 paragraphs | +0.08% | **−15.57%** |
| `write_text_to`, eager, the 16.79 MB harness corpus file | +2.06% | **−12.37%** |
| `write_text_to`, source-backed, 24 paragraphs | +3.85% | −14.91% |
| `write_text_to`, source-backed, 200 paragraphs | −0.71% | **−13.54%** |
| `write_text_to`, source-backed, 10,000 paragraphs | +3.08% | −15.81% |
| `text()` control, 24 paragraphs | +1.64% | −0.56% |
| `text()` control, 200 paragraphs | +0.69% | +1.13% |
| `text()` control, 10,000 paragraphs | +1.02% | +0.86% |
| `text()` control, harness corpus file | −1.24% | +2.14% |

p50 nanoseconds; the per-leg means, p95, p99 and minima are in
[`timing/sink-summary.txt`](results/change-0643/timing/sink-summary.txt). The
four `text()` rows are the control — that projection is not on the changed path —
and they move −0.56% to +2.14% against floors of 0.69% to 1.64%, i.e. nothing.
The sink rows fall **12.37% to 17.89%**, four to nineteen times their own floors.
Two source-backed rows carry floors above 3% and are reported as such; their
counts, not their timings, carry the result there.

**The cycles fall further than the instructions** (−15.6% against −9.7% at 10,000
paragraphs) because what was removed is 80,006 allocate-and-copy round trips, not
80,006 instructions' worth of arithmetic. The harness has no DOCX text-sink
selector, so no `tools/perf-baseline` case covers this path; the three DOCX file
selectors that exist were run above as part (a)'s controls and none of them
reaches `write_text_to`.

## Correctness evidence

### Differential corpus: 333 documents, both facades, byte-identical

`probe/src/corpus.rs` signs every main-document read on both facades — open
outcome, `text()`/`extract_text()`, `paragraph_count()`, `paragraphs()`, every
`paragraph(i)` to one past the end, `paragraph_text(i)`, `tables()`, `elements()`,
`blocks()` — and adds what this change touches: `write_text_to` under **five**
sink configurations, signing the exact bytes the sink accepted, the number of
write calls, and either the returned `TextOutputReport` or the complete
`TextOutputError` including its retained partial progress. The five are the
default policy; a 64-byte output ceiling; a one-object ceiling; a `|` separator
with empty objects excluded; and a sink that refuses its third write.

| corpus | documents | result |
| --- | ---: | --- |
| every `.docx` under `test-data/` | 62 | signature files **byte-identical**, sha256 `b8651f42a2ca5f7493ad03593901e54ae4e0f11a7906f7983001b985ba4eebe9` on both legs |
| every `.docx` retained under `docs/performance/results/` by the change-0483 and change-0495 fuzz campaigns | 270 | **byte-identical**, sha256 `53fe11c85c4d5e2c35cf51bbc26593d49e1f5f772332cf84409a172f48cdcc63` |
| the harness's own `docx_file_eager_paragraph_count` corpus file, retained | 1 | **byte-identical**, sha256 `bdf0c7a370fa7ad1ae71db6688493ded33348b498814a4dadf8085b32e655207` |

The refusals matter more than the successes, and there are many. Across the 333
documents and the five configurations the signatures contain **129
`TextOutputError::Document`** refusals — 12 of them `Error::Xml`, the
quick-xml syntax errors that the reader change could have moved, and 15
`Error::InvalidFormat` — **142 `TextOutputError::Limit`** refusals, and **74
`TextOutputError::Sink`** failures, each with its exact retained progress. All
identical on both legs. Eight `test-data/` fixtures are refused at open on both
legs with identical text.

### Tests

Two tests were added to `crates/litchi-docx/tests/sequential_text.rs`, each
pricing one property the borrowed reader could break. **Both pass on the
unmodified base as well**, which is the point: they pin behaviour the change does
not move, not behaviour it introduces.

- `every_accepted_event_kind_projects_identically_through_both_sink_facades`
  drives a document whose text runs are split across several events — a comment
  inside a `w:t`, a CDATA section inside a `w:t`, an entity and a character
  reference — plus a leading comment, a declaration, an empty control element and
  a second paragraph, and requires the eager and source-backed sinks to produce
  the same exact bytes and the same report.
- `truncated_documents_refuse_at_end_of_input_with_unchanged_message_and_progress`
  pins end of input, where a borrowed and a buffered reader are most likely to
  diverge, in both of its shapes: a tail ending on a token boundary
  (`semantic DOCX XML has unbalanced elements`) and one ending inside a start tag
  (`syntax error: tag not closed: \`>\` not found before end of input`), on both
  facades, each with zero progress and an empty sink.

Thirteen pre-existing tests in the same file carry the rest of the contract and
were not touched, including `eager_and_source_paragraph_sinks_have_parity`,
`malformed_later_paragraph_preserves_prior_progress` (the partial-progress path
through the emission loop), `output_and_object_limits_report_exact_progress`,
`oversized_namespace_event_is_rejected_before_resolution`,
`parser_depth_decoded_and_illegal_character_limits_are_bounded` and the four
source-freshness and cancellation tests.

### Gates

All in the worktree; tails in [`gates.txt`](results/change-0643/gates.txt).
`cargo fmt --all --check` clean; `cargo clippy -p litchi-docx --all-targets`
clean under the workspace's deny lints; `cargo doc -p litchi-docx --no-deps`
clean under the deny rustdoc lints; `cargo test -p litchi-docx` **1,467 passed,
0 failed, 31 ignored** across 52 test binaries. The two gaps the first wave found
were both run: the harness's own suite (`cargo test` in `tools/perf-baseline`,
**531 passed, 0 failed, 1 ignored** across 19 test binaries) and the
feature-bearing `cargo test -p litchi --features docx,xlsx,pptx,xls`
(**266 passed, 0 failed, 7 ignored** across 26 test binaries).
`python3 tools/non_iwork_gate.py check`, `python3 tools/check_crate_boundaries.py`
and `python3 tools/check_report_claim_classification.py` all exit 0.

One gate fails and it is **pre-existing and unrelated**:
`python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural` exits 2 with "landed claim
`claim-0251-xlsx-xml-borrowed` requires strict evidence verification". It was
reproduced verbatim on the untouched before checkout at `c7326f680`. This change
registers no claim, does not touch `claim-registry-v1.json`, and modifies no
`litchi-xlsx` file. Change 0592 recorded the same failure.

## Validation preserved

Nothing in this change is a validation, and no validation mutates anything. The
preflight (`preflight_semantic_xml`) runs over the whole part before emission
begins, exactly where it ran. The raw and processed XML byte ceilings, the event,
depth, name, attribute, namespace-binding, namespace-byte, reference, run,
paragraph-count, paragraph-byte and decoded-document-byte limits, the attribute
name validation, the element namespace validation, the XML character and comment
validation, and the `try_reserve` on every text append are all unchanged and
observe the same event sequence. On the source-backed facade the
`check_execution()` and `source_version()` fences still run before payload
materialization, before and after every read, before each paragraph emission and
around every sink write, and the declared-size check still precedes payload
materialization. The 270 fuzz-corpus documents and the 142 limit refusals in the
differential are the check that what refused still refuses, with the same message
and the same partial output.

## Limitations

**No claim is registered.** `performance_claim: none`. The counts are exact for
these fixtures on this host and build; the timings are warm, in-memory or
warm-page-cache, single-host, taken while other agents worked on the other cores.
No cold-cache, physical-device, peak-RSS, concurrency-scaling or cross-platform
result is taken or claimed.

**One allocation per paragraph remains on the sink path,** and it is not removed
here. `SemanticTextParser` allocates a fresh `String` per `w:p`
(10,045 allocations at 10,000 paragraphs, ≈1 per paragraph, against `text()`'s
51). Reusing one cleared buffer would remove it, but it changes what the parser
retains between paragraphs from "this paragraph" to "the largest paragraph so
far" and moves the `try_reserve` sites that report `Error::Allocation`, so it is
a separate change with its own bounded-resource argument to make.

**Part (a) fixes nothing, by design.** The +9.81% on
`docx_file_eager_paragraph_count` is attributed, not removed. The only
value-identical way to remove it is to restore the eager scan, which costs
−39.04% on the same corpus's eager text read and −30.24% on its source-backed one
in the same window. The non-identical option — memoizing the paragraph index on
the `Package` so a second `document()` reuses the first view's index — would help
exactly this caller shape and is **not** proposed here: it extends a cache
lifetime across view boundaries, and on the budget-managed source-backed route it
would move a `DocumentIndexAdmission` charge, which is the contract line change
0592 already declined to cross. It needs a frozen design record.

**The +0.85% layout term is measured, not explained.** The control proves it is
the `OnceLock` field and the code it generates rather than the placement of the
scan, and it is consistent across three preparations and two fixtures. This
record does not attribute it to a particular instruction, and it is below the
review trigger.

**The probe reproduces two thirds of the selector's regression, not all of it.**
+6.63% against +9.81%. The residual is what the harness child does that the probe
does not — reading the file, taking `process_metrics` snapshots and opening an
allocation region between the preparation and the clock — and the direction is
consistent with the mechanism, but this record measures it only by difference.

**Native cycles for the single timed call are inconclusive**, for the reason 0592
found: the enclosing process is 78–81 M cycles with an interquartile range of
3.2–4.8 M, and the call is under 1 M. The instruction counts and the page-fault
counts carry that part.

**Seven other `read_event_into` sites in the crate are untouched and unmeasured.**
`header_footer/codec.rs`, `modern_comments/codec.rs`, `validation.rs`,
`smartart.rs`, `font/codec.rs`, `chart/codec.rs` and
`source_backed/tail_append.rs` all read a slice through a caller buffer. None was
priced here; this change claims nothing about them, and `tail_append.rs` in
particular is on a save path where the argument above about borrows and fences
would have to be made again.

**The `text()` and `extract_text()` projections are untouched.** They use
`extract_word_text`/`for_each_word_text_chunk`, which have borrowed since change
0229; the controls in every table above confirm they did not move.

## Retained evidence

[`results/change-0643/README.md`](results/change-0643/README.md) — the probe and
its corpus checker, the two control patches, all nine capture scripts, the
allocation and instruction counts, the per-symbol callgrind extract, the raw
perf counters, the four harness timing JSONs, the six corpus signature files, the
`decision.json`, `gates.txt` and `log-sections.md`.

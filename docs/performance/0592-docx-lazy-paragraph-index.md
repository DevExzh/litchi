# 0592: the DOCX paragraph index is built by the first paragraph query, not by every `document()`

Status: retained. `performance_claim: none` — the paired medians and the
deterministic counts below are reported as evidence and are not registered as
claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **DOCX-2** (rank 10) of change
[0587](0587-remaining-opportunity-survey.md), the one item that record's ADR
matrix lists as *aligned* — "changes no byte of output and moves no refusal" —
and needing only a paired measurement. Scope: `litchi-docx`.

## What was changed

Two files, one mechanism. `ParagraphIndex` is a bounded array of `(start,
length)` offsets over the visible main-document XML, built by one full
`scan_word_element_ranges` pass. It serves the selected-paragraph query that
change [0283](changes/0283-docx-selected-paragraph-resource-evidence.md)
measured, but — worth saying, because change 0587 attributed it to 0283 — **the
index itself carries no numbered record**: it entered the tree in commit
`831e80d22`, "checkpoint uncommitted agent edits from docx-para-index-20260821",
and no record since has priced it on a read that does not use it.

Three operations read the index — `paragraph_count`, `paragraphs`,
`paragraph(i)`, plus `paragraph_text(i)` on the source-backed view — and every
other main-document read ignores it. It was built on **every** `document()`.

**`crates/litchi-docx/src/parts/document_part.rs`.** `DocumentPart`'s
`paragraph_index` field changes from `Option<Arc<ParagraphIndex>>` to
`OnceLock<Option<Arc<ParagraphIndex>>>`. `from_part` no longer scans; it keeps
the MCE visibility pass (`visible_document_xml`) exactly where it was and stores
an empty cell. A new private `paragraph_ranges()` fills that cell with the
*identical* expression the constructor used —
`ParagraphIndex::from_xml(xml).ok().map(Arc::new)` — and the three consumers call
it. `extract_text`, `xml_bytes`, `table_count`, `tables`, `elements`, `blocks`
and the remaining readers do not, so they no longer pay for a scan they cannot
use.

**`crates/litchi-docx/src/source_backed.rs`.** `Document`'s field becomes the
same `OnceLock`, and the same `paragraph_ranges()` serves
`paragraph_count`, `paragraph_text`, `paragraphs` and `paragraph`. The
**unmanaged** branch of `Package::document()` stores an empty cell. The
**budget-managed** branch is deliberately unchanged: it still reserves
`DocumentIndexAdmission`, admits the query parser and scans eagerly, and stores
the result through `OnceLock::from`, so its cell is already full and the lazy
initializer never runs for it. The reason is in *Why it is sound*.

Nothing else moves. No public signature, no error type, no output byte, no
limit, no `unsafe`, no new dependency; `OnceLock` is `std`.

## Why it is sound

**ADR 0005 prescribes exactly this shape.** "Semantic payloads load lazily into
thread-safe weighted caches. Clean parsed values are evictable; active handles
pin them […] **Cache behavior is semantically invisible.**" The paragraph index
is a clean parsed value derived from already-validated visible XML, and its own
doc comment says so: "the cache is an optimization and never changes the accepted
document surface". Building it in a
thread-safe cell on first use is the lazy load the ADR asks for; building it on
every open was the part that was not.

**Error identity does not move, because the index never owned an error.** The
constructor already swallowed the scan's failure with `.ok()`, leaving `None`,
and each of the three consumers then falls through to the streaming scanner
(`document_paragraph_count`, `document_paragraphs`, `document_paragraph`), which
reports it. Deferring the same expression to the first consumer keeps the
swallowed failure observable on exactly the same operations: the consumers are
unchanged and the non-consumers could never see it. The pre-existing
`rejects_unterminated_selected_elements` (eager) and
`malformed_source_document_errors_at_first_semantic_read` (source-backed) pin
that timing and still pass unmodified.

**No bound is relaxed.** `MAX_PARAGRAPH_INDEX_RANGES`, the per-push
`try_reserve`, and the depth and node limits inside `scan_word_element_ranges`
are untouched, and they run on the same bytes.

**The managed branch is excluded on purpose, and that is the contract line.**
ADR 0005 also says "every operation charges a hierarchical resource budget
supplied by an execution context". On the budget-managed source-backed route the
index has an explicit charge: `DocumentIndexAdmission::new` reserves memory and
objects for it and is *retained* for the view's lifetime, and
`admit_document_query_parser` reserves memory, objects and depth and **consumes**
`Work` for the scan. Deferring the scan while keeping those reservations would
run it with fewer live reservations than the managed contract established;
re-admitting a parser at the deferred site would consume `Work` twice and could
refuse a document that is accepted today. Both are contract changes, so neither
is made here: the managed branch keeps its eager scan, byte for byte. The
unmanaged branch reserves nothing for the index today, so deferring it charges
nothing differently.

**No freshness fence moves.** The deferred scan reads bytes the view already
owns; it is not a source read, so no `source_version()` check is skipped. Every
paragraph entry point still runs `check_execution()` (or
`check_selective_operation`) before the cell is filled, as it did before.

**Thread safety.** `OnceLock<Option<Arc<ParagraphIndex>>>` is `Send + Sync`
whenever its payload is, and `ParagraphRange` is `Copy`; `Part` is declared
`PartClone + Send + Sync`, so `DocumentPart<'a>` and `source_backed::Document`
keep the bounds they had. Two concurrent first queries can both compute an
index; one is stored and the other dropped, and the two are equal because
`ParagraphIndex::from_xml` is a pure function of the pinned bytes. Two new
concurrency tests and the pre-existing
`source_document_is_send_sync_and_managed_arc_views_refuse_consistently` pin it.

**Clone.** `source_backed::Document` derives `Clone`. A clone taken before the
first paragraph query starts with an empty cell and builds its own index — one
extra scan in that case, never a different answer. That is the only way the
change can add work, and it is bounded by one scan per clone.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Every process in this record ran pinned to **CPU 12** with
`taskset` while seven other agents built and measured on the other cores. Both
legs were built `--release --locked` from the same lockfile: the before leg is
the detached checkout of `08d968f8e`, the after leg is this branch.

### Deterministic counts

The harness has no selector that isolates `document()`, so the counts come from
a scratch probe with path dependencies on `litchi-core`, `litchi-docx` and
`litchi-opc`, retained at
[`results/change-0592/probe/`](results/change-0592/probe/). It builds one DOCX in
memory whose `document.xml` holds *N* paragraphs carrying the perf-baseline
corpus's own payload string, then repeats one operation *R* times. Each figure is
an **isolation pair**: the totals at *R* = 1 and *R* = 5 (10,000 paragraphs) or
*R* = 4 and *R* = 20 (200 paragraphs) are differenced and divided, so process
start, fixture construction and the package open are removed. Instructions come
from `valgrind --tool=callgrind`; allocations come from a counting global
allocator in the probe (`realloc` counted as one allocation of its new size).

10,000 paragraphs — `document.xml` 830,048 bytes, archive 29,624 bytes:

| operation | before Ir | after Ir | delta |
| --- | ---: | ---: | ---: |
| `document()` alone | 64,956,377 | 19,944,796 | **−69.30%** |
| `document()` + `text()` | 128,804,274 | 83,881,547 | **−34.88%** |
| `document()` + `write_text_to(sink)` | 230,772,059 | 185,842,549 | **−19.47%** |
| `document()` + `tables()` | 111,432,776 | 65,757,019 | **−40.99%** |
| `document()` + `paragraph(i)` | 64,950,698 | 64,551,637 | −0.61% |
| `document()` + `paragraph(i)` × 8 | 64,971,933 | 64,554,253 | −0.64% |
| `document()` + `paragraph_count()` | 64,986,456 | 64,542,120 | −0.68% |
| source-backed `document()` alone | 64,972,438 | 19,945,459 | **−69.30%** |
| source-backed `document()` + `extract_text()` | 128,796,130 | 83,939,457 | **−34.83%** |
| source-backed `document()` + `paragraph(i)` | 65,016,103 | 64,518,733 | −0.76% |
| source-backed `document()` + `paragraph_count()` | 64,986,967 | 64,570,004 | −0.64% |

200 paragraphs — `document.xml` 16,648 bytes, archive 1,532 bytes — reproduces
every share within a point: `document()` **−68.23%** (1,330,959 → 422,911),
`+ text()` **−34.77%**, `+ write_text_to` **−19.50%**, `+ tables()` **−35.14%**,
and every paragraph query between −0.61% and −0.63%. The full table is in
[`counts/instructions.txt`](results/change-0592/counts/instructions.txt) with the
88 raw callgrind totals beside it.

**What the numbers say.** The removed scan is **45.01 M Ir** on a 830 KB main
part, 4,501 Ir per paragraph, and it is *the same 45 M* in every row that loses
it, which is what a removed fixed pass should look like. Differencing the two
text rows prices the text pass itself at 63.85 M Ir before and 63.94 M after —
unchanged, as it must be — so the eager index was **70.5% of a full text
extraction** on this document. That is the survey's "about 75%", independently
reproduced on a different fixture. The paragraph rows are the control: the scan
is not removed there, only moved, and they come out **0.6–0.8% cheaper** —
the constructor's removed `Option<Arc<_>>` build-and-move, or code layout; the
record does not attribute it, because it is not a saving worth naming.

Allocations, same isolation pairs, 10,000 paragraphs:

| operation | before | after |
| --- | ---: | ---: |
| `document()` alone | 51 allocations, 345,293 B | 26 allocations, **2,558 B** |
| `document()` + `text()` | 76, 1,984,218 B | 51, **1,641,483 B** |
| `document()` + `write_text_to` | 90,076, 2,026,484 B | 90,051, **1,683,749 B** |
| `document()` + `tables()` | 61, 345,868 B | 36, **3,133 B** |
| `document()` + any paragraph query | 51, 345,293 B | 51, 345,293 B |

The index is 25 allocations and 342,735 bytes on this document — thirteen `Vec`
growth steps plus the `Arc<[ParagraphRange]>` it is frozen into. A text-only read
now allocates **99.26% fewer bytes** at `document()` and 17.27% fewer over the
whole read; a paragraph read allocates exactly what it did. The source-backed
rows are identical to the eager ones (`allocations.txt`).

### Native cycles for the same operations

Instructions rank work, not latency (change 0579: 1.24% of instructions, 6.19%
of cycles), so the probe's four headline operations were also priced with
`perf stat -e cycles` on CPU 12: isolation pairs at *R* = 4 and *R* = 20 on the
200-paragraph document, 15 pairs per leg, in `before after after before` order,
medians. The A/A column is the two before orders against each other.

| operation | A/A floor | before → after, pooled |
| --- | ---: | ---: |
| `document()` alone | −0.01% | **−68.34%** |
| `document()` + `text()` | +1.60% | **−39.81%** |
| `document()` + `paragraph_count()` | −1.67% | −0.69% |
| `document()` + `paragraph(i)` | +1.66% | −1.39% |

The cycle figures agree with the instruction figures on `document()` (−68.34%
against −68.23%) and on the paragraph control (−0.69% against −0.63%), and the
text read saves rather more cycles than instructions (−39.81% against −34.77%),
which is what a removed allocate-and-grow loop looks like against a straight-line
scan.

### Paired timing

Order `A1 B1 B2 A2` (before, after, after, before), 30 samples and 3 warmups per
leg per selector, both legs pinned to CPU 12, the four runs of each family inside
one window. The A/A column is A2 against A1 — the same binary against itself in
that window — and is the floor this record's timings are read against. Every
selector that was run is listed, including the ones that got worse.

**File-backed selectors.** These run a fresh child process per sample against a
real file on ext4, warm page cache, over the harness's `docx-source-backed-media`
corpus: 20 archive members, 16.79 MB, a 200-paragraph main document. Their timed
region contains `document()`.

| selector | A1 | B1 | B2 | A2 | A/A | pooled B−A |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_file_eager_full_text` | 210,221 | 126,965 | 126,291 | 213,416 | +1.52% | **−40.22%** |
| `docx_file_source_full_text` | 302,612 | 204,281 | 206,371 | 304,616 | +0.66% | **−32.37%** |
| `docx_file_eager_paragraph_count` | 126,480 | 134,160 | 133,716 | 127,466 | +0.78% | **+5.49%** |
| `docx_file_eager_open_full_text_lifecycle` | 11,240,922 | 11,026,581 | 10,993,215 | 12,200,432 | +8.54% | −6.06% |

p50 nanoseconds; means, p95 and p99 per leg are in
[`timing/summary.txt`](results/change-0592/timing/summary.txt) and the raw
distributions in the eight JSONs beside it. The two full-text selectors agree
across both orders to within 1.3 points and sit between twenty-six and
forty-nine times their own A/A floor. The open-plus-text lifecycle is dominated by reading 16.79 MB from the
file and its floor is 8.54%, so its −6.06% is **inside the floor and is not a
result**; the same operation's instruction count is the statement that stands.

**In-process semantic selectors**, on the 10,000-paragraph corpus:

| selector | what its timer covers | A/A | pooled B−A |
| --- | --- | ---: | ---: |
| `docx_semantic_full_text` | `text()`; the view is built untimed | −0.80% | −0.87% |
| `docx_semantic_open` | `Package::from_reader`; no `document()` at all | −17.34% | +3.87% |
| `docx_semantic_list_paragraphs` | `paragraphs()`; the view is built untimed | −2.61% | **+1,836%** |
| `docx_semantic_one_paragraph` | `paragraph(i)`; the view is built untimed | −69.2% | **×2.2 × 10⁴** |

### Reading the two regressions

Both are reported, neither is hidden in a mean, and they are different in kind.

**The two semantic paragraph selectors are a timer-boundary artifact, and change
0587 predicted it.** That record's falsification condition for DOCX-2 was "the
full-text harness timer starts after `document()`, in which case the gain is
invisible to the selector"; `docx_semantic_full_text` shows exactly that, moving
−0.87% against a −0.80% floor. The same boundary works the other way for
`docx_semantic_one_paragraph` and `docx_semantic_list_paragraphs`: they build the
view untimed and then time one query, so before this change the timed region was
an O(1) array lookup over an index built in the untimed setup (130 ns), and after
it is the scan itself (2.81 ms). No work was added — every paragraph operation
the probe counted costs 0.61% to 0.68% *fewer* instructions over the same
lifecycle — the harness simply stopped paying for it outside the clock. `docx_semantic_one_paragraph` additionally times 130 ns
and 40 ns on the two before legs, which is below this clock's usable resolution;
its A/A floor is −69.2% and it should not be read as a number at all.

**`docx_file_eager_paragraph_count`'s +5.49% is a real, reproducible regression
on that selector, and this record does not know what it is.** It is above the 5%
review trigger, both orders agree (+6.07% and +4.90%) and the floor is 0.78%, so
it is not noise. What it is not: the same operation pair measured directly on the
probe is **0.69% cheaper in native cycles** (A/A floor −1.67%) and 0.63% cheaper
in instructions, and the first measured call after a harness-style untimed
`text()` preparation takes the **same number of page faults** on both legs
(median 1.0, spread ±5, 18 pairs per leg —
[`counts/first-call-faults.csv`](results/change-0592/counts/first-call-faults.csv)),
which refutes the first-touch hypothesis. An attempt to price that first call in
cycles could not decide it either: its own A/A floor came out at ±26%
([`counts/cold-first-call.csv`](results/change-0592/counts/cold-first-call.csv)).
The difference therefore lives in something the harness's timed region does that
the probe's does not, and it is recorded here as an open question rather than
explained away. The counterweight, not an excuse: the same corpus and the same
child-process harness show `docx_file_eager_full_text` at −40.22%, so whatever
this is, it is worth **+7.0 µs** on one selector against **−85.2 µs** recovered
on another over the same corpus in the same window.

**The floor exceeded 5% at p50 on three selectors** in this window —
`docx_file_eager_open_full_text_lifecycle` (8.54%), `docx_semantic_open` (17.34%)
and `docx_semantic_one_paragraph` (69.2%) — with eight agents building and
measuring on the other 31 cores. For those three the counts, not the timings,
carry the result, and this record says which is which rather than quoting the
convenient half.

## Correctness evidence

### Differential corpus: 332 documents, ten query families, byte-identical

`probe/src/corpus.rs` signs every main-document read this change can reach —
open outcome, `text()`/`extract_text()` digest and length, `paragraph_count()`,
`paragraphs()` digest and length, every `paragraph(i)` from 0 to one past the
end, `paragraph_text(i)` on the source-backed view, `tables()`, `elements()`,
`blocks()` — for both facades, with every error captured as its `Display` text so
a moved refusal changes the signature. It ran on both legs over:

| corpus | documents | result |
| --- | ---: | --- |
| every `.docx` under `test-data/` | 62 | signature files **byte-identical**, sha256 `c7c6730b43b82fedd0e22dc68cd114aa738cc26946cde70a1753556326d956a1` on both legs |
| every `.docx` retained by the change-0483 and change-0495 fuzz campaigns under `docs/performance/results/` — seeds, accepted corpora, crash inputs | 270 | signature files **byte-identical** |

That is 332 documents and 664 facade signatures with zero differences. The
refusals matter more than the successes. Eight `test-data/` fixtures are refused
at open by the eager facade on both legs with identical error text (two of them
by the source-backed facade as well), and six fuzz-corpus documents —
`malformed-deflate.docx` and `malformed-stored.docx` in three retained
change-0483 corpora — open successfully and then refuse **every** structural
query on both facades: `paragraph_count`, `paragraphs`, `paragraph(i)`,
`tables`, `elements`, `blocks` and text, all with the identical message
(`invalid DOCX XML: syntax error: tag not closed: '>' not found before end of
input`). Those six are the case this change had to get right, because they are
where the best-effort index fails and the streaming scanner has to report it
instead, and their refusals did not move.

### Tests

Five tests were added, each pricing one property the change could break.

In `crates/litchi-docx/src/parts/document_part.rs`:

- `text_and_block_reads_never_build_the_paragraph_index` observes the cell
  directly: it is empty after construction, still empty after `extract_text`,
  `table_count`, `tables`, `elements`, `blocks` and `xml_bytes`, and full after
  the first `paragraph_count`. This is the assertion the whole change rests on,
  and it is a direct observation rather than an allocation proxy.
- `deferred_index_answers_every_paragraph_query_like_the_streaming_scan` pins
  `paragraph_count`, `paragraphs` and `paragraph(i)` for every *i* up to one past
  the end against the streaming selectors the index replaces, then repeats them
  to exercise the cached round.
- `concurrent_first_paragraph_queries_agree_on_one_index` asserts
  `DocumentPart<'_>: Send + Sync` statically and races eight threads into the
  first query on one shared view.

In `crates/litchi-docx/tests/source_backed_semantic.rs`:

- `source_paragraph_queries_are_order_and_thread_independent` opens the same
  fixture twice, extracts text first on one view and selects paragraphs first on
  the other, and requires the four paragraph projections to be equal both ways;
  then races eight threads into the first query on a shared view.
- `managed_source_paragraph_queries_keep_their_eager_index` pins the branch that
  was deliberately not changed, including that `paragraphs()` still refuses with
  `UnsafeEdit` on a managed payload.

Two pre-existing tests carry the error-timing contract and were not touched:
`rejects_unterminated_selected_elements` (the index build fails and is swallowed,
and the streaming scanner still reports the error from `paragraphs`, `paragraph`
and `elements`) and
`malformed_source_document_errors_at_first_semantic_read`.

### Gates

All in the worktree, tails in
[`gates.txt`](results/change-0592/gates.txt): `cargo fmt --all --check` clean;
`cargo clippy -p litchi-docx --all-targets` clean under the workspace's deny
lints; `cargo test -p litchi-docx` **1,454 passed, 0 failed, 31 ignored** across
50 test binaries; `cargo doc -p litchi-docx --no-deps` clean under the deny
rustdoc lints; `python3 tools/non_iwork_gate.py check` exit 0;
`python3 tools/check_crate_boundaries.py` exit 0 (64 packages, 241 internal
dependency declarations); `python3 tools/check_report_claim_classification.py`
exit 0.

One gate fails and it is **pre-existing and unrelated**:
`tools/check_perf_claims.py --mode structural` exits 2 with "landed claim
`claim-0251-xlsx-xml-borrowed` requires strict evidence verification". It was
reproduced verbatim on the untouched before checkout at `08d968f8e`. This change
registers no claim, does not touch `claim-registry-v1.json`, and modifies no
`litchi-xlsx` file.

## Validation preserved

Nothing in this change is a validation. The MCE visibility pass, the
content-type check, the relationship and catalog validation at open, the
source-version fences, the managed route's `ensure_source_document_xml`
preflight, its `DocumentIndexAdmission` and query-parser admissions, and the
`UnsafeEdit` refusal of Arc-backed views on managed payloads all run exactly
where they ran. The paragraph scanner's own limits —
`MAX_PARAGRAPH_INDEX_RANGES`, the per-push `try_reserve`, and the depth and node
ceilings inside `scan_word_element_ranges` — are unchanged and run on the same
bytes; the 270 fuzz-corpus documents are the check that they still refuse what
they refused.

## Limitations

**No claim is registered.** `performance_claim: none`. The instruction and
allocation counts are exact for the probe's fixture on this host and build; the
timings are warm, single-host, and were taken while seven other agents worked on
the other 31 cores.

**The work is moved, not removed, for paragraph readers.** A caller that opens a
document and then selects paragraphs pays the same scan; the counts say it pays
0.6-0.8% less, which this record does not attribute and does not count as a
saving. What changed for that caller is *when*: the cost has moved
from `document()` into the first paragraph query. That is the intended
consequence of a lazy cache and what ADR 0005 asks for, but it is a shift in the
latency distribution of a single call, and this record does not pretend
otherwise.

**The budget-managed source-backed route is unchanged.** It still scans on every
`document()`. Deferring it moves a budget charge, which is a contract change; if
it is wanted, it needs a frozen design record that states what the admission
covers and when it is reserved. Nothing here authorizes that.

**The probe's fixture is marker-free and synthetic.** Its `document.xml` carries
no `mc:` markup, so `visible_document_xml` takes its cheap presence-scan path.
On a real producer file the MCE term is larger (change 0587, XML-1), so the
index's *share* of `document()` on such a file is unknown and is probably lower
than the 69% measured here. The absolute 45 M Ir removed is a function of the
paragraph count, not of the markers.

**An observation the probe turned up and this change does not act on.** After
this change `write_text_to` costs **101.9 M Ir more** than `text()` on the same
10,000-paragraph document (185.8 M against 83.9 M) and **90,000 more
allocations** — about nine per paragraph. Change 0587's DOCX-3 names the
`preflight_semantic_xml` second pass as a candidate and lists it unmeasured;
this record does not attribute that 101.9 M to preflight, because the streaming
parse and the sink are inside the same difference. It reports the gap and leaves
the attribution to whoever takes DOCX-3.

**DOCX-1 is untouched.** The eager edit route's four `Snapshot::from_xml` scans
and its whole-document compaction — rank 4 of the survey, about 64% plus 14% of
the timed instructions of a one-paragraph edit — are a different mechanism in a
different file and are not affected by this change in either direction. The edit
path does not use `DocumentPart`'s index.

**Not measured:** cold cache, physical devices, range sources, peak RSS,
concurrency scaling, real-producer documents, other platforms, and any
`perf stat` cycle count. The instruction counts rank work; change 0579's 1.24%
of instructions against 6.19% of cycles is the standing warning that the two are
not interchangeable.

## Retained evidence

[`results/change-0592/README.md`](results/change-0592/README.md) — the probe and
corpus-checker sources, the capture scripts, the instruction and allocation
tables with the raw callgrind totals they are differenced from, the four
signature files, the paired timing JSONs, `decision.json`, `gates.txt`,
`log-sections.md` and `cleanup.json`.

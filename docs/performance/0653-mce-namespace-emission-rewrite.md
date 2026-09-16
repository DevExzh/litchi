# 0653: the MCE codec declares each namespace once, the processed worksheet falls from 3,540,261 to 194,508 bytes, and the slicing consumers re-declare at the slice boundary

Status: retained, implemented in `litchi-ooxml-common`, `litchi-docx`,
`litchi-pptx` and `litchi-xlsx`. `performance_claim: none` — the counts and
paired medians below are reported as evidence, not registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This lands the namespace-emission rewrite that change
[0588](0588-mce-codec-namespace-emission.md) designed, implemented, measured and
then withdrew, together with the consumer migration the withdrawal was waiting
for, and it removes 93.9% of the cost change
[0649](0649-pptx-opened-transaction-real-deck-edit.md) attributed to it.

## Authority

Decision 1 of change [0652](0652-owner-decisions-for-the-third-wave.md), in the
owner's words:

> "MCE namespace re-declaration: accept the public API changes."

0652 reads that as authorizing "the namespace-emission rewrite 0588 implemented
and withdrew: the codec stops re-declaring namespaces on every element, and the
slice consumers (`Paragraph::extensions`, `Shape::xml`, the five XLSX raw
accessors, the two publishing writers) change their contract to match", and
requires this record to prove "the new contract of each consumer, stated; every
pristine part still copied byte for byte; every regenerated part's bytes
enumerated where they change, with the reason; the 0588 corpus and the 0649 real
deck measured on both sides of the codec". Each is a section below. 0652 assigns
this change no ADR amendment, and none is made: ADR 0003, 0005 and 0006 are read
below and none of their statements moves.

## What was changed

### 1. The writer emits each declaration once

`crates/litchi-ooxml-common/src/mce/codec.rs`. `write_start` called
`ctx.ns.for_each_effective`, which walked the whole inherited chain and emitted
**every** in-scope binding on **every** emitted start tag. It now emits:

* the element's **own** `xmlns`/`xmlns:*` attributes, in source order, through
  the same loop that writes its other attributes; and
* when the output drops one or more ancestors between this element and its
  nearest emitted ancestor — an `mc:AlternateContent` wrapper, a
  `mc:Choice`/`mc:Fallback` branch, or an element unwrapped by
  `mc:ProcessContent` — the declarations those dropped ancestors made, innermost
  binding winning, a prefix the element declares itself suppressed, and the
  reserved `xml` prefix never emitted. `Namespaces::for_each_hoisted` walks from
  the inherited chain to the emitted boundary and stops.

Each `Frame` carries `emitted_ns`, the namespace layer in effect at its nearest
**emitted** ancestor. `Namespaces::with_local` shares the parent layer verbatim
when an element declares nothing, so "is there anything to hoist" is one
`Arc::ptr_eq` and the common element pays nothing.

A start tag that drops no attribute, hoists nothing, holds no `Cow::Owned`
attribute value and contains no `&` or `<` is copied from the source bytes
verbatim, which also removes the per-attribute re-escaping.

This is 0588's frozen design, reinstated. Its retained patch
(`results/change-0588/design/namespace-emission.patch`) was written against the
pre-0588 file, so it is re-applied by hand onto the landed allocation work
rather than with `git apply`; the result reproduces 0588's predicted output size
for the measured worksheet to the byte (194,508) and its predicted per-call
instruction count to 0.11%.

### 2. One shared way to make a slice self-contained

`crates/litchi-ooxml-common/src/mce/fragment.rs` (new) exports
`mce::InScopeNamespaces` and `mce::self_contained_fragment(document, start, len,
limits)`. Given a document and one complete element span, it collects the
declarations the span inherits and re-declares on the span's own root element
exactly those it does not declare itself. That is a bounded, local operation on
one element, and it reproduces for the fragment root precisely what the old
writer did for every element.

Two fast paths keep it cheap: a document whose bytes before the span contain no
`xmlns` binds nothing, and a document in which every `xmlns` occurrence before
the span lies inside the document's own root start tag — the shape every real
producer writes — needs only the root's declarations and no structural walk. The
general case drives a `quick_xml::Reader` with the crate's existing
`BindingTracker`, bounded by `Limits::max_depth` and
`Limits::max_namespace_bindings`.

For consumers that do not slice bytes but rebuild a fragment through a
`quick_xml::Writer`, `litchi_ooxml_common::private` gains
`in_scope_declarations(resolver)` — the resolver's bindings with shadowing
resolved and undeclared prefixes dropped, which `NamespaceResolver::bindings`
does not do — and `with_in_scope_namespaces(element, bindings)`, which adds to
one `BytesStart` the declarations it does not already make. `BindingTracker`
gains `declaration_count` and `for_each_in_scope`.

### 3. `litchi-docx`

`namespace::self_contained_element_xml` wraps the shared helper for the crate's
`Arc<Vec<u8>>`-plus-range retained spans. `XmlRef::self_contained_bytes` uses the
crate's **existing** `XmlRef::namespace_resolver` lease instead, so a managed or
source-backed paragraph keeps its namespace-scan admission accounting.

* `Paragraph::extensions` and `Row::extension_ids` (the only **hard** consumers:
  `paragraph/extensions/codec.rs::parse_root` parses the span standalone and
  requires its root to resolve `Bound` into the wordprocessing namespace, with
  no fallback) now parse the self-contained span. This covers every route that
  reaches them — `parts/document_part.rs`, `header_footer/model.rs::Story`,
  `footnote.rs`, `comment.rs`, `table.rs` — because the repair happens at the
  accessor, not at each range producer.
* `OpaqueBlock`, `OpaqueInline`, `OpaqueRunContent`, `Comment` and `Note` each
  gain `self_contained_xml()`, and their `xml_bytes()` doc states the new
  contract.

### 4. `litchi-pptx`

`Common::xml()`/`Shape::xml()` keep their borrowed, zero-copy
`Result<&'a [u8]>`; their doc comments state the new contract and point at the
new `Common::self_contained_xml()`/`Shape::self_contained_xml()`
(`Result<Cow<'a, [u8]>>`). The one in-crate standalone re-parse —
`presentation/source.rs::parse_picture_relationship`, which refuses with
`"picture descriptor XML does not have a p:pic root"` — now takes the
self-contained span.

### 5. `litchi-xlsx`

The five raw-markup accessors 0588 named keep their signatures and their
documented meaning; four of them are `quick_xml::Writer` re-serializations
rather than byte slices, so they inject at the capture root through
`private::{in_scope_declarations, with_in_scope_namespaces}`:
`IgnoredErrorsExtension::markup`; `sheet_view::{PivotArea, Extension}::markup`
and `Entry::retained_xml`; `named_sheet_view::Markup::xml`;
`conditional_formatting::{Differential, Component}::raw_xml` (only on the
`parse_differential_formats` path, which is the processed-buffer one — the
`parse_conditional_formattings` path deliberately reads the **original** bytes
and is byte-for-byte unchanged). `chain::Chain::extension_list_xml` is the one
true byte-range slice and uses `mce::self_contained_fragment`.

`conditional_formatting::codec::wrap`, a synthetic root that supplied hardcoded
`s`/`x`/`x14`/`xm` aliases so a non-self-contained fragment could be re-parsed,
is no longer applied on the processed path: the fragment now carries the part's
own bindings, and keeping the wrapper would inject ~336 bytes of synthetic
aliases into every published `dxf` and would pin the transitional default
namespace onto strict parts.

## Breaking changes

**No signature, type or error changes.** Decision 1 authorized them; the
migration turned out not to need any. Every public change is additive or
documentation:

| item | change |
| --- | --- |
| `litchi_ooxml_common::mce::InScopeNamespaces` | new public type |
| `litchi_ooxml_common::mce::self_contained_fragment` | new public function |
| `litchi_ooxml_common::private::{in_scope_declarations, with_in_scope_namespaces}` | new items in the existing `#[doc(hidden)]`, explicitly unstable plumbing namespace |
| `litchi_docx::{OpaqueBlock, OpaqueInline, OpaqueRunContent}::self_contained_xml` | new public method |
| `litchi_docx::{Comment, Note}::self_contained_xml` | new public method |
| `litchi_pptx::shape::{Common, Shape}::self_contained_xml` | new public method |
| `litchi_docx::{OpaqueBlock, OpaqueInline, OpaqueRunContent, Comment, Note}::xml_bytes` | doc contract only: the bytes are a fragment whose namespace declarations are those in scope where it appears in the part and are not necessarily repeated on its own root |
| `litchi_pptx::shape::{Common, Shape}::xml` | doc contract only, same wording |

**One behavioural break, stated plainly.** The **bytes** returned by every
accessor above, and by the XLSX raw-markup accessors, change: inner elements
lose the declarations the old writer stamped on every tag, and a fragment root
gains exactly the set that used to be there. No accessor's namespace-resolving
meaning changes; the `reads/` differential below measures that over the whole
corpus. Callers that compared these bytes literally against a recorded string
will see a difference; callers that parse them see the same document.

**The five XLSX accessors' contract also becomes unconditional, deliberately.**
Their fragments are now re-declared whether or not the part mentioned the MCE
namespace, so a retained `extLst`, `sheetView`, `pivotArea`, `ext` or `dxf`
parses standalone on every part. Before this change it did so only when the part
happened to declare `xmlns:mc`, because only then did the writer run at all —
the same conditional fragility that made `Paragraph::extensions()` refuse on the
one `.docx` fixture without markers, which 0588 reported and this change fixes.
Gating the injection on the borrow decision would have been the smaller diff and
would have kept the accessors' bytes unchanged on marker-free parts; it is
rejected because it would preserve a contract that depends on an unrelated
property of the input. `litchi-xlsx`'s `chain::Chain::extension_list_xml` is
where this shows most: its `extLst` now always carries the part's default
namespace, so a strict-conformance re-write no longer re-namespaces that subtree
through the root. That was already the behaviour on every marker-bearing part;
it is now the behaviour on all of them, and the round trip is idempotent, which
a new assertion in the chain test pins.

## Why it is sound

**Namespace equivalence, by induction.** For each emitted element E, let A be
its nearest emitted ancestor. In the output, E's in-scope set is A's set
overridden by what is emitted between them: the declarations of every dropped
ancestor strictly between A and E (innermost winning), then E's own. In the
source, E's in-scope set is A's set overridden by the declarations of every
ancestor strictly between A and E, then E's own. The two override sequences
contain the same bindings in the same precedence order, so the sets are equal.
A's sets are equal by induction, and the root's trivially so. The corpus
differential below is the same statement measured: a namespace-resolving
projection of the processed output is identical on 6,964 real fixture parts.

**No output can grow.** A start tag's new emission is a subset of its old one:
the old writer emitted every in-scope binding, and hoisting emits a subset of
those (only the dropped ancestors' declarations), with the element's own
declarations moving from the writer's synthesized list into the copied attribute
list. Measured, not only argued: of the 6,964 parts, **0 grew**, 980 shrank and
5,984 were byte-identical, and the corpus's total codec output falls from
197,811,418 to 36,761,542 bytes.

**A refusal that disappears, and why it is not a weakened defence.**
`Limits::max_output_bytes` is unchanged to the byte and is still checked before
every write. Some inputs that used to be refused with
`LimitExceeded("output bytes")` now succeed, because the output they produce is
genuinely smaller. 0587 named that behaviour a defect — "a legitimately sized
part can still be refused as oversized output" — and 0588 recorded it as still
open. The bound did not move; the manufactured bytes did.

**A third refusal that disappears, the one change 0664 asked this record to
witness.** `litchi-docx`'s text sink bounds `xmlns:` attributes with
`MAX_SEMANTIC_TEXT_NAMESPACE_BINDINGS = 4096`
(`crates/litchi-docx/src/paragraph/codec/text.rs:413`), counted **cumulatively
over the whole parse**. Because the old writer re-declared every in-scope
binding on every start tag, a main document part with a 33-declaration root
exhausted that budget after about 124 elements, and change 0664 recorded that
`Document::write_text_to` **refuses its marker-bearing DOCX corpus** while the
byte-identical marker-free control is admitted. A retained witness
(`sink/`) reproduces the pair at the smallest size that shows it — one
`word/document.xml` with a 33-declaration root and 200 paragraphs, and the same
bytes with the markup-compatibility URI replaced by an inert URI of the same
length:

| leg | marker corpus | marker-free control |
|---|---|---|
| before | `ERR semantic text conversion failed after 0 bytes and 0 objects: invalid DOCX format: semantic DOCX XML exceeds 4096 namespace bindings` | `OK 2,689 bytes, 200 objects` |
| after | **`OK 2,689 bytes, 200 objects`** | `OK 2,689 bytes, 200 objects` |

**The limit is untouched** — no line of `text.rs` changed and 4,096 is still
4,096. What changed is that the writer no longer manufactures 33 declarations
per element, so the marker corpus now projects exactly what its control
projects, byte for byte and object for object.

**A second, larger refusal that disappears, found by a failing test.**
`litchi-xlsx`'s `print_options::parse_options` skips attributes whose key
contains `:` but reports a bare `xmlns` as `unknown printOptions attribute
'xmlns'`. The old writer re-declared the default namespace on the `printOptions`
it selected out of an `mc:AlternateContent`, so **`parse_print_options` refused
its own preprocessor's output**, and every worksheet whose `printOptions` sat
inside an `mc:AlternateContent` was unreadable. The workspace test
`litchi-xlsx --test source_backed_print_options::mce_limits_and_partial_sink_are_checked`
asserted that refusal; it now asserts the Fallback's value is read, and a new
test, `the_old_writers_redundant_declaration_was_refused_by_this_parser`, pins
both shapes so the defect and its removal stay written down. No production line
of `print_options.rs` changed.

**One new refusal surface, deliberately accepted.** On a **managed**
source-backed paragraph, `Paragraph::extensions()` now also charges its
namespace scan against the execution budget, because it now performs that scan:
`XmlRef::self_contained_bytes` goes through the crate's existing
`namespace_resolver` lease, which takes a `namespace_scan_admission` and checks
it per event. A document whose budget is tight enough can therefore see
`extensions()` refuse where it used to succeed. That is the correct accounting
under ADR 0005 and ADR 0031 — the alternative, reading the owner's bytes without
charging for them, would be the defect — and it applies only to the managed and
source-authorized owners; an unmanaged owner's `namespace_scan_admission` is
`None` and nothing is charged. No limit value moved.

**Correctness over performance, twice.** (a) The self-containment helper
re-declares **every** inherited binding the fragment root does not declare, not
only the prefixes the fragment appears to use: deciding "appears to use" means
parsing the whole fragment, and a fragment containing an unfamiliar subtree
would be mis-served by a use-set that a scanner computed. The cost is a few
hundred bytes on one element. (b) `litchi-pptx`'s picture-descriptor read now
allocates one owned copy per descriptor (590 bytes on the fixture) where it used
to borrow; the borrowed alternative would have required seeding the re-parse's
resolver from the owner, which is a larger change than this migration is scoped
to. It is still a net reduction: the before leg borrowed 3,047 bytes out of a
16,925-byte processed owner, the after leg borrows 239 out of 1,481 and copies
590.

**ADR reading.** ADR 0003: no public type, signature or crate dependency is
removed or changed; the new items are additive and no format crate gains a
dependency. ADR 0005: no archive type, raw lock or executor is exposed; no read,
no positional source and no validation moves; `litchi-docx`'s managed
namespace-scan admission is reused rather than bypassed, so retained-state
accounting is unchanged. ADR 0006 is the substance — the processed stream is
still produced by the same mandatory pass, with the same checks in the same
order against the same values; what changes is only which namespace declarations
the writer emits, and the two writers that publish processed bytes publish
output that is strictly closer to their source. No `unsafe` is added and every
reservation stays fallible.

**What is deliberately not touched.** The eligibility gates
(`raw::worksheet::source_stream_eligible`, `MceRewriteEquivalence`), the limit
values, the presence scan, `active_offsets`' original-relative offset contract,
and `litchi-pptx`'s guides writer and `litchi-xlsb`'s drawing-anchor transfer.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0, `--release --locked`, every measured process pinned with
`taskset -c 8`, with seven other agents building and measuring on the other
cores. Base `70d7768cc`. Binaries staged outside every Cargo target directory
and their SHA-256s recorded in the packet README.

### Deterministic counts first

**Output bytes.** The measured worksheet
(`test-data/poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx`,
`sheet1.xml` 209,931 B, 11,578 elements, root declares `xmlns`, `xmlns:r`,
`xmlns:mc`, `xmlns:x14ac` with `mc:Ignorable="x14ac"`) goes from **209,931 in ->
3,540,261 out (16.86x)** to **209,931 -> 194,508 (0.93x)**, which is 0588's
predicted figure for this design to the byte. Change 0587's marker-stripped
control takes the borrowed fast path in both legs (194,257 -> 194,257).

Over the whole fixture corpus — 320 `.docx`/`.xlsx`/`.pptx` fixtures, every
`.xml` and `.rels` member, **6,964 parts** — the codec's total output falls from
**197,811,418 to 36,761,542 bytes (-81.42%)**. **0 parts grew**, 980 shrank and
5,984 were byte-identical (the borrowed fast path).

### Instructions (callgrind, isolation pairs where noted)

| Measurement | before | after | delta | tier |
|---|---:|---:|---:|---|
| MCE codec, real `sheet1.xml`, per call ((r6-r1)/5) | 305,273,402 | 61,464,892 | **-79.87%** | measured |
| MCE codec, control `sheet1.xml`, per call | 4,678,933 | 4,678,940 | +0.00% | measured |
| Eager open + one cell, real fixture | 496,700,777 | 154,479,583 | **-68.90%** | measured |
| Eager open + one cell, control fixture | 74,925,234 | 70,355,722 | -6.10% | measured |
| Source-backed read of one cell, real fixture | 651,068,133 | 304,172,603 | **-53.28%** | measured |
| Source-backed read of one cell, control fixture | 215,822,344 | 208,766,384 | -3.27% | measured |

The before column re-measures what 0588 reported as its after column
(305,095,214 / 497,483,936 / 654,148,756), on a binary rebuilt here, and agrees
to 0.06%, 0.16% and 0.47%. The after column reproduces what 0588 measured on its
withdrawn patch (61,399,465 / 156,028,289 / 309,307,843) to 0.11%, 1.0% and
1.7%; the small extra comes from the consumer migration, chiefly `litchi-xlsx`
no longer wrapping every captured `dxf` in a synthetic root. The borrowed fast
path is unchanged to five significant figures, as it must be: the change touches
only the rewrite path.

In the before profile the codec's self cost is `BoundedOutput::extend_from_slice`
21.81%, `__memcpy_avx_unaligned_erms` 16.45% and `esc` 9.10% — the expansion.
None of the three appears in the after profile's table, whose top entries are
parsing work (`CharSearcher::next_match` 4.46%, `codec::start` 3.18%,
`read_event_impl` 2.81%).

### Paired timing, public XLSX reads

80 samples per leg (two blocks of 40), ordered A1 B1 B2 A2, five untimed warmups
per block, one pinned process per block, with two further before-leg blocks as
the A/A floor in the same window. The floor quoted is the **widest p50 spread
across every before-leg block**, not a single favourable pair. Nanoseconds.

| Scenario | before p50 | after p50 | after vs before | before vs after | A/A floor |
|---|---:|---:|---:|---:|---:|
| Eager open + one cell, real | 24,980,930 | **8,748,352** | **-64.98%** | +185.55% | 2.17% |
| Source-backed read, real | 36,332,254 | **19,458,014** | **-46.44%** | +86.72% | 2.41% |
| Eager open + one cell, control | 4,407,803 | 4,110,072 | -6.75% | +7.24% | 2.83% |
| Source-backed read, control | 14,009,836 | 13,568,563 | -3.15% | +3.25% | 4.13% |

**The eager-real row took three windows and the first two are retained as
failures.** Its before-leg blocks came back bimodal — 26.5 ms in some blocks and
108 to 191 ms in others — on a host whose load average was 25 to 45 across 32
cores with seven other agents measuring, none pinned away from CPU 8; the first
window's A/A spread was 316.99% and a longer re-measurement's was 622.68%. Both
sample sets are retained under `timing/samples/`. The row above is a third
window measured after the other agents quieted (load average 9.5), six ABBA
rounds of 20 samples with a before-leg floor block in each: **twelve before
blocks spanning 24,705,350 to 25,241,252 ns, a 2.17% spread**, and the pooled
before p50 of 24,980,930 agrees with 0588's measurement of the same code
(24,892,964) to 0.35%. Its samples are under `timing/window3/`. The lesson is
recorded rather than smoothed: a 25 ms scenario on this host is destroyed by an
unpinned neighbour, and the two failed windows are what that looks like.

The control's -6.75% is above its floor but only 2.4x it, and is reported rather
than leaned on; 0588 explains why the control moves at all (`xl/workbook.xml`
and `xl/styles.xml` still declare the MCE namespace in both packages).

### The 0649 real deck

The 108 KB, 103-member deck
`test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx`,
through change 0649's own probe and its four documented public calls. 30 samples
per block, four blocks per leg ordered A1 B1 B2 A2 twice, plus two A/A blocks.

| Phase | before p50 | after p50 | delta | A/A |
|---|---:|---:|---:|---:|
| `opened_presentation()` — the capture | 59,719,392 | **11,056,451** | **-81.49%** | 3.93% |
| `Snapshot::edit()` — the working clone | 28,690 | 24,200 | -15.65% | 27.96% |
| `set_shape_text` | 4,803,917 | **1,183,656** | **-75.36%** | 2.84% |
| `Transaction::commit` | 62,763,016 | **11,822,591** | **-81.16%** | 3.29% |
| `apply_opened_presentation_commit` | 236,411 | 121,801 | -48.48% | 23.87% |
| **edit total** | **127,442,488** | **24,188,616** | **-81.02%** | 3.72% |

The two sub-millisecond phases have 27.96% and 23.87% floors; they are reported
and not relied on. Across all blocks the before leg spans 123.19-141.42 ms (a
14.80% spread, which is the honest A/A figure for this window) and the after leg
24.15-24.54 ms (1.60%); the -81% result is 5.5x the wider of the two.

**Against 0649's numbers.** 0649 measured this edit at **128.123 ms** on its own
probe and 0638 at **133.61 ms** through the harness, with a marker-stripped
control at **7.822 ms** — 93.9% of the cost removed by stripping the markers.
This change removes **81.0%** of it: 127.44 ms -> 24.19 ms on the probe here.
The control deck re-measured in this window is the reference point for what is
left: its four after-leg blocks span 8.34-8.46 ms and its least-contended
before-leg block is 8.32 ms, so the control is unchanged, as it must be for a
deck the codec never rewrites, and the after leg's 24.19 ms is **2.9x** the
control rather than 15.4x. The remaining gap is the codec's parse, which this
change does not remove: it still re-tokenizes every marker-bearing slide six
times per edit, it just no longer writes 16x its input while doing so.

### The 0638 ordinary-save PPTX real edit

`tools/perf-baseline`, change 0638's real-file selectors on the same deck, 20
samples per block, ABBA plus two A/A blocks, pinned to CPU 8.

| Selector | before p50 | after p50 | delta | A/A floor |
|---|---:|---:|---:|---:|
| `pptx_real_file_ordinary_save_edit` | 133,438,301 | **22,960,854** | **-82.79%** | 4.77% |
| `pptx_real_file_ordinary_save_lifecycle` | 129,960,346 | **24,680,588** | **-81.01%** | 1.68% |
| `pptx_real_file_ordinary_save_counting_publish` | 271,171 | 327,597 | **+20.81%** | 5.54% |

The first row is 0638's own 133.61 ms measurement reproduced to 0.13% on the
before leg, and it agrees with the probe's -81.02% from an independent
instrument.

**The +20.81% is a review trigger and is reported, chased and left open.** The
selector times `Package::to_bytes` alone, and its published bytes are identical
on both legs (the corpus publication differential below finds **0** differing
members on every PPTX save route except the guides writer). Re-measured in
isolation — its own process, that case only, 50 samples, six blocks per leg —
it is 269,141 ns against 304,761 ns, **+13.2%**, against a before-leg block
spread of 20.0% and an after-leg spread of 36.0% in the same window, so it does
not separate from the floor. A `perf stat` isolation pair over the same case
(samples 20 against 120, differenced) puts the whole per-sample work at
2,667,498,219 instructions before and 445,718,952 after, **-83.3%**, and cycles
at 586,314,113 against 104,426,084, **-82.2%** — the untimed open-and-edit that
precedes each timed publish is what collapsed. The most likely mechanism is
therefore heap state rather than work: the before leg's edit allocates and
frees roughly 28 MB of codec output per iteration, leaving the allocator holding
arenas the publish then reuses, while the after leg's publish has to grow the
heap itself. That is not demonstrated, only argued, and it is not resolved here.
In absolute terms the row is +56 microseconds on a route whose edit phase falls
by 110 milliseconds.

### Change 0664's marker-bearing selectors

Change [0664](0664-perf-harness-marker-bearing-corpora-and-save-allocations.md)
landed, after this change's own measurement window, the harness family this
program had been missing: generated DOCX and PPTX corpora that carry the
markup-compatibility namespace across the package, each with a **byte-identical
marker-stripped control** — the same members, the same per-member uncompressed
lengths, the same start-tag and attribute counts, proved before any sample runs
— so a marker/control pair measures the codec branch and nothing else. Its
harness-only diff (commit `e76897b0e`, `tools/` only) was applied to a fresh
checkout of the base and to this worktree to build the two binaries, and
reverted before the gates and the commit, exactly as change 0649 did with change
0638's diff. 30 samples per block, ABBA plus two A/A blocks, pinned to CPU 8.

| Selector | before p50 (ns) | after p50 (ns) | delta | A/A floor |
|---|---:|---:|---:|---:|
| `pptx_marker_ordinary_save_edit` | 155,537,122 | **21,340,230** | **-86.28%** | 2.96% |
| `pptx_marker_ordinary_save_lifecycle` | 157,658,453 | **22,290,470** | **-85.86%** | 2.34% |
| `pptx_marker_eager_full_text` | 60,072,104 | **8,999,773** | **-85.02%** | 2.01% |
| `pptx_marker_source_full_text` | 59,776,558 | **8,804,368** | **-85.27%** | 1.86% |
| `docx_marker_ordinary_save_edit` | 8,016,434 | **1,081,966** | **-86.50%** | 1.09% |
| `docx_marker_eager_full_text` | 6,779,992 | **525,727** | **-92.25%** | 4.60% |
| `docx_marker_source_full_text` | 6,450,520 | **480,192** | **-92.56%** | 3.58% |
| `pptx_marker_control_ordinary_save_edit` | 7,479,806 | 7,263,614 | -2.89% | 5.71% |
| `pptx_marker_control_ordinary_save_lifecycle` | 8,549,744 | 8,138,019 | -4.82% | 4.89% |
| `pptx_marker_control_eager_full_text` | 4,579,070 | 4,581,255 | +0.05% | 1.49% |
| `pptx_marker_control_source_full_text` | 4,441,479 | 4,448,969 | +0.17% | 0.84% |
| `docx_marker_control_ordinary_save_edit` | 594,098 | 606,968 | +2.17% | 1.13% |
| `docx_marker_control_eager_full_text` | 254,121 | 259,356 | +2.06% | 4.47% |
| `docx_marker_control_source_full_text` | 214,561 | 214,621 | +0.03% | 6.38% |

**Every control row moves less than its own floor**, which is the result that
makes the marker rows a measurement of the codec branch: the control packages
are byte-comparable to the marker ones and take the borrowed path, so the change
must not touch them, and it does not.

**The ratio 0664 asked this record to quote.** The marker/control ratio is the
whole cost of the rewriting branch, and it collapses:

| pair | before | after |
|---|---:|---:|
| `pptx_..._ordinary_save_edit` | **20.79x** | **2.94x** |
| `pptx_..._ordinary_save_lifecycle` | 18.44x | 2.74x |
| `pptx_..._eager_full_text` | 13.12x | 1.96x |
| `pptx_..._source_full_text` | 13.46x | 1.98x |
| `docx_..._ordinary_save_edit` | 13.49x | 1.78x |
| `docx_..._eager_full_text` | 26.68x | 2.03x |
| `docx_..._source_full_text` | 30.06x | 2.24x |

0664 measured `pptx_marker_ordinary_save_edit` at 150.368 ms and its control at
7.432 ms, a 20.23x ratio; the before leg here is 155.537 ms and 7.480 ms, 20.79x
— 3.4% and 0.6% from 0664's figures, inside its own 3.39% repeat spread. What
remains after the change is not the expansion but the parse: a marker package is
still rewritten rather than borrowed, and rewriting still costs about twice
borrowing.

### No regression on the marker-free harness corpora

0587 section 4 records that the harness's generated worksheets contain no `mc`,
`x14ac`, `dyDescent` or `<cols>` markup, so every existing XLSX selector takes
the codec's borrowed fast path. 0588's four selectors were run ABBA against an
A/A block, 30 samples, three corpus shapes each, twelve rows:

| Selector | shapes | worst move | worst A/A floor |
|---|---|---:|---:|
| `xlsx_first_cell` | tiny, medium, dense-wide | +1.22% | 0.94% |
| `xlsx_full_cell_scan` | tiny, medium, dense-wide | -1.79% | 1.38% |
| `xlsx_narrow_column_range_scan` | tiny, medium, dense-wide | -4.37% | 6.06% |
| `xlsx_open_owned` | tiny, medium, dense-wide | +0.41% | 2.25% |

Every one of the twelve rows moves by less than its own floor or within a factor
of two of it. **No selector regressed above 5%.** Per-row figures are in
`harness/summary.txt`.


## Correctness evidence

### The corpus differential: 6,964 parts, both modes

`oracle/mce_oracle_0653.py` walks every `.xlsx`, `.docx` and `.pptx` under
`test-data` — 320 fixtures — opens each archive and runs the codec over every
`.xml` and `.rels` member, a superset of the parts the codec is actually applied
to. For each of the **6,964 parts** it compares the before and after binaries on

* a **namespace-resolving projection** of the processed output: the resolved
  (namespace URI, local name) of every element and attribute with its normalized
  value, every text run, CDATA, comment and declaration, with namespace
  declarations themselves excluded — they are exactly what this change moves —
  and an unresolvable prefix rendered as `!UNKNOWN-PREFIX:<prefix>`, plus the
  complete `Report`; or, on refusal, the error's `Debug` identity and `Display`
  message;
* the **exact output length** and the borrow-versus-own decision.

**6,964/6,964 identical in the resolving projection. 0 canonical mismatches, 0
`Report` mismatches, 0 borrow-decision mismatches, 0 refusal mismatches, 0 parts
grew.** Byte identity is deliberately not asserted: this change is byte-visible
by design, and 980 parts are shorter.

### Adversarial refusal differential: 30,000 mutants

`oracle/mce_mutate.py` (0588's, unchanged) takes nine MCE-bearing seeds — the
real worksheet, a real `word/document.xml`, a real `ppt/slides/slide1.xml`, a
real `xl/workbook.xml`, and five synthetic documents covering
`AlternateContent`/`Choice`/`Fallback`, `ProcessContent` unwrapping,
`PreserveElements`/`PreserveAttributes`, opaque extension branches, entity and
whitespace-normalized attribute values, and `MustUnderstand` — and applies one to
three single-byte substitutions, deletions, insertions or truncations drawn from
an XML-aware alphabet, seeded deterministically.

**30,000 mutants, of which 16,661 refused. 0 mismatches.** Every accepted mutant
produced an identical resolving projection with identical counters; every refused
mutant produced the identical error variant, payload and message. The counts are
0588's exactly, which is the point: the refusal surface did not move.

### Publication: every pristine part still copied byte for byte

`publication/` runs twelve documented save routes over the whole corpus on both
legs — ordinary no-op save, ordinary save with an edit, and the source-backed
route for each of DOCX, XLSX and PPTX, plus PPTX's guides publisher — and
compares every published archive member against the same member of the source
and against the other leg. **20,421 member rows per leg, 0 rows present on one
leg only.**

* **19,896 pristine member rows** (published byte for byte from the source on the
  before leg). **0** stopped being identical to their source on the after leg,
  **0** pristine published hashes moved, **0** became newly identical. The
  copy-through routes are perfect on both legs: `xlsx_noop_save` and
  `xlsx_source_backed_tab_state` 3,003/3,003, `pptx_noop_save` 3,417/3,417,
  `pptx_source_backed_slide` 2,895/2,895.
* **Exactly 4 member rows differ between the legs**, all on route
  `pptx_guides_publish`, all member `ppt/presentation.xml` — the one writer that
  publishes codec output. They are the only four PPTX fixtures whose
  `ppt/presentation.xml` contains the MCE namespace, so the coverage of the
  changed path is complete for this corpus and narrow.

  | fixture | bytes before -> after | `xmlns:` declarations |
  |---|---:|---:|
  | `sd/qa/unit/data/pptx/slide-section-test.pptx` | 24,166 -> 3,235 | 309 -> 10 |
  | `sd/qa/unit/data/pptx/tdf89064.pptx` | 4,737 -> 1,377 | 58 -> 10 |
  | `office-interop/libreoffice-resaved/shapes-litchi.pptx` | 6,108 -> 1,488 | 76 -> 10 |
  | `ooxml/pptx/shape-soft-edges.pptx` | 8,603 -> 1,883 | 106 -> 10 |

  **All four are canonically EQUAL** under a namespace-resolving canonicalizer,
  with identical node counts (135/135, 31/31, 43/43, 58/58). Only namespace
  declarations moved, and the reason is the one this change is about: the writer
  used to repeat all ten root declarations on every emitted start tag.
* **Refusal identity: 0 differences.** 1,276 outcome rows per leg, 867 OK and 409
  ERR on both, with an identical per-route `Debug`-variant table. 91 rows differ
  in raw refusal *text*, and a before-against-before control shows 92 such
  differences with 0 member differences: `xlsx_source_backed_cell_values` names
  one relationship URI out of an unordered set, which is pre-existing
  nondeterminism, not a leg difference. After normalizing that one message,
  before-against-after differs in 4 rows and before-against-before in 0.
* Published bytes over all twelve routes: 39,831,947 -> 39,831,551, a difference
  of **396 bytes**, all of it the guides route; uncompressed, that route moves
  1,382,758 -> 1,347,127 (-2.6%) and every other route moves 0.

### Public reads: 179,205 observations per leg

`reads/` drives every documented public read of the three format crates over the
same 320 fixtures on both legs — paragraph and cell text, extensions, tables,
comments, notes, slides, shapes, sheet views, and every migrated raw-markup
accessor — and diffs the two digests line by line. The `(fixture, key)` sequence
is byte-identical between legs.

| class | count |
|---|---:|
| a — a text or typed value changed | **0** |
| b — an outcome changed (OK/ERR, or one error to another) | **4** |
| c — markup bytes changed, namespace-resolving canonical form unchanged | 1,886 |
| d — markup bytes changed and canonical form changed | **0** |

The four class-b rows are `Paragraph::extensions()` on the four paragraphs of
`sw/qa/extras/ooxmlexport/data/listWithLgl.docx`, where
`ERR InvalidFormat("Word extension XML must have one [112] root")` becomes
`para_id=none;text_id=none;no_spell_err=none`. **That is 0588's reported
pre-existing `litchi-docx` defect, fixed.** That fixture declares `xmlns:w` only
on `w:document` and carries no MCE markup, so the codec never rewrote it and the
`w:p` span never carried `xmlns:w` in either leg; the migration re-declares at
the slice boundary and the read succeeds. The fixture contains no
`w14:paraId`/`textId`, so `none` is the right answer. Corpus refusals fall from
122 to 118 and nothing else moves. The class-c rows are the accessors whose
bytes this change is expected to shrink: total markup bytes over 2,209
observations fall from **13,250,784 to 1,615,294, -87.81%**.

The canonical-equality verdict is not vacuous: `reads/compare.py` carries four
negative controls (an attribute value changed, an attribute added, the root
namespace rebound, and an unperturbed copy) and all four are detected.

### Unit tests

`crates/litchi-ooxml-common/src/mce/tests.rs` replaces 0588's
`namespace_emission_contract_tests` with the new contract and adds
`self_contained_fragment_tests` (fifteen and ten tests). They pin: each namespace
declared once while the resolving projection is unchanged; an inner span that no
longer resolves standalone and does resolve after
`mce::self_contained_fragment`; hoisting onto several emitted children of a
dropped wrapper; a child re-declaration shadowing a hoisted one; nested dropped
wrappers; a `ProcessContent` wrapper's default namespace travelling with its
content; a non-selected branch's declarations not leaking; the `xml` prefix never
re-declared; the source start tag copied verbatim when nothing changes; the
output bound no longer consumed by redundant re-declaration; and, for the
helper, the root-declarations path, an intermediate ancestor's declaration, a
span that re-binds a prefix, a sibling's declaration not leaking, a default
namespace undeclared by an ancestor not reinstated, the empty-element insertion
point, the borrowed no-op, the two range refusals by name, the
`max_namespace_bindings` bound, and an escaped namespace value round-tripping.

`crates/litchi-ooxml-common/src/mce/tests.rs` also keeps 0588's
`process_content_still_refuses_a_wrapper_that_binds_a_prefix`,
`the_borrowed_name_resolver_keeps_every_refusal_identity` and
`amortized_output_growth_keeps_the_bound_exact` unchanged.

In the consumer crates: `litchi-docx` gains
`namespace::tests::a_retained_span_resolves_only_after_the_inherited_declarations_return`
(a real Word fixture whose `w:p` span does not resolve and does after the
repair) and a real-fixture `paragraph_extensions` test that reads `w14:paraId`
from both paragraphs and table rows; `litchi-pptx` gains
`tests/pptx_shape_fragment_namespaces.rs` (4) and
`presentation::source::picture_fragment_tests` (2), which pin that a picture
descriptor with unconventional prefixes refuses on the borrowed span and parses
on the self-contained one; `litchi-xlsx` gains one standalone-resolution test per
migrated accessor.

### Gates

All exit 0; tails in `gates.txt`.

* `cargo fmt --all --check`.
* `cargo clippy --all-targets` on `litchi-ooxml-common`, `litchi-docx`,
  `litchi-pptx` and `litchi-xlsx` — workspace lints are `deny`, so any warning
  fails.
* `cargo test` on those four plus the consumer crates of the shared one:
  `litchi-xlsb`, `litchi-drawingml`, `litchi-spreadsheet-drawing` and
  `litchi-opc`. `litchi-ooxml-common` is 261 lib tests, `litchi-docx` 952,
  `litchi-pptx` 881 across 77 suites, `litchi-xlsx` 1,004 lib tests and every
  integration binary.
* `cargo doc --no-deps` on the four touched crates — rustdoc lints are `deny`.
* `cargo test -p litchi --features docx,xlsx,pptx,xls`.
* `cargo test` in `tools/perf-baseline`: **522 passed, 0 failed, 1 ignored**.
  The two flaky allocator tests in `docx_bounded_tail_append_compare` that the
  wave's briefing lists as known did not fire in this run.
* `python3 tools/non_iwork_gate.py verify`: 45 bulk tree roots, 35 facade-safe
  trees and the combined tree verified.

The differentials above are gates too, and each one's driver and raw report is
retained: the corpus oracle, the 30,000-mutant differential, the publication
census, the public-read digest and the text-sink witness.

Change 0664's harness diff (commit `e76897b0e`, `tools/` only) was applied to a
fresh detached checkout of the base and to this worktree to build the marker
selector binaries, and reverted — `git checkout -- tools/` plus removal of the
untracked `tools/perf-baseline/src/marker_shape.rs` — before these gates and
before the commit. `git status --short tools/` reports nothing.


## Validation preserved

Nothing about validation moves. The codec still performs its mandatory pass over
every part whose bytes contain the MCE namespace, in the same order, with the
same bounds: input bytes, output bytes, depth, namespace bindings, directive
tokens and choices per `AlternateContent`. Duplicate attribute detection stays on
(`with_checks(true)`), every attribute value is still decoded and normalized
before it is inspected, every attribute name is still expanded — and so still
refused when it is not a qualified name or its prefix is unbound — before the
writer decides whether the tag can be copied, DTDs and processing instructions
are still rejected, custom entities are still rejected, and `MustUnderstand`
still refuses with `Error::MustUnderstand`. ADR 0005's "validation must not
mutate" is unaffected. The new self-containment helper adds validation rather
than removing any: it refuses a range outside its document, a range that does
not begin with an element, a malformed document, and a scope that exceeds
`max_depth` or `max_namespace_bindings`, each with a typed error.

## Limitations

* **No claim is registered.** `performance_claim: none`, no claim-registry
  entry. Every number is scoped to one host, one build, one pinned CPU, this
  fixture corpus and the named selectors.
* **The eager-XLSX wall clock needed three windows.** The first two were
  destroyed by unpinned neighbours (A/A spreads of 316.99% and 622.68%) and are
  retained as failures; only the third, measured at load average 9.5, is quoted.
  Every other timing row in this record was measured in the contended window and
  carries the floor it earned there.
* **The `counting_publish` regression is not resolved.** +20.81% in the ABBA and
  +13.2% in isolation, against A/A block spreads of 20% and 36%. The published
  bytes are identical and the per-sample work fell 83%, so the argued mechanism
  is allocator heap state; it is argued, not demonstrated.
* **Six migrated XLSX accessors have no corpus witness.** The read differential
  found **zero** markup observations for sheetView `ext`, sheetViews-collection
  `ext`, `pivotArea`, `ignoredErrors` `ext`, named-sheet-view `ext` and
  calculation-chain `extLst`, and none for an inline conditional-formatting
  `dxf`. Their migration is covered only by the unit tests added beside them.
* **The `litchi-pptx` source-backed picture path has no corpus witness either.**
  `SourceBackedPresentation::images()` succeeded on 110 slides but on zero
  MCE-rewritten ones, because that reader refuses those up front with an
  identical `UnsafeEdit` on both legs. The migration there is covered by the
  synthetic tests in `tests/pptx_shape_fragment_namespaces.rs`.
* **`litchi-xlsb`'s drawing-anchor transfer is not exercised.** It is the second
  writer that publishes codec output; the OOXML fixture corpus has no `.xlsb`
  input for it, so the publication differential cannot reach it. Its code is
  unchanged and it already re-injects the source root's declarations explicitly,
  which is why 0588 named it the template.
* **DOCX source-backed publication is nearly uncovered.** Both source-backed
  DOCX routes refuse 61 of 62 fixtures, 59 of them with the identical
  `UnsafeEdit{reason: "source-backed document transactions do not support
  markup-compatibility branch selection"}` on both legs, so only 6 DOCX
  source-backed member rows exist.
* **The guides publisher is exercised on 4 fixtures.** It refuses 74 of 78 PPTX
  fixtures with `XmlPublication{NotCompact(FormattingWhitespace)}` before
  publication — decision 2 of 0652 is the change that loosens that — and the 4
  that pass are exactly the 4 MCE-bearing ones.
* **A stale production comment is left for another change.**
  `crates/litchi-xlsx/src/raw/worksheet/mod.rs` documents
  `MceRewriteEquivalence` and `MAX_REWRITTEN_DECLARATIONS` in terms of "the
  preprocessor re-declares every in-scope binding on every start tag it writes".
  The predicate stays **sound** — it accumulates every declaration seen so far
  and charges it against every start tag, which is an upper bound on what the
  new writer emits, and an over-conservative admission only costs a fallback to
  the authoritative two-pass path — but it is now much looser than it needs to
  be, which is an opportunity, and its prose is now wrong. That file belongs to
  change 0657 in this wave, so it is reported here and not edited.
* **Three adjacent processed-buffer slicing sites in `litchi-xlsx` are reported,
  not migrated**, because 0588 did not name them and this change's scope is the
  sites it did: `query_table/codec.rs::parse_extension_list` (highest risk: it
  retains `ext` subtrees and re-emits them through `write_query_table`),
  `auto_filter/package.rs::capture` (serializes a subtree out of the processed
  buffer and re-parses it standalone), and
  `data_validation/codec/wire.rs::capture_collections` (re-parsed under a
  hardcoded six-namespace prologue that guesses unknown prefixes as `x14`).
  `named_sheet_view/codec.rs::adapt_filter_root` rebuilds a fragment root and
  drops its declarations; that payload is never published, and it is likewise
  reported rather than changed.
* **The `ProcessContent` prefix refusal is not fixed.** An element unwrapped by
  `mc:ProcessContent` still cannot declare a prefixed namespace
  (`unbound prefix xmlns`). It is pre-existing, it stays pinned by a test, and
  fixing it would move a refusal.
* **The codec still parses every part it used to parse.** The remaining 24.19 ms
  of the real deck's edit is `litchi-pptx` asking for six whole-slide
  markup-compatibility passes per edit (0649's finding, unchanged by this
  record) against a marker-free control at 8.3 ms. Removing those passes is a
  different change.
* **The managed budget charge is not measured.** `Paragraph::extensions()` on a
  managed source-backed paragraph now charges its namespace scan against the
  execution budget, so a sufficiently tight budget can turn a success into a
  typed refusal. The corpus read differential exercises the ordinary readers, so
  it does not witness that boundary; the crate's managed test suites pass
  unchanged, and no budget sweep was run.
* **No allocation-count, RSS, cold-cache, physical-I/O, throughput, concurrency
  or cross-platform measurement was taken**, and none is claimed.


## Retained evidence

[`results/change-0653/`](results/change-0653/README.md) lists every file with
its provenance, the binary SHA-256s and the exact commands.

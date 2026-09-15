# 0623: one read per contiguous run of structural members — the OOXML open's request count falls again, for no extra bytes

Status: retained, implemented in `soapberry-zip` (a read-side accessor) and
`litchi-opc` (change [0577](0577-ooxml-open-relationship-parts.md)'s candidate
(c)). `performance_claim: none` — no claim-registry entry in this wave; the
paired medians and the deterministic counts below are reported as evidence, not
registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **ZIP-5** of
[0587](0587-remaining-opportunity-survey.md) (rank 15): the accessor change
0577's run-coalesced structural prefetch was blocked on, and the prefetch
itself. 0577 froze the design, its invariants, its admission gates and its
intra-run error-precedence decision; this record implements exactly that design
and reports what it cost and bought.

**The baseline moved before this landed.** 0577 modelled 89 open requests
falling to 10 on the 132-member workbook. Change
[0611](0611-zip-single-read-per-member.md) has since taken that open to **45**,
so everything here is measured and stated against 45, not against 89. The
modelled destination is unchanged, and it is reached exactly: **45 → 10**.

## What was changed

Two crates, four production files.

**`crates/soapberry-zip/src/office.rs` — additive only.** One new public type
and one new method. `git diff` shows no existing line touched, so no read path,
no grammar and no verdict in this crate can have moved.

* `LocalSpanHint` — an offset and a length, with `offset()`, `length()` and
  `end()`.
* `IndexedArchive::local_span_hint(EntryId) -> Option<LocalSpanHint>` — the byte
  span this archive's own first read of that member covers, or `None` for a
  member the span window does not admit. It returns exactly what change 0611's
  `member_span_length` computes, alongside the member's local-header offset.
  Both facts are already in the index; neither is read from the source, and
  neither is ZIP framing.

**`crates/litchi-opc/src/source_backed/read_ahead.rs`.** A second bounded
buffer beside the existing forward window, and the run arithmetic that aims it.

* `structural_runs(spans)` — groups per-member spans, streamed in increasing
  local-header order, into the runs one read each can deliver. Two neighbours
  join a run only when the earlier member's own span already reaches the later
  member's local header, which is what makes their union a range with **no
  hole**. Runs of fewer than two members are dropped: coalescing one member
  trades one request for one request and can only lose.
* `StructuralPrefetch` — holds the admitted runs, one positional read each, and
  answers a read only when one run holds **all** of it. A partly covered read, a
  zero-length read, and every read at all once the catalog is built take exactly
  the path they take today.
* Three named bounds:

| constant | value | why |
| --- | --- | --- |
| `MAX_STRUCTURAL_PREFETCH_RUN_BYTES` | 64 KiB | one run is one request, so this is the ceiling `MAX_SOURCE_READ_AHEAD_BYTES` already puts on a single speculative request. Measured: the longest run across the 533 ZIP containers under `test-data` is **9,298 bytes** — change 0577's own longest-run figure, reproduced. |
| `MAX_STRUCTURAL_PREFETCH_BYTES` | 256 KiB | unlike the forward window, the prefetch holds every admitted run until the catalog is read, because the relationship walk is a LIFO traversal that revisits runs in no order. Measured: the largest run set one package retains is **14,755 bytes**. |
| `MIN_STRUCTURAL_PREFETCH_RUN_MEMBERS` | 2 | below it there is nothing to coalesce. |

**`crates/litchi-opc/src/source_backed.rs`.** `SourceReader` carries an optional
`Arc<StructuralPrefetch>` and consults it before its existing dispatch.
`is_structural_member_name` names the members an open reads before it has parsed
anything — `[Content_Types].xml`, the package `_rels/.rels` and every
`*/_rels/*.rels` — **from the name alone**. `structural_prefetch_runs` streams
those members' span hints into `structural_runs` without buffering them. Both
source-backed catalog constructors prime the prefetch after the index is built
and release it the moment the catalog call returns, on every path.

**`crates/litchi-opc/src/pkgreader.rs`.** `CONTENT_TYPES_MEMBER` becomes
`pub(crate)`, so one constant names that member in both places.

Nothing else changed. No new `unsafe`, no new dependency, no weakened limit or
defence, no public API removed or altered, no change to any writer, publisher or
strict-layout path, and no existing test modified.

### Where it does *not* apply

The prefetch is the **default** open's mechanism and only the default open's.

* **A managed open** — one carrying an `ExecutionContext` — keeps the exact
  grammar. A read spanning several members is reserved and committed against
  `Resource::InputBytes` and observed for cancellation as a unit, so it would
  move where those observations land and how soon a finite input budget is
  exhausted. Change 0611 could accept that class of movement because its span
  read *was* the member read's first contact with the source, one for one; this
  read belongs to no member, so the same argument is not available. The gate is
  at the constructor and repeated inside `prime`, so the property is local. A
  test pins that a managed open of a package with an eleven-member run still
  costs one read per member.
* **An explicit forward window** — `SourceReadPolicy::forward_start(n)` — keeps
  its own grammar and its own diagnostics counters untouched. 0577 already
  showed candidate (c) strictly dominates candidate (d); running both at once
  would only make each harder to read.

## Why it is sound

**The run read fetches exactly the bytes the member reads fetch, and no others.**
This is not an argument but what `structural_runs` computes: two members join a
run only when the earlier one's own span already reaches the later one's local
header, so a run is the union of its members' own first-read spans with no hole
in it. Measured by logging every read of both legs on five real packages: the
**set of byte offsets the open touches is identical in both directions** — no
byte the after leg reads was left unread by the before leg, and none the before
leg read is skipped. That is change 0577's invariant 1, in the sharper form
change 0611's per-member span makes available, and it is why the byte totals
below are identical rather than 15% higher as 0577 modelled.

**Nothing is parsed to decide it.** The member list comes from member *names*
and from the central directory the index already holds. No member is read,
decompressed, classified or admitted to build it, which is what lets the
prefetch run before the catalog exists — and what keeps it from being able to
change the catalog.

**The buffer is a cache of source bytes and nothing else.** No value is trusted
because it came from it. Every local header is parsed, every size converted,
every CRC checked and every limit charged by exactly the code that does so when
each read goes to the source separately, in the same order, on the same bytes.
Only the *fetch* is reordered; the walk is not. That is invariant 3, and it is
what keeps the open's verdict a pure function of the package's bytes — the
property the `e4`/`e5` pair made decisive in 0577.

**Only a wholly covered read is answered.** A partly covered read would be
served in different-sized pieces than the source would have delivered, and
change 0611 *measured* that re-chunking moving which of two refusals a malformed
member reaches. The rule that removed the mechanism there removes it here.

**A zero-length read still reaches the source**, so a versioned or cancellable
adapter keeps exactly the chance to refuse it that it has today, and change
0317's fences and change 0600's monitored-read observations stay where they are.
Measured: source observations are unchanged on every one of the 533 containers.

**Intra-run error precedence, as 0577 froze it.** Within one run read, an I/O or
source-version failure at a later offset is reported before a parse, limit or
duplicate-ID error of an earlier member covered by the same read — the accepted
precedence change of changes 0565, 0566 and 0568. It is bounded in a way those
were not: because the prefetch is best effort, a failed run read is *not*
reported at all. It is abandoned, and the failure is observed again by whichever
member read actually needs those bytes, with that read's own identity. That is
invariant 5, and change 0611's reason for refusing invariant 5 one member wide
does not apply here: this read belongs to no member, so abandoning it cannot
give any member read a second refusal, a second cancellation observation or a
second reservation.

**ADR reading.** ADR 0005's bounded-resource clause is met by two named
ceilings, fallibly reserved buffers and a lifetime that ends with the catalog
call; its "cache behavior is semantically invisible" clause is met by the
measured verdict identity below and by the admission decision — the one 0577
showed has no later point to move to — being reached from the same bytes in the
same order. ADR 0006 is untouched: every check keeps its position and its
identity, and no output byte is produced on this path. ADR 0011 keeps its line:
`soapberry-zip` remains the ZIP grammar owner, and what crosses the boundary is
an offset and a length in the caller's own byte source — where bytes are, not
what they mean — which is the ownership-respecting addition 0577 said the design
needed. ADR 0010's unmeasured-cost posture is satisfied: the cost was measured
in 0572 and 0577 and the saving is measured here on a latency-bearing source,
which is the *after* half 0577 could not supply.

**Which contracts are untouched.** The strict target-scoped proof of
[0580](0580-zip-target-scoped-strict-layout.md) and its residual window
([0583](0583-zip-local-size-span-bound.md)) are on paths this change does not
enter; the three `stream_to` request counts are byte-identical. Change 0594's
one-decoder-per-open reuse and change 0611's per-member span are unchanged — the
prefetch sits *below* them, on the source the archive was built over.
`ZipOperationAccounting`'s counters are unchanged. Publication, splice and
artifact-restore paths all run after the prefetch has been released.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0.
Base `3156bff3b`, branch `perf/0623-zip-structural-span-accessor-and-prefetch`.
Every measured process pinned to CPU 17. Seven other agents were building and
measuring on the host throughout; every A/A floor was taken in its own window.

### Requests and bytes (deterministic counts, measured)

Change 0587's probe as change 0611 left it, re-run verbatim on both legs against
the same three fixtures (`results/change-0623/probe/`). The only edit is one
`classify` clause: a read that begins at a member's local header and covers a
second member's whole fixed header is a `run-span`. The before leg issues no
such read, so its output reproduces change 0611's retained after-counts exactly.

| scenario | before | after | requests | bytes |
| --- | --- | --- | ---: | ---: |
| open `ConditionalFormattingSamples.xlsx` (132 members, 42 structural) | 45 req / 26,066 B | **10 req / 26,066 B** | **−77.8%** | **0** |
| open `shapes.pptx` (48 members, 21 structural) | 24 / 11,191 | **6 / 11,191** | **−75.0%** | **0** |
| open `comment.docx` (10 members, 3 structural) | 6 / 1,690 | 6 / 1,690 | 0 | 0 |
| first read of `/xl/worksheets/sheet1.xml` | 1 / 1,305 | 1 / 1,305 | 0 | 0 |
| first read of `/ppt/slides/slide1.xml` | 1 / 4,807 | 1 / 4,807 | 0 | 0 |
| first read of `/word/document.xml` | 1 / 571 | 1 / 571 | 0 | 0 |
| all 90 XLSX parts, either order | 90 / 628,668 | 90 / 628,668 | 0 | 0 |
| all 27 PPTX parts, either order | 27 / 57,677 | 27 / 57,677 | 0 | 0 |
| all 7 DOCX parts, either order | 7 / 3,498 | 7 / 3,498 | 0 | 0 |
| `stream_to` of each of the three parts | 21 / 1,875, 16 / 5,227, 12 / 782 | identical | 0 | 0 |

0577 modelled **10** open requests on the workbook and 10 is what it costs. The
workbook's 42 structural members are served by four run reads and three
single-member reads; `shapes.pptx`'s 21 by two runs and one single member.

`comment.docx` does not move, and that is the degenerate case 0577's admission
gates required be shown to cost no more: its three structural members sit in
three separate runs of one, none of which is admitted.

**The byte columns are zero, not small.** A run read is the union of the spans
change 0611 already reads per member; where two adjacent members' spans overlap
by the 24 bytes of descriptor room, the union counts those bytes once instead of
twice, so the coalesced total is never larger and is sometimes smaller — see
`shape-soft-edges.pptx` below, 13,363 B → 12,859 B.

### The open differential over every ZIP container under `test-data`

533 containers — every file under `test-data` whose first two bytes are `PK`,
which is the same census change 0582's corpus generator takes. Both legs record,
per container: the open's verdict, its request, byte and source-observation
cost, every package and part relationship with its id, type, target and mode,
every admitted Part with its content type and relationship count, every non-part
member with its reason, and the decoded length and CRC-32 of every Part's bytes.
18,494 lines per report.

**The two reports differ on 316 lines, and every one of them is an
`open-cost requests=` line.** No verdict, no error identity, no part, no
relationship, no non-part member, no decoded payload and no source-observation
count differs anywhere in the corpus.

| | before | after |
| --- | ---: | ---: |
| open requests, 533 containers | 4,160 | **2,569** (−38.2%) |
| open requests, 321 OOXML-extension containers | 3,440 | **1,886** (−45.2%) |
| open bytes, 533 containers | 2,392,913 | **2,391,391** (−1,522) |
| containers costing **more** requests | — | **0** |
| containers costing fewer requests | — | 315 |
| containers reading fewer bytes | — | 16 |
| containers reading more bytes | — | **2** |
| `version()` observations, any container | — | **unchanged** |
| per-container request ratio, over the 315 that moved | — | min 0.188, median 0.667, max 0.958 |

Largest absolute savings: `bug62513.pptx` 52 → 12, `slide-section-test.pptx`
48 → 9, `ConditionalFormattingSamples.xlsx` 45 → 10, `45545_Comment.pptx`
45 → 13, `shape-soft-edges.pptx` 29 → 8, `shapes.pptx` 24 → 6.

**The two containers that read more bytes, named rather than averaged**, as
0577's byte-ceiling gate requires:

| container | requests | bytes | why |
| --- | --- | --- | --- |
| `poi/test-data/openxml4j/OPCCompliance_CoreProperties_OnlyOneCorePropertiesPartFAIL.docx` | 5 → 4 | 1,696 → 2,290 (**+594 B**) | the open *refuses*, part-way through a run the prefetch had already fetched |
| `poi/test-data/openxml4j/PackageRelsHasEntities.ooxml` | 5 → 5 | 2,484 → 2,816 (**+332 B**) | the same, with no request saved because the refusal comes before the run's second member |

Both refuse with the identical typed error before and after. That is the whole
of the byte cost across the corpus: a package whose open fails may have paid for
a run it did not finish reading.

### The read set, byte for byte

Every read of both legs was logged on five packages and the byte sets compared.

| container | before reads | after reads | bytes read only by the after leg | bytes read only by the before leg |
| --- | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xlsx` | 45 | 10 | **0** | **0** |
| `shapes.pptx` | 24 | 6 | **0** | **0** |
| `shape-soft-edges.pptx` | 29 | 8 | **0** | **0** |
| `comment.docx` | 6 | 6 | **0** | **0** |
| `poi/.../bug62513.pptx` | 52 | 12 | **0** | **0** |

### The run model

`results/change-0623/geometry/run_geometry.py` reimplements `local_span_hint`
and `structural_runs` from package bytes alone, with no help from the library,
and is what the two ceilings are set against.

| | 533 containers |
| --- | ---: |
| structural members | 2,560 |
| runs admitted | 542 |
| structural members covered by a run | 2,135 |
| bytes retained across the whole corpus | 1,029,321 |
| **largest run set one container retains** | **14,755 B** (`bug62513.pptx`) |
| **longest single run** | **9,298 B** (`ConditionalFormattingSamples.xlsx`) |
| containers whose longest run exceeds the 64 KiB clamp | **0** |
| containers whose run set exceeds the 256 KiB clamp | **0** |

The model predicts each container's request saving as `covered − runs`. It
reproduces the oracle's **measured** saving exactly on **530 of the 532**
containers both tools read. The two it misses are the two named above, where the
open refuses before reading every member of a run it had fetched — which is the
same fact from the other side.

### Paired timing

Both binaries `--release --locked` from the same sources with the same flags,
`taskset -c 17`, 30 samples and 5 warmups per case per leg, leg order
A1 B1 B2 A2 with an A/A floor A3 A4 in the same window
(`results/change-0623/timing/`, `results/change-0623/workbook-timing/`).

**The simulated latency-bearing transport of changes 0493 and 0572** — 1 ms of
fixed service per physical request, 100 MiB/s, 64 KiB maximum physical range.

*(a) The harness's own range-source selectors.* Physical requests per timed
iteration, from the simulator's counters, are deterministic and identical across
all 30 samples of every leg:

| case | requests before → after | bytes before → after |
| --- | --- | --- |
| `opc_range_source_open` | 5 → **4** | 1,011 → 1,011 |
| `opc_range_source_open_main_read` | 6 → **5** | 1,166 → 1,166 |
| `xlsx_range_source_open` | 7 → **6** | 1,899 → 1,899 |
| `xlsx_range_source_first_cell` | 10 → 10 | 666 → 666 |

| case | A1→B1 p50 | A2→B2 p50 | floor A3→A4 p50 | floor A1→A2 p50 |
| --- | ---: | ---: | ---: | ---: |
| `opc_range_source_open` | **−19.93%** | **−19.90%** | −0.01% | −0.04% |
| `opc_range_source_open_main_read` | **−16.59%** | **−16.56%** | −0.01% | −0.01% |
| `xlsx_range_source_open` | **−14.14%** | **−14.18%** | +0.02% | −0.03% |
| `xlsx_range_source_first_cell` | +0.04% | −0.00% | +0.02% | +0.02% |

These corpora are synthetic packages with exactly two structural members, so one
run of two is all there is to coalesce and the saving is one request — about
1.05 ms — in each case. The floor is under 0.05% at p50 in every direction.
`xlsx_range_source_first_cell` issues the same 10 requests on both legs; nothing
is expected or claimed there.

*(b) Real packages on the same transport*, because the harness's range corpora
cannot exercise a package with many relationship parts, and because 0577 set a
falsification criterion in exactly these terms
(`results/change-0623/workbook-timing/`):

| package | requests | before p50 | after p50 | A1→B1 p50 | A2→B2 p50 | saving | floor A3→A4 |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xlsx` | 45 → **10** | 48,208.6 µs | 11,292.7 µs | **−76.58%** | **−76.56%** | **−36.92 ms** | +0.00% |
| `shapes.pptx` | 24 → **6** | 25,610.3 | 6,621.1 | **−74.15%** | **−74.15%** | **−18.99 ms** | +0.01% |
| `shape-soft-edges.pptx` | 29 → **8** | 30,989.0 | 8,816.1 | **−71.55%** | **−71.55%** | **−22.17 ms** | −0.00% |
| `comment.docx` | 6 → 6 | 6,381.5 | 6,383.7 | +0.03% | +0.01% | +0.00 ms | −0.00% |

The two paired directions agree to within 0.02 percentage points on every row
and the floor is at most 0.01%, because the transport's service time dominates
and is deterministic. **0577's falsification criterion for ZIP-5 was "the
delayed-transport open saving is under 20 ms on the 132-member workbook". The
measured saving is 36.92 ms.**

**Local, in-process sources**, where a request costs about a microsecond
rather than a millisecond and no claim is made either way. The four-case local
window (`results/change-0623/timing/local-summary.txt`):

| case | A1→B1 p50 | A2→B2 p50 | floor A3→A4 p50 | floor A1→A2 p50 |
| --- | ---: | ---: | ---: | ---: |
| `opc_file_source_open` | +12.39% | +6.74% | −2.04% | **+7.07%** |
| `docx_file_source_open` | +3.28% | +4.91% | −0.94% | −0.78% |
| `pptx_file_source_open` | +4.70% | +1.62% | +0.31% | +0.80% |
| `docx_file_source_full_text` | −0.84% | +0.50% | −0.32% | −0.08% |

`opc_file_source_open` exceeds the programme's 5% review trigger in both
directions, so it is reported here rather than folded into a mean, and it was
followed up rather than explained away. A **third binary** was built from the
after tree with `is_structural_member_name` forced to `false`, so every line of
the mechanism is compiled and linked but no run is ever admitted, and the three
binaries were timed A N B B N A with an A/A floor in the same window
(`results/change-0623/timing/local-code-size-control.txt`):

| case | code size only, A→N | mechanism only, N→B | total, A→B | floor A3→A4 | floor A1→A2 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `pptx_file_source_open` | **+3.73% / +3.12%** | **−0.84% / −0.72%** | +2.86% / +2.38% | −0.82% | +0.31% |
| `opc_file_source_open` | +10.12% / −1.84% | −5.43% / +6.59% | +4.15% / +4.63% | **−7.41%** | **+5.92%** |

On `pptx_file_source_open`, the only local case whose A/A floor is under 1% at
p50 in both directions, the whole of the difference is present in the build
where the mechanism never fires, and the mechanism itself measures **−0.8% and
−0.7%** — the direction the counts predict, because it removes a `pread` and
adds a `memcpy`. What the local table is measuring is **`litchi-opc` growing by
about 490 lines of production code**, not the prefetch running. On
`opc_file_source_open` the A/A floor is −7.41% and +5.92%, larger than every
delta in its own row, so that case cannot resolve an effect of this size on this
host and nothing is read from it.

## Correctness evidence

**The oracle 0577's gates ask for, over a superset of the corpus they ask for.**
"Identical bytes, errors and verdicts for every member of every OOXML fixture"
is measured over all 533 ZIP containers under `test-data` — the OOXML packages,
the ODF packages, the iWork bundles and the bare `.zip` fixtures, opened
source-backed and compared field by field. The two reports differ on nothing but
the request count.

**The `e4`/`e5` pair is a regression test.**
`the_untyped_member_admission_verdict_does_not_move` builds the same untyped
member twice — once named by a deep relationship and once not — and pins that it
is tolerated archive junk in one package and a fatal `ContentTypeNotFound` in
the other, with the relationship part that decides it inside a coalesced run.

**Error identity at the start, the middle and the end of a run.** Three
fixtures each for a malformed relationship part and a duplicate relationship ID,
placed at positions 0, 3 and 7 of an eight-member run, all reporting the same
typed error; and five relationship-budget ceilings that refuse part-way through
the same run. Each is also run against a source that refuses the coalesced read,
so the verdict reached *through* the run read and the verdict reached through
the per-member fallback are compared directly, and are equal.

**Tests added.** `crates/litchi-opc/tests/structural_prefetch.rs`, 13 tests: one
read serves a contiguous run; a run read never covers an ordinary part's
payload; malformed, duplicate-ID and budget refusals anywhere in a run; the
`e4`/`e5` admission pair; a refused run read falling back to per-member reads
with an identical verdict; a scattered package costing no more reads; the
prefetch released when the catalog is built; a managed open keeping the exact
grammar; a forward window keeping its own counters; source observations
unchanged; and a guard that the fixtures really do coalesce, so the others are
not silently testing the fallback. Eight unit tests in
`source_backed/read_ahead.rs` for the run arithmetic: adjoining spans join, a
gap splits, a single-member run is not admitted, runs are disjoint and strictly
increasing, both ceilings, and out-of-order input admitting nothing. Three tests
in `crates/soapberry-zip/tests/member_span_read.rs` for the accessor: the hint
is the read the member actually issues, adjoining hints leave no hole, and a
member the window does not admit has no hint.

**The change 0582 read-grammar differential was not re-run, and here is why.**
That harness builds `soapberry-zip` against a slice source and exercises seven
per-member APIs; `litchi-opc` is not in its picture at all. The `soapberry-zip`
half of this change is a new `pub struct` and a new `pub fn` that no other code
in that crate calls, and `git diff` shows no existing line changed — so the read
grammar it would test is unchanged by construction, and re-running it would
prove that and nothing more. The grammar that *does* change is the sequence of
reads `litchi-opc` issues to a caller-supplied source, and the oracle above is
the differential for exactly that, over 533 containers rather than
`soapberry-zip`'s seven APIs.

**Gates.** Thirteen sections, all exit 0: formatting; Clippy for
`soapberry-zip`, for `litchi-opc` and for `litchi-opc --all-features`, with
workspace lints denied; rustdoc for both crates; a `--no-default-features`
check of `litchi-opc`; the `soapberry-zip` suite (613 passed, 2 ignored) and
the `litchi-opc` suite (707 passed, 1 ignored, and again under
`--all-features`); and the consumer suites
`litchi-xlsx`/`litchi-docx`/`litchi-pptx` (3,634 passed, 33 ignored), the
`litchi` facade and `litchi-ooxml-common`/`litchi-core` (477 passed). Every
gate tail is in `results/change-0623/gates.txt`. Beyond them, the 533-container
open differential and the three-fixture count probe above are themselves
reproducible checks and are retained with their runners.

## Validation preserved

No validation moved, weakened or changed identity. The relationship walk, the
orphan fallback loop, `classify_part_members`' untyped-member rule, every
`RelationshipLedger` charge and every `ReadLimits` ceiling run in the same order,
on the same bytes, with the same results — which the 533-container oracle
measures rather than asserts. No `unsafe` was added, no defence against
malformed input was relaxed, no limit was raised, and no typed refusal was
traded for a partial result. The one behavioural difference a caller can observe
is how many times its `read_at` is called, and in two of 533 packages, both of
which refuse, how many bytes it is asked for.

## Limitations

**No claim is registered.** `performance_claim: none`. The figures above are
evidence.

**The latency-bearing result is a simulated transport**, a fixed service time
per physical request, not a network or a real device. What it prices is the
request count, and the request count is measured deterministically. No cold page
cache, real device, peak-RSS, allocation-profile, instruction-count, syscall,
concurrency-scaling or cross-platform result is claimed.

**Nothing is claimed on a local, in-process source.** A warm local open of a synthetic
package costs a few microseconds more, and the controlled experiment above
attributes that to the compiled size of `litchi-opc` rather than to the
mechanism, which measures faster on the one local case with a floor tight
enough to read. That is a measurement on one host, one toolchain and one
allocator; it is reported because it was run, not claimed.

**Managed opens gain nothing**, by design and by the argument above, and no
managed measurement is offered. An explicitly configured forward read-ahead
window also gains nothing.

**The ceilings are measured against this repository's corpus.** 64 KiB per run
against a longest observed run of 9,298 bytes, and 256 KiB retained against a
largest observed set of 14,755 bytes. A producer that writes far larger
relationship parts, or far more of them contiguously, falls back to the
per-member grammar — never to an incorrect result.

**The member list is a name-only superset.** A relationship part the walk never
reaches is still fetched when it sits inside a run. Change 0577 measured that
across 167 OOXML packages every relationship part present is read at open, and
the corpus here has no container that costs an extra request; a package whose
*entire* contiguous run of relationship parts is unreachable would cost one.
Nothing in the corpus does.

**The 533-container run model is a model**, computed from package bytes by a
script that reimplements the accessor and the run arithmetic. It is validated
against the oracle's measured saving on 530 of the 532 containers both tools
read, and the two it misses are explained above; it is not a library
observation.

**The main part is not prefetched.** An XLSX or PPTX open reads
`xl/workbook.xml` or `ppt/presentation.xml` after the OPC catalog is built, by
the format crate, when the prefetch has already been released. Coalescing that
read with the structural run it usually adjoins would need format knowledge that
does not belong in `litchi-opc`, and is not attempted here.

## Retained evidence

Evidence packet, including both probes, the oracle, the run model, every timing
leg and the gate tails:
[`results/change-0623/README.md`](results/change-0623/README.md).

# 0611: one bounded positional read per ZIP member first-read — a source-backed OOXML open halves its requests, and halves its wall clock on a latency-bearing transport

Status: retained, implemented in `soapberry-zip`. `performance_claim: none` — no
claim-registry entry in this wave; the paired medians and the deterministic
counts below are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **ZIP-2** of
[0587](0587-remaining-opportunity-survey.md) (rank 14), which change
[0561](0561-opc-repeated-positional-reads.md) named as its target 1 and never
implemented. The frozen design, written before any code and amended once by the
differential, is
[`results/change-0611/design.md`](results/change-0611/design.md).

## What was changed

`crates/soapberry-zip` only, two files.

A member's *first* read on a positional source used to cost two or three
requests: the 30-byte fixed local header (`archive.rs`, `ZipArchive::get_entry`),
the compressed payload (`reader_at.rs`, `RangeReader::read` under `ZipReader`),
and a data descriptor when the central record declares one
(`DataDescriptor::parse_complete_at`, reached from `ZipVerifier::read`). The
bytes between the fixed header and the payload — the local name and the local
extra field — were skipped: their combined length was read out of the fixed
header and added to the payload offset.

It now costs one. `IndexedReadSession::read_entry_with_accounting`
(`office.rs`), the single production body behind `IndexedArchive::read`,
`read_entry`, `read_with_accounting`, `read_entry_with_accounting` and
`IndexedReadSession::read`, and therefore behind `litchi-opc`'s structural
member reads and its cold `PartView::data()` part loads, takes one bounded
positional read of the member's whole local record before the member is parsed,
and answers every read of that member from it.

* `SpannedSource<'archive, R>` — a `ReaderAt` that answers from one member's
  buffered local record and delegates every other read to the archive's own
  source. It is cloneable and carries `Option<Arc<MemberSpan>>`, so the
  passthrough form costs nothing. It never answers a zero-length read and never
  serves a payload it holds only part of.
* `IndexedArchive::member_span_length` — the window, `min(next local header +
  descriptor room, 30 + 610 + central compressed size + descriptor room,
  directory offset) - local header offset`, admitted only between 30 bytes and
  `MAX_MEMBER_SPAN_READ_BYTES`. `next_local_header_offset` is a binary search
  over the offset-sorted layout the index already holds.
* `IndexedArchive::member_source` — reserves the buffer fallibly, issues the
  read, and hands back the passthrough reader for a member the window does not
  admit or a buffer it cannot reserve.
* `ZipArchive::get_entry_from`, `ZipEntry::reader_over`,
  `ZipEntry::verifying_reader_over` — `pub(crate)` siblings of `get_entry`,
  `reader` and `verifying_reader` that take the reader to read from.
  `get_entry`, `reader` and `verifying_reader` keep their signatures and
  delegate.
* `IndexedReadSession`'s decoder becomes
  `DeflateDecoder<CountingReader<ZipReader<SpannedSource<'a, R>>>>`, so change
  [0594](0594-zip-session-reuse-per-open.md)'s one-decoder-per-open reuse is
  preserved exactly.

Two named constants carry the bounds:

| constant | value | why |
| --- | --- | --- |
| `MAX_MEMBER_SPAN_READ_BYTES` | 64 KiB | measured: of the 7,757 members in the 335 OOXML containers under `test-data`, `30 + local variable region + compressed + 24` is ≤ 4 KiB for 7,493, ≤ 32 KiB for 7,695 and ≤ 64 KiB for 7,732 (99.68%). It is also the largest single speculative window `litchi-opc` already admits (`MAX_SOURCE_READ_AHEAD_BYTES`), which [0577](0577-ooxml-open-relationship-parts.md) proposed borrowing. |
| `MEMBER_SPAN_METADATA_ALLOWANCE` | 640 | change [0573](0573-zip-single-local-header-read.md)'s measured window: 30 fixed bytes plus 610, which covers every OOXML member's local variable region in this repository's corpus, where a 512-byte allowance misses 301 of 4,270. |

Nothing else changed. No public API, no new `unsafe`, no new dependency, no
weakened limit or defence, no change to `litchi-opc`'s production code. The
strict-layout paths — `read_to`, `read_entry_to`, `stream_to`,
`with_verified_entry_reader*` and `capture_precompressed` — are untouched, which
the `stream_to` counts below confirm byte for byte.

## Why it is sound

**The buffer is a cache of source bytes and nothing else.** No value is trusted
because it came from the buffer. The fixed local header is parsed by exactly the
code that parses it today; the local variable length comes from that parse, not
from the central record; `local_end > directory_offset`, the ZIP64 size framing,
`body_end_offset > directory_offset`, the size conversions, the CRC and the
declared-size check all run unchanged, in the same order, on the same values.

**The buffer answers only reads it can answer in full.** A read the buffer does
not wholly hold reaches the source at exactly the offset and length it reaches
today, and the payload is checked for whole coverage against the member's own
`compressed_data_range` before the decoder is built. This is the one amendment
the frozen design took, and change 0582's differential is why: see "Correctness
evidence". Serving a partly-covered payload as a short read re-chunks the
decoder's input, and `ZipVerifier::read` completes a member on a per-read size
test, so on a member whose declared size *and* declared CRC are both wrong that
re-chunking moved which of two refusals fired. Full coverage or nothing removes
the mechanism rather than the symptom.

**One member read stays one refusal.** The span read is the member read's first
contact with the source and begins at the same offset as today's first read, so
a source failure is reported rather than retried behind the caller's back. This
is deliberately *not* invariant 5 of 0577's design ("a run read that fails falls
back to per-member reads"): that invariant is right for a read spanning several
members and wrong here, because a retry would give one member read two
refusals, two cancellation observations and two resource reservations.
`litchi-opc`'s `managed_input_budget_refusal_counts_only_the_terminal_read_reservation`
pins exactly that, and it passes unchanged.

**A short read is not a failure.** The read takes what the source returns;
whatever it did not cover is read exactly as before, so a truncated source still
fails at the byte, and with the message, it fails at today.

**A zero-length read still reaches the source.** A read that fetches nothing has
nothing for the buffer to answer with, and a versioned or cancellable adapter
gets exactly the chance to refuse it that it has today. `RangeReader` issues one
at the end of every member, so this is not a corner: it keeps change 0317's
fences and change 0600's monitored-read observations where they are.

**Error precedence, ADR reading.** Change
[0317](changes/0317-opc-source-read-error-precedence.md)'s brackets are
unchanged: `litchi-opc` fences a member read with a source-version observation
and an execution-context check on both sides, and this change is entirely inside
those brackets; source-version failure still precedes execution-context failure,
which still precedes the mapped ZIP member error. ADR 0005's bounded-resource
clause is met by a named ceiling and a fallibly reserved buffer that lives only
for the member read; ADR 0006 preservation is untouched because no output byte
is produced on this path; ADR 0011's ownership line is untouched because no
archive type crosses a crate boundary — the window is computed inside
`soapberry-zip` from facts the index already holds, which is why `litchi-opc`
needed no change.

**Which contracts are untouched.** The strict target-scoped proof of
[0580](0580-zip-target-scoped-strict-layout.md) and its residual window
([0583](0583-zip-local-size-span-bound.md)) are on paths this change does not
enter, so 0573's `the_speculative_window_never_reads_this_members_payload` and
0580's "a neighbour probe reads exactly 30 bytes … and never a variable region
or a payload byte" hold verbatim. The slice-backed `ArchiveReader` path is
unchanged (a slice source issues no requests). `ZipOperationAccounting`'s
payload counters are unchanged, because the payload still flows through the same
`CountingReader<ZipReader<..>>`.

**The one place a refusal can move.** Under a caller-supplied `ExecutionContext`
with a finite `Resource::InputBytes` limit, `litchi-opc` commits the bytes
actually read against that budget (`source_backed.rs`, `read_source_at_with_context`).
A spanned first read commits the member's local variable region as well — the
framing bytes today's two reads skip. The increase is bounded by 610 bytes per
spanned member first-read and measured at 84 bytes per structural member on the
132-member workbook. No error identity changes and no new refusal exists; what
moves is only how soon a finite input budget can be exhausted. This is the same
class and direction as the speculative window 0573 already landed on the strict
path, which paid +0.54% bytes for −47.3% reads under the same budget. Ordinary
opens pass no `ExecutionContext`, so this applies to managed packages only.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0.
Base `2d6fbeaed`, branch `perf/0611-zip-single-read-per-member`. Every measured
process pinned to CPU 31. Seven other agents were building and measuring on the
host throughout; the A/A floor was taken in the same window.

### Requests and bytes (deterministic counts, measured)

Change 0587's probe, re-run verbatim on both legs against the same three
fixtures (`results/change-0611/probe/`; the only edit is that `classify` now
names a read that starts at a member's local header and is longer than 640 bytes
`member-span`, which the before leg never issues, so its output reproduces
0587's retained counts exactly).

| scenario | before | after | requests | bytes |
| --- | --- | --- | ---: | ---: |
| open `ConditionalFormattingSamples.xlsx` (132 members, 42 structural) | 87 req / 22,546 B | **45 req / 26,066 B** | **−48.3%** | +15.6% |
| open `shapes.pptx` (48 members, 21 structural) | 45 / 9,104 | **24 / 11,191** | **−46.7%** | +22.9% |
| open `comment.docx` (10 members, 3 structural, descriptor-bearing) | 12 / 1,584 | **6 / 1,690** | **−50.0%** | +6.7% |
| first read of `/xl/worksheets/sheet1.xml` | 2 / 1,281 | **1 / 1,305** | **−50.0%** | +1.9% |
| first read of `/ppt/slides/slide1.xml` | 2 / 4,786 | **1 / 4,807** | **−50.0%** | +0.4% |
| first read of `/word/document.xml` (descriptor) | 3 / 530 | **1 / 571** | **−66.7%** | +7.7% |
| all 90 XLSX parts, physical order | 208 / 625,408 | **90 / 628,668** | **−56.7%** | +0.5% |
| all 90 XLSX parts, reverse order | 208 / 625,408 | **90 / 628,668** | **−56.7%** | +0.5% |
| all 27 PPTX parts, either order | 56 / 56,456 | **27 / 57,677** | **−51.8%** | +2.2% |
| all 7 DOCX parts, either order | 21 / 3,213 | **7 / 3,498** | **−66.7%** | +8.9% |
| `stream_to` of `/xl/worksheets/sheet1.xml` | 21 / 1,875 | **21 / 1,875** | 0 | 0 |
| `stream_to` of `/ppt/slides/slide1.xml` | 16 / 5,227 | **16 / 5,227** | 0 | 0 |
| `stream_to` of `/word/document.xml` | 12 / 782 | **12 / 782** | 0 | 0 |

0587 modelled open 87 → 45, 45 → 24, 12 → 6 and a part 2-3 → 1. All four are
met exactly. Forward and reverse traversal are identical on both legs, and the
three `stream_to` rows are byte-identical, which is the count-level evidence
that the strict path is untouched.

The extra bytes are the local variable region, the framing the two-read grammar
skipped: 3,520 B over 42 structural members on the workbook, 84 B per member.

Source observations (`version()` calls) are **unchanged in every row** — 4 per
open, 4 per first part read, 364 for the 90-part traversal — so nothing about
change 0317's fences or change 0600's monitored-read scopes moves.

### Allocations (deterministic counts, measured)

| scenario | before | after | delta |
| --- | --- | --- | --- |
| open `ConditionalFormattingSamples.xlsx` | 5,250 allocs / 2,284,294 B | 5,333 / 2,299,511 | +83 (+1.6%) / +15,217 (+0.7%) |
| open `shapes.pptx` | 1,951 / 589,982 | 1,992 / 596,892 | +41 (+2.1%) / +6,910 (+1.2%) |
| open `comment.docx` | 397 / 430,098 | 402 / 430,844 | +5 (+1.3%) / +746 (+0.2%) |
| first read of `sheet1.xml` | 27 / 86,291 | 28 / 87,640 | +1 / +1,349 |
| first read of `document.xml` | 27 / 83,091 | 24 / 83,337 | −3 / +246 |
| all 90 XLSX parts | 5,809 / 8,684,189 | 6,071 / 9,326,251 | +262 (+4.5%) / +642,062 (+7.4%) |

One fallibly reserved buffer per spanned member read, sized by that member's
local record and released with it. The 90-part row is the honest worst case in
this packet: +642 KB of transient allocation to deliver 620,758 B of member
spans.

### Paired timing

Both binaries `--release --locked` from the same sources with the same flags,
`taskset -c 31`, 30 samples and 5 warmups per case per leg, leg order
A1 B1 B2 A2 with an A/A floor A3 A4 taken in the same window
(`results/change-0611/timing/`).

**The simulated latency-bearing transport** — change 0493's and change 0572's
configuration: 1 ms of fixed service per physical request, 100 MiB/s, 64 KiB
maximum physical range.

Physical requests per timed iteration, from the simulator's own counters, are
deterministic and identical across all 30 samples of every leg:

| case | requests before → after | bytes before → after |
| --- | --- | --- |
| `opc_range_source_open` | 9 → **5** | 933 → 1,011 |
| `opc_range_source_open_main_read` | 12 → **6** | 1,039 → 1,166 |
| `xlsx_range_source_open` | 15 → **7** | 1,732 → 1,899 |
| `xlsx_range_source_first_cell` | 10 → **10** | 666 → 666 |

| case | leg | p50 | mean | p95 | p99 |
| --- | --- | ---: | ---: | ---: | ---: |
| `opc_range_source_open` | A1 | 9,530.6 µs | 9,649.7 | 10,713.3 | 11,615.0 |
| | B1 | 5,305.4 | 5,300.0 | 5,314.3 | 5,351.7 |
| | B2 | 5,303.8 | 5,296.8 | 5,309.5 | 5,315.1 |
| | A2 | 9,524.9 | 9,511.0 | 9,531.0 | 9,540.7 |
| `opc_range_source_open_main_read` | A1 | 12,695.6 | 12,723.7 | 12,717.4 | 13,998.7 |
| | B1 | 6,361.4 | 6,353.0 | 6,369.7 | 6,369.8 |
| | B2 | 6,359.8 | 6,352.3 | 6,366.5 | 6,368.1 |
| | A2 | 12,698.3 | 12,715.0 | 12,745.5 | 13,452.6 |
| `xlsx_range_source_open` | A1 | 15,869.2 | 15,871.9 | 15,918.1 | 15,959.3 |
| | B1 | 7,448.2 | 7,437.6 | 7,456.1 | 7,463.9 |
| | B2 | 7,450.7 | 7,447.4 | 7,459.8 | 7,473.4 |
| | A2 | 15,863.2 | 15,866.2 | 15,897.8 | 15,902.2 |
| `xlsx_range_source_first_cell` | A1 | 10,665.3 | 10,666.8 | 10,679.2 | 10,790.5 |
| | B1 | 10,664.0 | 10,683.4 | 10,676.4 | 11,363.0 |
| | B2 | 10,662.9 | 10,657.8 | 10,674.2 | 10,684.4 |
| | A2 | 10,666.3 | 10,662.7 | 10,676.7 | 10,676.8 |

Paired deltas in both directions, and the A/A floor taken in the same window:

| case | A1→B1 p50 | A2→B2 p50 | A1→B1 mean | A2→B2 mean | floor A3→A4 p50 | floor A1→A2 p50 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `opc_range_source_open` | **−44.33%** | **−44.32%** | −45.08% | −44.31% | +0.02% | −0.06% |
| `opc_range_source_open_main_read` | **−49.89%** | **−49.92%** | −50.07% | −50.04% | −0.01% | +0.02% |
| `xlsx_range_source_open` | **−53.07%** | **−53.03%** | −53.14% | −53.06% | +0.02% | −0.04% |
| `xlsx_range_source_first_cell` | −0.01% | −0.03% | +0.16% | −0.05% | −0.06% | +0.01% |

The floor is under 0.1% at p50 on every case in both directions, because the
transport's fixed service time dominates and is deterministic; the three
changed cases are two to three orders of magnitude outside it, and their two
paired directions agree to within 0.05 percentage points. The saving is
**1.05 ms per removed request** on a 1 ms transport — 4.22 ms for 4 requests,
6.34 ms for 6, 8.42 ms for 8 — which is the mechanism and not a coincidence of
the corpus. Tails: the A1 legs of the first two cases carry a p95/p99 outlier
from other work on the host, so their A1→B1 tail deltas are not read; the
A2→B2 direction, whose floor is clean, is.

`xlsx_range_source_first_cell` issues the same 10 physical requests on both
legs. Its timed interval is a selected-cell read that does not take the
ordinary indexed first-read path, so no change is expected there and none is
claimed.

**Local, in-process sources**, where no change is expected and none is claimed.

| case | leg | p50 | mean | p95 | p99 |
| --- | --- | ---: | ---: | ---: | ---: |
| `opc_file_source_open` | A1 | 184.1 µs | 188.4 | 211.4 | 221.8 |
| | B1 | 170.8 | 171.0 | 186.2 | 186.2 |
| | B2 | 180.7 | 182.5 | 195.0 | 195.2 |
| | A2 | 205.6 | 207.1 | 223.2 | 225.1 |
| `docx_file_source_open` | A1 | 295.1 | 300.4 | 334.2 | 334.8 |
| | B1 | 302.9 | 303.8 | 335.9 | 353.2 |
| | B2 | 284.8 | 286.9 | 323.2 | 323.9 |
| | A2 | 324.8 | 336.9 | 361.4 | 607.6 |
| `docx_file_source_full_text` | A1 | 208.7 | 210.3 | 236.3 | 241.6 |
| | B1 | 196.4 | 200.5 | 215.5 | 224.9 |
| | B2 | 201.8 | 203.6 | 217.3 | 221.3 |
| | A2 | 215.7 | 217.8 | 248.1 | 252.5 |
| `pptx_file_source_open` | A1 | 2,751.7 | 2,773.8 | 2,921.4 | 2,964.5 |
| | B1 | 2,724.3 | 2,732.1 | 2,807.6 | 2,821.7 |
| | B2 | 2,658.0 | 2,673.9 | 2,767.1 | 2,838.3 |
| | A2 | 2,815.1 | 2,833.6 | 3,006.0 | 3,030.1 |

| case | A1→B1 p50 | A2→B2 p50 | floor A3→A4 p50 | floor A1→A2 p50 |
| --- | ---: | ---: | ---: | ---: |
| `opc_file_source_open` | −7.21% | −12.08% | +2.25% | +11.67% |
| `docx_file_source_open` | +2.67% | −12.30% | −3.86% | +10.06% |
| `docx_file_source_full_text` | −5.91% | −6.42% | +3.18% | +3.33% |
| `pptx_file_source_open` | −1.00% | −5.58% | −0.28% | +2.30% |

**Nothing is claimed from this table, and it is reported because it was run.**
This window was contended: the gate run was compiling and testing the consumer
crates on the same host while it ran, and its A/A floor says so — 2.3% to 11.7%
at p50, against 0.06% in the range window. Every case's two paired directions
disagree in magnitude, one disagrees in sign, and no delta is outside its own
floor. What the table does establish is the absence of a regression: no p50 in
either direction is more than 2.67% adverse, and that one sits inside a +10.06%
floor. A member's payload is memory-to-memory on these sources, so removing a
positional request removes a `memcpy` and adds one, which is what the design
predicted and what the counts already showed as +0.2% to +1.2% allocation
bytes.

## Correctness evidence

### Change 0582's differential, extended and re-run in full

0582's harness exercises only the strict-layout APIs per member and never calls
`IndexedArchive::read_entry`, the path this change modifies, so re-running it
unchanged would have proved nothing. It was extended with two per-member verdict
APIs — `I.read_entry` (the changed path) and `R.read` (the slice-backed control)
— keeping every existing API, the verdict line format, the corpus generator and
the classifier's classes. The harness, the classifier and the runner are
retained in `results/change-0611/`.

Both builds, `git archive 2d6fbeaed` twice with `archive.rs` and `office.rs`
overlaid on the after tree, the full regenerated corpus of 22,875 archives
(334,625,602 bytes: 516 `test-data` containers, 7 fuzz seeds, 45 crafted, 22,307
mutations), every member, both directions and both limit profiles.

**2,886,786 member verdicts per API pair; 481,131 verdict lines per member API
per report.**

The first run found **two** class-E divergences — same refusal on both legs,
different typed identity — on one crafted mutation,
`mutations/crafted-many-tiny-one-reaching/cdh-crc-e0-ffffffff.zip`, member
`m000.bin`, under both profiles: `InvalidSize { expected: 4, actual: 5 }` before
and `InvalidChecksum { expected: 4294967295, actual: 211534962 }` after. That
member's central record declares `compressed_size = 4096` while its local header
declares 4, so the window stopped 42 bytes in, inside a payload the reader asked
4,096 bytes of; the partly-covered read was served short, which re-chunked the
decoder's input, which moved which of two refusals `ZipVerifier` reached first.
The design was amended — the buffer serves the payload only when it holds all of
it — and the harness re-run:

```json
{"class_A_inputs": 0, "class_B_inputs": 0,
 "class_E_distinct_pairs": 0, "class_E_pairs": {},
 "crafted_divergent": {}, "divergences": {},
 "oracle_failure_count": 0, "oracle_failures": [], "panics": [],
 "per_api": {}, "per_family": {}}
```

**The two reports are byte-identical** — `cmp` clean over 3,801,502 lines and
295,172,990 bytes each, with the same footer on both: `TOTAL inputs=22875
panics=0 fuzz[accept=1085207 overlap=147416 other=208259] wide[accept=1089102
overlap=148352 other=208450]`, and zero `PANIC` lines. That is stronger than
"no divergence classified": across the whole corpus the after tree reproduces
the before tree's read grammar exactly, for all seven per-member APIs, both
directions and both limit profiles. The before binary's sha256 reproduced
across the runs, so the control leg is provably the same code throughout.

Members examined: 2,952 crafted, 1,356,678 and 1,359,924 mutation members under
the two profiles, 114 seed, 81,138 and 82,914 `test-data` members.

### Tests added

`crates/soapberry-zip/tests/member_span_read.rs`, 13 tests: one request per
member first-read for Deflate, Store and descriptor-bearing members and for
every member of one archive; order independence across forward, reverse and
interleaved traversal and between a shared session and fresh reads; a member
above the ceiling keeping the 30-byte first read and the historical grammar; a
local variable region above the allowance still reading; a failing source read
reported once, not retried; a short-reading source served from what arrived; the
window never reaching past the central directory; one member read never fetching
more than its own local record plus the allowance and the widest descriptor; and
a payload the window cannot cover being fetched from the source.

### Test harnesses corrected, with their assertions unchanged

Sixteen tests in four files drive a source that fires — sleeps, cancels, bumps
its version, parks on a barrier, refuses the read, or counts it — when a read
*begins exactly at the member payload's offset*, and assert a source-change,
cancellation, flight, reservation or read-count contract around it. A member's
payload no longer has a read of its own, so fifteen failed and one deadlocked.
The contracts are untouched; only the trigger encoded the old grammar.

| file | tests |
| --- | ---: |
| `crates/litchi-opc/src/source_backed.rs` (`#[cfg(test)]`) | 10 |
| `crates/litchi-opc/tests/source_backed_batch.rs` | 2 |
| `crates/litchi-xlsx/tests/source_backed_cell_values.rs` | 3 |
| `crates/litchi-xlsx/tests/source_backed_row_visibility.rs` | 1 |

Each source now fires when a read **delivers the payload's first byte and began
inside the member's own local record** — `requested > 0 && offset <= payload &&
payload - offset < requested && payload - offset <= 640`. Both halves are load
bearing, and finding that out took three attempts:

* Matching "begins at the payload **or** at the member's local header" fires
  twice per member in a gapless archive, because the zero-length read a range
  reader issues at the end of member *n* lands on member *n+1*'s header. The
  non-empty requirement is what excludes it.
* Matching the local header alone fires twice during publication, because the
  preservation path issues a short local-record probe at the same offset.
  Requiring the payload byte to be delivered is what distinguishes the member
  read from that probe.
* Matching "delivers the payload byte" alone is tripped by a bulk publication
  copy that spans the member from 64 KiB away, which turned an exact-no-op
  publication into a refusal. The 640-byte prefix bound — change 0573's measured
  widest local record — is what keeps a fetch of the member distinct from a copy
  over it.

The result fires exactly once per member first-read under either grammar: the
one bounded read, or the payload read of the two-read form. Every assertion in
all sixteen tests is unchanged.

### Gates

`results/change-0611/gates.txt`. All green:

| section | result |
| --- | --- |
| `cargo fmt --all --check` | exit 0 |
| `cargo clippy -p soapberry-zip --all-targets --locked` | exit 0 |
| `cargo clippy -p litchi-opc --all-targets --locked` | exit 0 |
| `cargo test -p soapberry-zip --locked` | exit 0 — 610 passed, 0 failed, 2 ignored, 12 binaries |
| `cargo test -p litchi-opc --locked` | exit 0 — 679 passed, 0 failed, 1 ignored, 24 binaries |
| `cargo doc -p soapberry-zip --no-deps --locked` | exit 0 |
| `cargo doc -p litchi-opc --no-deps --locked` | exit 0 |
| `cargo test -p litchi-xlsx -p litchi-docx -p litchi-pptx --locked` | exit 0 — 3,629 passed, 0 failed, 33 ignored, 186 binaries |
| `cargo test -p litchi-xlsb -p litchi-core --locked` | exit 0 — 946 passed, 0 failed, 11 ignored, 20 binaries |
| `cargo test -p litchi-odt -p litchi-odc -p litchi-odf-common --locked` | exit 0 — 1,544 passed, 0 failed, 1 ignored, 90 binaries |

Workspace lints are `deny`, so every clippy and rustdoc section exiting 0 means
no warning at all. 7,408 tests, no pre-existing failure to record.

## Validation preserved

Every ZIP validation runs on the same values in the same order: the local
signature, the `local_end > directory_offset` and `body_end_offset >
directory_offset` bounds, the ZIP64 size framing through the local extra field,
the declared-size check and the CRC, and the data-descriptor comparison against
the central record. No limit was widened: `ArchiveLimits` is untouched and the
new ceiling only *narrows* what may be read speculatively. The strict
target-scoped proof, its neighbour probe and its residual window are not on this
path and are unchanged, as the identical `stream_to` counts show. No refusal was
turned into an acceptance anywhere in 2,886,786 member verdicts, and no accepted
member's bytes changed.

## Limitations

* **Not claimed:** any speedup on a local or in-memory source. The local
  selectors are reported for completeness and are inside the A/A floor.
* **Not claimed:** a claim-registry entry. `performance_claim: none`.
* The transport result is a *simulated* transport — a deterministic fixed
  service time per physical request, not a network or filesystem measurement.
  What it prices is the request count, which is measured deterministically.
* `xlsx_range_source_first_cell` shows no change and none is claimed: its timed
  interval is dominated by the strict-proof reads this change does not touch.
* The corpus is this repository's. The ceiling and the allowance are measured
  against 7,757 OOXML members here; a producer writing larger local extra fields
  or larger members would fall back more often, never incorrectly.
* Allocation *bytes* rise by up to 7.4% on a whole-package traversal. Peak RSS
  was not measured.
* Cold page cache, real devices, non-glibc allocators, non-Linux platforms and
  concurrency scaling were not measured.
* The extended differential's `members_reaching_proof` field now counts the two
  added APIs, which never enter the proof, so that field is no longer comparable
  with 0582's and 0583's retained values. The per-API verdict counts are.

## Retained evidence

[`results/change-0611/README.md`](results/change-0611/README.md).

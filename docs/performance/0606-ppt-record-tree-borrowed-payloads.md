# 0606: the PPT record tree spans the stream it already retains instead of copying every payload once per nesting level

Status: retained. `performance_claim: none` — this record carries deterministic
counts (allocations, retained bytes, `memcpy` calls, callgrind instruction
differentials), native instruction and page-fault counts, and paired timings
reported beside the host's A/A floor. No claim is registered.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements items **PPT-1** (rank 20) and **PPT-2** of change
[0587](0587-remaining-opportunity-survey.md), section "DOC and PPT", opportunity
R5 and R6. Change [0301](changes/0301-ppt-selected-slide-decoding.md) explicitly
left "the complete `Document` … parser" outside its scope; this record is that
parser. Changes [0116](changes/0116-ppt-source-backed-lazy-pictures.md) and
[0117](changes/0117-ppt-pictures-release-evidence.md) made `Pictures` lazy and
are untouched. 0199/0200 did the analogous elision for ODS parse events.

## The mechanism 0587 found

`Record::parse_impl` (`crates/litchi-ppt/src/records/record.rs`) copied each
record's declared payload into an owned `Vec<u8>` and then parsed that record's
children out of the same bytes. A container therefore held a private copy of
every byte its children also copied, once per nesting level. 0587 measured the
resulting copy factor — owned payload bytes over stream bytes — on three
fixtures (1.32, 1.63, 2.01). This record measures all 30 `.ppt` fixtures
(`results/change-0606/record-tree-census.txt`): the factor runs from 1.00 (the
four encrypted files, which parse as one opaque record) to **2.35**, and is
2.0 or more on 20 of the 26 files that parse. The retained `Presentation` keeps
the whole `PowerPoint Document` stream as well, so retention was 2.3-3x the
stream.

0587 also found that `parse_document_with_limits` ended with
`extract_slide_text_from_document`, an eager full-text extraction stored in
`slide_atoms_sets` and read only by `RecordParser::slides()` and
`slide_count()`. A workspace grep finds exactly one caller of either: a unit
test in `parsers/parser.rs`.

## What was changed

Scope: `crates/litchi-ppt` only. No other crate's files are touched.

**`records/payload.rs` (new).** `RecordPayload` is the type of `Record::data`.
It is `Storage::Owned(Vec<u8>)` or `Storage::Shared { buffer: Arc<Vec<u8>>,
start, end }`, 32 bytes either way. It `Deref`s to `[u8]`, compares by value,
and its `Debug` forwards to the slice, so the two forms are indistinguishable to
a reader; `From<Vec<u8>>`, `From<&[u8]>`, `From<[u8; N]>` and `FromIterator<u8>`
construct owned payloads, and `Deref`/`DerefMut`, `AsRef`, `Borrow` and
`IntoIterator` cover every read site. `to_mut()` materializes an owned buffer
first and follows `std::borrow::Cow::to_mut` exactly, including its re-match;
`DerefMut` routes through it, so an in-place edit of a borrowed payload copies
on demand and never writes through to the shared stream. A span outside its
buffer reads as empty rather than panicking.

**`records/record.rs`.** `parse_impl` takes a `PayloadStore` (`Owned(&[u8])` or
`Shared(&Arc<Vec<u8>>)`) and an explicit `window_end` instead of a re-sliced
`&[u8]`. Children are parsed from `header_end..record_end` of the same buffer
rather than from a fresh sub-slice, which is the same window the old code
created; every `data.len()` that bounded a child's available bytes became the
window end, so the lenient truncation and byte-at-a-time resynchronization are
byte-for-byte the same decisions. `Record::parse_shared_with_budget` is the new
entry point; every existing `&[u8]` entry point (`parse`, `parse_with_limits`,
`parse_strict*`, `parse_sequence_strict*`, `RecordParseSession`) keeps copying
its payloads and is unchanged in behaviour.

**`parsers/parser.rs`.** `parse_shared_document_with_limits` and (under the
`encryption` feature) `parse_shared_document_at_offsets_with_limits` parse from
an `Arc<Vec<u8>>`; the lenient top-level loop is factored into `parse_top_level`
so both payload sources run the identical loop. `slide_atoms_sets` becomes a
`std::sync::OnceLock<Vec<Vec<u8>>>` filled by `slides()` on first use instead of
by every parse. `OnceLock`, not `OnceCell`, so `RecordParser` and `Presentation`
keep whatever `Send`/`Sync` they had.

**`presentation/model.rs`, `presentation/package.rs`.**
`Presentation::powerpoint_document` becomes `Arc<Vec<u8>>` — the same `Vec` the
stream reader produced, wrapped without copying — and the record tree is parsed
from it. The decryption path still decrypts the `Vec` in place before it is
shared, so the `Arc` is only ever created over final bytes.

**Call sites.** 84 files change; all but the six above are one-line
conversions where a `Vec<u8>` is stored into `Record::data` (`x` → `x.into()`)
or read out of it as an owned vector (`.clone()` → `.to_vec()`), plus five
in-place editors that now write through `data.to_mut()` or `DerefMut`.

**Two public API changes inside `litchi-ppt`**, both authorized by the brief:
`Record::data` is `RecordPayload` rather than `Vec<u8>` (`RecordPayload` is
re-exported from the crate root), and `master::model::Element::bytes` loses its
`const` because `RecordPayload::as_slice` cannot be `const`. `RecordLimits`'s
`max_copied_payload_bytes` doc comment now says what the field has always
bounded.

## Why it is sound

**Every payload value is identical.** A borrowed payload is the same
`data[header_end..record_end]` the old code copied, taken from the same buffer
at the same offsets. `PartialEq`, `Debug` and every read path go through
`as_slice()`, so no reader can tell the forms apart. A unit test asserts that
an owned and a shared parse of the same bytes produce equal trees, equal
`extract_all_text` and equal `slides()` on four shapes, including a container
whose declared length runs past the stream and a stream with leading garbage
that forces resynchronization.

**Error identity is unchanged, including the copy budget.** `charge_copy` is
still called for every record with the same byte count whether or not the
payload is copied, so `max_copied_payload_bytes` refuses exactly the files it
refused before. A unit test parses the same tree under a budget equal to the
tree's payload bytes and one byte below it, and asserts both payload sources
accept and refuse identically. The only `Error` variant that becomes
unreachable on the shared path is `AllocationFailed("PPT record payload")`,
which fires only when the allocator refuses a payload the shared path does not
allocate.

**Deferring the slide-text extraction defers no refusal.**
`RecordParser::extract_all_text` has no `?` and no failing branch: it returns
`Ok` on every input. The value `slides()` returns is what the eager field held,
computed from the same records; `slide_count()` is `slides().len()` as before.
Nothing outside `parsers/parser.rs` calls either.

**Mutation cannot reach the stream.** The shared buffer is behind `Arc<Vec<u8>>`
and is never handed out mutably. `to_mut()` — and `DerefMut`, which calls it —
replaces a shared span with an owned copy before returning a `&mut`, so the five
in-place editors (`slide_order`, `document_structure/transaction`,
`font/transaction`) keep their exact byte semantics. A unit test edits one
record of a shared tree and asserts the stream and the sibling records are
unchanged. In practice those editors parse through the `&[u8]` entry points, so
their payloads are already owned and `to_mut()` is a no-op there.

**Validation does not mutate and no limit moved.** `RecordLimits` values,
`ParseBudget` charging order, the depth and record-count ceilings, the strict
and lenient container loops, and the truncated-container recovery are unchanged.
The four encrypted fixtures still refuse with the same errors.

**No contract moved.** No new `unsafe` (`litchi-ppt` keeps
`#![forbid(unsafe_code)]`), no weakened limit or malformed-input defence, no
hidden global pool, no ambient I/O, no public leakage of archive types, raw
locks or executors. `Arc` appears only inside `RecordPayload`'s private storage
and in a `pub(super)` field. The optimization-order step is 1 (eliminate
unnecessary work — the eager text extraction) and 2 (unnecessary copying and
allocation — the per-level payload copies).

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0, glibc's default allocator; every measured process pinned to CPU
26 while seven other agents built and measured on other cores. Both legs built
`--release --locked`: before from the shared read-only checkout of `6c4c1469b`,
after from this branch. Fixture unless stated: `45543.ppt`
(`test-data/poi/test-data/slideshow/45543.ppt`, 311,524-byte `PowerPoint
Document` stream, 286 records, depth 5, copy factor 1.63).

### Record-tree census, all 30 fixtures

`record-tree-census.txt`. Owned payload bytes over stream bytes, before:
minimum 1.00 (4 encrypted), maximum 2.35 (`ppt_with_png.ppt`), 2.0 or above on
20 of the 26 that parse. After the change the parser copies **no** payload byte
on the `Presentation` path; the census itself is produced through the `&[u8]`
entry point and is byte-identical on both legs, which is the intended proof that
that entry point did not change.

### Allocations and retained bytes (counting global allocator)

`allocations-before.txt`, `allocations-after.txt`. One operation, measured after
one untimed warm-up; "retained" is live bytes still held while the result is in
scope.

| 45543.ppt | allocations | allocated bytes | peak live | retained live |
| --- | ---: | ---: | ---: | ---: |
| eager open, before | 613 | 934,751 | 897,661 | 893,053 |
| eager open, after | **304** | **429,870** | **392,484** | **387,876** |
| source-backed open, before | 596 | 892,276 | 880,242 | 854,755 |
| source-backed open, after | **287** | **387,395** | **376,057** | **349,578** |
| open + list slides, before | 840 | 1,128,182 | 1,065,441 | 1,065,077 |
| open + list slides, after | **531** | **625,445** | **561,416** | **561,052** |
| open + full text, before | 1,115 | 1,155,486 | 942,002 | 896,685 |
| open + full text, after | **806** | **652,493** | **437,945** | **391,508** |
| open + one text edit + save, before | 1,714 | 4,785,035 | 1,282,701 | 16 |
| open + one text edit + save, after | **1,096** | **3,777,673** | **1,121,093** | 16 |

Retained bytes per byte of stream for the eager open fall from **2.87x to
1.24x**. Three more fixtures, eager open, retained live: `SampleShow.ppt`
(62,399-byte stream) 205,739 → 81,489 (3.30x → 1.31x); `headers_footers_2007.ppt`
(99,635) 322,742 → 124,347 (3.24x → 1.25x); `41246-1.ppt` (68,784, 804 records)
570,017 → 416,275 (8.29x → 6.05x — that fixture's cost is dominated by 804
record structs, not by payload bytes).

### Instructions and calls (callgrind isolation pairs, s=10 vs s=110)

`callgrind/an-*.txt`. Per operation; `memcpy` is counted per byte by callgrind,
so its `Ir` column ranks work and the call column is the deterministic count.

| per operation | before | after | delta |
| --- | ---: | ---: | ---: |
| eager open, `Ir` | 1,967,863 | 1,496,895 | **−23.93%** |
| eager open, `memcpy` calls | 948.7 | 653.4 | −31.1% |
| eager open, `memcpy` self `Ir` | 1,095,573 | 750,263 | −31.5% |
| eager open, `malloc` calls | 538 | 235 | −56.3% |
| eager open, `Record::extract_text` calls | 259 | **0** | PPT-2 |
| open + list slides, `Ir` | 2,173,545 | 1,714,572 | **−21.12%** |
| open + list slides, `memcpy` calls | 1,198.2 | 899.3 | −24.9% |
| open + full text, `Ir` | 2,358,040 | 1,906,028 | **−19.17%** |
| open + full text, `Record::extract_text` calls | 328 | 69 | −259 |
| open + edit + save, `Ir` | 7,457,200 | 6,543,115 | **−12.26%** |
| open + edit + save, `memcpy` calls | 2,990.6 | 2,402.8 | −19.7% |

The 295 removed `memcpy` calls per open are the 286 record payloads plus the
text extraction's; the 259 removed `extract_text` calls are the whole of PPT-2.

### Native instructions, cycles and page faults (`perf stat`)

`perf-stat-final.txt`. Single-shot is `-r 200`: two hundred separate processes,
each doing one operation, so the numbers include process start-up (about 3.0 M
instructions) and are the honest per-use figure.

| single shot, whole process | instructions | cycles | minor faults |
| --- | ---: | ---: | ---: |
| eager open, before | 4,399,332 | 2,491,522 | 545 |
| eager open, after | **3,689,782** (−16.1%) | **2,146,626** (−13.8%) | **419** |
| open + list slides, before | 4,729,921 | 2,689,576 | 586 |
| open + list slides, after | **4,022,839** (−14.9%) | **2,304,720** (−14.3%) | **458** |
| open + full text, before | 4,801,250 | 2,698,009 | 555 |
| open + full text, after | **4,111,759** (−14.4%) | **2,340,913** (−13.2%) | **428** |
| open + edit + save, before | 8,033,388 | 4,468,415 | 1,072 |
| open + edit + save, after | **7,175,112** (−10.7%) | **3,986,404** (−10.8%) | **945** |

Every mode touches about 130 fewer pages per process, which is the ~500 KB of
payload copies that are no longer written.

### Paired timing

40 timed samples per leg after 5 warm-ups, order A1 B1 B2 A2, then two more
before legs in the same window for the A/A floor; A and B are the medians of
their two legs. `timing/summary.txt` carries p50, mean, p95 and p99 for every
leg and both directions; the raw samples are in `timing/`.

**Registered selectors** (`litchi-perf-baseline`, corpus `ppt-tiny`: a 7,680-byte
synthetic package with a 4,127-byte document stream — two orders of magnitude
smaller than the real fixtures, and the only PPT corpus the harness has):

| case | before p50 | after p50 | delta | A/A floor p50 |
| --- | ---: | ---: | ---: | ---: |
| `ppt_semantic_open` | 17,108 ns | 14,555 ns | **−14.92%** | 1.18% |
| `ppt_semantic_list_slides` | 8,632 ns | 8,528 ns | −1.22% | 1.35% |
| `ppt_semantic_full_text` | 26,622 ns | 27,878 ns | **+4.71%** | 0.32% |
| `ppt_semantic_one_edit_save` | 268,198 ns | 253,894 ns | **−5.33%** | 1.13% |

**Whole operation on the real fixture** (`ppt0606 time`, 45543.ppt):

| operation | before p50 | after p50 | delta | A/A floor p50 |
| --- | ---: | ---: | ---: | ---: |
| open | 55,440 ns | 37,685 ns | **−32.03%** | 1.63% |
| open + list slides | 68,465 ns | 140,866 ns | **+105.75%** | 1.42% |
| open + full text | 78,065 ns | 59,825 ns | **−23.37%** | 0.49% |
| open + edit + save | 261,297 ns | 226,046 ns | **−13.49%** | 1.07% |

### The two results that got worse

**`open + list slides` in a repeat loop regresses 105.75%, and it is a glibc
heap heuristic, not the parser.** In that loop the after leg takes **171,889**
minor page faults per 1,000 iterations against the before leg's **768**
(`perf-stat-final.txt`), and its instruction count rises from 748,166,354 to
1,516,098,715 for the same work. Running the identical loop with
`MALLOC_MMAP_THRESHOLD_` and `MALLOC_TRIM_THRESHOLD_` pinned at 8 MiB removes
the faults (467 after, 593 before) and the same instruction count falls to
**603,047,547 after against 746,537,786 before (−19.2%)**, and the paired timing
becomes **−26.07%** (68,110 → 50,350 ns, floor 1.15%). The mechanism is that the
before leg's ~500 KB of per-record payload allocations kept the main arena above
glibc's dynamic trim threshold, so the 311 KB stream buffer was served from the
heap; the after leg's smaller arena lets the heap be trimmed between iterations,
so that buffer is re-`mmap`ed and re-faulted every time. A process that opens
one presentation never pays this: the single-shot table above shows 586 → 458
faults and −14.9% instructions for exactly this operation. The other three
operations do not cross the threshold in either leg (open and full text stay
under 1,600 faults per 1,000 iterations; edit-and-save already paid about
139,000 in **both** legs, from the CFB rewrite).

**`ppt_semantic_full_text` regresses 4.71% on the synthetic 4 KB corpus.** That
selector times only `Presentation::text()`, with the open outside the timer, and
`text()` re-parses each slide through `SlideFactory`, which still takes a
`&[u8]` and therefore still copies payloads. So the timed region pays the new
payload representation — one enum tag test per read and 8 more bytes per
`Record` — and collects none of its benefit. The regression is below the 5%
review trigger and well above the 0.32% floor, so it is reported, not dismissed;
the same operation measured end to end on a real fixture is −23.37%. An earlier
revision of this change used a four-field struct for `RecordPayload` (48 bytes);
compacting it to the 32-byte enum moved `ppt_semantic_list_slides` from +5.56%
to −1.22% and `ppt_semantic_full_text` from +5.55% to +4.71%, which is the
evidence that the residue is `Record`'s size and the tag test.

## Correctness evidence

**Differential oracle over all 30 `.ppt` fixtures** (`driver/src/main.rs`,
`oracle` mode; per-fixture digests in `oracle-digests.txt`). For every fixture
the driver dumps, from both the eager and the source-backed reader: slide count,
full text, `extract_text_fast`, and per slide the persist id, slide id, shape
count, text, the `Debug` of every shape, placeholders, shape flags, programmable
tags, shape and outline text interactions and speaker notes; plus the picture
inventory, document atom, document structure, colour schemes, fonts, header and
footer records, hyperlinks, comments and the comment catalogue, OLE objects,
embedded sounds, external media, smart tags, text metachars, outline text refs,
modify password, privacy, presentation advisor, HTML document and publish
settings, broadcasts, envelope data and settings, routing slip, both view-info
records, document comparison, shape flags, text special-info defaults, the
PowerPoint 12 document properties and main-master metadata, the sound-reference
validation, charts and the slide directory. It then runs the owned text editor
over 12 (slide, shape) targets, recording for each the read text, the length and
FNV digest of a no-op commit and of an edited commit, and the source-backed
editor's reads over 6 targets. Finally it re-parses the document stream and
reports the record census and an FNV digest over every record's type, version,
instance, declared length, payload length, child count and payload bytes.

**30 of 30 fixture dumps are identical**, including the four encrypted files
whose refusals are recorded verbatim. The only bytes that differ anywhere in the
9,927,532-byte dump are two `HashMap` `Debug` renderings inside `SlideDirectory`
(`by_slide_id`, `by_persist_id`); two runs of the **same** before binary differ
on those same lines, because `RandomState` is seeded per process. With those two
maps normalized, both dumps hash to
`9cdddef6e870aa35884738c3f1729c9e12b6ca7bc22b2e551a4dad586c236b7c`.

**Unit tests added** (`records/record.rs`, `records/payload.rs`): shared and
owned parses agree on four stream shapes including a truncated container and a
resynchronizing stream; the copied-payload budget refuses identically at and
one byte below the tree's payload total; editing a shared payload leaves the
stream and the siblings intact; a shared and an owned payload are
indistinguishable through `PartialEq`, `Debug`, `len`, `to_vec` and `into_vec`;
`to_mut` copies before editing; an out-of-range span reads as empty; and
`RecordPayload` is 32 bytes.

**Gates** (`gates.txt`): `cargo fmt --all --check` clean; `cargo clippy -p
litchi-ppt --all-targets` and `--all-features` clean under the workspace's
deny-level lints; `cargo test -p litchi-ppt` 1,212 passed, 0 failed, 11 ignored
across 32 test binaries, and 1,240 passed with `--all-features` (which is how
the `encryption` path is exercised); `cargo doc -p litchi-ppt --no-deps` clean;
`cargo clippy -p litchi -p litchi-pptx --all-targets` clean apart from three
pre-existing dead-code warnings in a `litchi` test file this change does not
touch. `cargo check --workspace --all-targets` fails only on four `litchi-iwa`
chart examples. Both the warnings and that failure reproduce on the untouched
before checkout of `6c4c1469b`; the examples are out of scope (iWork).

## Validation preserved

Every record limit keeps its value, its charge and its refusal point, including
`max_copied_payload_bytes`, which is still charged per record on the borrowing
path. The strict parser still rejects a record that extends past its container,
a truncated container header and zero-length progress; the lenient parser still
truncates an over-long container payload, still descends into what it has, and
still resynchronizes one byte at a time. Encrypted documents still refuse
identically. Nothing in the parse path became lenient, and nothing that
validated bytes stopped reading them.

## Limitations

- Nothing is claimed for any scenario, fixture, host or build not listed above.
  All timings are warm-cache, in-memory, single-threaded, on one host, with
  seven other agents active on other cores.
- The only registered PPT corpus is a 7,680-byte synthetic package; the real
  fixture numbers come from a scratch driver retained in this packet, not from a
  registered selector. No selector exists for a PPT open on a real fixture.
- The retained-bytes figures are counting-allocator live bytes, not RSS. No
  cold-cache, physical-device or RSS measurement was taken.
- The `open + list slides` loop regression is characterized, not eliminated. It
  is reported as a review trigger.
- `SlideFactory` — and therefore `Presentation::slides`, `slide_at` and `text`,
  and `NotesIndex` — still re-parses each slide through the copying `&[u8]`
  entry point. Extending the shared buffer to it is the obvious follow-on and
  is the reason `ppt_semantic_full_text` regresses; it is deliberately left out
  of this change, which is already 84 files.
- `41246-1.ppt` shows that on a record-dense, payload-light file the remaining
  cost is the `Record` struct itself (804 of them), which this change makes 8
  bytes larger. No packing of `Record` was attempted.
- The `owned-edit-save` path could not be exercised on `41246-1.ppt` (the owned text
  editor refuses it with `Refused(NoTextAtom)`); that cell is absent from both legs.

## Retained evidence

[`results/change-0606/README.md`](results/change-0606/README.md).

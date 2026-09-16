# 0634: the per-slide re-parse spans the stream the presentation already retains, and the glibc heap-trim artifact 0606 left open disappears

Status: retained. `performance_claim: none` — this record carries deterministic
counts (allocations, retained bytes, `memcpy` and `malloc` call counts,
callgrind instruction differentials), native instruction, cycle and page-fault
counts, and paired timings reported beside the host's A/A floor. No claim is
registered.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is item **16** of change [0630](0630-queue-refresh-after-the-first-wave.md)'s
refreshed queue, the follow-on change [0606](0606-ppt-record-tree-borrowed-payloads.md)
named in its "Limitations": *"`SlideFactory` — and therefore `Presentation::slides`,
`slide_at` and `text`, and `NotesIndex` — still re-parses each slide through the
copying `&[u8]` entry point. Extending the shared buffer to it is the obvious
follow-on and is the reason `ppt_semantic_full_text` regresses."* Change
[0301](changes/0301-ppt-selected-slide-decoding.md) left the complete document
parser out of scope; 0606 did that parser, and this record does the re-parse
0606 left behind.

## The mechanism 0606 left behind

After 0606 a `Presentation` holds its `PowerPoint Document` stream as an
`Arc<Vec<u8>>` and the top-level record tree spans it. But three paths re-parse
that same stream through `Record::parse_with_limits`, which copies every
record's payload into its own `Vec`:

- `SlideFactory::parse_slide_at_offset` — once per slide, for
  `Presentation::slides`, `slide_at`, `text` and `extract_text_fast`.
- `NotesIndex::try_build_with_limits` — a **whole `DocumentContainer`** re-parse,
  built lazily on the first `parse_slide` and therefore paid by every one of
  those operations.
- `SpeakerNotes::parse_with_limits` — once per slide that has a notes page.

So `Presentation::text()` paid 0606's new payload representation — one enum tag
test per read and eight more bytes per `Record` — and collected none of its
benefit. That is the mechanism behind the **+4.71%** `ppt_semantic_full_text`
0606 reported and did not hide.

## What was changed

Scope: `crates/litchi-ppt` only. Ten files; no other crate's files are touched.

**`records/record.rs`.** `PayloadStore` — the `Owned(&[u8])` / `Shared(&Arc<Vec<u8>>)`
selector 0606 introduced inside `parse_impl` — gains a `pub(crate)` `bytes()`
accessor and a `Debug` that forwards to `bytes()`, so a struct that used to hold
a `&[u8]` and derive `Debug` renders exactly the same text. `Record` gains
`parse_from_store_with_limits(store, offset, limits)`, which computes
`window_end = store.bytes().len()`, builds `ParseBudget::new(limits, window_end)`
and calls `parse_impl(store, offset, window_end, false, 0, &mut budget)` —
character for character what `parse_with_limits` did. `parse_with_limits` is now
exactly `parse_from_store_with_limits(PayloadStore::Owned(data), offset, limits)`
and keeps its behaviour; no other entry point changes.

**`records/mod.rs`.** `PayloadStore` is re-exported `pub(crate)` so the slide
layer can name it. It stays crate-private; nothing public leaks.

**`slide/factory.rs`.** `SlideFactory`'s `doc_data: &'doc [u8]` becomes
`source: PayloadStore<'doc>`. `new` and `new_with_limits` keep their `&[u8]`
signatures and construct `PayloadStore::Owned`; a new `pub(crate)`
`new_shared_with_limits(&'doc Arc<Vec<u8>>, …)` constructs `PayloadStore::Shared`.
`parse_slide_at_offset` bounds-checks against `self.source.bytes().len()` and
parses through `parse_from_store_with_limits`. `SlideData`'s field keeps the name
`doc_data` (so its derived `Debug` is unchanged) with type `PayloadStore<'doc>`;
`SlideData::doc_data()` still returns `&'doc [u8]`, now `self.doc_data.bytes()`.

**`slide/notes.rs`.** `NotesIndex::build_with_limits` /
`try_build_with_limits` and `SpeakerNotes::parse_with_limits` take a
`PayloadStore<'_>` instead of a `&[u8]` and parse through it. The test-only
`build` / `try_build` helpers wrap their slice in `PayloadStore::Owned`.

**`slide/types/mod.rs`, `types/model.rs`, `types/semantic.rs`.** `Slide`'s
`doc_data: &'doc [u8]` becomes `source: PayloadStore<'doc>`, carried across from
`SlideData` and handed to `SpeakerNotes::parse_with_limits`.

**`presentation/model.rs`.** The four `SlideFactory::new_with_limits(&self.powerpoint_document, …)`
call sites — `slides`, `slide_at`, `text`, `extract_text_fast` — become
`new_shared_with_limits`.

**`text_edit.rs`.** The source-backed resolve path wraps the `PowerPoint Document`
stream it reads in an `Arc` and uses the shared factory for the one selected
slide. The stream is a fresh local buffer there, so the `Arc` costs one
allocation and saves that slide's payload copies; every reader below still sees
`&[u8]` by deref.

No public API signature changes. `SlideFactory::new`, `SlideData::doc_data` and
every other `pub` item keep their types.

## Why it is sound

**Every record is value-identical.** `parse_from_store_with_limits` reaches
`parse_impl` with the same buffer, the same `offset`, the same `window_end`, the
same `strict = false`, the same `depth = 0` and a `ParseBudget` built from the
same `(limits, len)` pair as before. Only `PayloadStore::payload` differs, and
0606's `RecordPayload` makes the two forms indistinguishable through `Deref`,
`PartialEq` and `Debug`. `PayloadStore`'s own new `Debug` renders the byte slice
it stands in for, so `SlideData`'s derived `Debug` is unchanged.

**Error identity is unchanged, including the copy budget.** `charge_copy` is
still called for every record with the same byte count whether or not the
payload is copied, so `max_copied_payload_bytes` refuses exactly the slides it
refused before. A unit test parses the first slide of four real fixtures under a
budget equal to that slide's payload total and one byte below it, and asserts
both payload sources accept, refuse and format their refusal identically. The
one `Error` variant that becomes unreachable on the shared path is
`AllocationFailed("PPT record payload")`, which fires only when the allocator
refuses a payload the shared path does not allocate — the same narrowing 0606
recorded.

**Mutation cannot reach the stream.** `RecordPayload::to_mut` copies a borrowed
span before returning `&mut`, and `DerefMut` routes through it. A test parses a
slide from the shared stream, edits its payload in place and asserts both the
`Arc`'d stream and the sibling records are unchanged.

**The lenient parse behaviour is untouched.** No container loop, truncation
recovery, byte-at-a-time resynchronization, depth ceiling or record-count
ceiling is edited; the only edits inside `record.rs` are the new wrapper, the
`bytes()` visibility and the `Debug` impl.

**No contract moved.** No new `unsafe` (`litchi-ppt` keeps
`#![forbid(unsafe_code)]`), no weakened limit or malformed-input defence, no
hidden global pool, no ambient I/O, no public leakage of archive types, raw
locks or executors. `Arc` appears only in `pub(crate)` signatures and private
fields. The optimization-order step is 2 (unnecessary copying and allocation).

## Measured

Host: AMD EPYC 9R45, 32 logical CPUs, 123 GiB, Linux 7.0.0-1012-aws, rustc
1.95.0, valgrind 3.26.0, glibc's default allocator; every measured process
pinned to CPU 10 with `taskset` while seven other agents built and measured on
other cores. Both legs built `--release --locked`: **before** from the shared
read-only checkout of `c7326f680` (which contains 0606), **after** from this
branch. Fixture unless stated: `45543.ppt`
(`test-data/poi/test-data/slideshow/45543.ppt`, 311,524-byte `PowerPoint
Document` stream, 286 records). The speaker-notes mode uses
`headers_footers_2007.ppt` (99,635-byte stream), the only fixture family with
notes pages; `45543.ppt` has none.

### Allocations and retained bytes (counting global allocator)

`allocations-before.txt`, `allocations-after.txt`. One operation after one
untimed warm-up; "retained" is live bytes still held while the result is in
scope.

| operation | allocations | allocated bytes | peak live | retained live |
| --- | ---: | ---: | ---: | ---: |
| eager open, before | 304 | 429,870 | 392,484 | 387,876 |
| eager open, after | **304** | **429,870** | **392,484** | **387,876** |
| source-backed open, before | 287 | 387,395 | 376,057 | 349,578 |
| source-backed open, after | **287** | **387,395** | **376,057** | **349,578** |
| open + list slides, before | 531 | 625,445 | 561,416 | 561,052 |
| open + list slides, after | **371** | **467,494** | **416,400** | **416,036** |
| open + full text, before | 806 | 652,493 | 437,945 | 391,508 |
| open + full text, after | **646** | **494,542** | **398,396** | 391,508 |
| open + slides + notes (hf2007), before | 467 | 273,662 | 205,221 | 196,065 |
| open + slides + notes (hf2007), after | **369** | **205,743** | **145,057** | **135,901** |
| open + one text edit + save, before | 1,096 | 3,777,673 | 1,121,093 | 16 |
| open + one text edit + save, after | **936** | **3,619,722** | 1,121,093 | 16 |

The two open modes are **bit-identical**, which is the intended proof that the
open path — the part 0606 already borrowed — is untouched. Every other mode
loses exactly the payload allocations of the records it re-parses: **160** per
`slides()` on `45543.ppt` and **98** per notes read on
`headers_footers_2007.ppt`. Callgrind counts 161 extra `parse_impl` calls for
`slides()` over a bare open, so the removed allocations are one per re-parsed
record, less the one record whose payload is empty and never allocated.

### Instructions and calls (callgrind isolation pairs, s=10 vs s=110)

`callgrind/an-<leg>-<mode>.txt`. Per operation; `memcpy` is counted per byte by
callgrind, so its `Ir` ranks work while the call column is the deterministic
count.

| per operation | before | after | delta |
| --- | ---: | ---: | ---: |
| eager open, `Ir` | 1,494,743 | 1,498,532 | **+0.25%** |
| eager open, `memcpy` calls | 654.7 | 654.7 | 0 |
| eager open, `malloc` calls | 235 | 235 | 0 |
| open + list slides, `Ir` | 1,711,522 | 1,607,248 | **−6.09%** |
| open + list slides, `memcpy` calls | 899.3 | 738.5 | −160.8 |
| open + list slides, `malloc` calls | 439 | 279 | −160 |
| open + full text, `Ir` | 1,901,110 | 1,808,201 | **−4.89%** |
| open + full text, `memcpy` calls | 946.9 | 796.0 | −150.9 |
| open + full text, `malloc` calls | 662 | 502 | −160 |
| open + slides + notes, `Ir` | 834,969 | 790,604 | **−5.31%** |
| open + slides + notes, `memcpy` calls | 779.3 | 680.4 | −98.9 |
| open + slides + notes, `malloc` calls | 406 | 308 | −98 |
| open + edit + save, `Ir` | 6,532,253 | 6,433,187 | **−1.52%** |
| open + edit + save, `memcpy` calls | 2,404.9 | 2,237.0 | −167.9 |
| open + edit + save, `malloc` calls | 896 | 736 | −160 |

`parse_impl` is called the same number of times in both legs of every mode (485
per `slides()`, for instance): the same records are parsed, only their payloads
stop being copied. The per-op self-`Ir` deltas for `slides()` are
`__memcpy_avx_unaligned_erms` −65,275, `_int_malloc` −11,256,
`_int_free_maybe_trim` −7,308, `_int_free_merge_chunk` −6,315, `free` −5,046,
`malloc` −3,685, against +3,021 in `parse_impl` and +3,200 saved in the new
`PayloadStore::payload`.

### Native instructions, cycles and page faults (`perf stat`)

`perf-stat.txt`. Single-shot is `-r 200`: two hundred separate processes each
doing one operation, so the numbers include process start-up and are the honest
per-use figure.

| single shot, whole process | instructions | cycles | minor faults |
| --- | ---: | ---: | ---: |
| eager open, before | 3,847,165 | 2,168,268 | 423 |
| eager open, after | 3,838,539 (−0.22%) | 2,206,785 (+1.78%) | 424 |
| open + list slides, before | 4,197,128 | 2,372,566 | 463 |
| open + list slides, after | **3,972,746** (−5.35%) | **2,287,813** (−3.57%) | **429** |
| open + full text, before | 4,270,710 | 2,379,663 | 433 |
| open + full text, after | **4,168,140** (−2.40%) | **2,357,470** (−0.93%) | **426** |
| open + slides + notes, before | 2,975,384 | 1,784,079 | 267 |
| open + slides + notes, after | **2,894,466** (−2.72%) | **1,758,295** (−1.45%) | **254** |
| open + edit + save, before | 7,347,211 | 4,042,729 | 949 |
| open + edit + save, after | **6,816,432** (−7.22%) | **3,819,493** (−5.52%) | **878** |

### The glibc heap-trim artifact 0606 characterized but did not eliminate

0606 measured a 1,000-iteration single-process `open + list slides` loop at
**171,889** minor faults under glibc defaults against 768 before it, and
reported the resulting +105.75% paired timing as a review trigger it had
characterized but not fixed. Repeating that measurement here:

| `open + list slides`, 1,000 iterations | instructions | cycles | minor faults |
| --- | ---: | ---: | ---: |
| glibc defaults, before | 1,523,645,307 | 630,155,258 | 171,889 |
| glibc defaults, after | **551,758,441** (−63.79%) | **207,574,548** (−67.06%) | **667** (−99.61%) |
| thresholds pinned at 8 MiB, before | 600,275,872 | 224,519,772 | 474 |
| thresholds pinned at 8 MiB, after | **550,284,292** (−8.33%) | **206,824,993** (−7.88%) | 494 |

**The artifact is gone.** With the thresholds pinned the two legs now differ by
−8.33%, which is the real work removed; with glibc's defaults the after leg
lands on the same instruction count as the pinned run (551.8 M against 550.3 M),
while the before leg pays 973 M extra instructions and 171,000 extra faults.
The mechanism is the one 0606 named, read the other way: it was the per-iteration
churn of ~160 payload allocations inside `slides()` that grew and shrank the top
of the main arena past glibc's dynamic trim threshold each iteration, so the
311 KB stream buffer was released and re-faulted every time. Removing those
allocations removes the churn, so the arena is never trimmed. `open + full text`
falls from 1,590 to 1,265 faults per 1,000 iterations for the same reason; the
edit-and-save loop still pays about 138,500 faults in **both** legs, from the
CFB rewrite, unchanged.

### Paired timing

40 timed samples per leg after 5 warm-ups, order A1 B1 B2 A2, then two more
before legs in the same window for the A/A floor; A and B are the medians of
their two legs. `timing/summary.txt` carries p50, mean, p95 and p99 for every
leg and both directions; the raw samples are in `timing/`.

**Registered selectors** (`litchi-perf-baseline`). Both PPT corpora are
reported: `ppt-large` (40,960-byte package, 144 entries, 37,385-byte document
stream) and `ppt-tiny` (7,680-byte package, 4,127-byte stream). 0606's table is
labelled `ppt-tiny` but its numbers are the `ppt-large` block; the two are
distinguished here.

| case, corpus | before p50 | after p50 | delta | A/A floor p50 |
| --- | ---: | ---: | ---: | ---: |
| `ppt_semantic_list_slides`, ppt-large | 8,560 ns | 6,550 ns | **−23.48%** | 0.82% |
| `ppt_semantic_open`, ppt-large | 14,765 ns | 14,080 ns | −4.64% | 0.62% |
| `ppt_semantic_full_text`, ppt-large | 27,770 ns | 26,855 ns | **−3.29%** | 0.15% |
| `ppt_semantic_one_edit_save`, ppt-large | 245,681 ns | 244,246 ns | −0.58% | 0.04% |
| `ppt_semantic_full_text`, ppt-tiny | 1,990 ns | 1,945 ns | −2.26% | 1.01% |
| `ppt_semantic_one_edit_save`, ppt-tiny | 95,416 ns | 94,901 ns | −0.54% | 0.02% |
| `ppt_semantic_open`, ppt-tiny | 6,935 ns | 6,915 ns | −0.29% | 1.71% |
| `ppt_semantic_list_slides`, ppt-tiny | 1,405 ns | 1,410 ns | +0.36% | 0.72% |

**Whole operation on the real fixtures** (`ppt0634 time`). Under glibc defaults
and again with `MALLOC_MMAP_THRESHOLD_` and `MALLOC_TRIM_THRESHOLD_` pinned at
8 MiB, because the `slides()` loop crosses the trim threshold in the before leg
only.

| operation | before p50 | after p50 | delta | A/A floor p50 |
| --- | ---: | ---: | ---: | ---: |
| open + list slides, defaults | 143,591 ns | 46,220 ns | **−67.81%** | 1.02% |
| open + list slides, pinned | 50,975 ns | 46,750 ns | **−8.29%** | 0.27% |
| open + slides + notes (hf2007), defaults | 31,460 ns | 29,890 ns | **−4.99%** | 1.33% |
| open + slides + notes (hf2007), pinned | 31,300 ns | 30,320 ns | −3.13% | 0.29% |
| open + full text, defaults | 59,546 ns | 57,030 ns | **−4.22%** | 0.27% |
| open + full text, pinned | 59,845 ns | 57,845 ns | −3.34% | 1.15% |
| open + edit + save, defaults | 227,471 ns | 218,826 ns | **−3.80%** | 0.42% |
| open + edit + save, pinned | 151,511 ns | 149,350 ns | −1.43% | 0.16% |
| open, defaults | 37,560 ns | 37,790 ns | **+0.61%** | 1.04% |
| open, pinned | 37,890 ns | 38,196 ns | **+0.81%** | 0.26% |

### The two results that need explaining

**`ppt_semantic_open` improves 4.64%, and it is not the open.** That selector's
untimed per-iteration verification (`verify_semantic_ppt`) runs a full
`slides()` + `shapes()` + `text()` inside every iteration, outside the timer.
This change makes that verification allocate 160 fewer blocks per iteration, so
the timed open starts from a different allocator state. The open path itself is
provably untouched: identical allocation count, identical allocated bytes,
identical `memcpy` and `malloc` call counts, and the driver's own whole-operation
open on the real fixture — which has no such verification — moves the **other**
way.

**The open regresses about 0.8%, and it is codegen drift, not work.** Callgrind
puts it at +3,789 `Ir` per open (+0.25%) with every call count identical; the
native 1,000-iteration loop agrees at +3,737 `Ir` per open (+0.85%) with cycles
−0.10%; the paired p50 is +0.61% against a 1.04% floor under glibc defaults and
+0.81% against a 0.26% floor with the thresholds pinned. The listed per-symbol
deltas are −2,934 in `drop_in_place<Record>` against +1,680 in `parse_impl`,
with the remainder spread below the annotation threshold — the shape of
inlining drift after `parse_with_limits` became a wrapper, not of new work. It
is well below the 5% review trigger, above the floor in one configuration, and
reported rather than averaged away. No attempt was made to chase it: the
candidate fixes are inlining hints, which is speculative complexity for under
1%.

**`ppt_semantic_full_text` comes back to, but not measurably below, its 0606
baseline.** 0606 measured this selector at 26,622 ns before it and 27,878 ns
after (+4.71%, floor 0.32%). In this window the before leg — 0606's code —
measures 27,770 ns and the after leg 26,855 ns (−3.29%, floor 0.15%). So the
regression is recovered almost entirely, leaving about +0.9% against 0606's
absolute pre-change number and about +1.3% by compounding the two deltas. That
residue is what 0606 predicted would remain: `Record` is 8 bytes larger and
every payload read tests an enum tag, and this change removes neither.

## Correctness evidence

**Differential oracle over all 30 `.ppt` fixtures**, 0606's driver run verbatim
on both legs (`driver/src/main.rs`, `oracle` mode). For every fixture it dumps,
from both the eager and the source-backed reader: slide count, full text,
`extract_text_fast`, and per slide the persist id, slide id, shape count, text,
the `Debug` of every shape, placeholders, shape flags, programmable tags, shape
and outline text interactions and **speaker notes**; plus the picture inventory,
document atom, document structure, colour schemes, fonts, header and footer
records, hyperlinks, comments and the comment catalogue, OLE objects, embedded
sounds, external media, smart tags, text metachars, outline text refs, modify
password, privacy, presentation advisor, HTML document and publish settings,
broadcasts, envelope data and settings, routing slip, both view-info records,
document comparison, shape flags, text special-info defaults, the PowerPoint 12
document properties and main-master metadata, the sound-reference validation,
charts and the slide directory. It then runs the owned text editor over 12
(slide, shape) targets, recording the read text and the length and FNV digest of
a no-op and an edited commit for each, and the source-backed editor's reads over
6 targets, and finally re-parses the document stream and reports a record census
and an FNV digest over every record's type, version, instance, declared length,
payload length, child count and payload bytes.

**30 of 30 fixture dumps are identical** (`oracle-digests.txt`), including the
four encrypted files whose refusals are recorded verbatim. Both 9,927,532-byte
dumps hash to
`9cdddef6e870aa35884738c3f1729c9e12b6ca7bc22b2e551a4dad586c236b7c` once the two
per-process-randomized `HashMap` `Debug` renderings inside `SlideDirectory` are
normalized — **the same digest 0606 recorded**, so the reader surface is
unchanged across both changes.

**Unit tests added** (`presentation/tests.rs`), each over four real fixtures
(`45543.ppt`, `headers_footers_2007.ppt`, `SampleShow.ppt`, `41246-1.ppt`) and
asserting every fixture was opened:

- the owned and the shared factory return the same slide ids, the same `Record`
  trees, offsets and slide ids, the same note descriptors, the same
  `doc_data()`, and the same `text()`, `shapes()` and `speaker_notes()` renderings
  for every slide — and agree on failure where a slide fails;
- `max_copied_payload_bytes` set to a slide's exact payload total accepts on both
  sources and one byte below refuses on both, with the same formatted error;
- editing a borrowed slide payload in place leaves the `Arc`'d stream and the
  sibling records byte-identical.

**Gates** (`gates.txt`), all clean: `cargo fmt --all --check`; `cargo clippy -p
litchi-ppt --all-targets` and `--all-targets --all-features` under the
workspace's deny-level lints; `cargo doc -p litchi-ppt --no-deps`; `cargo test -p
litchi-ppt` **1,215 passed, 0 failed, 11 ignored** across 32 test binaries and
1,243 passed with `--all-features` (which exercises the `encryption` path);
`cargo clippy -p litchi -p litchi-pptx --all-targets`; `cargo test -p
litchi-pptx`; `cargo test -p litchi --features docx,xlsx,pptx,xls` and `--features
ppt,pptx,docx,xlsx,xls`; and `cargo test` in `tools/perf-baseline`.

## Validation preserved

Every record limit keeps its value, its charge and its refusal point, including
`max_copied_payload_bytes`, which is still charged per record on the borrowing
path so the refusal boundary cannot widen. The strict and lenient container
loops, the truncated-container recovery, the byte-at-a-time resynchronization,
the depth and record-count ceilings and the `NotesIndex` validations (duplicate
notes list, null `persistIdRef`, `NotesId` range, `SlidePersistAtom` header and
length, `NotesAtom` header, length and `slideIdRef` cross-check) are unchanged
and still read the same bytes. Encrypted documents still refuse identically.
Nothing in the parse path became lenient.

## Limitations

- Nothing is claimed for any scenario, fixture, host or build not listed above.
  All timings are warm-cache, in-memory, single-threaded, on one host, with
  seven other agents active on other cores.
- `performance_claim: none`. The paired medians and counts are evidence, not a
  registered claim.
- The registered PPT corpora are two synthetic packages of 7,680 and 40,960
  bytes; every real-fixture number comes from the scratch driver retained in
  this packet. There is still no registered selector for a PPT open on a real
  fixture.
- The retained-bytes figures are counting-allocator live bytes, not RSS. No
  cold-cache, physical-device or RSS measurement was taken.
- The open path regresses about 0.8% on the real fixture and 0.25% in callgrind
  `Ir`. It is characterized as codegen drift (identical call and allocation
  counts) and not eliminated.
- `ppt_semantic_full_text` recovers most but not all of 0606's +4.71%; about
  +1% against 0606's absolute pre-change number remains, and this change does
  not address its stated cause (`Record`'s extra 8 bytes and the payload tag
  test).
- Five copying re-entries remain in `text_edit.rs`, all through the *strict*
  entry point (`Record::parse_strict_with_limits` / `Record::parse_strict`):
  `reject_macro_records` and `inspect_source_text_atom` on the source-backed
  resolve, `inspect_slide` and `inspect_anchor` on a selected slide's bytes, and
  `rebuild`. They are once per resolve or per commit rather than once per slide,
  and two of them parse a slice that is already a copy; they were left alone. A
  shared strict entry point is the obvious next step there.
- `41246-1.ppt` still shows that on a record-dense, payload-light file the
  remaining cost is the `Record` struct itself. No packing of `Record` was
  attempted.
- The `eager-notes` measurements use `headers_footers_2007.ppt` because
  `45543.ppt` carries no notes page; only three fixtures in the corpus do.

## Retained evidence

[`results/change-0634/README.md`](results/change-0634/README.md).

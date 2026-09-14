# 0576: index the shared-string table without decoding it

Status: retained. `performance_claim: none` — this record carries paired medians
in two directions with a measured noise floor, isolated hardware counters,
per-open allocation counts and a corpus-wide differential, not a registry claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements opportunity 1 of change
[0574](0574-ole2-next-opportunity-survey.md). It takes the **conservative**
option that record left open: the walk still validates UTF-16 well-formedness, so
the refusal stays at open time and its message stays byte-identical. Deferring
SST decoding is not attempted here and is discussed under *What this does not
do*.

## What was removed

`scan_shared_string_records` builds an index of `(start, end)` logical offsets so
that a later cell lookup can fetch one shared string. It obtained each boundary
by **fully parsing the string and throwing it away**:

```rust
for string_index in 0..unique_count {
    let start = cursor.logical_position();
    parse_one_shared_string(&mut cursor, string_index)?;
    let end = cursor.logical_position();
    entries.push(SharedStringEntryLocation { start, end });
}
```

`parse_one_shared_string` reached `SstCursor::read_characters`, which allocated a
`Vec<u16>` of `count` code units, pushed them **one at a time through a
per-code-unit segment lookup**, and finished with `String::from_utf16` — a second
allocation plus a full UTF-16 to UTF-8 transcode. The `String` was dropped on the
next line. Per shared string that was two heap allocations, two frees, a
per-character indexing loop and a transcode, to advance a cursor.

The scan now walks the same framing without building either. Three things were
removed, and it matters which:

1. the `Vec<u16>`;
2. the `String` and its transcode;
3. the per-code-unit `self.current()` lookup — the framing walk now hands each
   contiguous run to the consumer as one slice, which benefits the materializing
   path too.

**One implementation of the framing, two consumers.** There is no second copy of
BIFF continuation handling. `SstCursor::walk_characters` is the only place that
knows about character counts, the high-byte flag, `Continue` boundaries and
continuation flags; it drives a `CodeUnitSink`. `CollectedCodeUnits` materializes
and `SurrogatePairing` validates. One level up, `walk_one_shared_string<T>` is the
only place that knows a shared string's record layout — header, flags, rich-text
run count, character data, formatting runs, `ExtRst` — and is instantiated with
`String` for retrieval and with `MeasuredText` for the scan. `parse_one_shared_string`
is now a one-line alias of the first instantiation and still serves
`decode_shared_string_entry`, the retrieval path, which is unchanged in what it
returns.

One duplication is **not** removed and is recorded rather than discovered later:
`SharedStringTable::parse_segments`, the eager BIFF path, keeps its own copy of
the per-string record layout. It was already a separate copy before this change.
It does share the new character framing, because it reaches it through
`read_characters`, so the part of the duplication that this change could have
made worse is the part that is now common.

## Measured effect

Two binaries built from a detached git worktree at `163ac1bd6` outside the
repository, differing **only** in `crates/litchi-xls/src/records.rs`. That commit
was `HEAD` when the legs were built; `HEAD` has since advanced by one
documentation-only commit (change 0572) that touches no file under `crates/`, so
the legs are still code-identical to the branch tip apart from this change. Both are the existing
`xls_source_attribution` harness; the per-leg driver is change 0574's
`capture_counters.sh`, unchanged, so the children are launched exactly as that
record's were: CPU 17 pinned, ASLR disabled, single-threaded, 50 warmups and
1,000 samples. Identities are in
[`environment.json`](results/change-0576/environment.json).

**Host quiescence is not established.** Other agents were active on the machine;
the load average was 4.86 when the capture window opened and 3.47 when it closed,
on 32 cores. That is why the noise floor below is measured rather than assumed.

### The control: not one byte of I/O moved

Every logical counter is identical between the legs, in all 16 captured cells:

| fixture | reads | read bytes | `version()` |
| --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 53 | 565,201 | 29 |
| `WithCustomViews.xls` | 16 | 110,242 | 22 |
| `54016.xls` | 40 | 317,171 | 25 |

The flagship row reproduces change [0565](0565-xls-globals-single-pass.md)'s and
change 0574's retained figures exactly. This change removes CPU work only.

### Wall clock, paired medians, A/B/B/A

`p50` of the measured operation, nanoseconds. `dir1` is the first
before→after pair, `dir2` the second; `A/A` and `B/B` are the **same binary**
against itself across the same window and are the noise floor.

| fixture | source | before p50 | after p50 | dir1 | dir2 | A/A | B/B |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `WithCustomViews.xls` | owned | 208,394 | **31,950** | **−84.57%** | **−84.77%** | +1.05% | −0.25% |
| `WithCustomViews.xls` | file | 214,986 | **39,505** | **−81.62%** | **−81.63%** | −0.01% | −0.07% |
| `54016.xls` | owned | 688,168 | **256,704** | **−62.63%** | **−62.77%** | +0.56% | +0.18% |
| `54016.xls` | file | 702,896 | **270,402** | **−61.56%** | **−61.50%** | +0.06% | +0.21% |
| `ConditionalFormattingSamples.xls` | owned | 90,520 | **80,138** | **−11.38%** | **−11.56%** | +0.07% | −0.14% |
| `ConditionalFormattingSamples.xls` | file | 108,778 | **98,320** | **−9.43%** | **−9.80%** | −0.44% | −0.85% |

**The noise floor, measured in this window.** Across all 16 captured cells the
same binary against itself moved between **−1.36% and +1.05%** at p50. Call the
floor ±1.4%. An earlier capture of the same matrix in a busier window — load 6.2
to 6.8 rather than 3.5 to 4.9 — measured a wider floor of −1.58% to +2.24% and
per-fixture improvements within 1.7 percentage points of these. It was discarded
and is not cited, because its `after` binary predated a comment-only edit to
`records.rs` and its retained identity would not have matched the tree. Its
agreement with this one is reported here and nowhere else.

**Change 0574 predicted the flagship's improvement would be inside noise. It is
not, and that prediction is recorded as falsified.** Its smallest measured
improvement, −9.43%, is 6.9 times the widest same-binary excursion, and the two
directions agree on every fixture to within 0.4 percentage points. The prediction
was reasonable — the flagship's SST is only 4,603 bytes — but it was made against
an unmeasured noise floor, and this host's floor is tighter than that.

Tails move with the medians rather than against them: p99 falls 79.1% / 61.3% /
9.0% on the three fixtures' `file-source` open, and 82.2% / 62.5% / 12.2% on the
`owned-readat` open.

The `one-cell` selector falls by the same amount as `open` on the two fixtures
where it is available (−84.42% and −11.20% owned), which is the expected shape:
the saving is all in the open, and decoding the one string the caller actually
asked for still costs exactly what it cost.

### Instructions per open

`perf stat`, isolated by differencing an 1,100-sample and a 100-sample child and
dividing by 1,000 — change 0574's method, on the `owned-readat` open.

| fixture | instructions before | after | change | branch misses before | after |
| --- | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 1,643,022 | **1,391,409** | **−15.31%** | 1,053 | 681 |
| `WithCustomViews.xls` | 5,312,769 | **627,989** | **−88.18%** | 4,850 | 195 |
| `54016.xls` | 17,139,968 | **5,979,090** | **−65.12%** | 10,241 | 364 |

The flagship's before leg is 1,643,022 against change 0574's independently
captured 1,645,028 — a 0.12% reproduction, which is the strongest available
evidence that the two records measure the same thing. Branch misses fall 35.4%,
96.0% and 96.5%: the per-code-unit loop is gone.

### Where the instructions went, by symbol

Callgrind, same isolation method, per open.

| | `Conditional…` | `WithCustomViews` | `54016` |
| --- | ---: | ---: | ---: |
| `scan_shared_string_records` **inclusive**, before | 382,361 | 4,946,335 | 14,848,546 |
| the same, after | **115,701** | **252,497** | **3,724,624** |
| | −69.7% | **−94.9%** | −74.9% |
| whole open, before | 2,798,346 | 5,608,091 | 18,219,149 |
| whole open, after | **2,535,964** | **922,514** | **7,305,937** |
| | −9.38% | −83.55% | −59.90% |

The before column reproduces change 0574's inclusive figures (382,422 /
4,947,696 / 14,890,837) to within 0.3%.

Self cost per open, for the symbols that move most:

| symbol | `Conditional…` | `WithCustomViews` | `54016` |
| --- | ---: | ---: | ---: |
| `SstCursor::read_characters` | 95,589 → **0** | 1,804,670 → **0** | 4,403,645 → **0** |
| `String::from_utf16` | 117,033 → 5,114 | 2,635,609 → **2,272** | 5,781,331 → **2,164** |
| `parse_one_shared_string` | 35,835 → 0 | 56,699 → 0 | 931,511 → 0 |
| `walk_one_shared_string` | 0 → 32,199 | 0 → 51,011 | 0 → 836,795 |

Allocator symbols follow: `malloc` plus `free` fall from 78,584 to 41,746 on the
flagship and from 987,208 to **9,129** on `54016.xls`, `RawVecInner::finish_grow`
from 322,602 to 6,922 on the latter, and `_int_malloc` from 83,557 to 11,883 on
`WithCustomViews.xls`. That is the same removal the allocation counts below
measure directly, seen from the instruction side.

`read_characters` is exactly **zero** on all three fixtures, which is the precise
claim: it is reachable only from the materializing instantiation and from
`SharedStringTable::parse_segments`, and a source-backed open now enters neither.
It has not been deleted — it is still the retrieval path — and the framing it used
to own now lives in `walk_characters`, which the optimizer inlines into
`MeasuredText::consume`.

`String::from_utf16` does **not** fall to zero, and the residue is worth naming
rather than rounding away. It falls by 95.6%, 99.91% and 99.96%, and what remains
is not shared strings: it is the other BIFF8 strings a source-backed open
genuinely decodes, chiefly the `BoundSheet8` worksheet names through
`utils::parse_string_record`. It tracks worksheet-name bytes, which is the
opposite of how it would scale if any of it were SST work:

| | `Conditional…` | `WithCustomViews` | `54016` |
| --- | ---: | ---: | ---: |
| `BoundSheet8` records | 16 | 3 | 1 |
| bytes of name payload | 191 | 21 | 8 |
| `from_utf16` Ir remaining | 5,114 | 2,272 | 2,164 |
| SST bytes | 4,603 | 101,125 | 225,007 |

The residue is ordered by the name bytes and inversely ordered by the SST
size.

**The framing walk remains, exactly as change 0574 predicted.** The scan still
costs 252,497 Ir per open on `WithCustomViews.xls` and 3,724,624 on `54016.xls`.
That residue is the header, flag, run-count and `Continue` handling, plus the
surrogate check. It is not free and this change does not claim it is.

### Allocations per open

Counted through a test-only global allocator in
`crates/litchi-xls/tests/sst_scan_allocations.rs`. The counters are per-thread,
so a measurement is unaffected by what the rest of the test binary is doing.
Reallocations count as one allocation and contribute their growth to bytes.

| fixture | unique strings | allocations before | after | bytes before | after |
| --- | ---: | ---: | ---: | ---: | ---: |
| `54016.xls` | 7,893 | **16,031** | **247** | 1,358,519 | 754,976 |
| `WithCustomViews.xls` | 474 | **1,478** | **236** | 533,308 | 153,369 |

`16,031 − 247 = 15,784`, against `2 × 7,893 = 15,786` — two allocations per
shared string, minus the one string in that workbook whose character count is
zero and which therefore reserved nothing and transcoded to nothing. That
workbook has **exactly one** empty shared string, read independently out of its
SST by [`sst_composition.py`](results/change-0576/sst_composition.py), so the two
allocations are accounted for rather than rounded away. `WithCustomViews.xls`
removes 1,242 for 474 strings rather than 948, because its strings are long
(101,125 bytes of SST across 474 strings) and `String::from_utf16` reallocated
while growing.

This is the one measurement that is **not** just faster: an open of a
string-heavy workbook now allocates a number of times bounded by the workbook's
structure rather than by its string count.

## Validation preserved

Nothing about which inputs are refused, when, or with what text has changed.

### The error-identity trap, and how it is closed

`String::from_utf16` fails with `FromUtf16Error`, whose `Display` is
`invalid utf-16: lone surrogate found`. `char::decode_utf16` fails with
`DecodeUtf16Error`, whose `Display` is `unpaired surrogate found: d800`. The
refusal the scan has always produced is

```
Encoding error: UTF-16 decoding error: invalid utf-16: lone surrogate found
```

and producing it from the wrong type would silently change the contract.

The fast path never produces this message at all. `SurrogatePairing` only decides
a **boolean**: it carries one pending high surrogate across chunks, and therefore
across `Continue` boundaries including one that switches between compressed and
uncompressed encoding, and reports whether any surrogate was left unpaired. When
that boolean says malformed, `measure_characters` takes a **cold path**: it
rewinds the cursor to where the character data began and calls `read_characters`,
so the error object is constructed by the identical code the materializing path
uses. Cost on that path does not matter; it is only reached for an input that is
about to be refused.

`a_lone_high_surrogate_keeps_the_from_utf16_message` asserts the exact string and
additionally asserts the message does **not** contain `unpaired surrogate`.
Mutation M3 below confirms the assertion has teeth.

### Ordering is preserved, including after a defect is known

`String::from_utf16` ran **after** the whole framing walk, so a string that is
both malformed UTF-16 and badly framed reported the framing error. The measure
walk reproduces that: `SurrogatePairing` records the defect and **keeps walking**;
the cold path runs only once the walk has completed successfully. Two tests pin
it — a `Continue` record that ends inside a code unit, and a `Continue` record
with invalid continuation flags, each following a record that already contains an
unpaired high surrogate. Both still report the framing error.

Likewise `read_characters` ran before `read_formatting_runs` and before the
`ExtRst` block, and still does:
`a_malformed_unit_is_reported_before_a_bad_formatting_run` pins that the encoding
refusal wins over a formatting run past the text.

### Every other refusal keeps its position

`a UTF-16 shared string is split inside a code unit`,
`shared string character data does not end at a record boundary` and
`invalid shared string continuation flags 0x%02X` are byte-for-byte the same
expressions, in the same order, inside the same loop, now in `walk_characters`.
The middle one of those three cannot fire, before or after; that is a
pre-existing property and is set out two sections below.
The rich-text run validation and the `ExtRst` phonetic block are **not** changed
at all: the measure path still calls `read_formatting_runs` and still reads and
parses the `ExtRst` payload, so their refusals and their allocation-failure
messages stay exactly where they were. That is deliberate conservatism, and it is
also cheap: read straight from the fixtures' own SSTs, **no** shared string in
any of the three carries an `ExtRst` block, and rich text covers 1 of 7,893
strings in `54016.xls`, 28 of 474 in `WithCustomViews.xls` and 3 of 303 in the
flagship. The only validation this change reimplements is the UTF-16 check.

### One error becomes unreachable on the measure path, and this is stated rather than hidden

`cannot allocate shared string characters: …` is raised when the `Vec<u16>`
reservation fails. **The measure path makes no such reservation, so on that path
the error is unreachable.** It remains reachable, in the same position and with
the same text, on the materializing path — that is, on `decode_shared_string_entry`
when a cell is read, on `SharedStringTable::parse`, and on the cold path a
malformed string takes.

The consequence is exact and bounded. Under an allocation failure that the old
code would have reported before walking, the new code walks first; if the framing
is also malformed it now reports the framing error instead of the allocation
error. The reservation is at most `65,535 × 2` bytes, because the character count
is a `u16`, so this divergence needs a 128 KiB allocation to fail. No test can
reach it and none is claimed.

### One branch is unreachable, before and after

`shared string character data does not end at a record boundary` cannot fire. To
reach it the walk needs `chunk_characters < wanted`, hence
`chunk_characters == available_characters`, which leaves `remaining % bytes_per_character`
bytes; that is zero for compressed data, and for uncompressed data an odd
remainder has already been refused one branch earlier as a split code unit. This
is a **pre-existing** property of the arithmetic, not something this change
introduces — the condition and its position are unchanged — and it is recorded
here because a reader comparing the two versions should not mistake an untested
branch for a regression.

### Policy

No `unsafe` in any production crate; `crates/litchi-xls/src/lib.rs` keeps
`#![forbid(unsafe_code)]`. The one `unsafe` block added by this change is the
test-only `#[global_allocator]` in
`crates/litchi-xls/tests/sst_scan_allocations.rs`, which is a separate test
binary and follows the pattern already established by
`crates/litchi-keynote/tests/animation_allocations.rs`. Fallible allocation is
unchanged — the reservation that remains is still `try_reserve_exact`. No public
API, no typed error, no ADR-governed boundary and no dependency edge changes.

## Correctness evidence

### The differential change 0574 asked for

Change 0574 set the falsification criterion: the measure-only walk must reproduce
**byte-identical `entries`** over every fixture with an SST.
`every_sst_fixture_indexes_identically_both_ways` does exactly that. It frames the
SST span out of each fixture's workbook globals **by hand**, so it shares no code
with the scan it checks, then runs both instantiations of the scan and compares.

| | count |
| --- | ---: |
| `.xls` and `.xlt` fixtures under `test-data` | 126 |
| refused by the CFB container before any SST is reachable | 2 |
| carrying no SST | 3 |
| **carrying an SST** | **121** |
| indexed, with byte-identical entries both ways | **117** |
| **shared-string entries compared, all identical** | **17,434** |
| refused, with byte-identical refusals both ways | 4 |

The two that never reach an SST are `1900DateWindowing.xls` and
`1904DateWindowing.xls`, both refused by `litchi-cfb` with
`Corrupted file: FAT sector 10 is not marked FATSECT`. The four refused are the
three encrypted fixtures — `password.xls`, `xor-encryption-abc.xls`,
`35897-type4.xls`, whose SST bytes are ciphertext and fail the signed-count check
— and `57456.xls`, which declares 1,761 unique strings and 0 total. All four are
refused **identically** by both paths, which is the same evidence as an identical
index.

### Tests that fail against the pre-change code

This change is behaviour-preserving by construction, so the honest statement is
that no *behavioural* test can flip. Two **resource** tests do, and they were run
against the pre-change worktree to prove it:

| test | pre-change | post-change |
| --- | --- | --- |
| `opening_a_string_heavy_workbook_allocates_less_than_once_per_shared_string` | **FAILS**: 16,031 allocations for 7,893 strings | passes: 247 |
| `opening_a_long_string_workbook_allocates_less_than_once_per_shared_string` | **FAILS**: 1,478 allocations for 474 strings | passes: 236 |
| `extracting_text_still_decodes_shared_strings` | passes | passes |

Both assert an invariant rather than a magic number — *one open allocates fewer
times than the workbook has unique shared strings* — and read the unique count out
of the fixture's own SST record, so they do not need re-tuning when the corpus or
unrelated allocation behaviour moves.

The third is the counterweight: it pins that extracting text still decodes shared
strings, so the materializing path has not been quietly disabled.

### Mutation checks

The 17 unit tests cannot fail against the pre-change code, because the measure
path does not exist there. They are therefore justified by what they catch in the
**new** code. Six mutations were applied to the merged tree and the suite re-run:

| mutation | caught by |
| --- | --- |
| M1: forget the pending high surrogate at a compressed continuation | `a_pending_high_surrogate_cannot_be_completed_across_an_intervening_run`, `…_three_run_delivery` |
| M2: ignore a high surrogate still pending at the end of the string | 7 tests |
| M3: report the refusal from the walk, with `decode_utf16`'s wording, instead of re-walking through `String::from_utf16` | 9 tests |
| M4: take the chunk-level fast exit even when a high surrogate is pending | `a_pending_high_surrogate_cannot_be_completed_across_an_intervening_run`, `…_three_run_delivery` |
| M5: swap which surrogate half opens a pair | both differential sweeps |
| M6: drop the re-examination of the unit that failed to complete a pair | **nothing — and it is equivalent** |

**M1 and M4 initially survived, and that found a real gap.** Both need a string
delivered in *three* runs — a high surrogate, then a run that cannot complete it,
then a low surrogate — because with only two runs a stale pending surrogate is
still unpaired at the end and the answer comes out right for the wrong reason.
Without the three-run tests, either mutation would have made the scan *accept* a
string `String::from_utf16` rejects, which is precisely the silent contract break
this change had to avoid. Three tests were added: the targeted shape, the
boundary case where an **empty** intervening run still lets the pair form, and a
640-case sweep over every three-run delivery.

M6 survives because it is genuinely equivalent: once `malformed` is set it can
never be cleared, so re-examining the offending unit cannot change the answer.
The re-examination is kept because it mirrors `char::decode_utf16`'s shape, and
the code comment now says it cannot change the answer.

The two exhaustive sweeps compare the two instantiations over **2,840** and
**640** deliveries of short code-unit sequences drawn from both surrogate halves,
both extremes of each half, the units immediately outside the surrogate block, a
plain unit and a compressed-representable unit — as a single record, split at
every position into two records, and delivered as three runs with mixed encodings.

The corpus supplies the same coverage on real data rather than only synthetic
data: **17** of `54016.xls`'s shared strings and **11** of
`WithCustomViews.xls`'s genuinely cross a `Continue` boundary, so the
continuation path the surrogate state has to survive is exercised by the
differential, not just by the sweeps.

### Suites

| slice | result |
| --- | --- |
| `litchi-xls`, all features | **1,379 tests, 0 failures**, 72 binaries |
| `litchi-xls`, `litchi-cfb`, `litchi-biff`, `litchi-ole-common`, `litchi-xlsx`, `litchi-xlsb` | **3,928 tests, 0 failures**, 160 binaries |
| `cargo fmt --check -p litchi-xls` | clean |
| `cargo clippy -p litchi-xls --all-features --all-targets -- -D warnings` | clean |
| `python3 -B tools/non_iwork_gate.py clippy` | clean, all 45 bulk packages plus the facade in both feature configurations |
| `litchi` facade library tests | 380 pass, **3 fail** — see below |

The three facade failures are `document::doc::tests::{filesystem_odt_keeps_source_owner_with_malformed_ooxml_catalog,
owned_odt_bytes_keep_odt_owner_with_malformed_ooxml_catalog,
managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal}`. They are
DOCX/ODT/OPC failures and are **pre-existing at HEAD**: they were reproduced in a
completely pristine worktree at `163ac1bd6` with no working-tree change of any
kind applied. They are recorded here rather than omitted, and they are not this
change's to fix — the same disposition `tools/check_example_targets.py` already
has for three duplicate example targets in the iWork crates.

## What this does not do

**It does not defer SST decoding.** The scan still visits every shared string at
open time to establish its extent, and after this change that visit is where the
remaining 252,497 and 3,724,624 instructions per open go on the two string-heavy
fixtures. A design that indexes lazily — or that stores enough to reconstruct
offsets without walking — is the larger follow-up, and it changes *when* a
malformed SST is refused, so it needs its own frozen design record of the kind
changes 0565 and 0566 wrote. This record deliberately does not establish it.

It does not touch opportunities 2 and 3 of change 0574: the globals payloads the
semantic pass never interprets, and the globals scan's re-walk of the allocation
chain. Both are still open, and both are now a larger share of a smaller open —
`next_chain_sector` and `read_stream_range` are unchanged at 139,104 and 120,463
instructions per flagship open, which is 10.2% of the post-change open against
9.3% before.

It changes nothing on a workbook without an SST, and nothing on DOC, PPT, XLSX or
any non-OLE2 path.

## Limitations

No cold-cache, physical-device, remote or range-source, peak-RSS,
concurrency-scaling, real-producer or cross-platform result is claimed. The
latency figures are warm-cache, page-cached, on a staged immutable copy, on one
host, from two binaries that differ in one file, and **host quiescence is not
established**: the load average was 3.5 to 4.9 on 32 cores throughout. The
A/A and B/B columns are the defence against that and are reported for every cell.

The allocation figures are counts and requested bytes, not live bytes and not
RSS. Peak live bytes almost certainly fall too — the removed allocations were
transient per-string buffers — but that is not what this counter measures and no
peak claim follows.

Callgrind instruction counts carry change 0574's caveat unchanged: the ERMS
string loops are instrumented per iteration, so `memset` and `memcpy` shares are
upper bounds. This change does not touch either, and the symbols that moved are
not string loops, so the caveat does not affect the attribution above.

Three fixtures dominate the latency story and they were chosen because change
0574's profile named them. No corpus-wide latency regression was run for this
change; the corpus-wide statement here is the **differential**, which covers
every fixture, not a timing.

`SharedStringTable::parse_segments` still carries a second copy of the per-string
record layout. It is pre-existing, it is not made worse, and unifying it is a
separate change.

[Captures, profiles, the differential log and the replay script](results/change-0576/README.md).

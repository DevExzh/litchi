# 0753: The fresh DOC, PPT and XLS writers encode each text once into its final buffer and hash each shared string once: payload-heavy writes take 0.29×, 0.19× and 0.37× the time, byte for byte the same

Status: retained, implemented. `performance_claim: none` — this record carries
paired ABBA timings from two windows, differenced hardware counters, exact
callgrind instruction counts and counting-allocator figures. They are reported
as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

Base `6d989cad63`; branch `perf/0753-legacy-fresh-writer-text-paths`. Production
commits: `c8984bce7b` (DOC), `f1531b1e26` (PPT), `f6ad75ed89` (XLS),
`88d2c18a9c` (buffer reservations only for large text), `630a24b17b` (XLS cell
sort) and `4c079ef4bf` (PPT reservation without a text scan); `42daeb4bc7` adds
tests only. Every measurement below is of `4c079ef4bf`. Evidence:
[results/change-0753](results/change-0753/README.md).

## Result

The three legacy fresh writers no longer walk text several times per character,
copy it through a chain of intermediate buffers, or hash a shared string three
times plus once per map resize. **Every output byte is unchanged**: the same
fixtures produce the same SHA-256 on the base and on this branch, and the
harness verified each process's output against its corpus hash.

| selector (`--writer-shape`) | before p50 | after p50 | paired ratio, window 1 [95% CI] | window 2 | instructions per write (callgrind) |
| --- | ---: | ---: | ---: | ---: | ---: |
| `doc_fresh_write_to` (payload-heavy) | 5.683 ms | 1.628 ms | **0.286** [0.285, 0.287] | 0.286 | 133.4 M → 12.2 M (**0.091**) |
| `doc_fresh_write_to` (large) | 0.1834 ms | 0.1442 ms | 0.787 [0.779, 0.793] | 0.788 | 4.99 M → 3.38 M (0.676) |
| `doc_fresh_write_to` (tiny) | 6.6 µs | 5.8 µs | 0.883 [0.873, 0.890] | 0.884 | 136.5 k → 115.9 k (0.849) |
| `ppt_fresh_write_to` (payload-heavy) | 4.504 ms | 0.860 ms | **0.191** [0.180, 0.197] | 0.200 | 162.3 M → 25.9 M (**0.160**) |
| `ppt_fresh_write_to` (large) | 0.1390 ms | 0.0733 ms | **0.527** [0.525, 0.529] | 0.532 | 2.85 M → 1.47 M (0.515) |
| `ppt_fresh_write_to` (tiny) | 10.5 µs | 9.5 µs | 0.910 [0.906, 0.914] | 0.913 | 209.6 k → 184.9 k (0.882) |
| `xls_fresh_write_to` (payload-heavy) | 6.227 ms | 2.321 ms | **0.374** [0.370, 0.388] | 0.387 | 85.2 M → 36.2 M (**0.425**) |
| `xls_fresh_write_to` (large, all numeric) | 1.205 ms | 1.102 ms | 0.915 [0.912, 0.941] | 0.951 | 18.50 M → 18.64 M (1.007) |
| `xls_fresh_write_to` (tiny, all numeric) | 4.8 µs | 4.9 µs | 1.011 [1.004, 1.023] | 1.014 | 82.7 k → 82.6 k (0.999) |

Allocated bytes per write fall by 13–73% on every DOC and PPT shape and by 30%
on the XLS payload-heavy shape; no shape's peak live heap rises, and the XLS
payload-heavy peak falls 38% (section "Allocations"). The two all-numeric XLS
shapes, which never reach the shared-string code, execute 0.7% more
instructions on `large`; their wall-time ratios (0.915–1.014) move with the
build's code layout (section "Regression flags").

## Why this record exists

Profile r2 (period-weighted `perf` of the timed regions at `009d515bef`; the
writer and harness code is identical at `6d989cad63`) found:

- `doc_fresh_write_to` (5.9 ms): per-character UTF-16 work about 60% of the
  timed cycles — `encode_utf16().count()` in `utf16_code_unit_len`, a
  per-character CP walk for field characters, and a per-unit `extend_from_slice`
  whose capacity check alone was 12.7% — plus 29% kernel page faults.
- `ppt_fresh_write_to` (4.9 ms): 27% re-counting UTF-16 units of text already
  stored as ASCII bytes, 31% `memcpy`, 34% page faults.
- `xls_fresh_write_to` (3.9 ms): 45% SipHash (every shared string hashed by
  `contains_key`, again by `insert`, again on each resize of an unreserved map,
  and again with a full comparison by the `LabelSst` lookup), 37% page faults.
- All three touched 12–17 MB of fresh memory per write for 4–5 MB of output.

This branch's own frame-pointer profiles of the base confirm the split
(`profiles/payload-heavy-attribution.json`): UTF-16 counting or decoding is 24%
of the DOC write and 28% of the PPT write, `memmove` 48% of the PPT write, and
SipHash 46% of the XLS write.

## What changed

Scope: `crates/litchi-doc`, `crates/litchi-ppt`, `crates/litchi-xls` (writers
only). No public API, durable format, limit or dependency changed; no `unsafe`;
litchi-cfb is untouched (other implementers own it this wave).

### DOC (`c8984bce7b`, `88d2c18a9c`)

`crates/litchi-doc/src/writer/core/model/codec.rs`:

- `utf16_units` counts UTF-16 code units without decoding: ASCII by its length,
  other text by summing, per byte, one for each non-continuation byte and one
  more for each four-byte lead byte, in 127-byte blocks with a `u8` accumulator
  (at most 254 per block, so it cannot wrap). `utf16_code_unit_len` keeps its two
  refusals on top of it.
- `TextStream` is the WordDocument stream under construction: the zeroed FIB
  placeholder, then the text. Its `text_len` and `text` exclude the placeholder,
  so every story builder computes exactly the FCs, offsets and overflow errors it
  computed on a separate text buffer. `push_utf16le` widens ASCII through
  `flat_map(|byte| [byte, 0])` (a `TrustedLen` iterator, one reservation, no
  per-unit capacity check) and encodes other text in one pass; `append_utf16le`
  checks the length first and appends nothing on refusal.
- `contains_field_character` scans bytes for U+0013–U+0015 in 256-byte blocks
  without an early exit inside a block, so it vectorizes (a first `any()`
  version was 20% of a development profile, not retained).

`crates/litchi-doc/src/writer/core/package/package.rs`:

- The main-story run loop checks the length before building the run's grpprl,
  as before, walks characters for field CPs only when the byte scan finds a field
  character, and appends with `push_utf16le`. Text boxes use `append_utf16le`.
- `word_document_capacity_hint` bounds the text from the model (UTF-16 units of
  every run, cell, note, comment, header/footer and text box, plus four marks per
  paragraph, run, cell, row, note, comment and text box). Below 32 KiB of text it
  reserves the placeholder and that bound only, and what follows grows the stream
  as before; above, it also reserves one FKP page per sixteen items, the SEPX and
  the final 4 KiB padding. The text never exceeded its bound in any of the 1,199
  litchi-doc tests (checked with a temporary assertion, then removed); a short
  estimate only costs growth.
- After the stories, the CHPX and PAPX FKP pages and the SEPX are built first, in
  the order the stream used to be appended, and appended where the text ends.
  Their first-page numbers and `fcMac` come from the same arithmetic the lengths
  used to give. The text is therefore written once, in place, instead of being
  grown by doubling, copied behind the placeholder, and doubled again.
- `populate_compound_document` hands the streams to the CFB writer with
  `create_stream_owned` instead of copying them; the FKP builders take the grpprls
  by value instead of cloning them.

`crates/litchi-doc/src/writer/core/package/semantic/{tables.rs,stories/*.rs}`,
`writer/core/codec.rs`: notes, comments, tables, headers and header text boxes
append through `TextStream`. The table and header walks keep their per-character
loops for runs with a field character. Without one, those loops recorded nothing
and could only refuse a CP that passes 32 bits; that refusal is decided by the
run's last character (tables) or its end (headers), and the new code makes that
one check with the same message.

### PPT (`f1531b1e26`, `88d2c18a9c`, `4c079ef4bf`)

- `escher/codec/text.rs`: `append_client_textbox_with_interactions` writes the
  `ClientTextbox` in one pass: header placeholder, `TextHeaderAtom`, the text atom
  (ASCII bytes counted by length; otherwise encoded once and counted from the
  encoded length), `StyleTextPropAtom`, interactions, then the container length.
  Refusals keep their order (the text is encoded before it is counted, the count
  checked before the style atom and the interactions).
- `records/model.rs`: `InPlaceRecord` writes a record header placeholder and
  patches its length after the body is appended, with `RecordBuilder`'s bytes and
  truncating 32-bit length.
- `escher/codec/shapes.rs`, `escher/codec/drawing.rs`, `core/package.rs`: the
  Slide record, its `PPDrawing`, the `DgContainer`, the root `SpgrContainer` and
  each user `SpContainer` are written straight into the document stream in both
  `write_to` and `save`. Each slide's offset is still checked after its children,
  so the first refusal is the one it was. The Vec-returning builders remain for
  notes, groups and tests. A text byte is now copied into the document stream
  once, instead of once per record builder and nesting level (about fifteen
  copies, by code inspection).
- `core/package.rs`: `document_stream_capacity_hint` reserves the stream when the
  text boxes hold at least 32 KiB of text, sized from their UTF-8 lengths without
  scanning them (`4c079ef4bf`; a first version ran `is_ascii` over every text,
  6.2% of the payload-heavy write in a development profile, not retained; the
  retained counters and callgrind figures are of the final version).

### XLS (`f6ad75ed89`, `630a24b17b`)

- `writer/core/stream/shared_strings.rs` (new): `SharedStringTable` is staged per
  write and borrows the cell strings, so nothing is cloned. Its map is keyed by
  the string and its hash, computed once as `RandomState::hash_one`: the same
  randomly keyed SipHash-1-3 a `HashMap<&str, _>` computes on every operation.
  The map uses that value as the bucket hash as-is, so insert, lookup and resize
  hash nothing again, and flooding resistance is the standard map's. The table
  also records every string cell's index in cell-map iteration order.
- `writer/core/stream/codec.rs`: each cell record carries its ordinal among the
  worksheet's string cells in the same iteration order; `index_for` returns the
  recorded index when the table's string at that index equals the cell's value
  (strings in the table are distinct, so that is the value's own index) and
  otherwise falls back to the hashed lookup. The cell keys are distinct map keys,
  so the sort is now unstable with an identical result, and a worksheet with no
  string cell skips the ordinal bookkeeping. The workbook stream reserves the SST
  and the worksheet records after it.
- `writer/biff/sst.rs`: non-ASCII strings are encoded once into one reused buffer
  (previously a `Vec<u16>` and a byte vector per string); `write_sst` takes any
  `&[impl AsRef<str>]`.
- The `Writer` fields `shared_strings`, `string_map` and `sst_total` are gone; the
  table is built from `&self` inside each write.

## Why it is sound

**Bytes.** Goldens were captured by running each golden test file on the base
first (it failed with the actual digests) and pinned; the final test files pass
on the base worktree and on this branch (gates):

- DOC, 9 fixtures: ASCII, Latin-1, CJK, supplementary-plane, empty and >64 KiB
  runs; field characters next to surrogate pairs in main, table and header
  stories; tables; headers and footers with fields; footnotes, endnotes and
  comments; main and header text boxes with CR/LF variants; a glossary document
  and an attached glossary.
- PPT, 9 fixtures: the same text classes; text interactions and rich text;
  tables, a picture, comments, a transition, a timing and a per-slide footer;
  notes written through `save`; the three benchmark shapes; an empty
  presentation.
- XLS, 6 fixtures with at most one distinct string per worksheet (below).

The harness corpora are pinned too: every measured process's output SHA-256
equals its leg's other processes and the base's (`same-output=True` in every
latency summary).

**XLS order and indices.** The SST lists strings in the order a worksheet's
`HashMap` iterates its cells, and that order depends on the map's per-process
seed, so a multi-string worksheet has no cross-process golden. The table's unit
tests therefore run the previous `contains_key`/`insert`/clone algorithm in
process over the same maps and require the same strings, order, total and every
cell's index, including for recorded indices; a read-back test checks all 3,600
string cells of a four-sheet workbook with repeated, empty, non-ASCII and long
strings; and a differential test keeps the previous SST encoder verbatim and
compares its bytes, including CONTINUE boundaries inside surrogate pairs and
strings past 0xFFFF code units.

**Refusals and order.** Every length check still happens before the work it
guards: DOC checks a run's length before its grpprl, PPT counts the text after
encoding it and before the style atom. The overflow checks the skipped
per-character walks could make are made once with the same messages. PPT's
in-place records leave a partial stream only on error, and every writer then
drops it: a test refuses a write after a slide's drawing is already in the
stream and finds the destination untouched. Writers written twice produce
identical bytes, and an XLS writer extended between writes stages its strings
again (tests pass on both legs).

**Hashing.** No hash algorithm changed. The only new hasher passes through a
value that is itself a randomly keyed SipHash-1-3 of the string, which is what
`std`'s `HashMap` feeds its table; a collision test keeps equal cached hashes
with different strings apart.

**Determinism.** Output order is unchanged. `RandomState` only seeds the table's
private map, whose iteration order never reaches the output.

## Authority and constraints

- ADR 0005: allocation and copy reductions measured, reservations exact or
  well-estimated, figures not claims.
- ADR 0006: preservation by default; output byte-identical and deterministic
  where it was.
- GOAL.md: no non-DoS-resistant hash for untrusted keys; no `unsafe`; limits and
  typed errors unchanged; allocation behaviour unchanged (these writers grow
  infallibly, as before).
- Change 0652's standing trade-offs 2 (correctness first) and 3 (optimize the
  benign common path). No owner decision was needed: nothing public changed.

## Measured

Harness: `tools/perf-baseline` at each leg, built with the identical command
(`scripts/build_leg.sh`: `cargo build --release --locked --offline
--manifest-path tools/perf-baseline/Cargo.toml`, `CARGO_BUILD_JOBS=6`); before
from a detached worktree at `6d989cad63`, after at `4c079ef4bf`. SHA-256:
`b700160a…2667db` (before) and `d56a416f…a19d23` (after); the allocation probes
`3e2b3674…84c234` and `b038d559…b1bc1bc9` (`binaries.sha256`). Every process pinned
to CPU 8; binaries at equal-length paths (`bin/A/lpb`, `bin/B/lpb`).

### Paired wall time

ABBA (A B B A A B B A) per case; tiny 100 warmups + 1,000 samples, large
10–20 + 100–200, payload-heavy 5 + 40; median of per-process p50s, median of the
four adjacent-pair p50 ratios, percentile bootstrap. Window 1 09:50–09:52Z,
window 2 09:52–09:53Z, host shared with other implementers (load average 6–11).

| selector | shape | before p50 ms | after p50 ms | paired ratio | 95% CI | window 2 ratio | 95% CI |
| --- | --- | ---: | ---: | ---: | --- | ---: | --- |
| `doc_fresh_write_to` | tiny | 0.0066 | 0.0058 | **0.883** | [0.873, 0.890] | 0.884 | [0.876, 0.889] |
| `doc_fresh_write_to` | large | 0.1834 | 0.1442 | **0.787** | [0.779, 0.793] | 0.788 | [0.778, 0.798] |
| `doc_fresh_write_to` | payload-heavy | 5.683 | 1.628 | **0.286** | [0.285, 0.287] | 0.286 | [0.286, 0.287] |
| `ppt_fresh_write_to` | tiny | 0.0105 | 0.0095 | **0.910** | [0.906, 0.914] | 0.913 | [0.902, 0.923] |
| `ppt_fresh_write_to` | large | 0.1390 | 0.0733 | **0.527** | [0.525, 0.529] | 0.532 | [0.528, 0.539] |
| `ppt_fresh_write_to` | payload-heavy | 4.504 | 0.8601 | **0.191** | [0.180, 0.197] | 0.200 | [0.195, 0.201] |
| `xls_fresh_write_to` | tiny | 0.0048 | 0.0049 | **1.011** | [1.004, 1.023] | 1.014 | [0.986, 1.017] |
| `xls_fresh_write_to` | large | 1.205 | 1.102 | **0.915** | [0.912, 0.941] | 0.951 | [0.911, 0.996] |
| `xls_fresh_write_to` | payload-heavy | 6.227 | 2.321 | **0.374** | [0.370, 0.388] | 0.387 | [0.379, 0.398] |
| `doc_semantic_one_edit_save` (control) | tiny | 0.0360 | 0.0357 | 0.993 | [0.984, 1.000] | — | — |
| `doc_semantic_one_edit_save` (control) | large | 0.7795 | 0.6424 | 0.827 | [0.815, 0.831] | — | — |
| `ppt_semantic_one_edit_save` (control) | tiny | 0.0775 | 0.0780 | 1.007 | [0.996, 1.010] | — | — |
| `ppt_semantic_one_edit_save` (control) | large | 0.1867 | 0.1876 | 1.005 | [1.001, 1.008] | — | — |
| `xls_semantic_one_edit_save` (control) | tiny | 0.0492 | 0.0497 | 1.010 | [1.004, 1.018] | — | — |
| `xls_semantic_one_edit_save` (control) | large | 1.545 | 1.541 | 0.995 | [0.978, 1.039] | — | — |

The timed region is the harness's `write_fresh_*`, which also generates the text
and builds the writer; those parts are unchanged. Payload-heavy absolute times
depend on the host's page-fault cost: this kernel zeroes pages at allocation
(`CONFIG_INIT_ON_ALLOC_DEFAULT_ON=y`) and free memory was low, and in an earlier,
busier window with intermediate builds the same cases measured 8.8–11.8 ms before
and 4.7 ms after for DOC, and 7.6–10.4 ms before and 0.85 ms after for PPT, while
the paired ratios stayed below 0.54 and 0.12. The ratios, not the absolute times,
are the comparable figures. The earlier window and the intermediate builds are
development measurements and are not retained; the tables are of the final
build only. `xls_fresh_write_to` large, whose code path gains only the ordinal
bookkeeping, measured 0.915 and 0.951 in the two retained windows and 0.997 and
0.931 in two development windows with the `630a24b17b` build (same XLS code):
its wall-time ratio is within the layout variation of the whole binary.

### Hardware counters per iteration

`perf stat` of each leg at 20 and 220 samples (200 and 2,200 for tiny), no
warmup, A B B A, differenced and averaged (`counters/`). An iteration includes
the harness's untimed output comparison, identical in both legs. User-cycle
differences of the tiny shapes are within the differencing noise (DOC tiny +13.4%
cycles with −12.4% instructions and a 0.883 wall-time ratio; XLS tiny +11.8% with
+0.1%). The kernel columns of the tiny and large shapes are a few tens of
thousands of instructions and move by that noise.

| selector | shape | user instructions | user cycles | kernel instructions | page faults |
| --- | --- | --- | --- | --- | --- |
| `doc_fresh_write_to` | tiny | 109,990 → 96,385 (−12.4%) | 29,355 → 33,295 (+13.4%) | 31,962 → 33,463 (+4.7%) | 0 → 0 |
| `doc_fresh_write_to` | large | 4,334,017 → 3,091,972 (−28.7%) | 854,811 → 627,267 (−26.6%) | 19,788 → 22,777 (+15.1%) | 0 → 0 |
| `doc_fresh_write_to` | payload-heavy | 127,359,998 → 8,392,913 (**−93.4%**) | 21,215,877 → 4,404,283 (−79.2%) | 17,288,633 → 13,153,883 (−23.9%) | 3,289 → 2,516 |
| `ppt_fresh_write_to` | tiny | 201,188 → 178,871 (−11.1%) | 45,942 → 44,767 (−2.6%) | 33,496 → 33,376 (−0.4%) | 0 → 0 |
| `ppt_fresh_write_to` | large | 2,684,755 → 1,365,744 (−49.1%) | 616,317 → 293,919 (−52.3%) | 55,815 → 67,126 (+20.3%) | 3 → 2 |
| `ppt_fresh_write_to` | payload-heavy | 68,465,979 → 7,284,677 (**−89.4%**) | 15,548,281 → 3,496,648 (−77.5%) | 15,908,694 → 634,443 (−96.0%) | 3,007 → 113 |
| `xls_fresh_write_to` | tiny | 82,241 → 82,292 (+0.1%) | 20,806 → 23,264 (+11.8%) | 32,986 → 33,365 (+1.1%) | 0 → 0 |
| `xls_fresh_write_to` | large | 17,464,979 → 17,650,975 (+1.1%) | 4,879,827 → 4,809,464 (−1.4%) | 2,361,893 → 1,756,422 (−25.6%) | 435 → 320 |
| `xls_fresh_write_to` | payload-heavy | 60,565,736 → 21,374,984 (**−64.7%**) | 20,385,877 → 7,741,097 (−62.0%) | 22,522,330 → 11,120,300 (−50.6%) | 4,240 → 2,103 |
| `doc_semantic_one_edit_save` | large | 41,395,150 → 41,406,783 (+0.0%) | 11,239,433 → 11,137,212 (−0.9%) | 4,398,561 → 2,737,335 (−37.8%) | 811 → 503 |
| `ppt_semantic_one_edit_save` | large | 8,420,787 → 8,418,762 (−0.0%) | 2,410,475 → 2,321,859 (−3.7%) | 109,428 → 39,774 (−63.7%) | 5 → 0 |
| `xls_semantic_one_edit_save` | large | 162,726,385 → 163,811,422 (+0.7%) | 46,600,095 → 46,973,643 (+0.8%) | 2,372,747 → 2,511,573 (+5.9%) | 422 → 449 |

### Exact instructions per write (callgrind)

The allocation probe (`probe/main.rs`: the harness's three `write_fresh_*`
bodies, built against each leg's sources with one command) under callgrind at
two iteration counts, differenced (`callgrind/`):

| writer / shape | before | after | after / before |
| --- | ---: | ---: | ---: |
| `doc_fresh_write_to/tiny` | 136,478 | 115,885 | 0.849 |
| `doc_fresh_write_to/large` | 4,990,087 | 3,375,792 | 0.676 |
| `doc_fresh_write_to/payload-heavy` | 133,399,783 | 12,163,994 | **0.091** |
| `ppt_fresh_write_to/tiny` | 209,639 | 184,852 | 0.882 |
| `ppt_fresh_write_to/large` | 2,848,812 | 1,465,992 | 0.515 |
| `ppt_fresh_write_to/payload-heavy` | 162,340,876 | 25,895,697 | **0.160** |
| `xls_fresh_write_to/tiny` | 82,684 | 82,631 | 0.999 |
| `xls_fresh_write_to/large` | 18,504,884 | 18,642,111 | 1.007 |
| `xls_fresh_write_to/payload-heavy` | 85,212,887 | 36,191,419 | **0.425** |

### Allocations

Counting global allocator in the probe, mean of five writes after one warm-up
(`alloc/`). Peak live bytes are the operation's own heap high-water mark above
its starting point, not process RSS.

| writer / shape | allocations | reallocations | allocated bytes | peak live bytes |
| --- | --- | --- | --- | --- |
| `doc_fresh_write_to/tiny` | 132 → 126 | 46 → 39 | 72,010 → 58,188 (−19.2%) | 47,613 → 47,613 (+0.0%) |
| `doc_fresh_write_to/large` | 3,213 → 3,207 | 116 → 100 | 1,164,071 → 982,825 (−15.6%) | 807,363 → 801,110 (−0.8%) |
| `doc_fresh_write_to/payload-heavy` | 967 → 961 | 1,126 → 1,103 | 40,064,521 → 21,433,355 (−46.5%) | 21,344,973 → 18,182,488 (−14.8%) |
| `ppt_fresh_write_to/tiny` | 277 → 249 | 182 → 152 | 59,043 → 51,467 (−12.8%) | 35,027 → 35,027 (+0.0%) |
| `ppt_fresh_write_to/large` | 3,712 → 2,176 | 2,473 → 757 | 963,333 → 489,621 (−49.2%) | 272,683 → 272,683 (+0.0%) |
| `ppt_fresh_write_to/payload-heavy` | 3,620 → 2,212 | 3,475 → 1,901 | 119,832,289 → 32,500,177 (−72.9%) | 20,797,687 → 20,777,831 (−0.1%) |
| `xls_fresh_write_to/tiny` | 76 → 77 | 30 → 30 | 22,326 → 22,478 (+0.7%) | 17,461 → 17,461 (+0.0%) |
| `xls_fresh_write_to/large` | 1,294 → 1,293 | 2,137 → 2,137 | 4,097,229 → 4,031,923 (−1.6%) | 1,972,688 → 1,972,688 (+0.0%) |
| `xls_fresh_write_to/payload-heavy` | 11,315 → 11,188 | 1,462 → 1,453 | 34,373,171 → 23,936,229 (−30.4%) | 27,766,458 → 17,309,788 (−37.7%) |

An intermediate version reserved the DOC and PPT streams with fixed allowances
and raised peaks on small outputs (development probe run, not retained: PPT tiny
35,027 → 97,079 bytes, PPT large +42%, DOC tiny +17%); `88d2c18a9c` reserves only
for large text, which removed every peak increase above.

## Regression flags (every case above 5%, and every smaller measured increase)

- **No measured fresh-writer case regresses by more than 5%** in wall time,
  instructions, allocated bytes or peak live bytes.
- `xls_fresh_write_to` large (all numeric): +0.7% exact instructions per write
  (callgrind; +1.1% per harness iteration). The string ordinal travels in the
  sorted cell tuples, 24 instead of 16 bytes. The stable sort of the wider tuples
  cost +4.1% (callgrind during development, not retained), which is why the sort
  is now unstable (identical order for distinct keys). Wall time 0.915 and 0.951
  (retained windows); user cycles −1.4%.
- `xls_fresh_write_to` tiny: 1.011 [1.004, 1.023] and 1.014 [0.986, 1.017] wall
  time on a 4.8 µs write with identical instructions (82,684 → 82,631); +152
  allocated bytes (the table's per-sheet index vectors).
- Controls, which run no changed code in their timed regions:
  `xls_semantic_one_edit_save` tiny 1.010 [1.004, 1.018],
  `ppt_semantic_one_edit_save` tiny/large 1.007/1.005, and
  `xls_semantic_one_edit_save` large, whose user instructions are +0.7%, kernel
  instructions **+5.9%** and page faults +6.4% (422 → 449) per iteration with a
  0.995 [0.978, 1.039] wall-time ratio (a development run of the `630a24b17b`
  build, not retained, measured this control's kernel instructions −28.7%). The
  only changed code these processes contain is the corpus-building writer and
  the crates' codegen-unit partition, which moves with any source change; these
  shifts are reported, not attributed.
- `doc_semantic_one_edit_save` large runs **0.827** of its base time, a
  *speed-up* in a control. Its user instructions are identical (41,395,150 →
  41,406,783); its page faults fall 811 → 503 and kernel instructions 38%. The
  changed fresh writer builds this process's corpus before the timed loop, so the
  edit path runs on a different heap state. Not claimed.
- Kernel instructions of the DOC large (+15.1%) and PPT large (+20.3%) shapes rise
  by 3,000 and 11,000 instructions per iteration, inside the differencing noise of
  numbers that small.

## Findings outside this change

- **The fresh XLS writer's SST order is not reproducible across processes** for
  a worksheet with more than one distinct string: it follows
  `worksheet.cells.values()`, a `HashMap` with a per-process random seed, so two
  processes can write the same workbook with different SST orders and `LabelSst`
  indices. Both files read back the same values. This predates the change, which
  keeps the order, and conflicts with ADR 0006's determinism rule. A fix (for
  example first occurrence in worksheet, row, column order) changes output bytes
  and is left to its own record.
- The SST writer silently truncates a string at 0xFFFF UTF-16 code units (it can
  split a surrogate pair there), and `write_string` accepts longer strings.
  Unchanged here; a typed refusal would be a behaviour change.

> **Amended 2026-09-23** by
> [0757](0757-xls-fresh-writer-sst-determinism.md), after this record's
> review. The truncation is more severe than the bullet above says. When the
> 0xFFFF-unit cut falls inside a surrogate pair (for example a cell holding
> 65,534 × `a` followed by one emoji), the SST entry ends with a lone high
> surrogate, and litchi's own reader then refuses the *whole* workbook
> ("lone surrogate found"); every other over-long string silently loses its
> tail. Change 0757 fixes both findings above: the SST lists each string at its
> first occurrence in worksheet, row and column order, and the writer refuses
> a string longer than an SST entry with the typed `Error::StringTooLong`
> before producing any output. Its regression test reproduces the lone
> surrogate case. The review's other follow-ups, which change no output byte,
> are recorded there too: the DOC and PPT stream reservations are made after
> the document-wide checks and fallibly, `InPlaceRecord` is written only
> through a closure that always patches its header, and `utf16_units`,
> `contains_field_character` and `InPlaceRecord` have direct unit tests.

- The CFB writer's `write_to` zero-fills its destination as it emits sectors:
  about 37% of the remaining DOC payload-heavy time is that `memset`, mostly in
  kernel page faults (owned by litchi-cfb, which records 0748 and 0749 are
  changing).

## What is not claimed

- No claim is registered; `performance_claim: none`.
- Timing is warm, in-memory, single-CPU, on a shared host; the payload-heavy
  absolute times move with the host's page-fault cost (section "Paired wall
  time"). No cold-cache, RSS, throughput or concurrency claim is made.
  Allocation figures are the probe's counting allocator, not RSS.
- Only the three fresh writers changed. The DOC, PPT and XLS editors,
  source-backed paths, readers and the CFB writer did not; their benchmark
  movements above are layout or heap-state effects, not improvements.
- Non-ASCII text is covered for correctness (goldens, differential and read-back
  tests) but no benchmark measures its speed; its DOC length count is 5× faster
  in a microbenchmark and its encoding makes one pass instead of three, which is
  not a measured writer claim.
- The DOC and PPT reservation estimates are heuristics. A document whose
  formatting needs more FKP pages than one per sixteen items, or a presentation
  whose non-text records exceed the allowance, grows its stream as before.
- The XLS recorded-index path relies on a worksheet's cell map iterating in the
  same order during staging and emission when unmodified. The index it yields is
  checked by content equality and falls back to hashing, so correctness does not
  depend on that order; only the saved hash does.

## What is left, with prices from the after profile

Frame-pointer profiles of the finished code (`profiles/`; shares of the timed
region, payload-heavy):

- DOC (1.63 ms): 62.5% kernel leaves. The CFB writer's `write_to` is 43.8%, most
  of it a `memset` of its destination in page faults (37%); building the streams
  is 46%, most of it the first touch of the WordDocument buffer while text is
  widened into it. Both are output-sized fresh memory, and the first is
  litchi-cfb's. The harness's text generation is 6.2%.
- PPT (0.86 ms): the harness's own text generation is 26.5% of the after time;
  the in-place shape containers 18.8% (the text copied into the document stream,
  and the `is_ascii` that chooses the text atom); converting shapes to drawing
  data 15% (a clone of each text box's text into the crate-internal
  `UserShapeData.text`, avoidable only with a borrowed or shared text there); the
  CFB output copy 11.5% and its `memset` 8%; copies inlined into the harness
  frame, chiefly `add_textbox`'s copy of its input into the writer, 10.8%.
- XLS (2.32 ms): staging the shared-string table 26.6% (one keyed SipHash per
  string cell, the floor while the table deduplicates by content with a
  flooding-resistant hash); the workbook stream 32% (the SST copied into fresh
  stream pages); the CFB output copy 22.7%; the harness 7.7%.

## Verification

All in the branch worktree with `CARGO_TARGET_DIR=…/targets/0753`, at the final
production commit `4c079ef4bf` (the last section of `gates.txt`; an earlier full
run at `630a24b17b` is above it):

- `cargo fmt --all --check` — 0
- `cargo check -p litchi-doc -p litchi-ppt -p litchi-xls --all-targets --locked` — 0
- `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets --locked` — 0
- `cargo clippy -p litchi-doc -p litchi-ppt -p litchi-xls --lib --no-deps --locked -- -D warnings` — 0
- the same with `--all-targets` — 101 on `tests/xls_query_index_cache.rs`
  (`unusual_byte_groupings`), which fails identically on the base (recorded in
  `gates.txt`, as record 0746 also noted); with that lint allowed — 0
- `cargo test -p litchi-doc` — 1,200 passed; `-p litchi-ppt` — 1,231 passed;
  `-p litchi-ppt --features encryption` — 1,242 passed; `-p litchi-xls` — 1,506
  passed; `-p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` — 382 passed
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-doc -p litchi-ppt -p litchi-xls --no-deps` — 0
- `python3 tools/check_crate_boundaries.py`, `python3 tools/non_iwork_gate.py verify`,
  `python3 tools/check_perf_claims.py … --mode structural` — 0
- the three golden test files, copied unchanged onto the base worktree — all
  seven tests pass there too (bytes identical, reuse and refusal behaviour equal)

The harness did not change, so its tests and the coverage validator were not
required.

## Cleanup

Recorded in `results/change-0753/cleanup.json`: the before-leg worktree
`0753-before-src` (removed with `git worktree remove --force`), the target
directories `targets/0753`, `0753-before`, `0753-prof`, `0753-before-prof`,
`0753-probe-A` and `0753-probe-B`, and the scratch directory
`scratch/0753/*` (binaries, perf data, folded stacks, micro-benchmarks, logs).
Only summaries, raw JSON reports, `perf stat` outputs, callgrind logs, the probe
source and the scripts are kept in the packet.

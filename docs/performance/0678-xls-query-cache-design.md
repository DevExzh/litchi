# 0678: bounded weighted cache design for repeated XLS queries

Status: design and pricing only. No production source, public API, cache, or
limit was changed by this item. `performance_claim: none`; no claim-registry
entry is made.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This item resolves the retained-index residue left by
[0605](0605-xls-retained-sheet-index.md) and
[0668](0668-xls-query-residues.md). Those records proved
that a whole-sheet visitor removes the large batch-query case and that selected
queries still repeat a validated worksheet walk, but deliberately stopped
before retaining a cross-query index. The missing piece was not another scan
shape: it was a cache with an explicit weight, eviction, pinning, and refusal
policy. This record supplies that design and prices it against the current
`litchi-xls` structures. It does not claim that the design has been
implemented.

## Current path and the remaining opportunity

The authoritative path is `crates/litchi-xls/src/workbook/source.rs`:

* `SourceBackedWorkbook` and `SourceBackedWorksheet` are cheap `Arc` handles
  over one immutable `SourceInner`. `SourceInner` already retains the validated
  sheet offsets, formatting, encoding, source identity, limits, and an
  `Arc<SharedStringSstScan>`.
* Opening parses workbook globals once. `SharedStringSstScan` retains one
  `SharedStringSstSegment` per `SST`/`Continue` payload and one
  `SharedStringEntryLocation` per unique string; it retains locators, not
  decoded string text.
* `cell()` takes a leading execution check and source fence, then
  `scan_worksheet` constructs a fresh `WorksheetScan` and a fresh
  `SharedStringResolver`. The scan frames and validates the selected sheet to
  EOF, even after it has found the requested coordinate, and takes the same
  trailing fence as the old query path.
* `WorksheetScan` currently reuses a window and one payload scratch buffer only
  for that call. `SharedStringResolver` can retain a whole SST window only for
  that call after its ninth resolve. Both are released before the next
  selected-cell query.
* `visit_cells()` already reports a sheet through one scan. Its visitor keeps
  no index, and it remains the right path for a caller asking for many cells in
  one pass.
* `text()` already shares one resolver and one worksheet-chain hint across the
  document operation. A persistent chain hint is not the proposed cache:
  `StreamChainHint<'a>` borrows `SharedOleFile`, so retaining it in
  `SourceInner` would require a self-referential owner and would mix an
  operation cursor position with snapshot state.

The repeated-query opportunity is therefore specific. After a successful
full-sheet validation, a later selected-cell query can avoid framing and
validating every preceding worksheet record. It still has to read and parse
the selected BIFF record, validate its XF, resolve its SST entry, and take the
operation's source fences. An index cannot make a malformed first query
successful, and it cannot replace `visit_cells()`.

## Proposed cache boundary

Add one private, snapshot-scoped weighted cache to `SourceInner`. The cache is
behind a private short-held `Mutex`; the lock is acquired only for lookup,
LRU touch, admission, and eviction. It is never held over a source read,
BIFF parse, visitor callback, execution check, or SST decode. Public worksheet
handles remain lifetime-free and expose no lock, archive object, or cache type.

The proposed limit is a new `SourceBackedLimits::max_query_index_bytes`, with a
small finite default of **2 MiB** and `0` meaning cache disabled. The value is
an optimization budget, not a correctness ceiling: if a candidate is too
large, its allocation fails, or all existing entries are pinned, the query
completes through today's scan and the candidate is dropped. Checked weight
arithmetic and fallible reservations are required. The default is priced below
and is intentionally large enough for the largest observed worksheet index and
its SST locator index together, while remaining far below the existing 128 MiB
source scan ceilings.

The cache entry key is private and snapshot-local:

```text
CacheKey::Worksheet { sheet_index: usize }
CacheKey::SstLocator
CacheKey::SstValue { string_index: u32 }   // optional second phase
```

The `SourceInner` identity and captured `expected_version` are part of the
cache boundary. An entry is never shared between two opened owners, even when
they use equal bytes. Each index also records the expected version that was
current when it was published; the leading `ensure_current()` remains the
authoritative source check before any cache hit is used.

### Worksheet occurrence index

The first entry type is a compact occurrence locator index, not a materialized
sheet or a map of `CellValue`s. It retains one locator for every recognized
stored-cell occurrence, including repeated coordinates:

```text
WorksheetCellIndex {
    expected_version: SourceVersion,
    slots: Box<[CellSlot]>,       // sorted by (row, column, stream_offset)
}

CellSlot {
    stream_offset: u64,           // BIFF frame header in Workbook stream space
    row: u16,
    column: u16,
    kind: u16,                    // scalar, Formula, MulRk, or MulBlank kind
    xf: u16,                      // already validated during the build scan
    ordinal: u16,                 // packed-cell ordinal; zero for scalar cells
}
```

This is the 0605 proposed slot shape; the current source does not yet have a
`CellSlot` type or a retained worksheet index. The occurrence chain is required
for semantic compatibility. `TargetCell::accept` returns early for a
non-target, but for each target occurrence it validates the record and resolves
SST or formula `STRING` semantics before setting `found`. Therefore an earlier
duplicate with a failing SST locator/read/decode, malformed pending `STRING`,
or another operation error must still refuse before a later valid duplicate can win. A
latest-only index would incorrectly turn that error into success. On a hit, the
equal-coordinate range is replayed in source order; each successful occurrence
updates the candidate result, so the last successful duplicate still wins while
earlier errors retain their current precedence. `visit_cells()` continues to
scan and report duplicates in stream order and never uses this index. An
out-of-range SST index is different: `resolve_shared_string_inner` returns
`Ok(CellValue::Error(...))`, so a later valid duplicate may overwrite that
cell value. Preserve both cases; do not turn a typed cell error into an
operation refusal. The existing `invalid_label_sst_index_is_typed_without_an_sst_read`
test pins that distinction.

This also covers the warm-query case that invalidates a latest-only shortcut:
a query for `A1` may warm the worksheet while `TargetCell` skips SST resolution
for a non-target `B1`; a later `B1` query must still see an earlier invalid
`LabelSst` occurrence before a later valid `B1`. The occurrence chain replays
that invalid locator and returns the same error, whereas a latest-only `B1`
slot would incorrectly return the later value.

The `kind`, `xf`, and packed `ordinal` keep the hit path independent of a
second worksheet-wide discovery scan. A hit seeks to `stream_offset`, reads
that frame, and runs the same private cell decode. `MulRk` and `MulBlank` are
revisited with their ordinal; they do not require one slot per BIFF frame.
For a string-valued `Formula`, every occurrence in the chain follows the
existing pending `STRING` and `CONTINUE` protocol before the next occurrence is
replayed. No formula result is guessed or cached by the index.

The logical charge is **24 bytes per published occurrence slot plus 64 bytes per
index**. This models the proposed field set on the measured 64-bit target: one
`u64` and five `u16` fields rounded to the target's 8-byte alignment. It is a
pricing rule, not a claim about a current source struct. The 64-byte charge
covers the `Arc`/entry metadata and LRU linkage as a conservative logical
overhead. It deliberately excludes allocator arena headers; a future
implementation must measure allocator bytes separately.

### SST value and locator handling

The current `SharedStringSstScan` is already a mandatory open-time catalog,
not a repeated-query residue. It is an `Arc` containing the exact source
segments and logical entry locators. A second selected-cell query reuses those
locators today; retaining another copy in a cache would remove no source read
and would duplicate memory. The design therefore keeps two concerns separate:

1. The first implementation may leave the validated `SharedStringSstScan`
   outside the query cache as mandatory snapshot catalog state. That is the
   smallest safe step and avoids turning an existing open-time validation
   result into a rebuild on every cache miss.
2. If memory profiling shows that this catalog must be evictable, retain a
   small `SstDescriptor` outside the cache: the SST wire start and end, total
   and unique counts, and the snapshot version. A `SstLocatorIndex` cache entry
   then owns the current `segments` and `entries`; an eviction rebuilds the
   exact SST/Continue region with the existing measured parser and limits.
   Rebuild failure is returned as the same typed error as the original parser,
   and no failed rebuild is cached. The descriptor must be recorded before
   dropping the first index; no source range is inferred from a missing
   segment table.

The proposed `SstLocatorIndex` weight is **24 bytes per segment, 16 bytes per
entry, plus 64 bytes overhead**. These are the actual current source fields:
`SharedStringSstSegment` is `(u64, usize, usize)` and
`SharedStringEntryLocation` is `(usize, usize)` on the target. The index stores
no string text. The optional second-phase `SstValue` entry is different: it
stores a successfully decoded `String` keyed by `u32`, charged as its UTF-8
bytes plus 32 bytes of entry overhead. It is evictable and useful for repeated
queries of the same `LabelSst`, but it must not be conflated with the locator
index already retained by the owner.

Only successful, clean string values are eligible for `SstValue`. Invalid SST
indices, `CellValue::Error` results, malformed decode results, source errors,
and cancellation are never cached. The value cache is optional evidence-led
follow-up work; the worksheet locator index is the primary 0678 scope.

## Admission, release, and pinning

The cache uses deterministic weighted LRU. Every entry stores its checked
weight, key, and `Arc` value. A cache hit clones the `Arc` while holding the
lock, moves the key to the MRU end, and releases the lock before doing any
source work. The cloned `Arc` pins the entry. Eviction removes only entries
whose cache-owned `Arc` is the last strong reference; an active query or
visitor that holds a clone cannot be evicted underneath it.

A worksheet candidate is built in private scan state and is published only
after all of these have succeeded:

1. every worksheet frame has passed the existing framing, byte, record-count,
   payload, XF, applicable formula-continuation, and unsupported-record rules;
2. the scan reaches worksheet EOF and the existing trailing execution and
   source fences succeed;
3. the candidate has a checked weight within the configured budget and its
   fallible reservations completed.

No error, partial index, or visitor-aborted scan enters the cache. If a source
mutation or cancellation occurs, the candidate is dropped. A candidate that
would exceed the cache budget is abandoned while the scan continues without
changing the query result; it is not a typed resource failure because caching
is an optional optimization. This follows the existing retained-SST-window
fallback rule in 0648 while keeping all mandatory scan allocations typed.

### Resident weight versus candidate peak

`max_query_index_bytes` is the owner-local ceiling for retained cache entries;
it is not a claim that a `2 MiB` value is the whole peak of a build. The private
cache state therefore tracks both `resident_weight` and `reserved_weight`.
`resident_weight` is the checked weight of published entries. A candidate's
reserved amount includes its current slot-vector capacity, index metadata,
sorting or duplicate-chain scratch, and any candidate-owned staging buffers.
The admission check is:

```text
resident_weight + all_in_flight_candidate_weights <= max_query_index_bytes
```

The candidate weight grows before each fallible allocation and is released on
candidate drop. At publication, the final slot capacity moves from
`reserved_weight` to `resident_weight`; unused capacity and scratch are
released. A losing concurrent builder releases its full reservation. This
state-level token is needed even when two different execution contexts build
the same worksheet at once; a per-entry weight checked only at publish would
temporarily allow duplicate builders to exceed the owner ceiling.

The candidate's peak also has to account for the current scan's live transient
state while slots are being collected: the worksheet window is capped at 64 KiB,
one BIFF payload is capped at 8,224 bytes, and the shared-string resolver may
retain a 256 KiB table window. Pending formula `STRING` continuation storage
and any sort or staging capacity are additional checked terms. These are
pricing constants from the current source, not permission to omit larger
configured continuation or allocator capacity; a future builder must use
checked arithmetic and either reserve the configured bound or abandon optional
index collection when the bound cannot be admitted.

The current `WorksheetScan` and `SharedStringResolver` have fallible allocation
checks but do not yet charge those buffers to an `ExecutionContext`; this
design does not claim that they are already managed by the hierarchical
budget. A cache implementation must include its candidate-owned share of those
live buffers in the managed peak, or leave the existing scan path unchanged
and make the cache builder's additional peak explicit in its own reservation.

When a query has an `ExecutionContext`, every candidate allocation reserves
`Resource::Memory` through `ExecutionContext::reserve` before it begins. The
reservation covers final resident capacity and candidate scratch, remains live
through build and publication, and the resident portion is held by the cache
entry until eviction so the hierarchical parent budget sees retained bytes.
The reservation belongs to the reference-counted entry value, not merely its
map slot: a pinned handle keeps the originating execution budget charged until
its last strong reference drops, even after owner teardown. A later query's
context does not transfer or release that earlier reservation. Local clean
cache weight must not be reported as total live memory across pinned handles.
Scratch reservations are dropped once publication no longer needs them. When
the context is absent, the owner-local weighted token and fallible allocation
are still mandatory. A failed optional reservation drops candidate state and
falls back to the ordinary query; it must not turn a cache admission miss into
a new result or refusal. No cache lock is held while a reservation, source
read, parse, or SST decode is in progress.

The first-query allocation policy needs to be explicit. Building a 935 KiB
index for a caller that asks for one cell once is wasteful, so the recommended
admission trigger is **two observations of the same worksheet on one
snapshot**:

* the first selected-cell miss runs today's path and increments a bounded
  private hotness counter;
* the second miss runs the same full validation scan while collecting a
  candidate, serves that query from the scan, and publishes the index only
  after success;
* the third and later queries use the index, subject to normal eviction.

This trigger keeps the index out of singleton and existing whole-sheet
workloads. A future measurement may show that first-scan collection is cheap
enough to publish immediately for large sheets, but that is an implementation
choice gated by a paired first-query allocation and p95/p99 result, not an
assumption in this design. The hotness counters are fixed-size per worksheet
state, saturating, and charged as cache metadata; they are not an unbounded
query-key map.

Concurrent misses may build duplicate candidates. The publish lock chooses the
already-published entry deterministically and drops the losing candidate. A
future implementation may add a per-key single-flight state only if contention
measurements justify it; no lock may span I/O to avoid turning a cache into a
worksheet-wide mutex.

When pressure requires eviction, the cache walks LRU from oldest to newest,
skips pinned entries, and removes enough unpinned weight to admit the new
entry. If the sum of unpinned entries is insufficient, admission is skipped.
There is no blocking wait for a pin to release. Dropping the last `Arc` releases
the slots or decoded string bytes immediately; no dirty edit state exists in
this read-only owner, so there is no dirty value that can be silently evicted.

## Query-hit semantics and error identity

`query_cell` remains fenced in the same order:

1. check the optional execution context;
2. call `ensure_current()` before cache lookup;
3. reject coordinates outside BIFF8's query range exactly as today;
4. use a hit if present, otherwise run the existing complete scan;
5. on a hit, replay every locator for the coordinate in source order, parse and
   validate each selected frame, resolve SST text or the formula `STRING` result
   for each occurrence, and retain the last successful value;
6. call the existing trailing fence before returning a successful hit.

The index is evidence that one earlier scan reached EOF successfully. It does
not authorize a caller to skip a failed scan. A malformed worksheet tail,
missing target `STRING`, bad XF, invalid packed range, unsupported formula
metadata, worksheet boundary error, or scan-limit error prevents admission
when the existing scan applies that check. A later hit on an admitted
immutable snapshot cannot newly encounter a structural defect at an unselected
record: the full scan already validated its framing and applicable checks. A
target occurrence is still parsed and resolved on the hit, so target-specific
SST and formula errors retain their current precedence. A source read or
source-version error while serving the selected frame still returns through
the existing typed mapping.

The cache never weakens source freshness. A source change before lookup is
reported by the leading fence. A source change during an indexed payload or
SST read is reported by the cursor/final fence, and the hit does not publish or
retain new state. The snapshot's version is not refreshed in place. If the
source is reopened, the new `SourceInner` receives a new cache.

Worksheet scan limits remain limits of the owner that built the index. A hit
does not charge a second full scan against
`max_worksheet_scan_bytes`/`max_worksheet_scan_records`, but it also cannot
serve an entry built under another `SourceBackedLimits` value because limits
are owned by the snapshot and the cache is snapshot-local. A configured cache
budget that refuses admission falls back to the current per-query limits and
result.

The cached locator is private. It does not expose CFB IDs, stream offsets,
archive handles, locks, or a new public index API, so ADR 0001's public-layer
rule and ADR 0002's crate ownership stay unchanged. No writer or edit path uses
the cache; exact no-op and source-preservation rules therefore remain outside
the cache's read-only scope.

## Corpus pricing

The independent probe is
[`results/change-0678/probe/xls_index_shape.py`](results/change-0678/probe/xls_index_shape.py).
It uses `olefile` 0.47 only to extract a Workbook/Book stream and walks its
BIFF frames arithmetically. Successful extraction is not a claim that
`litchi-xls` admits a fixture; the four giant or inconsistent SST declarations
are labelled and excluded from SST weight, and the source-backed corpus's
113-open distinction remains the authoritative semantic result recorded by
0648. Every analyzed fixture has a SHA-256 in the JSONL output. The command
was run at repository revision `389167b38c8b43c84c0d1799369e17360a59209c`:

```sh
python3 docs/performance/results/change-0678/probe/xls_index_shape.py \
  test-data \
  --output docs/performance/results/change-0678/probe/index-shape.jsonl \
  2> docs/performance/results/change-0678/probe/summary.json
```

The probe reports 126 physical fixtures and 372 worksheet substreams. It
counts **127,072 recognized stored-cell occurrences** at **126,152 distinct
coordinates**. The proposed occurrence index therefore charges **3,073,536
logical worksheet bytes** across the corpus if every sheet were retained at
once. This aggregate is a corpus census, not a permitted per-owner retention
amount; retaining all occurrences is required for the duplicate error
semantics described above.

| wire-level measure | p50 | p95 | largest observed |
| --- | ---: | ---: | ---: |
| worksheet occurrence slots per substream | 3 | 1,240 | 38,950 |
| worksheet index weight | 136 B | 29,824 B | 934,864 B |
| SST locator index weight per admissible group | 120 B | 5,456 B | 127,024 B |

The largest worksheet is `54016.xls` sheet 0: 38,950 occurrence slots (also
38,950 distinct positions), 37,929 BIFF frames, and a 934,864-byte logical
index charge. Its SST has 28 segments and 7,893 unique entries, charging
127,024 bytes. The two together charge 1,061,888 bytes, so a 2 MiB owner-local
cache can retain both with the 64-byte entry overhead used by this design.
Other large worksheet examples are
`45365-2.xls` at 398,416 bytes and `WithCustomViews.xls` sheet `Plan1` at
79,864 bytes. `15228.xls` has 18 worksheets and 423,816 bytes of worksheet
index weight in aggregate; the LRU budget, rather than sheet count, bounds what
survives.

The source fields explain why this is a bounded design rather than a raw
`Vec<CellValue>` retention proposal. The proposed slot charge is 24 logical
bytes for each occurrence; the current source's actual SST fields are 24 bytes
for each `SharedStringSstSegment` locator and 16 for each
`SharedStringEntryLocation` locator on the measured 64-bit target. The probe
adds only a fixed 64-byte cache-entry charge. It stores no worksheet values, no
SST strings, no whole sheet payload, and no allocation chain. Real allocator
headers, candidate capacity, cache lock/LRU time, and cache resident RSS remain
unmeasured and are explicitly outside these logical prices; the implementation
must charge candidate capacity and scratch separately as described above.

## Work removed and work retained

The cache does not alter the first complete scan. The first successful indexed
scan still pays the current worksheet framing and validation cost. On a later
hit, the work changes from `scan_worksheet` over every frame to one private
frame read plus selected-cell decoding:

* `54016.xls` scans 37,929 frames. A single-occurrence hit avoids approximately
  37,928 frame iterations for each additional selected-cell query, while still
  reading and decoding the target frame. If a coordinate has duplicates, the
  hit replays that coordinate's occurrence chain, so the saving is reduced by
  the number of target occurrences and preserves their error order.
* The existing 0605 paired measurement records a second `54016.xls` query at
  **25 additional reads, 615,822 additional bytes, and 19 observations** over
  the one-query path. Its indexed model was **2–3 reads and under 9 KiB** for
  the target record and any SST resolution. Those are prior measured/modelled
  numbers, reproduced here as pricing evidence rather than a new claim.
* For the same fixture, the selected-cell scan's 17,285,796 measured
  instructions had 70.5% attributed to the repeated worksheet scan in 0605's
  isolation. An indexed hit would parse one record rather than frame 37,929;
  instruction, cycle, allocation, p95, and p99 deltas still require a fresh
  implementation measurement.
* `WithCustomViews.xls` sheet `Plan1` costs only 1 extra read and 414 bytes in
  the 0605 second-cell row, so a 79,864-byte index should not be admitted on a
  singleton and may not amortize quickly. This is why the two-hit trigger and
  weighted admission are part of the design.

The current SST locator index is already shared by every query on a snapshot;
an additional locator cache would not remove those worksheet scans or SST
entry reads. A decoded `SstValue` cache could remove a repeated string's
source read and decode, but it would retain actual text and must be priced by
UTF-8 bytes, tested for source-freshness fences, and admitted only after a
separate repeated-string trace. It is not included in the worksheet index
saving above.

## Implementation gates for a future change

The next implementation item must clear all of these gates before claiming a
cache result:

1. Add the private weighted cache and deterministic LRU tests first, including
   checked weights, disabled/too-small budgets, pinned entries, all-pinned
   admission refusal, duplicate publish races, concurrent candidate
   reservations, hierarchical `Resource::Memory` reservation release, and
   immediate release after the last `Arc` drops.
2. Add a full-corpus differential: first query, indexed hit, an evicted hit,
   and a fresh-snapshot query must return identical values and `None` results;
   all malformed, encrypted, unsupported, cancellation, and source-change
   outcomes must retain their typed/display identity. Include a coordinate with
   an earlier malformed SST decode followed by a valid duplicate (still refuses),
   and an out-of-range SST index followed by a valid duplicate (later value wins)
   and assert that the hit still refuses. `visit_cells()` must preserve
   duplicate order and remain independent of the occurrence index.
3. Prove that no failed or partial scan publishes. In particular, mutate the
   source during the build, fail after a matching cell, fail at a malformed tail,
   and cancel during a pending formula `STRING`; each candidate must be
   discarded.
4. Measure first-query overhead under the two-hit trigger and compare it with
   the immediate-admission variant. Record allocations, allocated bytes, peak
   live bytes, RSS, source reads/bytes/observations, instructions, cycles,
   p50/p95/p99, and cache hit/miss/eviction counters. Use the existing A/A and
   ABBA protocol, cold and warm source states, owned and file/range sources,
   and a quiet pinned CPU.
5. Exercise budgets at 0, below one slot, 1 MiB, and 2 MiB. The largest
   observed worksheet plus its SST locator (1,061,888 logical bytes) must fit
   at 2 MiB, while a lower budget must evict or skip deterministically without
   changing values or errors.
6. Treat SST locator eviction as a separate gate. First prove that rebuilding
   from the retained `SstDescriptor` reproduces the exact segment and entry
   digest and refusal order; do not delete the current mandatory catalog until
   that proof and an RSS/read-cost comparison exist. Only then consider the
   optional decoded `SstValue` cache.

No code change is authorized by this design record alone. The smallest next
implementation is the private weighted cache plus the worksheet occurrence
index under the two-hit admission rule, with resident and in-flight weights
and managed execution reservations as specified above. SST catalog eviction
and decoded-value retention remain separately measured follow-ups.

## Limitations

The probe is an independent wire census, not a replacement for the Rust
source-backed differential. It does not parse formatting, formula metadata,
XF validity, all BIFF refusal order, or CFB source-version behavior. It uses
`olefile` to obtain streams and therefore cannot establish semantic admission.
The logical weights omit allocator headers, lock/LRU implementation details,
and RSS. No production code, Cargo build, timing run, or performance claim was
made by 0678.

The design intentionally does not cache a `StreamChainHint`; its lifetime and
operation ordering are tied to the borrowed `SharedOleFile`. A persistent hint
would require a separate CFB-owned cache design and must not be smuggled into
this worksheet index. The design also does not make an early-exit query: full
worksheet validation remains the admission proof required by the current
source owner and ADR 0005.

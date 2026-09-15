# Frozen design: one bounded positional read per ZIP member first-read (0611)

Frozen before implementation, at base `2d6fbeaed`. Item **ZIP-2** of
[0587](../../0587-remaining-opportunity-survey.md) (rank 14). This file is the
design record the survey required; the change record
[`0611-zip-single-read-per-member.md`](../../0611-zip-single-read-per-member.md)
carries the result and the evidence.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

## The mechanism today

A member's *first* read on a positional (`ReaderAt`) source costs two or three
requests:

| # | site | bytes |
| --- | --- | --- |
| 1 | `ZipArchive::get_entry`, `crates/soapberry-zip/src/archive.rs:2038-2042` | 30 (the fixed local header) |
| 2 | `RangeReader::read` under `ZipReader`, `crates/soapberry-zip/src/reader_at.rs:408-440` | the compressed payload |
| 3 | `DataDescriptor::parse_complete_at`, `crates/soapberry-zip/src/archive.rs:2634-2649`, reached from `ZipVerifier::read` | ≤ 24, only when the central record declares a descriptor |

The bytes between (1) and (2) — the local variable region, that is the local
name and the local extra field — are *skipped*: `get_entry` learns their
combined length from the fixed header and adds it to the payload offset without
reading them.

Change [0561](../../0561-opc-repeated-positional-reads.md) named this span model
as its target 1 and never implemented it. Changes
[0573](../../0573-zip-single-local-header-read.md) and
[0562](../../0562-zip-descriptor-read-once.md) removed duplicate reads on other
paths. [0587](../../0587-remaining-opportunity-survey.md) measured the shape: a
source-backed open of the 132-member workbook is 87 requests for 22,546 bytes,
3 for the locator and central directory plus 2 per structural member.

## What this change does

On the **ordinary indexed read path only**, issue one bounded positional read
covering the member's whole local record before the member is parsed, and serve
every read of that member from it.

The path is `IndexedReadSession::read_entry_with_accounting`
(`crates/soapberry-zip/src/office.rs:1262`), the single production body behind
`IndexedArchive::read`, `IndexedArchive::read_entry`,
`IndexedArchive::read_with_accounting`,
`IndexedArchive::read_entry_with_accounting` and `IndexedReadSession::read`, and
therefore behind `litchi-opc`'s structural member reads and its cold
`PartView::data()` part loads.

### D1. Scope: the non-proof path, and only it

The strict-layout paths — `read_to`, `read_entry_to`, `stream_to`,
`with_verified_entry_reader*` and `capture_precompressed`, all of which enter
`strict_layout_for` — are **not changed**. Change 0583's reachability audit
established that the most common OPC part read never enters the proof, and that
the proof paths run `validate_reader_entry_layout_with_name_policy` *before* any
payload byte is read.

Keeping the change on the non-proof path is what makes the interaction with the
target-scoped strict proof of [0580](../../0580-zip-target-scoped-strict-layout.md)
and its residual window of [0583](../../0583-zip-local-size-span-bound.md)
*empty*:

* 0573's speculative window is bounded by where the member's own payload must
  begin and is pinned by `the_speculative_window_never_reads_this_members_payload`.
  That window is untouched. The prover still reads only framing bytes, and a
  source that fails inside a payload still fails exactly where it failed before.
* 0580's neighbour probe still "reads exactly 30 bytes at an offset the central
  directory itself declares to be a local header, and never a variable region or
  a payload byte". The span read of this change is not a neighbour probe: it is
  issued only for the member the caller asked for, on a path where no
  neighbour's layout is proved at all.
* 0583's residual shapes (`residual-prune-distant-csize.zip`,
  `residual-zip64-sentinel-csize.zip`,
  `residual-descriptor-flag-disagreement.zip`) are verdicts of the proof. This
  change produces no verdict, so it cannot change one. Change 0582's
  differential is nevertheless run in full, because it is the gate the survey
  names for any read-grammar change.

### D2. The window

For the member with wayfinder `entry`:

```text
base   = entry.local_header_offset
desc   = if entry.has_data_descriptor { MAX_DATA_DESCRIPTOR_SIZE (24) } else { 0 }
next   = the first local-header offset strictly greater than `base` in the
         archive's offset-sorted layout, or the central-directory offset when
         there is none
end    = min( next + desc,
              base + MEMBER_SPAN_METADATA_ALLOWANCE + entry.compressed_size + desc,
              directory_offset )
len    = end - base
```

The read is taken only when `30 <= len <= MAX_MEMBER_SPAN_READ_BYTES`. Both
terms of the `min` matter:

* **`next + desc`** is the physical bound. In a gapless archive — 7,212 of the
  7,757 OOXML members in `test-data`, the remaining 545 being descriptor-bearing
  or padded — `next - base` is exactly `30 + variable + compressed + descriptor`,
  so the window is exactly the member's own local record and the read costs no
  byte the three reads did not already cost, except the variable region. The
  `+ desc` term exists because `DataDescriptor::parse_complete_at` asks for up to
  24 bytes and a real descriptor is 12, 16, 20 or 24; without it a 16-byte
  descriptor would leave 8 bytes of the descriptor request uncovered and cost a
  request back.
* **`MEMBER_SPAN_METADATA_ALLOWANCE + compressed + desc`** is the declared
  bound, and it is what keeps a padded archive from turning into a large read.
  `MEMBER_SPAN_METADATA_ALLOWANCE` is `STRICT_LOCAL_HEADER_WINDOW` (640 = 30
  fixed + 610 of variable region), the constant change 0573 measured over the
  same corpus: 640 bytes covers 4,270 of 4,270 OOXML members' variable regions,
  512 misses 301 of them because Microsoft Office writes a `0xa220` growth hint
  of 260 or 516 bytes.
* **`directory_offset`** keeps 0573's property that a speculative window never
  reaches past the central directory.

Never reading past `min(next, directory_offset) + desc` is also what keeps this
change from being the coalesced *header run* that
[0575](../../0575-zip-lazy-strict-layout-design.md) priced and set aside as
"fallback only": a run read at a 64 KiB gap cap cost 644,889 bytes to deliver a
1,251-byte member. This window is one member wide and is bounded by that
member's own declared size.

### D3. The ceiling

```rust
/// The largest member local record this path will read in one request.
const MAX_MEMBER_SPAN_READ_BYTES: u64 = 64 * 1024;
```

Measured, not guessed. Across 7,757 members of 335 OOXML containers under
`test-data`, `30 + local variable + compressed + 24` is at most 4 KiB for 7,493
members (96.6%), 32 KiB for 7,695 (99.2%) and 64 KiB for 7,732 (99.68%); the
largest is under 1 MiB. 64 KiB is also the value `litchi-opc` already names as
the largest single speculative window it will issue
(`MAX_SOURCE_READ_AHEAD_BYTES`, `crates/litchi-opc/src/source_backed/read_ahead.rs:20`),
which [0577](../../0577-ooxml-open-relationship-parts.md) proposed borrowing for
its run coalescing.

The buffer is fallibly reserved with `try_reserve_exact` and reports
`ErrorKind::Allocation` on failure like every other bounded buffer in the crate;
it lives only for the member read. The resident bound this change adds is one
buffer of at most 64 KiB per member read in flight.

A member whose window would exceed the ceiling takes **no span read at all** and
runs today's path unchanged, rather than a truncated one, so it can never cost
more requests than today.

### D4. When the local and central lengths disagree

The window is computed from central facts and one neighbour's declared
local-header offset. **It is never a verdict and nothing is trusted because it
is in the buffer.** The fixed local header is parsed out of the buffer by
exactly the code that parses it today; the local variable length comes from that
parse, not from the central record; `local_end > directory_offset`, the ZIP64
size framing, `body_end_offset > directory_offset`, the size conversions, the
CRC and the declared-size check all run unchanged, in the same order, on the
same values.

If the member's local variable region is longer than the window assumed — the
only way the central facts can mislead the window — the payload simply starts
past the end of the buffer. The reader then delegates those reads to the source
verbatim, at the same offsets and lengths as today, and the member read costs
one more request than today rather than one fewer. It cannot change what the
member read returns. (The amendment below makes that last sentence true by
construction: a payload the buffer does not wholly hold is read through the
passthrough reader, not served in pieces.)

The same holds for a duplicate or out-of-order local-header offset, which makes
`len < 30`: no span read is taken and today's grammar runs.

### D5. Error precedence within one member

Three rules, in the order they bind:

1. **A failed span read is the member read's failure, tried once.** The span
   read is the member read's first contact with the source and begins at the
   same offset as today's first read, so a source failure is reported rather
   than retried behind the caller's back. This is deliberately *not* 0577's
   invariant 5 ("a run read that fails falls back to per-member reads"): that
   invariant is right for a read covering several members, and wrong here,
   because a retry would give one member read two refusals, two cancellation
   observations and two resource reservations. `litchi-opc`'s
   `managed_input_budget_refusal_counts_only_the_terminal_read_reservation`
   pins exactly that, and it is the contract this rule keeps. The residual is
   that the first read of a member is longer than it was: a source that fails
   as a function of request length could therefore fail where it did not. The
   `ReaderAt` contract answers a short available range with a short read, not a
   failure, so a conforming source cannot; a source that refuses a long read
   refuses today's payload read too.
2. **A short span read is not an error.** The read takes what the source
   returns; whatever it did not cover is read exactly as before. So a truncated
   source still fails at the byte, and with the message, it fails at today —
   `read_exact_at`'s `"failed to fill whole buffer"` for the fixed header, and
   the payload's own `Eof` beyond it.
3. **0317's brackets are unchanged.** `litchi-opc` fences a member read with a
   source-version observation and an execution-context check on both sides
   (`changes/0317-opc-source-read-error-precedence.md`), and this change is
   entirely inside those brackets. Source-version failure still precedes
   execution-context failure, which still precedes the mapped ZIP member error.
   Under a monitored-read scope ([0600](../../0600-opc-cold-read-observations-and-name-lookup.md))
   each removed positional read also removes one
   `ensure_current_io_if_monitored` observation; the paths that take a monitored
   scope — streams, verified reads and publications — are the strict-proof paths
   D1 excludes, so their per-read observation density is unchanged.

### D6. The one place a refusal can move — stated, bounded and measured

Under a caller-supplied `ExecutionContext` with a finite `Resource::InputBytes`
limit, `litchi-opc` commits the bytes actually read against that budget
(`crates/litchi-opc/src/source_backed.rs:11509`). A spanned first read commits
the member's local variable region as well — the bytes today's two reads skip.

* The increase is bounded by `MEMBER_SPAN_METADATA_ALLOWANCE - 30 = 610` bytes
  per spanned member first-read, and is the local variable region's real size,
  which over the `test-data` OOXML corpus averages about 95 bytes per member
  (6,250 of 7,757 members have a variable region under 64 bytes).
* No error identity changes and no new refusal exists: the refusal is the same
  `ExecutionError::ResourceLimit`, mapped through the same path. What moves is
  only how soon a finite input budget can be exhausted.
* This is the same class and the same direction as the speculative window change
  0573 already landed on the strict path, which paid +0.54% bytes for −47.3%
  reads and is subject to the same budget.
* Ordinary opens pass no `ExecutionContext`, so the budget applies to managed
  packages only.

Nothing else in the change is byte-visible: `ZipOperationAccounting`'s payload
counters still see exactly the payload, because the payload still flows through
the same `CountingReader<ZipReader<..>>`.

### D6a. Test sources that watch for "the payload read"

Sixteen tests in four files — ten in `litchi-opc`'s `#[cfg(test)]` module, two
in its `tests/source_backed_batch.rs`, three in
`litchi-xlsx/tests/source_backed_cell_values.rs` and one in
`litchi-xlsx/tests/source_backed_row_visibility.rs` — drive a source that fires
when a read *begins exactly at the member payload's offset*, and assert a
source-change, cancellation, flight, reservation or read-count contract around
it. A member's payload no longer has a read of its own, so fifteen fail and one
deadlocks.

The contracts they assert are untouched; only the trigger encodes the old
grammar. Each source now fires when a read **delivers the payload's first byte
and began inside the member's own local record**:

```rust
requested > 0
    && offset <= payload
    && payload - offset < requested
    && payload - offset <= MEMBER_RECORD_PREFIX_BYTES   // 640, change 0573's window
```

Each clause was earned by a failure, in this order:

1. "Begins at the payload **or** at the member's local header" fires twice per
   member in a gapless archive: the zero-length read a range reader issues at
   the end of member *n* lands on member *n+1*'s header. `requested > 0`
   excludes it.
2. The local header alone fires twice during publication, because the
   preservation path issues a short local-record probe at that same offset.
   Requiring the payload byte to be *delivered* separates the member read from
   the probe.
3. "Delivers the payload byte" alone is tripped by a bulk publication copy that
   spans the member from 64 KiB away, which turned an exact-no-op publication
   into a refusal. The 640-byte prefix bound keeps a fetch *of* the member
   distinct from a copy *over* it.

The result fires exactly once per member first-read under either grammar, and
every assertion in all sixteen tests is unchanged.

### D7. Shape of the implementation

* `SpannedSource<'a, R>` (`crates/soapberry-zip/src/office.rs`) — a `ReaderAt`
  that answers from one member's buffered span and delegates every other read to
  `&'a R`. It is cloneable so the payload reader and the verifier can both hold
  it, and it carries `Option<Arc<MemberSpan>>` so the fallback costs nothing.
* `ZipArchive::get_entry_from`, `ZipEntry::reader_over`,
  `ZipEntry::verifying_reader_over` (`crates/soapberry-zip/src/archive.rs`) —
  `pub(crate)` siblings of `get_entry`, `reader`, and `verifying_reader` that
  take the source to read from. `get_entry` keeps its public signature and
  delegates.
* `IndexedReadSession`'s decoder becomes
  `DeflateDecoder<CountingReader<ZipReader<SpannedSource<'a, R>>>>`, so change
  0594's one-decoder-per-open reuse is preserved exactly.
* No public API changes, no new `unsafe`, no new dependency, and no change to
  `litchi-opc`'s production code — every `litchi-opc` edit is inside a
  `#[cfg(test)]` module or under `tests/`.

## Amendment, after the differential

The frozen text above said that a window that does not cover the payload leaves
those reads reaching the source "verbatim, at the same offsets and lengths as
today", and that it "cannot change what the member read returns". Change 0582's
differential, extended to `IndexedArchive::read_entry` and re-run over
2,886,786 member verdicts, falsified the second claim on one crafted shape, and
is the reason the design changed before landing.

**What it found.** `mutations/crafted-many-tiny-one-reaching/cdh-crc-e0-ffffffff.zip`,
member `m000.bin`, under both limit profiles: the member is refused on both
legs, but with a different typed error —
`InvalidSize { expected: 4, actual: 5 }` before,
`InvalidChecksum { expected: 4294967295, actual: 211534962 }` after. Two
verdicts out of 2,886,786; no class A, B, C or D divergence anywhere, no panic,
no oracle failure.

**Why.** That member's central record declares `compressed_size = 4096` while
its local header declares 4, and the next local header is 42 bytes in, so the
window is 42 bytes and stops inside a payload the reader will ask 4,096 bytes
of. A partly-covered read was served as a *short* read out of the buffer, and
`ZipVerifier::read` completes a member the first time
`size >= uncompressed_size_hint()`, so re-chunking the decoder's input moved
which of two refusals fired first: with a 5-byte piece the size check failed;
with a 4-byte piece the size check passed and the CRC check failed.

**The amendment.** The buffer now serves the payload only when it holds *all* of
it, and it never answers a zero-length read at all — a read that fetches nothing
has nothing for the buffer to answer with, and letting it reach the source keeps
a versioned or cancellable adapter's chance to refuse it exactly where it is
today. `SpannedSource::covers` is checked against the member's own
`compressed_data_range`, computed from the local header after it is parsed, and
a payload the buffer does not wholly hold is read through the passthrough
reader, so the decoder's input arrives in exactly the pieces the source would
have delivered. The fixed local header is still answered from the buffer either
way, so such a member costs the requests it costs today and never more.

This replaces an argument with a construction: the only reads the buffer answers
are reads it can answer in full, and a member whose local and central framing
disagree runs the historical grammar from its payload onward. The re-run of the
differential after the amendment is the evidence retained in this packet.

## Oracle and gates

* **Bytes and errors.** Identical bytes and identical errors for every member of
  every OOXML fixture, before and after, through `read_entry`.
* **Order independence.** New tests read the members of one archive forwards,
  backwards and interleaved and require identical results, and require a member
  read through a fresh session to equal one through a shared session.
* **Change 0582's differential**, in full, both builds, every member, both
  directions, both limit profiles: zero verdict changes.
* **Requests and bytes.** The 0587 probe, per open and per part.
* **Timing.** The simulated per-request-latency range source of
  [0572](../../0572-ooxml-range-source-attribution.md); plus local-file
  selectors, where no change is expected and none may be claimed.

## Falsification

The change is rejected if any verdict in 0582's differential moves; if any
fixture's bytes or error identity differ; if the request counts do not fall as
modelled (open 87 → 45 on the 132-member workbook, 45 → 24 on `shapes.pptx`,
12 → 6 on `comment.docx`, a part 2-3 → 1); if the request drop does not appear
as wall clock on the delayed transport; or if a local-file selector regresses by
more than 5%.

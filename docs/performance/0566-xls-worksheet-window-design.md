# 0566: design for a windowed source-backed XLS worksheet scan

Status: design only. No production change and `performance_claim: none`. This
record freezes a design and its admission gates; it measures nothing and claims
nothing.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## Why this record exists

Change [0564](0564-xls-open-read-attribution.md) measured that a source-backed
one-cell query costs 266 positional reads beyond the open: 177 four-byte
worksheet record headers, 88 payload reads and one shared-string fetch, so "a
skipped record costs exactly one read; a consumed record costs exactly two".
Change [0565](0565-xls-globals-single-pass.md) removed the equivalent pattern
from the globals scan. This record designs the same treatment for
`WorksheetScan` and states what would have to be measured to admit it.

It also answers a question change
[0325](changes/0325-cfb-frame-transaction-rejected.md) left open. 0325 closed
this line for want of "an explicit contract for distinguishing bytes that are
payload from bytes that may be skipped". That contract is supplied below.

## What the scan does today

`WorksheetScan` owns a `SharedOleStreamCursor` created once at the sheet start,
plus `upper_bound`, `scanned_bytes`, `scanned_records`, `limits`, `execution`
and a `scratch` payload buffer. It has three primitives:

- `next_frame` checks execution, bounds-checks `position + 4 <= upper_bound`,
  issues **one four-byte `read_exact`**, then applies the oversize, boundary,
  record-count and byte-count checks.
- `read_payload` and `take_payload` check execution, resize `scratch` and issue
  **one `read_exact` of the payload**.
- `skip_payload` checks execution and calls `cursor.skip_forward`, which reads
  nothing and takes no source observation.

### The decisive fact: the sheet boundary is already exact

`validate_sheet_offsets` runs at open, before any query. It sorts the sheets by
start offset, rejects duplicates, and gives each sheet an `end` equal to the
next sheet's start or the stream length, having already validated that every
start is at or after the globals end and within the stream.

So unlike change 0565 — where the globals end is unknown until framing completes
and the final fill must spill and be truncated — **the worksheet window can be
clamped exactly from the first fill**, and no byte outside the selected sheet's
own validated region is ever read. This design needs no over-read clause for the
sheet boundary, and `selected_queries_are_bounded_to_the_selected_owner` survives
unchanged.

The residual is intra-region. A sheet's `end` is the next sheet's start, not its
own EOF record, so a final fill can read trailing slack that belongs to the same
sheet. On `ConditionalFormattingSamples.xls` that slack is zero; on
`poi/Simple.xls` the last sheet is `[2024, 4096)` with its EOF at 2287, so up to
1,809 bytes of its own region could be read that today are not.

### Limits

`max_worksheet_scan_bytes` and `max_worksheet_scan_records` are both checked in
`next_frame` *after* the header has been read, against cumulative framed totals.
Today's real byte fence is therefore `sheet.start + limit + 4`: the header of the
record that trips the limit is always read. `max_worksheet_scan_records` bounds
parse work only and does not bound I/O.

## The read pattern today

Figures below are derived from the fixture's own CFB and BIFF geometry and
reproduce 0564's measured 266 exactly, so the model is validated against the
measurement rather than asserted. The selected perf cell is worksheet 1, row 20,
column 4, whose substream is `[599787, 646630)`.

| Case | headers | payloads | skips | shared string | reads | bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| one-cell, selected sheet (177 records) | 177 | 88 | 88 | 1 | **266** | 1,713 |
| full 16-sheet text | 2,880 | 1,739 | — | 658 | **5,277** | ~41,700 |

Open and list perform zero worksheet-scan reads; this design moves neither.

The shape is 87.5% four-byte reads by count and 2.2% of the sheet span by
volume: the one-cell scan touches 1,706 of the selected sheet's 46,843 bytes.

## The design

Keep the cursor and replace `scratch` with a retained window buffer. The cursor
matters: worksheet substreams start deep in the stream, and
`SharedOleStreamCursor` does not re-walk the allocation chain from the first
sector on each call, whereas the `read_stream_range` fill that change 0565 uses
does. One `read_exact` still batches physically contiguous runs into one
positional read each and takes exactly one source observation.

A single primitive ensures stream bytes `[position, need_end)` are resident. If
they already are it reads nothing. Otherwise it compacts the unframed tail to
the front of the buffer and issues **one** read whose length is the demand,
raised to the current window target and clamped by the sheet boundary and the
byte fence. The window starts at 512 bytes, one CFB sector, and doubles to a
64 KiB cap. `skip_payload` advances past a payload and, when that payload ends
beyond the filled end, seeks with `skip_forward` and discards the buffer,
reading nothing.

The traced fill schedule for the selected sheet is seven fills totalling 37,907
bytes; the gaps between them are skip-seeks that never read 8,936 bytes of
skipped payload.

### Invariants

1. Every byte read lies within `[sheet.start, min(sheet.end, sheet.start + max_worksheet_scan_bytes))`.
2. Reads are disjoint and strictly increasing in offset, so each byte is read **at most once**. Not "exactly once" as in change 0565, because skip-seeks mean some bytes are never read at all.
3. The read count is bounded by the doubling schedule over the scanned span, plus one read per skip-seek.
4. The boundary clamp is exact from the first fill: no byte of another sheet, of the globals, or past the stream is read.

### The byte limit becomes a true read fence

Every fill is clamped by `sheet.start + max_worksheet_scan_bytes`. The
replacement contract is **no byte at or past that offset is read**, which is
strictly tighter than today's `limit + 4`. The existing limit tests keep their
typed errors and their `observed` values. One accepted change: when fewer than
four bytes of budget remain, the header cannot be read, so `observed` becomes
`scanned_bytes + 4` rather than `scanned_bytes + frame_len`. That is unreachable
for any limit of four or more at a record boundary and no existing test
exercises it.

Folding `max_worksheet_scan_records` into the fill bound was considered and
**rejected**: it would degrade a legitimate small records limit into one fill per
record, and the records limit is a parse-work bound rather than an I/O bound.

### Replacement wording for the two payload contracts

Both current contracts are unachievable under any coalescing scheme, because a
record's kind is knowable only from its header and the header sits inside the
window. This is precisely the gap change 0325 named. The replacements:

> **Skipped payloads.** A skipped payload is never framed, interpreted or
> published. The scan issues no read whose purpose is a skipped payload: no read
> begins at an offset inside a skipped payload, and when a skipped payload
> extends past the filled end the scan advances past the remainder without
> reading it. Bytes of a skipped payload that lie inside a window already filled
> to frame earlier records are read; they are bounded by the window ceiling and
> confined to the selected sheet's own validated region.

> **Oversized records.** When a record header declares more than
> `litchi_biff::MAX_RECORD_BYTES`, the scan refuses on the header alone. It
> issues no read for that payload and never grows a window to cover it. Payload
> bytes already resident from an earlier fill are never framed, interpreted or
> published.

Worked example, the 4,096-byte unknown payload the existing skip test inserts:
the first fill covers the header and 201 of the payload bytes, the skip then
seeks past the remainder, and a four-byte fill reads the EOF header. Two reads
and 516 bytes, against 30 reads and 142 bytes today. No read begins inside the
payload.

### A density gate is required

Windowing trades bytes for syscalls, and the trade inverts on sheets of large
skipped records. At the roughly 116 ns per syscall that change 0564 derived, one
64 KiB fill pays for itself only while the mean frame is below about a kilobyte.
A simulated sheet of 200 near-maximum opaque records goes from 203 reads and 824
bytes to 31 reads and **1,613,496 bytes**.

The gate uses counters the scan already maintains: before each fill, if the mean
framed bytes per record exceeds 1 KiB, the fill is exact and the window target
resets. It is re-evaluated per fill, so the running mean is its own hysteresis.

| Case | today | ungated | gated |
| --- | ---: | ---: | ---: |
| selected sheet, one cell | 265 reads / 1,706 B | 7 / 37,907 B | **7 / 37,907 B** |
| 16-sheet text | 4,619 reads / 37,056 B | 110 / 650,137 B | **110 / 650,137 B** |
| synthetic 200 opaque records | 203 reads / 824 B | 31 / 1,613,496 B | **201 / 1,312 B** |

The gate is identical to the ungated design on every real sheet modelled and
caps the adversarial case at today's read count plus one sector.

### Cancellation, freshness and error precedence

Every existing execution check stays, including where bytes come from the buffer
and no I/O occurs, so the **CPU** interval between checks is unchanged at one
record. Only the **I/O** interval grows from one record to one window. The core
execution contract promises only that a check returns cancellation when
requested and prescribes no interval, and there is precedent at the same size on
this owner: `materialize_eager`'s documented contract is "checked between bounded
source reads" with a 64 KiB bound.

Source observations fall from one per read to one per fill: 265 to 7 on the
selected sheet. The query-level bracket is unchanged, observing before the scan
and after it.

The accepted precedence change mirrors change 0565: within one fill, an I/O or
source-version failure at a later offset is reported before a framing, limit or
parse error of an earlier record covered by the same fill. Two cases do not
change, because the fill is clamped: a fill never fails for exceeding the stream
or the sheet boundary, and the byte-limit error still precedes any read at or
past the fence.

### Memory

Peak retained per scan rises from one payload of at most 8,224 bytes to at most
64 KiB plus one taken payload, allocated fallibly and capped by the smaller of
the window cap, the sheet span and the byte limit, so a 311-byte sheet allocates
311 bytes. Scans run sequentially, so sixteen sheets do not hold sixteen buffers.

## Predicted effect, and the cost

| Selector | at HEAD | with 0565 | with 0565 and this design |
| --- | ---: | ---: | ---: |
| open | 655 | ~16 | unchanged |
| open and list | 655 | ~16 | unchanged |
| open and one cell | 921 | ~282 | **~24** |

Isolating this design's own delta:

| | reads | bytes | observations |
| --- | ---: | ---: | ---: |
| one-cell scan today | 266 | 1,713 | 266 |
| one-cell scan windowed | **8** | **37,914** | 8 |
| 16-sheet text today | 5,277 | ~41,700 | 5,277 |
| 16-sheet text windowed | **778** | **~654,800** | 778 |

Read reduction of 97.0% and 85.3%; **byte increase of 22.1 and 17.5 times**.
That trade is the central risk and is stated up front rather than buried.

Two consequences worth recording. This design removes 258 of the 266 reads that
change 0564 attributed to the scan. And after it, the per-cell shared-string
fetch becomes the dominant read source for the text path at 658 of 778 reads,
which is the next attribution target rather than part of this change.

## Admission gates

- **Counted reads.** Tests demonstrating at-most-once reads inside the exact sheet region, the byte-limit read fence, no read issued for a skipped or oversized payload, and the density fallback, each shown to fail without the change.
- **Syscalls.** Per-selector positional-read and stat counts isolated by differencing `strace -f -c` at 1 and 11 samples with one warmup, as change 0564 did. Open and list must be **unchanged**; that is the control proving the delta is the scan.
- **Bytes.** Logical bytes read per selector before and after, with the byte increase stated explicitly rather than omitted.
- **Opaque-heavy corpus.** Re-run on the comments-opaque-heavy corpus: reads must fall or hold and bytes must not exceed today plus one window per sheet. Without this the density gate is unvalidated.
- **Latency.** Paired A1/B1/B2/A2 p50 over the ten XLS selectors, warmup 5 and 60 samples per child, CPU pinned and ASLR disabled, with every cell adverse in both directions by more than 5% reported explicitly.
- **Correctness.** The full `litchi-xls` suite, Clippy and rustdoc with warnings denied, a no-default-features check, formatting, and the wider OLE2 and facade suites before commit.

## Does this supply what change 0325 required?

In form, yes, by instantiating change 0565's template. Item by item:

| 0325 requirement | Status |
| --- | --- |
| A contract separating payload bytes from skippable bytes | **Supplied.** No read is issued for a skipped or oversized payload; incidentally resident bytes are never framed, interpreted or published. |
| A bound on the over-read | **Supplied, and stronger than change 0565's.** The sheet boundary is exact before the scan begins, so there is no spill-and-truncate at all. |
| Limits that apply before extra bytes are accepted | **Supplied.** The byte limit becomes a true read fence, strictly tighter than today's. The records limit is explicitly identified as a parse-work bound that does not fold into the fill. |
| Freshness granularity | **Changed, not preserved**: 265 observations become 7, with the query-level bracket intact. |
| Error precedence unchanged | **Changed, not preserved**, the same accepted change as change 0565's. |
| Cancellation boundary | **Changed in the I/O interval only**, which the core contract permits and for which this owner has precedent at the same size. |

Change 0325's other claim, that the open and list selectors have no worksheet-scan
opportunity, is literally correct and remains correct: this design moves neither.
Change 0564 corrected the inference drawn from that claim; this record addresses
the 266 reads 0325 actually had in scope. Its stated ceiling for them was
"bounded by the frames one query touches"; this design reaches 8 of 266, at the
cost of two replaced contracts and 22 times the bytes.

## Limitations

Nothing here is measured. Every read and byte count is a model derived from the
fixture geometry, validated against change 0564's measured 266 for the one-cell
case and otherwise unvalidated. No latency, allocation, cold-cache,
remote-source, concurrency or cross-platform result is claimed. The design is
sequenced after change 0565 and depends on the test helpers that change adds.

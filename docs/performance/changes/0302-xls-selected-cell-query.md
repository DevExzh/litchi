# Change 0302: Reuse XLS selected-cell query scratch storage

Status: implemented

## Scope

Source-backed XLS selected-cell queries still scan the selected worksheet
sequentially through its valid `EOF`. This change reuses one fallibly allocated
worksheet payload buffer instead of allocating a fresh `Vec<u8>` for every
decoded frame. Formula `STRING` continuations retain owned payload chunks while
they are decoded because the decoder borrows all chunks at once.

Packed `MULRK` and `MULBLANK` records now expose validated internal visitors for
this query path. The visitors preserve the existing payload shape, count,
column-range, entry decoding, and XF validation checks without first building a
temporary `Vec<CellRecord>`. The existing allocating `parse_mul_*` APIs remain
unchanged for other callers.

Duplicate cells remain last-wins, and malformed or truncated records after an
earlier match remain errors. No row-order assumption, `INDEX`/`DBCELL` shortcut,
early return, or semantic relaxation was introduced.

## Performance claim

`performance_claim: none`

The change removes or reuses transient payload and packed-result-vector
allocations in the selected query path. It makes no quantitative claim about
I/O, latency, RSS, general allocation volume, or end-to-end throughput; the
full worksheet `EOF` scan remains.

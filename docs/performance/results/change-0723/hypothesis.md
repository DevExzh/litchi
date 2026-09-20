# 0723 pre-candidate hypothesis

Baseline: `45cb480eaa`; non-iWork performance program remains active.
The fresh baseline profile runs the unchanged 0684 selected-query probe for
2,000,000 owned-source 54016 first-cell queries. It reports 794,593,284 ns;
811 samples, zero lost samples, with 12.58% self samples in
`cursor_chain_sector`. This is attribution evidence, not a paired latency gain.

The existing worksheet occurrence index retains a worksheet-start checkpoint.
A later frame still traverses the intervening FAT links on each query. Retain
one additional immutable exact-frame checkpoint for the first occurrence of
the coordinate used to build that index. Obtain it only after the full scan
succeeds, with one metadata-only hinted seek. This avoids read-ahead cursor
positions and requires no CFB API change. Preserve the original worksheet
checkpoint for queries before the saved frame and the independent SST point.

The new optional `(offset, checkpoint)` is expected to add 40 fixed bytes per
admitted index, including numeric/missing indexes. Charge it before collection
through the existing weighted reservation mechanism. Price the one-time slot
search, metadata traversal, path scratch, layout change and capacity boundaries
explicitly. Keep ordinary scanning, full validation, errors, source/execution
fences, slot ordering and fresh selected-frame/SST decoding unchanged.
No decoded values, source bytes, errors, or per-slot checkpoints are cached.

Use unchanged 0684/0686 native, repeat, allocation, counted-I/O and corpus
probes. Freeze the detailed plan and analyzers before main captures. Native
and allocator/trace runs remain separate. Reject on correctness, freshness,
unpriced resource growth, or frozen primary latency gates. Disclose every
mean/tail/control drift flag; no resampling may erase original observations.
Any follow-up requires its own prospective plan and independent disposition.
No broad CRUD, cold-device, remote, concurrency, native Office, or cross-platform
performance claim follows from these warm selected-query measurements.

# Release pre-compaction worksheet storage

Control `f0ab67b55` retains the original rewritten `Vec<u8>` while a distinct
compacted vector is parsed into the verification Store. The candidate inserts
`drop(after)` immediately after successful `changed_worksheet` and before
optional grid parsing. `WorksheetOutput` owns its compacted bytes and scalar
web proof; neither it nor the owned Store borrows the obsolete vector.

An independent source review confirms that the phase order remains compaction,
grid, web, styles and change verification. Failure cleanup, exact output bytes,
source Arc retention, atomic publication, inverse patches and the 4,096-cell /
1 MiB Store handoff are unchanged. No public API, cache, execution context,
unsafe code or I/O changes. The accepted ADR inventory is unchanged from the
previously read 30 files; hashes are retained in `adr-refresh.json`.

| ADR | Obligation |
| --- | --- |
| 0001, 0003, 0006 | Preserve typed errors, byte identity, validation order, no-op behavior and atomic publication. |
| 0002, 0010, 0011, 0024 | Retain format-owned worksheet grammar and dependency boundaries. |
| 0005 | Reduce overlapping byte ownership only when measurement supports practical benefit; no assumption that deallocation returns pages to the OS. |
| 0008 | Full XLSX suite and scoped formatting, workspace checking, Clippy, rustdoc and boundary gates. |

The old rewrite reserves source length plus 128 bytes per effect. Releasing it
can remove worksheet-sized overlap with the verification Store, but an earlier
snapshot scan may still dominate peak heap. Whole-process high water and
allocator page retention can hide local lifetime effects. This experiment must
not call source inspection a measured memory improvement. It also does not
attribute the variable RSS in 0470 to this vector, which existed in both roles.

The frozen protocol uses 100-sample diagnostic ABBA for six ordinary XLSX rows
plus the previously flagged payload-heavy PPT writer. A 201-row short guard and
whole-process Heaptrack accompany it. Existing correctness tests cover the
unchanged behavior; no test is added merely to mirror an explicit drop.

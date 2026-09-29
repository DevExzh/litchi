# 0832 source and scope review

The candidate archive contains the four intended source files. I reviewed the
formatted `candidate/after` tree against `candidate/before` and the design;
no source correctness blocker was found.

`Assignments<T>` now keeps an empty map as an empty `Vec` plus an optional
complete `Assigned<T>` record (`column.rs:412-428`). The first assignment is
inline. `get` and `into_ranges` have corresponding inline paths, with the
latter reserving its one final range directly (`:478-517`). The second
assignment allocates the original checked `2 * COLUMNS` tree in a local value,
replays the first record and then the second, and only then installs the tree
and clears the inline record (`:436-475`). Thus a `try_reserve_exact` failure
leaves the prior inline record untouched; no partially initialized map is
published. With `COLUMNS = 16_384`, the existing tree indices remain within
the allocated vector, and later assignments retain the bounded recursive
algorithm. Work is constant for zero/one records, one-time `O(COLUMNS)` for
promotion, and `O(log COLUMNS)` per later assignment; no operation scans a
record's covered width.

Replay order preserves last matching-record semantics. The value is copied as
one complete `T`, so omitted properties are replaced rather than inherited,
and the existing ordered traversal still produces disjoint ranges with
adjacent equal values compacted. The parser now constructs without reserving
the tree and propagates assignment errors at `codec.rs:1008-1079`. The edit
validator and snapshot column writer make the same constructor/assignment
change at `validation.rs:68-75` and `write/columns.rs:27-30`. The validator
builds its owner map before the style-width check, and the writer builds it
before touching output, so promotion failure cannot leave a partial edit
output. The protected-sheet check remains before the empty-action return.
Error timing intentionally moves from construction to the second valid record
when promotion is needed, while attribute/range validation and the existing
semantic refusals remain in their surrounding order.

The new unit coverage checks empty and single-record storage, inline lookup
and ranges, overlapping promotion with equal-range compaction, repeated full
width replacement, and a 1,000-assignment grid oracle (`column.rs:704-933`).
The existing parser test `raw/worksheet/tests.rs:435-485` exercises two
overlapping records with a complete initial property set and a later record
that omits style, outline, and flags, asserting that the omitted properties
are reset rather than inherited. Existing edit tests at
`raw/worksheet/edit/tests.rs:824-948` and `:951-1034` cover overlapping owner
splits, preservation of unedited attributes, all column facets, style
retargeting, resets, and reparsing. Static caller inspection found no stale
`Assignments::new` or infallible `assign` call in the four archived source
files.

The focused limitation is that no existing test injects failure into the
`dense_nodes` reservation. The atomicity argument is explicit in the source:
the fallible allocation is local, both replay assignments are completed before
the map is installed, and a reservation error therefore leaves the inline
record readable and resolvable. This does not justify adding a production
fault-injection hook solely for this path. The existing grid, parser, writer,
transaction, build, allocation, and full-output gates should remain required.

The census and design scope are consistent: 43 of 202 worksheets with a
direct `<cols>` element have one physical record and avoid promotion, while
159 promote. The source change therefore supports the measured one-record
case without implying a broad multi-record latency claim.

# 0767 log sections (ready to paste)

## For HOTSPOTS.md

## 0767 — CFB open and read no longer clear a table-sized map per stream

[0767](0767-cfb-reparse-linear.md) removes the two super-linear terms of
CFB parsing:

- 0749's A5 term: `validate_stream_allocations` cleared a visited map the
  size of the FAT or MiniFAT for every stream, on every `OleFile::open`;
- a second one it found: `open_stream` allocated and zeroed such a map for
  every read.

The reusable maps now stay all-clear between walks. A walk is restored by
zeroing the words of its chain's sectors when the chain is under an eighth
of the table's words, and in bulk otherwise. The reader keeps one chain
scratch.

On Callgrind the base's open work per stream grew ×4.0 from 1,000 to 10,000
mini streams (A5 ×6.57, then ×9.67); the candidate's is flat (A5 ×3.00 and
×3.33). An open of 10,000 mini streams falls 31–32% in time, and a read of
every stream 31–39%.

What remains on these paths:

- the path lookup's O(log n) sibling-tree search per read (1,441 → 1,799 Ir
  per stream from 1,000 to 10,000 streams);
- memory-bound copying at 10,000 streams;
- a pre-existing gap: A5 admits a mini stream in a partial last mini sector
  that `open_stream` then refuses.

The non-iWork goal remains active.

## For REPORT.md

## 0767 — linear CFB structural parsing and reads retained

[0767](0767-cfb-reparse-linear.md) makes `OleFile::open`'s
stream-allocation validation and `open_stream`'s chain collection cost time
proportional to the chains instead of stream count × table size. The
verdicts, checks and error strings are unchanged. Two builds give
byte-identical verdicts on 30,400 byte-level faults over 104 files, and every
probe output matched. 12 paired layouts on core 28, 936 processes:

| Lane | Δ p50 |
|---|---:|
| `OleFile::open`, 10,000 mini streams (v3, v4) | −31.3%, −32.5% |
| `OleFile::open`, 10,000 regular streams (v3, v4) | −20.2%, −3.4% |
| Read every stream, 10,000 mini streams (v3, v4) | −39.2%, −31.2% |
| `SharedOleFile` open, 10,000 mini streams | −34.9% |
| Common-editor capture, 1,000 / 3,000 mini streams | −16.8% / −19.9% |
| Reuse `write_to` on 45543.ppt | +0.5% (instructions −1.0%) |
| `OleFile::open` of three real fixtures | −0.4% to +1.4% (instructions +0.10% to +0.16%) |

Native instructions per stream of an open are now flat: 7,296, 7,285 and
7,285 at 1,000, 3,000 and 10,000 v3 mini streams, against the base's 7,181,
7,387 and 8,157. A read of every stream allocates one buffer fewer per
stream (444.7 → 42.4 MB at 10,000 mini streams).

The first version's per-bit restoration cost small tables +0.2–0.3%
instructions per open; the retained version restores by word and keeps
small tables on the bulk clear.

Flags: a harness `cfb_list_streams/tiny` control, +20 ns on an unchanged
130 ns call with equal instructions (code placement); allocator-state tails
on DOC semantic open and the 10,000-stream capture.

Gates pass for `litchi-cfb`, its OLE2 dependents and the facade.
`performance_claim: none`. [Evidence](results/change-0767/README.md).

## For GOAL_AUDIT.md

## 0767 — CFB validation preserved, its per-stream term made linear

[0767](0767-cfb-reparse-linear.md) keeps every CFB validation, ownership,
cycle, overlap, FAT, MiniFAT, directory and truncation check the goal
requires. It changes only how the reusable visited maps are cleared: they
stay all-clear between chain walks, and a walk is restored by clearing only
what it set, or the table's words.

The walkers and their error strings are the same code (ADR 0006, ADR 0026).
Test-only oracles hold both collectors and A5 to their fresh-map forms on
random call sequences and 1,800 fault-injected files. A cross-build lane
shows byte-identical verdicts from the base and candidate builds on 30,400
byte-level faults.

The CPU cost an input with thousands of streams could force on every open
and every read is removed, and a work test bounds it: at most the tables
plus eight map words per chain sector.

Remaining debt:

- the path lookup's logarithmic search per read, linear in siblings for a
  list-shaped tree the reader accepts;
- a pre-existing validation/read disagreement on a root mini stream whose
  size is not a multiple of 64, which is refused at read time;
- a placement-only harness control flag.

No coverage, claim or timing-contract promotion; the non-iWork goal remains
open.

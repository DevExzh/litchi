# 0570: read contiguous CFB FAT runs in one call

Status: retained. `performance_claim: none` — this record carries deterministic
read counts only. **No timing is measured and none is claimed.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was removed

`load_fat` and `load_minifat` in `litchi-cfb` each called `read_sector_into`
once per sector, so a file with N FAT sectors issued N reads regardless of how
those sectors sit on disk. `read_sectors_batched` already existed in the same
file and was already used by `load_directory`.

The corpus survey retained with change
[0565](0565-xls-globals-single-pass.md) identified this and sized it honestly:
across 98 container fixtures, 12 have more than one FAT sector, and batching
contiguous runs saves reads on **5 of them, 30 reads in total**, dominated by one
file. MiniFAT batching saves nothing on that corpus. This change implements it
and the measurement matches the model exactly.

## Why the existing helper was not reused

`read_sectors_batched` position-checks only the **first** sector of each run
against the file size and then clamps the read. A run whose later sector begins
at or past the end of the file would therefore be silently zero-filled instead
of raising `"Sector N is outside the file"`. Reusing it would have changed error
identity, so it is left untouched, and a bounded run reader was added instead.
`read_sector_into` is also untouched and still serves the DIFAT loop, which
cannot be batched because a DIFAT chain stores its successor inside each sector.

## The memory constraint, and how it is respected

Both loaders deliberately used one reusable bounded sector buffer and parsed each
sector straight into the final table, precisely to avoid allocating a byte buffer
as large as the whole table. Batching naively would have reintroduced exactly
that.

The scratch buffer is instead sized by the **longest contiguous run the sector
list actually contains**, clamped to a named 64 KiB ceiling, so a file that gains
nothing pays nothing. Peak retained memory per load:

| Case | Before | After |
| --- | ---: | ---: |
| single-FAT-sector file | one sector | **512 B** |
| 22 isolated FAT sectors | one sector | **512 B** |
| `FloatingPictures.doc` | one sector | 2,048 B |
| `picture.doc` (longest run 21 of 23 sectors) | one sector | 10,752 B |
| hard ceiling, any file | — | **65,536 B** |

`load_fat` still carries its 4 KiB stack array for the DIFAT chain, so its peak
is that frame plus the heap scratch. `load_minifat`'s stack array is gone.

## Measured effect

Reads and bytes counted through the crate's existing instrumented reader, over
every container fixture under `test-data/ole`.

| | files | total reads | total bytes |
| --- | ---: | ---: | ---: |
| before | 98 | **493** | 276,992 |
| after | 98 | **463** | 276,992 |

**Bytes read are byte-for-byte identical.** The change removes calls, not work.
Exactly five files differ, and they are exactly the five the survey named, each
saving exactly the predicted amount:

| Fixture | before | after | saved |
| --- | ---: | ---: | ---: |
| `ole/doc/picture.doc` | 26 | 6 | 20 |
| `ole/doc/testPictures.doc` | 8 | 4 | 4 |
| `ole/doc/FloatingPictures.doc` | 12 | 9 | 3 |
| `ole/xls/WithCustomViews.xls` | 5 | 3 | 2 |
| `ole/ppt/SampleShow.ppt` | 7 | 6 | 1 |

The other 93 files are unchanged, read for read. In isolation the loaders show
the mechanism more clearly: a 23-sector FAT in 3 runs falls from 23 reads to 3, a
5-sector contiguous FAT from 5 to 1, a 17-sector run at a 4,096-byte sector size
from 17 to 2, and a 129-sector MiniFAT from 129 to 2.

**On this corpus this is a small, concentrated result.** Thirty fewer positional
reads across 98 files, twenty of them in one 1.4 MB file, is not separable from
timing noise, and no timing claim is made for it.

### On a large file it is much larger than the corpus suggests

Change [0568](0568-xls-worksheet-window.md)'s latency capture measured this
change on the synthetic `xls-comments-opaque-heavy` corpus, a 16,995,840-byte
container, and separated the two changes by region:

| Selector | CFB structural reads | CFB structural bytes |
| --- | ---: | ---: |
| `xls_source_backed_open` | 265 → **8** | 136,704 → **136,704** |
| `xls_source_backed_open_list_worksheets` | 265 → **8** | 136,704 → **136,704** |
| `xls_source_backed_open_one_cell` | 265 → **8** | 136,704 → **136,704** |

**257 reads removed from a single open, with the bytes again identical.** A file
that large needs hundreds of FAT sectors, and a generated file lays them out
contiguously, which is precisely the shape this change collapses. The real-file
corpus under `test-data/ole` is made of small documents and therefore understates
the effect badly: it is the right corpus for asking how often this helps, and the
wrong one for asking how much.

On that corpus the change is also responsible for the whole of the measured
`xls_source_backed_open` and `..._list_worksheets` latency improvement, −7.40% to
−13.23% at the median in both directions, because change 0568's worksheet
component is zero for those two selectors. That measurement was taken with both
changes in one binary, so it is reported as attributable to this change by
region, not as an isolated timing of it. MiniFAT batching saves nothing on any surveyed fixture; it was implemented
because it shares the mechanism and the same memory argument, and because a
MiniFAT is the only table that can exceed 109 sectors without a DIFAT, which is
what makes the bounded multi-piece path testable at all.

## Validation preserved

Every check keeps its position and its identity. The decisions worth recording:

- **Per-sector position checks are kept, not collapsed.** `sector_position` runs
  for every sector of a run in list order, so `"Sector N is outside the file"`
  still names the same N.
- **A bad sector ends the run rather than erroring in place**, so sectors ahead of
  it are read first and the next call reports it. An earlier sector's I/O error
  therefore still wins, exactly as before.
- **Runs require strict adjacency and never assume ascending order.**
  `FloatingPictures.doc` genuinely lists sector 650 before 649, and a test pins
  that shape by asserting the second read is at a *lower* offset than the first.
- **The truncated-final-sector zero fill is reproduced rather than approximated.**
  A run is contiguous, so per-sector present-byte counts tile its range exactly;
  the test asserts the same table entry by entry, including the partial value.
- **Allocation ordering is preserved**, so an allocation failure still reports the
  historical resource label first.
- Fallible allocation only, no `unsafe`, and loop termination is type-enforced by
  a non-zero return so the run walk cannot spin.

## Correctness evidence

Nine tests were added, of which **five fail against the pre-change code**: the
contiguous FAT and MiniFAT run tests, the two bounded multi-piece tests, and the
batched truncated-final-sector test. The four that pass on both sides are the
point of the exercise rather than filler: they pin that isolated sectors still
cost one read each, that typed errors and their order are unchanged for both
loaders, and that the scratch is bounded by the longest run it will hold.

With this change and change 0568 applied together, the OLE2 crates pass **4,261
tests with zero failures** across 154 binaries.

## Limitations

No timing, allocation-profile, cold-cache, physical-device or cross-platform
result is claimed. The read counts are logical calls through an instrumented
reader over warm fixtures.

Two coverage gaps are recorded rather than worked around. The
`expected_fat_sectors` mismatch error is only reachable past the header's 109
DIFAT entries, and **no corpus fixture has any DIFAT sector at all**, so that
check remains as untested as it was; the DIFAT loop is unchanged. And the
bounded multi-piece path at a 512-byte sector size is exercised through
`load_minifat` rather than `load_fat`, because a 512-byte FAT run past the
128-sector bound would need more FAT sectors than the header can express; both
reach it through the same run reader.

# Log paragraphs for change 0648

Four blocks, one for each log the coordinator merges. Each is written to sit
above that file's newest section.

## For `HOTSPOTS.md`

## 0648 — the XLS whole-sheet walk's hotspot was the per-string read, and the string table fits one window

Change 0636 attributed `54016.xls`'s whole-sheet walk to the shared-string
resolver, which builds one fresh `stream_cursor_at_hinted` cursor and reads one
string per `LabelSst` cell: the walk resolves 16,055 shared strings in 16,077
chunk reads (change 0621's count) out of 16,105 positional reads in all, 6,214 of
them three bytes long, jumping backwards 4,444 times. It named the resolver for the queue. The
cheap-looking answer, a bounded sliding window over those accesses, is wrong,
and a temporary trace of every resolve says why: the access order has almost no
spatial locality, so a 64 KiB sliding window cuts that walk to 5,539 reads while
taking **341,656,514 bytes** of them against 323,234 on the per-entry path, and
a 4 KiB one takes 34,540,295. The right answer is that the table is **small**:
across the 108 corpus fixtures that carry one, the extents run from 8 bytes to
225,003, none above 256 KiB. Reading the whole table once — after eight entry reads, so a
selected-cell query never pays for it — costs `54016.xls`'s walk **one read and
225,003 bytes** where the per-entry path took 16,055 reads and 323,234 bytes,
better on both axes at once because the walk resolves 16,055 strings out of
7,893 distinct entries. Corpus-wide the walk falls from **26,655 reads to 959**
and the text projection from **29,708 to 1,226**, both for **fewer bytes**
(−4.03% and −4.52%), with `open`, `list` and `validate` identical read for read
and all 565 frozen outcome digests matching. Natively that is −67.85% cycles on
the walk and −67.64% p50 on an owned source, −80.10% over `FileSource`, and
change 0627's `xls_range_source_open_all_cells` falls from 16,145 physical
requests to **77**. The remaining 37 reads of that walk are change 0568's
worksheet window and the open; the shared-string path is down to one.
[Change and limitations](0648-xls-shared-string-resolver-window.md);
[evidence](results/change-0648/README.md).

## For `GOAL_AUDIT.md`

## 0648 — the range-source scenario 0627 called the most expensive is no longer expensive

Change 0627 called the `54016.xls` whole-sheet walk "the single most expensive
thing a range-source caller can ask an XLS reader for" — 16,145 requests,
modelled at 16.16 seconds of fixed service at its transport — and change 0636,
having found the walk's cost was not where its brief expected, left it open as
the named next item. This change closes it: `xls_range_source_open_all_cells`
and `xls_range_source_open_full_text` fall to **77 physical requests and
1,158,091 bytes**, modelled at **88.04 ms**, and the other three selectors carry
**identical** complete ordered request-sequence digests, so the open, the
worksheet listing and a selected-cell query did not move at all. That is
`docs/GOAL.md`'s "caller-supplied remote and range sources" dimension advancing
on the row it was worst on, and it does so without touching the source-backed
CRUD surface: no public API, no limit, no error type and no output byte changes,
and the corpus differential holds 565 frozen outcomes identical across 113
fixtures. Two things are not supplied. The range figures are change 0627's model
arithmetic over a deterministic request count, not observed service. And a
workbook whose string table exceeds 256 KiB gains nothing and is not measured —
no fixture in this corpus is in that class, which is a statement about the
corpus as much as about the change. OLE2 and OOXML remain the active priority;
ODF stays deferred and iWork excluded.
[Change and limitations](0648-xls-shared-string-resolver-window.md);
[evidence](results/change-0648/README.md).

## For `REPORT.md`

## 0648 — one read for the string table, measured on two source kinds and four instruments

`SharedStringResolver` — change 0585's per-scan state — now retains the
workbook's shared-string table for the life of one scan, once that scan has
taken more than eight resolves, and decodes each entry from a slice of it; a
table larger than 256 KiB, or a scan that resolves fewer strings, reads entry by
entry exactly as before. `resolve_shared_string_inner` also replaces its linear
segment scan with a binary search and drops the three per-resolve `Vec`
allocations survey item XLS-10 named. Deterministic counts over every `.xls`
fixture under `test-data` (126 files, 113 of which open, 565 operation rows per
leg): `all-cells` falls from **26,655 positional reads to 959** and **27,019
source observations to 1,323** for **4.03% fewer bytes**; `full-text` from
**29,708 to 1,226** reads and **79,061 to 50,579** observations for 4.52% fewer
bytes; `open`, `list` and `validate` are identical read for read, byte for byte
and observation for observation; no fixture reads more times on any operation;
and all 565 frozen outcome digests match. Change 0605's `xls_source_attribution`, an independent instrument measuring the
whole lifecycle on both source kinds, agrees: **16,145 reads and 16,117
observations become 77 and 49** on `54016.xls` in both `owned-readat` and
`file-source` mode, and swept over the whole corpus it is a second differential
— 252 rows per leg, **every outcome string identical**, `all-cells` reads
−89.54% and `full-text` −88.83% with the open included, and no row reading more
times.
Allocations on `54016.xls`'s walk fall
from **84,182 to 36,042** and allocated bytes from **10,281,754 to 2,144,490** —
5.24 per resolved string to 2.24. Change 0627's five XLS range-source selectors
put `open`, `list` and `one-cell` at identical request-sequence SHA-256 and take
`all-cells` and `full-text` from **16,145 physical requests to 77**. `perf stat`
isolation pairs: **−67.85% cycles / −68.46% instructions** on the `54016.xls`
walk, −43.72% / −37.73% on `WithCustomViews.xls`, −37.61% / −30.78% on the
`54016.xls` text projection, and **+0.10% / +0.02%** on the `open` control;
callgrind agrees within six points on every scenario. Paired timing, 120 samples
per leg, order A1 B1 B2 A2 on CPU 23: **−80.10% p50** walking `54016.xls` over
`litchi_core::FileSource` (16.00 ms → 3.19 ms), **−67.64%** on the owned source,
−43.82%, −34.89%, −28.15% and −16.86% on the other four scenarios, and −0.22% on
the `open` control, against A/A floors of at most **0.87%** in absolute value;
three windows repeat every scenario to within 3.3 points at p50. No scenario
got worse. The honest cost is `59858.xls`'s walk, which reads **15,630 more
bytes** for seventeen fewer reads — a scan that resolves just past the threshold
on a 16,384-byte table — and is still 18.73% cheaper in cycles. A sliding 64 KiB
window was implemented first and rejected on measurement: it read 341,656,514
bytes on the walk that reads 323,234 per entry, and the existing
`the_whole_sheet_walk_costs_one_scan_not_one_per_cell` test caught it.
[Change and limitations](0648-xls-shared-string-resolver-window.md);
[evidence](results/change-0648/README.md).

## For `ADR_COMPLIANCE.md`

## 0648 — a cache whose ceiling is the region it caches, and whose region the open already read

The window is **all or nothing**, and that is the compliance argument as much as
the performance one. It holds the string table's whole source extent or holds
nothing, so ADR 0003's bounded-resource rule is met by a ceiling that is a named
256 KiB constant *and* by the region itself: one window per resolver, grown with
`try_reserve_exact` and reported as `SourceBackedError::Allocation { resource:
"retained SST window" }` on failure, released when the scan ends, and refused
outright rather than truncated when a table is larger. The window is exactly the
extent the open's own segment table describes, so it can read **no byte the open
did not already read** to build the entry locators; a segment table that reached
past the declared workbook stream length would disable it rather than be clamped
into a region that could leave a late entry outside. ADR 0005's lazy-payload
contract is observed: the table is read on demand, through the cursor's own
validated chain walk, by a scan that has shown it will use it, and no stream is materialized —
change 0576's rule that the owner retains locators and not text is untouched,
because nothing here outlives the resolver. ADR 0006 is untouched: no execution
context, no worker pool, no ambient I/O, no lock, and the cancellation checks are
taken in the same places and the same numbers. No new `unsafe`, no new
dependency, no public API change and no weakened limit; `max_sst_entries`,
`max_global_bytes`, `max_text_cells` and `max_text_bytes` all run where they
ran. The one contract that moves is change 0621's observation width, and it
moves exactly as change 0636 moved it for the validation walk: a resolve served
from retained bytes takes no read and so no observation, which reports a
mutation at the next **fill** rather than at the next resolve, bounded by the
`SourceInner::ensure_current` bracket every operation still ends with. Change
0621's change-under-read sweep is extended to that boundary and requires the
same typed refusal from every observation window of a walk that crosses it. An
entry is never assembled from two reads, because the table is read whole or not
at all.
[Change and limitations](0648-xls-shared-string-resolver-window.md);
[evidence](results/change-0648/README.md).

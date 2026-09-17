# Log paragraphs for change 0668

Four blocks, one for each aggregate log. Each block is written to sit above
the newest section of its file.

## For `HOTSPOTS.md`

## 0668 — the remaining packed-cell work is measure-only, and locator scratch is gone

Queue row 12's `MulRk`/`MulBlank` residue is implemented. A selected-cell scan
now validates each packed cell's range and XF through `MeasuredCell`, creates an
88-byte `CellRecord` only when the coordinates select the target, and decodes
an RK value only for that target. The 0641 census puts the affected path at
**60,575 of 127,072 cell values (47.7%)**. The source-open SST locator scan
also drops its temporary `Vec<&[u8]>`; it uses the `RecordRef` payloads already
held by the open and preserves every locator and refusal. The base/after
allocator leg removes one allocation and **448 bytes** on `54016.xls` and one
allocation and **208 bytes** on `WithCustomViews.xls`, exactly the fat-pointer
vector sizes. [Change and limitations](../../0668-xls-query-residues.md);
[evidence](README.md).

## For `GOAL_AUDIT.md`

## 0668 — XLS query work advances without widening the retention contract

The XLS query path now eliminates proven transient work for packed numeric and
blank cells while retaining the existing validation and output behavior. The
open-time locator path uses the source records it already owns, and its pinned
SST digest remains `0x9cb14f5daa02eebc` across 117 indexed fixtures and four
typed refusals. The retained cross-query sheet index and snapshot chain hint
remain queued behind ADR 0005's bounded weighted evictable cache requirement;
the two 0633 framing passes remain queued behind their coverage-proof order.
No latency or speedup claim is registered, so the goal audit advances the
correctness-preserving work item with deterministic allocation evidence.

## For `REPORT.md`

## 0668 — packed-cell measure-only validation and a smaller SST locator scan

The `scan_worksheet` `MulRk` and `MulBlank` arms now follow the same
measure-only rule as the seven scalar cell kinds. They still visit and validate
every packed cell, while selected queries retain only their target. The SST
open scan replaces a temporary slice-of-payloads vector with a cursor over the
existing `RecordRef` list. Base versus after allocation counts on the two SST
fixtures move from 246/754,968 to 245/754,520 and 208/153,205 to 207/152,997;
the whole-sheet walk remains 35,793 allocations and 1,389,460 bytes. The full
`litchi-xls` suite passes, including the packed-cell equivalence and corpus SST
differential tests. The packet reports no timing floor, so these counts carry
no registered performance claim.

## For `ADR_COMPLIANCE.md`

## 0668 — the XLS residue follows the existing lazy and bounded-state rules

The packed visitor validates the complete BIFF record before invoking the sink,
keeps XF validation for every cell, and materializes a value only after a
selected coordinate matches. The SST cursor consumes the same already-read
`RecordRef` payloads and keeps the existing segment, entry, and framing
boundaries. The change adds no cache, lock, reader-free retained position, or
global pool. Row 11 remains deferred because ADR 0005 requires a bounded,
weighted, evictable clean-value cache with accounting; row 12's framing fusion
remains deferred because 0633's refusal order and coverage proof have priority.
No public API, limit, error, fence, output byte, dependency, or unsafe code
changed. [Decision record](../../0668-xls-query-residues.md);
[packet](README.md).

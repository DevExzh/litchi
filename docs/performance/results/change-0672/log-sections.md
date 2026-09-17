# Log sections for change 0672

## For `HOTSPOTS.md`

## 0672 — the stored-route XLSX `cells` vector now reserves its exact sparse size

Record: [0672](../../0672-xlsx-stored-cell-allocation.md). Queue row 15's
second residue is landed: `collect_stored_cells` no longer grows one slot at a
time. `Rect::ALL` uses the parsed store's exact record count and narrower
ranges count their sparse selection before one `try_reserve_exact`. On the
75,770-cell POI worksheet, warm `cells` falls from 16 allocations and
10,485,760 bytes to one allocation and 6,061,600 bytes; the generated dense
control keeps the same 5,608,448 bytes and loses 14 vector-growth calls. The
cold selected scanner remains intentionally materialized because yielding would
move a refusal past a partial result and needs its own ADR 0006 design.

## For `GOAL_AUDIT.md`

## 0672 — bounded owning conversion sized from validated sparse records

Record: [0672](../../0672-xlsx-stored-cell-allocation.md). The change follows
the goal's allocation-elimination tier while preserving the typed error and
refusal order. The stored `Store` is immutable after validation, so counting
the selected records and copying the same iterator cannot produce a partial or
guessed result. Allocation failure still occurs before the returned vector is
observable. The cold scanner's EOF-before-publication guarantee is unchanged;
its vector is explicitly left open for a separate refusal-order design.

## For `REPORT.md`

## 0672 — exact stored-route reservation removes the 42% capacity residue

`SourceWorksheet::cells` previously used `try_reserve(1)` for each stored cell.
On the 75,770-cell POI worksheet that reached 131,072 slots, or 10,485,760
bytes. The landed route reserves 75,770 slots once, retaining 6,061,600 bytes:
one allocation and 42.2% fewer bytes in the warm result. A paired pinned warm
run has a 1.770% p50 same-binary floor and reports 3.751% and 4.134% p50
improvements in the before/after directions. The p95/p99 floor is noisy and no
tail claim is made. Cold POI allocation bytes fall by 4,424,160 while peak live
bytes stay unchanged, and the selected scanner's own record vector is outside
this change.

## For `ADR_COMPLIANCE.md`

## 0672 — no refusal, ownership or public API boundary moved

Record: [0672](../../0672-xlsx-stored-cell-allocation.md). ADR 0003's explicit
owned conversion remains `cells` and returns the same sparse values in the same
order. ADR 0005's resource boundary remains fallible through `try_reserve_exact`;
the change introduces no cache, lock, executor or ambient resource. ADR 0006's
preservation and refusal rules are unchanged. In particular, the selected
scanner still reaches EOF and validates retained records before a callback, so
no typed refusal is traded for a partial result. The 0652 decision record and
0642's accepted refusal-order design are cited as authority; no ADR amendment
is proposed.

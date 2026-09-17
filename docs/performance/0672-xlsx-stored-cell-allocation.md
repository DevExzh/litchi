# 0672: reserve the stored XLSX cell result exactly

Status: retained, implemented in `litchi-xlsx`. `performance_claim: none` —
the allocation and paired timing measurements below are scoped evidence, not a
registered claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

## Authority

This change implements the second half of queue row 15 in change
[0651](0651-queue-refresh-after-the-second-wave.md): the stored-route
`cells` result over-allocated by 42%. The first half remains deliberately
unimplemented: the cold selected scanner retains its `Vec<SelectedRecord>`
until the worksheet reaches EOF, because making it yield would move a typed
refusal after a partial result.

Change [0642](0642-xlsx-visit-cells-streaming.md) measured the residue on the
stored route: a 75,770-cell whole-sheet result grew to 131,072 slots, retaining
10,485,760 bytes for 6,061,600 bytes of actual `SourceCell` elements. The
owner's standing decisions in [0652](0652-owner-decisions-for-the-third-wave.md)
keep correctness, preservation, bounded resources and refusal-before-result
ahead of speed. This change changes only the owning result's reservation.

## What changed

`collect_stored_cells` now determines the number of sparse records selected
before cloning them into `Vec<SourceCell>`, then calls `try_reserve_exact` once.
The full-grid `Rect::ALL` route uses the store's already-known
`stored_cell_count`, so the common whole-sheet read does not traverse the
records twice. A narrower range counts its sparse iterator first, preserving
the old sparse result shape rather than reserving the enclosing rectangle.

The store is immutable for the duration of this read. The count and copy
therefore see the same validated records. Allocation failure remains the same
typed `allocation("source-backed selected cells", ...)` error and still occurs
before any owning value is returned. Cell order, cloning, fences, cancellation,
and all selected-scan behavior are unchanged.

The selected scanner's vector is intentionally untouched. Its records are
published only after worksheet EOF, dependency resolution and validation, which
is what keeps a malformed later row from arriving after an earlier callback.
Moving that refusal boundary needs a separate ADR 0006 design record.

## Breaking changes

None. The public `SourceWorksheet::cells` signature, result type, errors,
ordering and semantic values are unchanged. No new dependency or `unsafe` was
introduced.

## Why this is sound

The stored route owns a parsed `Store` that has already passed worksheet
validation. `Store::stored_cell_count` is the exact length of the immutable
stored record slice, and the subset count is obtained from the same sparse
iterator used by the copy. Reserving that count cannot omit a selected record;
the loop still pushes every iterator item in its existing row-major order.

The cold selected route is not affected. `visit_cells` still reaches EOF and
finishes all dependency and source fences before the first callback, and the
scanner's retained records still remain available for the same refusal order.

## Evidence

The focused before/after probe compares base `5fa92d7ce` with this branch on
the POI `no_drawing_patriarch.xlsx` worksheet (75,770 stored cells) and a
generated 2×256×256 integer workbook. On the warm stored route, the POI result
falls from 16 allocations and 10,485,760 allocated/peak bytes to one allocation
and 6,061,600 bytes. The dense result already had a power-of-two capacity, so
allocated bytes stay at 5,608,448 while allocation calls fall 65,551 → 65,537.
Cold POI allocation bytes fall 174,657,136 → 170,232,976 while its peak live
bytes stay 57,025,552; the cold selected dense route is unchanged.

Two fresh before/after differential binaries agree on all three scoped
worksheets (the POI sheet and both generated sheets), with the same counts and
FNV digests on cold `visit_cells`, cold `cells`, warm `visit_cells` and eager
`Workbook` reads. A 30-sample pinned `cells-warm` pair on the POI sheet has a
same-binary p50 A/A floor of at most 1.770%; the before/after p50 direction is
3.751% faster in A1→B1 and 4.134% faster in B2←A2. The p95/p99 floor is noisy
in this one window and no tail claim is made.

The focused 0642 refusal/order tests plus the new exact-capacity test pass, and
the full `litchi-xlsx` library suite passes 1,031 tests. The retained packet is
[`results/change-0672/`](results/change-0672/).

## Limitations

The packet measures two worksheet shapes on one host and does not claim RSS,
cold-cache latency, throughput, range-source behavior, other formats or other
platforms. Narrow stored ranges use a count pass and were covered by the unit
test but were not timed. The selected scanner's cold peak remains open work
under the refusal-order design question.

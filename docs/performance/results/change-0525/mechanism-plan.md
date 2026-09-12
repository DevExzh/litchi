# XLSX reconstruction work elimination: candidate contract

This is a pre-capture design, not an adopted optimization or a performance
claim. OLE2/OOXML remain the priority; ODF is deferred and iWork excluded.

## Evidence and choice

The 451 current XLSX source files match the 0522 baseline manifest exactly
(`prior-source-binding.json`). That baseline puts reconstruction at roughly
61% of the isolated commit instruction total. Full raw worksheet parsing is
74,805,030 of 197,179,569 commit instructions for medium repeat 1, separately
from the full-output value-only XML validator. These are nested instruction
costs, not additive latency estimates or per-operation allocation counts.

The first row-only design was ruled out before implementation or captures.
`closure_coverage.py` reproduces the deterministic corpus selectors: a 1%
dense-sparse edit touches rows containing 16,769 of 17,792 cells. Omitting only
untouched rows could avoid parsing just 1,023 cells (5.75%). Omitting unchanged
cells could avoid 17,614 cells (99.00%). Medium has the same approximate 99%
cell opportunity. These static counts justify the finer closure; they do not
predict a measured speedup.

## Required ownership and independent readback

The ordinary writer continues to emit identical full worksheet bytes. A
private value-only route may additionally record spans that it copied from
the immutable source, together with the exact original cell addresses those
spans own. The proof must be bound to that source; a proof from a different
snapshot cannot authorize reuse of its Store.

`Snapshot` must validate the complete emitted XML before using any omission.
The reduced semantic readback document retains the emitted changed cells,
all enclosing namespace context, row tags, worksheet metadata, columns,
defaults, and declared dimension. It omits only proven unchanged cell owners.
The existing raw worksheet parser reads this reduced document. The resulting
Store supplies all parsed metadata and every changed cell; only proven
omitted cells come from the original Store. Indexes and extents must describe
the combined Store. Publication readback still compares parsed emitted values
with staged expectations, never staged expectations with a copy of themselves.

## Closure rules

- A row with replacements only keeps its membership and inferred coordinates.
  The writer already emits an explicit address for each changed cell, so its
  coordinate remains independently parseable after preceding cells are omitted.
- Insertions/removals within an existing row require parsing the entire
  emitted row, including unchanged followers whose implicit coordinates may
  shift. Their old cell semantics must not also be merged back.
- New rows use the existing full parser path: insertion can change later
  implicit row numbering. Shared-formula worksheets also use that path because
  formula resolution can cross omitted cell owners. These are explicit
  dependency boundaries, not claims that the fallback is optimized.
- Full validation continues to reject unsupported dependency-bearing, MCE,
  foreign and unknown markup. Namespace declarations inside an omitted cell
  cannot escape its complete element scope. All outside context is retained.
- Keep source/version/cancellation fences, output and aggregate byte limits,
  style and cell provenance, calculation-chain invalidation, exact no-op
  sharing, reversible patch bytes and publication checks.
- Reserve scratch and merged cell storage with checked sizes. Record the
  extra reduced-document scratch in allocator/peak results. Do not claim
  bounded constant memory or zero copying.

## Admission

The frozen `plan.json` will bind exact candidate files before builds/captures.
Fresh matched normal and allocator binaries are required despite unchanged
historical source. Use native ABBA, separate allocator repeats and isolated
commit Callgrind profiles. Whole-child hardware counters remain explicitly
separate from operation-scoped attribution. The plan's primary total/commit
latency and commit instruction gates decide retention; retain and review
individual adverse results. No architectural proof alone admits the candidate.

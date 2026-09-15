# Change 0622 log paragraphs

The coordinator merges these into the four repository logs. Written in the style
of each log's newest section.

## HOTSPOTS.md

**XLSX-1 is implemented (change 0622).** The source-backed value editor's commit
no longer re-scans the touched worksheet into a `Layout`. The planning traversal
now retains sixteen bytes per cell — the resolved address and the source span of
its `<c>` element — plus one record per row and the `<sheetData>` and
`<dimension>` spans, and the commit rewrites from those, materializing the
scanner's slot only for the at most 256 cells one commit may touch. Measured on
`taskset -c 15`, callgrind isolation pairs (N=1, N=11): the whole
`rewrite_value_only_with_provenance` falls from 18.4 M to 0.21 M Ir per operation
on the harness `medium` shape, 129.7 M to 1.32 M on `dense-sparse` and 25.6 M to
0.21 M on `noncompact`, while planning rises 15.0-17.7%; planning plus commit
together fall 33.6-41.4%. Commit allocation calls fall 53.2-71.9% and commit
allocated bytes 51.4-75.2%; planning allocated bytes rise 3.90-6.09% and planning
allocation calls 0.02-0.12%. Paired p50 is 12-20% faster on the source-backed
cell-value edit-and-save selectors, against an A/A floor stated in the record.
The remaining XLSX opportunities are unchanged: XLSX-2 (admission-surface
widening, gated on the `litchi-opc` compactness prerequisite change 0602 named),
XLSX-3 (publication audit reuse) and XLSX-4 (overlay the candidate `Store`).
**A new hotspot is now visible on this path**: with the second scan gone, the
planning traversal is the whole cost of an XLSX value commit, and the fact
builder is 15-18% of it; the next reduction on this route is the builder's
per-cell `r` parse and per-element ampersand probe, not the writer.

## GOAL_AUDIT.md

Change 0622 closes item XLSX-1 of the 0587 survey and, with it, the line
0550 opened ("the next task is a private source-bound layout proof that could
avoid the second complete source scan while preserving lexical spans, scanner
facts, error order, resource limits and output validation"). It satisfies 0550's
admission condition — a candidate must improve planning *plus* commit and
publication rather than shift work — with a measured 33.6-41.4% reduction in
planning-plus-commit instructions and no change to publication. It also clears
the two gates that rejected changes 0552 and 0553: valid no-op planning is
unaffected (the builder never runs when the plan is empty, and an empty plan
returns before the fact route is consulted), and process peak RSS rises 4.23% at
the median on the shape where 0552 measured +6.97% and 0553 +5.16%, inside that
leg's own 6.49% run-to-run spread. The optimization order in `docs/GOAL.md` is
respected: this removes unnecessary work and unnecessary parsing, ahead of
layout, algorithms and parallelism; nothing was vectorized and no parallelism was
introduced. **The audit gap this change does not close**: the route it optimizes
is unreachable on real producer files. Of 391 real worksheet parts in
`test-data/`, the value-only planning validator accepts exactly one, and that one
publishes facts; change 0602 already recorded that the editor admits none of the
95 real `.xlsx` fixtures. The measured win is on the harness corpora.

## REPORT.md

`litchi-xlsx`: the source-backed cell-value commit stopped reading the worksheet
a second time. Previously an open-edit-save of one cell walked the touched sheet
about six times; one of those walks — the commit's layout scan, measured by
change 0550 at 53.24-54.94% of commit instructions — is gone, replaced by a
sixteen-byte-per-cell note taken during the walk that already happens. On the
harness corpora a one-cell source-backed edit and save is 12-15% faster at p50
and a one-percent edit 15-20% faster, with commit allocations roughly halved.
Output is byte-identical: `output_sha256` matches on both legs for every case and
shape measured, and an in-crate differential oracle compares the two commit
routes cell by cell over the synthetic shapes, a dense grid and every worksheet
part of every `.xlsx` fixture in the repository. No public API changed and no
error moved. The saving does not reach files produced by Excel: the value editor
that owns this path refuses them for reasons change 0602 recorded, and this
change neither widens nor narrows that surface.

## ADR_COMPLIANCE.md

Change 0622 (`litchi-xlsx`, compact planning facts) is compliant and adds no ADR
question. **ADR 0003 (bounded resources):** the retained state is a flat
`Box<[CellFact]>` and `Box<[RowFact]>` with no per-cell heap object, bounded by
the worksheet the planning `Store` already holds; every push uses `try_reserve`
and a failed reservation declines rather than errors. The 256-action commit cap
is unchanged and now also bounds how many cell slots are materialized. No new
`unsafe`, no new dependency, no weakened limit. **ADR 0005 (validation
placement):** no validation moved. The value-only XML validator still sees every
event of the source first and still owns the first error; the raw parser still
builds the `Store` from the same transition function; the commit's `scan` remains
the authority and still runs whenever the builder declined. The builder never
returns an error and never stops the traversal, so no provisional diagnostic can
become a public planning error. **ADR 0006 (lossless preservation):** output
bytes are unchanged, proved by the harness's package hashes and by the in-crate
oracle; change 0525's independent changed-cell readback and change 0528's two
publication audits are untouched. **Contract movement: none.** The builder's
allow-list narrows only which worksheets take the fast route, never which
worksheets are accepted, which errors are raised or in what order.

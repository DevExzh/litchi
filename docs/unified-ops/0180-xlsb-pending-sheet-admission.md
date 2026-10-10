# 0180 — XLSB pending sheet admission

Pending homogeneous visibility runs now admit each direct positional command
without rescanning the complete tab catalog. Shared admission retains property,
path and quota validation; the run retains its immutable catalog and visible-tab
count. Selected unbound tabs retain their explicit capability refusal. First
commands and post-barrier replans retain full-catalog validation. Repeated targets,
state changes, active-tab hides and quota boundaries still flush and replan.
Source fences, deferred error precedence and last-visible policy are unchanged.

Independent implementation and read-only review found no semantic regressions.
All 561 XLSB-enabled facade tests/doctests passed, including the twelve focused
sheet-batch regressions. Minimal XLSB all-target Clippy with warnings denied,
scoped rustfmt and diff checks passed. Verification used an external copied
lockfile with cached cc 1.6.0; repository lockfiles are unchanged. No base-only
commits were pending. Temporary worktree and build/profile artifacts were removed
once their reduced evidence was recorded.

## Measurements

[Before](results/0180-xlsb-sheet-batch-before.json),
[after](results/0180-xlsb-sheet-batch-after.json) and
[repeat](results/0180-xlsb-sheet-batch-repeat.json) cover 1, 10, 100 and 1,000
commands, eager/source owners, three ABBA rounds and six samples per leg on
CPU 0 without concurrent builds/tests. The unchanged harness and all three
builds use release debug information. Evidence records binary hashes and exact
source provenance. Every case has exact byte equivalence, one commit per leg
and an empty failure matrix.

| 1,000 commands | Before native / facade ms | After native / facade ms | Repeat native / facade ms |
|---|---:|---:|---:|
| eager | 137.105 / 154.024 | 140.815 / 143.529 | 143.190 / 144.816 |
| source | 171.362 / 188.758 | 174.552 / 175.829 | 175.338 / 175.830 |

After/repeat facade overhead is 1.1–1.9% eager and 0.3–0.7% source at 1,000
commands. These scoped measurements do not establish full acceptance or linear
native scaling; 10,000-command acceptance remains open.

## Native publication follow-up

The [reduced CPU profile](results/0180-xlsb-sheet-profile.json) records 492
samples without lost samples. It covers the whole benchmark process, including
construction and warmup, and its timings are not acceptance evidence. Stack
unwinding was truncated; sampled symbols identify signature scans as a hotspot,
but caller attribution comes from independent source inspection.

`apply_sheet_view_dependencies` calls `get_part_mut` for each changed worksheet
view. That calls `revoke_exact_source`, which scans the full graph through
`is_signed` for source-ingress packages. N changed views in an N-part package
therefore repeat quadratic signature scanning. Unmarshalling itself performs
one signature scan rather than one per part.

A follow-up should stage payload replacements behind one validated package
mutation boundary. Simply removing revocation scans risks losing the sticky
signature-disposition requirement after low-level graph changes. Native policy
code is unchanged here. General sheet-command coalescing, native publication
scaling and the all-format/Python/CLI goal remain open.

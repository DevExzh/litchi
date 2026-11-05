# 0173 — XLSB direct cell write resolution

## Change

A scoped direct cell write resolves its sheet and checked coordinate through
the shared path resolver, then leaves the worksheet read grid deferred. The
native package cell editor continues to own record, formula-group, dependency
and publication validation. Reads, queries and range expansion keep their
existing grid-backed navigation. No second path parser or document model is
introduced.

The deferred cell source retains the native sheet handle and an inline
`OnceLock<Grid>`; resolving a write target does not allocate a cache or acquire
the read grid. If an internal consumer inspects that node, its normal cell,
format and merge accessors load the same grid lazily. Root-level `/A1` remains
eligible to name a sheet, and complete-catalog tab checks still reject chart
sheets as cell destinations.

Four source/eager regression tests cover named, case-insensitive and quoted
paths, ranges, chart-first catalog positions, missing names, invalid paths and
coordinates, snapshot atomicity, and read/rename transaction barriers. Sheet
names containing spaces use the existing quoted Excel-path grammar.

## Validation and evidence

The XLSB facade suite passed 525 tests/doctests. The final combined
`docx,xlsx,pptx,xlsb,encryption` suite passed 1,194 tests/doctests. Combined
all-target Clippy and rustdoc passed with warnings denied; the minimal-feature
check, scoped formatting and diff checks passed. Root builds used an external
lockfile copy selecting cached `cc` 1.6.0, offline resolution, external targets,
zero dev/test debug information and no incremental compilation. Repository
lockfiles remain unchanged.

The crate-boundary audit passed for 66 packages and 268 internal dependency
declarations. Independent review found no correctness blocker.

[Before](results/0173-xlsb-cell-values-before.json) and
[after](results/0173-xlsb-cell-values-after.json) retain nine ABBA rounds,
18 samples per leg, CPU 0, exact warm-up output hashes and build provenance.
No validation or build job from this turn ran during either measurement.
Both runs have twelve exact ZIP-byte comparisons, empty failure matrices and
one native commit per case. The after report includes a source-file hash
manifest; the before report identifies the exact baseline commit.

The eager native leg starts from `Package`, while the eager facade starts from
`Workbook`. Source promotion is timed in both source legs. Caller path/property
preparation, opening and serialization are outside the clocks; facade property
recollection, admission and coercion remain timed. The harness source and README
now state those distinctions explicitly. This compares complete edit routes,
including their different publication requirements.

| Shape | Cells | Owner | Before unified/native | After unified/native |
|---|---:|---|---:|---:|
| set | 1 | eager | 1.107 | 1.109 |
| edit | 1 | eager | 1.103 | 1.118 |
| set | 1 | source | 1.049 | 1.033 |
| edit | 1 | source | 1.056 | 1.036 |
| edit | 10 | eager | 1.149 | 1.155 |
| edit | 10 | source | 1.066 | 1.057 |
| edit | 100 | eager | 1.421 | 1.371 |
| edit | 100 | source | 1.225 | 1.199 |
| edit | 1,000 | eager | 1.677 | 1.611 |
| edit | 1,000 | source | 1.604 | 1.515 |
| edit | 10,000 | eager | 1.726 | 1.654 |
| edit | 10,000 | source | 1.740 | 1.643 |

Large batches remain below the 2x review threshold and show lower ratios in
this paired run. **The small eager edit requirement remains open:** one-shot
native/unified medians are 105.100/116.561 microseconds (1.109), and the
one-operation transaction ratio is 1.118. The before ratios were 1.107 and
1.103, respectively. This change does not demonstrate an eager small-edit
improvement or close the 10% requirement. Native workbook parse reuse is the
next concrete follow-up; no broader performance acceptance is claimed.

Reproduce using the harness README's build and provenance commands and
`taskset -c 0 .../unified-xlsb-cell-value-baseline --rounds 9`.


## Remaining scope

The native package cell editor still reconstructs a workbook to resolve its
worksheet, and facade adoption reconstructs the committed workbook again.
A future native owner seam could reuse validated projections, provided it
retains the physically admitted OPC graph and exact patch endpoints. These
changes are outside this target-resolution optimization.

This generated numeric workload does not cover formula-group-heavy or
shared-string-heavy scaling, allocation counts or RSS. XLSB formula authoring,
formatting, structural facade CRUD, creation and durable cell history remain
open, as do the other format and Python requirements of the full goal.


The branch already contained the rewritten base at turn start. Temporary
implementation worktrees and branches were removed after integration. Scratch
builds, copied locks, patches and raw logs were removed after recording the
reduced evidence. No binary or large evidence artifact is committed.

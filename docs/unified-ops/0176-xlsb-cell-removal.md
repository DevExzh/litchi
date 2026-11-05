# 0176 — XLSB cell removal

## Native contract

The package-owned `cell_values::Edit::remove` now returns a changed flag.
It removes the stored cell record, including its header and formula payload,
without shifting adjacent coordinates. An absent coordinate or a repeated
removal is an exact no-op. The raw worksheet primitive retains its stricter
low-level behavior.

Admission refuses array/shared/data-table formula-group members, even if a
member's own record is absent, and ordinary formula deletion when an opaque
calculation chain is present. Unsupported or ambiguous cell records are not
silently treated as missing. Existing publication reconciles shared-string
occurrence counts and retains the source policies and exact inverse image.
Deleting the final shared-string cell leaves a valid empty SST envelope.

Native regressions and an extended calculation-chain regression cover
idempotence, insertion cancellation, exact inverse replay, neighboring cells,
shared-string closure, grouped formulas including missing followers, and
ordinary numeric/Boolean/string formula-cache record removal. Reserved cell-family records are refused before their coordinate layout is
interpreted. Cancelling a newly staged formula insertion remains an exact no-op,
even when the source has an opaque calculation chain.

## Unified integration and evidence

`File::remove` resolves direct cell paths, ranges and cell queries against the
immutable snapshot. Ranges expand into coordinates without shifting adjacent
cells; queries can span worksheets. Admission bounds target and staged-byte
counts before native publication, with direct range charge checked before target
expansion. One package-owned native edit covers each selected worksheet.

Explicit transactions and atomic JSON batches coalesce distinct direct scalar assignments and removals
in the same worksheet edit. Repeated coordinates, reads, ranges, queries and
worksheet changes close the pending run. Best-effort batches isolate individual
items. Native failure attribution retains the originating operation index;
failed atomic work leaves the source snapshot unchanged.

Capabilities advertise removal only for cell elements, with explicit contextual
gaps for non-cell targets, formula groups, calculation chains and unmodeled
cell-record families. The generated OfficeCLI delta report follows the updated
capability golden. Coordinate-anchored comments, hyperlinks, merges and other
metadata remain valid on empty coordinates; removing a stored record neither
shifts coordinates nor garbage-collects those objects.

The final native suite passed 1,144 tests/doctests. The combined
`docx,xlsx,pptx,xlsb,encryption` facade suite passed 1,207 tests/doctests,
including ten new removal tests. The 1,000-command atomic batch regression
retains one native commit and reports a missing-cell removal as unchanged.
Native all-target Clippy and rustdoc JSON passed with warnings denied; the full
XLSB glossary profile passed all 31 name bans. The boundary audit passed for
66 packages and 268 internal dependency declarations, retaining the existing
11 explicit debt items. The generated OfficeCLI parity report is deterministic,
and its generator's 14 unit tests passed.

Combined facade all-target Clippy and documentation passed with warnings
denied, as did all-target Clippy with no format features. Scoped formatting
and diff checks passed. Builds used offline resolution, external target
directories and an external workspace lock copy selecting cached `cc` 1.6.0;
the inherited repository lockfiles remain unchanged.

## Paired performance

The standalone [harness](../../tools/perf-baseline/unified-xlsb-remove/README.md)
measures eager and source owners, 1/10/100/1,000/10,000 cells, direct/range,
closure transaction and atomic command routes as applicable. Each of four
matrices contains 24 cases with nine ABBA rounds (18 samples per leg), pinned
to CPU 0 with no concurrent Cargo/test job. Input construction, opening,
serialization, output checks and teardown are outside the timer. Source
promotion is inside both measured edit routes. Eager native starts from its
Package owner; the facade starts from its retained Workbook owner.

The first two matrices exposed quadratic native removal: the raw primitive
linearly searched every stored cell for every removal. It now uses the existing
checked coordinate index, retaining duplicate/absent admission. This change
benefits both routes. All four matrices have empty failure matrices and exact
native/unified ZIP equality; output hashes also match across the optimization.
Every command batch retains one native commit.

| Owner / route | Cells | Before unified ms | After native ms | After unified ms | After ratio | Repeat ratio |
|---|---:|---:|---:|---:|---:|---:|
| eager / range | 1,000 | 2.227 | 1.715 | 1.747 | 1.018 | 1.025 |
| eager / command_batch | 1,000 | 2.901 | 1.598 | 2.484 | 1.555 | 1.563 |
| source / range | 1,000 | 2.907 | 2.404 | 2.433 | 1.012 | 1.022 |
| source / command_batch | 1,000 | 3.713 | 2.259 | 3.316 | 1.468 | 1.467 |
| eager / range | 10,000 | 60.612 | 14.269 | 14.541 | 1.019 | 1.027 |
| eager / command_batch | 10,000 | 71.728 | 14.347 | 23.070 | 1.608 | 1.598 |
| source / range | 10,000 | 64.626 | 18.309 | 18.534 | 1.012 | 1.022 |
| source / command_batch | 10,000 | 76.839 | 18.330 | 29.016 | 1.583 | 1.578 |

Across both after matrices, one-cell ratios range from 1.007 to
1.072, below the 10% review threshold. All measured 1k/10k ratios
remain below 2x. Unified 10k/1k growth ranges from 7.617 to
9.757x, consistent with linear scaling over this measured interval.
These are generated numeric-cell removal results, not an all-workload claim.

Evidence: [before](results/0176-xlsb-remove-before.json),
[before repeat](results/0176-xlsb-remove-before-repeat.json),
[after](results/0176-xlsb-remove-after.json),
[after repeat](results/0176-xlsb-remove-repeat.json).
Each records raw samples, corpus/output hashes, source-file SHA-256 manifest,
source commit, compiler/machine, build command and binary SHA-256. The before
manifests intentionally retain the old lookup source hash. The two after
manifests match the final Rust/manifests used in this change. No binary is
committed.

After the lookup optimization, all 1,144 native tests/doctests and 34 focused
facade tests passed again, along with native all-target warning-denied Clippy.
The base branch is already contained (zero commits ahead of this branch).
The three milestone worktrees and branches were removed; external build
outputs and scratch logs are cleaned after recording evidence.


## OfficeCLI delta and remaining scope

OfficeCLI's ordinary single-cell removal also deletes the cell without a shift,
but refuses an absent stored cell and drops the calculation chain. Litchi's
coordinate projection includes virtual empty cells, so removing one is an exact
no-op. Calculation-chain and grouped-formula deletion remains a typed refusal
until the native owner can update that dependency closure. Explicit cell shifts,
worksheet lifecycle, formula authoring and formatting remain follow-ups.
The full all-format, Python, CLI and acceptance goal remains active.

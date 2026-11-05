# 0068 — XLSB unified read adapter

## Scope

R4d connects XLSB to `File::get`, `find`, `query`, node views, document
views, and capability discovery. The adapter reads the existing eager or
source-backed leaf owner. It shares spreadsheet vocabulary and inline coordinate
labels with XLSX; format-specific capability tables describe leaf differences.

Workbook, sheet, cell, range, row, column, table and defined-name nodes use
semantic paths. Complete catalog ordinals include chart tabs. Unicode caseless
name collisions reject named selection while positional selection remains
available. Unified cell coordinates are zero based and path ordinals one based.
The legacy worksheet trait's source/eager coordinate discrepancy remains a
separate migration task.

## Ownership and fallback

Nodes pin leaf grids instead of copying worksheets into an operations model.
The adapter retains a semantic tab name index and defined-name metadata once.
Sparse cell iteration uses a bounded coordinate buffer, terminates at exhaustion,
and honors query limits. Dense range expansion requires an explicit node budget.
Table-only traversal reads selected-sheet table parts without acquiring a grid.
Header detection retains sparse regions, without allocating every covered cell.

A source full-grid refusal classified as `unsupported_capability` can use the
leaf's binary index to build a scalar-only grid. Corruption, cancellation,
resource-limit and stale-source errors do not trigger that fallback. This cache
has its own 32-entry/16-MiB logical retention bound, in addition to the existing
128-entry/64-MiB full-grid cache. Live node snapshots can outlast eviction.
Fallback tolerance reports publish only after the final source/execution fence,
survive trimming, and are superseded by a full-grid report without double counting.

Cached-only reads expose scalar text, type and coordinates. Formula provenance,
formula cached-result properties, styles, merges and row/column metadata remain
unavailable; formula counts are omitted rather than reported as zero. Row and
column coordinates can still derive from occupied scalar cells.

Source freshness and execution checks bracket reads, including retained nodes
and iterator advancement. Mixed worksheet/chart eager grid selection maps the
catalog position exactly once. XLSB unified error messages and debug output hide
leaf-authored payloads while preserving explicit typed causes. Leaf worksheet
Debug likewise reports structural counts only.

## Capability publication

The XLSB golden lists 26 element entries and 75 explicit gaps. It advertises
only the cell properties the leaf can currently project, and only `name` for
defined names. Missing font/fill/border/protection/number-format projections and
extended elements have follow-ups. Write verbs retain their M2 gap.

The generated OfficeCLI delta includes XLSB against the XLSX vocabulary as a
parity target; this does not claim OfficeCLI accepts XLSB files.

## Verification

Offline verification passed:

- `cargo test -p litchi-xlsb --all-features`: 1,076 tests/doctests passed,
  12 ignored.
- `cargo test -p litchi-ops --no-default-features --features xlsb,test-support`:
  418 tests/doctests passed.
- The corresponding XLSX-only ops run: 468 tests/doctests passed after shared
  vocabulary extraction.
- XLSB leaf all-feature/all-target Clippy, plus XLSB-only and all-feature
  ops all-target Clippy, with warnings denied.
- Workspace library compilation and XLSB-only ops rustdoc with warnings denied.
  Conditional family documentation links were adjusted to work without the
  document/presentation features.
- Both XLSB glossary profiles: 10 banned names checked per profile.
- Capability golden checks and all 12 OfficeCLI delta-generator tests.

Regressions cover Unicode ambiguity and positional access, mixed chart tabs,
sparse iterator exhaustion, dense output budgets, deferred malformed sheets,
table-only reads, cached-only row enumeration, indexed scalar fallback,
corruption/cancellation/source changes, source publication rollback, tolerance
retention, and content-free diagnostics. Retained cancellation tests configure
execution on the package source limits; `OpenOptions::execution` currently
applies only to open-time work.

## Remaining work

This checkpoint does not establish performance acceptance, complete OfficeCLI
parity, write integration, Python bindings or the remaining format tiers.
Extended XLSB element/property gaps remain listed in the capability report.

## Commit and cleanup

The rewritten base `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains an ancestor
of HEAD. `git add -A` fails because the linked worktree's Git metadata is read
only; no commit can be created in this environment. Both required staged size
checks were silent, and the changed/untracked artifact audit found no prohibited
evidence or compiled executables.

Removed 34.5 GiB of generated Cargo output, the generated untracked lockfile,
Python bytecode, and 27 checkpoint scratch paths. No temporary worktree was
created. All delegated work has finished.

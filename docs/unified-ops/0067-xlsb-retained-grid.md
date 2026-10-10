# XLSB retained worksheet grids and checked sparse reads

This checkpoint implements the L-XLSB-R2 read foundation. The complete unified
operations goal remains active; the XLSB unified node/query/view adapter is the
next integration wave.

## Read and retention contract

`sheet::Grid` owns an immutable `Arc<Worksheet>` snapshot. Eager workbooks select
it through `grid(sheet::Selector)`; source worksheet handles acquire it through
`grid()`. Snapshots survive owner drop, eviction and later eager publication.
Grid acquisition reuses the owner's decoded worksheet cache. Legacy eager
`worksheet()` and source `materialize()` retain their owned return types by
cloning the retained snapshot.

`cell_at(At)` and `cells_in(Area)` validate zero-based coordinates before reading
the sparse cell map. Missing cells remain absent. Range traversal seeks across
occupied rows and selected columns without allocating dense cells. Grid and
iterator Debug output omit authored values. Shared A1 coordinate/range errors
now omit the supplied reference from both their value and diagnostic text.

Both leaf owners retain complete decoded worksheets in a weighted LRU, with
128 entries and 64 MiB logical retained weight as the current internal ceilings.
Oversized worksheets bypass retention. Eviction drops the owner's reference;
caller-held snapshots remain valid. Parsing runs outside the cache bookkeeping
mutex, and concurrent publication returns the already-retained canonical value.
Source reads check freshness and execution fences on cache hits as well as
misses. Compact source worksheet tolerance reports survive eviction separately,
so reparsing does not duplicate repairs in `read_tolerance_report()`. The
existing `tolerance_report()` remains workbook scoped.

The weight includes owned string/vector capacities, cell/formula/rich-string
payloads, worksheet metadata, and nested scenarios, selections, formatting,
sparklines, slicers, timelines and tolerance records. Sparse B-tree storage uses
a conservative node-slot allowance; the metric is a logical cache admission
weight, not an allocator or RSS measurement. Weight calculation is outside the
ordinary cache-hit path.

`litchi-ops` now borrows these grids for its source XLSB trait operations instead
of permanently retaining every materialized worksheet in its own cache. Legacy
one-based cell reads now reject zero and out-of-grid coordinates before
materialization; A1 arguments are checked and redacted on failure. Fallible
source selectors and iteration are fenced, and iterators fuse after an error.
The legacy trait’s infallible dimensions/count methods still cannot propagate
errors; their migration to a fallible boundary remains open.

## Verification

Offline verification passed:

- `cargo test -p litchi-xlsb --all-features`: 1,072 tests/doctests passed,
  12 ignored.
- `cargo test -p litchi-ops --no-default-features --features xlsb,test-support`:
  401 tests/doctests passed.
- `cargo test -p litchi-sheet`: 14 tests passed.
- The affected XLSX column-layout regression passed.
- XLSB all-feature/all-target and ops XLSB all-target Clippy, with warnings
  denied; XLSB all-feature rustdoc with warnings denied.
- Both XLSB glossary profiles: 10 banned names checked per profile.
- Workspace library compilation passed.

Coverage includes sparse bounds and missing cells, authored-input redaction,
concurrent canonical publication, byte-driven eviction, pinned snapshots,
oversized bypass, spare allocation capacity, source changes/cancellation,
tolerance retention across trim/reparse, late-fence publication rollback,
and eager edit publication preserving an old snapshot.

The XLSB-only lint run also exposed read-dispatch arguments used solely by the
signed wrapper when none of the current unified adapters is enabled. A narrowly
conditional lint allowance documents that interim feature configuration; it
does not suppress the check in DOCX/XLSX/PPTX builds.

## Remaining work

R4d still needs XLSB unified nodes, query/view dispatch and capability publication.
Deferred PivotTable/slicer/timeline cases still require the planned cached-value
fallback, with typed refusal where the leaf lacks a representation. This record
does not establish OfficeCLI parity, performance acceptance, complete policy
migration, write integration or the remaining format tiers.

## Commit and cleanup

The rewritten base remains an ancestor of HEAD. `git add -A` failed because
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` cannot be
created on the read-only filesystem; no commit was created. Both required
staged evidence checks were silent, and the changed/untracked file audit found
no prohibited evidence or compiled executables.

Removed 24.1 GiB of generated Cargo output, generated untracked lockfiles,
Python bytecode cache, and this checkpoint’s scratch logs. No temporary
worktree or fixture copy was created.

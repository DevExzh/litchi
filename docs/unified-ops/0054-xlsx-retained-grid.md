# 0054 — XLSX retained grids and bounded worksheet metadata reads

## Scope and provenance

Implement the leaf prerequisites L-XLSX-R1, R2 and R6 from record 0010:
retained grid reads, worksheet delegates, catalog names, metadata blocks and
early-stop traversal. XLSX File dispatch still reports a pending adapter.
Formatting/table access (R3/R4), the concrete adapter and paired performance
acceptance remain required; this record does not close M1 or the full goal.

Code-only portions of candidate `0586063e19` provided the starting point.
Candidate history and measurement artifacts were not merged. The current local
base tip `32ad64e3207655b4ba558fc2793d9a7396655cdf` remains an ancestor, so no
rebase against that tip is needed.

## Behavior and API policy

`SourceWorksheet::grid()` returns a cheap-to-clone, immutable
`workbook::source::Grid` pin. It exposes checked cell/range, row/column, merge,
default, extent and tolerance reads. A grid owns its parsed snapshot and can
outlive the workbook; it performs no later source I/O. Reads through the live
worksheet still enforce source-version and execution fences.

Worksheet handles and the workbook share a leaf-owned clean-value cache,
with logical ceilings of 64 MiB and 128 entries. Clean least-recently-used
entries are evicted first. Pinned grids remain valid after eviction; oversized
values are returned but not cached. Per-sheet build locks share successful
retained builds. Repair reports survive eviction. `semantic_cache_diagnostics`
returns focused `workbook::source::CacheDiagnostics`, and
`trim_semantic_cache` releases unpinned cache entries. No new root aliases are
added. Both XLSX glossary profiles reject the candidate's prefixed type names.

Cold cell/range reads preserve selected streaming. Explicit grid/structural
reads or parser fallback materialize the grid; later reads reuse it. This
differs from the candidate's automatic eager materialization below 64 MiB.
Ordinary read scans now admit inert worksheet children under parity checks;
edit-verifier target scans keep their narrower admission rules.

`visit_cells_until` accepts `ControlFlow`: resident-grid traversal stops at the
breaking cell, and final source/execution errors outrank the break value.
Cold selected scans still validate the complete worksheet before callbacks;
early break then saves callbacks, not the preceding scan. No constant-time
cold query claim is made.

Metadata delegates retain a projection outside `sheetData`. Its physical XML
scan counts **all** events, including elided cells, against the caller's
`SourceLimits`, and checks depth and attribute-size limits with typed OPC
limit errors. Attribute validation uses the shared checked iterator to avoid
quadratic duplicate checks. Downstream metadata parsers see only projected
block events. Large sheets need an explicitly raised physical event budget;
elision does not bypass it. Malformed or foreign roots, multiple documents,
malformed end tags and excessive parser depth preserve the original input for
authoritative parser refusals. MCE parser results are compared with the eager
owner.

The cache's weights are logical retention estimates, not measured heap/RSS or
execution-budget reservations. The existing parsed-store accounting policy is
unchanged; caller-selected semantic limits and execution charging remain a
separate required migration. Other existing source read routes still need the
full limits audit. Existing XLSX owner/transaction vocabulary also remains to
be migrated under ADR 0034.

## Verification

The initial all-feature library checkpoint passed 1,436 tests, including the
full-corpus cell-facts oracle. Review then strengthened physical XML budgets,
root/depth validation and MCE parity; final all-feature results follow below.

- Minimal XLSX source suite: 84 tests passed, including grid sharing, racing
  builds, eviction/oversized bypass, repair persistence, source/cancellation
  fences, eager/source metadata parity and early-stop behavior.
- Operations XLSX owner integration: 13 tests passed across retained/eager
  and real-fixture routes.
- All-target/all-feature XLSX Clippy with warnings denied passed. The initial
  attempt reported a missing existing drawing fixture; it was supplied from
  the base checkout for the successful run.
- All-feature workspace library check passed, including Python binding build.
- XLSX minimal and full glossary profiles each passed five banned names
  (10 profile/name checks total). The overall policy now registers 148 bans;
  unchanged non-XLSX profiles were verified in record 0053.
- Python glossary and generated-delta tests: 31 passed.

Final all-feature XLSX suite passed 1,442 library tests, 753 integration
tests and 6 doctests. The final all-feature source-only confirmation also passed
84 tests after the shared attribute-iterator change. Crate boundaries passed for 66 workspace packages, 267 internal dependency
declarations and 11 existing debt items. Changed-source formatting and
`git diff --check` passed. No paired
performance, retained-heap/RSS or Python runtime acceptance is claimed.

## Commit and cleanup

Git staging failed again at the read-only worktree metadata path
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`.
The old partial index must be refreshed before any commit. Records 0051–0054
remain uncommitted; no commit is claimed. Implementation continues, so this is
not an impasse for the full active goal.

No worktree is created. One existing drawing test's missing 12,470-byte fixture
is temporarily copied from the read-only base checkout for verification, then
removed with only directories created for that copy. Build output, generated
lockfiles, patch scratch and logs are removed after checks. Repository artifact
and staged-size restrictions are checked before the attempted commit.

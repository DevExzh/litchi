# 0206 — Unified XLS cell removal

`File::remove` and transactional removal now accept XLS cell paths, ranges,
and cell queries. A single native edit removes the selected stored records;
neighboring coordinates do not shift. Targets are deduplicated, missing cells
are exact no-ops, and semantic limits bound expansion, staging, operations,
and retained history. Generic batches retain rollback and item attribution.
Native patches support exact replay and inverse through the unified owner.

Formula cells remain an explicit refusal. Capabilities distinguish ordinary
cell removal from formula/dependency closure, worksheet, row and column
removal. Existing signature, encryption and source-freshness policies apply.
Removing the last rich or opaque shared-string identity still requires a
native semantic inverse that can retain that metadata.

The native cell owner now indexes packed blank cells. Removal splits interior
MulRk and MulBlank records, preserves neighboring XF styles, and accepts the
standalone records produced when successive removals reduce a packed family.
Semantic inverse can restore numbers after every packed member was removed;
exact inverse retains the original artifact. Native publication continues to
rebuild row blocks, INDEX/DBCELL, dimensions and shared-string dependencies.

## Source coverage correction

The previous write suite's common helper used byte opening for both `Eager`
and `Auto`; CFB byte opening uses eager ownership. Dedicated filesystem tests
were valid, but the broad source/eager coverage claim in record 0205 was too
broad. Both write and removal suites now use real filesystem `Auto` owners
and keep their temporary directories alive. Eager cases still open bytes.
The removal benchmark likewise constructs positional native and unified
owners from the same filesystem fixture.

## Verification

All-feature native XLS tests and doctests pass (1,783, one ignored), as do all-feature
facade tests and doctests (1,535). The corrected filesystem/eager
write suite has 21 tests and the removal suite has 15. Coverage includes
physical absence, neighboring coordinates/styles, ranges with holes, duplicate
queries, formula refusals, replay/inverse, batch rollback/attribution,
semantic limits, signatures, encryption and stale sources.

Warning-denied all-target Clippy and rustdoc pass for XLS and the facade.
The minimal facade build and XLS glossary audit (24 banned names, one
transaction family) pass. Scoped
formatting and the crate-boundary checker pass (66 packages, 268 internal
dependencies, 11 existing debt items). The capability golden and generated
OfficeCLI delta are refreshed; all 15 generator tests pass. Repository
lockfiles remain unchanged; offline commands use an external lockfile
selecting cached cc 1.6.0 and an external build target.

Reproduce the suites with `cargo test -p litchi-xls -p litchi-ops --all-features`.
The paired harness is `crates/litchi-ops/examples/xls_remove_bench.rs`; it
removes one middle cell in 1, 100, 1,000 and 10,000-cell workbooks, comparing
native edit/remove/commit with unified removal. Both eager and filesystem
positional owners use identical input bytes. Setup, opening, serialization
and verification are outside timing. Every remaining coordinate/value,
output bytes and source immutability are checked.

## Remaining scope

XLS add/move/swap, worksheet and axis removal, formula/style authoring,
security publication, source query streaming, transaction coalescing and
batch performance remain open. The all-format goal, Python/CLI and full
acceptance program remain active. The rewritten base is still an ancestor at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; no further rebase is needed.

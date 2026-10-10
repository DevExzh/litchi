# XLSB semantic tab catalog

This checkpoint implements the L-XLSB-R1 catalog prerequisite for the M1 XLSB
adapter. The complete unified operations goal remains active.

## API and ownership

Canonical semantic vocabulary lives in `litchi_xlsb::sheet`:
`Kind`, `State`, `Tab`, and `Selector`. Selection uses authored names or checked
zero-based `Position`, never native tab or relationship IDs. The eager workbook
exposes `sheets()` and `sheet(selector)` over a retained validated catalog; the
source owner exposes the same selector form and freshness-checked kind/state
facts on its existing handles. Legacy worksheet-only entry points keep their
worksheet ordinal meaning. New catalog selection follows the existing leaf's
case-insensitive workbook name comparison.

Eager catalog validation checks internal targets, exact recognized relationship
kinds, content types and duplicate targets before publishing the catalog. Tab,
selector and workbook Debug output omit authored names. Source reads check
freshness and cancellation around metadata lookup.

## Integration review

BIFF12 permits a very-hidden tab with a null relationship. The catalog must
retain that semantic tab without inventing a worksheet or physical target.
Both owners now retain it as `Kind::Unbound` with `State::VeryHidden`,
without a worksheet ordinal. Source materialization and indexed payload reads
refuse with typed `UnsupportedFeature`; visible/hidden null relationships
remain invalid. Tests cover current and beta `BrtBundleSh` layouts.

The full crate test build also identified three stale test accesses to the old
DrawingML run-property layout. Those tests now use the shared run properties
introduced in record 0060.

## Remaining M1 work

L-XLSB-R2 still needs leaf-owned weighted decoded-worksheet retention and checked
zero-based sparse range traversal. The ops bridge's per-sheet permanent cache
must be replaced by that leaf retention policy. Unified XLSB node/query/view
integration (R4d), capability publication, paired performance evidence and
OfficeCLI corpus acceptance remain open.

## Verification

Offline verification passed:

- `cargo test -p litchi-xlsb --all-features`: 1,055 tests/doctests passed,
  12 ignored.
- All-feature/all-target Clippy and rustdoc with warnings denied.
- XLSB minimal/full glossary profiles: 10 banned names each.
- Glossary checker Python tests: 21 passed.

- `cargo test -p litchi-ops --features xlsb,test-support`: 400 tests/doctests
  passed.

## Commit and cleanup

The rewritten base remains an ancestor of HEAD. `git add -A` still fails because
the linked-worktree Git index is read-only; no commit was created. Both required
staged artifact checks were silent. Generated Cargo output, untracked generated
lockfiles, Python caches and checkpoint scratch files were removed. No temporary
worktrees were created. The full goal remains active; the next XLSB prerequisite
is leaf-owned weighted worksheet retention.

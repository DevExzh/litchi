# 0217 — Unified legacy PowerPoint slide structure

Unified `remove`, `move_to`, and `swap` now dispatch to the retained native PPT
primary edit. The facade resolves slide paths and scoped removal selectors,
validates placement, and applies semantic target/staging/snapshot limits.
Native dependency validation and immutable publication remain authoritative.

Removal resolves the complete selection against the source snapshot and stages
positions in descending order. It removes slide-list entries and their owned
outline records; native incremental history can retain unreachable payloads.
This is ordinary document editing, not content sanitization.

Move positions are measured after detachment. Start/end, explicit positions,
before/after anchors, and placement into the presentation root share the common
placement contract. Swap resolves both original positions before one native
edit; adjacent swaps need one move and non-adjacent swaps need two. Self moves
and self swaps retain the exact original snapshot. Other element kinds remain
explicit unsupported targets.

Single calls, explicit edits, atomic batches, best-effort batches, and reversible
history use the same dispatch. Capabilities expose the supported verbs and
slide removal, retaining the released-source refusal and native closure gaps.
Slide insertion, shape insertion/removal, nested writes, rich text, consecutive
command coalescing and the full all-format/Python/CLI goal remain open.

## Verification

The existing all-feature facade suite passes 1,585 tests/doctests. The nine
new structural integration tests pass separately with all features and with
only `ppt`; PPT-only capabilities/read/write tests pass another 26 cases.
Twelve native primary-edit contract tests pass. Coverage includes both owner
routes, all sixteen ordered pairs of four-slide swaps, relative/root/start/end
placements, exact no-op history, descending query removal, removal of all
slides, speaker notes, native-byte equivalence, replay/inverse, batch rollback,
stale/signed/released refusals and expansion/staging/snapshot budgets.

All-feature/all-target Clippy and rustdoc pass with warnings denied, as does
the no-feature build. Crate boundaries pass for 66 packages / 268 declarations
with 11 pre-existing debt items. OfficeCLI delta freshness and its 15 unit tests
pass; legacy PPT is outside OfficeCLI's OOXML parity matrix.

## Paired structural probe

Implementation: `4f53162ac9`. [Raw samples and build identity](results/0217-ppt-structure.json)
record seven alternating native/unified pairs with ten operations per leg.
Paths are parsed before timing; input creation, open, warmup, exact output-byte
comparisons and source checks are also outside timing. Release LTO is disabled.
The native and unified outputs match exactly in all eighteen cases.

| Operation | Slides | Eager median ratio | Positional median ratio |
| --- | ---: | ---: | ---: |
| remove | 3 | 1.020 | 1.054 |
| remove | 10 | 1.026 | 1.042 |
| remove | 100 | 0.992 | 1.018 |
| move | 3 | 1.021 | 1.068 |
| move | 10 | 1.019 | 1.041 |
| move | 100 | 1.016 | 1.017 |
| swap | 3 | 1.017 | 1.066 |
| swap | 10 | 1.020 | 1.055 |
| swap | 100 | 1.010 | 1.018 |

All local median ratios are below the 1.10 single-write review trigger. This
probe does not establish frozen-harness qualification, broader corpus behavior,
allocation bounds or coalesced batch scaling. The earlier shape-text write
budget gaps from record 0216 remain open.


The base remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`, already an ancestor.
The repository lockfiles remain unchanged; verification uses an external
scratch lock with offline cached cc 1.6.0 instead of unavailable cc 1.7.0.

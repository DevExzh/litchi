# 0134 — PPTX package-root three-way merge

## Implementation

Native three-way planning no longer refuses every patch that changes
package-root relationships. Compatible relationship-ID edits merge automatically
with their retained metadata. Divergent IDs or opaque root XML changes produce
an immutable conflict whose physical resource is `/`, alongside any ordinary
part conflicts in deterministic order.

Root candidates retain their graph and metadata together. Base drops both root
candidates; Left or Right selects that candidate's graph and exact metadata.
Metadata-only spelling changes participate even when neither candidate has a
semantic root delta. Selection preserves automatic non-root resources and
metadata. Root and payload conflicts can be resolved independently. This uses
the existing physical-owner resolution contract; it does not add a semantic
conflict path or a last-writer-wins rule.

The plan admits each retained input under the intersected policy before cloning
it. Compatible merged relationship counts are checked before graph allocation.
Finish accounts for selected root graphs, metadata and part states before
cloning the final write set, including the exact combined wire byte limit.
An invalid resolution index is rejected before cloning the plan.

Five new regressions exercise disjoint root additions, deterministic branch
ordering, durable forward/inverse replay, same-ID and opaque-spelling conflicts,
mixed semantic/spelling candidates, all three resolutions, independent part
metadata, independently selected root/payload conflicts, merged count limits,
and exact/one-under wire budgets. The original plan stays immutable and cannot
finish with unresolved conflicts. Chosen results and their serialized inverses
restore exact publication bytes.

## Verification

The final native PPTX suite passed 983 tests with encryption enabled and one
existing ignore. All 28 focused patch tests passed. The PPTX/encryption facade
passed 533 tests/doctests across 26 result groups. Native all-target Clippy
passed with warnings denied. Documentation also passed with warnings denied,
and both PPTX glossary profiles passed (79 minimal and 70 full banned names).
Read-only review found no remaining checkpoint
correctness, atomicity or admission blocker.

Commands were offline and serialized:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib opened::patch::
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx --features encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

## Remaining scope

Part/content-type metadata-coupled divergent conflict resolution remains an
explicit refusal. The conservative self-closing-root and unchanged-content-type
proof restrictions recorded in 0133 also remain. Public slide insertion,
remaining native/facade coverage, performance acceptance and the complete
all-format goal are unfinished. No performance acceptance claim is made.

## Repository checkpoint

The rewritten base is still an ancestor of this branch, so no rebase is needed.
Staging failed because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created and no permission workaround was
attempted. Both required staged evidence guards printed nothing; the
changed/untracked artifact audit found no prohibited evidence or compiled
binaries. After all gate handles were terminal, checkpoint scratch logs, the
generated root lockfile and build output were removed. No temporary worktree
was created; inherited changes and intentional evidence files were preserved.

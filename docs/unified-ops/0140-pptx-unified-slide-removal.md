# 0140 — Unified PPTX slide removal

## Implementation

File, transaction and batch removal now delegate eager PPTX slide removal to the
native immutable edit. Paths and scoped queries resolve against one unchanged
snapshot. All matches must be slides; unsupported predicate diagnostics and empty
matches refuse. Unique positions are bounded before allocation and sorted in
descending order before one native transaction stages removal. A later native
refusal discards the complete candidate and leaves the caller unchanged.

Native removal owns relationship/content-type changes, lexical preservation and
inverse proofs. It currently requires a dependency-free slide with just its layout
relationship and refuses the final remaining slide. Layouts stay shared and
retained. Additional incoming or outgoing links, media/notes/comments and richer
topology are not silently discarded. Capabilities expose the eager operation and
explicit source/dependency/final-slide gaps. The facade introduces no XML mutation
or serialize/reopen fallback.

Owner, signature and retention checks happen before source promotion. Target and
staged action limits are admitted before the native editor; the transaction kernel
also accounts for retained proof bytes. Source-backed removal remains an explicit
capability refusal without promotion.

## Verification

Eight public regression tests cover path and scoped-union removal, descending
reindexing, single/multiple patch replay and inverse, saved/reopened results,
following text edits, zero/mixed/unsupported/bare selectors, exact and insufficient
staged limits, expanded target limits, final-slide refusal and atomic/best-effort
batches. The notes fixture also proves rollback after a higher-position slide was
already removed in the local native candidate. Native no-op checks verify that a
refused final-slide edit leaves its original package exact.

The tests exposed a facade error-mapping defect: native `SlideRemovalPlan`
refusals fell through to `CorruptFile`. Classification now retains the typed leaf
error and maps final-slide refusal to `InvalidArgument`, signature/protection
refusals to `DocumentProtected`, and other unsupported removal prerequisites to
`UnsupportedCapability`. A unit test covers all ten existing refusal families.
Another unit test proves source-owner refusal initializes neither promotion cache.

The PPTX/encryption facade suite passed 558 tests/doctests across 29 result groups.
The combined DOCX/XLSX/PPTX/encryption suite passed 1,074 across 29 groups.
Read-only review found no concrete ownership, atomicity or dependency blocker.
All-target facade Clippy and rustdoc passed with warnings denied. The DOCX-only
build passed, checking isolation of PPTX capability state. Both glossary profiles
passed (79 minimal and 70 full banned names), as did scoped formatting and
`git diff --check`.
The capability golden was reviewed: its removal refusal is now owner-specific,
with explicit source-owner, dependency-topology and final-slide contexts. The
regenerated OfficeCLI delta passed its deterministic check and all 14 harness
tests. Builds were offline and serialized:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo check -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-ops --no-default-features --features pptx,encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
python3 -m unittest tools.test_officecli_capability_delta
```

## Remaining scope

Richer slide-removal dependency closures and source-native structural editing
remain unfinished. The latter requires the topology/catalog/proof work in record
0137. Other element CRUD, remaining formats and performance acceptance remain
open; the full goal is active. No performance acceptance claim is made.

## Repository checkpoint

The rewritten base is an ancestor, so no rebase is necessary. No temporary
worktree was created. Inherited changes and evidence are preserved.

Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
Both required staged evidence guards printed nothing; the changed/untracked
artifact audit found no prohibited evidence or compiled binaries.

All verification handles finished before checkpoint scratch logs, the generated
root lockfile/build output and the two delta-harness bytecode files were removed.
Inherited fixtures, tracked lockfiles and intentional evidence were preserved.

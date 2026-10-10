# 0135 — PPTX part relationship metadata resolution

## Implementation

Three-way planning now resolves conflicting part relationship metadata with
the part's payload and graph, using the existing physical-owner Base/Left/Right
contract. It collects both semantic part conflicts and lexical relationship
owner conflicts, removes those owners from automatic metadata, and retains each
candidate's metadata alongside its optional part delta. Metadata-only conflicts
therefore work without inventing a payload delta.

Base retains the original owner; Left or Right selects that branch's complete
owner candidate. Unrelated owners remain automatic, including independent
content-type mappings and added parts. This replaces the broad refusal whenever
any non-root metadata accompanied a part conflict. Root and part choices can be
resolved independently in the same immutable plan.

Finish admits the combined selected metadata group count and exact wire length
before cloning the final metadata/write set. Each selected owner is disjoint
from automatic metadata because planning partitions its relationship groups
first. Existing native catalog validation and bounded source installation remain
in use. No new public conflict type or durable wire version is introduced.

Three regressions cover matrices of same-ID graph conflicts, opaque XML spelling
changes, mixed semantic/spelling branches, payload changes with one-sided graph
changes, all three resolutions, and independent content-type/part/root edits.
Combined root-and-part metadata selection exercises exact and one-byte-short
wire budgets. Resolved patches and inverses are serialized and decoded before
replay; assertions compare complete publication bytes, not only typed graphs.

## Verification

The final native PPTX suite passed 986 tests with encryption enabled and one
existing ignore, including all three new regressions. The focused suite passed
30 tests before the final combined-root/part budget regression was added; that
regression passed in the full run. The PPTX/encryption facade passed 533
tests/doctests across 26 result groups. Native all-target Clippy passed with
warnings denied. Documentation passed with warnings denied, and both PPTX
glossary profiles passed (79 minimal and 70 full banned names). Read-only review found no remaining checkpoint correctness,
atomicity or admission blocker.

Verification commands were offline and serialized:

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

A conflicting part whose candidate changes content type or existence still
requires dependency-aware content-type projection and is explicitly refused.
Divergent content-type metadata itself also remains an explicit refusal.
Selecting one physical part selects all that branch's edits to the owner; this
does not introduce per-relationship conflict choices within a divergent owner.
The conservative self-closing-root and unchanged-content-type proof restrictions
from 0133 remain. Public slide insertion, remaining native/facade coverage,
performance acceptance and the full all-format goal are unfinished. No
performance acceptance claim is made.

## Repository checkpoint

The rewritten base remains an ancestor of this branch, so no rebase is needed.
Staging failed because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created and no permission workaround was
attempted. Both required staged evidence guards printed nothing, and the
changed/untracked artifact audit found no prohibited evidence or compiled
binaries. After all verification handles were terminal, checkpoint scratch
logs, the generated root lockfile and build output were removed. No temporary
worktree was created; inherited changes and intentional evidence were preserved.

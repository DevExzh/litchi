# 0136 — Native PPTX slide insertion

## Implementation

`Edit::insert_slide(Position, layout_name)` inserts an empty slide linked to an
exact registered layout name, including into a fresh package or an empty slide
list. It returns the staged slide for subsequent shape edits in the same
transaction. Layout placeholders are not copied. Missing and ambiguous layout
names have typed errors; the facade maps a missing layout to `not_found`.

Insertion uses the native authoring codec, bounded OPC metadata plans and the
existing lexical presentation splice. Strict and Transitional packages retain
their dialect. The native durable patch retains source spelling and physical
metadata, supports serialized forward replay and inverse, and preserves unrelated
destination metadata. Failed insertion preserves earlier staged edits.

Part allocation reserves existing internal relationship targets, including
case-insensitive dangling targets. Generic patch application now checks its final
private graph before publication: a removed part cannot have a surviving incoming
edge, and a newly added part cannot activate an unrelated dangling edge. Incoming
edges to added parts must be declared new or changed by that owner's patch.
External relationships do not count as internal dependencies. Owner lookup honors
OPC case equivalence.

Insertion also shares removal's conservative presentation schema validation.
The new relationship ID must have no existing XML bindings; positional slideshow
ranges and unknown extension surfaces are refused. This prevents a dangling
custom-show reference from becoming live and avoids shifting unmodeled positional
dependencies. Presentation payload bytes are admitted against half the patch
budget before insertion scans. The existing removal contract still requires
exactly one selected slide-list binding.

New public regressions cover empty/populated lists, both dialects, fresh packages,
staged shape edits, exact-name ambiguity, rollback, patch budgets, signature
policy, lexical metadata overlays, forward dangling-reference activation and
inverse incoming references, root/external edges, and presentation dependencies.
Adversarial allocation and inverse tests initially exposed missing checks and
passed after the fixes. One existing forged-removal fixture now constructs its
malicious target directly because generic patch application correctly refuses to
construct that invalid graph; its original negative assertion remains.

## Verification

The encrypted native run passed 986 unit tests, 10 insertion tests and 5
signature-policy tests, with one existing unit-test ignore. Native doctests passed
20 tests with two existing ignores, including the new insertion example. The
PPTX/encryption facade passed 533 tests/doctests across 26 result groups. Native
all-target Clippy and rustdoc passed with warnings denied. Both glossary profiles
passed (79 minimal and 70 full banned names). Read-only review found no remaining
checkpoint blocker after the owner-equivalence, XML-binding and early-byte-budget
fixes. Quick-XML already enables closing-tag name checks by default; review
withdrew its initial concern about that configuration. All ten insertion
regressions passed again after the final byte-budget admission reorder.

Verification commands were offline and serialized:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib --test slide_insertion --test opened_signature_policy
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --test slide_insertion
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx --features encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

## Remaining scope

This is native slide insertion. Unified facade slide CRUD, content-type/topology
three-way conflict projection, remaining format coverage and performance
acceptance remain unfinished. Existing conservative metadata replay restrictions
from records 0133–0135 remain. No performance acceptance claim is made; the full
goal remains active.

## Repository checkpoint

The current base remains an ancestor of this branch; no rebase is needed.
Staging failed because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created and no permission workaround was
attempted. Both required staged evidence guards printed nothing, and the
changed/untracked artifact audit found no prohibited evidence or compiled binaries.
After all verification handles were terminal, checkpoint scratch logs, the
generated root lockfile and build output were removed. No temporary worktree was
created; inherited changes and intentional evidence were preserved.

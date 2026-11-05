# 0141 — Composable source-backed PPTX slide ordering

## Implementation

The native source-backed slide-order editor now composes repeated moves and
swaps against its currently staged order. Positional APIs take the common checked
`Position` type. The obsolete `move_position` alias and `NoMove`/`MultipleMoves`
refusals were removed; an empty edit commits as an exact no-op.

Every changed candidate is rebuilt from the original presentation XML's opaque
entry spans and the current binding permutation. Returning to the original
permutation restores the original payload, including lexical inter-entry trivia,
and publication takes the exact no-op ZIP path. The editor retains no chain of
intermediate XML buffers. Managed execution reserves the exact length-preserving
root output before rewriting; the reservation follows the candidate through
snapshot, commit and patch clones. Invalid later positions leave the prior staged state
unchanged; source freshness, execution, signing and positional-owner checks remain
in the native planning/publication paths.

This removes the one-move restriction that obstructed composing a source order
batch into a single proof. It is a native prerequisite, not a claim that the
primary source package or default facade now supports structural edits. Their
catalog/proof integration remains explicit work below. No eager materialization
or serialize/reopen fallback was added.

The facade's typed `SlideOrderPlan` classification now distinguishes modify
protection (`DocumentProtected`), ambiguous bindings (`CorruptFile`) and unsupported
positional-owner contexts (`UnsupportedCapability`). It retains the native error
and context instead of flattening them into a string.

## Verification

The native source-edit integration suite passed 35 tests. New coverage includes
repeated moves/swaps using current positions, unchanged payload cache diagnostics,
composed patch replay and inverse, cancellation with lexical trivia and exact ZIP
equality, empty signed commits, and failure atomicity after an earlier valid move.
Managed-context tests verify refusal before output allocation, release when an
edit returns to the original order, and continued charging through cloned patches
and inverse patches until the last holder drops. Existing tests retain bounds,
invalid owner/MCE, cancellation, stale/foreign
source, signing and partial-output checks. A facade unit test covers all eight
remaining order-refusal classifications.

The first full native run failed two SVG tests because corpus inputs were missing.
Three inputs required by those tests were temporarily copied from the base checkout for verification, without changing
the tests or base checkout, and removed after all runs:

| Fixture | Bytes | SHA-256 |
| --- | ---: | --- |
| `tdf163852.pptx` | 9579 | `2525665a1e430a9ee543faf456424bd7b994cce0a2d76470a5a07453e3e2ab39` |
| `tdf164622.pptx` | 18696 | `33b86388804554e35cb3b07e81cf5db5cb6bb02e76716c71f233b57325c2d199` |
| `tdf169496_hidden_graphic.pptx` | 35202 | `f717ca4657257b538466c772eca2a4d4886cb5b919c2ad5f044377fcee124fc8` |

Their base path is `/home/zhuhe/code/litchi/3rdparty/libreoffice-core/sd/qa/unit/data/pptx/`;
copy the listed files into the corresponding worktree directory to reproduce the
full native check. They are verification inputs, not new committed artifacts.
The final full PPTX/encryption suite passed 1,557 tests/doctests with three
existing ignores across 99 result groups. The PPTX/encryption facade suite passed
559 across 29 groups. The combined DOCX/XLSX/PPTX/encryption suite passed 1,075
across 29 groups. Read-only review found no concrete correctness blocker in
the composition, lexical cancellation, bounds or reservation ownership. The
standalone editor continues to admit changed signed candidates during staging and
refuse them at publication; primary integration must preserve its own signed-revision
policy rather than adopt a changed signed package.

Both native and facade all-target Clippy passed with warnings denied, as did
rustdoc for both crates. The glossary profiles passed (79 minimal and 70 full
banned names). Scoped formatting and `git diff --check` passed.

Checks were serialized and offline:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

## Primary source integration remaining

Read-only architecture review identified an order-only path that does not require
a topology plan. Add a logical catalog keyed by physical URI, slide ID and
presentation relationship ID, alongside the retained root payload. Keep existing
slide snapshots bound to physical identity, and translate logical positions before
text snapshotting, change filtering and proof construction. Raw source positions
alone are insufficient once insertion/removal is added.

Primary `same_content`, no-op detection, inverse and retained-byte accounting must
include the root/order state. Keep the immutable source/topology digest and add a
current-order digest. A text-only patch authored after a reorder must require that
order context; applying it to the original order must not silently lose the reorder.
Bump durable wire schema to carry stable bindings, ordered keys and optional exact
root before/after deltas, and validate all root/slide preimages before publication.

For pure reorder, publish root and slide replacements in one bounded shared-overlay
write. Reuse the native span permutation via a private lowering helper; do not call
the standalone publisher separately from primary slide-overlay publication.
Retain managed root output reservations through the candidate and proof. Signed
no-ops must stay exact, and changed signed revisions need explicit policy.

The facade must route all position-sensitive reads through the revised catalog:
slide count/catalog, text, scenes, facts, notes, comments and modern comments.
`adopt_source` must retain the same source freshness and execution fences. Add/remove
still require the candidate topology work in record 0137. Default structural
capabilities stay explicit gaps until this integration is implemented and verified.
Other format/element coverage and paired performance acceptance remain open.
No performance acceptance claim is made; the full goal stays active.

## Repository checkpoint

The rewritten base remains an ancestor, so no rebase is necessary. No temporary
worktree was created. Inherited changes and evidence are preserved.

Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the read-only
filesystem. No commit was created and no permission workaround was attempted.
Both required staged evidence guards printed nothing; the changed/untracked
artifact audit found no prohibited evidence or compiled binaries.

All verification handles finished before the three temporary corpus copies,
checkpoint scratch logs and generated root lockfile/build output were removed.
Inherited fixtures, tracked lockfiles and intentional evidence were preserved.

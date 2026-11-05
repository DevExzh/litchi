# 0122 — PPTX committed source revision readback

The facade can now adopt a native source shape-text revision without saving,
reopening, or replacing the admitted source catalog. This is an internal
prerequisite for unified PPTX mutation dispatch: revised text, shape scenes,
paragraphs and runs must be visible to reads before that dispatch is exposed.

Native `pptx::source::Slide` adds `text`, `text_with_tolerance`, and
`text_and_name`. These use the existing native `SlidePart` readers, preserving
flattened text, producer name fallback, markup-compatibility handling and
repair reporting. Original payloads retain their source-backed part owner;
edited payloads are shared for the duration of the read while the snapshot
keeps their managed reservation alive. Physical XML limits admit the immutable snapshot, and source/execution
fences surround each read. Text reads reuse that physical admission instead of
scanning the same immutable XML again. Repairs are appended only after successful parsing
and the final fence. Native parser/output bounds remain in force; this does not
claim that returned `String` allocations have gained managed reservation tokens.

The facade retains an optional native revision beside its original source
projection. Text, selected slide, slide-list, shape-scene and tolerance reads
use the revision. Subsequent native edit promotion also starts from it, so a
second commit composes with the first. Catalog identities, layouts, notes and
other non-text facts keep their original projection because this adapter accepts
only source shape-text commits, whose native proof preserves those resources.
The internal adoption method requires a committed revision from the same owner;
it is not an arbitrary public package-replacement API.

No facade write capability is advertised by this change. Unified setters,
transaction history, source/eager publication and broader PPTX writes remain
required. The full goal, including other format tiers, bindings and performance
acceptance, remains active.

## Verification

Twenty-six native primary-source tests and 519 facade tests/doctests passed.
The readback regressions compare empty, escaped and multiline edited text
against reopened native slides; check valid MCE, combined text/name semantics,
untouched-slide laziness, source changes, cancellation and unchanged repair
reports on failure; and compare facade text, slide, shape, paragraph and run
reads against reopening. A second adopted edit retains the first, while the
original remains byte-exact. Catalog, metadata, size and tolerance parity are
also checked.

Warning-denied native/facade all-target Clippy and 17 native doctests passed
(two existing doctest ignores). Read-only review found no blocker under the
private same-owner, shape-text-commit adoption contract. Scoped formatting,
`git diff --check`, staged artifact guards and changed/untracked artifact audit
passed. The rewritten base remains an ancestor of HEAD; no rebase replay is
needed. The PPTX glossary profile passed with 68 banned names and zero
transaction families checked.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib presentation::source::tests::primary
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
```

## Commit and cleanup

Staging failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only shared Git metadata filesystem. No commit was created; changes remain
in this worktree.

All build/check processes completed before removal of this checkpoint's
`/tmp/litchi-unified-0122-*` files, generated root `Cargo.lock`, build target and
Python cache. Tracked tool lockfiles and inherited work remain untouched.
Review was read-only; no temporary worktree was created. The full goal remains
active.

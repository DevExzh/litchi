# Primary source-backed PPTX slide removal

The immutable source-backed primary owner now supports `Edit::remove_slide`
for the dependency-free profile admitted by the native removal planner.
Removal composes with shape text and slide ordering, publishes through a lazy
OPC prepared topology, and leaves the original owner unchanged. Readback uses
the effective stable-key catalog, including slides restored by inverse replay
when their Parts no longer exist in the reopened physical source.

## Structural revisions and durable replay

Each structural revision retains a prepared candidate and a declarative recipe
relative to the physical source. Later edits rebuild that recipe without
retaining the entire chain of previous candidates. Removing a restored slide
normalizes redundant resource changes; returning to the physical state restores
the original owner. Physical `source_position` remains unavailable for restored
Parts; callers read the effective slide through `Presentation::slide`.

The new bounded `LPSX0003` patch stores distinct before/after catalogs, exact
root and metadata images, canonical typed-topology digests, and changed slide
resources. Existing v2 text/permutation semantics are unchanged. Reopened
inverse replay restores typed relationships before applying guarded exact XML
images. Regression tests compare the complete decoded ZIP member map, including
relationship members and content types. This is not a claim that a rewritten
ZIP has identical container bytes; a local inverse normalized to the original
physical owner does retain its original bytes.

Admission checks ownership, layout registration, native slide/root policies,
source freshness, cancellation, finite resource limits, exact metadata images,
and typed topology. Root proofs preserve retained entry bytes and whole trivia
gaps; same-presence slide changes retain the existing shape-text proof. Wire
decoding preflights lengths and counts, rejects duplicate identities and
inconsistent presence changes, and checks its checksum before materialization.
Payload copies and retained recipes have managed reservations. Repeated edit
tests cover bounded retention and release after the last handle is dropped.

## Explicit limits and next work

The final slide, hidden or multiply referenced slides, unmodeled dependencies,
and unsupported policy/metadata surfaces remain refused. Primary durable
removal additionally requires canonical owner-relative target spelling for
the presentation binding and slide layout edge. OPC's current exact metadata
restore cannot accept a semantically equivalent alternate spelling after typed
relationship reconstruction. Such inputs are refused before changing state;
broader admission requires extending that guarded OPC behavior first.

The unified facade's source-removal capability remains gated. Its catalog and
related reads still use physical source positions and must consume the effective
native catalog before removal can be exposed through `File`. The OfficeCLI
capability delta therefore does not change in this increment. Source structural
creation, broader format CRUD, Python/CLI closure, and paired performance
acceptance remain open. No performance improvement is claimed.

## Verification

The final native PPTX suite passed 1,614 tests and doctests, including 17 focused
primary-removal integration tests. Coverage includes drop rollback, repeated
removal, text/order composition, local replay and inverse, durable forward and
reopened inverse, subsequent edits/removal of restored Parts, final-slide and
hidden-slide refusals, stale sources, cancellation, managed retention, and
resealed malformed patch rejection.

The combined DOCX/XLSX/PPTX/encryption facade passed 1,078 tests and
doctests. A final codec-only run passed five tests after replacing an internal
linear key scan with hash lookup. Native and PPTX facade all-target Clippy and
rustdoc passed with warnings denied. Both PPTX glossary profiles passed
(79 minimal and 70 full banned names checked). Scoped rustfmt, `git diff --check`, the
OfficeCLI delta check and its 14 generator tests passed. The changed/untracked
artifact audit found no prohibited evidence or compiled binaries. Crate
boundaries passed for 66 workspace packages and 268 internal dependency
declarations, with 11 explicit migration-debt items.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,pptx,encryption
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib topology_wire
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo doc -p litchi-pptx -p litchi-ops --no-default-features --features litchi-pptx/encryption,litchi-ops/pptx,litchi-ops/encryption --no-deps
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_crate_boundaries.py
python3 tools/officecli_capability_delta.py --check --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI
python3 -m unittest discover -s tools -p test_officecli_capability_delta.py
```

The rewritten `feat/office-format-completeness` base remains an ancestor of
this branch, so no rebase was needed. The full goal remains active.

Both required staged guards printed nothing. The scoped staging attempt failed:

```text
fatal: Unable to create '/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock': Read-only file system
```

No commit was created. No permission workaround was attempted; implementation
and progress changes remain in the working tree.

All verification handles completed before cleanup. The three temporary corpus
fixtures were removed after checking their recorded hashes. The generated root
lockfile, build output, this turn's Python caches and scratch logs were removed.
No temporary worktree was created. Inherited changes, tracked lockfiles, the
base checkout, and existing worktrees were preserved.

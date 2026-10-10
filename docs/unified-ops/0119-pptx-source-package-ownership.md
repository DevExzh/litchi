# 0119 — PPTX source package ownership

Checkpoint 0118 established eager package-carrying commits. This checkpoint
adds the primary immutable deferred-source owner at `pptx::source::Package`.
It uses the existing cheap-clone OPC source rather than serializing, reopening,
or materializing the package at the commit boundary.

## Contract and retained state

The new source family exposes `Package`, `Presentation`, `Slide`, `Edit`,
`Commit`, `Patch`, and `Diagnostics`. Package owns `edit`, immutable `apply`,
`write_to`, `to_bytes`, atomic `save`, and explicit `save_with_durability`.
Commit carries package/patch/diagnostics, with `into_package` and
`into_parts -> (Package, Patch)`. Introspection is `diagnostics().changed()`
and `patch.is_noop()`; patches have `inverse()`.

The immutable package retains the admitted native source and a bounded slide
XML overlay set. The primary edit composes multiple atomic shape-text batches,
including repeated changes to the same slide. All selected before-states are
validated before replay constructs its candidate. Replay preserves unrelated
destination edits. Inverse replay removes physical-source-equal overlays, so
restoring the initial revision reproduces the original complete archive.

`from_presentation` shares a retained read projection even while clones or
slide handles are live. It retains tolerance evidence and reuses the existing
strict edit-admission seed. Existing read handles still observe their original
revision. The owned primary presentation/slide projection reads its package's
revision and obtains detached native owned shape scenes.

Package operations and retained slide reads check caller execution policy and
physical source freshness. Changed signed packages refuse while staging;
no-op commits and replay preserve signed archives exactly. Writes reuse the
native source overlay publisher and raw-copy unselected members. Atomic saves
use the guarded sibling-temporary publisher, including source and execution
checks after file synchronization and before rename.

## Bounds and ownership

At most 32 physically changed slides may be retained across successive commits,
not just within one edit. Patches also bound changed-slide count; retained XML
payloads are checked against the native primary patch-byte ceiling. Each native
shape-text batch retains its existing selector, XML, part-size and text limits.

On managed sources, conservative input/output/escaping headroom is reserved
before lowering. The reservation stays attached to candidate payloads through
package clones, patches, inverse patches, detached scenes, and shape handles;
failed stages and final owner drops release it. Original parts retain their
managed `PartData` handles. Primary reads admit valid MCE independently of the
edit-only validator, and physical depth/event/attribute limits apply before MCE
can discard branches. A separate scene reservation follows processed XML and
its index when the original source payload can be released.

Managed scene admission currently reserves `64 * input_bytes + 1024` bytes,
capped at 128 MiB, and checks the actual retained scene size before returning
it. This conservative bound can refuse managed slides above roughly 2 MiB and
charges additional headroom when original `PartData` is also retained. It is a
known admission/performance tradeoff to tighten before broad managed large-slide
acceptance; unmanaged reads retain their native finite parser limits. This is bounded payload admission, not a claim that
all allocations throughout the older source families have been migrated.

## Verification

Twelve focused tests cover immutable source preservation, exact no-op/inverse,
composed edits, stale and foreign replay, unrelated destination edits, live
read handles, untouched-slide laziness, signed policy, source changes,
cancellation, atomic save/reopen, failed-stage atomicity, reservation lifetime,
budget recovery, the cumulative 32-slide bound, valid MCE reads with retained
processed-scene charges, and physical limits in discarded MCE branches. The final encrypted source
suite passed 92 tests and the owned-scene suite passed eight. Warning-denied
native all-target Clippy passed. Native doctests passed 17 examples (including
the new source ownership example), with two existing ignores. The facade
PPTX/encryption suite passed 516 tests/doctests. The PPTX glossary profile passed
with 68 banned names and zero transaction families checked. `git diff --check`,
the required staged artifact guards, and the changed/untracked artifact audit
passed. The rewritten base is still an ancestor of HEAD; no rebase replay is
needed.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib presentation::source
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib shape::owned
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx --features encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
```

## Remaining goal scope

These source patches are process-local exact-source patches. Durable encoding,
full source-owner vocabulary migration, structural verbs, new-package authoring,
and facade PPTX write dispatch remain required. Existing long-named source
slide/SVG/order families are not claimed conformant by this new primary family.
The unified capability registry is unchanged because no facade write verb is
enabled here. All later format tiers, bindings, performance, and acceptance
requirements remain in the active full goal; these scoped checks do not prove
full-workspace acceptance or performance parity.

## Commit and cleanup

The required staging attempt failed with exit 128 because the shared Git
metadata filesystem is read-only: Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit was
created; the changes remain in this worktree. No escalation or alternative
repository was attempted.

Builds and verification processes completed before removal of this checkpoint's
`/tmp/litchi-unified-0119-*` files, generated root `Cargo.lock`, build target,
and Python cache. Tracked tool lockfiles and inherited work are preserved.
Review agents were read-only; no temporary worktree was created. The full goal
remains active.

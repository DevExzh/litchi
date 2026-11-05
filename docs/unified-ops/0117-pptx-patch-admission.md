# 0117 — PPTX patch admission and visible payload ownership

Checkpoint 0116 made verified progress on the primary edit vocabulary. While
preparing the package-carrying commit migration, inspection found that its
underlying patch capture trusted `Part::blob_arc()` even when it did not back
`Part::blob()`. The package-equivalence shortcut also accepted matching shared
owners without proving equality of visible bytes. This checkpoint fixes those
correctness prerequisites rather than publishing a new owner on top of them.

## Change

Patch capture first constructs a bounded plan containing borrowed references
only for changed resources. It freezes each visible payload slice, content type,
and relationship collection, then measures the entire encoded write set before
retaining payload owners or copying resource strings. This avoids re-reading a
custom part's visible state after admission. The plan also checks changed-part
counts, relationship counts, wire-format `u32` fields, and checked aggregate byte
arithmetic. Unchanged large parts are not charged to the patch budget.

Ownership capture accepts a shared payload only when both its length and slice
pointer match the admitted visible bytes. Otherwise it fallibly reserves and
copies those visible bytes. Ordinary aliasing parts keep their existing shared
allocation. A custom part's unrelated shared bytes cannot hide a change or enter
the resulting patch.

Package equivalence, stale checks, and publication readback compare visible
values. The validation checks no longer build owned resource snapshots just to
compare them. Root and part relationship counts are also revalidated by the
common patch constructor after join/three-way policy intersection, so a result
cannot exceed its own decoder's relationship-count policy.

The `LPTX0002` encoding is unchanged. No public API was added or renamed in this
checkpoint. Package-carrying commits, diagnostics, immutable patch apply, and
unified PPTX shape-text writes remain required.

## Verification

Ten new regressions cover foreign shared owners, exact no-ops, validation without requesting shared payload owners, aggregate byte and changed-part boundaries, unchanged
large resources, root/add/remove accounting, relationship-count admission and
joined policy intersection, original shared-owner retention, visible prefixes
and switched accessors, real presentation publication/inverse replay, and wire
width boundaries. The initial default-feature opened suite passed 203 tests
with one existing ignored transcript generator; the three signed-source and
marked-XML integration suites passed 22 tests. The final encrypted opened suite passed 203 tests with the same ignored
transcript generator, including the joined-limit regression. All 516
PPTX/encryption facade tests and doctests passed. Warning-denied all-target
Clippy passed for native PPTX and the PPTX/encryption facade. Clippy initially
identified one needless borrow in a test fixture helper; it was corrected before
the final gate. Scoped formatting and `git diff --check` passed. The audit of
490 changed/untracked paths found no oversized evidence or compiled binaries;
both mandatory staged artifact checks printed nothing.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --lib opened
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib opened
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --test opened_signature_policy --test package_signature_policy --test byte_order_marked_parts
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
```

## Next ownership step

The reviewed design is to retain a small native owner context alongside the
opened snapshot: tolerance report, physical-source provenance, encryption mode,
and optional font state. A commit must move its already-validated OPC candidate
into the returned package and project digest memos onto that candidate; it must
not serialize/reopen or silently merge a committed result onto another caller's
graph. The migration must update consumers and remove superseded APIs in the
same wave, not introduce compatibility aliases. Snapshot-derived edits must
preserve that owner context. Source-backed publication remains a separate native
owner path with its freshness and execution fences.

This is scoped correctness evidence, not full-workspace acceptance or a paired
performance result. All later format tiers, bindings, CLI, and acceptance gates
remain in the active goal. The rewritten base remains an ancestor of HEAD, so
no rebase replay is needed.

Staging failed because Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only filesystem. No commit was created. Removed this checkpoint's generated
root Cargo.lock, build directory, Python cache, and temporary logs. No temporary
worktree was created; both read-only reviewers finished.

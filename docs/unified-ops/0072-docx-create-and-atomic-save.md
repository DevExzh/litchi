# 0072 — DOCX creation and atomic path saving

## Scope

This checkpoint adds unified `create`/`create_with` and
`File::save`/`save_as`/`save_with` for DOCX. It closes the basic DOCX
create/read/edit/publish loop using the existing leaf owners. Other format
creators and path publishers, batch execution, the remaining DOCX structural
families, Python, and performance acceptance remain open.

## Creation

`CreateOptions` selects document/template flavor, OOXML conformance, and
optional initial metadata. The default is a Transitional document. DOCX
creation adopts `Package::new()` or `new_template()` directly into the eager
unified owner. This follows ADR 0034's typed-owner rule and supersedes record
0015's proposed serialize/reopen creator route.

Initial core metadata is applied through the leaf's typed properties. Strict
conformance and metadata fields not representable by that part are refused
explicitly. The metadata admission match is exhaustive, so adding a field
cannot silently drop it. Creation reads no clock; absent created/modified
values remain absent, and identical creation requests publish identical bytes.
Created handles have no filesystem origin.

## Origin and immutable publication

The common `Opened<S>` handle carries optional shared path metadata. DOCX
direct, eager, source-backed, signed, unlocked, and prepared path routes
retain it. Relative paths are anchored to the open-time working directory;
symlink spelling is retained so save applies the normal symlink refusal.
Immutable DOCX commits and signature transitions preserve the origin.
Byte-created handles have none. The origin is omitted from diagnostic output.

`save()` publishes to that origin or returns `file_required`. `save_as(path)`
does not change it. `Saved` reports the destination and accepted byte count;
its Debug output does not expose the authored path. A recognized extension
from another supported office format is refused; a filename never converts
the content or promotes its flavor.

`SaveOptions` defaults to `Durability::Full` and atomic replacement.
`Overwrite::Refuse` is enforced at publication, not merely by an existence
precheck. The unified layer implements no rename or synchronization sequence:
it dispatches to OPC's sibling-temporary publisher. Untouched source output
retains its exact bytes; edited owners publish directly without ZIP reopen.
Source and cancellation fences run before replacement, with the retained
execution context governing staging. A failure before replacement leaves the
destination untouched. A directory-sync failure after replacement retains the
leaf's typed `Committed` error.

## Leaf additions and security

OPC adds guarded and unguarded no-clobber variants of its atomic publisher,
using `NamedTempFile::persist_noclobber`. They share the existing permission,
staging, cleanup, durability, and committed-error paths. Tests simulate a
destination appearing between staging and publication, rather than relying
on timing. `AlreadyExists` maps to `file_exists` without losing its typed
I/O/OPC source.

The DOCX source owner exposes its retained execution context without borrowing
the physical OPC graph. This is necessary after an immutable edit, because
the effective owner deliberately refuses the physical-graph escape hatch.

Signed no-op saves preserve signatures. Explicit unsigning publishes removal
even without a subsequent content edit. The successful unsigned package is
memoized, guarded by source checks, and cloned for each publisher; failures
are not cached. Source conversion preserves the owning package directly,
including effective edits, instead of serializing and reopening it.

Unlocked encrypted DOCX refuses ordinary publication. An explicit password
can re-encrypt through the leaf's retained encryption profile with full
durability; weaker durability is refused. Re-encryption of a plain input is
invalid. Only ciphertext is staged in a temporary file. The broader explicit
plaintext/security vocabulary migration is still open.

## Capabilities and verification

`Verb::Create` exposes creation discovery. Saving is publication rather than
an edit, so signed and unretained DOCX handles can advertise it. Unsupported
creator options and publication combinations remain explicit gaps. Format
goldens and the generated OfficeCLI delta are updated.

Tests cover creation and edits through publication/reopen, deterministic
metadata, origin retention, prepared relative paths, source/eager owners,
repeat saves and immutable clones, signatures and explicit removal,
re-encryption, overwrite policy, source changes, symlinks, directories,
missing parents, and cross-format suffix refusal. These are correctness
checks, not evidence for the goal's performance budgets.

Validation completed:

- Unified ops, all features: 970 tests/doctests passed.
- Unified ops, DOCX only: 596 passed; minimal features: 389 passed.
- OPC: 1,097 passed, one ignored, including 25 focused atomic tests.
- Thirteen save integration tests include encrypted origin publication;
  six creator tests include the full creation/edit/publication loop.
- All-feature and minimal ops all-target Clippy passed with warnings denied;
  OPC library/test Clippy passed.
- Ops, DOCX and OPC rustdoc passed with warnings denied.
- Workspace library check (including Python bindings), DOCX facade check,
  scoped formatting and `git diff --check` passed.
- Crate boundary checker and DOCX glossary gate (40 banned names) passed.
- OfficeCLI delta: 14 generator tests and regeneration check passed.

The DOCX golden has separate plain and encryption-enabled baselines, since
the encryption feature adds an option-policy gap. The delta generator merges
that conditional operation gap without changing the plain baseline; its
`document_protected` code matches the live publisher. Both feature matrices
are checked. Review also verified the prepared-origin, unsigned-cache and
pre-replacement source-fence changes.

Representative regeneration commands:

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --all-features
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features
CARGO_NET_OFFLINE=true cargo test -p litchi-opc
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --all-features --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo check --workspace --lib
python3 -m unittest tools.test_officecli_capability_delta
python3 tools/officecli_capability_delta.py --officecli /home/zhuhe/code/litchi/3rdparty/OfficeCLI --check
```

## Commit and cleanup

The rewritten local base remains `32ad64e3207655b4ba558fc2793d9a7396655cdf`,
already an ancestor of this branch, so no rebase is needed for that revision.
The required staged evidence/binary checks printed nothing. The artifact
audit found no prohibited evidence or compiled executable among the changed
and untracked files.

`git add -A` still fails because
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` cannot be
created on the read-only filesystem. No commit was created; the accumulated
implementation remains in the checkout. The full goal remains active.

Final cleanup removed 66,015 build files (71.0 GiB) and 21 scratch/cache
paths, including generated untracked lockfiles. This checkpoint created no
temporary worktrees or external fixture links. The final artifact audit
covered 320 changed/untracked paths and found no prohibited artifacts.

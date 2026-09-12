# 0524 temporary-file substitution test review

This is an independent, bounded review of the test-only adjustment to
`crates/litchi-cfb/src/writer/sequential.rs`. I did not edit production or test
source and did not run a build, test, profile, capture, or cleanup command.
The reviewed change is the addition of a displaced sibling path and the
identity assertion in
[`detected_temp_substitution_is_not_deleted_by_cleanup`](../../../../crates/litchi-cfb/src/writer/sequential.rs#L2089).

## Applicable contract

The relevant accepted decisions were read before this review:

- [ADR 0005](../../../adr/0005-io-memory-and-performance.md) requires a
  sibling temporary artifact, validation before publication, atomic filesystem
  replacement, an untouched destination on cancellation or failed staging,
  and best-effort identity-aware cleanup of an unpublished temporary artifact.
- [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md)
  keeps validation non-mutating and requires unsafe publication or validation
  failures to stop before output is published.
- [ADR 0008](../../../adr/0008-migration-and-verification.md) requires
  source-bound, reproducible verification and treats focused test output as
  evidence only for the behavior it actually exercises.
- [ADR 0001](../../../adr/0001-priorities-and-api-layers.md), [ADR
  0002](../../../adr/0002-crate-topology.md), [ADR
  0024](../../../adr/0024-current-topology.md), and [ADR
  0026](../../../adr/0026-ole-directory-metadata-binding.md) keep this
  behavior inside the private CFB writer/container owner and do not authorize
  a new public archive or filesystem abstraction.

These constraints are consistent with the 0524 manifest at
[`adr-manifest.json`](adr-manifest.json). OLE2/OOXML remains the active
performance priority; ODF remains deferred.

## Failure mechanism and corrected sequence

The failed candidate test removed the staged path and immediately created an
attacker file at the same name. On a disk-backed temporary directory, unlinking
the original made native inode reuse possible. The cleanup guard compares the
path identity with the identity captured from the original open file; if the
filesystem reuses that identity, cleanup can remove the attacker file. The
retained [candidate test log](candidate/check-cfb-tests.stdout) records the resulting
`writer::sequential::tests::detected_temp_substitution_is_not_deleted_by_cleanup`
failure as `NotFound` while reading the attacker path after the save returned,
with 277 tests passed and one failed.

The adjusted test now exercises the intended known-substitution case in this
order:

1. `test_destination("displaced-temp")` creates a second unique name in the
   same temporary directory as the staged file.
2. The replacement hook renames the original staged file to that displaced
   name. `rename` keeps the original file object and its native identity alive.
3. The hook writes a new attacker file at the original staged name and asserts
   that the staged and displaced paths have different `path_identity` values.
4. The hook returns an injected replacement error. The cleanup guard observes
   the new identity at the staged name, refuses to remove it, and leaves the
   attacker file present.
5. The test checks the attacker contents and the unchanged destination, then
   removes the attacker, displaced original, and destination paths.

This ordering matches the writer implementation: the original identity is
captured from the created `File`, candidate validation uses cloned handles, the
path check runs before the injected replacement hook, and the guard performs
identity-aware cleanup after the hook returns an error. The test therefore
checks the guard's known-substitution behavior without claiming protection
against a concurrent race between a path identity check and a later filesystem
operation. The existing save documentation correctly requires a trusted,
private parent for that stronger property.

## Portability and cleanup audit

Both `destination` and `displaced` are produced by `test_destination`, so they
use the same `std::env::temp_dir()` parent. This keeps `fs::rename` on one
filesystem and avoids a cross-device rename failure. The generated names
include the process ID and a process-local atomic counter, so the test is safe
under the normal parallel Rust test harness and does not collide with the
other sequential-writer cases. The writer drops its staging and validation
handles before invoking the replacement hook, which also satisfies Windows'
share/delete rules.

The passing path explicitly removes all three names. The panic path can leave
test files behind, as the neighboring sequential-writer tests can; this does
not alter production cleanup behavior and process-specific names prevent a
normal later process from treating those files as its own. A crash-clean test
would need a test-only RAII cleanup guard, but that is outside this bounded
regression and is not required to establish the writer contract.

On Unix, `path_identity` compares device and inode from `symlink_metadata`; on
Windows it compares volume serial and file index. Other targets use the
documented length fallback. The assertion uses the same private identity
function as production cleanup and the attacker payload is a different length
from the generated CFB artifact, so the fallback also observes a mismatch.
The test uses only regular files in a writable temporary directory and does
not follow a replacement symlink.

The identity assertion is sufficient for this setup because the original file
is kept alive by `rename`, while the new staged path is created independently.
Capturing the original identity before the rename would make the evidence more
explicit, but it is not needed to distinguish the two path objects or to prove
that the attacker file survives cleanup. The `fs::read` after `save_with_hooks`
is the stronger observable assertion: it fails if cleanup incorrectly deletes
the replacement.

## Disposition

Accept the test adjustment as a valid portability and semantic repair. It is
test-only, preserves the production best-effort security boundary, keeps the
destination assertion, and closes the inode-reuse false failure under the
owned disk-backed `TMPDIR` used by the quality gates. The retained [final
all-features CFB test log](final/check-cfb-tests.stdout) reports this test
passing with exit code zero; that log is the execution evidence, while this
review is the independent source-level assessment. No performance,
native-producer, physical-provider,
fuzz, scaling, or broader cleanup guarantee is inferred from this test.

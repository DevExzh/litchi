# 0362utc — ODG active-state validation and payload-effect proof

This successor follows 0362utb without amending it.  The implementation is
`ba582aac9a`, whose parent is `bac9d58fd4`.

## Cost and active-state validation

Replacement-containing plans now defer package-dependent page-name validation
until the execution segments have opened the active replacement.  Pure
semantic plans retain admission-time final-name validation.  The replacement
path therefore opens each replacement once: `apply_pending_joined` performs
the policy check, opens the replacement, and lets
`commit_replacement_operations` validate page names and commit that segment;
admission does not reopen the blob.  The new-page rename regression is a
focused proof of this call path because its `page/2` selector is absent from
the source and exists only after replacement.

Projected style dependency walking is likewise skipped for replacement plans.
The actual-payload package/semantic guard is checked first, and effect
declarations are validated before it, so cross-sub-edit package dependencies
cannot be hidden by forged declared write sets.  Remaining replacement style
operations are intentionally ordered within one sub-edit.  Pure semantic
plans retain projected direct/inherited style checks.

## Name and effect correctness

Final-name validation now checks collisions involving explicitly renamed page
targets only.  Unrelated pre-existing duplicate names remain tolerated, while
duplicate touched targets, collisions with untouched names, swaps, and the
old touched no-op behavior remain checked.  Direct, joined, and replacement
segment tests cover an `A,A,C` source renaming `C` to `D`, plus duplicate
refusals and exact replay/inverse behavior.

`Edit::join` derives every durable write key from the payload and requires the
declared write set to match.  Repeated non-package semantic keys are refused
atomically, closing concealed same-target writes; repeated `package` keys are
retained for ordered replacement capability.  Staged and cross-sub-edit
package conflict checks also use payload-derived keys.  Underdeclared
replacement plus ordinary staged work and duplicate semantic payloads return
recoverable typed refusals without changing the owner edit.

The active-policy regression creates signed state in replacement one and then
performs replacement two: default security refusal fails at the second
replacement, while `RemoveSignatures` permits the ordered sequence.  This
proves current-package policy state rather than source-only policy checks.

## Validation

The source commit was validated with the memory-capped cargo shim, one job,
`CARGO_INCREMENTAL=0`, debug information disabled, offline locked resolution,
and `TMPDIR=(deleted-coord-artifact)`.  The
owned isolated target was
`(deleted-coord-artifact)`, measured at
414 MiB after the run.  External receipts are intentionally not tracked:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check --offline --locked --all-targets --all-features -p litchi-odg -j 1` | `(deleted-coord-artifact)/final-check.log` | 162 bytes | `e69de7b1c10bfef88f12096e8175f2d2fa1c38099d53490f0593a5652cfd26fc` |
| `cargo clippy --offline --locked --all-targets --all-features -p litchi-odg -j 1 -- -D warnings` | `(deleted-coord-artifact)/final-clippy.log` | 162 bytes | `8d25179bc5fc58c574f54aa89d66b84337247fc986e4753a2ab7ee5bf44a6069` |
| `cargo test --offline --locked --all-features -p litchi-odg --test package_semantics --test transition_review --test capability_transactions -j 1` | `(deleted-coord-artifact)/final-tests.log` | 5063 bytes | `cca93855dafe53f62bfa48af1c5891a0d131e1f9a480e26ae1170d63907ee196` |

The focused run passed 8 capability tests, 26 package-semantic tests and 20
transition-review tests (54 total).  `cargo fmt -p litchi-odg -- --check`,
`git diff --check`, the API glossary policy check and the repository artifact
guards passed.  No benchmark, fixture, binary or large repository evidence
artifact was added.

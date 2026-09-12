# 0533 CFB claim-sector test review

`scope: read-only test and differential-coverage review before implementation`

`baseline: 0dd079b95982aa1fbef5eddd1ec3743261dcaf20`

No candidate-dependent test source exists yet. This review identifies the
smallest useful tests for the planned cold-error-helper and ordinary-inline
change. No Rust source, build, test, benchmark, or profiler was run here.

## Existing coverage

The private tests in [`file.rs`](../../../../crates/litchi-cfb/src/file.rs)
and [`allocation_validation_tests.rs`](../../../../crates/litchi-cfb/src/allocation_validation_tests.rs)
already exercise the surrounding ownership pipeline:

| behavior | existing coverage | remaining direct gap |
| --- | --- | --- |
| Successful role publication through real CFB opening | real compound-file corpus, FAT/DIFAT, MiniFAT, directory, mini-stream, and regular-stream tests | No focused assertion that one direct claim writes each production role. |
| Duplicate physical ownership | `rejects_self_referential_difat_chains`, `rejects_directory_and_fat_sector_overlap`, and the real open paths | No direct exact-message assertion for `claim_sector` with controlled old/new roles. |
| Missing physical slot | higher-level sector/range checks and malformed input tests | No direct exact-message assertion for `claim_sector` when `sector_roles` is empty or too short. |
| `u32` to `usize` conversion failure | source branch exists | Unreachable on the current 64-bit host; use source review or a supported narrower-target test rather than a misleading host test. |
| Chain order and failure behavior | chain collector and overlap/marker tests | The candidate does not touch `claim_chain`; a small direct test can protect earlier-claim retention if the implementation is accidentally broadened. |
| Physical reconciliation and stream allocation order | open tests and dedicated allocation-validation cases | Must remain in the existing full package/test gates; do not duplicate their implementation in a microtest. |

The existing tests therefore provide broad integration coverage, but they do
not isolate the three extracted formatter branches or prove that the role map
is unchanged when a direct claim fails. A focused private test is warranted
because the candidate moves code across closure/function boundaries while
claiming exact error and mutation behavior.

## Smallest useful candidate-independent test

Add one test in the existing private `file.rs::tests` module, for example
`claim_sector_preserves_checked_success_and_failure_contract`. It can build a
minimal `OleFile<Cursor<Vec<u8>>>` with the other fields empty and a controlled
`sector_roles` vector. The test should:

1. On fresh one-slot maps, call `claim_sector(0, role)` for each production
   role (`Fat`, `Difat`, `Directory`, `MiniFat`, `MiniStream`, and
   `RegularStream`) and assert `Ok(())` plus the exact stored role.
2. On an empty map, call a valid role and assert the exact bounds display:
   `directory sector 0 is outside the file` (or use a table of role/expected
   label pairs to cover more than one label).
3. On a map containing `Fat`, request `RegularStream` for sector zero and
   assert the exact conflict display:
   `Sector 0 is claimed by both FAT and regular stream`.
4. Snapshot the role map before each failure and assert it is byte-for-byte
   unchanged afterward.

This test calls the unchanged private API and remains valid if the production
candidate is rejected, so it is eligible to be retained as regression coverage.
It should compare `OleError::to_string()` or the exact `CorruptedFile` payload,
not merely `is_err()`, because formatter extraction is the behavior under
review. If the test uses a role table, include every production role while
excluding `Unclaimed`, which is the sentinel state rather than an ownership
claim.

The conversion branch should be documented as statically preserved. A
`#[cfg(target_pointer_width = "16")]` assertion is useful only if the project
supports and runs that target. On the current 64-bit host, constructing a
large `u32` cannot trigger `usize::try_from` failure; adding a host test that
pretends otherwise would test a different implementation.

## Optional call-sequence guard

If the candidate patch also touches `claim_chain`, add a second private test
with roles `[Unclaimed, Fat]` and a request for `[0, 1]` as `RegularStream`.
It should assert the conflict at sector one and verify that sector zero
remains `RegularStream` while sector one remains `Fat`. This records the
current sequential, partial-on-late-error behavior of `claim_chain`; it must
not be added merely to justify changing that behavior. The proposed 0533
source scope does not require a `claim_chain` edit, so the one direct
`claim_sector` test is the preferred minimum.

No differential test should reimplement `claim_sector` or compare machine
instructions. The old checked semantics are simple enough to assert through
the private method, and the performance lane owns assembly/code-size
comparison. Likewise, allocation counts, peak memory, and latency do not
belong in a unit test; they are covered by the frozen native, allocator, and
profile lanes.

## Required broader gates

The focused test must be accompanied by the frozen CFB/OLE2 quality commands:

- all-feature and no-default-feature `litchi-cfb` tests;
- all-feature `litchi-xls`, `litchi-doc`, and `litchi-ppt` tests;
- workspace check, warning-denied Clippy, rustdoc, crate-boundary, and strict
  performance-claim checks; and
- the existing real DOC/XLS/PPT and malformed CFB corpus tests.

The quality matrix is necessary because `claim_sector` is shared by direct
CFB, XLS, DOC, and PPT ingress. It is not necessary to add duplicate
role-specific fixtures to each consumer crate when the private direct test
and existing corpus already exercise the common implementation.

The candidate should be retained only after the frozen 0533 native and
allocation protocol passes. The four primary XLS workflows each require at
least 3% lower p50 in both paired repeats and lower constructor instructions
in both repeats. CFB tiny, many-small, and few-large are sensitivity guards;
every row's mean, tail, RSS, and allocation changes still require review.
Any same-build or matched adverse variation over 5% is a review trigger. A
failed performance gate rejects the production attribute/helper change while
leaving the direct test above available as candidate-independent correctness
coverage.

The test review disposition is therefore: **focused direct coverage required;
broader existing gates remain applicable; no candidate-dependent differential
test is needed beyond exact private success/error/state assertions unless the
implementation expands beyond `claim_sector`.**

## Applied test review

The final `tests.patch` contains two compact tests with a literal seven-role label table. The repeated-claim matrix covers all 49 old/new role pairs, exact conflict payloads, sentinel behavior, and unchanged neighboring/previous ownership. The boundary matrix covers all seven roles over eight explicit empty/first/last/out-of-file cases, including `u32::MAX`. Expected valid indices are table data, not a copied implementation. The boundary test is compiled on 32/64-bit targets; only the current x86-64 host was executed. Conversion failure remains a static source-review obligation. Both CFB feature configurations passed with 308 executions each.

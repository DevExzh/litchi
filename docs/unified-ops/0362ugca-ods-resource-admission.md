# 0362ugca — ODS resource admission stays out of package open

- Date: 2026-10-24.
- Status: coordinator review correction and pinned ODS validation complete on
  `uo2/ods-lazy-entry`; integration and the final workspace gate remain
  pending.
- Performance claim: no native timing or baseline claim.

## Review return and correction

Review of `0362ugc` rejected the first resource-count implementation because
`Archive::from_family_package` called `family::Package::files()`, allocating a
full member-name `Vec` and scanning it during every package open. That moved
edit-admission work into construction and did not close the construction-cost
requirement.

`Archive` now retains an empty `OnceLock<usize>` at open. The non-reserved
resource count is computed only when a caller-selected resource limit is first
admitted by an operation on an `Edit` or by commit/replay validation; the
count is cached for subsequent validations. `Package::edit`, `from_vec_with`,
and shared-package handoffs do not perform this census. The first staging
boundary admits both the current source and the candidate archive before
retaining either as edit state, so an over-limit source or candidate is
refused without staged retention. A first `commit()` performs the same check
before publishing an unchanged snapshot. Dropping an untouched edit leaves
the cache empty. The existing archive index, content validation, encryption
checks, parser limits, source checks, and publication behavior are unchanged.

The flat ODS edit remains sparse: only touched sheets enter its `BTreeMap`,
and the aligned sheet vector is materialized at a changed commit. DDE,
sheet-metadata, and source-cell lazy parser/protection behavior from
`ce3c80bc74` and `2e2f902179` is unchanged.

## Source and validation boundary

The correction is in the successor source commits for `0362ugc`, with the
redundant `#[must_use]` Clippy fix in
`1fdfcdcdba11b2860c08037aa9939450770ecc84`. The focused regression
`resource_admission_stays_cold_until_edit_operation_or_commit` records the
construction/drop, reopening/handoff, first-operation, and pre-staging
failure proof. The final pinned ODS suite and strict all-target/all-feature
Clippy receipts are recorded in `0362ugc`; the full workspace gate must rerun
after integration.

The full suite exited `0` with 119 result blocks and 1,768 passed tests;
Clippy exited `0` with no diagnostics under Rust/Clippy 1.95.0. No binaries,
fixtures, or evidence files were added.

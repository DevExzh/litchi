# XLSB primary package transactions

The native structural family now starts with `Package::edit()` or
`edit_with_limits(TransferLimits)`. Commits carry the canonical triad:
`package()`, `patch()` and `diagnostics()`, with `into_package()` and
`into_parts() -> (Package, Patch)`. `Package::apply(&Patch)` returns a new
package and leaves the source untouched on success or failure.

The replaced `Workbook::edit_workbook_structure`,
`edit_workbook_structure_with_limits` and `apply_workbook_structure` methods
are removed, without aliases. Both glossary profiles ban them. Native tests
and performance harness callers are migrated. `Workbook::to_package()` is a
policy-preserving shared-graph conversion for callers that still hold a workbook
projection; `into_package()` remains the consuming conversion. Neither copies
the semantic projection or worksheet caches.

Edits retain their package base. Replay consumes the exact serialized before
image using retained OPC source limits and projection policies. The final
validated replay workbook supplies the commit package without a second semantic
parse. Its generated OPC image is admitted under the same finite read policy
before commit publication, so oversized output fails at commit rather than only
at later application. Effective `read_limits()` are retained even when no source
archive is retained; archive-only limits are not used as a substitute. Exact no-op commits retain the original package owner, including its
signature and source metadata. Application checks the exact source first,
returns the unchanged owner for exact no-ops, and otherwise requires an unsigned
source and validates the candidate with destination source/projection policies.

Starting replay directly from the in-memory OPC graph was rejected by the
regression suite: two drawing cases produced after-images different from durable
replay. The implementation therefore retains the canonical before-image parse.
It does not replace durable byte proof with semantic equivalence.

## Consumer and performance scope

The native benchmark lifecycle now uses `Package::edit`, `Commit::into_package`
and `Package::write_to`. Its metadata names the new timing boundary; historical
results describe the previous mutable-workbook lifecycle and are not directly
comparable to this route. The lifecycle test verifies deterministic output and
reopened semantics. No speedup or global performance acceptance is claimed.

Full-image staging, serialization and durable replay remain. This wave does not
yet enable unified XLSB writes or migrate every Workbook mutator. Mutable
workbook patch/history helpers remain for existing focused native consumers.
Their merge/decode helpers still use default OPC limits; they are outside the
new package-application policy contract. The committed tolerance report is the
semantic replay report, not a newly computed semantic report of the physical
readmission. The structural family still uses the existing package error/limit types rather
than a complete scoped Error/LimitKind contract, so it is not registered as a
complete glossary family. Detached worksheet-family leniency propagation and
explicit lenient durable-decode policies also remain follow-ups. The all-format
Rust/Python/fuzz/performance goal remains open.

Review identified the bounded-chaining issue before finalization. Regression
coverage now includes exact-input-budget growth refusal, unchanged source on
failure, successive within-budget edits, identical commit/apply output, and
retained effective limits on both paths. Existing signed and skipped-drawing
coverage runs through the package entry points. The three primary-package
integration tests also exercise durable forward/inverse application, stale
refusal, no-ops, failed staging and the consuming commit accessors.

## Verification

The final native `sign,vba-inspection` suite passed 1,087 tests and doctests,
including the budget/chaining regression. The earlier default-feature suite
passed 1,075 before that last test and admission fix; it is not the final-code
full default-suite claim. All 489 XLSB-only facade tests and doctests passed. Native and facade all-target
Clippy passed with warnings denied. Warning-denied minimal/full rustdoc JSON
builds and both glossary profiles passed (20 banned names each). Scoped formatting
and diff checks passed; crate boundaries passed for 66 packages and 268 internal
dependency declarations. Independent review found no blocker in the package route.

Builds use `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0
CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true` and a temporary target outside the
checkout. Full-feature and standalone harness builds use temporary copies of
their own lockfiles with `cc` resolved offline to 1.6.0 through Cargo's unstable
lockfile-path option; neither inherited lockfile is changed.

Reproduce with `cargo test -p litchi-xlsb --features sign,vba-inspection` and
`cargo test -p litchi-ops --no-default-features --features xlsb`. Use matching
feature sets for all-target Clippy with `-- -D warnings`. Build warning-denied
minimal/full XLSB rustdoc JSON and check profiles `xlsb-minimal` / `xlsb-full`.
For the standalone harness run `cargo test --manifest-path
tools/perf-baseline/Cargo.toml --lib xlsb_lifecycle_selectors_are_opt_in_and_hash_stable`,
and Clippy with that manifest plus `--features xlsb-crud --lib --bin xlsb_crud`.

The final migrated harness lifecycle test passed, including deterministic hashes
and semantic reopening. Harness library and `xlsb_crud` Clippy passed with
warnings denied. These are correctness/build checks, not paired performance
acceptance.

The affected harness skipped-drawing reparse regression also passed. Temporary
build targets, rustdoc JSON, copied lockfiles, scripts and logs were removed
after verification. No worktrees or borrowed fixtures were created. Both
inherited lockfiles are unchanged and the rewritten base remains an ancestor.

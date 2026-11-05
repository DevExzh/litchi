# XLSB publication retains the caller's read policy

A workbook explicitly opened with `Leniency::TolerateCachedNumbers` now retains
that policy after cell-value, cell-watch and sparkline publication. Previously,
the detached publication helpers reparsed candidates with strict defaults and
replaced the owner with that strict projection. Structural cell insertion and
transfer also read the target worksheet under the retained policy, allowing
unrelated edits without rejecting or normalizing an existing tolerated cache.

The common internal reparse seam takes the selected leniency alongside the
external-link and drawing policies. Workbook callers pass their retained policy;
raw OPC helpers pass `Strict`. A raw OPC package carries no semantic projection
policy, so reopening its output with defaults still produces strict cell reads.
The change neither relaxes authored-value validation nor rewrites tolerated
number bytes. Exact no-ops retain their original owner and bytes.

Legacy mutable structural `Patch::apply` now shares the primary package route's
changed-candidate validator. Both routes enforce the destination's finite OPC
limits, external-link policy, drawing policy, leniency and signature guard. The
mutable route adopts the already validated workbook, avoiding a second semantic
parse. Exact-source checks still precede no-op handling, and rejected candidates
leave the caller untouched.

## Verification

The focused regression corpus injects an invalid NaN cache into an otherwise
valid workbook. Tests compare the complete original cell record, including its
BIFF12 header, after unrelated edits; inspect the reported `#NUM!` projection;
and verify that default reopened cell reads still refuse the cache. Coverage
includes all three worksheet publication families, structural transfer,
mutable forward/inverse replay, exact no-ops, invalid dependency atomicity and
retained destination input budgets.

The full native `sign,vba-inspection` suite passed 1,096 tests/doctests and the
XLSB-only facade suite passed 501. After adding the destination-budget assertion,
all seven focused publication-policy tests passed. Native/facade all-target
Clippy passed with warnings denied. Scoped formatting and diff checks passed;
crate boundaries passed for 66 packages and 268 internal dependency declarations.
Warning-denied native rustdoc JSON and the full XLSB glossary profile passed
(20 banned names). Independent review found no production blocker.

Reproduce using `cargo test -p litchi-xlsb --features sign,vba-inspection`,
`cargo test -p litchi-ops --no-default-features --features xlsb`, and the native
`--test publication_leniency` target. Builds used an external target directory,
offline dependencies, zero dev/test debug information and disabled incremental
compilation. Full-feature runs used an external copy of the root lockfile with
cached `cc` 1.6.0 selected through Cargo's unstable lockfile-path option.

## Remaining work

This closes the detached publication leniency gap recorded in 0167–0168. It does
not migrate the worksheet families to their canonical package-owned API or add
unified cell writes. The next owner migration can place the existing raw
worksheet transaction family under `cell_values::worksheet`, then expose
package-owned `cell_values::{Edit, Commit, Patch, Diagnostics}` through
`Package::edit_cell_values` and `apply_cell_values`. Replaced Workbook mutation
entries must be removed, not aliased. Scoped errors/limits, explicit lenient
durable decoding, command coalescing and the all-format Rust/Python goal remain
open. No new performance claim is made for this policy correction.


Temporary targets, copied locks and raw logs were removed after recording these
results. No temporary worktrees or borrowed fixtures were created. Both inherited
lockfiles are unchanged; the rewritten base remains an ancestor. The goal stays
active.

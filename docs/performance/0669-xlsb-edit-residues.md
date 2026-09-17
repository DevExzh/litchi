# 0669: XLSB edits publish the parse they validate and repair resource relationship guards

Status: retained, implemented in `crates/litchi-xlsb`.
`performance_claim: none` — this change reports structural work and correctness
evidence only. No timing or speedup claim is registered.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This closes row 13 of change [0651](0651-queue-refresh-after-the-second-wave.md),
under the correctness-first trade-off in
[0652](0652-owner-decisions-for-the-third-wave.md). Change 0599 removed the
duplicate parse from `Workbook::apply_cell_values`, but its workbook-structure
replay helpers still called the package publication wrapper and then parsed the
same candidate again. Change 0647 also identified two resource paths whose
presence test examined a part while the operation actually needed a workbook
relationship. The sparkline and cell-watch facades had the complementary state
problem: they published the changed worksheet package while retaining the
pre-edit derived `Workbook` fields.

## What changed

`cell_values/root.rs` now calls the existing crate-private
`cell_values::workbook::apply_retaining_parse` from both
`insert_candidate_cell` and `transfer_cell`. The helper returns the complete
`Workbook` parse that validated the candidate. Each caller installs that boxed
parse directly, so the candidate is cloned and parsed once for this publication
path. The old second `Workbook::from_opc_package...` call is gone. These
operations always produce a changed candidate; an unchanged result therefore
leaves the staged owner untouched.

`cell_watches/workbook.rs` and `sparkline/workbook.rs` now have the same
crate-private retaining outcome: `Unchanged(snapshot)` for an exact byte no-op,
or `Published { snapshot, workbook }` after the changed candidate passes a
complete XLSB workbook parse. The public cell-watch package wrapper keeps its
existing signature and takes the package back out of the validated workbook.
The typed `Workbook::apply_cell_watches` and `Workbook::apply_sparklines`
facades install the returned workbook directly. Their worksheet snapshots and
source checks remain unchanged, while all package-derived fields now come from
the parse that validated the published candidate.

The publication order remains staged and atomic: the input package is borrowed,
the worksheet patch is applied to a clone, the updated worksheet is decoded,
signatures are removed on the changed candidate, and the complete workbook is
parsed before one final owner assignment. A malformed patch, a stale commit, a
worksheet decode failure, a workbook dependency failure, or a complete parse
failure returns before the caller's package or workbook is changed. Exact
no-ops still return the commit snapshot without cloning or parsing a candidate.

`cell_values/resources.rs` now treats a resource part and its workbook
relationship as two independent obligations. `ensure_styles_part` creates
`styles.bin` only when needed and then adds the exact internal styles
relationship if it is absent. `intern_shared_string_for_new_cell` follows the
same rule for `sharedStrings.bin`, including the strict relationship variant.
The relationship test matches the internal target mode, exact relationship type
and exact relative target used by `Relationships::get_or_add`; an unrelated or
external relationship cannot satisfy it. Existing parts with missing
relationships are therefore repaired, while existing relationships are reused
without duplication.

Two focused resource tests construct those previously missed states: an
existing styles part without its relationship, and an existing shared-strings
part without its relationship. The cell-watch and sparkline facade tests also
compare the published workbook's debug projection with a fresh parse after a
successful edit, and retain their stale-commit atomicity checks.

## Scope and coordination

The branch started at `5fa92d7ce` in
`/home/zhuhe/code/litchi-worktrees/0669` on
`perf/0669-xlsb-edit-residues`. It does not modify `litchi-opc` and does not
assume any new lazy OPC API from the concurrent 0661 migration. The two changes
use the existing `OpcPackage` and `Workbook` ownership seams, so the coordinator
should cherry-pick this commit with the 0661 branch's normal compile gate and
resolve only any incidental context drift.

The shared performance harness is intentionally untouched because 0674 owns its
selectors and gates. The selector handoff for that agent is the complete
workbook-structure path: build an edit through
`Workbook::edit_workbook_structure`, commit it, call
`Workbook::apply_workbook_structure`, save, and perform a typed readback. The
selector should separate planning/commit from publication so the candidate
reparse removed here is visible without attributing unrelated open or save work
to it. This record makes no measurement claim for that selector.

## Verification

All commands below used `CARGO_PROFILE_DEV_DEBUG=0`,
`CARGO_PROFILE_TEST_DEBUG=0`, and Cargo jobs `2`; no existing target directory
was cleaned.

* `cargo fmt --all -- --check` passed.
* `cargo check -p litchi-xlsb -j2` passed.
* `cargo clippy -p litchi-xlsb --all-targets -j2 -- -D warnings` passed.
* `cargo test -p litchi-xlsb --lib -j2` passed: 575 tests.
* `cargo test -p litchi-xlsb --tests -j2` passed: 575 library tests and 144
  integration tests.

The complete evidence packet is
[`results/change-0669/`](results/change-0669/). It contains the decision,
four log sections, and the gate summary. The packet records no benchmark result
and no claim-registry entry.

## ADR and goal audit

ADR 0003's staged publication rule is preserved by the retained parse APIs and
the single final owner assignment. ADR 0005's bounded, typed validation remains
in force: no limit is relaxed, and all existing worksheet, workbook,
external-link and resource checks still run before publication. ADR 0006's
preservation default is preserved: no-op worksheet bytes remain exact, changed
bytes are published only after validation, and resource relationship repair is
performed on a staged package. The accepted lazy OPC ADR 0030 is not amended or
referenced by production code here; this branch leaves its API migration to
0661. No unsafe code, dependency, global cache, ambient I/O, executor or public
archive ownership is introduced.

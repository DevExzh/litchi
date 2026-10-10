# XLSB package-owned cell-value transactions

`Package::edit_cell_values(selector)` owns the complete publication boundary.
Names and checked catalog positions select worksheets; missing selections return
`None`, and non-worksheet tabs are refused. An edit stages cell commands in one
worksheet transaction and validates the resulting package at commit. The commit
exposes `package()`, `patch()` and `diagnostics()`, with consuming accessors for
the published package and patch. The original package remains immutable.

The replaced `Workbook::edit_cell_values*` and `apply_cell_values*` methods are
removed without aliases. Native tests, facade consumers and the existing XLSB
CRUD harness migrate in the same change. Raw worksheet transactions move to
`cell_values::worksheet`; their wire format and history operate on worksheet
images, not complete package images. Their no-op predicate is `is_noop()`.

Package patches bind exact before/after package images. Application checks the
source before accepting an exact no-op, and inverse replay restores the original
image. Changed publication retains the destination's finite OPC limits,
external-link policy, drawing policy and leniency. Signature and dependency
checks happen before a changed package is returned. Package image buffers stop
at the transfer-byte quota before growing, and semantic patch-change allocation
is capped before inserting changes. Worksheet parsing and temporary codec maps
remain governed by the separate finite worksheet limits.

## Verification

The final native `sign,vba-inspection` suite passed 1,109 tests/doctests and
XLSB-only facade tests passed 501. The combined DOCX/XLSX/PPTX/XLSB/encryption
facade suite passed 1,170 before the final patch-change allocation cap; the
final native and XLSB-only reruns cover that tightening. Eleven package-owner
regressions cover the commit triad, immutable source, exact forward/inverse
replay, stale no-ops, ambiguous catalog refusal, a chart tab before a worksheet,
contextual style validation, raw worksheet access, retained input budgets and
transfer image/change quotas.

Final native/facade all-target Clippy passed with warnings denied. Native
warning-denied rustdoc JSON and the full XLSB glossary profile passed (26 banned
names). The glossary checker passed 21 unit tests, and OfficeCLI capability
parity regeneration matched the committed report. Independent review found no
remaining blocker after the bounded image/change allocation fixes.

All nine XLSB CRUD harness tests passed after consumer migration.
Final crate boundaries passed for 66 packages and 268 internal dependency
declarations. Scoped formatting and diff checks passed.

Reproduce with `cargo test -p litchi-xlsb --features sign,vba-inspection`,
`cargo test -p litchi-ops --no-default-features --features xlsb`, and the same
facade command with `docx,xlsx,pptx,xlsb,encryption`. The standalone consumer
harness uses `cargo test --manifest-path tools/perf-baseline/Cargo.toml
--features xlsb-crud --bin xlsb_crud`. These are correctness checks, not timed
performance evidence. Builds used external target directories, offline
resolution, zero dev/test debug information and disabled incremental builds.
Root builds used an external lockfile copy with cached `cc` 1.6.0 selected by
Cargo's unstable lockfile-path option; both inherited lockfiles are unchanged.

## Scope and remaining work

This migrates native ownership and consumers. Unified XLSB cell properties are
still unavailable; the existing unified worksheet-name route is unchanged.
Scoped family errors/limit kinds, durable package cell-value patch encoding,
package cell-value history/merge, facade cell dispatch and paired cell-write
measurements remain open. The raw worksheet codec and the primary structural
family retain their existing durable operations. No cell-write performance
claim is made. Complete before/after package images are retained; reducing that
cost without weakening exact replay remains performance work.

The all-format Rust/Python goal remains active. Other leaf mutation families,
remaining format coverage and full acceptance gates are not claimed complete.

Implementation tasks used isolated temporary worktrees; those worktrees and
branches were removed after integration. Scratch build targets, copied locks
and raw logs were removed after recording validation. The base rebase was
already up to date at the start of this milestone. No benchmark binary or
large evidence file is committed.

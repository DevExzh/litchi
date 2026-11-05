# 0362uo — Workspace Clippy v2 consumer repairs

This source-only repair starts at integration commit `af415286f3d4671bef80c243b9710fc2f70f6c53`.
The completed all-workspace, all-target, all-feature Clippy run with warnings
denied reported 35 consumer diagnostics across OPC, ODS, ODT, ODP, ODI and
OTH. The fixes remove only redundant bounds, non-`Drop` `drop` calls,
needless `Ok(...?)` wrappers, same-type error conversions, and one redundant
borrow. The OPC scoped sinks now end their borrows by leaving a block; no
lint allowance or behavior change was introduced. Existing error conversions
remain where the source and destination error types differ, and all selected
source policies and publication fences are unchanged.

The immutable terminal log is retained outside the checkout at
`(deleted-coord-artifact)/check-0.log`
(28,036 bytes, SHA-256
`7254dd2fba24fb3c589e934d2b5e2034f033fa66ba6063b62f40d56f6044ef57`). The
command exited 101 on the unmodified integration source after 448.53 seconds;
the failures are the diagnostics this source repair addresses. Cargo was not
run in this isolated leaf.

## Validation

- `rustfmt --edition 2024` and `git diff --check` passed.
- The repository artifact guards pass for the staged source and record files.
- A coordinator rerun of the affected workspace Clippy and tests is required
  after this source-only commit is integrated.

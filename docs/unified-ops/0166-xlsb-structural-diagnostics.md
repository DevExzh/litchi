# XLSB structural diagnostics and exact no-ops

The structural transaction family now reports the ADR 0034 introspection pair:
`commit.diagnostics().changed()` and `patch.is_noop()`. Diagnostics are computed
once during commit and returned as a small, copyable value. Patch introspection
compares the exact canonical package images without serializing or reparsing;
its comparison can inspect all image bytes.

Operation counts are not change counts. A same-name rename retains a forward
operation but can leave the package unchanged. An inverse patch has no forward
operations but can change the package. Tests cover both cases, empty edits,
changed edits, durable decode, inverse restoration and stale no-op refusal.

A signed-package regression exposed an existing replay failure: a same-name
rename left the package unchanged, then attempted to serialize the validated
signed graph again. That failed with `SignedSourceRequiresExplicitPolicy`.
Rename replay now reports whether its encoded workbook part changed. When all
operations leave the package unchanged, replay retains the exact original
bytes after validation. Cell and drawing publication remain conservatively
classified as changed. The signed regression proves that empty and same-name
transactions preserve the signature and archive bytes, while a real rename
still fails closed.

Patch application checks the exact source before testing `is_noop()`. Any
identical-image patch returns without replacing the workbook, including a patch
with retained no-op operations. Changed patches still require an unsigned source
and complete candidate readback. The durable wire format is unchanged.

## Owner migration findings

`Package` already wraps an OPC owner but `Workbook` retains public mutators and
another OPC graph. The facade uses the latter. Adding aliases would leave the
ownership model wrong. The next owner wave should:

1. Introduce primary package edit/apply entry points and capture physical lineage
   and ingress policies in the edit.
2. Keep workbook replay internal and adopt its already validated candidate into
   the returned package, avoiding a second parse.
3. Move consumers, remove the structural edit/apply methods from public Workbook,
   and supply the commit package/patch/diagnostics triad and consuming accessors.
4. Preserve external-link, source, leniency and drawing policies, signatures,
   source freshness and source retention across conversions. The current Package
   does not carry the full projection policy; byte reconstruction still uses
   default OPC limits.
5. Move remaining mutations to immutable package publication and convert the
   facade write kernel to that owner. Keep source-backed reads lazy.

Full-image staging, source comparison and replay remain. This wave makes no
throughput or changed-closure performance claim and does not enable unified XLSB
writes. Complete owner/commit/error contracts, Python coverage, Tier 2/Tier 3
write loops, fuzz and global performance acceptance remain outstanding.

## Verification

The final native default-feature suite passed 1,067 tests and doctests; the
`sign,vba-inspection` suite passed 1,078. The new signed same-name regression
failed before the replay fix and passed afterward. Native full-feature
all-target Clippy passed with warnings denied. XLSB-only facade tests and
doctests passed 489 cases, and facade all-target Clippy passed with warnings
denied. Warning-denied minimal/full-feature rustdoc JSON and both glossary
profiles passed (17 banned names each). Scoped formatting and diff checks passed.
Independent review found no blocking issue in the diagnostics or replay changes.

Tests use `CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0
CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true` and a temporary target directory.
Full-feature checks use a temporary lockfile copy with offline `cc` 1.6.0 via
Cargo's unstable lockfile-path option; the workspace lockfile stays unchanged.
Reproduce with `cargo test -p litchi-xlsb`, then with
`--features sign,vba-inspection`; facade tests use `cargo test -p litchi-ops
--no-default-features --features xlsb`. Clippy uses the same latter two feature
sets, `--all-targets -- -D warnings`. Run warning-denied XLSB rustdoc JSON
builds for minimal/full features and check `xlsb-minimal` / `xlsb-full` with
`tools/check_api_glossary.py`.

The crate-boundary checker passed for 66 packages and 268 internal dependency
declarations. Temporary targets, generated documentation, logs, scripts and the
copied lockfile were removed after verification. No worktrees or borrowed
fixtures were created. The workspace lockfile is unchanged and the rewritten
base remains an ancestor.

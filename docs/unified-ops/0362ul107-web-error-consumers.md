# 0362ul107 — DOCX web-extension typed-error consumers

- Date: 2026-10-25.
- Source under repair: `e130751fc503099c4d5fff74a5746b5099d625e1`, branch
  `uo2/web-error-consumers-0362ul107`.
- Scope: DOCX web-extension and durable replay test consumers. No production
  error mapping or publication code is changed.

## Frozen runtime finding

The coordinator's all-feature runtime reached the DOCX web-extension batch
with 26 passing and 8 failing tests. One failure,
`conformance_noop_ignores_missing_unused_prefix_and_vendor_other_dialect_text`,
is an unrelated production conformance/no-op capability issue and is left for
its assigned owner. The other seven failures were stale assertions that still
looked for source, stale, relationship, or durable forward-replay diagnostic
text after the shared OOXML adapter began returning DOCX's typed
`Error::PatchConflict` with its fixed content-free reason.

The captured diagnostic remains outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/ooxml-pptx-joint-e130751fc5-f5-runtime-remaining33/11-tests-docx-web-extensions-attempt-actual-f5-11.log`
(22,345 bytes, SHA-256
`a1ff45c4d8d9a30985407292221578b1ed77024ce95af156b15dceff2af61d25`).

## Consumer repair

`web_extensions.rs` now has one typed `assert_patch_conflict` helper. The
stale content-types/relationship cases, semantic no-op cases, and the signed
patch after explicit `unsign` all assert the exact public
`Error::PatchConflict { .. }` variant. `web_extensions/durable.rs` now retains
the rejected `litchi_docx::Error` instead of flattening it to a string, so the
durable stale-source, altered-restore, signed-after-unsign, and forged inverse
closure cases use the same typed assertion. Their existing package-byte
atomicity checks remain in place. Operation-tamper, target-tamper, and
resource-limit diagnostics retain their existing purpose-specific checks.

The DOCX-only audit found no analogous stale web/durable consumer in PPTX,
XLSX, or XLSB. Their existing web/task-pane tests either have no shared
durable web route or already assert their format-level typed variants. No
production capability, source fidelity, inverse behavior, or signature policy
was widened or removed.

## Validation

No Cargo, rustdoc, metadata, or runtime rerun was performed; the coordinator
owns the frozen runtime rerun. This leaf is source-only and records the
pre-fix failure honestly. `rustfmt --edition 2024 --check` was run against
both changed Rust files; it reports only pre-existing formatting differences
in untouched match-arm punctuation in `web_extensions.rs` (the changed hunks
have no rustfmt diagnostics). `git diff --check` and the repository size
guards pass.

## Configured-format follow-up

The initial source-only check above was run before the sparse leaf materialized
the tracked root formatter inputs, so it used the host formatter and was not
the workspace gate. The follow-up materialized `rust-toolchain.toml` and
`rustfmt.toml` from this commit's source tree and ran the full owned files
under the pinned `1.95.0` toolchain:

```text
rustfmt --edition 2024 \
  crates/litchi-docx/tests/web_extensions.rs \
  crates/litchi-docx/tests/web_extensions/durable.rs
rustfmt --edition 2024 --check \
  crates/litchi-docx/tests/web_extensions.rs \
  crates/litchi-docx/tests/web_extensions/durable.rs
```

Both commands exited 0; the files were already compliant under the actual
workspace configuration, so this follow-up has no Rust source delta. A fresh
`git diff --check` and the repository size guards also pass. No Cargo,
runtime, or production validation was run.

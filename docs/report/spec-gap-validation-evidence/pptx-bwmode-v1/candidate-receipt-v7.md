# PPTX p14:bwMode candidate receipt v7

- Isolated worktree: `/var/tmp/litchi-pptx-bwmode-validation-20260910`
- Base commit: `cf98bee37455e8e6d0e73002ff2e94299ef546cb`
- Publication status: no commit, no staging, primary worktree untouched.
- Candidate patch: `/var/tmp/litchi-pptx-bwmode-validation-20260910-artifacts/bwmode-candidate-v7.patch`
- Candidate patch SHA-256: `6a7aeab3762b4f8fbbe53651c3009dc29beb52ef71cfc51cb3ded20dd83a26ef`

## Scope

The candidate adds inert typed `BlackWhiteMode` read accessors and source-checked `Transaction::set_black_white_mode` / `set_bw_mode` editing on the qualified `p14:bwMode` attribute of `p:contentPart`. It retains unknown anchor markup, authored omission versus `auto`, inherited/custom prefixes, local namespace insertion, MCE inactive branches, atomic validation, inverse patches, and opaque payloads. It does not add a playback engine or infer rendering behavior.

This v7 delta also:

- Processes content-part MCE with an explicit bounded capability set that understands the PowerPoint 2010 `p14` namespace, so canonical `mc:Choice Requires="p14"` activates while unsupported choices still fall back.
- Refuses non-empty content-part patches while `OpcPackage::is_signed()` or `requires_signature_edit_policy()` is true. Empty commits return before that guard and retain signature infrastructure. A caller can explicitly authorize the mutation with `package.unsign()` first; the patch path no longer unsigns a staged clone implicitly.
- Commits and reopens all 11 `ST_BlackWhiteMode` tokens through the package XML path.
- Rejects typed interpretation of unqualified or foreign `bwMode`, and keeps `p14:media` opaque.

Changed paths:

- `crates/litchi-pptx/src/presentation/embedded/content_parts/model.rs`
- `crates/litchi-pptx/src/presentation/embedded/content_parts/codec.rs`
- `crates/litchi-pptx/src/presentation/embedded/content_parts/mod.rs`
- `crates/litchi-pptx/src/presentation/embedded/content_parts/package.rs`
- `crates/litchi-pptx/src/presentation/embedded/content_parts/transaction.rs`
- `crates/litchi-pptx/src/presentation/embedded/content_parts/validation.rs`
- `crates/litchi-pptx/src/presentation/embedded/content_parts/tests.rs`
- `crates/litchi-pptx/docs/FEATURE_MATRIX.md`

## Local specification evidence

- `3rdparty/specs/[MS-PPTX]/2 Structures/2.3 http---schemas.microsoft.com-office-powerpoint-2010-main.md:553-561`: §2.3.2.2 defines `bwMode` in the p14 target namespace, type `a:ST_BlackWhiteMode`, and says it changes rendering interpretation only.
- `3rdparty/specs/[MS-PPTX]/2 Structures/2.2 Extensions.md:118-128`: §2.2.3 adds `bwMode` to the `p:contentPart` / CT_R owner.
- `3rdparty/specs/[MS-PPTX]/2 Structures/2.2 Extensions.md:159-181`: §2.2.4 media extensions describe `p14:media` and `tracksInfo`; they do not own `bwMode`.
- `3rdparty/specs/[MS-PPTX]/5 Appendix A - Full XML Schemas/5.1 http---schemas.microsoft.com-office-powerpoint-2010-main Schema.md:193-201`: the top-level p14 `bwMode` attribute declaration has no explicit default; absent is retained as `None`.
- `3rdparty/specs/[MS-ODRAWXML]/3 Structure Examples/3.2 Content Parts and Ink.md:179-190`: native example emits `<p:contentPart p14:bwMode="auto" ...>` under an `mc:Choice`.

## Validation

Commands were run from the isolated worktree with Rust `1.95.0`, `TMPDIR=/var/tmp`, and separate targets:

- Focused: `rustup run 1.95 cargo test -p litchi-pptx presentation::embedded::content_parts::tests --all-features --no-fail-fast`; 22 passed. Log: `/var/tmp/litchi-pptx-bwmode-validation-20260910-focused-v7.log`.
- Full crate: `cargo test -p litchi-pptx --all-features --no-fail-fast`; 571 unit tests passed and all package/integration test binaries passed. Log: `/var/tmp/litchi-pptx-bwmode-validation-20260910-all-tests-v7.log`.
- Clippy: `rustup run 1.95 cargo clippy -p litchi-pptx --all-features --all-targets -- -D warnings`; passed. Log: `/var/tmp/litchi-pptx-bwmode-validation-20260910-clippy-v7.log`.
- Rustdoc: `cargo doc -p litchi-pptx --all-features --no-deps`; passed. Log: `/var/tmp/litchi-pptx-bwmode-validation-20260910-doc-v7.log`.
- Formatting: `rustup run 1.95 rustfmt --edition 2024 --check` on the seven changed Rust paths; passed.
- `git diff --check` on the candidate paths; passed.

Targets:

- `/var/tmp/litchi-pptx-bwmode-validation-20260910-target`
- `/var/tmp/litchi-pptx-bwmode-validation-20260910-target-clippy`
- `/var/tmp/litchi-pptx-bwmode-validation-20260910-target-doc`

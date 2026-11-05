# 0362ul105 — DOCX Ribbon signed-error consumer correction

- Date: 2026-10-25.
- Source under repair: `ea5c95e212e1f8e934a0485eb0b1f51355f13ba5`, branch
  `uo2/signature-family-runtime-0362ul105`.
- Scope: the DOCX package-signature consumer for the Ribbon edit family and
  its source record. No production signature or publication code is changed.

## Runtime finding

The coordinator's frozen all-feature runtime reached
`ribbon_family::a_signed_document_refuses_ribbon_changes_and_keeps_exact_noops`
with 8 of 9 tests passing. The one failure was at line 557 of
`crates/litchi-docx/tests/package_signature_policy.rs`: the operation returned
DOCX's direct `Error::Signed`, while the shared `assert_signed` helper expected
the lower-level `Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy)`.
The diagnostic is retained outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/ooxml-pptx-joint-ea5c95e212-f5-runtime-remaining33/08-tests-docx-package-signature-policy-attempt-actual-f5-08.log`
(18,612 bytes, SHA-256
`bba781aa6161f4ba549d45b9ccce907609b8555198f0dec9c49c9d525b58cf6d`).

The shared Ribbon operation intentionally maps its signed fence to the
format-level error. This leaf adds a dedicated DOCX test helper matching
`Error::Signed` and uses it for all three changed Ribbon assertions. The
existing OPC-backed helper remains strict for the other package families; it
is not widened to accept both variants. The Ribbon exact no-op, signature-byte
publication, explicit `unsign`, changed-write, and removal assertions are
otherwise unchanged.

## Consumer audit

The XLSB Ribbon/task-pane signature consumer already asserts its format-level
`PackageError::Signed` for the Ribbon routes. Its separate helper retains the
lower-level OPC alternative because other XLSB package routes can surface that
wrapped error. The PPTX and XLSX package-signature tests have no Ribbon
consumer; their direct format-level and wrapped source-backed assertions are
already separated. No analogous stale assertion was changed in those crates.

This is a source/test correction only. Cargo, rustdoc, and runtime reruns are
owned by the coordinator; this record makes no post-fix pass claim.

## Source-only validation

`rustfmt --edition 2024 --check crates/litchi-docx/tests/package_signature_policy.rs`
and `git diff --check` pass. The staged tree contains only the DOCX test and
this reduced record.

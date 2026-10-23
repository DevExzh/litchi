# 0362upb — DOCX consumer runtime repairs

This leaf starts at `55051a8e4f6fd94e9fb82418ad93336369e3ab59` on
`uo2/docx-runtime`. It repairs the three DOCX failures reported by the
all-feature facade runtime batch without changing the public edit families or
weakening their typed error and signed-source policies.

## Repairs

- The recorded replay test now expects `UnsupportedCapability` for a
  `RefusalKind::UnsupportedContent` refusal. This follows the 0351 mapping;
  the fixed reason and refusal kind remain unchanged.
- A committed source-backed DOCX retains its physical OPC read policy even
  when it also carries an effective edited topology. The facade reads that
  retained policy in O(1) through `Package::read_limits()` instead of asking
  the effective owner for an unavailable physical package. Hyperlink
  relationship resolution therefore remains on the effective owner while
  source admission still uses the original policy.
- DOCX paragraph no-ops compare the projected visible text before scanning the
  paragraph owner. Signed paragraphs containing empty Word bookmark range
  markers consequently preserve their exact bytes and no-op classification.
  Changed text still uses the existing structural writer; only empty
  `w:bookmarkStart`/`w:bookmarkEnd` children directly under `w:p` are
  admitted, while nested or unrelated structural content remains refused.
  Managed edits retain the existing staged-operation path so an edit that
  changes and then cancels is still checked and charged as before.

No source-backed capability, signature protection, or publication path was
removed. No compatibility alias, alternate commit route, or lint suppression
was added.

## Validation

Owned Rust files were formatted with `rustfmt --edition 2024`; `git diff
--check` passed. The repository-wide `cargo fmt --all -- --check` still reports
three unrelated pre-existing files outside this leaf:
`crates/litchi-odf-common/src/package/tests.rs`,
`crates/litchi-ops/src/format/xlsx/sheet_write.rs`, and
`crates/litchi-ops/tests/xlsx_features.rs`.

With one build job, debug information and incremental compilation disabled,
the capped offline locked all-feature tests passed:

```text
cargo test -p litchi-ops --all-features --lib \
  error::tests::docx_recorded_replay_conflicts_are_invalid_input \
  --offline --locked                         # 1 passed
cargo test -p litchi-ops --all-features --lib \
  format::docx::unified_tests::effective_source_owner_resolves_retained_hyperlinks \
  --offline --locked                         # 1 passed
cargo test -p litchi-ops --all-features --lib \
  tests::signature::signed_docx_paragraph_setter_preserves_exact_noops_and_refuses_changes \
  --offline --locked                         # 1 passed
```

The command logs remain outside the repository:

| Log | Bytes | SHA-256 |
| --- | ---: | --- |
| `effective-owner-test.log` | 405 | `0acb228a8de77c942b693d017719620039ac90e057cec78de859ec2ce8786041` |
| `error-test.log` | 5,780 | `23add1d624940414f37c5dd3f6f8bf42b2e2b6b443a7ecc306dec1b37ce1accc` |
| `signature-test.log` | 620 | `1c35c30bd86ced50055231616fbd422b96d16bf74b4fb768ea4d7b9618c72830` |

They are regenerated under the leaf's capped target with the three commands
above and are not committed as evidence.

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
- DOCX paragraph no-ops use the same authoritative owner scan as changed
  text. Signed paragraphs containing empty Word bookmark range markers
  consequently preserve their exact bytes and no-op classification without a
  second parser or a fast path that could discard typed refusal, namespace,
  limit, or cancellation errors. Changed text still uses the existing
  structural writer; only empty `w:bookmarkStart`/`w:bookmarkEnd` children
  directly under `w:p` are admitted, while nested or unrelated structural
  content remains refused. Managed edits retain the existing staged-operation
  path so an edit that changes and then cancels is still checked and charged
  as before.

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

## v16 correction

The v16 runtime exposed that a marker-preserving paragraph scanner must not be
used by an operation that removes the rest of a cell. The policy is now
explicit: `scan_text_owner` remains strict for `replace_table_cell_text` and
the path-addressed `replace_cell_text_at`, while
`scan_paragraph_text_owner` is used only by paragraph-preserving editors and
their managed/base/replay plans. Thus a paragraph edit retains empty direct
bookmark markers byte-for-byte, but whole-cell replacement still returns its
existing typed `ComplexContent` refusal rather than admitting content it could
drop. The new native test also commits a changed bookmarked paragraph and
checks both markers and the exact reversible patch.

The focused follow-up checks passed with the same one-job capped offline
configuration:

```text
cargo test -p litchi-docx --lib \
  document::transaction::element_tests::cell_text_replacement_refuses_content_it_would_drop \
  --offline --locked                         # 1 passed
cargo test -p litchi-docx --lib \
  document::transaction::element_tests::paragraph_text_replacement_preserves_bookmark_markers \
  --offline --locked                         # 1 passed (TMPDIR on external target)
cargo test -p litchi-ops --all-features --lib \
  tests::signature::signed_docx_paragraph_setter_preserves_exact_noops_and_refuses_changes \
  --offline --locked                         # 1 passed
```

The first paragraph-marker attempt was not a source failure: the linker could
not create a temporary file because `/tmp` had exhausted its inode quota
(`df -i /tmp`: 100%). It was rerun unchanged with
`TMPDIR=(deleted-coord-artifact)`.

| Follow-up log | Bytes | SHA-256 |
| --- | ---: | --- |
| `cell-test-v17.log` | 1,110 | `6a1c0a7a624e084573aa181b55e000b0ebd34ce4d34b05a98866eca53f07020d` |
| `paragraph-marker-test-v17-tmpdir.log` | 522 | `0d8ed3814042c535075cc243fcb8e6f513b494968c40de1892c4d547fbe21c63` |
| `signature-test-v17.log` | 620 | `c61b1a8c28f079f1a40fbeb72be5eec47bb9835cb81911d4920f133511a4e26b` |

The inode-exhaustion failure is retained outside the checkout as
`paragraph-marker-test-v17.log` (4,369 bytes,
`6636d22591be4d2cdbf48d4006117c34ccb9489ab4080e56e7db60fb9c9f1ed7`) for
diagnostic provenance only.

## v18 correction

The selected paragraph readers and replay preconditions now use the same
marker-preserving scanner as their writers for:

- `replace_table_cell_paragraph_text`;
- `replace_block_content_control_paragraph_text`; and
- `replace_nested_table_cell_paragraph_text`.

Their whole-cell counterpart remains strict. This keeps changed text,
readback, forward replay, inverse replay, and stale preconditions coherent
without a second scan or an untyped fallback. The focused native test edits a
bookmarked paragraph through all three routes, checks every marker and new
text, then proves durable replay and inverse restoration through the shared
`commit_and_check` path. The all-feature signed DOCX no-op/cancellation test
also passes after the route changes.

| v18 log | Bytes | SHA-256 |
| --- | ---: | --- |
| `selected-routes-v18.log` | 529 | `552a514ade4600fedec1f6b999d696c2bb39d9bc5306ae7d944898e7a3c9ec94` |
| `signature-test-v18.log` | 620 | `f7b6c9a758fa2848e33e3b9194b9f9891564445fabb9acc5a69d5228237f8dbc` |

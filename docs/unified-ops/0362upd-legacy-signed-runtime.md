# 0362upd — legacy signed DOC/XLS runtime repair

- Date: 2026-10-23.
- Source under validation: `77bc2d3ac2f58808b153f50c8f76c1d2c552cde2`, a
  successor of `be5905ed457b239c52652afebe68974bd1a786ff`.
- Scope: the signed DOC fallback in `crates/litchi-doc/src/body_text.rs`, the
  legacy XLS cancellation path in `crates/litchi-xls/src/cell_values/mod.rs`,
  and the signed-policy consumer in `crates/litchi-ops/src/tests/signature.rs`.
- Qualification: no native timing or F3 qualification is claimed. These are
  focused correctness, policy, and feature-compile checks only.

## Repair

The DOC owner retains the exact signed source bytes and uses a private
signature-stripped candidate only when strict mutation admission rejects the
source. Ordinary unsigned opens stay on the original one-parse path. Before
the fallback is attempted, the original CFB is reopened under the caller's
`SourceLimits` and every declared stream is preflighted. This prevents a
large signature stream from disappearing from entry or aggregate accounting.
The profiled path preserves balanced CFB events when that original preflight
returns a typed limit error. Changed signed edits remain refused; only an
exact candidate no-op maps back to the original bytes.

The XLS editor now treats a final Workbook stream equal to the original as a
real empty commit even when a coalesced operation briefly staged and restored a
resource. A structural cell removal followed by reinsertion restores the
original cell outcome; independent mixed-family changes still refuse. The
final byte equality, semantic readback, and signed changed-edit refusal remain
the correctness fences. No performance result is inferred from these focused
tests; the existing XLS source-backed F3 routes still require native ABBA
qualification.

## Source-limit regressions

The test fixture copies a valid generated DOC into a CFB writer and adds a
128 KiB `_xmlsignatures` stream. With a 64 KiB entry ceiling it returns the
typed `EntryBytes` limit, and with a 64 KiB aggregate ceiling it returns the
typed `AggregateBytes` limit. The same entry-limit proof runs through the
performance-diagnostics/profiled open path.

## Receipts

All commands ran with one build job, debug information disabled, incremental
compilation disabled, locked/offline dependencies, and the coordinator memory
shim first on `PATH`. The exact source pin for every receipt is
`77bc2d3ac2f58808b153f50c8f76c1d2c552cde2`. Logs remain outside the checkout
under `(deleted-coord-artifact)`.

The shared environment prefix was:

```text
PATH=(deleted-coord-artifact):$PATH
CARGO_TARGET_DIR=(deleted-coord-artifact)
CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0
CARGO_PROFILE_TEST_DEBUG=0
TMPDIR=(deleted-coord-artifact)
```

| Check | Exact command after the shared prefix | Exit | External log (bytes, SHA-256) |
|---|---|---:|---|
| Signed-input policy matrix | `cargo test -p litchi-ops --features doc,xls --lib --locked --offline tests::signature::every_signed_input_follows_the_explicit_signature_policy -- --nocapture --exact --test-threads=1` | 0 | `every-signed-post-77bc.log` (629, `f1a378113a3989bd95336d22925bf482467c8a9575fc659ab2ab87793efa5a27`) |
| Signed XLS setter/cancellation | `cargo test -p litchi-ops --features doc,xls --lib --locked --offline tests::signature::signed_xls_sheet_setter_preserves_exact_noops_and_refuses_changes -- --nocapture --exact --test-threads=1` | 0 | `signed-xls-post-77bc.log` (417, `b5075110148f191eb5352a0edb8accc175a26278bcc1ce3867895cf7348c8677`) |
| Signed DOC setter/cancellation | `cargo test -p litchi-ops --features doc,xls --lib --locked --offline tests::signature::signed_doc_paragraph_setter_preserves_exact_noops_and_refuses_changes -- --nocapture --exact --test-threads=1` | 0 | `signed-doc-post-77bc.log` (421, `75e08b50c0e058bef2c8d9a16a628802959c4c369115e3ec2837f629d62507d5`) |
| DOC performance-diagnostics compile | `cargo check -p litchi-doc --features performance-diagnostics --lib --locked --offline` | 0 | `doc-perf-check-post-77bc.log` (170, `893e7f39edc501f5e9d27bb3aabcaef57ee306d8f1305a3f8d8c102d08c15f15`) |
| Profiled failed-open event balance | `cargo test -p litchi-doc --features performance-diagnostics --lib --locked --offline body_text::tests::profiled_open_matches_ordinary_error_and_closes_failed_phase -- --nocapture --exact --test-threads=1` | 0 | `doc-profiled-error-post-77bc.log` (517, `176d41ef579e4d656ad3ca758ff35898b2e1ac1a5f2155e779735070cb3a3383`) |

Additional source-limit regressions were also run against the same pin:

| Check | Exact command after the shared prefix | Exit | External log (bytes, SHA-256) |
|---|---|---:|---|
| Original entry/aggregate limits | `cargo test -p litchi-doc --lib --locked --offline body_text::tests::signed_fallback_preserves_original_entry_and_aggregate_limits -- --nocapture --exact --test-threads=1` | 0 | `doc-signed-limit-post-77bc.log` (414, `08d7b2fe2c4f22c729f2a17a971941e76a308de41ac2bba8a5f925a6b6eb778e`) |
| Profiled original entry limit | `cargo test -p litchi-doc --features performance-diagnostics --lib --locked --offline body_text::tests::profiled_signed_fallback_preserves_original_source_limits -- --nocapture --exact --test-threads=1` | 0 | `doc-profiled-signed-limit-post-77bc.log` (410, `295557140aaebeecc6ebdf391f15be6bd00c1a92b5b7bb15e8604340951045ba`) |

Final external Cargo target size: **778,307,116 bytes** at
`(deleted-coord-artifact)`.
No target contents, binaries, or evidence logs are tracked in this commit.

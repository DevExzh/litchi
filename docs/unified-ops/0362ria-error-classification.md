# 0362ria — replay and freshness classification repair

Status: repair complete on `uo2/0362ria-error-classification`.

The implementation commits are:

- `023d8d51b2` — DOCX recorded replay conflicts and OPC artifact-proof
  mismatches;
- `abb002bf3e` — iWork freshness and wrong-kind adapters.

The commits are based on `50122ae80a` (the coordinator integration tip at the
start of this repair). No raw iWork substrate API, package owner, replay
algorithm, output bytes, or redaction rule was changed.

## Classification repairs

The shared OPC classifier now reports the following distinct cases:

| Typed error reach | Unified code | Reason |
|---|---|---|
| `SourceArtifactMismatch { artifact: "current" | "original" | "source", .. }` before output | `invalid_input` | The caller supplied a recorded length or digest proof that does not identify the artifact being authenticated. DOCX durable current/source checks and exact-inverse current/original checks use this path. |
| Candidate proof mismatch before publication output | `invalid_input` | `write_to_with_expected_artifact` rejects the caller's expected candidate proof before publishing bytes. |
| Candidate proof mismatch after publication has accepted bytes, wrapped in `IncompleteOutput` | `io_error` | The outer error records partial output and must retain the output-stream failure semantics. `Classify for OpcError` classifies `IncompleteOutput` before inspecting its boxed source. |
| `SourceChanged`, either directly or boxed by `IncompleteOutput` | `io_error` | The positional source changed during the operation; this is source freshness/I/O failure rather than a bad caller proof. |
| Ordinary `IncompleteOutput` | `io_error` | Output was accepted before the publication failed, regardless of whether its boxed source is a sink, spool, cancellation, or freshness error. |

The candidate and post-output behavior is visible in
`crates/litchi-opc/src/source_backed/splice.rs:532-562`: the preview returns a
top-level candidate `SourceArtifactMismatch`, while the publishing pass wraps a
candidate mismatch in `IncompleteOutput` after accepted bytes. The common
classifier preserves that distinction at
`crates/litchi-ops/src/error.rs:1198-1205`.

The following paragraph records the pre-0362rga repair state and is retained
for engineering history. It is superseded by 0362rga and the current
canonical error contract: base mismatches use the shared `PatchConflict`, and
unsupported document-variable preservation uses
`Refused(UnsupportedContent)`. None of the former variant names below is
public in the current DOCX API.

Before 0362rga, DOCX's nine public recorded-base and authorization conflicts
were listed as mapping to `invalid_input` at
`crates/litchi-ops/src/error.rs:1294-1303`:

`ExternalHyperlinkDetachmentConflict`, `ExternalHyperlinkRedactionConflict`,
`DocumentVariablesConflict`, `SectionLayoutStaleSource`,
`SectionLayoutRevisionConflict`, `SectionLayoutForeignSource`,
`SectionLayoutFingerprintConflict`, `SectionLayoutPolicyConflict`, and
`SectionLayoutAuthorizationConflict`.

At that historical point, `DocumentVariablesPreservation` remained
`corrupt_file`; the redaction audit was deferred. The 0362rga repair supersedes
that classification for the current implementation: document-variable markup
that cannot be preserved is `Refused(UnsupportedContent)`, while recorded-base
mismatches are `PatchConflict` and classify as `invalid_input`. The historical
classifier tests remain evidence of the pre-0362rga state.

The iWork facade adapters now preserve source freshness in
`crates/litchi-ops/src/format/iwork/error.rs:375-403`:

- archive `SourceChanged` and `DirectoryChanged`, and detector `SourceChanged`,
  are `io_error`;
- Pages `UnexpectedFormat`/`NotOfficeFile`, Keynote `UnexpectedFormat`, and
  Numbers `UnexpectedFormat` are `unsupported_format`.

The focused adapter tests are at
`crates/litchi-ops/src/format/iwork/error.rs:760-816`.

## Validation receipt

The focused runs reused the coordinator integration target. They used `CARGO_BUILD_JOBS=1` and the 20 GiB memory cap from
`.coord/w8-part-a-integration/env.sh`; the target was
`/home/zhuhe/code/litchi-worktrees/unified-ops-api/target`. Those cargo
processes have terminated; the coordinator resumed its semantic run after the
target was released. Combined-source validation remains separate from these
worker receipts.

| Command | Result | External receipt |
|---|---|---|
| `source /home/zhuhe/code/litchi-worktrees/.coord/w8-part-a-integration/env.sh && cargo test --offline --locked -p litchi-ops --no-default-features --features docx --lib error::tests::docx_recorded_replay_conflicts_are_invalid_input --no-fail-fast` | 1 passed, 0 failed; exit 0 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ri-logs/docx-replay.txt`, 800 bytes, SHA-256 `b37878b171a04c7b4f37e3e0cc0edee90235fce14dc737d2865ce9c155f5059f` |
| `source /home/zhuhe/code/litchi-worktrees/.coord/w8-part-a-integration/env.sh && cargo test --offline --locked -p litchi-ops --no-default-features --features pages,keynote,numbers --lib format::iwork::error::tests::iwork_source_freshness_errors_are_io_errors --no-fail-fast` | 1 passed, 0 failed; exit 0 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ri-logs/iwork-freshness.txt`, 718 bytes, SHA-256 `dfe029916c542ef60255d2007f3fb7316563c53ea435ab4731b310fdbc7aaf33` |
| `source /home/zhuhe/code/litchi-worktrees/.coord/w8-part-a-integration/env.sh && cargo test --offline --locked -p litchi-ops --no-default-features --features pages,keynote,numbers --lib format::iwork::error::tests::iwork_package_wrong_kinds_are_unsupported_formats --no-fail-fast` | 1 passed, 0 failed; exit 0 | `/home/zhuhe/code/litchi-worktrees/.coord/0362ri-logs/iwork-wrong-kind.txt`, 729 bytes, SHA-256 `f8828112d5cb9f13cd3267729eacab49445a860869fadd482a55bc5f35bb3233` |

`rustfmt --edition 2024 --check` passed for the three changed Rust files, and
`git diff --check` passed. The staged repository size guard and binary guard
both produced no output. The full-workspace `cargo fmt --all -- --check` was
not used as a gate because this intentionally sparse checkout lacks unrelated
generated and corpus files; the changed files were checked directly.

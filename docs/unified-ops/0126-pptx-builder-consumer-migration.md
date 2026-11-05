# 0126 — PPTX builder consumer migration

## Changes

PPTX authoring in 23 root examples, two native examples, the unified write
fixture and five generators across three live performance source files now uses
`builder::Builder`. Authoring ends with `build()` before package publication or
OPC edits. Archive round trips used only to flush mutable authoring are removed;
reopens that verify saved output or establish physical source ownership remain.
A further empty-artifact SmartArt example now publishes through atomic native
`save`, as do the two migrated SmartArt authoring examples.

The builder gains feature-gated `font_embedding(mode)`, delegating native policy
admission before finalization. Font discovery and embedding failures propagate
from `build()`. No font policy is relaxed. Five section examples now supply
valid deterministic GUIDs instead of labels rejected by the existing validator.

The native legacy `Package::new` / `presentation_mut` authoring implementation
and internal fixtures remain pending migration. Empty-artifact `Package::new`
callers are intentionally retained for the eventual canonical constructor.
No live authoring call to `presentation_mut` remains in root/native examples,
the unified facade or performance harness sources. Historical evidence sources
were not rewritten. This is a consumer migration, not completion of the native
constructor split, the all-format glossary or the full goal.

## Validation and known failures

Both root and native example builds passed. Both native smoke examples passed;
21 of 24 root examples passed at runtime, including the password example with
its required environment input. The three remaining failures are detailed below.
Thirty-five generated plain PPTX archives passed ZIP integrity and content-type /
presentation XML parsing. These checks do not certify Office interoperability.
The 12 public unified PPTX write tests passed. Ten selected live harness tests
passed: semantic tiny corpora (1), source edit (1), cross-copy (1), marker shapes
(5), and slide boundaries (2). Cross-copy checks include deterministic output,
dependency closure, source immutability, durable patch replay and borrowed,
stale and foreign-source refusals. Independent read-only review found no builder
boundary regression and confirmed the separate codec/adapter failures.

All-target native/facade Clippy passed with warnings denied. The native builder
regression and 18 native doctests passed (two existing doctest ignores). Both
PPTX glossary profiles passed: minimal checked 72 banned names, full checked 63;
neither profile claims transaction-family verification.

Runtime artifacts and command logs live only under `/tmp/litchi-unified-0126-*` during verification;
no generated presentation, harness executable or raw evidence is committed.
The fixture tests verify semantics/determinism, not paired performance parity
or preservation of previous historical fixture digests.

Runtime checks identified three existing codec/adapter gaps, independent of
builder finalization:

- `pptx_handout`: handout serialization omits configured header/footer flags and
  text; reopened-value preservation assertions fail. The assertions remain.
- `pptx_handout_comprehensive`: its layout-token assertion fails because the
  handout codec does not emit that token. The assertion remains; print layout
  ownership and the example contract need a separate correction.
- `pptx_with_fonts`: Allsorts' PDF subset profile removes `OS/2`, which the
  PowerPoint EOT encoder requires. The subset example still returns the typed
  embedding failure. Changing the example to full embedding would hide this
  subset-path defect, so it was not done.

Four other examples initially failed on invalid section GUID fixture values
before builder finalization; those inputs were corrected. The encrypted example
requires `LITCHI_PPTX_PASSWORD`; its initial missing-variable failure was corrected
by supplying a test-only value for the rerun. No Office application visual/interoperability validation was performed.

Verification commands (offline, serialized):

```sh
CARGO_NET_OFFLINE=true cargo build -p litchi --no-default-features --features docx,xlsx,pptx,encryption,automatic-fonts --examples
CARGO_NET_OFFLINE=true cargo build -p litchi-pptx --features encryption,automatic-fonts --examples
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption --test pptx_unified_write
CARGO_NET_OFFLINE=true cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --target-dir target --lib semantic_docx_and_pptx_tiny_corpora_are_deterministic_and_editable
CARGO_NET_OFFLINE=true cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --target-dir target --lib pptx_source_edit_is_deterministic_and_emits_complete_evidence
CARGO_NET_OFFLINE=true cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --target-dir target --lib pptx_cross_copy_selectors_emit_separated_phase_evidence_and_gates
CARGO_NET_OFFLINE=true cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --target-dir target --lib marker_shape::tests
CARGO_NET_OFFLINE=true cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --target-dir target --lib pptx_slide_boundaries::tests
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption,litchi-pptx/automatic-fonts --all-targets -- -D warnings
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption,automatic-fonts --lib builder::tests
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption,automatic-fonts --doc
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-full --target-dir target
```

## Commit and cleanup

The rewritten `feat/office-format-completeness` tip remains an ancestor of HEAD,
so no rebase replay is necessary. Staging failed with exit 128 because Git
cannot create `/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`
on the read-only Git metadata filesystem. No commit was created. Required staged
evidence/binary guards and the changed/untracked artifact audit found no forbidden
files. No temporary worktree was created; review was read-only.

After every verification process terminated, this checkpoint's generated root
`Cargo.lock`, `target`, Python cache and `/tmp/litchi-unified-0126-*` scratch files
were removed. Tracked tool lockfiles and inherited changes were preserved.

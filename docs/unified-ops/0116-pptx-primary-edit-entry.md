# 0116 — PPTX primary edit entry

The previous checkpoint made verified progress on XLSX borders. This checkpoint
starts the PPTX write prerequisites required by ADR 0034; it does not enable
unified PPTX writes yet.

## Change

The primary native editor is now `opened::Edit`, entered through
`Package::edit()` or the existing `Snapshot::edit()`. Removed the old
`opened::Transaction` name and `opened_presentation_transaction()` entry without
aliases. Removed the primary editor's `rollback()` method: dropping the detached
edit discards staged changes. The package entry has a runnable example that
stages text, drops the edit, and verifies exact original package bytes.

Migrated native tests, facade editor storage and signature tests, the public
Office CRUD example, and performance harness consumers. The slide-boundary
harness retains the source snapshot before dropping the edit, preserving its
existing rollback verification. Updated the current native feature matrix and
harness documentation; historical evidence remains historical.

The `pptx-minimal` glossary profile now rejects all three removed names in their
primary-family scopes. Other formats and scoped PPTX transaction names are
outside this rename. The glossary describes the exact completed entry-point
scope instead of claiming full transaction-family alignment.

The PPTX-only configurations exposed existing feature
gating errors in origin persistence: an unnecessarily mutable eager result, an
unused decrypt-origin parameter, and DOCX/XLSX-only origin accessors compiled
without either feature. The unencrypted build additionally needed the stable-origin
helper gated to its actual consumers. Conditional shadowing and narrow feature gates fix
those configurations without changing their runtime behavior.

## Verification

- Native opened-editor suite: 193 passed, one existing ignored transcript
  generator, 725 filtered out; corpus revision and slide-root checks included.
- Signed-source and byte-order-marked-part integration suites: 22 passed.
- New `Package::edit()` drop-preservation doctest: one passed.
- PPTX/encryption facade tests and doctests: 516 passed.
- PPTX glossary: one profile, 62 banned names, no transaction-family compliance
  claim. Warning-denied all-target Clippy passed for native PPTX and the
  PPTX/encryption facade configuration.
- The standalone performance harness library and the public Office CRUD
  example compiled successfully. The harness lockfile includes the facade's
  current dependency set.
- Scoped Rust formatting and `git diff --check` passed. The audit of 487
  changed/untracked paths found no oversized evidence or compiled binaries.
  Both mandatory staged artifact checks printed nothing.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --lib opened
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --test opened_signature_policy --test package_signature_policy --test byte_order_marked_parts
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --doc package::model::Package::edit
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$PWD/target" cargo check --manifest-path tools/perf-baseline/Cargo.toml --lib
CARGO_NET_OFFLINE=true cargo check -p litchi --no-default-features --features docx,pptx,xlsx,ooxml-common --example office_crud_demo
```

## Remaining work

The existing signed-source refusal remains in the native primary editor.
Package-carrying commits, diagnostics, canonical patch introspection, immutable
`Package::apply`, authoring-writer removal, and source-backed editor publication
still require migration. The current mutable publishers return `Snapshot` and
cannot be renamed to the canonical immutable `apply` without changing their
ownership contract. Source-backed facade editing must retain its native source
owner and execution/freshness fences.

The next step is that publication/ownership migration, followed by PPTX
shape-text writes with retained native patches, atomic failure, exact no-ops,
inverse/replay, and preservation tests. XLSB and all later goal tiers remain
required. No full-workspace acceptance or performance claim is made by this
name migration. The rewritten base remains an ancestor of HEAD; no rebase
replay is needed.

Read-only review identified follow-up areas to exercise during the ownership
migration: preserve tolerance, physical-source provenance, encryption and font
state in the returned package; keep the candidate graph rather than replaying
a write-set-only patch onto a possibly changed caller; audit payload aliasing
and aggregate admission in patch capture; redact content-bearing debug output;
and validate failure atomicity of notes and other composite mutators. These are
review findings requiring focused regressions, not verified fixes in 0116.

The requested staging attempt failed: Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` because the
shared Git metadata filesystem is read-only. No commit was created. This does
not prevent further implementation in the writable checkout.

Final PPTX-only and no-feature facade checks passed, followed by another
warning-denied all-target PPTX/encryption facade Clippy run. Removed the generated
root Cargo.lock, build directory, Python cache, and checkpoint-specific temporary
logs. No temporary worktree was created; both read-only reviewers finished.

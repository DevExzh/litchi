# 0118 — PPTX committed package ownership

Checkpoint 0117 made verified progress on patch admission and payload ownership.
This checkpoint migrates the eager primary PPTX commit/publication contract
toward ADR 0034. It does not yet enable unified PPTX write verbs.

## Native contract

`opened::Commit` owns its candidate package and exposes `package()`,
`into_package()`, `patch()`, `diagnostics()`, and
`into_parts() -> (Package, Patch)`. `Diagnostics::changed()` and
`Patch::is_noop()` replace the primary editor/commit `is_changed()` and patch
`is_empty()` methods. The old public commit `snapshot()` and mutable
`apply_opened_presentation_commit/patch()` entry points are removed without
compatibility aliases. `into_patch()` remains a consuming convenience for callers
that only retain history: it drops the candidate without cloning it.
`into_parts()` is the canonical package-plus-patch extraction. Native white-box tests retain a test-only snapshot
accessor; it is absent from the public library API and rustdoc.

A changed commit moves its already-validated OPC candidate into the returned
package. An exact no-op uses the source OPC owner, preserving its exact source
archive. Neither route serializes/reopens or applies the result onto another
caller's graph. Commit clones share their package; consuming a shared clone
copies its OPC graph while retaining shared source buffers. Digest memos are
projected onto the actual candidate allocations.

A small owner context carries tolerance evidence, encryption mode, and optional
font state through snapshots and snapshot-derived edits. Physical-source
provenance remains explicit in the snapshot. The primary slide-copy,
cross-slide-copy, and slide-removal publication routes attach their destination
owner context before returning snapshots as well.

`Package::apply(&patch) -> Result<Package>` is immutable durable-patch replay.
It validates the write set and uses the destination's unrelated resources and
owner policy. Direct committed-package publication is therefore independent of
later source drift, while replay intentionally preserves unrelated destination
drift. The native replay builder returns its candidate directly, avoiding an
extra wrapper-level graph clone.

Changed signed replay refuses both signed sources and signature introduction
into an unsigned candidate. Exact no-ops retain signatures. Primary immutable
edits/replay retain encryption provenance: plaintext output remains refused,
and explicit retained-mode re-encryption remains available. Legacy scoped
mutation families keep their existing policies.

Debug implementations for primary packages, edits, commits, snapshots, slides,
patches, and conflict/delta wrappers expose counts and policy facts rather than
payloads, authored names, native identifiers, or part names. Adding safe package
Debug enabled the standard `expect_err` assertions in four existing read-limit
tests. Primary transfer-refusal and notes-readback details now omit authored
names, identities, and text; transferred-XML and patch-part-name parsing errors
use static details. Regression checks cover authored frame-name and malformed
part-name redaction. This is not a claim that all remaining native errors have
completed the glossary migration.

## Consumers and verification

Native tests, facade signature tests, the public Office CRUD example, the
performance harness, and native-resave consumers are migrated to the new
ownership contract. Registry/property capabilities do not change because no
facade write verb is enabled here. The glossary profile rejects six additional
superseded primary names.

The final encrypted native opened-editor run passed 210 tests, with one existing
ignored transcript generator. This includes six new owner regressions for
immutable replay/inverse, clone and snapshot owner state, later source drift,
content-free Debug, encrypted commit/replay/re-encryption, and signed replay/
signature injection. Existing byte-order-marked and signed-source integration
suites passed 22 tests. Native doctests passed 16 examples, including the new
commit/replay/inverse example; two existing examples remain ignored.
The facade PPTX/encryption run passed 516 tests/doctests. The performance
harness library, Office CRUD example, native-resave PPTX tool, and native
`automatic-fonts,encryption` feature build passed. The glossary profile passed
with 68 banned names; it still checks zero transaction families. Warning-denied native/facade all-target Clippy passed. `git diff --check`,
the two required staged-file guards, and the changed/untracked evidence and
binary audit passed.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --lib opened
CARGO_NET_OFFLINE=true cargo test -p litchi-pptx --features encryption --test opened_signature_policy --test package_signature_policy --test byte_order_marked_parts
CARGO_NET_OFFLINE=true RUSTDOCFLAGS='-D warnings' cargo test -p litchi-pptx --features encryption --doc
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-pptx -p litchi-ops --no-default-features --features litchi-ops/pptx,litchi-ops/encryption --all-targets -- -D warnings
CARGO_NET_OFFLINE=true python3 tools/check_api_glossary.py --profile pptx-minimal --target-dir target
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$PWD/target" cargo check --manifest-path tools/perf-baseline/Cargo.toml --lib
CARGO_NET_OFFLINE=true cargo check -p litchi --no-default-features --features docx,pptx,xlsx,ooxml-common --example office_crud_demo
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$PWD/target" cargo check --manifest-path tools/native-resave/Cargo.toml --no-default-features --features pptx
CARGO_NET_OFFLINE=true cargo check -p litchi-pptx --features automatic-fonts,encryption
```

## Remaining work

Source-backed commit ownership, new-package authoring without a mutable writer,
scoped PPTX families, and remaining selector/error vocabulary still require
migration. Source-backed facade editing must continue to use its native owner
and freshness/execution fences. The full transaction-family glossary gate is
not claimed yet. Unified PPTX shape-text writes follow these native owner
prerequisites. XLSB and every later format tier, Python/CLI, performance, and
acceptance requirement remain part of the active goal.

These are scoped correctness/API checks, not full-workspace acceptance or
paired performance evidence. Existing benchmark routes retain their identities;
current API descriptions and callers are migrated. The rewritten base remains
an ancestor of HEAD, so no rebase replay is needed.

## Commit and cleanup

The requested staging attempt failed with exit 128 because Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` on the
read-only shared Git metadata filesystem. No commit was created; the changes
remain in the worktree. No escalation or alternate repository was attempted.

All build processes completed before removing this checkpoint's `/tmp` logs,
generated root lockfile, build target, and Python bytecode cache. Tracked tool
lockfiles and inherited unrelated work are preserved. No temporary worktree
was created. The full goal remains active.

# 0121 — Shared immutable PPTX native promotion

The facade's internal PPTX adapter now returns an immutable native package
through `editable(&self)`. It no longer consumes the facade handle or requires
exclusive source ownership. This removes the old `into_editor`, `EditorRefused`
and `PptxEditor` handoff and its sharing-dependent refusal. The unused
`Opened::try_unwrap` / `rebuild` helpers are removed with that handoff.

The source route promotes through `pptx::source::Package::from_presentation`,
sharing the already admitted OPC catalog and keeping slide/media payloads
lazy. Successful promotion is cached behind an owner-local mutex; concurrent
callers receive native snapshots from the same cached package. Failed
promotion is not cached. Cache lookup and publication still enforce source
freshness and execution policy. Live facade readers and native slide handles
remain usable while an independent edit commits and replays.

The eager route shares its native package through `Arc` and retains native
edit admission, including refusal of decrypted encrypted edits. Explicit
`Signed::unsign` uses a separate cache of native unsigned packages. It cannot
replace the original source promotion or change the signature disposition of
another handle. Signed and `retain_source(false)` handles retain their typed
refusals. This is an internal ownership change; unified PPTX mutation dispatch
and capability advertisement remain pending.

## Verification

The PPTX/encryption facade suite passed 518 tests and doctests. Regressions
cover exact no-op output, no additional cold part loads at source promotion,
live readers alongside a changed native commit, repeated and concurrent source
promotion, shared eager owners, transient fence failure and retry before and
after successful promotion, cached cancellation and source-change refusal,
retained-source policy, and reuse of explicitly unsigned packages. Existing
signature tests verify that original signed handles still refuse after another
handle explicitly removes signatures and edits. Broader format coverage and
performance acceptance remain open.

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features pptx,encryption
CARGO_NET_OFFLINE=true cargo clippy -p litchi-ops --no-default-features --features pptx,encryption --all-targets -- -D warnings
```

Warning-denied facade all-target Clippy passed. The strengthened final signature
regression also passed. The PPTX glossary profile passed (68 banned names,
zero transaction families checked). Scoped formatting, `git diff --check`,
required staged artifact guards and changed/untracked artifact checks passed.
Read-only review found no concrete correctness or security blocker.
The rewritten base is still an ancestor of HEAD; no rebase replay is needed.

## Commit and cleanup

Staging failed with exit 128 because the shared Git metadata filesystem is
read-only: Git cannot create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit was
created. Changes remain in this worktree.

All build/check processes completed before removal of the checkpoint's
`/tmp/litchi-unified-0121-*` logs, generated root `Cargo.lock`, build target and
Python cache. Tracked tool lockfiles and inherited changes are preserved.
Review agents were read-only; no temporary worktree was created. The full goal
remains active.

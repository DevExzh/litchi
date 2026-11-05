# 0052 — File unified reads over the DOCX adapter

## Scope

Continue record 0051 with production `File::root/get/find/query/view/view_to/`
`view_json` dispatch and facade exports for `node`, `path`, `query` and `view`.
DOCX is the first concrete target. Signed files delegate to their retained
read owner, and locked files refuse before parsing paths or selectors. Other
formats retain their existing family read APIs and return a typed pending
capability error on this new surface. This is an intermediate implementation,
not an acceptable final whole-format capability gap.

Adapt the query ownership and forwarding portions of source candidate
`b9c6fd0bf4`; no candidate history is merged. A walk owns its adapter, whose
children borrow the retained document rather than a temporary adapter value.
Reference forwarding preserves source checks and supports the existing fake
adapters. `Matches` statically dispatches its iterator, reusable plan and
diagnostics; errors fuse it and zero limits touch no source payload. The
existing depth/visit bounds and raw-admission policy remain in place.

Native paths and query predicates still require explicit raw wrappers, and
returned DOCX paths remain positional. Typed paths and borrowed query plans use
the same public entry points. Node and iterator lifetimes remain tied to File.

Text, annotated, outline and statistics views are reachable through File with
default options. JSON views own their decoded projection, serialize directly
through serde without a JSON string round trip, and redact Debug output. A
boxed private format enum keeps the generic JSON handle small. Public typed view
options and their limits remain follow-up work; this record does not expose
the candidate's raw numeric option fields as a public API.

## Verification

Public integration coverage includes typed/text paths, path-not-found versus
invalid-path behavior, reusable borrowed plans, first/limit and diagnostics,
raw admission, positional round trips, signed-file reads, source changes before
lazy iteration and after caching nodes, zero-limit behavior, source/eager view
agreement, JSON serialization and sink errors. Locked-file tests verify error
precedence. A compile-fail example forbids an iterator outliving File.

- Operations all-feature checkpoint: 609 library tests, 22 integration tests
  and 132 doctests passed. Public integration tests were rerun with JSON views.
- DOCX-only/test-support: 345 library tests, 22 integration tests and 120
  doctests (including two compile-fail examples) passed.
- No-default-feature library tests: 223 passed.
- The facade DOCX integration test passed through `litchi::open` and the
  facade's exported path/query/node/view types.
- Warning-denied all-target ops Clippy, ops/facade rustdoc and the all-feature
  workspace library check passed. The latter includes Python bindings.
- API glossary: 11 profiles, 144 banned names, zero transaction families
  checked; passed.
- Crate boundaries: 66 packages, 267 internal dependency declarations and
  11 existing migration debt items; passed.
- Query smoke: 792 inputs, 748 compiled family plans; this is deterministic
  smoke over fake adapters, not a coverage-guided fuzz campaign.

No paired read/write performance acceptance or complete OfficeCLI parity is
claimed. Python runtime tests and generated capability deltas remain outstanding
for milestone closure; no Python API implementation changes are made here.

## Remaining work

The full goal remains active. Public typed view options, capability/registry
publication, all other format adapters, unified edits/batches, Python/CLI,
generated OfficeCLI deltas and paired benchmarks still need completion. The
existing facade compatibility views remain until their consumers migrate.
No iWork sources or base-checkout files are edited by this change.

## Commit and hygiene

Git staging remains blocked by the sandbox's read-only worktree Git metadata
(`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`). This is the
second consecutive turn observing that restriction, but implementation and
verification can still progress in the writable checkout. The goal is not at
an implementation impasse. No commit is claimed; the old partial index must be
refreshed before committing the complete working-tree changes.

No worktree is created. Build output, generated lockfiles and scratch logs are
removed after checks. Repository evidence-size and binary rules are checked.

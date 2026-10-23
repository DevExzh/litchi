# 0362ugd — DOC lazy edit entry

- Status: implemented on `uo2/doc-lazy-entry`, based on integration
  `7559f44aa975254bd96f6e28eee17366b0ebeb6a`; source, consumers, tests and
  this record are intended to land together.
- Scope: the DOC construction-cost contract. Opening an edit retains the
  immutable package and effective options only; it does not copy the complete
  CFB or scan/parse the document. The first operation that needs semantic
  state performs the existing bounded, source-fenced parse.
- Performance claim: none. Native timing is not part of this record; the
  change is qualified by source shape, fidelity/refusal tests and the scoped
  checks below.

## 1. Deferred shape

The primary body edit stores `Package`, `EditOptions` and an optional body
transaction. `body_snapshot` is materialized once on first body operation and
continues to use the existing source-version, limit, leniency, protection and
candidate-validation gates. Profiled opens retain their observers until that
same point.

The shared scoped-family base stores `Package`, options and a mutex-protected
optional exact byte snapshot. The first family read, mutation or commit
materializes the bytes; a dropped or untouched edit has no artifact-sized
allocation. Metadata, property-set, auxiliary-part, tracked-revision and
embedded-object edits use the same optional-inner pattern. Source-backed
paragraph edits retain only the source and position and resolve the paragraph
on `text`, `set_text` or `commit`.

No source/security/limit policy was bypassed. In particular, malformed and
encrypted inputs still fail at the first operation requiring their deferred
state, signed-source checks remain at publication, and failed setters do not
consume staged state. Existing no-op, inverse/replay and byte-fidelity paths
remain the publication authority.

## 2. Consumer migration

All changed fallible getters now propagate their family error (`Result`), and
all direct DOC tests and the legacy W5 probe use the new receivers. The
leniency, encrypted-metadata, source-fingerprint and caption rejection tests
were updated to assert the deferred failure boundary rather than treating
construction as a parse operation. The signed-container validation guard in
`body_text` remains in place for the original-CFB safety contract.

## 3. Validation

Commands ran with `/home/zhuhe/code/litchi-worktrees/.coord/doc-lazy-entry/env.sh`
(`--offline --locked`, one job, incremental/debug disabled, 20 GiB cap). Raw
logs are kept outside the repository:

| Check | Result | Raw log |
|---|---|---|
| `cargo test --offline --locked -p litchi-doc --tests` | library: 1,211 passed, 2 ignored; all integration targets passed | `.coord/doc-lazy-entry/litchi-doc-tests-final2.log` |
| `cargo test --offline --locked -p litchi-doc --lib --features performance-diagnostics` | 1,222 passed, 2 ignored | `.coord/doc-lazy-entry/litchi-doc-tests-profiled.log` |
| `cargo clippy --offline --locked -p litchi-doc --all-targets -- -D warnings` | passed for `litchi-doc` and its dependencies with the tracked workspace `clippy.toml` thresholds | `.coord/doc-lazy-entry/litchi-doc-clippy-final2.log` |
| `cargo clippy --offline --locked -p litchi-doc --all-targets --features performance-diagnostics -- -D warnings` | passed for the profiled callback/lifetime path and dependencies | `.coord/doc-lazy-entry/litchi-doc-clippy-profiled.log` |
| `rustfmt --edition 2024` on all changed Rust files; `git diff --check` | passed | terminal receipt |

Both strict Clippy commands used the pinned Rust 1.95.0 toolchain from
`rust-toolchain.toml`; no dependency or unrelated lint workaround is included
in this change. The profiled opening path receives the retained package
`SourceLimits` rather than default limits, and callback lifetimes remain
caller-scoped.

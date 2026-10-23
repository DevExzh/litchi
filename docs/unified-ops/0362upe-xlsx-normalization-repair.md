# 0362upe — XLSX normalization correction and receipt

- Date: 2026-10-23.
- Status: implementation receipt only. Coordinator integration/review is
  pending; native and performance qualification remain open, so this record
  makes no acceptance claim.
- Source: `a58dbaceb3726ae2ace3a10576f8ecb02d7adb9f` on
  `uo2/xlsx-normalization-repair`.

## Correction

The XLSX cell-normalization and join paths now preserve the distinctions that
the native endpoint needs:

- explicit empty creation (`Clear`) is distinct from an absent-cell no-op
  (`ClearIfPresent`), and shared-string identity is compared by exact SST
  index;
- rich-text comparison includes storage representation and run properties,
  so plain/rich transitions and same-text formatting changes are retained;
- unbounded setters keep their fast path and do not eagerly read source rich
  storage; bounded normalization performs the representation-aware proof;
- pending full styles, typed facets, removals, fills, hyperlinks, merges and
  staged no-ops are normalized against the effective state before admission,
  with shrinking replacement charges released;
- joins use compact normalization plans and move-based undo. All fallible
  checks happen before publication, so a refused join leaves both edits and
  their staged maps unchanged rather than cloning or partially mutating them;
- capture-ledger rollback and mixed bounded/unbounded joins retain the actual
  final semantic charge. Typed source/package/XML/invalid errors from staged
  census are preserved through `JoinError::source()` and classification,
  rather than being replaced by a synthetic resource-limit error.

These repairs cover the returned review cases for explicit-empty and shared
storage identity, rich-run fidelity, pending-style/remove interactions,
failed-admission atomicity, join normalization, typed error propagation and
the unbounded fast path. Earlier successors also corrected reversible-join
borrow handling and the structured join-error coercion before this pinned
source.

## Validation receipt

The clean frozen source was tested with the external capped target using this
exact command:

```text
env CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-xlsx-normalization-repair/.coord/signed-native-repair/target CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 LITCHI_MEM_MAX=20G PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH cargo test -p litchi-xlsx --lib --all-features --offline --locked -- --nocapture
```

Result: `1871 passed; 0 failed; 0 ignored; 0 measured; finished in 347.14s`.

- External log: `/tmp/signed-native-a58-all-features-xlsx-lib.log`
- Log size: `183735` bytes
- Log SHA-256:
  `48f1ce55dfffd9ac274fe352660b47d3f829d7cbe6f8864c4ba3f08cd3b2d8d1`
- External target: `/home/zhuhe/code/litchi-worktrees/uo2-xlsx-normalization-repair/.coord/signed-native-repair/target`
- Target size after the run: `821800822` bytes

The focused a58 receipts were retained outside the repository:

- join/source error path: `/tmp/signed-native-a58-focused-join-source.log`,
  SHA-256 `476d51e60d454970c3c538ff5f656cb83f1623c73fa489e5ed6c29207b862d64`;
- unbounded same-text rich fast path: `/tmp/signed-native-a58-focused-unbounded-rich-broad.log`,
  SHA-256 `39cc058ab9a3b5532bd5add927aa9922535d8da90169d7409d189b92d5c93edd`.

No build artifacts or binaries are committed. Native/performance review and
coordinator acceptance remain pending.

The coordinator verified the three log hashes and retained durable copies
with the same filenames under
`/home/zhuhe/code/litchi-worktrees/.coord/xlsx-normalization-final-a58/`.
These mirrors preserve the receipts independently of the temporary directory.
The source correction passed independent scoped review and is composed in
the diagnostic integration; final performance and complete gates remain open.

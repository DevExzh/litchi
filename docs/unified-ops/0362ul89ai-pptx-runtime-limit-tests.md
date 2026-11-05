# 0362ul89ai — PPTX typed limit test repair

Date: 2026-10-25

## Qualification failure

The frozen source base is
`ae6f32d202a04e3d6675fa431b2579c0d5f27192`.  The retained full-feature
qualification command was:

```text
cargo test -p litchi-pptx --all-features --lib --no-fail-fast --message-format json-render-diagnostics
```

Its external receipt is at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-ae6f32d202/fullfeatures-tests/`.
The 378,736-byte `0000.log` has SHA-256
`4924bcd19683868b0cc259024ceefdb93f3ba4ef0d3ec7d8653fcf564639fb6a`.
The run exited 101 after 1,356 passed, 23 failed, and 1 ignored test.  The
retained `checks.json` is 1,076 bytes with SHA-256
`25a4e9375b216cdc639b2ea553fe84ca4b370f489e298bc5ed9d03d3b5b8fea1`.
Cargo, build, rustdoc, and native commands were not rerun for this repair.

## Diagnosis and repair

The source tip is
`e412373e0e3bd95a761a753bbb8ffa240036fbe1`.  Only these two test-bearing
files changed:

- `crates/litchi-pptx/src/shape/designer/model.rs`
- `crates/litchi-pptx/src/tag/tests.rs`

The Designer failure expected `CodecLimitKind::Names`, but
`validation::validate_tag_text` measures the selected per-string byte
profile and returns `CodecLimitKind::StringBytes` with observed 3 and maximum
2.  The test now asserts that actual typed dimension and values.

The four programmable-tag edit assertions and the canonical-escape parsing
assertion expected `OutputBytes`.  They exercise `ensure_list_budget`, which
measures the serialized PresentationML wire size against the static
`MAX_PART_BYTES` physical XML ceiling and returns `CodecLimitKind::XmlBytes`.
`OutputBytes` remains the selected caller output-profile check in `write`; no
production limit mapping, default capacity, input/output bound, parser
behavior, or allocation path was changed.  All failed-edit snapshots and
unchanged-list/unchanged-value atomicity assertions remain intact.

## Handoff

The exact two-file patch from the frozen base to the source tip is retained
outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-runtime-89ai-e412373e0e.patch`.
It is 2,918 bytes with SHA-256
`9161bbee6b1b1c383f2bf3275b0d7db737d651ca19b1064d3296e980914f7370`.

Source-only checks passed: `rustfmt --edition 2024 --check` on both changed
Rust files, `git diff --check`, exact changed-path scope, no `3rdparty/`
references, and repository large-file/binary commit guards.

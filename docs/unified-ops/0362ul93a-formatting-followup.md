# 0362ul93a — follow-up full-workspace formatting repair

- Date: 2026-10-25.
- Source before this repair: `2106417c0de63417206c9553750e9fb3791d6427`,
  tree `4ff9ac0898c2100cd8bb8805ae637bc2e965fc85`.
- Status: source-ready for independent review; complete workspace formatting
  remains a coordinator-owned final check.

## Preserved failed full-format row

The fresh full-workspace formatter attempt is preserved at
`/home/zhuhe/code/litchi-worktrees/.coord/format-0362ul93-f5-actual-2106417c0d/`.
Its only row was the canonical `cargo fmt --all -- --check`, which exited 1
after 29.048958 seconds. The row is a failure receipt, not an acceptance
claim; its target stayed at 0 bytes and its post-exit owned/builders samples
were empty.

| artifact | bytes | SHA-256 |
| --- | ---: | --- |
| `00-workspace-format-full-attempt-actual-f5-00.log` | 15,383 | `89560d4116586b59e0fe088919a8792b5bf7300c7f589c0c2c20904aaef8c530` |
| `result.json` | 4,321 | `3ca28d7a6a8eef239d2f76da1e57a49e6c15e3154c091d5e030b1f45b40d3b80` |
| `rows.json` | 3,619 | `1f9e9cc84fcb6db0f47ecd8219ea0b2131536e7dc844798789167abe1afb4514` |
| `plan.json` | 6,845 | `07b1162c8a4ed84e7b4d7b4ce2186c36e4e84ab8d2eed7d1106dc8013de9d1da` |
| cargo target trace | 105 | `e93eab815065deb33ece25b877725eb922e4958bb0e3837ea8a6cdf3f6e8dd2a` |

The source proof in that receipt is clean at the stated commit and retains
the three approved role locks:

```text
Cargo.lock                         efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811
tools/perf-baseline/Cargo.lock     de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea
tools/native-resave/Cargo.lock     989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae
```

## Why the first scoped capture did not contain these paths

The earlier `0362ul86` capture was bound to source `a84f576138` and its
335-path sorted relative list, whose SHA-256 was
`6b53128af05235f8aef4bb89b0898fb743215b470af59acbd880fe0e09d2ebb5`.
All ten paths below existed at that source, but commit
`0f098a25d9` (`feat(numbers-wire): integrate bounded raw policy lane
(0362ul66)`) changed every one after that capture. The first formatter repair
therefore correctly used the old capture's scope; the later full-workspace
check exposed this post-capture drift.

The new sorted relative path list, with one LF after every path, hashes to
`d0500a0c30144648a8550106eb4dab748148d8d3126379ecae5d7e4d827546bc`.
It is disjoint from the earlier set. The combined known scope is 345 paths
and hashes to
`9457723fb5c428e5ab70eb1156a2c793b665a43d529014a3fad5db1b99106556`.

## New paths and formatter decisions

All ten paths use edition 2024 inherited from the workspace and the root
`rustfmt.toml` (SHA-256
`5b08a5db07e6acaee62319a09314b9302c8cbc0f8e6b88b5ee8695f6303383d3`) with
`skip_children=true`. The owning manifest mapping is:

| owning manifest | paths |
| --- | ---: |
| `crates/litchi-iwa/Cargo.toml` | 1 |
| `crates/litchi-keynote/Cargo.toml` | 1 |
| `crates/litchi-numbers/Cargo.toml` | 5 |
| `crates/litchi-numbers-wire/Cargo.toml` | 2 |
| `crates/litchi-pages/Cargo.toml` | 1 |

The exact paths are:

- `crates/litchi-iwa/src/numbers/editor/tests.rs`
- `crates/litchi-keynote/src/package/slide_table_cells.rs`
- `crates/litchi-numbers/src/cell/mod.rs`
- `crates/litchi-numbers/src/package/comments.rs`
- `crates/litchi-numbers/src/package/comments_reply_native.rs`
- `crates/litchi-numbers/src/package/table_cell_edit/tile.rs`
- `crates/litchi-numbers/src/package.rs`
- `crates/litchi-numbers-wire/src/lib.rs`
- `crates/litchi-numbers-wire/src/table_cells.rs`
- `crates/litchi-pages/src/package/table_cells.rs`

## Source-only repair and guards

Only the ten paths above were passed to direct rustfmt:

```text
rustfmt --edition 2024 --config-path rustfmt.toml --config skip_children=true <the ten paths>
```

The ten new paths and all 345 paths in the combined known scope pass their
explicit `rustfmt --check` invocations. A lexical audit found no changed
comments, string/character literals, identifiers or numeric literals;
punctuation differences are rustfmt's equivalent closure-brace and
match-arm formatting. `git diff --check` is clean. No Cargo, cargo fmt,
rustdoc, build, glossary, native, corpus or performance command was run for
this repair.

The commit guard covers only the ten source files and this record; the
repository size and binary guards pass. The final full-workspace formatter
run remains required after all source merges.

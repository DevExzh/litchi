# 0362ul89af — PPTX InkAction aggregate-limit assertion

Date: 2026-10-25. This is a required correctness repair, not a runtime,
corpus or performance acceptance. Only redaction audit 22 is deferred.

The isolated full-feature library capture started from
`ae6f32d202a04e3d6675fa431b2579c0d5f27192` (tree
`361e2d24d8c9d0cb2c838f109646b3ede9150545`) on
`uo2/pptx-combined-0362ul89`. The requested test failed with the exact
aggregate error:

```text
Limit { kind: Codec(RetainedBytes), observed: 1630, maximum: 815 }
```

The test had been migrated to expect `PayloadBytes`, although its two
distinct 815-byte targets pass the per-target limit and then correctly exceed
the package-wide retained-target limit. Shared targets remain charged once by
the target cache.

Commit `0f9839b007d47081d525c93bbe2d7974affbd58a` changes only
`crates/litchi-pptx/src/presentation/embedded/ink_actions/package_tests.rs`.
The assertion now requires `RetainedBytes`, observes the checked sum
`payload_len + payload_len`, and retains `maximum == payload_len`. No cache,
budget, atomicity, payload, or fidelity implementation behavior changes.

The retained terminal capture remains outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-ae6f32d202`.
It reported 1,356 passing tests and 23 failures; this source repair addresses
only the named InkAction assertion. Fresh qualification is required, and no
runtime result is inferred from this source change.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `fullfeatures-tests/0000.log` | 378,736 | `4924bcd19683868b0cc259024ceefdb93f3ba4ef0d3ec7d8653fcf564639fb6a` |
| `result.json` | 42,519 | `4b8e178cf9d63f55e1cc445c85e6d32ce32e49df2b70c91e84c9d978b3f7758e` |
| `fullfeatures-tests/checks.json` | 1,076 | `25a4e9375b216cdc639b2ea553fe84ca4b370f489e298bc5ed9d03d3b5b8fea1` |

The source patch is retained outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/0362ul89af-pptx-ink-actions.patch`:
1,096 bytes, SHA-256
`03d2e271a481f212fb7e89a07bf96c0c55e978019ef968831e024a4b8a84fe27`.
No Cargo, build or rustdoc command was run for this repair.

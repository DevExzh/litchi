# 0362ul89y — PPTX strict-Clippy cleanup

Date: 2026-10-25. This is a required source-preparation repair, not runtime,
corpus or performance acceptance. Only redaction audit 22 is deferred.

The retained qualification-v2 capture started from
`02b2489eada2bfcd3d1b9bc285b0ea0d3c7ae0d6` (tree
`05942a20a7599b39d23b407831dac81b600a61cd`) on
`uo2/pptx-combined-0362ul89`. Scoped formatting and the minimal all-targets
check passed. The strict row failed at
`cargo clippy --offline --locked -p litchi-pptx --all-features --all-targets
-- -D warnings`, with 76 indexed diagnostics. This worker owns 13 of them.

Commit `dd06bfdafb5d586ab4adc9d61819ad7d80cee69d` repairs only these seven
source files:

- `animations/codec/semantic.rs` and `laser/model.rs` remove redundant
  `must_use` attributes where the return type already carries the contract.
- `animations/mod.rs` and `notes/mod.rs` remove no-op struct updates after
  retaining every explicitly selected limit field.
- `notes/codec.rs` removes the redundant `must_use` on its `Result` return.
- `opened/model.rs` elides a lifetime whose relationship is already expressed
  by ordinary reference elision.
- `package/replay_index.rs` uses equivalent `?`, direct optional part
  references, and `or_default()` forms. These preserve lazy admission,
  bounded reservation and persistent path-copying behavior.

No public capability, typed limit, rollback path or allocation policy changes.
No runtime, corpus or performance result is inferred; a fresh strict-Clippy
capture and downstream qualification remain required.

The actual capture remains outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-v2-02b2489ead`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `clippy-diagnostics-index.json` | 13,168 | `8c08ff22b3d5942f141cc6e488944838ebf35aef6c6a5c7e476c6930c80ca6ce` |
| `02-pptx-strict-clippy-all-features-attempt-qualification-f5-02.log` | 59,454 | `bfe533c51350b5e9bfdeeb60ec6db7485f033679a4962e47d0264b14751785f6` |
| `rows.json` | 35,949 | `4e3db5c9fea5e6e15802195ead886716c161b0d3ea9f7645a9e0f417c31e921a` |
| `result.json` | 37,850 | `5878d7e3ed980656fefa0ddb83823d3fa077e3041ee8cde4cf058f0f092f1274` |

The source patch is retained outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/0362ul89y-pptx-clippy.patch`:
6,007 bytes, SHA-256
`c40f95414777d4b3c9e2bbfd29c631b7a11b36c8a927d72e4cc6a55a6227840c`.
No Cargo, build or rustdoc command was run for this repair.

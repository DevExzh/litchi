# 0362ul89ax — Harness formatter normalization

The actual eight-row compiler packet on `cc2ab3f79d` stops at formatting in
5.662 seconds with target size zero. No Cargo row executes. Root verifies the
terminal receipt and log. The only formatter delta is module declaration
ordering and one RTF closure layout in `tools/perf-baseline/src/lib.rs`.

Root reproduces that exact delta with the repository formatter:
`rustfmt --edition 2024 --config skip_children=true tools/perf-baseline/src/lib.rs`.
The normalized source preserves all function calls, arguments and behavior.
Before normalization, the isolated 89aw baseline has identical bytes. It must
receive exactly this normalization before any paired build.

Preimage and postimage remain outside Git in
`.coord/pptx-combined-0362ul89/format-89ax/`:

| File | Bytes | SHA-256 |
|---|---:|---|
| `lib.rs.before` | 3,088,486 | `aa74247cb8b1131eb5bb9c593ac17302de4dcedb220deb0f9c635f0779106429` |
| `lib.rs.after` | 3,088,522 | `bc99dc96f72de24ca00e3dbab687bb9b9d059edd59f392a7be1885015a0732ab` |

The 0362ul89-f5-consumer-v2-cc2 packet is retained at
`.coord/pptx-combined-0362ul89-f5-consumer-v2-cc2ab3f79d`. Its `rows.json`
is 28,747 bytes, SHA-256
`0a230a2d2df8a128a8b350cb27ca7e072de73eef3d06e98f657ff5c3a69a0cd6`;
`result.json` is 30,250 bytes, SHA-256
`63b79798ac571d25ff3832d230f17d123318642a6e40e086bb134bcbadb0bc21`.
No compiler/runtime/performance result is inferred. Fresh source-bound
qualification remains required after pending reviewed composition.

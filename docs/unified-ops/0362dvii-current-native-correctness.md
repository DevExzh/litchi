# 0362dvii — composed DrawingML correctness and source-preparation receipt

This receipt records the exact-source c796 candidate correctness handoff and
the source-only preparation of the four reviewed historical boundaries. It is
not a performance or ABBA qualification, and it does not close the pending
DOCX/OPC receipt-type source comparison or make a whole-goal claim.

The complete reduced receipt is external, at
`/tmp/dml-final-paired-baselines.Wz0XIZ/source-only-readiness-receipt.json`
(31,727 bytes, SHA-256
`57937d3eea590ef4d294c9b8f316fd7dd101a8b5efa0601000c2832ca7feafcc`). It
binds the source, tracked plan, adapter hashes, lock identities, all ten
preparation receipts, and the exact commands needed for a later paired build.

## Source and lock identities

```text
candidate source: c796d0acc24c7c6e6fd08980b9e31af6bb86b18b
candidate workspace Cargo.lock: 4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1 (139606 bytes)
historical workspace Cargo.lock: 52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6 (139591 bytes)
perf-baseline Cargo.lock: eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c (57227 bytes)
tracked plan: docs/unified-ops/results/0362dvie/qualification-plan.json
plan SHA-256: e0cb370e0399c072de9fe4b141ceff682b87d7936bec5332cedc9a181fc13505
tracked manifest SHA-256: a94e262d2fd2d9055de1036f2c87a31cc198142002bc27fec92f586eda2283e8
```

The detached baseline worktrees and revisions are:

```text
owner routes:  /tmp/dml-final-paired-baselines.Wz0XIZ/source-945
               945eb1d6b9ef6c56fa26c82e12758d38c1b9015d
raw routes:    /tmp/dml-final-paired-baselines.Wz0XIZ/source-2f7
               2f7db7ee7fa01893e9099217fc4243687d196f63
theme bulk:    /tmp/dml-final-paired-baselines.Wz0XIZ/source-996
               996651b6b19717a0192dfd4144bccd367fb2ea44
theme family:  /tmp/dml-final-paired-baselines.Wz0XIZ/source-11728
               11728db894181ffbcdbcea3ae0008d51fdd2d8af
```

Each source has only the reviewed adapter allowlist applied. The reviewed
paired manifest was expanded to
`/tmp/dml-final-paired-baselines.Wz0XIZ/expanded-manifest.json` (12,332
bytes, SHA-256
`15c320211b8880f9e466037f3a2c8c5ecdd1d6ab598e4dbafe7d64a319537cb1`). Its
ten profiles and ten distinct preparation-receipt paths passed the runner's
fail-closed manifest and receipt validation. No Cargo command was run for
these historical sources.

## Current native correctness

The c796 release binary was built with one Cargo job, the 20 GiB cap, the
candidate 4100 lock, and `--locked --offline`. The binary is external at
`/tmp/dml-c796-current-native-target/release/drawingml-raw-native` (11,371,984
bytes, SHA-256
`ba974b34b6d6c71eda6c9840aa8a422e14bf3b847d3f4c1a8c92930fe2b71fc3`). Ten
routes then completed with the reviewed native protocol: six raw routes and
four owner routes, each with three warmups, five blocks, twenty samples, and a
150 ms minimum leg. Independent assertions verified raw size/limit/refusal and
fidelity receipts plus owner commit/replay/inverse, stale conflict, mutation,
and unchanged-remainder semantics.

The detailed current receipt is external at
`/tmp/dml-c796-current-native-routes-20261024T172202Z/current-native-receipt.json`
(23,648 bytes, SHA-256
`38547057b3b9b5ef5ed7f232fbcf780b4eaa359cbb34e46f5ecfe25db24a1e22`). It
explicitly records `abba: not run`, `performance_claim: not claimed`, and
`paired_qualification: not run`.

The future paired build command is recorded in the external readiness receipt;
it uses the expanded manifest, exact 52ed/eac862 inputs, fresh external output
and target directories, `LITCHI_MEM_MAX=20G`, and `CARGO_BUILD_JOBS=1`. The
historical build and ABBA phases remain unexecuted here.

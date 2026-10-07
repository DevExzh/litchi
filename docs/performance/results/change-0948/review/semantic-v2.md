# 0948 semantic review v2 addendum

Status: **PASS**. This addendum binds the prior semantic review to the final
`candidate-build-v2` source snapshot; it does not redo `semantic-v1.md`.

## Prior review

- Report: `/home/zhuhe/litchi-performance-0948/review/semantic-v1.md`
- SHA-256: `6af2bce79c4f0e0cf963f6e75f4ca0144d0d09c43ce55ab5079218a7c8e5f26c`

## Final snapshot binding

- Input root: `candidate-source-v4`; frozen/build root: `candidate-build-v2`
- Base commit: `b6c5c80ffbcb60736d6c4c2f3b590838640008a9`
- Source tree: 10,789 files, 260,076,537 bytes
- Source tree digest: `b62c060f2a1da0af75982813357c7c7bf7ddbfeccefb2a799e435ce3363ea4c1`
- Build manifest SHA-256: `9897879f118d5ca143cd1f0f7bcd81626f4aa26b7587e27c80e1cce89e058220`
- Freeze receipt SHA-256: `de3d4354c55e8e29644aafe6ed15fb171929312f5ff32f04ab08ef3037466a38`

The manifest and receipt bind the complete copied tree and the eleven changed
source files reviewed in v1. Key final source hashes are: XML `audit.rs`
`29efd9aa99c9491a1e86cfc25f1bbb31b3da7b6cbfc4438a1638c15b18b95cee`, XML
`namespaces.rs` `2b24bf90377ba75db3a15eeaa53d010bf817b17b2741b8b20e6e377fd34bbdfe`,
OPC `lib.rs` `1a9c1cc6810b343aa5647aab9aa60e09adf14fc63b6cc0459196316ffac442a2`,
OPC `source_backed.rs` `571ec3e2e08d11626b5ce489c5d7de498445c1dd72441692306772c1893c3498`,
and XLSX `package.rs` `16983ebda28ace9081be68e6d6a84698f91e95b152e25481c348e6a14d288303`.

## v3 to v4 delta

The only source delta is `crates/litchi-xlsx/src/raw/worksheet/edit/package.rs`:
the two empty `Box::<[T]>::new([])` expressions were replaced with
`Vec::<T>::new().into_boxed_slice()`. The v3 hash was
`56572673773c85ed18fc0e6c75478c7958223a5bb8d054a01b3c7927a92c2e7b`; the
final hash is the one above. This is a type/compile-compatibility fix for
empty boxed slices. It preserves empty values and has no XML proof, OPC
ownership, metadata, or budget semantic effect.

## Disposition

The v1 PASS carries to the exact final snapshot. No new semantic blocker was
introduced. The v1 advisories remain: bounded namespace/search state can be
held briefly before a fallback decision, the 8 MiB namespace admission limit
is a conservative logical charge rather than a physical RSS guarantee, and
the existing Managed ownership copy route remains required for publication.

No Cargo/build command was run by this reviewer; compile/test results are
external quality evidence associated with the freeze, not this static review.

# Compression backend review after 0836

This is a read-only dependency and source review. It does not claim a new
build, benchmark, or adopted production change. The question is whether the
0836 compression profile measured a backend with runtime CPU dispatch and what
the compatible next investigation should use as its baseline.

## Result

The 0836 pinned profile already had `zlib-rs/std`, and therefore had zlib-rs
runtime feature detection. It did **not** measure the current root workspace
lock's `flate2 1.1.10` graph. The profile's nested probe workspace retained
`tools/perf-baseline/Cargo.lock`, which resolves `flate2 1.1.9` and `zlib-rs
0.6.7`; the 0836 build log records the same `flate2 v1.1.9` resolution.

This changes the next-step ranking. A runtime-dispatch experiment against the
0836 profile must not be described as enabling SIMD for the first time. The
first compatible investigation is to establish whether the current production
`litchi-opc` graph lost that dispatch when it moved to `flate2 1.1.10`, then
measure an explicit `flate2/runtime_detection` candidate if restoration is
worthwhile. Keep the zlib-rs backend, compression level, framing, staged I/O,
preservation behavior, and durability protocol fixed.

## Exact dependency evidence

The 0836 nested lock has these entries:

```text
tools/perf-baseline/Cargo.lock:591-597
flate2 1.1.9
  checksum 843fba2746e448b37e26a819579957415c8cef339bf08564fe8b7ddbd959573c
  dependency zlib-rs

tools/perf-baseline/Cargo.lock:2494-2497
zlib-rs 0.6.7
  checksum 34b31d188d9d685a4f9c7b46d6e36631b07058d2cfe190267adce54dc230bf12
```

The retained 0836 build log says `Compiling zlib-rs v0.6.7` followed by
`Compiling flate2 v1.1.9` (`commands/build-fp/output.log`, lines 108 and
112). A package-scoped read-only feature tree for that nested manifest reports:

```text
zlib-rs v0.6.7
├── zlib-rs feature "rust-allocator"
│   └── flate2 v1.1.9
└── zlib-rs feature "std"
    └── flate2 v1.1.9
```

The reason is in the normalized registry manifest for `flate2 1.1.9`: its
optional `zlib-rs` dependency explicitly requests `std` and `rust-allocator`
(`Cargo.toml`, lines 248-255). The root and `soapberry-zip` declarations only
spell `features = ["zlib-rs"]` with `default-features = false`; that spelling
does not by itself prove whether the selected flate2 release supplies
`zlib-rs/std` transitively.

The current root lock and package-scoped `litchi-opc` feature tree differ:

```text
Cargo.lock:1258-1265       flate2 1.1.10 -> zlib-rs
Cargo.lock:5803-5806       zlib-rs 0.6.7

zlib-rs v0.6.7
└── zlib-rs feature "rust-allocator"
    └── flate2 v1.1.10
```

`flate2 1.1.10` adds a `runtime_detection` feature that maps to
`zlib-rs?/std`, while its optional zlib-rs dependency itself requests only
`rust-allocator` (`flate2 1.1.10/Cargo.toml`, lines 72-75 and 244-248). The
current `litchi-opc` tree therefore has no `zlib-rs/std` edge. A whole-workspace
tree can show `std` through unrelated default-feature consumers, so the
package-scoped graph is the relevant one for an OPC production claim.

`allsorts` does not resolve this ambiguity: its `flate2_zlib-rs` feature maps
only to `flate2/zlib-rs`. The root comment intentionally disables allsorts'
C-zlib default so the workspace stays on zlib-rs (`Cargo.toml`, lines 70-74).
The 0836 perf target's package-scoped tree contains no allsorts edge.

Both `zlib-rs-0.6.7` and a newer `zlib-rs-0.6.8` are present in the local
registry cache, but every relevant lock above resolves **0.6.7**. The 0.6.8
cache must not enter this evidence or candidate.

## What `std` changes

In the locked 0.6.7 source, `src/cpu_features.rs` gates runtime checks on
`feature = "std"`: x86 checks call `std::is_x86_feature_detected!`, and the
AVX2/BMI2 result is cached atomically (lines 11-60). Without `std`, those
runtime checks return false unless all required target features are compile
time enabled. The dispatch sites include:

- `src/deflate/compare256.rs`, which can select AVX2/BMI2 for match comparison;
- `src/deflate/slide_hash.rs`, which can select AVX2/BMI2 for hash-table sliding;
- `src/adler32.rs` and `src/crc32.rs`, which can select checksum paths.

The level table remains scalar algorithm selection: level 6 points to
`medium::deflate_medium`, and that routine calls `longest_match`. Enabling
runtime detection does not select another level or strategy. It may reduce
helper-loop cost, but it does not remove the dominant algorithm itself from
the 0836 profile.

Relevant source hashes from the selected 0.6.7 registry tree are:

```text
Cargo.toml                                      3b6d5c17260ea4045a9911573ed30e71350d68718e5370918215f7dab67b98d3
src/cpu_features.rs                             05ecb3dcd40ef6bebca2850767a5f8f59488ebf0916e1534fe42946873212630
src/deflate/compare256.rs                       8665cb4170e55935bd386e0c76c24ee2977801e86f48a32c6c743942aeebaab4
src/deflate/slide_hash.rs                       3b75be5113161e1408c12556a7c0ce9ffaedb4ffe1013c1bc2ddc83405d433fd
src/adler32.rs                                   17479b18d0596299a83f38d84142b6ae868ffdeb5bec590abd9b2c3887713b2c
src/crc32.rs                                     2e5fde22a4fc26bd7a9e87a8716c7c7a381d8ee2da8763e8440c8cec0049681b
src/deflate/algorithm/mod.rs                    dca1e0ce47926948da214c73c40b18e7827e6fe8750d5390bd9334c6d61aa09e
src/deflate/algorithm/medium.rs                  28e7a9139aca14aeadd30d5a158e1c3c15f57fd6afa2ea57fd86aa36a0f28a51
src/deflate/longest_match.rs                     fd2897dc481b9bb3e3e37a36a3fdc1d1fab746eb6b3d87e74b6dabec87bff014
```

The selected crate archives hash to `843fba2746e448b37e26a819579957415c8cef339bf08564fe8b7ddbd959573c`
(`flate2-1.1.9.crate`) and
`34b31d188d9d685a4f9c7b46d6e36631b07058d2cfe190267adce54dc230bf12`
(`zlib-rs-0.6.7.crate`).

## Proposed next investigation

1. Treat 0836 as the `flate2 1.1.9` plus `zlib-rs/std` baseline. Do not
   attribute any difference from it to a newly enabled runtime-dispatch path.
2. Record a current package-scoped production baseline for `litchi-opc` under
   the root lock, including the resolved feature tree and exact artifact bytes.
   The current source graph uses `flate2 1.1.10` with only
   `zlib-rs/rust-allocator` for that package.
3. If current measurements show the missing dispatch is a regression, prepare
   a bounded candidate that adds only `runtime_detection` to the workspace
   `flate2` dependency and the direct `soapberry-zip` declaration. Keep the
   same zlib-rs 0.6.7 lock, and inspect the resulting graph before any workload
   run.
4. Gate the candidate on exact compressed-member and whole-archive bytes,
   semantic reopen, resource limits, cancellation behavior, and unchanged
   default durability. Compare the changed-Part owner stack and latency; a
   speedup must be attributable to the feature edge, not to level, strategy,
   framing, or source-preservation policy.

If runtime dispatch is already equivalent in the actual production target, the
next highest-ROI source-preserving work remains coverage of the existing
verified precompressed transfer path for eligible copied media. A changed
member still has no reusable source compressed stream, and 0837/0838 continue
to rule out weakening the durable inverse contract to avoid regeneration.

## Evidence identities

```text
0836 build-environment.json  a5386053ed8886db7acf6106643da22314e154364d5a8d11b67554e64d92a16d
0836 build-fp output.log     a14a8d844cb6d8ed96605231e65c555ffd9024e5cb54c0abf79529e5c3de9fb4
0836 build-baseline log      91fee6d0ac0d5d0da3aaf26ef8b17c70c2354cf9c0256dd937e6ada1a34d3883
0836 freeze-baseline.json    89fba9596a91c92b7d2d874361abf8600cf05339ca694adcc77bdf9c7e6b021f
0836 freeze-fp.json          3cea8f1c54a961d710b4b5b550dbf4bea6fcdc1fb97e90e62e285d9793b5912
tools/perf-baseline/Cargo.lock
                             a373fcbf868c5e74da4adc8de1c39249f3e900b69bf7e735aaa419fe85a2801b
root Cargo.lock              d80aa535dcf880a85fa31e92fe2e5ce78b91d93a55953ae734e513aed83e0061
```

The final line items are generated lock snapshots; the retained 0836 build
identity is the nested lock and build receipt, not the current root lock.

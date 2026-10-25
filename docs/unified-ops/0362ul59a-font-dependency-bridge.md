# 0362ul59a — pinned font dependency bridge

Date: 2026-10-25. Status: source and metadata validation only; no Cargo build,
native run, performance claim, or final acceptance.

This record covers `1d0a8b533286d65270c254ca6746d7d955f40604` (the explicit
Allsorts → flate2 → zlib-rs bridge) and
`a4f874413b84238c1c5b70ca59769a655f908013` (the workspace-membership repair).
Both commits are based on `21b845406d9582249a7eee3132850484663f431a` on
`uo2/font-dependency-bridge-0362`.

## Source provenance

`vendor/flate2` is the crates.io `flate2` 1.1.10 package copied from
`/home/zhuhe/.cargo/registry/src/index.crates.io-1949cf8c6b5b557f/flate2-1.1.10`.
The source comparison had 64 upstream files and 65 vendored files: the only
added file is `PROVENANCE.md`, and the only changed upstream file is the
generated `Cargo.toml`. Every Rust source, test, example, documentation and
license file is byte-identical. The comparison receipt is
`/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-candidate/flate2-source-proof.txt`
(SHA-256
`da0c8f27babf35b015b9bf52c366aba376d6b0cc748f43ed25cedf31602dfcd9`).

The generated manifest pins its optional `zlib-rs` dependency to the reviewed
0.6.8 path. Allsorts 0.17.0 and zlib-rs 0.6.8 retain their reviewed source
snapshots and licenses. Workspace and standalone manifests use explicit paths;
the workspace patch remains as a workspace-local fallback.

## Root lock comparison against approved 4100

The approved root lock is
`/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-baseline/old-4100-Cargo.lock`,
139,606 bytes, SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1`.
The final candidate lock is
`/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-candidate-exclude/root-Cargo.lock`,
139,189 bytes, SHA-256
`efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811`.
Both contain 586 package entries. Comparing every package name, version,
source and checksum gives exactly these differences:

| Package | Approved source/checksum | Candidate source/checksum | Version |
| --- | --- | --- | --- |
| `allsorts` | crates.io / `83b5535c25712f8ba509f2e9f134581e080ff8b6699f3fbf12bfda95d82fcb36` | local path / no checksum | 0.17.0 |
| `flate2` | crates.io / `6e634e2e0ebac1ee034020da1ca582e17ffe4e0f5e985823721e168928136dcb` | local path / no checksum | 1.1.10 |
| `zlib-rs` | crates.io / `b268e58e7c693d7c271f93ffc4ba3b380412554231c85bf61ca7af91042a4112` | local path / no checksum | 0.6.8 |

No other package name, version or registry checksum differs. The candidate
Allsorts lock entry additionally lists its reviewed direct `zlib-rs` edge; the
flate2 dependency names remain `crc32fast`, `miniz_oxide 0.9.1` and `zlib-rs`.

The preserved baseline sidecars are root `4100…`, perf lock SHA
`eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`, and
native lock SHA
`51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`.

The root regeneration started by copying the approved lock, then ran:

```sh
cp /home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-baseline/old-4100-Cargo.lock Cargo.lock
/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo metadata --offline --format-version 1 \
  --manifest-path /home/zhuhe/code/litchi-worktrees/uo2-font-dependency-bridge/Cargo.toml \
  > /home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-candidate-exclude/metadata-unlocked.json \
  2> /home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-candidate-exclude/metadata-unlocked.stderr
/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo metadata --offline --locked --format-version 1 \
  --manifest-path /home/zhuhe/code/litchi-worktrees/uo2-font-dependency-bridge/Cargo.toml \
  > /home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-candidate-exclude/metadata-locked.json \
  2> /home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-candidate-exclude/metadata-locked.stderr
```

The locked receipt is SHA-256
`2c9959c7bb2434bbe69715caa47eac9dfe14555589e30f1798c5d1a3bc8b0006`.

## Workspace membership and consumers

Before the exclusion, the bridge graph reported 71 workspace members and 554
packages; all three vendor packages were implicit members and the ignored lock
had 600 entries. After `exclude = ["vendor/allsorts", "vendor/flate2",
"vendor/zlib-rs"]`, default metadata reports 68 members and 526 packages, and
`--all-features --locked` reports 68 members and 586 packages. No vendor package
is a workspace member in either graph. The all-features receipt is
`/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-candidate-exclude/metadata-all-features-locked.json`
(SHA-256
`14cc9f12b5a3eecc3d5b5faa65c637981e76c1fd312275968a3a823940c06486`).

In the final workspace graph there is one path `flate2` package and one path
`zlib-rs` package, with no registry duplicate. `png`, `tiff`, `usvg` and `zip`
have four direct edges in total, all to the path flate2 package. This is a
workspace result only: an external root does not inherit the workspace patch.

The font-only standalone fixture is
`/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-standalone-candidate/Cargo.toml`.
Its locked metadata receipt has 108 packages and resolves `litchi-fonts`,
Allsorts, flate2 and zlib-rs all as path packages. Receipt SHA-256 is
`33db6132a0b70d668e665f90a4fc8ab3f8d1b67637c4e0c8736f47ec6227403b`; its lock
SHA-256 is
`9fb9eab0b6a4cd31747596afdcac8317931e5e19eb062f24cb7664d8ec00f6ed`.

A mixed standalone fixture without any workspace patch was generated at
`/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-standalone-mixed`.
Its lock has 183 packages, two flate2 entries (one path and one registry) and
two zlib-rs entries (one path and one registry). `png`, `tiff`, `usvg` and
`zip` point to the registry flate2 entry there. The metadata command was not
declared passing: offline Cargo stopped at uncached `bzip2 0.6.1`; exact
stderr is retained with SHA-256
`b6e88e5875bed5bc4c473379c7c8fa4ef06715f3476328c02ede7773fc74b9de`.

## Tool lock transitions

The approved tool locks were copied unchanged to the external validation
directory before each candidate regeneration:

* `tools/native-resave/Cargo.lock`: approved SHA
  `51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378`;
  candidate `/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-tool-consumers/native/candidate-Cargo.lock`, SHA
  `989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae`.
  Both have 243 package entries. The only identity transitions are flate2
  1.1.10 registry/checksum `6e634e2e0ebac1ee034020da1ca582e17ffe4e0f5e985823721e168928136dcb`
  → path, and zlib-rs 0.6.8 registry/checksum
  `b268e58e7c693d7c271f93ffc4ba3b380412554231c85bf61ca7af91042a4112` → path.
* `tools/perf-baseline/Cargo.lock`: approved SHA
  `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c`;
  candidate `/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-tool-consumers/perf/candidate-Cargo.lock`, SHA
  `de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea`.
  Both have 239 package entries. flate2 1.1.10 changes from the same registry
  checksum to path; zlib-rs changes from registry 0.6.7/checksum
  `34b31d188d9d685a4f9c7b46d6e36631b07058d2cfe190267adce54dc230bf12` to the
  reviewed path 0.6.8. This is a real historical lock transition, not a
  relabeled performance baseline.

Candidate locked metadata passed for both normal and `--all-features` runs.
Receipts are retained under
`/home/zhuhe/.cache/litchi-api-contract/font-deps-21b845-tool-consumers`:

| Consumer | Normal metadata SHA | All-features metadata SHA | Packages |
| --- | --- | --- | --- |
| native-resave | `b8844f9a77277aa4901a47cf4e57c887890678272047c372a24d1ce9c5fc031c` | `8b3207ace1df708dd8a804126dc32c741fdc733bb1d445d6a66cd59ee92f519b` | 67 / 243 |
| perf-baseline | `a00f3ac491cb1c708fe5cfbfe37e8f202c4be0c1aa182eabe19e0425768ac2fd` | `78de2588fc6159a04656368be6b8c0cb43381cbe1c3b6a3ca797f692b5745f4f` | 239 / 239 |

The all-features metadata command was:

```sh
/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo metadata --offline --locked --all-features --format-version 1 \
  --manifest-path /home/zhuhe/code/litchi-worktrees/uo2-font-dependency-bridge/tools/<consumer>/Cargo.toml
```

The normal receipt used the identical command with `--all-features` omitted.

The original tool locks were restored after each candidate run. No build or
native qualification was performed.

No final acceptance or global single-instance claim is made for arbitrary
external roots; only the font chain is guaranteed by its explicit path edges.

## Coordinator assembly

Root independently compares every package entry against the approved original
locks. Only the exact transitions above occur: three path substitutions in
the root lock, two in each tool lock, and the explicit performance-tool zlib-rs
0.6.7 to 0.6.8 transition. The Allsorts direct zlib-rs edge is the sole added
root dependency edge. All other package entries remain byte-for-byte equal
after TOML parsing.

The reviewed candidate locks are installed only in the isolated shared-font
assembly. The root lock remains ignored; the two tracked tool locks migrate
with the dependency manifests. Earlier captures and measurement receipts keep
their original locks. New builds and actual qualification must bind these new
identities, including the performance-tool version transition. This metadata
review does not establish build correctness or performance acceptance.

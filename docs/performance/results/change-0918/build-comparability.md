# 0918 candidate build comparability

The 0917 baseline and 0918 candidate probe executables were built from the
same standalone probe manifest and lock. The manifest hash is
`0e6f34cc57d56b2725034d4ed067e06a9a28d571f59348d0412c3c6966ea157d`; the lock
hash is `abd9f47a5aaa973a760cd61392de79c2b6143026c300dc1329831652aa54c06c`.
The probe manifest contains its own `[workspace]`, so the repository root
profile is not inherited. Its only explicit release override is
`[profile.release] debug = 1`.

The archived base tree, candidate tree, and current root all contain the same
`.cargo/config.toml` (SHA-256
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`); it
defines a lint alias and no release build settings.

| setting | baseline and candidate evidence |
|---|---|
| command/profile | `cargo build --release --locked --offline -j2 --manifest-path .../probe/Cargo.toml --bins` |
| Rust/Cargo | 1.95.0, rustc `59807616e`, Cargo `f2d3ce0bd`; ELF `.comment` is identical |
| debug info | explicit `debug = 1`; both ELF files retain debug sections and symbols |
| opt-level | omitted by the probe manifest; same Cargo release default (inferred, not separately logged) |
| LTO | omitted; the root workspace's `lto = true` is outside the standalone probe workspace, so the same Cargo release default applies (inferred) |
| codegen-units | omitted; same Cargo release default applies (inferred) |
| panic | omitted; the root workspace's `panic = "abort"` is outside the standalone probe workspace, so the same Cargo release default applies (inferred) |
| strip | omitted; both binaries visibly retain debug/symbol sections, consistent with the same unstripped default |
| incremental | omitted; same Cargo release default applies (inferred) |
| RUSTFLAGS / wrapper | absent in both recorded environments |
| toolchain override | both source-tree toolchain files hash to `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` and select Rust 1.95.0; the candidate final environment records rustup resolving the identically hashed repository path |
| `CARGO_PROFILE_RELEASE_*` and related profile environment | not recorded in either build receipt; the omitted profile fields and these environment variables remain a comparability limitation |
| affinity | baseline CPUs 14–15; candidate CPUs 16–17; this affects build placement, not release settings |

The candidate executable is smaller. Its `.text` and debug metadata changes
are consistent with removing the temporary `Vec` clone from the three stream
handoffs, but the size evidence does not establish that exact causal path. For
`doc_retention_probe`, the file size is 55,146,832 → 50,630,928 bytes
(`−4,515,904`, `−8.19%`). The `.text` section is 2,652,899 → 2,649,395 bytes
(`−3,504`), while the six DWARF sections shrink by 4,487,548 bytes. `.rodata`
grows by 384 bytes and `.eh_frame` by 20 bytes. The allocator binary shows the
same pattern: 55,150,520 → 50,634,624 bytes (`−8.19%`). This is a binary-size
observation, not a latency claim; capture measurements remain authoritative.

The baseline target's Cargo fingerprint directory was removed during earlier
cleanup, so its numeric Cargo profile fingerprint and original rustc argument
line are unrecoverable. The shared manifest/lock, matching `.cargo` config,
recorded commands and environments, identical ELF compiler comments, and
candidate fingerprint (`rustflags = []`, the same release profile hash for both
probe bins) are the available comparison evidence. The omitted profile fields
and unrecorded `CARGO_PROFILE_RELEASE_*` variables prevent treating that
evidence as a complete historical environment capture; the defaults shown above
are inferred from the manifests and toolchain.

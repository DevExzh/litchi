# vendored allsorts 0.17.0

This directory is a pinned source snapshot of the crates.io `allsorts` 0.17.0
package, copied from the local Cargo registry source identified by the package
name, version, and checksum in the approved workspace lock. The package is
licensed Apache-2.0; `LICENSE` is retained beside the source. The upstream
repository is `https://github.com/yeslogic/allsorts`. The checked allocation
adaptation is intentionally a reviewable patch on that snapshot; it does not
claim that the files below remain unmodified.

The vendored package is selected by the workspace `[patch.crates-io]` entry.
The budget changes in this branch are kept as a reviewable diff against this
source snapshot; the registry cache is never modified.

This repository ignores `Cargo.lock`, so this leaf deliberately does not add or
rewrite one. The patch keeps package version `0.17.0` and changes only the
source identity from the registry archive to this path. The integration owner
must regenerate the lock deliberately and record that source-identity change;
no prior registry checksum or performance receipt is relabeled here.

The zlib-rs bridge is likewise pinned to the crates.io `zlib-rs` 0.6.8 source
snapshot (registry source revision `0254e07204884bfc902d6ebb68c8b703d0ea49cf`,
checksum `b268e58e7c693d7c271f93ffc4ba3b380412554231c85bf61ca7af91042a4112`).
The upstream repository is
`https://github.com/trifectatechfoundation/zlib-rs`, and its Zlib license is
retained in `vendor/zlib-rs/LICENSE`. Only the public query exposing
zlib-rs's compiled inflate workspace size is added. The workspace patch changes
the zlib-rs source identity to this path when the `flate2_zlib-rs` feature is
used; the integration owner must record that lock/source transition separately.

Source snapshot checksums before budget changes:

| File | SHA-256 |
| --- | --- |
| `Cargo.toml` | `75d6fe2a55c818f05fdfdd9d46abd8b1fb2cbaf09f0f833ed79309767254f56f` |
| `LICENSE` | `b40930bbcf80744c86c46a12bc9da056641d722716c378f5659b9e555ef833e1` |
| `src/subset.rs` | `fabb447bec1910ab69d1ed31d85b475c5f029d16db849f7ded976264b6feb66c` |
| `src/tables/glyf/subset.rs` | `50c5c9ab5df0d491e73b1fed1d377d2c1364c0395fef9f6983888209fd7d3cd` |
| `src/cff/cff2.rs` | `6d80753cada75a6f704bfad0ab10377875c6ef36427137bdbfdb36f036be6e37` |
| `src/cff/charstring.rs` | `bf67c6c7354e525282f7e071581626aa656378a68bd06dd30bee9f1931be1922` |
| `src/binary/write.rs` | `c3bd46a29f0d1c20c90e2bb3252ba8dc3049544f8cad10be3a3de65de3e7ef38` |
| `src/binary/read.rs` | `9a4eba0d62abd7d27fbfe9bec6e24e3d8ed2d07b1f22e3d1cdb5ff1c2143ddbf` |
| `src/cff.rs` | `746b8b751cec1ddfc443c3f460329cb08711390c9511b789056f373f4283826c` |
| `src/cff/subset.rs` | `e8a6846ca09d9f9ab56b0e2f2b50bc405526343667b6b2db9b505d80a9fea76b` |
| `src/glyph_info.rs` | `241a3509af8c2ba168d7568d86295e9de281433c936c62d89f20623ec7dd3038` |
| `src/tables/cmap/subset.rs` | `bf94675faa103050da907047afdd1002154f1925c39dc2dde6d1ba26c5e34201` |
| `src/tables/glyf.rs` | `3b92d2a0ffa839b93dcbe21cbc067b0febe79adde206d804131cd7c6195c9d7b` |
| `src/error.rs` | `63b1a857400cec57876a3d8f252360bf8628a11144fff359ab3d3ca0848e1e17` |

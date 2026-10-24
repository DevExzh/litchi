# Apache POI fixture provenance (record 0257)

Files under `test-data/` here keep their path inside the Apache POI source
tree (<https://github.com/apache/poi>). License: Apache-2.0 (the repository
root `LICENSE`); the upstream notice is retained in `NOTICE`. Metafile, XLSB
and VBA fixtures are described in
`../libreoffice-core/METAFILE_PROVENANCE.md`.

Copied byte-for-byte on 2026-10-13 from the workspace's local, untracked
Apache POI checkout (no git metadata; pinned by SHA-256):

| Path | SHA-256 | Used by |
|---|---|---|
| `test-data/poifs/60320-protected.xlsx` | `deed30663d59c44ea36728f81c8428ed5ff85373b30f4e4c92974d0f8b3e5329` | `crates/litchi-crypto/tests/native_ooxml_encryption.rs` |
| `test-data/slideshow/42485.ppt` | `333ef0b2bf22158cd4ca8a8f650e534e3740437156cd47e05951ae08f03f58c9` | `crates/litchi-ops/tests/crud_matrix.rs`, PPT group (record 0289; copied 2026-10-14) |
| `test-data/slideshow/ppt_with_embeded.ppt` | `afa1dd5adc36765e6e40725a133b0330e5ad1febd1506ca93ceae8c5c5400b53` | `crates/litchi-ppt/tests/scoped_families.rs`, embedded-object edits, PPT lazy-native objects routes (record 0362ugja) |
| `test-data/slideshow/bug61881.ppt` | `ba2bf551e87f5fd7667dec52ac19005396bdf9af00f0006f9a247a9fede2e8ee` | `crates/litchi-ops/tests/crud_matrix.rs`, PPT table (record 0294; copied 2026-10-16), PPT lazy-native picture-bullets route (record 0362ugja) |
| `test-data/slideshow/slide_master.ppt` | `e9e2f53aa1b66b1d7121ee9bb9554b0cc47d7a0bff84abe38ffb5da129a1356c` | `crates/litchi-ppt/tests/scoped_families.rs`, title-master edits (two main masters and one title master; record 0357a; copied 2026-10-20) |

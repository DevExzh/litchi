# Apache POI fixture provenance

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
| `test-data/slideshow/42485.ppt` | `333ef0b2bf22158cd4ca8a8f650e534e3740437156cd47e05951ae08f03f58c9` | `crates/litchi-ops/tests/crud_matrix.rs`, PPT group (copied 2026-10-14) |
| `test-data/slideshow/ppt_with_embeded.ppt` | `afa1dd5adc36765e6e40725a133b0330e5ad1febd1506ca93ceae8c5c5400b53` | `crates/litchi-ppt/tests/scoped_families.rs`, embedded-object edits, PPT lazy-native objects routes |
| `test-data/slideshow/bug61881.ppt` | `ba2bf551e87f5fd7667dec52ac19005396bdf9af00f0006f9a247a9fede2e8ee` | `crates/litchi-ops/tests/crud_matrix.rs`, PPT table (copied 2026-10-16), PPT lazy-native picture-bullets route |
| `test-data/slideshow/slide_master.ppt` | `e9e2f53aa1b66b1d7121ee9bb9554b0cc47d7a0bff84abe38ffb5da129a1356c` | `crates/litchi-ppt/tests/scoped_families.rs`, title-master edits (two main masters and one title master; copied 2026-10-20) |
| `test-data/slideshow/42520.ppt` | `96b485b11ea966d17df68712c34a19507d25d7667c47d94a328742b90e7e53a1` | `crates/litchi-ppt/src/equation/tests.rs`, `crates/litchi-ops/tests/ppt_equations.rs`, an Equation Editor 3.0 object (`Equation.3`) in a compressed storage on slide 15, saved by PowerPoint (copied 2026-11-16) |

## Derived fixtures

| Path | Derived from (POI path, SHA-256) | Change | Used by |
|---|---|---|---|
| `test-data/spreadsheet/LIBRE_OFFICE-128382-0-small-media.xlsx` (SHA-256 `2cd8f7c73e7ef6b9f67cafa018df954bf114be647d0a981abde49a761e7e5715`) | `test-data/spreadsheet/LIBRE_OFFICE-128382-0.xlsx`, `e6dbe8f5da66804b39f17f02183a63005a7c418d636ad5d43f21b9d39f22e800` | the four pictures under `xl/media/` replaced by 8×8 images of the same format (1.8 MB to 28 KB); every other entry keeps its bytes and order. Its first sheet's drawing holds an ONLYOFFICE-written (`docProps/app.xml`: ONLYOFFICE 5.4.1) Office Math equation (`a14:m`) inside a text box paragraph | `crates/litchi-xlsx` worksheet-equation tests, `crates/litchi-ops/tests/xlsx_equations.rs` (copied 2026-11-16) |

# 0362ul47 — imgconv raw record limits

This leaf migrates the variable-size PICT, EMF, and WMF raw parser boundaries
to the existing `litchi_imgconv::Limits` profile.  The profile is now a
required argument at every public variable parser/record entry owned here;
fixed-width headers and scalar zerocopy helpers remain unchanged.

## Canonical API

* `PictParser::new(data, limits)` replaces the public `new_bounded` twin.
* `EmfParser::{new,new_compatible,from_owned,from_owned_compatible,new_lazy}`
  require `&Limits` and retain the selected profile for zero-copy iteration.
* `EmfRecord::{parse}` and `EmfRecordRef::parse_ref` require `&Limits`;
  `EmfRecordRef::to_owned` also checks the selected allocation ceiling.
* `WmfParser::{new,from_owned}` and `WmfRecord::parse` require `&Limits`.
* Allocating EMF point/count array helpers require `&Limits`; fixed-width
  scalar/header readers retain their existing signatures.  Borrowed DIB
  extraction also requires `&Limits` so a public variable fragment boundary
  cannot bypass input/selected-range admission.

The raw paths reject input/record bytes before copying or reserving and cap
record/path counts before allocation.  Owner converters pass their retained
`Options::limits`; convenience conversion functions select their documented
owner default only at the high-level facade boundary.

## Consumer and test migration

PICT, EMF SVG, WMF SVG, and the shared metafile codec were migrated in the
same source change.  Compatibility raster converters retain and forward the
parser-selected policy into their SVG playback path, so output/object/state
ceilings cannot silently fall back to defaults.  Focused parser and converter
tests cover exact input-size acceptance, record-count rejection, record-byte
rejection, allocating path-array boundaries, and compatibility output-limit
refusal. No `3rdparty/` fixture is referenced. This record makes no
performance or full-workspace gate claim.

## Coordinator validation and integration (2026-10-25)

Independent source review accepts
`e392867ef9390982a0e6bdd5f810d7ade15e7ae3`, including retained EMF/WMF
compatibility playback limits. The final atomic commit
`aaeaa79a0bc0689e87069c78c4c87726c7bc8523` has the identical complete tree
`0416f04405b4c9209faf7195a1fe3cd79f095b56` and is based on 30aa. Root verifies
this equality and the terminal validation artifacts.

Actual capped validation is retained externally at
`/home/zhuhe/code/litchi-worktrees/.coord/imgconv-raw-limit-validation-corrected.9FTSYC/`.
The approved f5 runner has SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`;
the unchanged root lock is approved `4100b1fb…bfaea1`. The source overlay's
imgconv `src` is a symlink to the frozen, clean leaf, whose HEAD/status the
runner checks before and after each command.

On exact e392 source, both compatibility-output-limit regressions pass in
5.914645 seconds, and all-features/all-targets imgconv clippy with warnings
denied passes in 6.834357 seconds. The final target remains below 1.54 GB,
live guards report no overflow, and post-exit observations contain no owned
processes.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `08-imgconv-compatibility-limits-final.receipt.json` | 14,773 | `7a0fc185a1ea61ee886297dc9601e6e0c27a995f520a26dc810f645e1a0ac461` |
| `08-imgconv-compatibility-limits-final-attempt-imgconv-final-e392-8-1792887816418233494.log` | 682 | `8af49ebee1e9edd92ba381c0b8c54b9d2dc3a9e95706fb62ab2a659cbe94cf70` |
| `09-imgconv-clippy-all-targets-final.receipt.json` | 18,648 | `4290d08101d9c9c891c0d175dbc1fe508e925a8799cd5e66bbc78ac7969bb9cf` |
| `09-imgconv-clippy-all-targets-final-attempt-imgconv-final-e392-9-1792887822584487933.log` | 219 | `8c17eabf5561dae044f2c16f55de341ced06ee8d3ffbae8d9574b9da5730d04b` |

The earlier corrected-source focused parser/array tests, imgconv check,
facade check and facade clippy pass on `5a46333b64812aab2b05a5d4e93aee5cf922f45d`.
Its final facade clippy receipt is 7,427 bytes, SHA-256
`b3c6627c6f1868925acffe4da288116547a70bed39a55811f09a0ea69a4f5ecb`.
Those receipts keep their original source attribution. The final difference
repairs test initialization lint and restores original renderer helper
signatures; compatibility production paths are unchanged. The restored
renderer implementations are byte-identical to the lane base and pass the
final imgconv clippy check.

Initial lint failures remain preserved. Missing tracked `clippy.toml` caused
the first renderer diagnostics; the correct configuration is SHA-256
`f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c`.
The later four test-initialization warnings are corrected before the final
pass. No lint policy is weakened. The composed glossary policy parses with
all XLSX drawing and imgconv signature rows retained.

# Upstream aggregate observations

`cached-results.json` retains 48 selected numeric observations across the
seven aggregate functions in this batch: `SUM`, `PRODUCT`, `SUMSQ`,
`SUMPRODUCT`, `SUMX2MY2`, `SUMX2PY2`, and `SUMXMY2`. The observations come
from the LibreOffice mathematical and array FODS fixtures at commit
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`. `provenance.json` records the
raw GitHub source URL, every input SHA-256, the source coordinates, and the
profile exclusions.

The extractor reads only the seven selected source files and bounds every
reference closure before emitting it. Each referenced cell is retained as a
literal number, text, logical, or empty cell. A reference to an upstream
formula cell is refused, so the Rust test never evaluates a source formula or
silently imports its cached result as an input. The selected closures are at
most 14 cells.

The retained receipt was generated from a temporary source tree populated
with the exact raw files at the pinned commit. Reproduce it, including the
download and byte checks, with:

```text
python3 reproduce.py
```

`reproduce.py` downloads all seven raw files into a bounded temporary source
tree, verifies each SHA-256 against `provenance.json`, regenerates both JSON
receipts, compares them with the retained files, and removes the temporary
tree on exit. It does not modify a checkout. To extract from an independently
prepared checkout instead, run `extract.py` with that source root; all seven
files must match the retained hashes.

The repository's auxiliary `3rdparty/libreoffice-core` checkout is not the
receipt source for this batch: its `sumproduct.fods` is a later Collabora
editing variant (`bc09d1f31ad0d68b6e45577f3ab11dd87a0d7abb2e143cfae9d85317821ab14d`,
481,386 bytes), while the pinned raw fixture is
`3dec3d33ca5425d883dae35a6c51bc0784e8929c029929793f5a1d9cf273ee60`
(457,698 bytes). The other six local files match their pinned raw hashes.
The local file is preserved and never overwritten.

`PRODUCT()` at `product.fods` row 12 is retained as an explicit arity-profile
exclusion because the selected profile requires at least one argument.
`SUMPRODUCT()` at `sumproduct.fods` row 15 is excluded because its source
range includes `M17`, an upstream formula cell. Row 16 is also retained as an
explicit conversion variance: its `N17` source is the literal text
`"Unknown"`, which the LibreOffice cache treats as zero while the selected
contract reports malformed text in a referenced forced array as `#VALUE!`.
The exclusion receipt preserves that formula, cache, coordinate, and source
cell detail. None of these exclusions is treated as a supported result.

The public test compares the finite numeric caches with relative tolerance
`1e-13`, requiring exact positive zero. This is bounded cached-value
corroboration, not LibreOffice application execution, file resave or
acceptance, full-fixture evaluation, recalculation, or complete OpenFormula
conformance. Upstream fixture data is distributed under the retained MPL 2.0
license.

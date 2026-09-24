# LibreOffice descriptive-statistics observations

This directory retains bounded native observations for `AVEDEV`, `DEVSQ`,
`GEOMEAN`, `HARMEAN`, `KURT`, `SKEW`, and `SKEWP`. The source files are direct
FODS fixtures from the pinned LibreOffice core revision
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`; no workbook conversion,
recalculation, or formula-cache import is used.

The extractor retains 39 finite numeric formula caches across all seven
functions. A retained reference closure contains only literal number, logical,
text, or empty cells. Formula cells, named ranges, multi-area unions, and other
unsupported closures remain explicit exclusions in `provenance.json`; the
extractor refuses to promote any such closure.

The Rust integration test reconstructs each literal closure and evaluates its
formula in scalar and matrix value modes. It compares finite values with a
relative tolerance of `1e-13` and exact bit equality for zero. This is bounded
cross-checking against upstream cached observations, not a claim of complete
LibreOffice conformance.

Reproduce the receipt, including bounded raw downloads, SHA-256 checks,
temporary materialization, and byte-for-byte output comparison, with:

```text
python3 reproduce.py
```

The helper removes its temporary source and output trees on exit and never
modifies a LibreOffice checkout. Upstream fixture licensing is retained in
`LICENSE-MPL-2.0.txt`.

# LibreOffice paired-statistics observations

This directory retains bounded native observations for `CORREL`, `COVAR`,
`PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, and `FORECAST`. The source
files are direct FODS fixtures from the pinned LibreOffice core revision
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`; no workbook conversion,
recalculation, or formula-cache import is used.

The extractor retains 42 finite numeric formula caches across all eight
functions. A retained reference closure contains only literal number, logical,
text, or empty cells. Formula cells, named ranges, invalid probes, a
reference-valued forecast argument, and other unsupported closures remain
explicit exclusions in `provenance.json`; the extractor refuses to promote any
such closure. Forty-one rows use the ordinary native comparison policy. The
`STEYX([.I2:.I101];[.J2:.J101])` row is retained as an explicit host numeric
deviation: the independent proof computes `1.6021504609570352`, while the
LibreOffice cache is `1.60215046095659` (2005 binary64 ULP lower).

The `INTERCEPT` rows intentionally include nonzero native intercepts, including
`6.75000000513889` and `-9`. ODF 1.4 §6.18.38 describes `INTERCEPT` using
`LINEST(Data_Y, Data_X, FALSE())`, while §6.18.41 defines `Const=FALSE` as a
zero intercept. The native cache therefore records LibreOffice behavior and
does not resolve that normative contradiction. The local normative member is
`3rdparty/specs/OpenDocument-v1.4-os.zip!/part4-formula/OpenDocument-v1.4-os-part4-formula.html`
(the archive and member SHA-256 values are recorded in `provenance.json`). The
paired-statistics contract must state its chosen interpretation separately.

The Rust integration test reconstructs each literal closure and evaluates its
formula in scalar and matrix value modes. It compares the 41 ordinary rows
with the relative tolerance of `1e-13` and exact bit equality for zero. The
STEYX deviation is checked against the exact reference bits and the retained
host difference recorded by `steyx_deviation.py` and
`steyx-deviation-proof.json`. This is bounded cross-checking against upstream
cached observations, not a claim of complete LibreOffice conformance.

Reproduce the receipt, including bounded raw downloads, SHA-256 checks,
temporary materialization, and byte-for-byte output comparison, with:

```text
python3 reproduce.py
```

The focused numerical proof can also be checked directly:

```text
python3 steyx_deviation.py --check
```

The helper removes its temporary source and output trees on exit and never
modifies a LibreOffice checkout. Upstream fixture licensing is retained in
`LICENSE-MPL-2.0.txt`.

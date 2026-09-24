# LibreOffice text-function observations

This directory retains bounded native observations for all 26 normative
OpenFormula 1.4 §6.20 text functions: `ASC`, `CHAR`, `CLEAN`, `CODE`,
`CONCATENATE`, `DOLLAR`, `EXACT`, `FIND`, `FIXED`, `JIS`, `LEFT`, `LEN`,
`LOWER`, `MID`, `PROPER`, `REPLACE`, `REPT`, `RIGHT`, `SEARCH`, `SUBSTITUTE`,
`T`, `TEXT`, `TRIM`, `UNICHAR`, `UNICODE`, and `UPPER`. The source files are
direct FODS fixtures from pinned LibreOffice core revision
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`.

The receipt retains 91 typed native formula caches: 68 text values, 18 finite
numbers, four logical values, and one `#VALUE!` error. Every selected formula
has a bounded dependency closure containing only literal number, logical, text,
or empty cells. The raw source SHA-256 manifest and the exact native cache
attributes are in `provenance.json` and `cached-results.json`.

`DOLLAR`, `FIXED`, and numeric `TEXT` rows use the deterministic en-US
formatting profile represented by the selected cache. Date serials, Japanese
locale variants, formula-dependent or computed arguments, nested calls,
invalid-arity rows, and other incompatible host cases remain explicit
exclusions in `provenance.json`. Missing or untyped native results are not
replaced with synthetic values.

The Rust integration test reconstructs each retained literal closure and
evaluates the formula in its declared scalar or matrix value mode. Native
matrix-span rows are matrix-only and their complete output shape is checked;
the retained top-left cache is compared as the corresponding first element.
Ordinary formulas are checked in both modes. Formula cells never enter the
resolver. This is bounded cross-checking against upstream cached observations,
not a claim of complete LibreOffice conformance.

Reproduce the receipt, including bounded raw downloads, SHA-256 checks,
temporary materialization, and byte-for-byte output comparison, with:

```text
python3 reproduce.py
```

The helper removes its temporary source and output trees on exit and never
modifies a LibreOffice checkout. Upstream fixture licensing is retained in
`LICENSE-MPL-2.0.txt`.

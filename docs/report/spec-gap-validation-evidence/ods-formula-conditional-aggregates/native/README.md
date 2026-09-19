# Upstream conditional-aggregate observations

`cached-results.json` retains 64 selected numeric observations across
`SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, and `AVERAGEIFS` from
LibreOffice sources at commit
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`.  The exact raw input hashes,
coordinates, cached values, converter receipt, and profile exclusions are in
`provenance.json`.

The five ordinary FODS inputs contributing selected observations are read
directly.  Two supplemental wildcard/regex FODS inputs are also staged and
hashed for explicit exclusion provenance.  The upstream revision has no
SUMIFS FODS fixture, so the selected SUMIFS observation is taken from the pinned
`sc/qa/unit/data/xls/opencl/math/sumifs.xls` workbook.  `extract.py` converts
that exact raw workbook to FODS in a temporary directory using the recorded
LibreOffice 26.2.5.2 binary, then retains one bounded `Main` formula and its
literal `Data` closure.  The conversion is part of the receipt and is checked
by version and `soffice.bin` SHA-256; no checkout is modified.  The selected
SUMIFS value is the value emitted by that pinned conversion.  This evidence
does not claim that the original BIFF8 cached value was preserved byte-for-byte
or that conversion did not recalculate the formula.  The SUMIFS closure
contains 40,003 literal cells because its native source formula spans 10,000
rows.  This larger closure is intentional and remains below the extractor's
per-reference bound.

Reproduce the receipt, including raw downloads, SHA-256 checks, temporary
materialization, and byte-for-byte regenerated JSON, with:

```text
python3 reproduce.py
```

The helper removes its temporary source and conversion trees on exit.  It
requires the exact converter recorded in `provenance.json`; a different
LibreOffice build is refused rather than silently changing the native cache.

Rows using LibreOffice regular expressions, wildcards, inline case controls,
or host case behavior remain explicit exclusions.  Empty-side relational
criteria, invalid arity, inline-array range inputs, stale cached formula
attributes, and unmaterialized source closures are also retained as
exclusions.  Those rows preserve useful host observations without promoting a
host profile choice to the selected evaluator contract.

The Rust integration test reconstructs only literal number, logical, text,
and empty cells.  It refuses any formula-cell input, evaluates each retained
formula through the value resolver, and compares finite caches with relative
tolerance `1e-13`, requiring exact zero.  The test is bounded cached-value
corroboration, not full LibreOffice conformance, complete fixture recalculation,
native Office acceptance, or a general application round-trip claim.  Upstream
fixture data is distributed under the retained MPL 2.0 license.

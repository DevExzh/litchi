# Upstream statistical-reducer observations

`cached-results.json` retains 54 selected numeric observations across
`COUNT`, `COUNTA`, `COUNTBLANK`, `AVERAGE`, `AVERAGEA`, `MIN`, `MAX`, `MINA`,
and `MAXA` from the LibreOffice FODS fixtures at commit
`d804d6aff49054bad1719ec3c2d136b545bbc7e7`. The exact raw input hashes,
coordinates, caches, literal dependency closures, and explicit exclusions are
recorded in `provenance.json`.

Every selected source is a direct FODS fixture. No workbook conversion,
LibreOffice recalculation, or evaluator-generated golden is involved. The
extractor stages nine raw files, bounds each download to 2 MiB, checks every
SHA-256, and refuses formula cells in a retained closure. It accepts direct
text and a single direct `text:a` hyperlink as literal text while recording
the nested markup in the JSON cell record; other nested markup, dates, named
ranges, and nested formula dependencies remain explicit exclusions.

Reproduce the receipt, including raw downloads, hash checks, temporary
materialization, and byte-for-byte regenerated JSON, with:

```text
python3 reproduce.py
```

The helper removes its temporary source and output trees on exit. It does not
modify the repository or a LibreOffice checkout. A different upstream byte
revision is rejected by the pinned input hashes.

The Rust integration test reconstructs only the literal number, logical,
text, and empty cells retained in each closure. It evaluates each formula in
both scalar and matrix value modes where the selected formula is valid, then
compares finite native caches with relative tolerance `1e-13` and exact zero.
This is bounded cached-value corroboration, not full LibreOffice conformance,
fixture recalculation, application round-tripping, or a claim that host date,
error, named-range, or nested-expression behavior is normative. Upstream
fixture data is distributed under the retained MPL 2.0 license.

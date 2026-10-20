# ODS conditional-aggregate native observations (record 0358)

`observations.json` is a lossless re-encoding of the 64 selected LibreOffice
conditional-aggregate observations (`SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`,
`AVERAGEIF`, `AVERAGEIFS`) that
`docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/cached-results.json`
held before the 2026-10-05 history rewrite purged it (5,510,619 bytes,
SHA-256 `8a93522797d5f3cd6a61d13305d15bd2f77a4bf7057f12d96d65ab908174e369`).
Every field is kept; only each observation's literal cell closure is stored
as row-major blocks instead of one JSON object per cell (`compact.py`
documents the format, rebuilds the file and asserts the decoded closure
equals the original).

- Upstream: LibreOffice core (<https://git.libreoffice.org/core>) at commit
  `d804d6aff49054bad1719ec3c2d136b545bbc7e7`; the source fixtures, their
  SHA-256 hashes, the converter receipt and the exclusions are in
  `docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/provenance.json`,
  and the extraction is `extract.py`/`reproduce.py` beside it.
- License: MPL-2.0; the retained license text is
  `docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/LICENSE-MPL-2.0.txt`.
- `observations.json`: 352,876 bytes, SHA-256
  `a2c80e2295669bbef8c771e7623fa31d2d258eebbfd58052a2b55810c8256b3d`.
  Regenerate with `python3 compact.py <cached-results.json> > observations.json`.
- Used by: `crates/litchi-ods/tests/ods_formula_conditional_native.rs`.

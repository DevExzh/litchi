# 0362dvec — DrawingML chart raw review return

Status: repaired source packet, focused validation complete.

## Source repair

The source repair is split from the validation receipt:

- `a5c79ccc2a93844c4518b0233eb4b9f02810539f` bounds classic data-label and
  pivot-format collections before parsing, validates nested label text and
  number-format/separator fields, uses bounded duplicate tracking, streams the
  selected label fields, preflights ChartEx text/CDATA using retained decoded
  UTF-8 bytes, and checks ChartEx external relationship IDs before output.
- `140d6eefd9f2c53d10f9daeebc34817c2a95f59c` fixes the streamed format-code
  attribute quote.
- `9280092a3cb35473efa75e4ec57debad800d5e95` adds focused regressions and
  resolves chart-style general references before retaining text, so entity
  spellings are budgeted by their decoded value.

The classic reader admits each `dLbl`/`pivotFmt` count before calling its
semantic parser. The writer's preflight covers group, series, and pivot labels;
the writer then streams the nested format/separator values into the bounded
sink. ChartEx text admission measures XML line-ending normalization before
`xml_content`, while its CDATA path measures `String::from_utf8_lossy` without
allocating a replacement string first. Style text and CDATA follow their
existing decoder semantics.

## Validation

All commands used the memory-capped cargo shim, one build job, `--locked
--offline`, and an external target directory. No native timing was run.

Pre-repair identity (`.coord/drawingml-chart-raw-validation-vc/pre-identities-a5c79ccc2a.txt`):

- source `a5c79ccc2a93844c4518b0233eb4b9f02810539f`
- source tree `1e8b7826112b0d61fa4d9a146e7a6ef28b9936cb`
- `Cargo.lock` SHA-256
  `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`
- lock blob `ce39fdc781093aa0819bcd17d41af891d4f33bcc`

Final identity (`.coord/drawingml-chart-raw-validation-vc/post-identities-9280092a3c.txt`):

- source `9280092a3cb35473efa75e4ec57debad800d5e95`
- source tree `83584e39954becba18543aad559ff7245386bd8c`
- the lock SHA-256 and blob are unchanged
- target directory size `6,090,876,243` bytes, below the 20 GB cap

Passing logs under `.coord/drawingml-chart-raw-validation-vc/`:

- `dml-all-features-lib-tests-9280092a3c.log`: 245 library tests and all
  DrawingML integration targets passed.
- `chart-limits-vc-rerun.log`: 12 focused limit tests passed.
- `chartex-tests-vc-rerun.log`: 9 ChartEx tests passed.
- `style-entity-test-vc.log`: decoded style-entity regression passed.
- `consumer-check-9280092a3c.log`: DOCX, PPTX, spreadsheet-drawing, XLSX,
  XLSB, and OPS all checked with all features.
- `docx-chart-tests-9280092a3c.log`: 9 DOCX chart tests passed.

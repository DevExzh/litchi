# ODS formula goldens

Expected values compiled into the `litchi-ods` formula tests
(`crates/litchi-ods/tests/ods_formula_*.rs`). Each directory is one function
family and keeps the layout it had under
`docs/report/spec-gap-validation-evidence/ods-formula-<family>/`, where these
files lived before they moved here.

Two kinds of file:

- `*-goldens.json`, `oracle-vectors.json`: project-generated oracle values,
  computed with exact rational or high-precision arithmetic independently of
  the Rust evaluator. They carry no third-party content.
- `native/cached-results.json`, `native/native-results.json`,
  `native/steyx-deviation-proof.json`: formula results that LibreOffice stored
  in its own function-test spreadsheets. `native/provenance.json` beside each
  one names the upstream files and their SHA-256 hashes.
  - Upstream: LibreOffice core (<https://git.libreoffice.org/core>), under
    `sc/qa/unit/data/functions/`, at the commit recorded in each
    `provenance.json`.
  - License: MPL-2.0 (`LICENSE-MPL-2.0.txt` in this directory).

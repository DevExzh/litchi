# Change 0904 exporter and artifact review

Reviewed at HEAD `0283e96b73af76906c1a83ce90956e1c6afa1116`. The exporter uses the benchmark's deterministic XLSX corpus, update selection, source-backed editor, commit, and stream-publication route. The review made no Rust or Python source changes.

The artifact contract passes in the retained verification files:

- `identity-check.json` is `pass`; every case's source and output hashes match the prior identities, and the independently computed changed XML member set equals the manifest's selected parts. The cases are medium one-percent (workbook plus four worksheets), dense-sparse one-percent (workbook plus four worksheets), and medium one-edit (workbook plus sheet 1).
- `analysis.json` records byte-verified ZIP sidecars and full original/replacement XML census results. Its quick-xml counts agree with the independent Expat census. The primary artifacts are over 99.9% zero/one-attribute events without `xml:space`; the one-edit control is approximately 99.87–99.89% (root corrected the initial all-artifact percentage wording before sealing).
- `replay-run.json` records byte-identical census replay and refusal to overwrite an existing result; `census-mutations-run.json` records the fail-closed mutation checks with inputs unchanged.

The proof profile is `xml_minifier::audit::Limits::default()`, matching the OPC source-replacement audit. The medium one-percent XML parts have `Complete` proofs; dense-sparse has four `Complete` proofs and a sheet-4 `Window` at `[1494,1508]`; medium one-edit has a workbook `Complete` proof and a sheet-1 `Window` at `[190,198]`. A `Complete` proof covers the full replacement audit. A `Window` records only its bounded span; the Python census does not replay the Rust auditor or claim full replacement coverage for that proof.

The exact changed-member check is enforced independently by `verify.py`/`identity-check.json`, rather than inferred from the Rust exporter’s selected-part list. The hashes and sidecars establish lexical decompressed XML bytes. This review makes no claim about preserving raw compressed ZIP member bytes.

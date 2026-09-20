# Date/time isolated gate preparation

This directory contains preparation tooling for the ODF 1.4 §6.10 date/time
batch. It has no frozen candidate or gate result yet. The scripts do not carry
historical test totals, source hashes, or identifiers from another batch.

`stage.py` reads the date/time `baseline.json`, requires the isolated checkout
to be at `preparation_commit`, and copies the retained isolated `Cargo.lock`
whose path and SHA-256 are declared by that baseline. The selected closure is
derived from the current tree:

* every path in `selected_baseline_sources`, workspace manifests, and the
  positive immutable date/time input allowlist: `baseline.json`,
  `contract.md`, `coverage-scope.json`, the explicit oracle plan/vector/
  verifier inputs, native
  fixtures/results/provenance, the fixed performance plan/case matrix/runner, and its
  harness sources;
* the complete recursive `crates/litchi-ods/src/codec/formula/evaluation`
  subtree, including generic dispatch, cache, reference, calendar, parser,
  timestamp, and date/time helpers;
* the formula expression/reference Rust dependencies and both feature-matrix
  documents;
* every current `ods_formula_date_time*.rs` integration test.

`coverage-requirements.json`, review documents, completion reports, audit
records, and verification/result receipts are intentionally excluded from the
frozen source map. Coverage-scope and oracle vectors are immutable freeze
inputs; coverage bindings and verifier receipts are validated as separate
outputs.

The staged map, batch-format list, closure categories, and preparation warnings
are written to `stage-manifest.json`, `staged-profile-sources.json`, and
`batch-files.json`. A missing future date/time module or test file is reported
as a preparation warning. The tool never writes `freeze.json`, runs Cargo, or
claims that the candidate is ready.

After semantic/resource review and source freeze, root should create
`freeze.json` from the reviewed `stage-manifest.json` with this shape:

```json
{
  "schema": "ods-formula-date-time-freeze-v1",
  "base_commit": "<baseline preparation_commit>",
  "isolated_lock_sha256": "<baseline isolated_lock.sha256>",
  "selected_files": {"<repository-relative path>": "<sha256>"}
}
```

Only then may root run the following commands in the isolated checkout:

```text
python3 gates/run.py <isolated-checkout> <isolated-target>
python3 gates/verify.py
python3 gates/closure_negative_cases.py
```

Before that handoff, `run.py --allow-pending` and `verify.py --allow-pending`
report `verified: false` without starting a command. The runner records the
seven locked/offline package, lint, documentation, formatting, boundary, and
diff checks dynamically; `verify.py` checks their exact command lines, logs,
source stability, lock identity, and all reported test summaries. It does not
invent a test count. `closure_negative_cases.py` checks path traversal,
missing-source, malformed-hash, and closure-category refusal without Cargo.

The date/time evidence verifier one directory above remains the owner of
contract/coverage bindings. These gate scripts only custody the isolated source
and execution receipts; they do not turn pending coverage into a PASS.

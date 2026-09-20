# Lookup integration gates

All seven isolated gates pass against `freeze.json`: 1,699 tests, zero failures and zero ignores. `verify.py` verifies the source closure, commands, logs and test summaries. Performance and independent review acceptance are separate requirements.

`selected-source.tar.gz` contains the exact 61 selected inputs, with each member verified against the freeze. Extract it into the baseline commit recorded by the manifest to reconstruct this candidate. Archive SHA-256: `6e15caf59814b5124dd720f87be26eb706ca33edaf71c92f8119c021f4d8867a`. The retained `Cargo.lock` is the isolated dependency input; do not replace the ambient root lockfile.

The `diagnostic-attempt-*` directories retain earlier passing snapshots and explain why they were superseded. They are historical evidence, not additional final-candidate gate runs.

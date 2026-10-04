# Change 0904 evidence

See [the report](../../0904-xlsx-attribute-census.md) for interpretation.
`plan.json` freezes the three-case diagnostic scope and initial source hashes.
`source-final.json` binds the final harness files; `exporter-v1.rs` preserves
the initial source before two lint fixes and an introductory comment correction.

Build and export from the repository root to a new directory:

```sh
cargo +1.95.0 build --manifest-path tools/perf-baseline/Cargo.toml --locked \
  --offline --example xlsx_attribute_fixtures
./tools/perf-baseline/target/debug/examples/xlsx_attribute_fixtures \
  /path/to/new-fixtures --include-one-edit
```

Use the actual Cargo target directory if overridden. This packet's commands
used the isolated `target/litchi-0904-build`, removed after verification.
The debug binary is diagnostic, not a timing baseline. Source/output/member
identities are in `artifacts/manifest.json`; prior benchmark identities are
bound by `prior-identities.json`.

Reproduce the independent full-artifact census with a new output path:

```sh
python3 -B docs/performance/results/change-0904/census.py \
  docs/performance/results/change-0904/artifacts --output /path/to/new-analysis.json
python3 -B docs/performance/results/change-0904/verify.py
python3 -B docs/performance/results/change-0904/census_mutation_test.py
```

`verify.py` checks prior report hashes, archive identity, exact changed-member
scope, byte-identical canonical replay and reducer overwrite refusal. It removes
its temporary output. `final-export-verification.json` records final-export
parity and exporter overwrite refusal. `command.py` writes create-only command,
log and source-stability receipts; failed invocation and initial Clippy outputs
are preserved. No helper was changed after execution.

`analysis.json` contains structural counts only. Rust proof kinds are diagnostic
observations; Python does not reproduce the auditor's semantic/window proof.
The selected full replacement is counted even when only a window was re-audited.
No native timing, allocation, RSS or performance improvement is claimed.

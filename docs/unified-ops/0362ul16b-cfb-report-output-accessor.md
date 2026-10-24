# CFB report output accessor ordinary-harness closure

The composed source at `a3ec9799939071d535925133c483bc998aa9667c` still
contained eight ordinary performance-harness consumers of the removed
`PublishReport::bytes()` compatibility shorthand. This return closes the
consumer migration in the same wave: `filesystem.rs:5888,5935`,
`xls_numeric.rs:1472,1662`, and `lib.rs:30621,30681,31231,31296` now call
`output_bytes()`. The remaining `report.bytes()` calls in the performance
package belong to different report types and are unchanged; the ordinary
all-targets check confirms that no further `PublishReport` consumer remains.

The DrawingML preparer now pins the reviewed
`tools/perf-baseline/src/lib.rs` bytes to the post-migration SHA-256
`280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a`.
The other fixed protocol pin was audited without modification:
`prepare_xlsx_raw_inputs.py` still matches the reviewed XLSX identity
`tools/perf-baseline/unified-xlsx-raw-identity/src/lib.rs` hash
`d65ce0630a50d1b972a0b2f93044091d1a5c50df1e9c8ef06e9c3ba65cf8a36a`.
Formatting-only closure of the existing XLS consumer line is included so the
workspace formatter remains clean.

The ordinary performance package check passed with the repository Cargo shim,
one build job, offline locked resolution and a fresh external target:

```text
cargo check --manifest-path tools/perf-baseline/Cargo.toml --offline --locked --all-targets PASS
cargo fmt --all --check                                      PASS
23 DrawingML/XLSX preparer tests                             PASS
```

The check emitted the existing DrawingML probe warnings but exited zero. It
does not make an ABBA or performance-acceptance claim. The exact source, lock,
command-log and target identities are bound by the [validation receipt](results/0362ul16b/validation-receipt.json).

# Final quality gates

All ten gates ran against the final tree. `cargo fmt --all --check` failed on an
earlier run — three rustfmt wrappings in `#[cfg(test)]` code — and passes after
`cargo fmt --all`; every other gate passed on the first run.

| # | Command | Result | Counts |
| --- | --- | --- | --- |
| 1 | `cargo test --offline --locked -p litchi-cfb -p litchi-ole-common -p litchi-xls -p litchi-doc -p litchi-ppt --all-features -- --test-threads=2` | PASS | 4,238 passed, 0 failed, 27 ignored |
| 2 | `cargo check --offline --locked --workspace --all-features` | PASS | 64 packages, 0 warnings |
| 3 | `cargo clippy --offline --locked -p litchi-cfb -p litchi-ole-common -p litchi-xls -p litchi-doc -p litchi-ppt --all-features --lib -- -D warnings` | PASS | 0 lints |
| 4 | `cargo doc --offline --locked -p litchi-cfb -p litchi-ole-common -p litchi-xls -p litchi-doc -p litchi-ppt --all-features --no-deps` with `RUSTDOCFLAGS=-D warnings` | PASS | 5 crates, 0 warnings |
| 5 | `cargo fmt --all --check` | PASS after `cargo fmt --all` | 3 test-only wrappings corrected |
| 6 | `python3 -B tools/check_crate_boundaries.py` | PASS | 64 packages, 241 internal declarations, 11 pre-existing iWork debt items |
| 7 | `cargo check --offline --locked -p litchi-cfb -p litchi-xls --no-default-features` | PASS | — |
| 8 | `python3 -B tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --evidence-root . --mode strict` | PASS | 10 claims validated |
| 9 | `cargo test --offline --locked -p litchi-opc -p litchi-docx -p litchi-pptx -p litchi-xlsx --all-features -- --test-threads=2` | PASS | 4,325 passed, 0 failed, 34 ignored |
| 10 | `cargo clippy --offline --locked -p litchi-cfb -p litchi-ole-common --all-targets --all-features -- -D warnings` | PASS | 0 lints across 11 units |

Per-crate library unit tests: `litchi_cfb` 287, `litchi_ole_common` 129,
`litchi_doc` 1,009, `litchi_ppt` 1,092, `litchi_xls` 1,028, `litchi_opc` 423,
`litchi_docx` 946, `litchi_pptx` 552, `litchi_xlsx` 991 — all with zero failures.
Gate 10 was re-run after `cargo clean -p litchi-cfb -p litchi-ole-common` because
the first invocation returned fully cache-fresh; the re-run genuinely recompiled
all 11 units.

Also run outside the numbered set and passing:
`python3 -B tools/check_report_claim_classification.py` (167 rows, 2 tables) and
`python3 -B tools/validate_crud_coverage_index.py` (15 categories, 33 selectors).

`cargo fuzz` is not installed on this host, so no fuzz gate was run. No Miri,
sanitizer, native Office security-corpus, cold-cache, cross-platform, or scaling
result is claimed by this batch.

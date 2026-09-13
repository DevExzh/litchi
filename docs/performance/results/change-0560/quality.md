# Final quality gates

Run against the tree containing changes 0560, 0561 and 0562. The nine numbered
gates were run by an independent verification pass that detected concurrent
edits mid-run and re-ran every gate against the final state; the additional
commands below were run directly.

| # | Command | Result | Counts |
| --- | --- | --- | --- |
| 1 | `cargo fmt --all --check` | PASS | 0 diffs |
| 2 | `cargo test --offline --locked -p litchi-cfb -p litchi-ole-common -p litchi-xls -p litchi-doc -p litchi-ppt --all-features -- --test-threads=2` | PASS | 4,240 passed, 0 failed, 27 ignored |
| 3 | `cargo check --offline --locked --workspace --all-features` | PASS | 64 packages, 0 warnings |
| 4 | `cargo clippy --offline --locked -p litchi-cfb -p litchi-xls --all-targets --all-features -- -D warnings` | PASS | 0 lints |
| 5 | `cargo doc --offline --locked -p litchi-cfb -p litchi-xls --all-features --no-deps` with `RUSTDOCFLAGS="-D warnings"` | PASS | 0 warnings |
| 6 | `python3 -B tools/check_crate_boundaries.py` | PASS | 64 packages, 241 internal deps, 11 pre-existing iWork debt items |
| 7 | `cargo check --offline --locked -p litchi-cfb -p litchi-xls --no-default-features` | PASS | 0 warnings |
| 8 | `python3 -B tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --evidence-root . --mode strict` | PASS | 10 claims validated |
| 9 | `cargo test --offline --locked -p litchi-opc -p litchi-docx -p litchi-pptx -p litchi-xlsx --all-features -- --test-threads=2` | PASS | 4,325 passed, 0 failed, 34 ignored |

Additional commands for the crates this batch touched beyond that set:

| Command | Result | Counts |
| --- | --- | --- |
| `cargo test --offline -p soapberry-zip --all-features` | PASS | 439 lib tests plus every integration binary, 0 failed |
| `cargo clippy --offline --locked -p soapberry-zip --all-targets --all-features -- -D warnings` | PASS | 0 lints |
| `cargo test --offline --locked -p litchi-odf-common -p litchi-odt -p litchi-ods -p litchi-odp --all-features` | PASS | 0 failed across every binary |
| `cargo lint` (the repository's pinned `clippy --workspace --all-features --lib --no-deps -- -D warnings`) | PASS | 0 lints |
| `python3 -B tools/non_iwork_gate.py verify` | PASS | 45 bulk tree roots, 35 safe facade trees, 1 combined tree |
| `python3 -B tools/non_iwork_gate.py deprecated` | PASS | no deprecated non-iWork API use |
| `python3 -B tools/check_report_claim_classification.py` | PASS | 167 rows, 2 tables |
| `python3 -B tools/validate_crud_coverage_index.py` | PASS | 15 categories, 33 selectors |

`cargo fuzz` is not installed on this host, so no fuzz gate was run. No Miri,
sanitizer, native Office security-corpus, cold-cache, cross-platform or scaling
result is claimed by this batch.

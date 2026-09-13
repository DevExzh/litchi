# Final quality gates

Run against the tree containing change 0565. All cargo gates used
`--offline --locked` with the pinned 1.95.0 toolchain.

| # | Command | Result | Counts |
| --- | --- | --- | --- |
| 1 | `cargo fmt --all --check` | PASS | 0 diffs |
| 2 | `cargo test -p litchi-cfb -p litchi-ole-common -p litchi-xls -p litchi-doc -p litchi-ppt --all-features -- --test-threads=2` | PASS | 154 binaries, 4,244 passed, 0 failed, 27 ignored |
| 3 | `cargo check --workspace --all-features` | PASS | 64 packages, 0 warnings |
| 4 | `cargo clippy -p litchi-cfb -p litchi-xls --all-targets --all-features -- -D warnings` | PASS | 0 lints |
| 5 | `cargo doc -p litchi-cfb -p litchi-xls --all-features --no-deps` with `RUSTDOCFLAGS="-D warnings"` | PASS | 0 warnings |
| 6 | `python3 -B tools/check_crate_boundaries.py` | PASS | 64 packages, 241 internal deps, 11 pre-existing iWork debt items |
| 7 | `cargo check -p litchi-cfb -p litchi-xls --no-default-features` | PASS | 0 warnings |
| 8 | `python3 -B tools/check_perf_claims.py --mode strict` | PASS | 10 claims validated |
| 9 | `cargo test -p litchi-opc -p litchi-docx -p litchi-pptx -p litchi-xlsx --all-features -- --test-threads=2` | PASS | 210 binaries, 4,328 passed, 0 failed, 34 ignored |
| 10 | `cargo test -p litchi --all-features --no-fail-fast -- --test-threads=2` | FAIL, 4 pre-existing | 25 binaries, 532 passed, 4 failed, 11 ignored |
| 11 | `cargo lint` | PASS | 0 lints |
| 12 | `python3 -B tools/non_iwork_gate.py verify` and `deprecated` | PASS | 45 bulk tree roots, 35 safe facade trees; no deprecated non-iWork API use |
| 13 | `python3 -B tools/check_report_claim_classification.py` | PASS | 167 rows, 2 tables |
| 14 | `python3 -B tools/validate_crud_coverage_index.py` | PASS | 15 categories, 33 selectors |

## Gate 10's four failures are pre-existing

`cargo test` aborts after the first failing target, so the gate as previously
written ran only 1 of the facade's 25 test targets and its integration tests
never executed. With `--no-fail-fast`, exactly four tests fail and nothing else:

- `document::doc::tests::filesystem_odt_keeps_source_owner_with_malformed_ooxml_catalog`
- `document::doc::tests::owned_odt_bytes_keep_odt_owner_with_malformed_ooxml_catalog`
- `document::doc::tests::managed_docx_facade_paragraph_text_avoids_rich_paragraph_refusal`
- `document_limit_arbitration::ordinary_odt_uses_native_policy_even_with_ooxml_suffix_but_polyglot_honors_docx_input_limit`

All four were established empirically to fail at HEAD, by reverting the two
`crates/litchi-xls/` files and rerunning. The fourth needed care: the facade
change sits inside a `#[cfg(test)]` module, and an integration-test target links
the library compiled without `cfg(test)`, so reverting the two `litchi-xls`
files reproduces HEAD exactly for that target. It fails identically with and
without this batch, on an ODT and DOCX input-limit arbitration path with no
dependency on `crates/litchi-xls`.

**Recommendation for the standing gate list**: add `--no-fail-fast` to gate 10,
or the facade's integration tests go unobserved whenever a library test fails.

## Not run

`cargo fuzz` is not installed on this host. No Miri, sanitizer, native Office
security-corpus, cold-cache, cross-platform or scaling result is claimed.

`tools/perf-baseline` has its own workspace table and is outside every cargo
gate above; its locality-gate change was compiled by building both measurement
binaries from trees carrying it.

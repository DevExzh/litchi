# 0359qg — DOC body-toggle final-472b build and correctness receipt

- **Status:** build and focused correctness qualification complete; performance
  acceptance remains open.
- **Date:** 2026-10-24.
- **Scope:** the final production endpoint bound by 0359qf, its reviewed
  two-file body-publication baseline, and the existing W5 source probe. This
  follow-up records actual build/output checks only; it does not replace the
  quiet-host ABBA/allocation protocol or accept a performance result.

## External run and source binding

All generated targets, logs, binaries, probe JSON, OLE outputs, and the full
receipt are outside the checkout:

~~~text
(deleted-coord-artifact)
~~~

The candidate and baseline were both prepared from the final production
endpoint:

| item | identity |
|---|---|
| candidate source | `(deleted-coord-artifact)` |
| baseline source | `(deleted-coord-artifact)` |
| commit (both) | `472b6705475b073e7fe08f6e11ee17fb1e29f838` |
| tree (both) | `ebc1cee05f124c0c63977aff07a23288c19e7d00` |
| Cargo.lock SHA-256 (both) | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` |
| rust-toolchain.toml SHA-256 | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |
| `.cargo/config.toml` SHA-256 | `50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d` |

The candidate is clean. The baseline has only the reviewed reversals in
`crates/litchi-doc/src/body_text.rs` and
`crates/litchi-doc/src/tracked_revision/package.rs`; its applied two-file
diff SHA-256 is
`0e0bf3c606277da07e53c085fa52a2e1db99839a3fc8ca5e869154ae60ab59c0`.
The candidate and baseline body/revision file hashes are:

| source | `body_text.rs` | `tracked_revision/package.rs` |
|---|---|---|
| candidate | `db01cfce38c13c6c806382ffede990fda2b3bbe4dd53f59c4c5cb9feba9777d3` | `029721d34e7ad658e0064afd258cdb2b5a40e330110255b157aafb83f0d7b12a` |
| baseline | `6dc26da8130659c212dd29f26499cdaba7c4849268583b44dfcd9413a36887a3` | `b7155dc7c9aae6e318845b5073da4c0f275889125d38eae61df091aa2a2166df` |

The probe source is identical in both legs, SHA-256
`b808e4dcba33dcafd4dff7dd22dc250e4170f2613839044fa8097263f5020965`.
The prior reviewed overlay recipe is recorded as SHA-256
`10765febb64c4e099f5c3ef560366124dc1d28d540d138768ae0f7cf782da2ab` in the
source-bound packet; the prepared final baseline diff above is the artifact
used for this run.

The complete external receipt, including every source, command, lock, log,
binary, output, and target-size identity, is:

~~~text
(deleted-coord-artifact)/actual-receipt.txt
SHA-256 a0dfb1d1b66c16249939a702d26565466a3ca8b86d11e9bc87461bf5df414b3a
~~~

The build, diagnostic-test, qualified-test, and probe logs are also bound by
the full receipt. Their SHA-256 values are recorded here for quick checking:

| log | SHA-256 |
|---|---|
| `build-candidate.log` | `d36851d11bf903910805c95c09bc9ace1cc4e850de385f0931bac16011b2d5ec` |
| `build-baseline.log` | `37b7dd64ec9473afdcf754b26b2b6e26c88e4ab49281ecd89c4df00a6f0074a7` |
| `test-candidate-deferred.log` (zero-test diagnostic) | `bf423b2577cf6866aea365167c201cd5702b9df367de2700d1b8661cb6937d55` |
| `test-candidate-retained.log` (zero-test diagnostic) | `df66fd7853e07a33ec6c680d57f2ce30ae03445ae1ed49f9317bf66b178ca3ba` |
| `test-baseline-retained.log` (zero-test diagnostic) | `0d88054b219a13497db8a9791c94783d56cde72c7d4c05a6099b311fa7e7568a` |
| `test-qualified-candidate-deferred.log` | `b44c3df2bd8892cffa1591b572cc18acdd0544b7bac6a6eb398d030583b97f07` |
| `test-qualified-candidate-retained.log` | `905c24a386aa140435749e5dc09860c4ed6598e863cb31d6b8a15ca32b4d81a1` |
| `test-qualified-baseline-retained.log` | `73b17c849deef3ec100628e20b5a207851a5c893b1e2dacdee240d9f1c4926c2` |
| `candidate-leaf-doc_leaf_write.log` | `b0bb9b4667a13dc437061a4936ca2c7414fc63590e56332800e028b1abd3fa81` |
| `baseline-leaf-doc_leaf_write.log` | `6f144967d09837cb07f9e052aed68241b1018cea3756fc2f11867d373fd09e65` |
| `candidate-batch-doc_batch1k_leaf_write.log` | `0f7a36fff267cb258f566da02a45c815486f1d7ca76e03dd180d2fcc85346a28` |
| `baseline-batch-doc_batch1k_leaf_write.log` | `6cf03671c6cfd12ef49b8b9f60754d736b07ecea61a74d500c60807ebd9d9d16` |
| `chpx-audit.json` | `1df6ae24c74cca29fb22b620dad27ad44134a6cc10f42282c29b22e099fff62b` |

## Builds and exact focused tests

Both release example builds passed with the memory-capped environment,
`CARGO_BUILD_JOBS=1`, `--offline`, and `--locked`:

~~~sh
cargo build --offline --locked --release -p litchi-ops \
  --features doc,xls,rtf --example legacy_w5_source_probe
~~~

The candidate and baseline example binaries have SHA-256
`ba7042fd413c26764986b0ed9dcd79beca85e91231f0beb2f850550152e14877` and
`2d62dacae12b49214a086c109b506ca8eb69a33cf5e770627794efecdbcba810`,
respectively. Their release targets are 567,455,995 and 567,533,769 bytes;
the debug test targets are 309,231,179 and 309,229,555 bytes. All are below
the 20,000,000,000-byte limit.

The first test attempt is retained as a diagnostic because unqualified names
combined with `-- --exact` selected zero tests in every leg. Its three logs
explicitly report `running 0 tests`; those rc=0 results are not validation.
The corrected commands used the full module-qualified names, separately:

~~~sh
cargo test --offline --locked -p litchi-doc --lib \
  'body_text::tests::deferred_body_formatting_flushes_before_table_edits_and_keeps_failures_atomic' -- --exact
cargo test --offline --locked -p litchi-doc --lib \
  'body_text::tests::retained_render_is_cleared_by_resource_reopen_and_formatting_flushes_at_commit' -- --exact
cargo test --offline --locked -p litchi-doc --lib \
  'body_text::tests::retained_render_is_cleared_by_resource_reopen_and_reenabled_for_formatting' -- --exact
~~~

The first two commands ran against the candidate and the third against the
baseline. Each log reports exactly `1 passed; 0 failed`; no test leg was
combined with another filter.

## Output and CHPX fidelity checks

The probe was run only for output generation, with one round, and its timing
fields were not used. The ordinary one-toggle publication output is byte
identical:

| route | candidate | baseline |
|---|---:|---:|
| `doc/leaf_write` bytes | 168,960 | 168,960 |
| `doc/leaf_write` SHA-256 | `6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11` | same |

The 1,000-toggle route intentionally does not claim raw OLE equality:

| route | size | SHA-256 |
|---|---:|---|
| candidate `doc/batch1k_leaf_write` | 182,784 | `bd016a87f5d0015d6f8b910862eeeef9c0ce490033b4951c32ee5f403b7eebdf` |
| baseline `doc/batch1k_leaf_write` | 7,715,328 | `4a5bcc2888ab819488fbb5a1ced1b088010440bd8293f10556e0903993e448d3` |

The reviewed `audit_doc_chpx_history.py` oracle compared the FIB-selected
`PlcBteChpx` ranges and referenced FKP payloads and returned:

~~~text
PASS: skipped blocks are not reachable through the FIB-selected PlcBteChpx
~~~

The external audit JSON SHA-256 is
`1df6ae24c74cca29fb22b620dad27ad44134a6cc10f42282c29b22e099fff62b`.

## Explicit non-acceptance boundary

This receipt is not a performance result. No quiet-host ABBA, native timing
run, allocation capture, Callgrind, Massif, Valgrind, `systemd-run`,
`/usr/bin/time`, or `taskset` run was performed. The probe's internal timing
fields are retained only in the external raw JSON and are deliberately not
used here. The queued performance/allocation protocol in 0359qf remains open.

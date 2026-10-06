# 0922 shared-owner DOC quality receipt

Created `2026-10-06T07:56:49.918Z`. The reviewed production patch is based on `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0` and is limited to the three Rust files listed in `shared-owner-quality.json`; the quality workspace regenerated `Cargo.lock` offline solely for locked checks. All commands used `/tmp/litchi-performance-0922/source` with Rust/Cargo 1.95.0.

The final source-only patch SHA-256 is `dd7f0bafb81c23a3762e74246eccf991b56ff1f7e62816087e27441db636bd84`. The three final source hashes are recorded in the JSON. The borrowed FIB/Selsf path preserves public `parse_at`/`raw_data` behavior, legacy and counted pointer bounds, Selsf missing/zero versus strict range errors, malformed-record verdicts, and preflight atomicity.

Affected checks passed:

- full `litchi-doc --lib --tests`: **1349 passed, 0 failed, 2 ignored**;
- FIB unit: **8 passed**;
- SavedSelection unit: **15 passed**;
- affected integrations: **8 + 10 passed**;
- clippy with `-D warnings`: passed;
- scoped rustfmt and `git diff --check`: passed.

Two repository baseline failures remain recorded: workspace fmt reports three unrelated iWork files, and strict rustdoc reports four unchanged private intra-doc links. Rustdoc passes with the existing baseline lint allowance; details and hash-bound logs are in the JSON.

# 0362ul87ae — XLSB source no-op cancellation fence

Status: source-only follow-up to 0362ul87ad.  No Cargo, rustdoc, native
qualification, performance measurement or final gate was run for this repair.

Both source-backed XLSB metadata no-op paths now call
`SourceBackedPackage::check_execution()` before `source_version()`.  Managed
cancellation therefore wins over a simultaneous source revision, matching the
workbook source preflight ordering.  The ordinary no-op and revised-owner
replay paths retain their O(1) source-version freshness check after the
cancellation check.

The source-backed integration suite covers cancellation for no-op commit and
revised-owner no-op replay, with both tests arranging cancellation and a stale
source together to verify the typed cancellation-before-source-change
precedence.

## Source-only checks

`git diff --check` and scoped `rustfmt --edition 2024 --check` pass for the
changed Rust sources.  The coordinator should run the XLSB metadata source
tests and full gate after integration.

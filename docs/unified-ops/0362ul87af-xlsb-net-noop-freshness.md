# 0362ul87af — XLSB metadata net-no-op freshness fence

Status: source-only follow-up to 0362ul87ae.  No Cargo, rustdoc, native
qualification, performance measurement or final gate was run for this repair.

The source-backed XLSB metadata edit now shares one no-op preflight helper
across exact no-op commits, edits whose new patch reverses the retained
revision to a net no-op, and no-op patch replay.  The helper checks managed
execution cancellation first and then the retained source version.  A
non-empty undo patch can therefore no longer clear a revision and publish an
unchanged owner after its source advanced, and cancellation still wins over a
simultaneous stale-source condition.

The source-backed integration suite adds stale-source and
cancellation-before-stale-source regressions for revised-owner undo commits.

## Source-only checks

`git diff --check` and scoped `rustfmt --edition 2024 --check` pass for the
changed Rust sources.  The coordinator should run the XLSB metadata source
tests and full gate after integration.

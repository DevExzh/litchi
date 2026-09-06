# Retained development checks

The first filtered ODS unit receipt is retained as a failed check because the
copy guard blocked the requested filter and the test binary reported zero
selected tests. It has `exit_code: 0`, `passed_tests: 0`, and no failed tests;
it is a selection/setup failure rather than a failing assertion
([`ods-text-unit-first.json`](ods-text-unit-first.json),
[`ods-text-unit-first.log`](ods-text-unit-first.log)). The corrected retained
unit run selected four tests and passed all four
([`ods-text-unit-v2.json`](ods-text-unit-v2.json)).

The first unscoped strict Clippy run stopped on the pre-existing common-crate
`ArchiveReaderKind` `large_enum_variant` lint
([`ods-strict-first.log`](ods-strict-first.log)). After that existing debt was
explicitly excluded, the first scoped run found exactly two new
`int_plus_one` assertions in `streaming_text_spans.rs`; those assertions were
changed to the equivalent strict-inequality form. The retained scoped rerun
passes with warnings denied
([`ods-strict-scoped-first.log`](ods-strict-scoped-first.log),
[`ods-strict-scoped-v2.json`](ods-strict-scoped-v2.json)).

The final release ODS receipt records 454 passed, 0 failed, and 0 ignored
tests, with unchanged source custody
([`ods-tests-final.json`](ods-tests-final.json)). These functional and lint
receipts authorize no performance or speedup claim.

The first sealed portable replay failed because summary rederivation exposed
`.log.gz` resource paths after sealing where the pre-seal summary recorded the
logical `.log` paths. All 24 differences were paths; measurements were unchanged.
The summary driver now emits stable logical resource paths in either storage
form. Its previous bytes are retained in `versions/summary-pre-compression-fix.py`,
and the failed replay receipt/log remain retained. The successful precleanup
retry uses a new tag; no benchmark process or sample was replaced.

The second copied replay passed all six mutation probes and summary derivation,
but its source-inventory guard failed because root edited README and the ADR
record while replay was running. This is retained as a failed replay with
`exit_code: 0` and `inventory_members_unchanged: false`; it does not authorize
cleanup. The next replay runs after documentation edits stop. Both previous
cleanup/lifecycle driver versions are retained. No workload or profile changed.

The first staged whitespace check found one extra blank line at EOF in
`next-work.md`. The line was removed and a separately tagged retry passed.

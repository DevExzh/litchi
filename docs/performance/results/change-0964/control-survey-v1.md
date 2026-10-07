# 0964 conditional confirmation control survey v1

Status: read-only design. No build, benchmark, or workload was run.

## Source-custody correction

An earlier note incorrectly treated the different 0957 and current commit hashes as proof that the retained 0957 source was stale. That inference is retracted. Root's audit proves that all 10,932 retained 0957 runtime/harness files match the current committed runtime/harness closure in content and mode. Different commit hashes do not establish source drift.

The retained 0957 managed benchmark and allocator executables therefore remain valid for the inherited 0964 screen. The reason a new generic baseline is required is narrower: 0957 retained no generic perf-baseline executable, only the two managed executables. Generic controls need fresh matched release baseline and candidate builds from the proven equivalent current source and harness closure. The 0706 generic artifacts require a separate content audit before reuse; their commit hash alone is not a rejection reason.

## Conditional control matrix

The managed harness accepts only medium/dense-sparse one-percent and medium one-edit. Those are the three existing screen cases, and their marker-free cell-CRUD source is SourceAdmission::Borrowed. They exercise both changed validation and shared raw transport, but no authoritative fallback.

If the screen qualifies, use the smallest additional matrix that covers the route classes:

| Control | Expected route | Purpose |
| --- | --- | --- |
| xlsx_source_backed_cell_values_one_edit_save on medium and dense-sparse | Borrowed | Ordinary unmanaged context over the same changed validator/raw loop. |
| xlsx_producer_medium_source_planning (and, if publication needs a control, xlsx_producer_medium_source_one_edit_save) | Rewritten, subject to a current route census | Producer-shaped namespace transport; planning isolates source admission from publication. |
| A new generic fallback fixture/selector derived from the accepted worksheet in cell_values/admission_tests.rs::compatibility_markers_on_the_worksheet_root_are_copied_through | None because of x14ac:dyDescent, or a deliberate ProvisionalFailed proof-refusal fixture | Exercises changed validate_xml on authoritative fallback. The current producer Read/real-file shapes refuse earlier and are not valid timed fallback controls. |

The producer Edit variant carries the markup-compatibility declaration but no dyDescent or AlternateContent; current source_stream_admission therefore predicts Rewritten, while the old producer sidecar mirrors the pre-0603 source_stream_eligible predicate and cannot establish that route. A fresh untimed census must record SourceAdmission, rewrite-proof/parse outcome, validator finish, and the authoritative result or error before timing.

A markup-compatibility namespace means Rewritten, not automatically fallback. Proof failures such as processing instructions, late declarations, custom entities, invalid QNames/namespaces, duplicate or malformed attributes, and undecodable attributes attempt the shared route and then fall back authoritatively. dyDescent, AlternateContent, invalid UTF-8, or admission-limit failures return None directly. Route classification must use the actual predicate and proof result rather than marker-prefix presence.

The historical 0603 census remains useful qualification context: 207 real worksheet parts contained 71 Borrowed, 31 Rewritten, and 105 refused parts; no shipped worksheet completed the shared traversal. It is not a timed fallback corpus.

## Future matched run boundary

Build baseline and candidate from identical current source, harness, lockfile, and toolchain, with only the approved production diff between them. A future normal control invocation is:

~~~sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 20 --samples 200 \
  --case xlsx_source_backed_cell_values_one_edit_save,\
xlsx_producer_medium_source_planning,xlsx_producer_medium_source_one_edit_save \
  --xlsx-cell-crud-shape medium,dense-sparse \
  --producer-evidence <outside>/producer-evidence.json \
  --json <outside>/generic-control.json
~~~

The existing ordinary-cell oracles remain outside the timer: immutable source, exact output/package preservation, semantic reopen, inverse restoration, bounded sink, and resource cleanup. Producer planning requires exact no-op/source-unchanged checks; producer one-edit-save requires the exact target splice, semantic readback, and untouched-member preservation. The fallback control must additionally prove its route (None or ProvisionalFailed) and compare the candidate with the authoritative validator/parser for exact value, typed error, and error ordering. Route census, fixture generation, and all oracle work remain outside measured samples.

The exact route tests should use their full module paths and reject a zero-test match, for example:

~~~sh
cargo test --release --locked -p litchi-xlsx --lib \
  cell_values::shared_traversal_tests::marker_admission_matches_the_authoritative_path_on_every_real_worksheet \
  -- --exact
cargo test --release --locked -p litchi-xlsx --lib \
  cell_values::shared_traversal_tests::rewrite_only_refusals_survive_marker_admission \
  -- --exact
~~~

These commands are future qualification commands only; they were not run for this survey.

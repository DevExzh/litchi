# 0728 final probe review resolution

This follow-up checks the four findings in `final-review.md` against the
final source handoff and the regenerated qualification packet.  The reported
build-5, quality-3, qualification-3, and synthetic analyzer-preflight
receipts are successful.  I did not run Cargo, native timing, or profiler
commands.

## Public expected metadata gate: resolved

`raw_directory_oracle` now requires `source_expected_ok` for every operation,
including the public format expected artifact.  Format reports use the
`public_format_raw_source_model_gate` mode, while Reuse applies the same
source-model gate and Rewrite remains a report-only physical-normalization
route.  `analyze.py` and `audit.py` require zero source/expected normalized
directory difference, and measured format/Reuse output also has to match that
source model.  Qualification-3 reports the source-model gate as passed on all
18 routes.

## Direct semantic comparison: resolved

DOC stories, projections, table cells, fields, revisions, and embedded
objects now compare the returned values directly; `direct_match` is part of
the semantic acceptance path.  PPT survivor order compares the full witness,
including the in-memory slide text, outline values, notes/comments values, and
live persisted record bytes.  Lengths and SHA-256 values remain diagnostic
report fields.  The final witnesses identify these bases as
`direct_text_and_collection_values_with_sha256_report_fields` and
`direct_public_values_and_live_record_bytes`.  Optional or unsupported
projections carry explicit absent/unavailable states and are not presented as
coverage beyond the projection that was decoded.

## PPT dependency scope: resolved

The PPT witness now states that unaffected top-level CFB streams and storages
are covered by the complete stream/path/directory oracle, while dependencies
embedded inside the allowed `PowerPoint Document` stream are not separately
projected.  This is the accurate boundary for the current save contract and
does not claim semantic coverage for those embedded records.

## Terminal audit/schema synchronization: resolved

`audit.py` now expects
`common_container_open_replace_and_validate_finish_control`, loads the
regenerated `oracle-contract.json`, and validates semantic witnesses, frozen
identities, raw source-model evidence, and every named control's rejected
status and reason.  `analyze.py` and `prepare.py` enforce the same control and
identity contract.  The synthetic analyzer preflight passed against the final
qualification schema.

There is no substantive remaining blocker from the previous review.  The
qualification packet records all controls rejected and all three fixtures
matching their source normalized directory model across the 18 route/lane
qualifications.

# 0728 final probe review

This review covers the final expanded probe handoff at the 3163-line
`probe/src/lib.rs`.  The source now has raw directory extraction, source and
output unchanged-stream checks, DOC collection witnesses, PPT public and raw
slide witnesses, logical length witnesses, and live corruption controls.  The
reported build and 18 qualification runs are useful evidence; this review did
not run Cargo, native timing, or profilers.

The following are the remaining substantive qualification issues.

## Public expected metadata is not fail-closed

`raw_directory_oracle` treats the public format route as
`public_format_raw_differences_report_only`, and `enforce_policy == false`
leaves `raw_directory_policy_ok` true regardless of
`source_expected_normalized`.  The current qualification fixtures happen to
report `match_after_allocation_normalization` for source versus expected, but
the acceptance path would still admit a future public DOC/PPT output that
silently loses non-owned directory state, links, colors, or timestamps.  That
expected artifact is then the model for the common route, so expected-versus-
actual agreement cannot detect consistent loss.

The current source-backed save contract requires the public expected output to
preserve the source normalized raw directory image outside planner-owned
allocation fields.  Make that source-to-expected comparison a format gate (or
explicitly mark the route unqualified when it differs).  Keep Rewrite as the
policy-selected physical-normalization report, while Reuse remains a
source-model gate.

## Semantic witnesses still use digests as the comparison

The ordered DOC paragraph projection is compared exactly, and the PPT ordered
identity tuple is compared directly.  Auxiliary DOC collections are reduced to
`Debug` SHA-256 digests, while PPT list text, outline references/interactions,
slide text, notes/comments, and live records are represented by lengths and
SHA-256 digests.  Those fields are therefore still hash-based semantic
oracles.  The probe should compare the already available values or raw record
bytes directly and retain digests only as report fields.  This follows the
probe contract's explicit false-positive rule that hashes/lengths alone do not
prove preservation.

## PPT dependency coverage is overstated

`PptSemanticWitness::dependency_scope` says masters/media/fonts/embedded
storages are covered by complete CFB path/byte/directory checks.  `Pictures`
and separate storages are covered when they are unchanged streams, but master
and related records can live inside the allowed changed `PowerPoint Document`
stream.  The current witness does not inspect those records.  Either add the
small current-save check for dependencies present in the fixture, or report
them as absent/unavailable; do not publish the broader coverage sentence as a
proof of the changed document stream.

## Independent terminal audit must match the frozen schema

The current `analyze.py`/`prepare.py` path checks the control names and requires
every control to be rejected, and binds semantic witness/identity contracts.
The separate `audit.py` still expects the old
`common_container_open_stage_finish_control` scope and does not validate the
new `oracle_controls`/`oracle-contract.json` fields.  Synchronize it before
freeze, or clearly exclude it from terminal acceptance; otherwise the final
audit either rejects valid captures or can accept a report without the new
negative-control contract.  `qualify.py` alone also checks no control status,
so `prepare.py` must remain a mandatory gate.

The current qualification packet shows all listed controls rejected and all
three fixtures with a matching source/expected normalized directory image, so
these findings describe acceptance and claim boundaries rather than a failed
current fixture.  Until the public metadata gate, digest-only semantic fields,
and terminal-audit contract are resolved or explicitly scoped, retain the
result as qualified harness evidence rather than a fully closed current
baseline.

# 0781 CI repair re-review

Reviewed the frozen CI repair at `fdca3e6303` in the 0781 worktree. The review
covers `.github/workflows/perf-baseline.yml`,
`tools/validate_perf_default_matrix.py`,
`tools/test_validate_perf_default_matrix.py`, and
`tools/test_perf_workflow_policy.py`, against the checked default manifest,
the current performance harness, and the active CRUD coverage index.

## Verdict

No release-blocking defect remains in the repaired default-matrix check. The
earlier numeric-sample finding is resolved. The helper now requires `ns`, an
exact requested sample count, non-negative integer elapsed samples, and an
exact permutation of sample indexes (`validate_perf_default_matrix.py:399-428`).
It also compares the report's default cases, sample count, filesystem flags,
range simulation, and applicable corpus configuration with the manifest
(`validate_perf_default_matrix.py:299-345`). The focused tests exercise the
boolean/NaN/negative sample cases and unit, ordering, sample-count,
filesystem, and range mutations (`test_validate_perf_default_matrix.py:133-214`).

The identity contract is aligned with the current matrix: the manifest records
41 default cases and 213 full rows (`perf-regression-default-manifest-v1.json:8-10`),
and `Case::DEFAULT` currently contains 41 entries in `tools/perf-baseline/src/lib.rs:1639-1682`.
Full validation derives keys and the digest from that manifest; smoke validation
uses the requested tiny/compressible selectors. The workflow invokes the helper
after report generation for both modes (`.github/workflows/perf-baseline.yml:421-438`
and `:1646-1652`), and the full job remains schedule/manual-only
(`.github/workflows/perf-baseline.yml:1416-1419`).

The CRUD references in the repaired workflow consistently use the active v2
index for validation, copying, upload, and path triggers. The policy test also
requires the manifest-backed helper and rejects the former hard-coded 37/201
shape checks (`tools/test_perf_workflow_policy.py:537-591`). This preserves the
existing smoke scope and does not expand ODF/iWork coverage.

## Remaining scoped caveats

1. The manifest and both matrix/CRUD validators still describe the harness as
   `tools/perf-baseline/src/main.rs:Case::DEFAULT`, while the actual declaration
   is in `tools/perf-baseline/src/lib.rs`. The helper therefore checks a stale
   metadata literal rather than resolving the declaration from source. This is
   a provenance follow-up shared with the existing CRUD validator, not a defect
   in the 0781 matrix repair; correcting it should update the manifest and both
   validators together.
2. The workflow policy test checks the helper command, arguments, paths, and
   removal of stale row-count assertions, but does not execute the workflow or
   independently assert step ordering. The checked workflow places validation
   after each report-producing command, and the helper performs the runtime
   identity checks. This is a low-priority static-policy coverage limitation.
3. The helper establishes matrix identity and capture-shape invariants. It does
   not recompute every derived statistical field or independently bind the
   manifest to the Rust source. Those are existing comparator/provenance
   responsibilities and remain outside this repair's scope. The focused tests
   also do not add a dedicated malformed sink-bucket fixture, although the
   helper retains the existing bucket-key and sum checks.

The recorded CI-quality evidence in `validation-scope.md` reports the repaired
Python suite and six validation gates passing. This review did not run tests,
Cargo, native builds, or profilers.

# 0362ul94ha — XLSX join error classifier follow-up

This source-only follow-up closes the chained consumer return found during
independent review of 0362ul94h. It is based on the 0362ul94h source tip
`aded3daf77b02c107c5df0b86822e1f281d35a1e` and must be composed after that
repair.

## Chained classification

`litchi-ops` now handles `litchi_xlsx::Error::Join` explicitly. When the join
retains a typed `SourceChanged` normalization source it reports `io_error`; a
typed `Internal` source reports `internal_error`. A join with no normalization
source keeps the existing `corrupt_file` fallback. Any retained source is
classified by its existing XLSX rules, including refusals, patch conflicts,
limits, and nested shared semantic errors. The source chain and the
format-owned XLSX leaf are unchanged.

The ops test constructs a real pair of XLSX edits from different snapshots,
converts its resulting `JoinError` through the XLSX error and unified error
boundaries, and checks the retained typed leaf plus the unchanged fallback.
The 0362ul94h transaction test continues to cover the two semantic source
classes at the XLSX boundary; no wildcard classification or compatibility
alias was added.  The private branch coverage and the complete typed
`JoinError` source-chain check are recorded in 0362ul94hb.

## Validation status

No Cargo, build, rustdoc, runtime, or performance command was run. The
changed source is checked with `git diff --check`; the coordinator must rerun
the composed OOXML consumer and ops gates.

# ODS aggregate performance evidence

This directory will retain the reproducible process-level performance receipt
for the ODS numeric aggregate evaluator batch.  The capture protocol is
defined in [PLAN.md](PLAN.md).

The profile is intentionally paired with the semantic evidence in the parent
directory.  It records matched controls from the previous committed evaluator
and candidate-only aggregate cases separately, because the previous revision
does not have a valid implementation baseline for the seven new functions.

The final receipt will include the harness source, lockfile, profile-input
hashes, candidate freeze, build/compiler metadata, raw JSON/time files, source
manifests, independent verifier output, and a generated report.  No result is
valid until the root agent records a quiet source-freeze window and both lanes
have been captured under the same profile inputs.

The candidate matrix also includes small and large three-matrix `SUMPRODUCT`
rows so the K>2 scaled-product path is measured separately from the paired
matrix controls.

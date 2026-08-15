# Change 0137: plan-only native XLS numeric publication selectors

## Scope

Change 0137 adds two opt-in `tools/perf-baseline` selectors for the additive
forward-only native XLS numeric publication plan:

- `xls_numeric_plan_only_number_edit_save`
- `xls_numeric_plan_only_rk_mulrk_edit_save`

The Number selector edits deterministic `Untouched!E21` from `42` to `43`.
The packed selector edits one standalone `RK` value and both values in one
two-cell `MulRK` record. Both use the same corpora, edits, bounded sink and
untimed gates as the four selectors in [change 0135](0135-xls-numeric-source-publication.md).

## Production boundary

The selectors call `Transaction::commit_source_backed_plan`, which returns a
`SourceBackedPlanCommit`. The plan retains the immutable source, a validated
same-length CFB overlay plan and bounded numeric splice metadata. It does not
retain a reopened target `Snapshot`, complete target byte vector, or reversible
artifact `Patch`. It is therefore a forward-only publication contract; callers
that require patch/inverse semantics continue to use
`Transaction::commit_source_backed`.

Commit timing includes plan construction and the composed-source semantic,
security and fingerprint validation performed by the production API. Complete
publication through `write_to` is timed separately. Source ingress, expected
output preparation, full output reopen/readback, CFB topology/member identity,
no-op plus exact source/target fingerprint preflights, partial-sink,
unsupported/security refusal and the
real-producer `54016.xls` gate remain outside timing.

## Evidence contract

The `source.xls_numeric` evidence explicitly records:

- `target_artifact_retained_at_commit: false`;
- `target_artifact_materialized_at_commit: false`;
- `patch_or_inverse_supported: false`;
- zero `complete_target_materialized_bytes` at the commit boundary; and
- complete published bytes through `sink_bytes`, sink write counts and sink
  digests.

The plan still emits a complete CFB artifact at publication time. Composed
semantic validation may read and allocate a candidate `Workbook` model before
the plan is returned; zero complete-target-artifact bytes therefore do not mean
zero target-semantic allocation or a bounded total-memory claim. These fields
are evidence of the publication contract, not a memory measurement.

The selectors are correctness/descriptive evidence only. Change 0136 is the
committed before baseline for the four fixed-width numeric selectors; no
plan-only latency, allocation, RSS, I/O, or speedup claim is made until a
balanced CPU-pinned release ABBA comparison uses matched configuration and
records the relevant resource evidence.

## Review and gates

Focused selector tests, strict Clippy/deprecation gates, the full harness, and
an independent adversarial review must run after the production plan API is
frozen. No selector is added to the default 36-case / 198-record matrix.

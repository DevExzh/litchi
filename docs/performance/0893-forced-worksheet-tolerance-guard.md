# 0893 — force worksheet tolerance reports after full-cell scans

The producer XLSX full-cell benchmark now forces every worksheet's tolerance
report before accepting an iteration. Previously it checked only the workbook
aggregate, which includes already published stores. A future eligible streaming
route could leave those stores unmaterialized and conceal repairs or parse
errors from that check. This closes the qualification gap identified in
[0892](0892-current-producer-full-cell-profile.md); production parser eligibility
is unchanged.

## Guard contract

The runner preserves the timed open, worksheet enumeration, full stored-cell
visits, and canonical callback digest. After stopping the timer, finishing the
allocation region, and taking the process endpoint snapshot, it:

1. Forces each worksheet's `tolerance_report()` and rejects repairs or errors.
2. Checks the workbook aggregate, retaining package/catalog/shared-string checks.
3. Checks source freshness after all forced reports.
4. Validates the existing independent cell oracle before publishing a sample.

Typed worksheet errors remain in the error source chain with worksheet and
iteration context. Regression cases place a dirty unpublished worksheet at each
of the four positions independently, and separately require malformed worksheet
errors to retain their typed source. The new `guard_contract` is
`post-endpoint-per-worksheet-tolerance-v1`. Measured report vectors record forced
worksheet counts and clean results in elapsed sample order. Separate warmup
vectors retain execution order; the successful guard count includes both.

The fixed producer archive remains 22,739 bytes, SHA-256
`55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8`, with four
worksheets and 4,096 cells (3,686 numeric, 410 text). The canonical digest remains
`8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b`. Default case
selection and the static marker census are unchanged.

## Actual selected-scanner qualification

Test-only calls to the public raw scanner use the generated worksheet payloads,
`Rect::ALL`, default capabilities, and the same declared-byte processing ceilings
as the source-backed route. All four complete producer worksheets return
`NotEligible(UnsupportedStructure)`. Removing earlier metadata exposes these
successive barriers:

| Remaining producer markup | Scanner result |
| --- | --- |
| `sheetViews` | `UnsupportedStructure` |
| `sheetFormatPr` | `UnsupportedStructure` |
| `cols` | `Styles` |
| row `spans` | `UnsupportedStructure` |
| `pageMargins` | `UnsupportedStructure` |
| relationship sheet's `pageSetup` | `UnsupportedStructure` |
| all listed metadata removed | Eligible records, retaining 3,686 numeric and 410 deferred shared-string payloads across four sheets |

The current scanner code and controlled removals identify `sheetViews` as the
first barrier. These are actual raw-scanner observations, not instrumentation
of the enclosing timed `SourceWorksheet` route. `runtime_route_observed` remains
false. Malformed root markup must produce a typed MCE/XML error, distinct from
an eligibility refusal.

## Measurement limits and next work

This is correctness qualification, with no latency, allocation or RSS claim.
The report checks are outside the current timer and region metrics, but forced
materialization still contributes to whole-process RSS and profiler totals and
can affect allocator/high-water state in later iterations. In particular,
Callgrind's runner scope includes the guard. Historical 0885/0891/0892 captures
use the aggregate-only guard and remain valid for their recorded fallback
workload; they remain immutable and do not qualify the new contract. Future paired trials must rebuild both lanes with identical guards
and schedules.

Next implement and differentially qualify a narrow selected-route metadata
subset, preserving authoritative fallback for malformed, repairable, unknown,
styled, duplicated or misordered forms. EOF verification, resource limits,
MCE/x14ac behavior, dependencies, source fences and callback ordering remain
mandatory. The guard is one prerequisite, not proof of candidate correctness
or predicted performance. The broader non-iWork performance program remains open.

## Evidence

The [packet README](results/change-0893/README.md) provides exact commands and
scope. Attempts retain source copies, executed helpers, command environments,
logs, terminal statuses and passing test counts. The first formal attempt hit
`/tmp` user quota during dependency compilation, before tests; its log and
source copies remain, and its marked target was removed before moving builds
to workspace storage. An earlier agent-run exploration and its incorrect
malformed-dimension assumption are recorded separately and are not used as
qualification evidence.

Final [quality-v3](results/change-0893/attempts/quality-v3/quality-receipt.json)
passes all eight gates: formatting; four full-cell tests; three scanner-control
tests; two forced-report tests (including dirty stores at all four positions);
seventeen existing producer tests; warning-denied Clippy and rustdoc; and crate
boundaries. Test groups overlap, so these counts are not a unique-test total.
The source remained byte-identical throughout that attempt.

The [independent report checker](results/change-0893/qualification-v3/verification.json)
accepts the retained one-warmup/three-sample guard report, rejects twelve altered
reports, and rejects the historical aggregate-only report. Debug timings are
incidental test output, not performance measurements.
[Post-cleanup replay](results/change-0893/qualification-after-cleanup/verification.json)
also passes. The [cleanup receipt](results/change-0893/cleanup.json) records
removal of the dedicated 1.73 GB logical build target; the earlier failed tmpfs
target is also absent. No worktree was created, and shared targets and unrelated
worktrees were left intact. The [final review](results/change-0893/review.json)
binds the source, validation, report evidence, failure history and cleanup.

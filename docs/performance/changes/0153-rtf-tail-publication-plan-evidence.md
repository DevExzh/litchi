# Change 0153: RTF tail publication-plan evidence

Date: 2026-08-16

Status: opt-in matched publication-boundary and correctness evidence only; no
end-to-end, rich-format, allocation/RSS, physical-I/O, or ABBA latency claim.

## Scope

The standalone harness adds four selectors over the existing plain,
uncompressed RTF lifecycle corpus and its tiny/medium/large shapes:

- `rtf_logical_tail_commit_append`
- `rtf_logical_tail_plan_append`
- `rtf_logical_tail_commit_noop_save`
- `rtf_logical_tail_plan_noop_save`

The first pair appends the deterministic 4/64/256 one-run paragraph workload;
the second pair submits the exact empty append. The Commit selectors are the
matched `TailAppendCommit` controls. The PublicationPlan selectors use the
public bounded `TailAppendPublicationPlan` path. The existing two logical-tail
selectors remain unchanged and retain their historical staging/commit/
publication timing scope; these four selectors are the isolated publication
comparison.

The current harness has 295 selectable names. The default remains 36 cases /
198 records. No iWork selector or evidence is included.

## Timing boundary

Each iteration stages the equivalent control or candidate before the timer.
`elapsed_ns` is exactly the pre-staged publication-call interval around
`TailAppendCommit::write_to` or `TailAppendPublicationPlan::write_to`, after
the staged object and fixed 16 KiB sink exist. The four new selectors use a
`WindowedCountingSink`; the historical pair continues to use
`WindowedHashingSink` and is not publication-only evidence. Commit emits its
retained candidate, while PublicationPlan rechecks source/version/fingerprint/
limits, reserves execution budget, emits bounded source/insertion windows, and
performs final verification. These calls intentionally have asymmetric
validation and publication work; this is not a symmetric-work comparison.

The result's `source.rtf_tail_publication` keeps independent planning,
publication, reopen, and lifecycle vectors. `planning_ns` and
`publication_ns` have one entry per retained sample. `reopen_ns` and
`lifecycle_ns` are one-element preflight-only vectors: planning is construction
of the tail transaction plus Commit or PublicationPlan; reopen parses and
projects the exact output; lifecycle is all correctness work. The expensive
correctness gates run once outside the sample loop rather than repeating for
every sample, and all four vectors are excluded from `elapsed_ns`. The
`sink.rtf_tail_append` record explicitly reports source
retention, complete candidate retention for a changed Commit, zero distinct
target-candidate retention for exact no-op selectors, and the 16 KiB
publication window. These are ownership/accounting boundaries, not RSS,
allocator, physical-I/O, or total-memory measurements.

## Untimed acceptance gates

Before and alongside measured samples the harness proves identical output
bytes and SHA-256, semantic paragraph projection, exact no-op snapshot/bytes,
in-memory patch replay and inverse, durable deterministic-JSON apply and
inverse, stale/foreign source refusal, cancellation before output, sink
failure with exact partial progress, publication-window validation, output
limits, and source-version checks. The Commit control supplies the reversible
durable patch contract; the PublicationPlan candidate separately exercises
its source-proof, cancellation, sink-progress, and bounded-publication
contracts. A successful one-sample smoke is required before any future
measurement discussion.

The tranche is evidence-ready for a later review-approved release ABBA run,
but this change record does not run or claim that measurement. No rich-format
RTF, end-to-end save, allocation/RSS, physical-I/O, or generic CRUD
performance conclusion follows from these selectors.

## Reproduction

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --semantic-shape tiny,medium,large \
  --case rtf_logical_tail_commit_append,rtf_logical_tail_plan_append,rtf_logical_tail_commit_noop_save,rtf_logical_tail_plan_noop_save \
  --json target/perf/rtf-tail-publication-plan.json
```

The command is a correctness/evidence tranche only until the root agent
approves a clean release measurement after code review.

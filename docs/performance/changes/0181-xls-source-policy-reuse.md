# Change 0181: reuse immutable XLS source policy facts

Date: 2026-08-18

## Decision

Retain the native XLS plan-only fixed-width numeric optimization. An immutable
`cell_values::Snapshot` already performs a complete public `Workbook` open and
requires every privately inventoried worksheet to survive that open. It now
also records three private, content-free policy facts from the same validated
model: worksheet coverage, workbook-versus-worksheet protection
classification, and macro-free classification.

`commit_source_backed_numeric_plan` consumes those facts instead of reopening
the unchanged source through `Workbook::new` and repeating the same policy
walk. The composed target still receives an independent complete `Workbook`
open, worksheet-coverage and protection/macro checks, and public numeric
readback inside the CFB planner's fingerprint bracket. Container-level VBA
storage refusal, signed/encrypted/DRM ingress, exact source/version checks, CFB
reopen and range verification, fingerprints, emission hashing, partial-output
typing, and exact no-op refusal remain.

The first review found that a boolean protection fact collapsed the prior
distinct workbook/shared and worksheet protection messages. The frozen
candidate uses a three-state classification and real protected-workbook and
protected-worksheet fixtures to preserve the exact plan-time refusal behavior,
including exact no-op edits. Both final independent reviews are safe.

## Deterministic work reduction

Each effective plan now performs zero source policy `Workbook` reopens instead
of one. The one target semantic reopen and every CFB/publication pass remain.
For the measured corpora, the source Workbook payloads are 80,946 bytes
(Number) and 1,665 bytes (RK/MulRK), but those sizes are context rather than an
I/O counter: the removed parser consumes a seekable in-memory `Cursor`, while
the harness's source counters cover only owned-source ingress. No logical
`ReadAt`, physical-I/O, decompression, or copied-byte reduction is claimed.

The cache is three small private facts inside the already immutable snapshot.
It retains no additional document bytes, public handle, lock, executor, global
state, unsafe code, or dependency edge.

## Clean release A/B/B/A

Control `5b6eff538` and corrected candidate `d3df35ffe` were built as distinct
locked release binaries with SHA-256 `74663614ca...` and `001a771a1b...`.
Every retained report is clean, pinned to CPU 2, exposes one logical CPU, and
contains 20 warmups plus 500 samples. The order is strict
`A1 control, B1 candidate, B2 candidate, A2 control` for both existing
plan-only selectors. No harness selector, corpus, default, or matrix count
changed.

All four Number legs retain the same 16,995,840-byte source, one splice,
80,946-byte Workbook stream, output SHA-256 `f8f37064...`, source/target
fingerprints, 16,995,840-byte bounded sink, semantic readback, opaque streams,
and directory topology. Positive values are candidate reductions:

| Number metric | A1 -> B1 | B2 -> A2 | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| total p50 | 5.91% | 1.92% | 3.29% | 0.82% | accepted |
| total mean | 5.36% | 1.92% | 3.13% | 0.39% | accepted |
| total p95 | 6.60% | 0.48% | 5.89% | 0.28% | accepted |
| total p99 | 5.99% | 1.64% | 5.76% | 1.40% | accepted |
| commit p50 | 8.27% | 3.95% | 3.44% | 1.11% | accepted |
| commit mean | 7.58% | 3.64% | 3.52% | 0.59% | accepted |
| commit p95 | 8.17% | 2.02% | 6.51% | 0.25% | accepted |
| commit p99 | 9.63% | 2.75% | 7.06% | 0.01% | accepted |

The complete Number workflow and its isolated commit phase pass the
predeclared 5%/5%/10%/15% p50/mean/p95/p99 stability thresholds in both paired
directions. Publication is deliberately not claimed: the second direction is
near-neutral and slightly mixed, consistent with an optimization that changes
only commit-time source policy validation.

RK/MulRK total p50 is 8.57%/5.80% lower and commit p50 is 12.87%/10.83% lower,
but candidate p50/mean/p95/p99 drift reaches 7.38%/8.18%/16.40%/18.81%, and
control p95 drift is 10.84%. All RK/MulRK latency remains descriptive and
withheld. The deterministic `1 -> 0` source policy reopen applies to both
families independently of timing stability.

## Verification and scope

- 1,017 XLS library tests and the all-target suite apart from the established
  unrelated writer-encryption fixture failure;
- focused cached-policy, real protection/no-op, macro/storage, stale/foreign,
  topology/CLSID, partial-sink, and numeric semantic regressions;
- production Clippy with warnings and deprecations denied, formatting and diff
  checks;
- exact raw release semantic/source/sink vector checks and two independent
  final reviews.

This adds no selector or CRUD closure and leaves the matrix at 322 names and
the historical default at 36 cases / 198 records. No RK/MulRK, publication,
allocation/RSS, physical-I/O, cold-cache, atomic-save, formula/string,
structural, broad native XLS, or new real-producer performance claim follows.

Artifacts:

- [summary](../results/xls-source-policy-0181-summary.json)
- [manifest](../results/xls-source-policy-0181-manifest.json)
- raw Number and RK/MulRK A1/B1/B2/A2 reports listed in the manifest

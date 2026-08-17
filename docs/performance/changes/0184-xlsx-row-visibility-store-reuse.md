# Change 0184: XLSX row-visibility cell-store reuse

Date: 2026-08-18

## Decision

Retain a private, capability-specific handoff from the existing row-visibility
rewriter to the source-backed scalar worksheet snapshot. A changed direct
`hidden`-attribute commit now reuses the immutable parsed cell store instead of
running one complete scalar-cell parse over the rewritten worksheet. Candidate
XML grammar validation, a fresh row-visibility scan, source/cancellation
checks, patch readback, and complete publication validation remain.

The handoff is not a generic worksheet trust flag. `VisibilityRewrite` is a
private-field token constructed only by the bounded direct-`hidden` rewriter.
It borrows and checks the exact source byte slice for its lifetime; another
worksheet cannot consume it. Generic cell-value rewrites continue to parse the
complete candidate. Tests prove `Arc` store identity, unchanged numeric and
boolean cells, refreshed row state, full-parse parity, and foreign-source token
refusal.

## Measurement contract

The clean control is revision
`f41a76ecf9ffc38586ac5838748e3109c70a0468`, release binary SHA-256
`f16fac1f2028ae616f3dccf675ab6784b328e0f9b4744344c8ca761a4367d3ba`.
The clean candidate is revision
`9adbc95bd0b29b79915fae61b17010f78165d14a`, release binary SHA-256
`04408b8fa0c0011db989a3f6bdf7f37512dd8a0744ca460d5e216a9b230bd410`.
Fresh CPU-2-pinned processes run A1/B1/B2/A2 with 20 warmups and 500 retained
samples for each existing selector and medium/large shape.

The timer covers source-backed editor open, row-visibility staging, commit, and
sequential publication through the zero-retention hashing sink. `commit_ns` is
the primary metric because it contains the removed parse; complete
`elapsed_ns` is secondary and `publication_ns` is a shifted-work guard. The
predeclared p50/mean/p95/p99 same-implementation drift ceilings remain
5%/5%/10%/15%.

### Commit phase

| Shape / edit | Statistic | A1 control | B1 candidate | B2 candidate | A2 control | A1 -> B1 | B2 -> A2 | Control drift | Candidate drift |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| Medium / unhide 256 | p50 | 15.632 ms | 8.988 ms | 9.376 ms | 14.912 ms | -42.50% | -37.12% | 4.61% | 4.32% |
| Medium / unhide 256 | p99 | 23.641 ms | 11.606 ms | 13.302 ms | 20.214 ms | -50.91% | -34.20% | 14.49% | 14.61% |
| Large / hide one | p50 | 121.446 ms | 75.547 ms | 71.962 ms | 121.721 ms | -37.79% | -40.88% | 0.23% | 4.75% |
| Large / hide one | mean | 123.628 ms | 76.170 ms | 73.635 ms | 123.130 ms | -38.39% | -40.20% | 0.40% | 3.33% |
| Large / hide one | p95 | 137.360 ms | 88.974 ms | 84.780 ms | 135.388 ms | -35.23% | -37.38% | 1.44% | 4.71% |
| Large / hide one | p99 | 153.025 ms | 101.813 ms | 95.565 ms | 151.128 ms | -33.47% | -36.77% | 1.24% | 6.14% |
| Large / unhide 256 | p50 | 122.421 ms | 68.646 ms | 69.445 ms | 122.361 ms | -43.93% | -43.25% | 0.05% | 1.16% |
| Large / unhide 256 | mean | 125.596 ms | 69.289 ms | 70.333 ms | 125.030 ms | -44.83% | -43.75% | 0.45% | 1.51% |
| Large / unhide 256 | p95 | 143.984 ms | 74.378 ms | 76.550 ms | 142.740 ms | -48.34% | -46.37% | 0.86% | 2.92% |
| Large / unhide 256 | p99 | 156.300 ms | 77.065 ms | 80.527 ms | 152.202 ms | -50.69% | -47.09% | 2.62% | 4.49% |

The table contains only accepted commit statistics. Medium hide-one commit
latency is withheld because candidate drift is 7.35%-24.24%. Medium batch mean
and p95 are withheld because their control/candidate drift exceeds the
corresponding ceilings.

### Complete lifecycle

Large unhide-256 accepts all complete-lifecycle statistics: p50 is
21.70%-22.23% lower, mean 21.97%-23.36%, p95 24.53%-28.14%, and p99
24.53%-29.98%. Large hide-one accepts mean at 13.49%-17.46% lower, p95 at
11.45%-14.24%, and p99 at 10.91%-13.79%; its p50 is withheld because candidate
drift is 6.36%. Every medium total statistic is withheld. No independent
publication-phase latency claim is made.

## Correctness and work evidence

Deleting the environment revision and all timing vectors produces the same
canonical projection SHA-256
`2515c11ccbcbb58bc223803197c94ab5e7f64ffdb3dcf48b4cd7ba6475bf502d`
for all four legs. Output, semantic reopen, untouched-member, source/sink,
cache, and refusal evidence is therefore identical across implementations.
Medium records retain 204 logical source reads and one selected-worksheet read;
large records retain 209 and six. Unselected-worksheet reads remain zero.
These are logical `ReadAt` observations, not physical-I/O evidence.

The production verification is green for two focused token/store tests, 16
row-visibility integrations, 30 shared cell-value integrations, all 768 XLSX
library tests, formatting/diff checks, and strict all-target Clippy with
warnings and deprecations denied. Independent current-tree review is SAFE. A
strict public rustdoc attempt remains blocked by unrelated pre-existing broken
or private intra-doc links elsewhere in `litchi-xlsx`; none is in this change.

Change 0176 rejected a different conditional-formatting parsed-readback
experiment on a smaller corpus. This result does not revive that generic
handoff: it uses a source-bound row-only proof and retains only the claims that
pass the current workload's paired gates.

## Scope retained

No allocation/RSS, physical-I/O, decompression, cold-cache, throughput,
scaling, real-producer, formula, structural-row, insert/delete, or broad XLSX
CRUD claim is made. The capability remains direct visibility changes on
existing explicit row owners and keeps its existing MCE, macro, signature,
protection, relationship, stale-source, and bounded-publication refusals.

Artifacts:

- [summary](../results/xlsx-row-visibility-store-0184-summary.json)
- [manifest](../results/xlsx-row-visibility-store-0184-manifest.json)
- compressed raw A1/B1/B2/A2 reports listed in the manifest

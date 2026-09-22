# 0738 — PPT Sample restoration and startup argument controls

The same original probe binary changes secondary owner p50 by **+6.14%** when
its command repeats the already selected `--operation format` option. All nine
pairs exceed 5%; the bootstrap interval is +5.66–6.42%. Production and owner
allocation fields are unchanged. Startup argument/setup choices can affect the
later owner timer on this host. These measurements identify no unique allocator,
cache or address mechanism and do not explain the rejected 0735 optimization.

Restoring the original `Sample` source fields does not resolve the earlier
secondary observer gap. Keep the revised probe diagnostic-only, retain the 0735
rejection, and qualify complete invocation/startup profiles before transferring
timing conclusions or retrying a production candidate. The non-iWork goal remains
active. [Sealed evidence and replay instructions](results/change-0738/README.md).

## Two separately planned matrices

The main matrix compares the original 0735 probe (`archive`), the 0737 lifecycle
probe (`prior`), and duplicate invocations of the restored-field probe
(`restored-a` / `restored-b`). The supplemental argv matrix was planned and frozen
**after** the main results, before its own capture. No observations are pooled
across phases. Each matrix has 72 native processes (four arms, two fixtures,
nine rotated rounds, 50 samples after three warmups) and 24 allocation processes
(three rounds, one sample, no warmups). All run serially on CPU 12.

Together these are 192 measured processes, 7,200 native samples and 48 allocation
samples. Each output passes the exact prior full semantic oracle and the eight
preservation corruption controls. Commands, timestamps, raw reports and stderr
are retained. The scope is these two fixtures, this host, and warm serial edits;
it does not establish cold, concurrent or general workload performance.

Percentages below are medians of nine process-paired changes. Intervals use
10,000 seeded bootstrap resamples (seed 7338). Process p50 uses the midpoint;
tails use nearest rank. At 50 samples p99 equals maximum, so those are one signal.

## Restoring Sample fields

The new probe restores the exact original `Sample` fields, order, types and
derives. A serialization wrapper supplies retained-witness counts without
storing that count in each sample. Legacy still retains `Vec<Sample>`; strict
lifecycle buffering and witness policies remain unchanged. The source proof
checks that definition and five owner/oracle function blocks byte-for-byte.
Its fixed-block brace extractor is not a general Rust parser; it proves neither
ABI size nor runtime addresses. Wire schema and semantic checks are retained.

| Fixture / left → right | p50 change (95% bootstrap CI) | Mean | p95 | p99 = max |
|---|---:|---:|---:|---:|
| primary: archive → prior | +0.99% [+0.80, +1.51] | +1.97% | +0.21% | +0.16% |
| primary: prior → restored-a | -2.01% [-2.50, -0.86] | -2.67% | -3.08% | -0.08% |
| primary: restored-a → restored-b | +0.26% [-1.00, +0.51] | -0.03% | +0.28% | -0.17% |
| secondary: archive → prior | +7.10% [+5.19, +7.59] | +2.50% | +0.38% | +0.10% |
| secondary: prior → restored-a | +0.27% [-0.04, +0.36] | -0.07% | +0.52% | +1.04% |
| secondary: restored-a → restored-b | -0.06% [-0.29, +0.24] | +0.06% | +0.02% | +0.11% |

The secondary archive → prior p50 gap is +7.10%, with all nine pairs above 5%.
Restoring the fields changes secondary p50 by only +0.27% with an interval
crossing zero. Duplicate restored A/A p50 also crosses zero. This fails the
hypothesis that restoring these fields alone would remove the observer gap.
Primary restored p50 improves 2.01% versus prior; that does not establish full
observer equivalence. Primary last-ten/first-ten drift remains positive in every
arm; secondary remains negative. Sample-position plots retain the interior bands.

Secondary prior → restored has a +16.30% p99/max pair (round 0). Restored A/A
has −12.01%, +7.87% and +15.53% tail pairs (rounds 0, 1, 7). No other main
full-window mean or tail metric crosses the absolute 5% threshold.

## Same-binary startup arguments

`archive-standard` uses the original invocation; `archive-extra` appends the
redundant `--operation format`. `prior-explicit` supplies `--lifecycle legacy`;
`prior-default` omits it, selecting the same default. Each same-binary pair has
identical executable bytes and effective operation/lifecycle. Both added option
pairs have string lengths 11 and 6. Equal argument counts and lengths do not
establish identical process memory layout or all startup behavior.

| Fixture / left → right | p50 change (95% bootstrap CI) | Mean | p95 | p99 = max |
|---|---:|---:|---:|---:|
| primary: archive-standard → archive-extra | -1.49% [-2.12, -0.93] | -1.03% | -3.33% | -0.63% |
| primary: prior-default → prior-explicit | +7.04% [+6.65, +8.10] | +5.00% | +0.71% | +0.23% |
| primary: archive-standard → prior-default | -6.14% [-6.36, -5.85] | -3.14% | -1.04% | -0.86% |
| primary: archive-extra → prior-explicit | +2.20% [+1.76, +2.57] | +2.90% | +3.47% | +0.72% |
| secondary: archive-standard → archive-extra | +6.14% [+5.66, +6.42] | +1.84% | +0.42% | -0.10% |
| secondary: prior-default → prior-explicit | +7.78% [+7.27, +8.21] | +2.84% | +0.47% | -0.34% |
| secondary: archive-standard → prior-default | -1.05% [-1.70, -0.19] | -0.47% | +0.01% | +0.66% |
| secondary: archive-extra → prior-explicit | +0.72% [-0.01, +1.04] | +0.55% | -0.06% | +0.55% |

Both same-binary secondary contrasts exceed +5% in all nine rounds. Prior
explicit versus default also increases primary p50 +7.04% in all nine rounds.
Matching argument counts across binaries brings secondary gaps closer to zero,
but the primary default contrast remains −6.14% (eight of nine pairs below −5%).
Thus matched counts alone do not qualify the observers as equivalent.

Additional full-window flags are primary archive standard → extra tail round 0
−5.87%; primary prior default → explicit mean rounds 0/4/7/8
+5.11/+5.17/+5.59/+5.07%; primary archive standard → prior default tail round 0
−5.63%; secondary prior default → explicit tail round 2 +13.81%; and secondary
archive extra → prior explicit tail round 2 +13.95%. Every flagged paired value,
including window diagnostics, is retained in
[threshold-flags.md](results/change-0738/threshold-flags.md). Full precision,
ranges and intervals for all metrics are in each phase's analysis and summary.

## Allocation and validation

All five owner allocation fields are identical in every one of the 48 allocation
processes. These counters cover the owner measurement interval, not RSS or the
whole process heap.

| Fixture | Allocated bytes | Deallocated bytes | Calls | Peak live bytes | Retained bytes |
|---|---:|---:|---:|---:|---:|
| Primary | 11,391,768 | 11,001,624 | 5,658 | 1,992,885 | 390,144 |
| Secondary | 10,444,631 | 10,159,447 | 18,662 | 1,976,223 | 285,184 |

The final three builds pass formatting, release library tests (4 / 8 / 9),
all-target Clippy with warnings denied, documentation with warnings denied,
and release binary builds: 15 gates total. The first restored build failed
`--locked` because its copied lockfile still named package 0737. The failed
attempt is preserved; correcting only that package name allowed all five gates.
No new production test run is claimed. Twenty-four qualification processes,
17 offline negative-contract controls, and both synthetic schedule preflights
pass. The main preflight rejects altered statistics; the argv preflight rejects
an altered command.

The main matrix has independent custody, schedule, statistic and bootstrap
replay. The supplemental runner checks its scalar arithmetic against that
independent implementation; its grouping is shared within the runner. A separate
[results review](results/change-0738/results-review.md) records the independent
raw grouping check. The original source review remains frozen separately.

The 7,206-file production census and 34 goal/CRUD/accepted-design constraint
hashes remain unchanged. Host identity and some environment metadata are
inherited from 0737; toolchain versions, affinity and memory information were
refreshed. No claim depends on fresh dynamic-frequency telemetry.

All six temporary binaries were checked against their recorded hashes before
removing the owned target and binary directories. Both analyses, the independent
main audit, negative controls and source proof replay successfully after cleanup.
Raw Cargo test logs `build-0/1.log`, `build-1/1.log` and `build-3/1.log`
retain their original terminal blank lines; these are the only staged
whitespace-check exceptions.
The evidence manifest seals exact inventory and bytes, including all three
lockfiles. The unrelated API design draft is excluded from this batch.

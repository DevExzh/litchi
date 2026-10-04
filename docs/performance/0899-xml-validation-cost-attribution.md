# 0899 — XML validation cost attribution

The retained evidence does not justify another duplicate-check implementation.
Two baseline allocation profiles each attribute 6,106 scan-closure allocation
calls per scan to stacks containing the duplicate checker, but its native
sampled self share is only 0.7515–0.8327%. These different populations do not
establish a latency saving. Production is unchanged; this batch runs only
source audits, offline annotation, custody checks, and deterministic reduction.
iWork is excluded.

## Evidence and population boundaries

[The packet](results/change-0899/README.md) retains exact inputs, helper hashes,
reconciliations, failures, and independent review. Current source identity
matches all 10,771 selected files from the 0894 manifest. Production crates
also match the older 0892 profile revision. Three harness files changed since
0892, including the forced worksheet guard introduced in 0893; therefore the
0892 CPU profiles and 0895 allocation profiles are contextual populations,
not a matched performance comparison.

The two 0895 baseline Heaptrack exports each cover 20 scans. Native 0892 data
covers three whole processes, each with 2,000 measured scans and ten warmups;
poor stack unwinding limits interpretation to leaf samples. Two 0892 Callgrind
runs cover ten scans each with no warmup. Guest instruction totals include
runner/oracle work and are not native cycles or elapsed time. Retained
Callgrind brk-segment warnings remain part of the original capture evidence.

## Allocation attribution

Both baseline exports reproduce these integer counts exactly:

| Stack population | Calls per process | Calls per scan |
| --- | ---: | ---: |
| Entire process | 673,663 | — |
| Scan closure | 518,660 | 25,933 |
| Duplicate checker anywhere in process stack | 144,432 | — |
| Duplicate checker in scan closure | 122,120 | 6,106 |
| Closure duplicate checker, nearest codec owner `start` | 85,560 | 4,278 |
| Other closure duplicate checker | 36,560 | 1,828 |

Duplicate-check stacks account for 23.5453% of closure allocation calls. This
is stack-presence attribution, not source-line attribution or a count of
allocations proven removable. Across all closure allocations, nearest MCE
`start`, other codec, and no-codec owners partition 343,320, 2,060, and 173,280
calls respectively, reconciling to 518,660.

## CPU context

Native ranges below span three children; they are not confidence intervals.
Exclusive guest counts repeat exactly in both Callgrind runs, whose total
instruction counts are 717,244,483 and 717,348,440.

| Selected function | Native leaf period share | Exclusive guest Ir per run |
| --- | ---: | ---: |
| Duplicate checker | 0.7515–0.8327% | 6,700,360 |
| Bounded output reserve | 0.9685–1.0827% | 18,127,870 |
| MCE `start` | 12.9937–13.4970% | 39,019,560 |
| Quick XML iterator `next` | 1.2730–1.4754% | 25,556,590 |
| MCE escaping | 0.0276–0.1669% | 1,335,010 |
| MCE `expand_parts` | 2.9316–3.6182% | 10,415,520 |
| QName predicate | 0.3337–0.5274% | 8,686,020 |
| NCName predicate | 0.9186–1.8320% | 10,172,600 |

The reducer partitions exact leaf symbols and exclusive annotation rows,
including an explicit remainder. All counts reconcile within their own
population. Duplicate checking is about 0.934% of exclusive guest instructions.
Its self cost excludes allocator and other callees, so this is not a hard
upper bound on potential latency improvement.

Inclusive Callgrind rows are retained separately: run 00 has duplicate
checking at 1.8168%, iterator `next` at 5.38%, MCE `start` at 30.3413%, and
MCE `process` at 46.4626%. Those rows overlap and must not be added or mixed
with exclusive, native, or allocation percentages.

## Source findings and decision

[The duplicate-check audit](results/change-0899/source-audit.md) reviews the
first-32 checked iterator, later map transition, malformed-value precedence,
error positions, fused termination, and five mirrored implementations. Prior
0794–0806 candidates repeatedly traded short-input improvements for replay,
state, allocation, or protected malformed-input regressions. Another inline
32-entry design would add substantial iterator state and map-transition work.
Its priority remains low without evidence that resolves those failures.

[The MCE audit](results/change-0899/mce-start-audit.md) also rules out treating
unqualified-attribute validation reuse as new: 0890 already tested that
mechanism, and its boolean fit existing padding without layout growth.
A distinct source hypothesis is to return borrowed QName parts from lexical
validation, eliminating `expand_parts`' subsequent delimiter scan. Retained
Callgrind paths cover 197,460 expansions per ten scans, including 53,745 writer
calls, but do not isolate the redundant split or demonstrate qualifying savings.
This hypothesis remains deferred, with no prototype or adoption claim.

Any future qualification must preserve full lexical validation before prefix
resolution, exact invalid-name errors, Unicode and malformed QName handling,
opaque writer paths, namespace rules, directive ordering, duplicate checks,
and resource limits. Source redundancy alone does not establish emitted
machine-code redundancy. New evidence must justify a candidate before the
existing correctness and paired native gate: p50 ratio at most 0.98 with an
upper confidence endpoint below 1.0.

## Verification and cleanup

The 0895 strict allocation custody replay passes. The original 0892 strict
verifier fails because two raw Callgrind data files have mode 0664 instead of
0600; both retain their exact size, SHA-256, and non-executable status. No
historical file was changed. The bounded content replay permits only those two
mode differences and runs the original semantic/custody checks; it passes
while explicitly preserving `strict_historical_verifier_pass: false`.
The first adapter failed during Python module registration before producing
output; its helper and both failure records are retained.

Independent review passes with 33 exact artifact bindings and separately
recomputed allocation, native, and guest totals. A fresh reducer replay is
byte-identical to canonical analysis (SHA-256
`7dc7cca5f7fa426e48f3dc4c9f90e54f67299ab1b47917835127092f93fbfc33`).
The root closure receipt binds the final report and packet; it is distinct
from the independent numerical review. No production build, test suite,
workload, or profiler capture was needed for this documentation batch. No
worktrees were created. The owned 154,317-byte duplicate replay output and its
temporary directory were removed; unrelated workspace files were preserved.

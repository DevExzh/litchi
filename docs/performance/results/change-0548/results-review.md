# 0548 result review

The 0548 scalar checkpoint candidate is rejected for adoption. The frozen
native requirement is not met, and the independent required profile gates
regress in both repeats. The allocation and malformed public guard gates pass,
but those passes do not override the failed primary and profile gates.

The reviewed candidate source is
[`candidate-sources/file.rs`](candidate-sources/file.rs), SHA-256
`c773ab45a0d6b0fe0918406432933e52f7e7e2432d0897c3a4d3f877081ec427`. Its
patch SHA-256 is
`4a4168a097d6490ba399b71101ecb05d3ede7dec29273ef26c8bf1c4163d8d65`.
The canonical input digests are recorded in
[`adverse-review.json`](adverse-review.json). The comparison, guard,
profile, hardware, instruction, source-manifest, plan, and preparation
digests there are the inputs reviewed here.

## Gate results

| Gate | Result | Evidence |
| --- | --- | --- |
| Native primary XLS p50 | **Fail** | `xls_source_backed_open_one_cell`, repeat 1 changes from 102,760 ns to 104,935 ns (+2.116582%); the requirement is at least 3% improvement in both repeats. Repeat 2 improves 3.701809%. |
| Allocation guard | Pass | Calls, allocated bytes, and incremental region peak remain unchanged under the recorded allocation comparison. |
| Required profiles | **Fail** | XLS-owned constructor inclusive Ir rises 2.865078% / 2.891198%; XLS-owned `collect_exact` self Ir rises 5.855415%; CFB few-large self Ir rises 5.871289%, in repeats 1 / 2 as applicable. |
| Public malformed guard | Pass | All 12,800 samples match the exact oracle. The maximum same-invalid ratio is 2.233526x and the maximum baseline-valid ratio is 1.312102x, below the 4x and 2x limits. |
| Candidate quality run | Candidate only | The candidate run reported all 15 checks passing across 4,386 tests. The final restored-source quality session was still running when this review was written; the candidate count is not used as final disposition evidence. |

The complete native primary matrix was:

| Case | Repeat | Baseline p50 (ns) | Candidate p50 (ns) | Change | Required result |
| --- | ---: | ---: | ---: | ---: | --- |
| `xls_source_backed_open` | 1 | 104,521 | 96,290 | −7.874972% | Pass |
| `xls_source_backed_open` | 2 | 103,710 | 100,260 | −3.326584% | Pass |
| `xls_source_backed_open_one_cell` | 1 | 102,760 | 104,935 | +2.116582% | **Fail** |
| `xls_source_backed_open_one_cell` | 2 | 107,515 | 103,535 | −3.701809% | Pass |
| `xls_owned_source_open` | 1 | 96,635 | 88,930 | −7.973302% | Pass |
| `xls_owned_source_open` | 2 | 96,545 | 93,520 | −3.133254% | Pass |
| `xls_owned_source_open_one_cell` | 1 | 97,905 | 92,751 | −5.264287% | Pass |
| `xls_owned_source_open_one_cell` | 2 | 99,596 | 95,830 | −3.781276% | Pass |

## Row review coverage

[`adverse-review.json`](adverse-review.json) retains every canonical row
with its original fields and an individual review string:

* 30 main matched adverse rows and 61 main same-build variation rows;
* 87 public guard adverse rows and 14 public guard same-build drift rows;
* 12 profile diagnostic increases above 5%, plus the six required profile
  gate failures, including the two sub-5% constructor increases;
* no profile core metric with absolute repeat variation above 5%; and
* two matched and two same-build whole-child hardware rows above 5%, all
  context-switch decreases.

The main timing adverse rows include coarse system-time samples, repeat spread,
and maximum or tail observations. The same-build rows are retained as repeat
drift and are not treated as candidate gains. Guard rows cover valid input,
root and later cycles, early termination, reserved markers, invalid indices,
and excess chains at both 128 and 16,384 sectors. A positive malformed row is
consistent with speculative traversal followed by authoritative replay; a
negative drift row is still retained as repeat variation. Exact error/value
agreement and the aggregate guard envelopes do not erase the individual
observations.

Hardware counters are whole-child diagnostics covering setup, copies, queries,
oracles, destruction, and report construction. The only matched changes above
5% are context switches falling from 86 to 81 and from 71 to 66. Baseline and
candidate same-build context-switch counts also fall from 86 to 71 and 81 to
66. These scheduling observations cannot establish operation-local speedup,
and no positive hardware event increase above 5% was observed.

## What the measured assembly shows

The canonical instruction comparison reports 313 collector instruction rows
for the baseline and 280 for the candidate, while explicitly treating static
instruction counts as diagnostic. The offsets below are from the
`profile-r1-xls-owned` XLS-owned group, repeat 1, timed dump 1:
`baseline/profile-r1-xls-owned.callgrind.1`
(SHA-256 `15a0ae520b1b7c8a4dbe35867b48149b94cb578d89f695442d95cfa1aa3eeecb`)
and `candidate/profile-r1-xls-owned.callgrind.1`
(SHA-256 `69de2c450f63f9344085d3774bb69dafa06bc3b1c60615e0d888106c1638ef0a`).
The baseline collector’s hot loop contains the visited-map bound calculation,
word load, bit test, and word OR around offsets `0x1c0`–`0x1fd`. The candidate
removes that per-sector map sequence from the ordinary path, but the
replacement loop around `0x291`–`0x345` adds:

* a checkpoint equality compare and replay branch;
* distance and power state updates;
* checked-bound handling expressed as compares, increments, conditional moves,
  and a branch back to the loop.

The successor load and ENDOFCHAIN/reserved-marker checks remain in both loops
before advancing; they are retained structural work required by the existing
error contract, rather than candidate-added work.

The candidate still pays the existing visited-map preparation and zero-fill
before entering the loop. Therefore the measured path keeps the fixed map
setup and adds scalar checkpoint scheduling work on every traversed sector;
the removed in-loop map operations did not compensate for that schedule on
these inputs. The resulting Callgrind collector self Ir rises by 327,145
(5.871289%) for CFB few-large and 327,970 (5.855415%) for XLS-owned. The
XLS-owned constructor inclusive Ir rises 2.865078% and 2.891198% in the two
repeats. These measurements explain the direction of the regression without
claiming a particular cache, branch-prediction, or wall-time mechanism.

The cold authoritative replay is necessary for exact diagnostics and preserves
the existing error precedence, allocation order, and scratch-buffer boundary.
It does not provide a measured valid-path benefit. The candidate-only Rust
differential and boundary tests are useful correctness evidence, but they
cannot turn the semantic or arithmetic model into a native performance proof.

## Next direction

The exact measured baseline is the retained production direction for this
batch. The next OLE2/OOXML experiment should first attribute the baseline
visited-map loop and the candidate checkpoint loop with a narrowly scoped
measurement, then reduce or remove per-sector checkpoint state while retaining
preflight, reservation order, resource labels, zero-fill, exact error
precedence, and authoritative replay. It must earn fresh matched primary,
profile, allocation, and public-guard evidence before adoption. No gain is
inferred from the current instruction count, proof, or guard envelope.

ODF work remains deferred until the OLE2/OOXML optimization goal is complete.

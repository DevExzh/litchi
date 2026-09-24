# Lookup performance capture analysis

The authorized frozen-source run used runner session `69932`, three warmups, and fifteen fresh child samples in both `evaluate` and `parse-evaluate`. It retained 990 baseline rows (33 matched controls) and 3,630 candidate rows (33 controls plus 88 lookup cases), for 4,620 timed rows. The candidate freeze is [`gates/freeze.json`](../../gates/freeze.json) (SHA-256 `53261588…c86fa60`); the contract hash is `b112d66d…889e6aaf` and the isolated gate lock hash is `58b4be6c…1340a3e3`.

## Capture integrity

Candidate preflight passed all 121 named cases and the complete per-case read contract. The retained `verify.py` passed 990 baseline and 3,630 candidate records across both phases. Source/profile hashes stayed stable before and after timing, allocator bytes balanced, all matched control accounting sets and output checksums were unchanged, and target cleanup was recorded. Baseline and candidate binary identities are recorded with SHA-256 `6c399d71…ac0f2b24` and `00f01f3d…498ed2b`, respectively.

After the five rejected harness preflights retained in the diagnostic directories, the first full-capture launch was setup-only: a duplicated directory component made the relative freeze path invalid. It produced no build, preflight, or timed row. The command/error is retained under [`diagnostic-freeze-path-typo-20260920T1400Z/`](diagnostic-freeze-path-typo-20260920T1400Z/); the corrected single run is the 4,620-row result.

## Matched controls

Across 66 control/phase groups, normalized latency deltas range from `-5.2266%` to `+12.6005%`. The only positive latency flag above +5% is `reference-conditional-256x4-sumifs` evaluate: `84,480.5` to `95,125.5` ns/repeat (`+12.6005%`). The one negative latency outlier below −5% is `scalar-control-average` parse-evaluate (`-5.2266%`).

For SUMIFS evaluate, work is 1,839/repeat, resolver reads 1,792/repeat, allocator calls 42, and bytes/repeat 52 on both builds. An independent unpaired median-ratio bootstrap using exact per-sample `elapsed_ns_p50 / repeat` values, seed `20260920`, and 10,000 resamples gives the descriptive 95% interval `[ +8.0045%, +17.2941% ]`. This interval describes sample variability; the receipts do not establish a cause for the latency observation.

RSS deltas range from `+2.3305%` to `+14.7577%`; 36 groups exceed +5% and none is below −5%. RSS is reported as an observation only.

The retained root audit reports 37 positive review triggers in total: 36 RSS groups and this one latency group. The existing SUMIFS control remains useful as a matched evaluator control; lookup-specific additions do not alter its resolver reads, work, allocator counters, or checksums.

## Lookup lanes

The 88 lookup lanes cover ADDRESS A1/R1C1 formatting, lazy and projected CHOOSE, exact linear and sorted approximate HLOOKUP/LOOKUP/MATCH/VLOOKUP scaling, selected-value reads, zero-read INDEX/OFFSET/INDIRECT descriptors, matrix outputs, union/shape refusals, resource and cancellation failures, lazy projected consumers, and dynamic projected INDIRECT descriptor scaling.

| lane | evaluate ns/repeat | parse-evaluate ns/repeat | reads | work |
| --- | ---: | ---: | ---: | ---: |
| `lookup-match-exact-8-end` | 2,241.875 | 2,664.375 | 8 | 27 |
| `lookup-match-exact-64-end` | 4,867.500 | 5,402.500 | 64 | 84 |
| `lookup-match-exact-256-end` | 13,740.000 | 14,130.000 | 256 | 277 |
| `lookup-hlookup-exact-8-start` | 2,215.625 | 2,675.969 | 2 | 26 |
| `lookup-hlookup-exact-64-end` | 5,167.500 | 5,655.000 | 65 | 90 |
| `lookup-hlookup-exact-256-end` | 13,900.000 | 15,140.000 | 257 | 283 |
| `lookup-vlookup-exact-8-start` | 2,196.281 | 2,670.031 | 2 | 26 |
| `lookup-vlookup-exact-64-end` | 5,172.500 | 5,675.000 | 65 | 90 |
| `lookup-vlookup-exact-256-end` | 13,960.000 | 15,250.000 | 257 | 283 |
| `lookup-indirect-projected-dynamic-8` | 29,425.781 | 31,207.969 | 16 | 443 |
| `lookup-indirect-projected-dynamic-64` | 210,166.656 | 218,498.281 | 128 | 3467 |
| `lookup-indirect-projected-dynamic-256` | 834,372.875 | 873,716.531 | 512 | 13835 |
| `lookup-match-projected-invariant` | 5,530.000 | 6,383.750 | 2 | 70 |
| `lookup-match-position-sensitive-key` | 3,157.500 | 3,822.500 | 5 | 32 |
| `lookup-match-nested-munit-key` | 10,363.750 | 11,453.875 | 5 | 141 |

The dynamic INDIRECT lanes read 16, 128, and 512 cells at 8, 64, and 256 projected coordinates, with work 443, 3,467, and 13,835; this is linear descriptor/selector work in the retained fixture. The nested MUNIT MATCH lane returns distinct position-sensitive results 1/2 and reads five cells, while the invariant projected MATCH lane reads two. Descriptor, union-refusal, and resource lanes retain zero resolver reads. Cancellation records one first attempted read across four sticky-cancel repeats; its normalized integer report field is zero by floor division.

## Limits and retained artifacts

The environment manifests identify a 32-CPU AMD EPYC 9R45 host and matching Rust/Cargo/libc/toolchain, but the runner did not retain a load or CPU-affinity snapshot. A process check immediately before launch found no competing Cargo, rustc, or benchmark process; this is not an isolated-host claim. Latency and RSS differences therefore remain descriptive observations. The profile measures evaluator and typed resource paths, not save, recalculation, native producer acceptance, or cross-platform timing identity.

The complete raw receipts are under [`baseline-635fd2e1348b621426b50909cbd5765c91837306/`](baseline-635fd2e1348b621426b50909cbd5765c91837306/) and [`candidate-final/`](candidate-final/). The generated tables are [`performance-report.md`](performance-report.md), the machine-readable analysis is [`capture-analysis.json`](capture-analysis.json), and the final flat retained manifest is [`retained-files.json`](retained-files.json).

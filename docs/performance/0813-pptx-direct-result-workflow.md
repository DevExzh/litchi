# 0813 — direct PPTX event-result matching

The one-file optimization is retained after fresh public-workflow qualification.
Large-capture p50 improves 10.544% and large lifecycle 7.269%; six rows meet
the frozen benefit threshold, with no latency veto or allocation increase.

[0812](0812-pptx-scanner-instruction-localization.md) localized the leading
fresh native scanner samples to payload copies in the inline `Result::map_err`
at `notes/codec.rs:389`. This trial changes only that result dispatch: each
existing event arm is wrapped in `Ok`, and an explicit `Err(error)` returns
`xml_error(error)`. Reader configuration, event-arm bodies, namespace timing,
limit/refusal ordering, and the entire buffered differential test module remain
byte-identical. No public API, dependency, concurrency, or ownership contract
changes.

The baseline is commit `03305a82aa`. Its complete 9,196-file source census
matches sealed 0810 after, allowing reuse of the six production quality gates
(1,241 tests, zero failures, three ignored). Both probe legs receive fresh
format/test/Clippy gates, and the candidate receives all six production gates.
All 35 previously read normative inputs are unchanged. iWork is excluded.

Before application, all eighteen baseline public-workflow reports matched the
sealed source/output/semantic oracles. Ordinary and profile scanner disassembly
confirm that all three 40-byte pre-dispatch payload copies disappear before
native trials. In both ordinary and profile binaries, the scanner shrinks
from 2,616 to 2,320 bytes and the inspected window has zero vector moves
instead of twelve. These are static mechanism observations, not timing results. The fresh
comparison covers six counterbalanced blocks across 18 capture/commit/lifecycle
rows, two paired allocation blocks, and four owner-scoped Callgrind publications:
310 reports/6,718 measured outputs including qualification. Historical timings
are not pooled.

The frozen policy requires a capture or lifecycle p50 improvement of at least
3% with bootstrap upper endpoint below one, no p50 case above 1.05 with lower
endpoint above one, and no increase in paired allocation-call/byte/net-live/
peak medians. Quantiles use nearest rank; the bootstrap uses six-block medians,
10,000 resamples, seed 813813, and sorted endpoints 250/9749. Tail latency and
RSS observations remain explicit diagnostics. Guest instruction counts are
mechanism evidence, not native phase fractions or causal savings.

[Protocol](results/change-0813/protocol-review.md),
[candidate design](results/change-0813/candidate/design.md), and
[source review](results/change-0813/source-review.md) state the scope and
semantic constraints. The broader OLE2/OOXML performance goal remains active.

Fresh measurements pass the frozen numerical policy: six eligible benefits,
zero latency vetoes, and all 144 paired allocation comparisons exactly equal.
Independent review agrees; the retained source disposition is recorded separately.

The table reports medians of six process p50s in milliseconds. The ratio is
the median of six paired ratios, so it need not equal the quotient of the
two displayed medians. Intervals are the frozen bootstrap ratio intervals.

| Shape | Workflow | Before ms | After ms | After/before | 95% interval |
| --- | --- | ---: | ---: | ---: | --- |
| tiny | capture | 0.233002 | 0.229296 | 0.980929 | 0.979133–0.985976 |
| tiny | commit | 0.209356 | 0.207601 | 0.990727 | 0.986880–0.995786 |
| tiny | lifecycle | 1.407942 | 1.408527 | 0.998801 | 0.990806–1.002944 |
| medium | capture | 0.447577 | 0.425953 | 0.952520 | 0.947184–0.959294 |
| medium | commit | 0.292992 | 0.290751 | 0.991866 | 0.988700–0.992901 |
| medium | lifecycle | 1.999650 | 1.981175 | 0.989083 | 0.987453–1.000920 |
| large | capture | 17.926675 | 16.045099 | 0.894556 | 0.882107–0.899554 |
| large | commit | 1.284752 | 1.260447 | 0.980945 | 0.980036–0.981292 |
| large | lifecycle | 27.785777 | 25.769371 | 0.927313 | 0.920132–0.933996 |
| vendor | capture | 0.528198 | 0.507833 | 0.961294 | 0.954689–0.970224 |
| vendor | commit | 0.321452 | 0.319337 | 0.996120 | 0.985896–0.998960 |
| vendor | lifecycle | 2.157126 | 2.136136 | 0.990393 | 0.987771–0.992217 |
| unicode-vendor | capture | 0.532003 | 0.515273 | 0.968801 | 0.962577–0.972043 |
| unicode-vendor | commit | 0.322111 | 0.319927 | 0.993234 | 0.990443–0.995369 |
| unicode-vendor | lifecycle | 2.159607 | 2.145652 | 0.995860 | 0.989817–0.997549 |
| valid-4attr | capture | 0.507393 | 0.492403 | 0.966499 | 0.959556–0.970461 |
| valid-4attr | commit | 0.315092 | 0.313157 | 0.994941 | 0.986502–0.997456 |
| valid-4attr | lifecycle | 2.115381 | 2.099656 | 0.992306 | 0.985593–0.995650 |

Four individual block p99 increases exceed 5%: tiny commit block 5 (+10.328%),
medium lifecycle blocks 3/5 (+7.835%/+26.331%), and valid-4attr lifecycle
block 4 (+8.553%). The corresponding paired-median p99 changes are
−0.093%, +1.967%, and −0.715%. All are retained as diagnostics; no general
tail-latency improvement is claimed. Eighteen native spread flags remain
(nine p99 and nine process-RSS groups). Allocation-lane diagnostics have no
regression or spread flags.

Process RSS medians are shown below in KiB. No median increase exceeds 5%;
process variation and single-block changes are diagnostic, not an allocation
or memory-saving claim. The four paired allocation metrics—calls, allocated
bytes, net live bytes, and peak above entry—are equal in both blocks of all
eighteen cases.

| Shape/workflow | Before RSS KiB | After RSS KiB |
| --- | ---: | ---: |
| tiny/capture | 5170 | 5016 |
| tiny/commit | 4920 | 4966 |
| tiny/lifecycle | 5088 | 5176 |
| medium/capture | 5164 | 5148 |
| medium/commit | 5388 | 5356 |
| medium/lifecycle | 5348 | 5340 |
| large/capture | 18636 | 18572 |
| large/commit | 18570 | 18570 |
| large/lifecycle | 18634 | 18604 |
| vendor/capture | 5420 | 5406 |
| vendor/commit | 5388 | 5516 |
| vendor/lifecycle | 5484 | 5514 |
| unicode-vendor/capture | 5462 | 5388 |
| unicode-vendor/commit | 5484 | 5452 |
| unicode-vendor/lifecycle | 5506 | 5478 |
| valid-4attr/capture | 5484 | 5484 |
| valid-4attr/commit | 5420 | 5452 |
| valid-4attr/lifecycle | 5452 | 5486 |

Paired RSS increases above 5%:

- medium/lifecycle, block 0: 5168 → 5452 KiB (+5.495%).
- medium/lifecycle, block 3: 5132 → 5516 KiB (+7.482%).
- tiny/capture, block 2: 4900 → 5240 KiB (+6.939%).
- tiny/commit, block 4: 4856 → 5112 KiB (+5.272%).
- tiny/lifecycle, block 3: 4920 → 5176 KiB (+5.203%).

The measured host is AMD EPYC 9R45, x86_64 Linux, pinned to logical CPU 12.
Rust/Cargo 1.95.0 use the frozen release profile (optimization 3, thin LTO,
one codegen unit, unwind); native builds have no instrumentation feature or
frame-pointer override. The six deterministic public fixtures and probe are
byte-identical to the sealed 0806 harness; source/output hashes and complete
verification objects match its oracles. Each timing process uses three warmups
and thirty measured samples. This experiment does not cover cold-cache,
cross-format, concurrent, remote-source, or universal document performance.

The six fresh candidate gates cover `litchi-pptx` formatting, all-feature /
all-target checking, tests, warning-denied Clippy and rustdoc, and the repository
crate-boundary checker. Tests pass 1,241 cases across 85 result groups, with
zero failures and three pre-existing ignored cases. Each probe leg passes its
own formatting, 36 all-feature tests, and warning-denied Clippy. Exact baseline
source identity permits reuse of the sealed 0810 production checks. The
buffered differential oracles and test module remain byte-identical.

| Constraint | Evidence in this change |
| --- | --- |
| ADR 0001/0002/0024: API and crate ownership | One private codec function; no public API or dependency change. |
| ADR 0003: snapshots, atomic edits, reversible patches | Existing public capture/commit/lifecycle semantic verification passes all measured outputs. |
| ADR 0005/0031: bounded resources and explicit execution | Same limits and execution model; all 144 allocation comparisons equal; no concurrency introduced. |
| ADR 0006/0008: preservation, refusal, verification | Event bodies, namespace timing, checked attributes, error conversion and order remain unchanged; existing refusal/differential tests pass. |
| ADR 0010/0011: physical ownership | No archive or physical-package change; source and output identities remain exact. |

Reproduce numerical and custody checks with
`python3 -B docs/performance/results/change-0813/validate.py --final`.
The [packet README](results/change-0813/README.md) lists the original fresh-build
and capture sequence. Archived receipts bind all six binaries before cleanup;
post-cleanup replay uses those receipts and retained raw measurements. The
[independent code-generation review](results/change-0813/codegen-review.md),
[profile review](results/change-0813/profile-review.md), and
[result review](results/change-0813/results-review.md) delimit the claims.

All four scoped Callgrind publications pass exact-owner, full self-cost sum,
termination, source, and semantic checks. Scanner self Ir is 22,295,758 before
and 16,724,054 after in both repeats. Reader calls remain 282,612 and namespace
push self Ir remains 13,019,166. Whole-owner Ir is 537,151,032/537,174,712 before
and 531,661,794/531,570,923 after. This is guest instruction attribution;
software cryptography and code generation under Valgrind differ from native
execution. The Ir difference neither predicts nor explains a native percentage
speedup by itself. Native timing and static assembly provide separate evidence.

Further work requires refreshed current-source attribution before selecting
another hotspot. This trial does not justify namespace/attribute fusion, whose
error-ordering risks were documented in 0812, nor a broad parsing rewrite.

Final aggregate validation passes before and after cleanup. The retained
[decision](results/change-0813/decision.json) and
[disposition](results/change-0813/disposition.json) bind the outcome to this
one-file candidate. Cleanup removes 11,216 owned files / 4,153,889,847 logical
bytes after verifying all six executable identities. The three unrelated
workspace files remain excluded; the broad project goal remains active.

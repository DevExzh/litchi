# 0804 — direct equality in the bounded linear scan

This diagnostic compares the exact 0803 after-control with the same helper
using byte-slice equality in its bounded linear name scan. Neither leg is
current production. Workflow advancement and production adoption remain false
regardless of the measured result; the rejected 0802 candidate stays rejected.

The linear scan previously called `Name::cmp(...).is_eq()`, computing byte
ordering to answer an equality question. The after leg tests the borrowed
byte slices for equality and preserves test-only comparison counting.
`Name::Ord` and its ordered-map use stay unchanged, as do empty-check placement,
the first/second/third prefix, the 32-name bound, ordered handoff, key/error
preflight, offsets, cloning and fused exhaustion. Shared tests are unchanged.

The hypothesis is that direct equality removes unnecessary ordering work in
the bounded scan. This is one source intervention, including associated
compiler and code-layout effects. It does not establish a unique instruction
as the cause of any native latency change.

## Protocol and verification

The [packet](results/change-0804/README.md) retains the same 39 literal cases
and construct/consume owners. Six alternating paired native blocks use 30
samples, three warmups and 4,096 iterations per child on CPU 12. Ratios use
nearest-rank process p50s and a 10,000-resample median bootstrap with fresh
seed 804080. Historical native timing is not pooled. A diagnostic regression
requires a median ratio above 1.05 and a 95% interval wholly above 1; all
individual rows and variability flags remain visible.

Construction includes common dispatch, an opaque iterator reference through
`black_box(&iterator)`, drop and checksum work. Consumption includes
construction and checksum work. These are hot repeated direct-helper inputs,
not workflow timing. Two separate Callgrind repeats retain guest instruction
and branch diagnostics, not native time, hardware branch rates, allocator API
counts or phase fractions. Object size is measured, not interpreted as heap use.

Both five-copy mirror legs pass 100 tests and warnings-denied Clippy. The
release probe passes formatting, build, check, Clippy, the independent fixture
audit and semantic self-check across all 39 cases and clone advances
0, 1, 2, 3, 4, 5, 32 and 33. These are isolated helper/probe gates, not full
production-workspace verification. Both iterator layouts measure 128 bytes.

A root setup script initially failed when candidate metadata was missing.
Incorrect shell sequencing then continued into quality, which completed
successfully on unchanged source archives. The quality witness freezes those
source files; metadata was completed and separately bound before the release
build. The execution note retains this setup failure. No quality test, build or native child was retried, and no failed-build
binary exists.
The archived helper comments still refer to the inherited 0802 design, including
its stale immediate-map sentence; the actual backend retains the bounded scan.

The initial profile lane stopped after its first child: the process exited
zero but produced only an empty termination dump, so the driver correctly
failed its positive-dump assertion. The optimized binary exposes only
`after_construct`; the original `before_construct` selector could not collect
that operation. This failed attempt, its zero dump and receipts are retained
under `profiles-failed-0`.

A separate, hash-bound profile owner resolution maps both construction legs to
`after_construct`, while retaining the two distinct consumption owners. Its
symbol witness and separate driver are archived. The same binary, native
samples, original plan and build inputs remain unchanged. A fresh profile lane
must prove one executed owner call and counter conservation in every child;
the original zero dump is not counted as a successful profile. This is an
explicit scope amendment after a failed profile, not a native timing rerun.

## Results

There are no diagnostic regression flags. Consumption improves most clearly
at 16–33 distinct names, with a smaller benefit after the ordered-map handoff.
The short rows include small increases; the frozen 5% rule does not hide them.
The full 78-row table remains in [summary.md](results/change-0804/summary.md).

| Consume case | After/before p50 | Change | 95% interval |
| --- | ---: | ---: | ---: |
| distinct-0 | 1.012988 | +1.299% | 0.997856–1.067949 |
| distinct-1 | 1.002094 | +0.209% | 1.001698–1.002979 |
| distinct-2 | 1.012556 | +1.256% | 1.009678–1.014030 |
| distinct-4 | 1.010841 | +1.084% | 1.008342–1.017765 |
| distinct-8 | 0.981138 | -1.886% | 0.958503–0.996608 |
| distinct-16 | 0.844107 | -15.589% | 0.812481–0.864115 |
| distinct-32 | 0.751753 | -24.825% | 0.725271–0.761829 |
| distinct-33 | 0.757407 | -24.259% | 0.741482–0.765053 |
| distinct-64 | 0.908234 | -9.177% | 0.877709–0.920927 |
| duplicate-valid-after-33 | 0.821211 | -17.879% | 0.810609–0.831526 |

Consumption medians range from 0.751753 to 1.024252. The maximum is equals-value
syntax after four attributes. Construction medians range from 0.944528 to
1.145025. Four construction medians exceed 1.05, but their intervals include 1:

| Construct case | Ratio | 95% interval |
| --- | ---: | ---: |
| distinct-32 | 1.128906 | 0.995643–1.155884 |
| duplicate-valid-after-1 | 1.145025 | 0.902776–1.157673 |
| syntax-unique-tail-after-4 | 1.054806 | 0.867220–1.131440 |
| syntax-equals-value-after-2 | 1.070591 | 0.928181–1.156830 |

There are 75 spread flags among 156 case/mode/leg groups: 63 construction and
12 consumption. All paired ratios and spread flags are retained. Both
construction legs execute the same surviving function symbol in this binary;
construction timing differences do not establish different constructor work.

Both profile repeats agree on these guest instruction totals:

| Case and mode | Before Ir | After Ir |
| --- | ---: | ---: |
| distinct-0/construct | 70 | 70 |
| distinct-0/consume | 122 | 122 |
| distinct-1/consume | 407 | 407 |
| distinct-2/consume | 870 | 870 |
| distinct-4/consume | 2,390 | 2,372 |
| distinct-8/consume | 5,052 | 4,914 |
| distinct-16/consume | 12,578 | 10,010 |
| distinct-32/consume | 35,978 | 26,130 |
| distinct-33/consume | 38,142 | 27,487 |
| distinct-64/consume | 100,838 | 90,145 |
| duplicate-valid-after-33/consume | 53,370 | 43,078 |

The 32-name operation removes 9,848 guest instructions in this comparison.
This supports direct equality as a useful source-level simplification to test
in a fresh production candidate. It does not assign native latency to those
instructions, and names with different lengths contribute to this fixture’s
comparison behavior; the result is not a universal per-name cost model.

All 312 resolved owners qualify with exactly one call each. Independent scalar
conservation covers all five counters in 312 positive dumps and 312 zero
termination dumps. Successful captures contain 936 native reports and 28,080 samples plus 312
profile reports with one sample each: 1,248 reports and 28,392 samples. The
failed initial profile adds one separately retained report/sample and one
zero dump; it is excluded from successful scope counts and timing analysis.

## Disposition and limits

The decision remains diagnostic-only. A source-level improvement against the
0803 control cannot be multiplied by historical ratios to establish a gain
over production. Any combined candidate needs a fresh production comparison
and the protected error-boundary gates before workflow, resource and
cross-format qualification. The constructor/profile scope issue also shows
why requested symbol names alone are insufficient evidence of collected work.

The source intervention adds no allocation, concurrency, ambient provider or
unsafe code. Borrowed names, bounded linear checking, ordered fallback and
first-error behavior stay fixed. All 9,196 production source hashes and 35
architecture inputs remain unchanged. No memory, cold/range, producer,
concurrency or CRUD coverage is promoted. iWork is excluded and the broader
GOAL remains incomplete.

Owned-target cleanup removed 178,385,224 logical bytes after exact binary
identity verification. The failed profile used that same final binary, so no
separate failed-build executable existed. Post-cleanup replay and final sealing
verify all retained evidence and the unchanged production source census.

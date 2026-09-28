# 0826 — allocation-vector admission repair

Added a reusable Python checker for the allocation-vector schema that stopped
0825 qualification. The unchanged Rust harness emits each counter as a metric
object containing a per-sample `values` array; the frozen packet checker
incorrectly expected an integer. The new checker accepts that exact retained
report and reproduces the old checker's rejection as a regression witness.

This is validation-tooling work. Production Rust, benchmark runtime, locks,
all 35 normative inputs, and the three unrelated workspace files retain their
recorded hashes. No Cargo command, workload, new timing sample, profiling run,
or optimization decision is part of this batch. The complete ordinary-save
effect of the shipped 0824 optimization remains to be measured in a fresh trial.

## What the checker validates

[`tools/perf_allocation_schema.py`](../../tools/perf_allocation_schema.py)
checks a single-case schema-v1 capture against caller-supplied case, sample,
warmup, and native/observer expectations. It validates binary/instrumentation
labels, the v3 allocator revision, elapsed sample ordering, exact sample-index
alignment, and all eleven allocation metric objects.

Measured vectors must have the expected length, matching status/scope, and
unsigned 64-bit integers. Booleans, floats, negative values, scalar metrics,
missing/unknown fields, and partial vectors are rejected. Each sample must
preserve signed live-byte conservation; region peaks must cover both live-byte
endpoints and stay within the lifetime peak. Lifetime peaks remain monotonic,
and successful reallocation calls cannot exceed successful allocation calls.
Native reports must explicitly report unavailable counters without numeric
payloads. The checker uses explicit exceptions, so `python -O` cannot disable
its checks.

Signed net-live decreases are valid. Region peaks and process lifetime peaks
are distinct. Nonzero failed-allocation counters are also valid schema data;
an explicit trial policy can reject them separately. Independent review caught
that policy/schema distinction in the first draft. Its passing positive
preflight and exact helper source remain archived; the final helper and tests
use the corrected behavior.

The helper does **not** validate the `source.ordinary_save` object, corpus or
output hashes, OPC/XML preservation, executable hashes, process counters, or
statistical summaries. Matching a case label alone cannot establish those
properties. Existing independent artifact admission and source custody remain
required, and allocator-instrumented elapsed values remain ineligible for
native latency claims.

## Evidence

| Check | Result |
| --- | --- |
| Regression suite | 30 tests pass in normal Python |
| Assertion-independent behavior | Same 30 tests pass under `python -O` |
| Historical schema preflight | 325 pinned reports / 6,733 sample envelopes pass |
| Scenario coverage | DOCX/XLSX/PPTX × lifecycle/edit/atomic/counting publication |
| Original failure | Frozen 0825 scalar checker rejects the retained report; new helper accepts it |
| Runtime/source custody | Production, benchmark runtime, locks, 35 normative inputs, unrelated files unchanged |
| New measurements | Zero |

The fixture matrix comprises 108 reports from 0819, 216 from 0821, and the
single failed-qualification report from 0825. Each report is bound to its
historical seal and a pinned path/size/SHA-256 descriptor. The preflight checks
sample envelopes and alignment without calculating ratios, quantiles, or any
new performance result. Historical reports serve only as schema fixtures.

The adversarial tests cover every allocation field, status/scope mismatch,
vector cardinality, bool/float/negative/overflow values, unsigned boundaries,
conservation and peak violations, relabeled mode/revision expectations,
misaligned/duplicate sample indices, ties, unavailable native metrics, and
successful/rejected CLI calls. Temporary CLI fixture directories close before
tests return. There were no failed test or reader invocations in this batch.

## Replay and next measurement

```sh
python3 -B -m unittest tools.test_perf_allocation_schema -v
python3 -B -O -m unittest tools.test_perf_allocation_schema -v
python3 -B docs/performance/results/change-0826/preflight.py --check
python3 -B docs/performance/results/change-0826/validate.py
python3 -B docs/performance/results/change-0826/seal.py --check-head
```

The [packet](results/change-0826/README.md) retains fixture hashes, both test
logs, tested source snapshots, the source review, unchanged-source census,
preflight results, and cleanup witness. No build or filesystem scratch roots
were created. The final seal binds the exact two Python tools and all report
and evidence files to the commit.

The next matched trial must freeze this helper's hash as an execution input
and call `validate_report` with expectations from its frozen plan. Before
freezing the comparative reader, also resolve accepted admission attempts from
bound descriptors instead of hardcoded attempt-directory names, and compare
ZIP-preservation replay through canonical JSON so timestamp tuples match
serialized arrays. Then perform fresh builds, independent admission, both
qualification legs, and the complete matched capture. Nothing in this repair
retroactively admits 0825 or establishes a complete-save performance benefit.

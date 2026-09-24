# ODS paired-statistics performance profile plan

This package measures the eight paired OpenFormula reducers `CORREL`,
`COVAR`, `PEARSON`, `RSQ`, `SLOPE`, `INTERCEPT`, `STEYX`, and `FORECAST`.
The matched baseline is `aa48eee68cb6ee0904523394d4ff8015dce6e595`. The
profile uses the retained order-statistics gate lock
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`, copied
into the harness for reproducible dependency resolution. The reviewed
contract is frozen at
`9ce79b7b5014c6e29c1d30c61a77187d6dc4bce9be7810e1f5006e8ef147aadd` and
selects the included-constant `INTERCEPT` profile described there.

## Matched controls

Both sides run the 24 existing dispatch and evaluator controls, plus 15
controls for the five descriptive reducers whose centered-moment machinery
is shared by this batch. The descriptive controls deliberately have three
forms per reducer: a normal 64x4 reference, a sensitive dyadic inline array,
and an extreme signed dyadic inline array. This gives before/after evidence
for the exact consumers affected by the shared extraction without inheriting
the previous descriptive family's full candidate cross-product.

```text
scalar-control-arithmetic
scalar-control-sin
scalar-control-imsum
database-control-dsum
scalar-control-average
scalar-control-counta
scalar-control-var
scalar-control-stdev
database-control-dvar
database-control-dstdev
array-control-4x4-arithmetic
array-control-4x4-sin
array-control-16x16-arithmetic
array-control-16x16-sin
reference-array-16x4-arithmetic
scalar-aggregate-sum
literal-aggregate-4x1-sum
reference-aggregate-64x4-sum
reference-conditional-256x4-sumifs
reference-control-average
reference-control-counta
representative-median
representative-rank
representative-percentrank
matched-descriptive-{reference,sensitive-inline,extreme-inline}-{avedev,devsq,kurt,skew,skewp}
```

The concrete runner order is 39 controls total: the first 24 entries above,
then each descriptive reducer's reference, sensitive inline, and extreme
inline case in function order.

The reference `AVEDEV` control is expected to make two physical passes over
its 64x4 range; `DEVSQ`, `KURT`, `SKEW`, and `SKEWP` use the shared centered
moment state in one pass. The verifier records those exact read bounds so a
reducer cannot hide an output-independent second scan.

## Candidate matrix

The candidate adds 66 bounded groups: eight lanes for each paired function,
plus two `FORECAST` query-array lanes. The candidate-wide preflight runs all
105 named cases before baseline timing begins.

| lane | functions | fixture and purpose | resolver-read intent |
| --- | --- | --- | --- |
| `small` | all eight | 16x4 `known_y`/`known_x` references with an exact affine relation | 128 reads, two 64-cell sequences |
| `large` | all eight | 256x4 references for input-size scaling | 2048 reads |
| `extreme` | all eight | inline finite dyadic values spanning large and small magnitudes | 0 reads |
| `offset` | all eight | equal-length references beginning at row 3 | 128 reads |
| `pairwise-skip` | all eight | aligned M:P/Q:T references with Text/Empty positions omitted as pairs | 128 physical reads |
| `shape-reject` | all eight | ReferenceList refusal, or RSQ's unequal checked-cell count | 0 reads after descriptor refusal |
| `cancellation` | all eight | 64x4 references with sticky cancellation after the first successful read | exactly 1 total read per child |
| `resource` | all eight | the same references with `max_reference_cells = 0` | 0 reads before typed resource failure |
| `forecast-query-array` | `FORECAST` | two query values over one shared fit | 128 reads, 2x1 output |
| `forecast-query-array-cache` | `FORECAST` | projected query array with complete fit reuse | 128 reads, 2x1 output |

The direct rows use complete ForceArray descriptors. Pair construction is
position-wise and admits only finite Numbers on both sides; Empty, Text, and
Logical members skip the aligned pair. Formula Errors remain retained while
the scan continues. The harness reports actual fixture values and checksums
for every preflight, so a fixture correction cannot silently become timing
data.

The cancellation fixture intentionally shares one `ExecutionContext` and
resolver across the four internal repeats. The first successful read cancels;
the next three repeats refuse before `read_cell`. Verification therefore
requires `reference_reads = 1`, `repeat = 4`, and integer-floor normalized
reads per repeat of zero. This is a mixed-batch latency measurement, not four
fresh mid-read cancellation events.

## Correctness, resource, and cache gates

The reducer equations, included-constant `INTERCEPT`, RSQ `#N/A` cardinality
case, shape/list refusal, pairwise omission, and FORECAST matrix query shape
come from the reviewed contract and `numeric_oracle.py`/`numeric-goldens.json`.
Untimed preflight checks numeric output, typed failures, output shape, and
resolver reads for each case. Typed cancellation/resource failures remain
evaluator failures. The borrowing resolver exposes no materialized cell
vector; allocator calls/bytes, live and peak bytes, work, memory budget, and
reads are recorded alongside external RSS.

Complete paired descriptors are the cache unit for projected `FORECAST` query
arrays. The query position remains observable, and a scalar query or nested
array-producing argument is not treated as a free-standing invariant. The
profile records the two FORECAST rows separately so an output-multiplied fit
scan is visible.

## Source closure and capture protocol

The selected closure includes the shared evaluation dispatch, dyadic and
descriptive moment helpers, paired kernel/value owners, matrix/reference
helpers, existing aggregate/statistical/order/conditional paths, paired and
descriptive tests, contract/oracle/goldens/native inputs, and the harness
lock. Recursive paired and descriptive globs retain later module splits.
`stage.py` requires the generator, numeric goldens, and native provenance in
the closure; missing authored inputs fail closed.

Once root authorizes capture, `run_profile.py` uses three untimed warmups and
15 fresh child processes in both `evaluate` and `parse-evaluate`. It preserves
raw stdout, `/usr/bin/time -v` RSS logs, source/profile manifests, and cleanup
receipts, while removing temporary targets and the baseline worktree. No
source or profile input may change after timing starts.

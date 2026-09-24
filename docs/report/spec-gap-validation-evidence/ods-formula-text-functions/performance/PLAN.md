# ODS text-functions performance profile plan

This package profiles the 26 OpenFormula 1.4 section 6.20 functions in the
frozen text batch: `ASC`, `CHAR`, `CLEAN`, `CODE`, `CONCATENATE`, `DOLLAR`,
`EXACT`, `FIND`, `FIXED`, `JIS`, `LEFT`, `LEN`, `LOWER`, `MID`, `PROPER`,
`REPLACE`, `REPT`, `RIGHT`, `SEARCH`, `SUBSTITUTE`, `T`, `TEXT`, `TRIM`,
`UNICHAR`, `UNICODE`, and `UPPER`. The matched baseline is
`8f09231e36982248eface4d599432143a67f6e49`.

The harness uses the retained order-statistics gate lock
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`.
The text contract is deliberately a capture prerequisite. The reviewed input
is pinned by `CONTRACT_SHA256` to
`b03ec074efe30c4261509a70b6188010d74a9d48d8bc92a57ad39a2db4a3b611`, and the
runner fails closed if that exact file changes. This keeps timing tied to the
reviewed domain, list admission, Unicode profile, and optional-argument rules.

## Matched controls

The baseline and candidate use the 24 existing ordinary evaluator controls
from the order/paired profiles, plus four explicit concatenation controls.
The latter cover a borrowed literal result, an owned left operand, an owned
right operand, and a repeated growth chain. They keep `&` allocation and
borrow behavior visible when the text dispatch is added.

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
concat-borrowed-literals
concat-owned-left
concat-owned-right
concat-growth-chain
```

These 28 controls are measured in both `evaluate` and `parse-evaluate`. The
ordinary controls preserve the prior profiles' exact fixture reads, allocator
counters, work, retained budget, and RSS fields, so a text dispatch change is
visible even when the text rows themselves are candidate-only.

## Candidate matrix

The candidate contains 133 named cases, giving a bounded 161-case capture
including controls. Every function appears once in each of the four broad
function lanes (`tiny`, `large-unicode`, `reference-64`, and `refusal`).

| group | count | coverage | read expectation before contract freeze |
| --- | ---: | --- | ---: |
| `tiny-text-*` | 26 | scalar literals, optional arguments, and direct scalar results | 0 |
| `large-unicode-text-*` | 26 | bounded mixed Unicode input, with numeric-only functions using their numeric domain lane | 0 |
| `reference-64-text-*` | 26 | one borrowed `Main!A1:D16` descriptor and scalar companion arguments | 64 (`EXACT`: 128 because it compares the descriptor twice) |
| `matrix-broadcast-text-*` | 13 | `CONCATENATE`, `EXACT`, `FIND`, `LEFT`, `LEN`, `LOWER`, `MID`, `PROPER`, `REPLACE`, `RIGHT`, `SEARCH`, `SUBSTITUTE`, and `TRIM`; output shape is 16x4 | unary lanes 64; two-reference row/column broadcast lanes 128 (each output selects both descriptors) |
| `refusal-text-*` | 26 | type/domain/list or arity refusal; the descriptor/type gate must precede resolver access | 0 |
| `cancellation-text-{concatenate,len,substitute,search}` | 4 | shared sticky execution context: first child read cancels and later repeats refuse | exactly 1 total read per child; repeat 4 |
| `resource-text-{concatenate,len,substitute,search}` | 4 | `max_reference_cells = 0` typed resource failure | 0 |
| `search-worstcase-{find,search,substitute,exact}` | 4 | long near-match/non-match strings, bounded search work | 0 |
| `rept-growth` | 1 | output growth below the text limit, with input/output byte metrics | 0 |
| `text-format-fraction-six` | 1 | scalar `TEXT(0.3333333333333333;"######/######")`, exercising the six-position bounded improper-fraction kernel and exact `"1/3"` output; default repeat 1 and scoped scalar `max_steps=2,000,000` | 0 |
| `asc-jis-expansion-{asc,jis}` | 2 | full-width/half-width conversion and expansion/contraction | 0 |

The matrix is intentionally smaller than a function-by-lane cross product.
`reference-64` isolates descriptor streaming and borrowed resolver text;
`matrix-broadcast` exercises output mapping and scalar/broadcast selection;
the refusal/cancellation/resource rows isolate typed gates. Numeric-only
functions in the Unicode lane use representative numeric inputs, while text
consumers receive a bounded UTF-8 fixture rather than an unbounded generated
literal.

The reference resolver returns `CellRead::Text` from static fixture strings.
It does not construct a cell vector. The evaluator must retain reference
geometry and select one borrowed cell per output position. The harness records
input bytes, output bytes, work, exact successful reads, allocation calls and
bytes, live and peak bytes, retained execution budget, and external RSS.
`output_bytes_p50` is the sampled output total across the fixed repeat count;
`bytes_per_repeat_p50` adds one input payload and divides that output total by
the repeat count so byte throughput is comparable across scalar and mapped
lanes.

Reference read bounds are checked during untimed preflight and again in the
capture verifier. A one-reference 16x4 lane is exactly 64 reads. `EXACT`
intentionally passes the same 16x4 descriptor twice in its reference lane and
reads 128 cells. A broadcast
lane with a row or column descriptor has an explicitly recorded bound derived
from its output shape and descriptor reuse; the two-reference 16x4/16x1 lanes
(`CONCATENATE`, `EXACT`, `LEFT`, `MID`, `REPLACE`, and `RIGHT`) therefore read
128 cells, not 80, because each output selects both descriptors.
A cancellation lane has one total successful read across four internal
repeats because the same sticky context is intentional. Resource and refusal
lanes must fail before any resolver read.

## Correctness and cache checks

Preflight parses every case and validates the scalar or matrix shape, text or
typed-error result, output checksum/byte count, exact resolver reads, and
resource/cancellation failure type. The reviewed contract and its numeric/text
oracle own the expected Unicode, locale, search, optional-argument, list, and
shape semantics. The pinned contract hash and input manifest remain a
fail-closed capture gate.

Projected matrix branches must carry complete text reference descriptors into
`map_function`; they may not materialize an input cell vector. The scalar
payload demand cache remains conservative for text values, so a projected
lazy branch selects the reference cell at its output coordinate rather than
assuming text is an invariant cache payload. Position-sensitive
search/substring parameters therefore remain associated with that coordinate.
Repeated `&` and `CONCATENATE` rows retain the output until checksum/drop so
owned text storage appears in the allocator and live-byte measurements.

## Source closure and capture protocol

The profile closure includes the evaluator dispatch, scalar/text owners,
`value.rs` matrix/reference/owned helpers, unchanged aggregate/statistical/
order/paired/descriptive controls, all landed text tests, the reviewed text
contract/oracle/generator/goldens, native cached results and provenance, the
feature matrix, and the harness lock. Recursive `evaluation/text*.rs`,
`evaluation/text/**/*.rs`, `evaluation/value/text*.rs`,
`evaluation/value/text/**/*.rs`, `evaluation/unicode*.rs`, text tests,
contract, generator, oracle, golden, native provenance, and the complete
`unicode-data/` generator/provenance/license directory are included so a
later module split or Unicode table refresh cannot escape the frozen source
closure.

After root authorizes capture, `run_profile.py` performs all candidate and
matched-control preflights before either side is timed, then runs three
warmups and 15 fresh child processes in both phases. It preserves raw child
JSON, `/usr/bin/time -v` logs, source/profile manifests, and cleanup receipts.
No source, contract, generator, lock, or profile input may change after the
freeze hash is recorded. Dynamic results stay outside the frozen input set.

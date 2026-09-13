# Radix harness review

This is a read-only audit of the frozen four-file harness and the corrected
baseline receipt. The production revision under test is
`e1976c59d5c4785e0c73a5d27e9349ba082cca2f`. The candidate radix implementation
had not been captured when this review was written.

## Frozen harness

The current files and the baseline custody receipt agree byte-for-byte:

| file | SHA-256 |
| --- | --- |
| [`Cargo.toml`](radix-harness/Cargo.toml) | `f337daad9e20c4b63c10fc84fc998f2ee16762170d513c59e1e7c4b471c9cdff` |
| [`Cargo.lock`](radix-harness/Cargo.lock) | `0bd5dc7cfe53d0a6625406040d86b10f3f257f9cc061747c4a77077e5f1c8df0` |
| [`run.py`](radix-harness/run.py) | `a96ad565525121c9f00135326c21798ffde574c39e4caf8dd34453ae32876c6f` |
| [`src/main.rs`](radix-harness/src/main.rs) | `11c3d7df4156199ae15e4e9931551d2dcc4644601ce3dbfea0a273a13e576c61` |

The package and emitted workload are both `ods-formula-radix-evaluation-profile`
and `radix-evaluation`. The comparable group has 47 cases: the prior 39
scalar/logical cases plus eight representative bitwise cases. The radix group
has 53 candidate-only cases: one fixed success for each of the fourteen
functions, six spelling/truncation controls, seven scalar formula-error
controls, four bounded-refusal controls, sixteen input/padding scale lanes at
64/256/1024/4096, and six large finite-number lanes. Every case has parse,
evaluate, and parse-evaluate phases. The runner uses three warmups and fifteen
timed iterations.

## Corrected baseline receipt

The valid baseline receipt is [`baseline/comparable/raw.csv`](baseline/comparable/raw.csv),
with command and sequence metadata in [`capture.json`](baseline/comparable/capture.json)
and [`sequence.json`](baseline/comparable/sequence.json). It was captured after
the corrected harness was built; the captured interval is
`2026-09-13T13:45:57.390Z` through `2026-09-13T13:45:58.763Z`.

| phase | rows | expected scalar success | expected refusal | process status 0 |
| --- | ---: | ---: | ---: | ---: |
| parse | 47 | 47 | 0 | 47 |
| evaluate | 47 | 40 | 7 | 47 |
| parse-evaluate | 47 | 40 | 7 | 47 |
| total | 141 | 127 | 14 | 141 |

All rows have status `0`. The 14 refusal rows are the seven existing control
failures in each evaluation phase: unsupported reference, array, and name;
Work, Memory, and Objects limits; and cancellation. Each expected label occurs
twice, and all other 127 rows report `failure=none`. Maximum resident set size
in these process receipts ranges from 2384 to 5096 KiB. The build completed
with status `0` in [`baseline/build.log`](baseline/build.log); its binary is
recorded as SHA-256
`8437de6ad6b53d4883960f2b20fdb7e13b6af80822fd740934f42525ce015733`, size
1,266,984 bytes, in [`baseline/binary-provenance.json`](baseline/binary-provenance.json).
The source and harness provenance are in
[`baseline/source-sha256.json`](baseline/source-sha256.json),
[`baseline/harness-sha256.txt`](baseline/harness-sha256.txt), and
[`baseline/harness-custody.json`](baseline/harness-custody.json).

## Normative audit

The current expected vectors agree with the approved profile and the checked
integration vectors:

- Each of `BASE`, the twelve `xxx2yyy` functions, and `DECIMAL` has a fixed
  case. The small vectors use the specified digit alphabet and signed widths.
- `BASE(15.9;16)` expects `F`, reflecting the selected generic Integer
  truncation rule. The direct decimal-to-radix fraction case expects a scalar
  `Value` error. Radix 1 and 37 expect scalar `Number` errors, and the large
  direct-conversion cases exceed their signed target widths and also expect
  scalar `Number` errors.
- The `DECIMAL` controls cover leading spaces, a leading tab, `0x`, trailing
  `H`, and binary `B`. These are scalar results. Formula errors and bounded
  evaluator refusals are represented separately: the former have
  `expected_success=true` and are checked as `ScalarValue::Error`; the latter
  have `expected_success=false` and are checked by typed failure labels.
- `BASE` minimum-width lanes use the third argument as output padding, which is
  distinct from the direct-converter `Digits` argument. The 4096-width BASE
  output is intentionally a single bounded output, while the DEC2HEX scale
  uses width 7 and 4096 concatenated results (28,672 bytes), below the 32,767
  text envelope used by the evaluator.
- The finite-number lanes include `2^53-1` and an exact `f64::MAX` BASE/DECIMAL
  pair. The 256-digit hexadecimal oracle for `f64::MAX` is independently
  derived from `(2^53-1) * 2^971`, rather than copied from evaluator output.
- `main.rs` compares exact Number, Logical, Text, and scalar-error values in
  `validate_value`; a stable allocation/checksum line alone would not make a
  candidate pass. The candidate-only group is intentionally absent from the
  baseline performance receipt, so no new-function speedup claim can be made
  from the baseline.

## Findings and coverage limits

No correctness blocker remains in the current frozen harness. The baseline
receipt is internally consistent and supersedes the earlier provisional
capture. That provisional version was not suitable for comparison: it treated
large `DEC2BIN`/`DEC2HEX` values as successful 53-bit text, classified direct
fractions and radix bounds with the wrong scalar errors, and used an oversized
padding construction. Those cases were corrected to signed-width `Number`,
direct-fraction `Value`, radix-bound `Number`, width-7 DEC2HEX concatenation,
and direct BASE padding before this receipt was rebuilt. The cancellation
predicate was also corrected at all three execution-context construction
sites to include `radix-refusal-cancelled`. The finite `f64::MAX` BASE and
DECIMAL vectors were added at the same correction point.

The following are deliberate coverage limits rather than receipt failures:

- The radix group is candidate-only. Baseline evaluation of those formulas is
  supported by the runner as an `unsupported-function` control, but it was not
  run in this baseline capture.
- Fixed `xxx2yyy` cases primarily use text operands. The integration suite
  separately exercises integral numeric operands and fractional rejection;
  the performance corpus does not repeat every numeric/text pairing.
- The performance corpus uses hexadecimal BASE at the finite boundary. The
  integration suite covers the independent base-2, base-10, and base-36
  `f64::MAX` vectors.
- The text corpus exercises the accepted prefix/suffix forms but leaves the
  broader malformed-text matrix to the integration tests.

These limits should remain explicit when candidate results are reported: the
candidate comparison can establish exact behavior for the listed vectors,
bounded refusal labels, and comparable legacy lanes, but it cannot establish
performance or semantics for unlisted radix spellings or for a baseline
implementation of the new functions.

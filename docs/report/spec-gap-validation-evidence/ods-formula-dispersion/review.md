# ODS dispersion semantic and numerical review

This review covers `VAR`, `VARA`, `VARP`, `VARPA`, `STDEV`, `STDEVA`,
`STDEVP`, and `STDEVPA` in the explicit, read-only formula evaluators. It
compares the frozen implementation with the repository-local ODF 1.4 source
and the profile in [`contract.md`](contract.md). The review made no production
or test edits.

The semantic and numerical disposition is **pass** for the frozen evidence.
The result is a finite-binary64 compensated-kernel assessment; it is not a
claim of exact rounding or of a universal error bound. Performance is
measured separately; the final receipt is summarized below.

## Frozen identity

The implementation freeze is [`gates/freeze.json`](gates/freeze.json), based
on commit `55e147bfa0676ce6ecdc609efc682b98568b8a5f`. The contract hash is
`5c4b881fcddee6bd5a1d4a143d74fbfb794c4d58726e3922ae492c963a516852`.
The normative source recorded by the contract is the repository-local ODF
archive `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`
and its Part 4 HTML member
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.

The selected freeze hashes are:

| Frozen file | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `fb9ebd5e684fa2cde84d5fce5960e77184571bffc47ec3ef8bcbe73a8dc40ed4` |
| `crates/litchi-ods/src/codec/formula/evaluation/statistical.rs` | `1045c81b8d7780b6d4edd5073f66240701b2f7e6c4a2ad3b894e326f9eb0328a` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/statistical.rs` | `6d6b1d1432d4a9367c4c29e1a672aed0e92368fb149c8615df8bd03292fcdbf3` |
| `crates/litchi-ods/tests/ods_formula_dispersion_evaluation.rs` | `ab89fb34f4053444320b00d2a3479bee50cf7c7171fe9c211b442edc1ad34693` |
| `crates/litchi-ods/tests/ods_formula_dispersion_limits.rs` | `1d20d12e1c99d2e926f9993880bc96035411bbd3d48d4544f93a1739208d9244` |
| `crates/litchi-ods/tests/ods_formula_dispersion_native.rs` | `223f05af0b81613059fc49b0f5b20e3e6e2ed82f566eb9e78b263ed0b9ed2fc0` |
| `crates/litchi-ods/tests/ods_formula_dispersion_oracle.rs` | `f42a9a96b5bb26b82b623b4c26b8e6fa22c4459c59a280ddadd5701cdc687825` |
| `numeric-goldens.json` | `0a5a52c6f38e045f0221038d4563a845b4730845579068e7a042532469df9e8b` |
| `native/cached-results.json` | `88c785f9b36e98a8b7e88f0bbfc76ad99b845d1592070ea1c32feea7e46f02f5` |
| `native/provenance.json` | `d1f1e0317ae44a7b03178d0a0b4fe19bac6d9561f439123eee53522d2cf8ef35` |
| isolated `Cargo.lock` | `58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3` |

## Semantic findings

The function table in the frozen scalar and resolver-aware paths maps the
eight names to the intended operations: sample variance, population variance,
sample standard deviation, and population standard deviation, with the A
variants sharing the corresponding operation after their `Any` conversions.
The selected profile follows the ODF pseudotypes:

* `VAR`, `VARP`, and `STDEVP` use `NumberSequence`; `STDEV` uses
  `NumberSequenceList`; and all four A variants use `Any`.
* A scalar Number is admitted directly. Direct Logical values become zero or
  one. Direct finite decimal Text is parsed for the NumberSequence family;
  A-variant Text becomes zero, including empty Text. Scalar Empty is zero for
  the ordinary sequence extension and is omitted for A variants. Missing and
  Complex values produce the selected formula error.
* Referenced NumberSequence cells admit Numbers and retain formula Errors;
  referenced Text, Logical, and Empty cells are omitted. Referenced A-family
  cells admit Text as zero and Logical as zero or one while omitting Empty.
  Numeric-looking reference Text is therefore not parsed cell by cell.
* The frozen shape gate rejects an explicit ReferenceList for `VAR`, `VARP`,
  and `STDEVP`, and accepts it for `STDEV` and all four A variants. The
  resolver-aware path records the rejection before its reducer scan, while
  admitted lists retain occurrence order.
* Sample reducers require two admitted members. Population reducers require
  one and return canonical positive zero for a singleton. Empty and
  insufficient selections, including zero supplied arguments, use the
  contract's `ScalarError::Value` profile. Formula Errors retain source order
  and supersede a count decision; typed resolver, resource, cancellation, and
  source failures remain evaluator failures.

The shared reducer dispatch is used in both scalar and value evaluation. The
value path scans references in retained area, sheet, row, and column order,
keeps the complete descriptor for admitted list shapes, and applies the same
conversion rules to inline arrays. The integration hooks classify all eight
names as scalar-valued statistical reducers and preserve complete descriptors
through projected branches. The focused projection case confirms that an
invariant nested reducer can reuse its complete reference while a nested
`MUNIT` criterion remains position-sensitive.

The state is bounded: the reducer retains only counters and the fixed-size
`VarianceAccumulator`/`NumericAggregate` state. The compensated first-value
offset kernel rescales its normalized moments as larger magnitudes arrive and
publishes variance or standard deviation through checked denominators. It
handles adjacent representable values, opposite extreme values, subnormals,
and a finite standard deviation whose variance would overflow without
materializing the admitted sequence.

## Independent numerical assessment

[`numeric_oracle.py`](numeric_oracle.py) constructs exact `Fraction` results
from the represented binary64 operands and takes standard-deviation square
roots with 600-digit `Decimal` precision before the final binary64 conversion.
It does not invoke the Rust evaluator or a native spreadsheet application.
The retained [`numeric-goldens.json`](numeric-goldens.json) contains 512
observations over 56 fixtures, 64 observations for each of the eight
functions, including both scalar and ReferenceList descriptors where the
profile admits them.

The comparison policy is deliberately split by fixture:

* 207 ordinary numeric rows pass at no more than the declared 8 ULP bound.
* 150 marked cancellation rows pass at no more than the declared `2e-12`
  relative bound. This set includes adjacent values, opposite extremes,
  subnormal and underflow deltas, reordered values, and forward/reverse
  majority-equal runs of 10,000 and 100,000 elements.
* 155 formula-error or insufficient-cardinality rows match their typed
  expected errors. Rows whose mathematical result is zero require exact
  positive zero; the Rust oracle checks the sign bit directly.

The 8 ULP allowance is adequate for the retained ordinary corpus: every
ordinary successful observation passes it in both scalar and matrix modes.
The cancellation allowance is the existing database numerical profile and is
kept narrow enough for the majority-equal probes to expose order-dependent
error. These observations establish acceptance for the frozen fixture set
only. The first-offset compensated algorithm remains an approximate
binary64 profile, so this review makes no universal bound claim for larger or
different cancellation populations. A future kernel change must regenerate
the independent oracle and recheck those bounds rather than widening them to
hide a failure.

## Validation receipts

The focused dispersion targets pass:

* `ods_formula_dispersion_evaluation`: 9/9;
* `ods_formula_dispersion_limits`: 11/11;
* `ods_formula_dispersion_oracle`: 8/8, evaluating all 512 observations in
  scalar and matrix modes;
* `ods_formula_dispersion_native`: 2/2, covering the 48 retained native
  observations and pinned-source exclusions.

`python3 numeric_oracle.py --check` reports
`{"observations": 512, "fixtures": 56, "verified": true}`. The isolated
package gate records 1,378 tests with zero failures and zero ignored tests;
strict all-target Clippy, rustdoc, crate and selected-file formatting, crate
boundary, and diff checks all pass with stable compiled source hashes.

The native receipt retains 48 cached observations, six per function, from
eight pinned LibreOffice FODS inputs. Its independent reproduction regenerated
identical bytes and cleaned its temporary source tree. Native values are
corroboration for ordinary host behavior; they do not override the local
pseudotype or `#VALUE!` cardinality choices.

The final performance capture retains 510 baseline and 3,090 candidate
samples. Root independently checked source custody, all raw observations,
reported medians, and resolver-read bounds. Matched allocation metrics are
unchanged; median time shifts are −3.24% to +2.35% and RSS shifts are −2.52%
to +4.22%. These are bounded single-host observations; no universal timing
or throughput conclusion follows from this semantic review.

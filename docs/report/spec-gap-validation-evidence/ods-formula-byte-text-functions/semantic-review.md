# ODF byte text function semantic and integration review

This review covers the seven OpenFormula 1.4 §6.7 functions `FINDB`,
`LEFTB`, `LENB`, `MIDB`, `REPLACEB`, `RIGHTB`, and `SEARCHB`. It checks the
frozen implementation against the repository contract, with attention to
conversion and domain errors, scalar and matrix contexts, lazy projection,
reference descriptors, and source/provider failures. The review made no
production or test edits.

The disposition is **PASS**. I found no unresolved semantic or integration
blocker in the frozen source.

## Frozen identity

The authority is [`gates/freeze.json`](gates/freeze.json), based on commit
`3844f235bac545ff0ae1580b97612883c1fd9f89`. The semantic implementation and
its integration points have these frozen SHA-256 values:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `971f7589f5abd32930e039e7cb07a1b6f20dcab776836ade37327648ebc85847` |
| `crates/litchi-ods/src/codec/formula/evaluation/value.rs` | `e9e266ad59658d378fb758a436fd7a4a12e66b9924c225e074a03a82c8ceb51f` |
| `crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs` | `cd882fa914bffa26ede20c51a4483c53fc36f4070e37cfbfef2bfec6e2db69f3` |
| `crates/litchi-ods/src/codec/formula/evaluation/text.rs` | `11d7e945eeb00b733941c34026ae670d997d58a76be16808c2d10b0a9221ac41` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/bytes.rs` | `3a625110a77697bba2e7414f9c898bedf61a70bed8517cc9d67db6036add3d61` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/core.rs` | `17d1a1d1f9bbd4a30eec6a48d1aa1d049e29040b489779c0ca867a6f27f962f0` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/search.rs` | `523b93ce563976bc617c626133e3dd90760d0580b63ef677a7eb73bda0f78d0f` |
| `crates/litchi-ods/src/codec/formula/evaluation/text/unicode.rs` | `336584ece56ee00e9e68a0c2165b5396ac1f007d5ad671f1fe8eba30aac877b6` |
| `crates/litchi-ods/tests/ods_formula_byte_text_evaluation.rs` | `57be711b33521d02197552bba33a09e105e4c927e8cc7e1fc5a6451e68ca8007` |
| `crates/litchi-ods/tests/ods_formula_byte_text_limits.rs` | `ad4723fe4e3ee6a31a82e59a63e036fcabbff01a25b52ec06778f09f534981d2` |
| `crates/litchi-ods/tests/ods_formula_byte_text_oracle.rs` | `be7897323c2dbb937a19acc3546be18da2f4739ec843326c74e7325c5261783b` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/contract.md` | `88dafc6f0d111672e724b4238289afc0a17d879b856059ac3886ca60bbc99db5` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/byte_oracle.py` | `0b677d1e3b67d2597fdfa292d656abd98debf308f3a7be4387a8e37341888859` |
| `docs/report/spec-gap-validation-evidence/ods-formula-byte-text-functions/byte-goldens.json` | `9f30f244b3f98d1f6ddbdc481165efbf153c80205ad60a85cab1c7561f47746c` |

The full selected-file manifest, native receipts, gate logs, and verification
status remain in `gates/freeze.json`, `gates/results.json`, and
`gates/verification.json`. The isolated gate lock is SHA-256
`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`; the
ambient root lock has the separately recorded SHA-256
`aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02`.

## Semantic review

The registry and dispatch cover exactly the seven contract names. The scalar
bridge validates arity before consuming arguments, propagates an input formula
error, rejects a missing required argument with generated `#VALUE!`, and
applies the documented optional defaults: `Start = 1` for `FINDB` and
`SEARCHB`, and `Length = 1` for `LEFTB` and `RIGHTB`. The required three-argument form of `MIDB` and four-argument form of
`REPLACEB` remain strict. An explicitly supplied
value is converted and validated even when it is zero, false, empty, or
otherwise falsy.

The selected profile uses UTF-8 octets of semantic Text. `LENB` counts those
octets; slicing snaps an interior start backward to its containing scalar and
returns complete UTF-8 scalars only. `LEFTB` and `RIGHTB` use the ordinary
integer-length rule, `MIDB` uses the ordinary integer start and length rules,
and `REPLACEB` validates finite raw arguments before truncating toward zero and
mapping the resulting boundaries. Zero lengths, an end start, and a start
beyond the text follow the contract's empty, sentinel, and append rules.
There is no normalization or host-code-page conversion.

`FINDB` performs a literal, case-sensitive search at source scalar boundaries.
`SEARCHB` performs the pinned Unicode full case fold, including fold
expansions, while mapping a match back to the original UTF-8 byte position.
Its empty-query end sentinel and non-empty-query end behavior are handled
separately. The implementation does not introduce wildcard, regular-expression,
or native-width behavior.

Text numeric arguments use the invariant decimal grammar. Malformed numeric
Text produces generated `#VALUE!`; successfully parsed non-finite Text is
retained long enough for the byte integer validator to produce `#NUM!`. Other
non-finite values and invalid negative positions or lengths retain their
specified `#NUM!` or `#VALUE!` distinction. Integer conversion truncates
toward zero where the selected ordinary counterpart requires it. This scoped
Text path avoids changing the common conversion behavior of unrelated
reducers.

The value integration preserves ordinary scalar and matrix context. Scalar
evaluation uses implicit intersection; matrix evaluation lifts each scalar
argument by output position and applies the normal singleton, row, and column
broadcast rules. The result is a scalar Text or Number at every output
position. `Missing`, `Complex`, formula errors, Empty, Logical, Number, and
Text inputs follow their profile conversions. An admitted provider or
resolver `Unsupported`, resource, cancellation, source, allocation, or
source-version failure remains a typed `EvaluationFailure`; it is not changed
to a formula error and cannot be hidden by `IFERROR` or `IFNA`.

Known `ReferenceList` and unsupported multi-area/3-D descriptors are rejected
before selecting their cells in the direct scalar/matrix byte path. The
focused probes confirm zero resolver reads for these direct refusals. The
contract deliberately permits an admitted computed expression to perform its
own reads before its resulting descriptor is refused; that behavior is
per-argument and does not weaken the direct descriptor gate.

The byte functions remain position-sensitive scalar consumers under projected
`IF` and demand-cache evaluation. The value cache classifier rejects a text
function node while walking a computed matrix/reducer branch, preventing a
`LENB` or `LEN` result for one coordinate from being reused at another. The
frozen regression checks record:

* `AVERAGE(LENB([.A1:.A2]))` -> `[10, 3]` with two reads;
* `AVERAGE(LEN([.A1:.A2]))` -> `[4, 2]` with two reads; and
* `SUM(LENB([.A1:.A2]))` -> `[13, 13]` with four reads, preserving the
  existing complete-matrix aggregate context.

Nested conditional-criterion propagation remains coordinate-aware, and the
scalar size input of `MUNIT` remains position-sensitive. This closes the
observed projected-branch failure without admitting unsafe cross-coordinate
text-result reuse.

## Validation and disposition

The frozen focused targets passed 6/6 evaluation tests, 5/5 limit tests, and
2/2 oracle tests. The independent Python oracle covers 1,376 observations.
The package gate records 1,605 passing tests with zero failures or ignored
tests. Clippy with `-D warnings`, rustdoc, package format, batch rustfmt,
crate-boundary checks, and `git diff --check` all passed; the gate manifest
reports `stable_sources: true` and `all_required_checks_passed: true`.

The native receipt contains 19 observations. All seven ASCII observations
match the selected UTF-8 profile; the documented non-ASCII differences expose
the native application's width/DBCS behavior and are an intentional profile
divergence, not an implementation defect or portability claim. The retained
performance review reports 3,360 frozen samples; this semantic review relies
on its receipt and makes no independent cross-machine timing claim.

No unresolved semantic or integration finding remains for the seven-function
scope. The source is suitable for the frozen handoff.

# Independent semantic review: ODF 1.4 text functions

## Review disposition

The frozen source passes the retained semantic, limit, native, oracle, and
repository gates, but I am holding semantic acceptance for three concrete
`TEXT` grammar cases. The source accepts format forms that the finalized
contract rejects, and it drops a percent marker in another form. These cases
are outside the current 219-row oracle, so the passing receipts do not close
them.

The issues are localized to `text/format.rs`:

* `NumericSpec::parse` accepts both U+0009 tab and ordinary spaces anywhere in
  the numeric core (the `b' ' | b'\t'` arm). The selected grammar admits a
  literal U+0020 space only as the mixed-fraction separator. An unquoted tab
  is also accepted by `has_invalid_date_punctuation`, although date sections
  admit the listed date separators and quoted/escaped literals only. For
  example, `TEXT(1.25;"#\t?/?")` and `TEXT(0;"yyyy\tmm")` should be
  malformed `#VALUE!` under the current contract.
* A percent marker inside the placeholder span is used for scaling but is not
  emitted. For example, `TEXT(1.2;"0%0")` scales to 120 and loses `%`, while
  the contract requires the one permitted marker to be emitted. The profile
  needs either positional emission or an explicit rejection of internal
  markers.
* Repeated percent markers outside the placeholder span bypass the duplicate
  check in `NumericSpec::parse`; `TEXT(1;"0%%")` and `TEXT(1;"%%0")` are
  accepted as literal repeated suffix/prefix text. The contract makes a
  repeated marker a grammar error regardless of its position.

These are semantic blockers for a strict contract disposition. No production
or test files were changed by this review. The focused receipts below remain
useful evidence for the covered profile once the grammar decision is landed
and the selected source is re-frozen.

## Normative and source identity

The contract is the repository-local ODF 1.4 Part 4 profile in
[`contract.md`](contract.md), SHA-256
`b0b7f6ffd8a33c93f98c9728c938eb2d55476680e11797a9b87dcf49223312e7`.
Part 4 leaves `TEXT`'s number-format grammar implementation-defined. The
selected profile nevertheless fixes the accepted tokens, the one-percent
rule, literal handling, and the checked u64 whole/improper numerator bound.

The frozen production hashes are recorded in `gates/freeze.json` and verified
unchanged:

| source | SHA-256 |
| --- | --- |
| `evaluation/text.rs` | `7959904110633b298cc2fcfb3affbe205b91c8f408a864cb57dc7818648f696e` |
| `evaluation/text/core.rs` | `e2832703658582c516a6605fc9197de6af8d636266777bca8e66ab5b025809ba` |
| `evaluation/text/format.rs` | `23c0df92d8659ce9f5630b676fa1e1a1ebae471f110a2597ae92802d59d9619f` |
| `evaluation/text/fraction.rs` | `a2196dd849e2d85faaacd49f876aba8920f9aff00b50e1fad474ae6c278f276e` |
| `evaluation/text/search.rs` | `523b93ce563976bc617c626133e3dd90760d0580b63ef677a7eb73bda0f78d0f` |
| `evaluation/text/unicode.rs` | `336584ece56ee00e9e68a0c2165b5396ac1f007d5ad671f1fe8eba30aac877b6` |
| `evaluation/text/width.rs` | `69a4fae2bb52744e6fa041e3b5de2ffcb9ffa17f6fa7ed2139e433ffe9136b7c` |
| `evaluation/value.rs` | `3f59c37f883cd1a91278c224d3d97edcb86d5a6a3d7ea7c1cc4adfc7848cc55a` |
| `evaluation/value/scalar.rs` | `cd882fa914bffa26ede20c51a4483c53fc36f4070e37cfbfef2bfec6e2db69f3` |

## Semantic and numerical coverage

The reviewed implementation covers the 26-function §6.20 batch with scalar
and value evaluation integration. It uses Unicode scalar positions, the
selected full case-fold SEARCH profile, literal FIND/EXACT/SUBSTITUTE
matching, the selected CLEAN/TRIM/PROPER Unicode tables, and the exact ASC
and JIS mappings. Optional defaults, required missing arguments, source
formula-error precedence, Complex/Text gates, Empty handling, and typed
resolver/evaluator failures are retained through the scalar bridge.

The formatter uses the invariant DOLLAR/FIXED provider and the selected TEXT
section rules: four-way section selection, positional `0`/`#`/`?` slots,
scientific mantissa normalization, quoted/backslash/underscore literals,
Gregorian signed serial dates, and improper versus mixed fractions. The exact
binary64 fraction helper uses checked cross-products, the denominator cap
`10^d-1`, lower-denominator/lower-numerator tie breaking, and the explicit
u64 whole/improper numerator profile. The `2^64` refusal and represented
predecessor are independently retained by the oracle; the boundary guard
rejects the former rather than allowing a saturating cast.

Matrix evaluation preserves the existing scalar projection and array-lifting
context, while references and formula errors retain source order. Output
text remains borrowed when identity-preserving and constructed output is
budgeted through the evaluator's storage/work paths.

## Receipts

The independent oracle verifies 26 functions over 219 observations:

```text
{"functions": 26, "observations": 219, "verified": true}
```

The focused frozen targets pass 13 evaluation tests, 12 resource-limit tests,
2 native-comparison tests, and 26 oracle tests. Root's isolated verifier
reports 1,592 ODS tests passed, zero failed, and zero ignored. All seven
isolated gates pass: the ODS test suite, strict all-target Clippy, rustdoc,
crate formatting, batch formatting, crate-boundary checks, and diff checks.
`gates/verification.json` records stable source manifests and all required
checks passed.

This semantic review makes no performance claim. Any performance disposition
and matched-control threshold evidence belongs to root's separate performance
receipt.

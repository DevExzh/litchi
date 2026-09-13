# Radix evaluator implementation review

Review scope is the fourteen OpenFormula 1.4 Part 4 §6.19 conversion
functions implemented in `evaluation/radix.rs`: `BASE`, `DECIMAL`, and the
twelve binary, decimal, hexadecimal, and octal directed conversions.  This
review covers numerical representation, fixed-width signed conversion,
formula error behavior, bounded work and storage, cancellation, and owned
text lifetime.  It does not review unrelated evaluator families.

## Review snapshot

The reviewed source snapshot is anchored to base commit
`e1976c59d5c4785e0c73a5d27e9349ba082cca2f` and has these SHA-256 digests:

| artifact | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula/evaluation.rs` | `15ff312b132e6d3d0d97a2cd6ff897a8b33fcbd8b366b35e20737c6d7cfeaa84` |
| `crates/litchi-ods/src/codec/formula/evaluation/radix.rs` | `48cf638a72a4946f8f3c8284b34ff817e4c5ba77c962776ad6c2ae981991cf9d` |
| `crates/litchi-ods/tests/ods_formula_radix_evaluation.rs` | `c2d0ee542fdc4e465b6236b1e6234d1b7980791a82995d7c315239af48ba32bf` |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| Part 4 formula HTML entry | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| `numerical-review/check_nearest_even.py` | `cd55d5aeabfd94e936886b82a851bf99004499c8315f3980694283a88455d801` |
| `numerical-review/receipt.json` | `fe5637cf897c2b9cac9d5036be3b84d6275eb380080226b6edc3cd35eb544d4d` |

## Numerical and semantic disposition

`Uint1024` is a fixed 32-limb, 1024-bit unsigned magnitude.  It is large
enough for every finite integer `f64`, including `f64::MAX = (2^53 - 1) *
2^971`, and the implementation does not narrow `BASE` or `DECIMAL` through
`u64`.  Multiplication and division use checked fixed-limb arithmetic with
small radices.  Decimal text is rounded once to `f64` with nearest-even logic;
zero is handled explicitly and values rounding beyond finite `f64` become the
profile's `#NUM!` value.  An independent inline Python oracle compared 19,251
random values across all bit lengths and neighborhoods around the finite
maximum; every finite result and overflow decision matched Python's nearest
even conversion.  The reproducible command and seed are recorded in
`numerical-review/receipt.json`; that receipt explicitly records that the
oracle mirrors the bit decisions and does not execute the Rust binary.

The direct converters parse the documented ASCII digit alphabets, use fixed
10-, 30-, and 40-bit two's-complement widths, preserve the source sign bit,
and reject values outside the target signed range before rendering.  Number
inputs are required to be integral for the twelve direct functions; fractional
Number inputs return `#VALUE!`.  Decimal text accepts one leading minus for
the `DEC2*` family, while binary, octal, and hexadecimal text remains strict.
`Digits` is evaluated eagerly, truncated toward zero for positive results,
bounded to ten digits, and ignored for negative fixed-width results after
formula errors have been propagated.  `DECIMAL` applies only its specified
leading space/tab and radix-specific prefix/suffix rules.

## Bounded resource disposition

Parsing, limb arithmetic, sign conversion, and the reverse digit buffer use
fixed storage.  Input and emitted text bytes are charged to the evaluator's
work budget, long loops check the caller's cancellation context in bounded
chunks, and output length is checked against the finite text limit before
`reserve_storage` and fallible `String` reservation.  Successful output keeps
its reservation in `TextValue` and then in `EvaluatedScalar`; failure paths
drop temporary values through normal ownership.  The radix dispatch adds no
evaluator frame variants or value-stack escape hatch.

The review initially found R1 in the pre-fix snapshot: `apply_decimal` dropped
the input reservation while its owned `Cow<String>` was still live.  The
current source resolves it by keeping the `TextValue` wrapper intact through
parsing and calling `drop(text)` before rounding and evaluator-stack insertion.
`TextValue` declares its text field before its reservation, so the backing
string is dropped before the reservation.  Direct source review confirms that
the corrected path covers both successful and formula-error parse results.

## Receipts checked

The scoped gate receipt `gates/results.json` records status zero for test,
clippy, documentation, doctest, and format checks, with source custody
unchanged.  The final test receipt reports 915 tests, and the focused radix
test file contains the twelve required groups, including the zero and
fractional-input boundaries and 1024-byte storage-limit correction.

The candidate performance capture is status zero for all 141 comparable rows
and all 159 radix rows (53 cases in each of parse, evaluate, and
parse-evaluate).  The radix rows contain 151 successful exact-value checks
and the eight intended typed resource/cancellation refusals.  The candidate
source custody receipt records the corrected `48cf638a…` module digest.
These receipts establish bounded behavior for the captured corpus.

**Verdict:** numerically and semantically satisfactory after the documented
profile choices; the numerical, fixed-width, cancellation, and storage
ownership review is closed with no remaining blocker.

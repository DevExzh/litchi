# Independent text-function oracle review

Status: **accepted scoped verdict**. This review owns
`text_oracle.py`, `text-goldens.json`, and
`crates/litchi-ods/tests/ods_formula_text_oracle.rs`. It does not derive an
expected value from the Rust text implementation, LibreOffice, or a host
locale, and it makes no production-source change.

The semantic boundary is the repository-local ODF 1.4 text-function contract,
currently SHA-256
`b03ec074efe30c4261509a70b6188010d74a9d48d8bc92a57ad39a2db4a3b611`.
The normative ODF archive is
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`, and the
retained Part 4 HTML member is
`ace07938ef54303b57af8472e0b66b289fc6946c32440fc23b8e13fdeeb5ffa1`.

The corpus covers all 26 §6.20 names: `ASC`, `CHAR`, `CLEAN`, `CODE`,
`CONCATENATE`, `DOLLAR`, `EXACT`, `FIND`, `FIXED`, `JIS`, `LEFT`, `LEN`,
`LOWER`, `MID`, `PROPER`, `REPLACE`, `REPT`, `RIGHT`, `SEARCH`, `SUBSTITUTE`,
`T`, `TEXT`, `TRIM`, `UNICHAR`, `UNICODE`, and `UPPER`. It contains 244 typed
observations (TEXT 118; DOLLAR/FIXED 8 each; SEARCH 9), with compact repeated
boundary coverage for each function.

## Independent derivation

The generator keeps every input as a typed JSON value. Number inputs are
constructed from and compared by their IEEE-754 binary64 bit pattern. Text
results are compared as exact Unicode scalar sequences encoded as UTF-8;
logical and formula-error results are exact typed values. There is no numeric
tolerance in this corpus.

The boundary rows also include Complex producers consumed by `CONCATENATE`,
`T`, and `TEXT`, missing required input, Empty and scalar reference inputs,
formula-error precedence, and a known ReferenceList. Fractional `REPLACE`
boundary rows remain intentionally outside this scoped corpus. These rows keep
Complex rejection, generic conversion, and structural refusal separate from
ordinary text conversion.

The scalar reference implements the ODF Table 33 and Table 34 ASC/JIS maps
directly, including voiced and semi-voiced Katakana, Yen/reverse-solidus, and
quotation exceptions. Slicing, length, search, replacement, and repetition
operate on Unicode scalar values. The corpus includes astral and combining
characters, a normalization-variant pair, an unassigned scalar, control and
format characters, NBSP, full-width and half-width Katakana, empty needles,
literal case-sensitive FIND, case-insensitive SEARCH, expansion cases such as
`Straße` and dotted-I, partial-expansion rejection (`s` in `ß`), reverse
expansion (`ß` in `ss`), overlapping-boundary fallback, and final-sigma
context.

Python 3.14 in the gate environment exposes UCD 16.0.0. The selected
Unicode-17 additions used by this independent oracle are pinned explicitly:
the official simple mappings include lowercase `U+A7CE → U+A7CF`,
`U+A7D2 → U+A7D3`, and `U+A7D4 → U+A7D5`, with inverse uppercase mappings.
The retained Unicode-17 source receipts and generated evaluator table are in
`unicode-data/provenance.json`; the oracle does not import the Rust table.
Ordinary Unicode casing uses Python's default context-sensitive algorithms,
with the pinned additions applied after that independent operation. CLEAN
uses the selected `Cc`/`Cn` category rule, TRIM uses only tab, LF, CR, and
space, and PROPER uses letter boundaries without normalization.

The formatting rows use the contract's explicit invariant provider: `$`, comma
grouping, period decimals, two default DOLLAR/FIXED places, half-away rounding,
negative DOLLAR parentheses, and TEXT sections covering optional/required
placeholders (including aligned `?` spaces), grouping, percent, scientific
notation (including subnormal and max-finite inputs), simple fractions with
`Fraction.limit_denominator`, quoted/backslash/underscore literals, four-way
positive/negative/zero/Text selection, numeric-`@` selection, Gregorian
date/time serials around 1900 and 9999 including signed dates, Logical/Text/
Complex/Empty inputs, quoted and backslash-escaped `AM/PM` literals, unquoted
tab/stray-space grammar refusal, repeated-percent refusal, literal quoted and
escaped percent, positional percent emission across fixed/fractional/rational
spans and a scientific suffix, atomic scientific-token rejection when `%`
splits the E/sign/slot sequence, zero/optional-slot cases, positional mixed
placeholders, improper and mixed fractions, normalized scientific mantissas,
date underscore padding, the exact binary64 `2^64` fraction overflow boundary,
and malformed tokens. The
reference formatter is
implemented with Python `Decimal`, `ROUND_HALF_UP`, and represented-binary64
`Fraction` arithmetic, independently of the Rust rounding or format parser.
The finalized profile permits one through six denominator placeholders and
uses the strict bound `10**placeholder_count - 1`; the `0.1`/`# ?/?` row
therefore records `1/9`, while a near-one value records carry to `1`.

Typed reference rows use a resolver that returns borrowed text and counts every
cell read. Both Scalar and Matrix value modes are exercised. Scalar mode
projects a one-cell reference; Matrix mode checks the corresponding one-cell
rectangular result and its read receipt. A known ReferenceList refusal and a
required Missing argument are retained as zero-read structural/error cases.
Formula errors remain formula values; the target does not catch or rewrite
typed evaluator failures.

## Comparison and validation policy

The policy was fixed before the evaluator run:

* Text is exact UTF-8, including NUL, combining marks, astral scalars, and
  case-mapping expansions.
* Number results require exact binary64 bits, including zero sign.
* Logical and formula-error results require exact enum values.
* Resolver read counts require the exact retained `expected_reads` value in
  both evaluation modes.

The independent byte check currently passes:

```text
python3 text_oracle.py --check
{"functions": 26, "observations": 244, "verified": true}
```

The focused target compiles with `cargo test --locked --offline -p litchi-ods
--test ods_formula_text_oracle --no-run`, and the latest normative run passes
all 26 function tests over 244 observations in both Scalar and Matrix modes.
That receipt includes the tab/space and repeated-marker grammar refusals,
quoted/escaped and @ literal-percent rows, positional marker emission in
fixed/fraction/rational forms, scientific suffix emission, atomic scientific
token rejection for markers inside `E`/sign/slot syntax, zero/optional-slot
marker retention, positional mixed placeholders, improper versus mixed
fraction rendering and rational `?` padding, normalized scientific mantissa
padding, underscore date padding, and the `2^64` fraction overflow boundary
with a representable predecessor.
The earlier formatter mismatches for `?` alignment, strict fraction bounds,
malformed grammar, numeric `@`, quoted `AM/PM`, grouping, and signed dates are
also covered by the passing run. The source and contract support this accepted
scoped verdict.

The retained LibreOffice receipt contains 91 bounded native observations for
cross-checking ordinary behavior. It is not used to define the contract,
Unicode profile, error precedence, matrix shape, or formatting provider.

Current owned hashes are:

```text
text_oracle.py                 597a7fa70a56fb8f96d1826d9463b1e54f2325963511ffeafc877abb248bb19c
text-goldens.json              75c1e87100cc7f2bb999a9cf9d9b35f1ab202bc7a834493285f6bace0240b0c4
ods_formula_text_oracle.rs     5341f4dbf29e9ecb5bc04505313971535398e2cc17840ebe124301f80767e703
contract.md                    b03ec074efe30c4261509a70b6188010d74a9d48d8bc92a57ad39a2db4a3b611
```

These hashes are the current grammar-review custody record. The generator,
goldens, focused target, and this review are byte-stable for handoff; no
production or diagnostic edits were made by this oracle owner.

Root follow-up: after the oracle agent reached its usage limit, root added
five zero-field percent cases and completed the corresponding independent
Python rule for a required zero numerator with a suppressed denominator.
The 244-observation byte check and all 26 Rust oracle targets pass. These
last additions were authored and verified by root, not independently reviewed
by the oracle agent. The earlier review is retained under diagnostics.

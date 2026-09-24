# ODS formula string literal performance review

Status: bounded review complete; no blocking finding remains in the literal
decoder. The review covers decoded string sizing, doubled-quote and UTF-8
semantics, NUL/unterminated refusal ordering, and the existing formula limits.
It does not make a whole-parser time or memory claim and does not review the
remaining OpenFormula grammar.

## Reviewed snapshot

The baseline is commit `cbc60f1123d105f67f4abdd0d53034fb3325c25f`. The frozen
candidate and focused test source digests are:

| Path | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula.rs` | `0b59a9f20b2c9c62fadce3172c646178f8c3cc596c594a48787a0250a651e748` |
| `crates/litchi-ods/tests/ods_formula_string_regression.rs` | `d58abe8eaa7bd9e7ba96139407d9ffcfe5f57350302a5ee795df5297cf7597ac` |
| `crates/litchi-ods/src/codec/formula/reference.rs` | `3707b3176e2714aa794bf28628eb27ec72923e001faaab22a987883b596bd27e` |

The reference parser is included only as an unchanged source-boundary
receipt; this batch changes the string path in `formula.rs`.

## Normative rule

The bundled source is `3rdparty/specs/OpenDocument-v1.4-os.zip`, whose SHA-256
is `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`.
The Part 4 formula HTML entry
`part4-formula/OpenDocument-v1.4-os-part4-formula.html` has SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
Part 4 §5.4 defines the relevant production as:

```text
String ::= '"' ([^"#x00] | '""')* '"'
```

Thus quotes delimit the value, a doubled quote contributes one quote to the
decoded value, and U+0000 is excluded. Other characters accepted by the
production, including line feed, tab, and other non-NUL control characters,
remain accepted by this bounded change.

## Decoder and allocation review

`parse_string` now performs a syntax pass before creating the payload
`String`. It locates the first unpaired closing quote and counts
doubled-quote pairs, then checks the admitted content slice for U+0000. For a
content span of `S` source bytes and `P` doubled-quote pairs, the decoded UTF-8
byte length is exactly `S - P`: every ordinary byte, including every byte of a
multibyte UTF-8 scalar, is retained; each two-byte ASCII quote pair becomes one
output byte. The subtraction and pair count are checked.

Only after that pass succeeds does the decoder make one fallible
`try_reserve_exact(decoded_bytes)` call. Empty literals skip the reservation
and retain `String::new`'s zero-capacity state. The second pass copies complete
UTF-8 spans and collapses each already-admitted doubled pair. It uses checked
`from_utf8` conversions and does not use an unsafe byte-copy or an intermediate
decoded buffer. The outer parser validates the borrowed formula as UTF-8
before it allocates the retained original text, so a valid public input cannot
split a scalar while the literal scanner advances bytewise.

This removes the baseline's reservation of the entire remaining formula for
each literal. The reservation is now proportional to that literal's decoded
content, subject only to allocator rounding. No additional literal scratch
storage is retained.

## Error and limit ordering

`parse_with_limits` keeps its established order: validate the input as UTF-8,
check `FormulaLimits::max_bytes`, reserve and copy the original formula text,
then tokenize. During tokenization, the pre-existing token-count admission
check still runs before `next_token`. Consequently, a zero-token limit can
refuse before entering string parsing, and a formula over the byte limit is
refused before the retained text allocation.

Once a string token is reached, an unterminated literal is returned as
`Error::InvalidFormat` by the locating pass. For a complete literal, a NUL is
returned as `Error::InvalidFormat` by the following content check, before the
literal reservation. The original formula text has already been retained, as
it was before this change. If an input is both unterminated and contains NUL,
the unterminated error wins because no complete content span exists yet; both
cases remain typed syntax refusals. A valid literal allocation failure remains a typed
`Error::Allocation` with resource `formula string literal`; the focused
allocator test exercises that path. Successful parsing leaves the original
formula spelling in `Formula::text`, advances past the closing quote, and
retains the decoded token value.

## Compatibility cases and validation

The focused regression target covers many short and empty literals, UTF-8
values, doubled quotes, formula source retention, NUL refusal with LF/tab and
U+0001 acceptance, malformed and unterminated quote runs, exact and
one-under byte/token limits, and a targeted typed allocation failure. All six
tests passed. The frozen package receipt at
[`gates/results.json`](gates/results.json) records status zero for test,
clippy, documentation, doctest, and formatting gates, with `828` tests across
`50` targets and unchanged reviewed source hashes.

The paired literal workload artifacts retain the baseline cases for short,
empty, UTF-8/doubled-quote, 64 KiB, and unterminated strings. They support the
bounded allocation claim above; this review does not extrapolate those cases
to unrelated parser paths.

Disposition: no additional blocker was found in this bounded string-literal
review.

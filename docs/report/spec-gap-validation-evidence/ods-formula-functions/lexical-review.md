# ODS OpenFormula function and lexical review

Reviewed 2026-09-13 UTC against the candidate formula lexer:

- `crates/litchi-ods/src/codec/formula.rs` SHA-256:
  `0c3108e95e4b5fd92e616917171ebd840f26e103a2076e27592a9a330f2bd647`
- `crates/litchi-ods/src/codec/formula/functions.rs` SHA-256:
  `1d736271e9f9cf3743295f8895f184e32842640eaba48084dfafc9ca7333e2a1`

This review covers the bounded lexical recognition changes only.

| Area | Disposition | Evidence and boundary |
| --- | --- | --- |
| Canonical function lookup | Closed | `lookup_function` checks the bounded input envelope, uses the PHF catalog directly for canonical ASCII names, and uses a fixed stack buffer for ASCII case folding and the retained Unicode-uppercase compatibility behavior. No caller-controlled uppercase `String` is allocated. |
| Catalog and scope claims | Closed | `STANDARD_FORMULA_FUNCTIONS` contains 393 canonical Part 4 chapter 6 names. Module and public-query documentation state that recognition is lexical, with no evaluation, arity validation, or complete expression-grammar claim. |
| Function-call lookahead and cursor rollback | Closed | `try_parse_function_call` scans a complete ASCII identifier, permits whitespace before `(`, and restores the original cursor when the identifier is not a recognized invocation. The 19-byte guard returns through the same rollback path for longer invalid identifiers; every catalog name is at most 19 bytes. The existing speculative `try_parse_cell_ref` path still rewinds on failure. |
| Bare-cell-reference preservation | Closed | A1-shaped identifiers without an invocation parenthesis continue through `try_parse_cell_ref`; the tested `LOG10` spelling remains column `LOG`, row `10`, while `LOG10(...)` and dotted digit-bearing catalog names become function tokens. |
| UTF-8 string literals | Closed | `parse_string` copies complete UTF-8 slices between ASCII quote delimiters, converts doubled quotes to one quote, preserves Unicode text, and returns an error when no terminating quote remains. No byte is decoded as an independent character. |
| Grammar and evaluation boundary | Closed | The changes only affect function-name recognition and string tokenization. Token extraction remains inert; arguments are neither evaluated nor validated for arity, and no source access or publication is introduced. |

Validation passed:

- `cargo test -p litchi-ods --test ods_formula_functions` — 7 tests
- `cargo test -p litchi-ods --lib codec::formula::tests` — 30 tests

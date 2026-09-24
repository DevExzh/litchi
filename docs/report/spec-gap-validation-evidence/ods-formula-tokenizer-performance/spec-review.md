# ODS formula tokenizer boundary review

This review covers the tokenizer optimization candidate based on commit
`b88341bf4` and the adjacent reference-parser annotations. It is limited to
legacy function/cell tokenization, compatibility handoff, and typed failure
propagation; the OpenFormula reference grammar was reviewed separately.

## Reviewed source

The reviewed source digests are SHA-256 values:

| Path | SHA-256 |
| --- | --- |
| `crates/litchi-ods/src/codec/formula.rs` | `4084ab6a41f61477aeaa15d3dbe896544a15147b75b2e1028d079569f52a0e0a` |
| `crates/litchi-ods/src/codec/formula/functions.rs` | `1d736271e9f9cf3743295f8895f184e32842640eaba48084dfafc9ca7333e2a1` |
| `crates/litchi-ods/src/codec/formula/reference.rs` | `3707b3176e2714aa794bf28628eb27ec72923e001faaab22a987883b596bd27e` |
| `crates/litchi-ods/src/codec/formula/reference/iri.rs` | `a09fa610c25de63e5665f9de88698d7578f8b80c70ca027382ddb4d13ae05961` |
| `crates/litchi-ods/tests/ods_formula_tokenizer_regression.rs` | `55acf738b05b86c375e3bf96b31c7c0f6378c1f4554ba9d8842e78b6cd7814bc` |

`reference.rs` contains only the reviewed `inline`/`cold` code-generation
annotations relative to the preceding reference candidate; its parsing,
limits, and error behavior are unchanged.

## Findings

The compact path scans one bounded ASCII candidate and gives a known function
precedence only when the complete name is followed by optional ASCII
whitespace and `(`. It leaves that whitespace for the ordinary token loop, so
`SUM (A1)`, lowercase calls, and dotted calls such as
`BINOM.DIST.RANGE(A1;1;2)` retain their token boundaries. A bare
cell-shaped name remains a cell: `LOG10` is `LOG` row 10, while `BIN2DEC`
falls back to the established partial-cell/name behavior. Malformed suffixes
such as `LOG10X(` are refused as syntax errors.

The compact cell parser admits only a complete column-plus-decimal-row
spelling (with an optional compact sheet prefix) and checks the row conversion
before copying owned components. Inputs containing `$` are deliberately sent
through the legacy parser, which preserves column and row absolute flags.
Malformed suffixes (`A1B2`, `Sheet1.A1B2`, and dotted partial forms) likewise
rewind to that parser, retaining its established partial-token behavior.

The literal-space handoff is required by the legacy unquoted sheet scanner.
When a compact cell is followed by a contiguous `[A-Za-z0-9_ ]*.` locator,
the candidate returns to the legacy parser; this preserves `Sheet1 .A1`,
`A1 .B2`, and multi-space sheet names. Tabs are excluded from this handoff,
matching the legacy scanner, which treats them as formula whitespace.

Speculative cell parsing now rewinds only `Error::InvalidFormat`. Allocation
failures and other typed errors from sheet/column component copies propagate
to the caller instead of being reinterpreted as a name. The focused allocator
regression exercises compact, current-sheet, and absolute legacy paths.

No blocking semantic finding remains in this candidate. The compatibility
handoff can rescan a long contiguous space/alphanumeric suffix for every cell
in an input such as repeated `=A1 A1 A1 ...`; that retains the legacy worst-case
quadratic behavior for this unusual space-separated spelling. The ordinary
operator-delimited compact path has the intended single-candidate behavior, so
performance claims should exclude the compatibility fallback rather than call
every repeated-cell input linear.

## Validation

The focused tokenizer receipt records four passing tests, including function
and cell boundary cases, spaced-sheet preservation, malformed suffixes, and
allocator failure propagation. The formula unit tests and the package-wide
`--all-features --all-targets` test, clippy, docs, doctests, and formatting gates
also passed in the frozen receipt at
[`gates/results.json`](gates/results.json).

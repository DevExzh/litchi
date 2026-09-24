# ODS OpenFormula reference-family review

Status: bounded normative and parser integration review; no production changes in this review.

## Sources

The normative source is the bundled OpenDocument v1.4 formula work product (6 October
2025), Part 4 §5.8, `References`:

- package: `3rdparty/specs/OpenDocument-v1.4-os.zip`
- package SHA-256: `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`
- Part 4 HTML SHA-256: `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`
- HTML anchor: `#a_5_8_References` (also `#References`)
- normative HTML anchors: `#a_5_8_References`, `#a_3_7_Basic_Limits`,
  `#a_5_13_Inline_Arrays`, and `#a_7_2_Inline_constant_arrays`

The Source value is normatively an RFC 3987 IRI-reference under §5.8. The primary
reference is [RFC 3987 §2.2 at the RFC Editor](https://www.rfc-editor.org/rfc/rfc3987.html#section-2.2),
especially its `IRI-reference` and `irelative-ref` productions.

## Current implementation boundary

At review baseline `b7a66574a` (before the candidate reference files were wired),
`crates/litchi-ods/src/codec/formula.rs:32-81` exposed `CellRef`, `RangeRef`, and a
closed `Token` enum with no source, whole-row, whole-column, cuboid, subtable, or
reference-error representation. The parser's bracket path is at lines 400-428; it
captures a payload and delegates to `parse_open_formula_reference` at lines 616-762.
That helper splits at an unquoted colon and then reverse-splits at the last dot, so it
can handle the existing simple cell/range subset but cannot represent the §5.8 grammar.
`extract_cell_refs` at lines 782-790 would silently omit any new token unless a new
reference accessor is added.

`crates/litchi-odf-formula` is the MathML/StarMath formula-content crate. Its formula
validation explicitly leaves spreadsheet OpenFormula to `litchi-ods`; it is not an
existing owner for this feature.

The audit's `~150` function-catalog wording in `docs/report/spec-gap-audit.md:716-722`
is stale after the separate 393-name catalog work. The remaining grammar gap is still
real, including source-qualified references and arrays.

## Normative §5.8 contract

The constant reference production is:

```text
Reference ::= '[' (Source? RangeAddress) | ReferenceError ']'
```

All six `RangeAddress` alternatives must remain distinguishable:

```text
1. SheetLocatorOrEmpty '.' Column Row (':' '.' Column Row )?
2. SheetLocatorOrEmpty '.' Column ':' '.' Column
3. SheetLocatorOrEmpty '.' Row ':' '.' Row
4. SheetLocator '.' Column Row ':' SheetLocator '.' Column Row
5. SheetLocator '.' Column ':' SheetLocator '.' Column
6. SheetLocator '.' Row ':' SheetLocator '.' Row
```

The first alternative covers a cell and a same-locator cell range; a right-hand
locator is omitted and inherits the left locator. Alternatives 2 and 3 are
whole-column and whole-row ranges. Alternatives 4-6 are two-locator cuboids for
cells, columns, and rows. A missing locator means the current sheet. A locator on
the left with no locator on the right inherits; two explicit locators are retained
as a cuboid, not flattened to two unrelated cells.

The locator grammar is:

```text
SheetLocatorOrEmpty ::= SheetLocator | /* empty */
SheetLocator ::= SheetName ('.' SubtableCell)*
SheetName ::= QuotedSheetName | '$'? [^\]\. #$']+
QuotedSheetName ::= '$'? SingleQuoted
SubtableCell ::= (Column Row) | QuotedSheetName
Column ::= '$'? [A-Z]+
Row ::= '$'? [1-9] [0-9]*
ReferenceError ::= "#REF!"
```

The `[^\]\. #$']+` class excludes `]`, `.`, SPACE, `#`, `$`, and `'`. It does not
exclude `:`, `[`, or `\\`; in particular, an unquoted sheet name such as
`Sheet:Name` is grammar-valid in `[Sheet:Name.A1]`. Therefore a parser cannot split
the payload at the first unquoted colon. It must first recognize the endpoint/locator
shape, while respecting quoted segments. The same applies to a range such as
`[Sheet:Name.A1:.B2]`, where the colon in the sheet name and the range colon have
different roles. The right-hand leading dot is required by alternative 1 when the
sheet locator is inherited; `[Sheet:Name.A1:B2]` is not that form (an explicit
second locator would instead be `[Sheet:Name.A1:Sheet2.B2]`).

The `SingleQuoted` production is defined in §5.2 as:

```text
SingleQuoted ::= "'" ([^'] | "''")+ "'"
```

The `+` requires at least one content item. An empty quoted sheet or subtable,
`[''.A1]`, is therefore not accepted by this production. This must not be conflated
with an empty Source IRI: RFC 3987 permits the empty `irelative-ref`, so
`[''#.A1]` is valid Source syntax. Part 3 §19.677.14 gives `table:name` the generic
`string` datatype and supplies no narrower host sheet-name regex; reference parsing
should preserve the grammar spelling and defer actual table lookup to the host.

Whitespace needs an explicit boundary. §5.14 permits whitespace only at the positions
allowed by the surrounding grammar and forbids it inside a terminating lexical rule.
The new reference path may inherit whitespace around the formula/reference expression,
but must not call `.trim()` on the bracket payload or accept whitespace inside an
unquoted sheet name, `Column`, `Row`, or the Source marker. Whitespace inside a quoted
sheet name is part of that name; ASCII whitespace in a Source IRI is outside RFC3987,
while Unicode characters such as U+00A0 remain eligible through `ucschar`. The
current codec's global formula whitespace skipping and public-wrapper trim are
compatibility behavior and should not be presented as complete §5.14 conformance. NUL is also outside XML 1.0
formula input; the Source validator's RFC control rejection covers it, while the
existing general string-token path's direct-in-memory NUL behavior remains outside
this narrow reference slice.

## Source and RFC 3987 boundary

`Source ::= "'" IRI "'" "#"`. The source IRI is decoded by replacing every pair of
consecutive apostrophes with one apostrophe before applying RFC 3987 §2.2's generic
syntax. A scanner should consume the opening quote, decode `''` pairs, and treat only
an unpaired apostrophe followed immediately by the Source marker `#` as the closing
quote. This also handles apostrophes and `#` inside the IRI without confusing them
with the outer reference grammar. For example, `['file:///O''Brien.ods'#.A1]` has the
decoded IRI `file:///O'Brien.ods`.

The validator must accept absolute, relative, empty, and same-document-fragment
IRI-references, for example `['../book.ods'#.A1]`, `['#frag'#.A1]`, and
`[''#.A1]`. It must reject malformed percent escapes (`%ZZ`), ASCII controls and
ASCII whitespace, invalid Unicode code points, and characters that are outside the
RFC3987 generic syntax. RFC3987 describes a Unicode character grammar, not a
byte-only URL grammar: U+00A0 (NBSP), for example, is an allowed `ucschar` and must
not be rejected merely because a host language classifies it as whitespace. The
grammar also permits relative references and an empty path; requiring a scheme or a
non-empty value would be incorrect.

No existing validator is sufficient. `crates/litchi-oth/src/codec/structure.rs:3657-3664`
only rejects controls and whitespace. The workspace `url` crate is not a RFC3987
validator, rejects/normalizes relevant relative IRI forms, and is not a dependency of
`litchi-ods`. Use a small pure, bounded lexical helper (private to formula initially,
or promoted to a shared core utility later) implementing the RFC3987 generic syntax:
scheme/relative-path forms, authority/IP-literal grammar, percent-encoding, query,
fragment, and the RFC3987 Unicode character ranges. This is lexical conformance only;
IRI normalization, resolution, opening, and external dependency fetching remain
host-defined/inert per §5.8. A controls/whitespace check alone must be documented as a
guard and must not be called RFC3987 validation.

## Candidate IRI validator review

The candidate `crates/litchi-ods/src/codec/formula/reference/iri.rs` was reviewed
against RFC 3987 §2.2 at blob SHA-1
`f7e65f4f167e9fc8ee3f37ea4d295c1e41026eda`. The helper covers the absolute and
relative reference branches, scheme first-match handling, authority/user-info,
IPv4, IPv6, IPvFuture, path, query, fragment, percent-encoding, and the exact
`ucschar`/query-only `iprivate` ranges. Its component-specific ASCII sets preserve
the RFC distinction between path, query, fragment, authority, and first relative
segment. It is allocation-free and deliberately does not resolve, normalize, or
apply scheme policy. The seven in-module validator tests pass when compiled as a
standalone Rust test binary. No generic RFC 3987 syntax blocker was found in this
bounded review. The caller remains responsible for ODF doubled-apostrophe decoding
before validation and for the surrounding formula-reference limits.

The source parser must decide whether a leading quoted segment is Source or
QuotedSheetName by its following marker (`#` for Source, `.` for a sheet), rather than
classifying every leading apostrophe as Source. A quoted sheet may contain doubled
apostrophes, `]`, `:`, spaces, and dots; an IRI may contain `:` and `#` inside its
quotes. Do not use a first-colon split before Source extraction.

## Final parser integration review

The final review covers the following Git blob SHA-1 identifiers:

- `crates/litchi-ods/src/codec/formula.rs` — `8b03fd9f4583a3b65c1dda785bf79958104b387c`
- `crates/litchi-ods/src/codec/formula/reference.rs` — `e57b60b6c90c50019cee6b7706a7f6d96e38ab15`
- `crates/litchi-ods/src/codec/formula/reference/iri.rs` — `f7e65f4f167e9fc8ee3f37ea4d295c1e41026eda`
- `crates/litchi-ods/tests/ods_formula_references.rs` — `dab747f16c45479a04a9120683a66d629baaeaba`

The focused integration gate `cargo test -p litchi-ods --test
ods_formula_references` passes 10 tests. The parser preserves all six §5.8
range alternatives, explicit and inherited locators, colon-bearing sheet names,
quoted Unicode/delimiter content, paired apostrophe decoding, source boundaries,
and checked row/column axes. Malformed mixed endpoints, suffixes, zero or
overflowing rows, lower-case rich columns, and invalid source syntax are refused.
The formula bracket scanner advances only over UTF-8-safe byte boundaries and
keeps `]` inside quoted components; no additional parser integration blocker was
found in this bounded review.

The public `reference::parse_with_limits` trims only whitespace surrounding the
complete bracket wrapper for compatibility. Its `parse_body` path receives the
interior bytes verbatim, so whitespace cannot be accepted inside an unquoted
sheet or coordinate; formula-level trimming remains the pre-existing outer
formula compatibility behavior. The recorded `max_bytes` reference budget is
the bracket body, as exercised by the focused boundary test.

The final formula allocation refinement keeps token-limit admission before
`next_token`, while reserving the token vector only after a token parses
successfully. This preserves zero-budget precedence and malformed-token error
classification without changing the accepted token stream.

## Discovery proposal (implemented by this batch)

The next coherent slice should be an inert §5.8 reference parser in `litchi-ods`:
parse all six range forms, source decoding plus complete RFC3987 lexical validation,
quoted/unquoted sheet locators, subtable segments, locator inheritance, and `#REF!`.
Do not resolve or fetch external workbooks, evaluate references, implement named
expressions, or claim reference-list (`~`) and intersection/automatic-label semantics.
This is a direct extension of the existing bracket path and fits the Small Group
syntax boundary; `litchi-odf-formula` need not change.

Existing public structs should not gain a `source` field casually: adding it changes
downstream struct literals and still cannot model row/column endpoints or cuboids.
A dedicated `FormulaReference`/`FormulaRange` representation and a new token/accessor
are safer. If compatibility is required, retain old `CellRef`/`RangeRef` tokens for
the old simple local forms and expose qualified/whole-range forms through a dedicated
variant plus `extract_references`; do not silently drop them through
`extract_cell_refs`. Preserve the original formula text for round-trip purposes and
store a bounded decoded Source value only when exposing it in the parsed model.

Resource handling must be explicit. Reuse the caller's formula text budget and charge
the source spelling/decoded value once; do not allocate based on an unchecked length.
Parse `Row` with checked decimal accumulation into the existing `u32`-compatible
model (reject overflow rather than wrapping), and check column-index conversion with
overflow/bounds handling. The ODF basic-limit floor is 1024 formula characters,
30 list parameters, 32,767 ASCII string characters, and seven nesting levels
(§3.7, HTML anchor `#a_3_7_Basic_Limits`); a bounded parser may impose a larger
project ceiling but must not claim these limits while accepting unchecked input.
An explicit row/column value beyond the evaluator's supported capability is an Error,
not a valid reference (§5.8); with no evaluator, the inert layer should preserve this
as a typed refusal/error rather than silently truncating it.

Minimum regression cases should include:

- valid `[.A1]`, `[$Sheet1.$A$1]`, `[Sheet:Name.A1]`, `[Sheet:Name.A1:.B2]`,
  `['Q1 Sales'.$C$3:.$D$4]`, `[Sheet.A1.B2]`, `[.A:.C]`, `[.1:.3]`,
  `[$S1.$A$1:$S2.$B$2]`, `['../book.ods'#.A1]`, `['#frag'#.A1]`, `[''#.A1]`,
  and `['file:///O''Brien.ods'#.A1]`;
- refusal for `[.a1]`, `[.A01]`, `[.A0]`, `[''.A1]`, malformed source quoting,
  `%ZZ`, a source missing its outer `#`, and malformed/ambiguous endpoint shapes;
- preservation of a quoted sheet containing `]` and a quoted subtable segment,
  rather than terminating at the inner `]` or flattening locator ancestry.

Inline arrays are a separate follow-on. §5.13 defines `{ MatrixRow ( '|' MatrixRow )* }`
with `;` between expressions in a row; §7.2's constant-array capability requires at
least one rectangular row/column and Number, Text, TRUE/FALSE, and Error constants.
§7.3 requires full Expression syntax for non-constant arrays, and §3.3 adds evaluation
intersection/iteration and matrix broadcasting. A flat token stream and the existing
`CellMatrixSpan` metadata cannot claim those semantics, so arrays should not be bundled
into the Source/reference slice.

## Discovery disposition and final closure

Source-qualified references are a viable next grammar feature only with the pure
RFC3987 lexical helper and a structured reference model. The discovery blockers for
an implementation that merely adds `Source` text are: first-colon splitting; failure
to decode paired apostrophes; rejecting valid empty/relative/same-document IRIs;
accepting malformed percent/Unicode syntax; and losing locator kind, subtable ancestry,
row/column range kind, or `#REF!`. The candidate IRI helper review clears the
RFC3987-specific items above. The final parser integration review recorded above
closes the surrounding §5.8 representation review with no remaining blocker,
within the documented bounds. Broader expression and evaluation semantics remain
separate. This reviewer did not change production implementation.

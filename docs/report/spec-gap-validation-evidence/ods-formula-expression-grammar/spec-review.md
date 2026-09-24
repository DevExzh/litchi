# OpenFormula expression grammar specification review

Status: normative and bounded implementation review complete. This review
covers ODF 1.4 Part 4
§5.1–§5.14 and the API and resource constraints for a bounded, non-evaluating
expression tree. It does not certify an evaluator, recalculation, host lookup,
or external I/O.

## Source and baseline

The normative source is `3rdparty/specs/OpenDocument-v1.4-os.zip`, SHA-256
`9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`. The
entry `part4-formula/OpenDocument-v1.4-os-part4-formula.html` has SHA-256
`ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`.
The reviewed baseline is commit
`b609407cc672896f6954491ac9501cd1a131e941`; its existing
`crates/litchi-ods/src/codec/formula.rs` digest is
`0b59a9f20b2c9c62fadce3172c646178f8c3cc596c594a48787a0250a651e748`.
The HTML anchors are `#a_5_1_General`, `#a_5_2_Basic_Expressions`,
`#a_5_3_Constant_Numbers`, `#a_5_4_Constant_Strings`,
`#a_5_5_Operators`, `#a_5_6_Functions_and_Function_Parameters`,
`#a_5_7_Nonstandard_Function_Names`, `#a_5_8_References`,
`#a_5_9_Reference_List`, `#a_5_10_Quoted_Label`,
`#a_5_11_Named_Expressions`, `#a_5_12_Constant_Errors`,
`#a_5_13_Inline_Arrays`, and `#a_5_14_Whitespace`.

## Required grammar

Section 5.2 defines one formula expression, with an optional introduction:

```text
Formula ::= Intro? Expression
Intro ::= '=' ForceRecalc?
ForceRecalc ::= '='
Expression ::=
    Whitespace* (
        Number | String | Array | PrefixOp Expression |
        Expression PostfixOp | Expression InfixOp Expression |
        '(' Expression ')' | FunctionName '(' ParameterList ')' |
        Reference | QuotedLabel | AutomaticIntersection |
        NamedExpression | Error
    ) Whitespace*
```

The recursive BNF is disambiguated by §5.5's precedence table. A checked
parser must admit exactly one root and consume the complete body. The existing
flat token loop accepts adjacent roots, dangling operators, empty/mismatched
parentheses, and trailing separators; those are grammar-validation gaps.

Section 5.3 defines numbers as:

```text
Number ::= StandardNumber |
           '.' [0-9]+ ([eE] [-+]? [0-9]+)?
StandardNumber ::= [0-9]+ ('.' [0-9]+)? ([eE] [-+]? [0-9]+)?
```

Signs are prefix operators, so `.5` is valid, `1.` is not, and an exponent
requires digits after its optional sign. A syntax tree may retain the numeric
lexeme instead of converting it to finite `f64`; conversion is evaluator
behavior and must not turn a syntactically valid value into a parser refusal.
Section 5.4 is:

```text
String ::= '"' ([^"#x00] | '""')* '"'
```

The decoder must preserve UTF-8 and non-NUL content while collapsing doubled
quotes.

## Operators and precedence

Section 5.5 defines:

```text
PrefixOp ::= '+' | '-'
PostfixOp ::= '%'
ArithmeticOp ::= '+' | '-' | '*' | '/' | '^'
ComparisonOp ::= '=' | '<>' | '<' | '>' | '<=' | '>='
StringOp ::= '&'
IntersectionOp ::= '!'
ReferenceConcatenationOp ::= '~'
RangeOp ::= ':'
```

`InfixOp` is the union of the arithmetic, comparison, string, and reference
operators. Table 1 orders operators from highest to lowest precedence as
follows: `:` left; `!` left; `~` left; prefix `+`/`-` right; postfix `%` left;
`^` left; `*`/`/` left; binary `+`/`-` left; `&` left; and all comparisons
left. In particular, unary minus binds tighter than `^`, exponentiation is
left-associative, and intersection binds tighter than union. Parentheses
override precedence and should remain observable through source spans or an
explicit group node.

The current `Token::Operator(char)` cannot represent multi-character
comparisons or whether `+`/`-` is prefix or binary. A new AST operator enum
must preserve that distinction. A colon inside an already parsed bracketed
reference is address structure; a colon outside it is the infix range operator.
`!` and `~` are absent from the current tokenizer, while `%` is absent and
`|` is needed by arrays.

## Functions, names, labels, references, and errors

Section 5.6 defines a function name as an XML `LetterXML` followed by zero or
more XML letters, digits, underscore, dot, or `CombiningCharXML`; names are
case-insensitive. A call always has parentheses, but its parameter list may be
empty. The parameter grammar is:

```text
ParameterList ::= /* empty */ |
    Parameter (Separator EmptyOrParameter )* |
    Separator EmptyOrParameter (Separator EmptyOrParameter )*
EmptyOrParameter ::= /* empty */ Whitespace* | Parameter
Parameter ::= Expression
Separator ::= ';'
```

`F()` means zero parameters. Leading, trailing, and interior missing slots in
forms such as `F(;A)`, `F(A;)`, and `F(A;;B)` must remain explicit in the AST;
dropping semicolons changes arity. Commas are not the standard separator.
Section 5.7 allows nonstandard/host-defined functions. Catalog membership and
arity are semantic checks, so a syntax parser must retain an unknown function
name rather than reject it merely because it is absent from the 393-name
standard catalog. `is_valid_function` should remain a separate catalog query.

Section 5.8's bracketed references are already owned by the bounded
`reference::Reference` model. Section 5.9 defines a reference list as
`Reference (Whitespace* '~' Whitespace* Reference)*`; the AST must retain the
union operator and must not flatten the list or resolve it.

Section 5.10 adds:

```text
QuotedLabel ::= SingleQuoted
AutomaticIntersection ::= QuotedLabel Whitespace* '!!' Whitespace* QuotedLabel
```

Labels and automatic intersections are inert text until a host performs its
defined-label or automatic-label lookup. Section 5.2's shared lexical rule is
`SingleQuoted ::= "'" ([^'] | "''")+ "'"`; the `+` forbids an empty quoted
label or quoted sheet name.

Section 5.11 defines named expressions:

```text
NamedExpression ::= SimpleNamedExpression |
                    SheetLocalNamedExpression | ExternalNamedExpression
SimpleNamedExpression ::= Identifier | '$$' (Identifier | SingleQuoted)
SheetLocalNamedExpression ::= QuotedSheetName '.' SimpleNamedExpression
ExternalNamedExpression ::= Source
    (SimpleNamedExpression | SheetLocalNamedExpression)
```

`Identifier` is an XML-letter-led sequence of XML letters/digits,
underscore/combining characters, excluding cell-shaped `[A-Za-z]+[0-9]+`
spellings and case-insensitive `TRUE`/`FALSE`. Name lookup and dependency
resolution are outside this parser; existing name-dependency scanners should
eventually consume the structured AST rather than rescan raw text.

Section 5.12 defines constant errors as:

```text
Error ::= '#' [A-Z0-9]+ ([!?] | ('/' ([A-Z] | ([0-9] [!?]))))
```

The specific spelling must be retained and cannot contain whitespace. `#REF!`
outside brackets is a constant error; inside a bracket it remains the
invalidated-reference form.

## Inline arrays

Section 5.13 defines:

```text
Array ::= '{' MatrixRow ( '|' MatrixRow )* '}'
MatrixRow ::= Expression ( ';' Expression )*
RowSeparator ::= '|'
```

The grammar requires at least one row and at least one expression per row;
semicolons separate columns in this context. It does not require all rows to
have equal lengths. The following prose and §7.2 describe evaluator capability:
an evaluator claiming inline constant-array support must accept rectangular
matrices with one or more rows/columns and at least Number (optionally unary
`-`), Text, TRUE/FALSE, and Error elements. §7.3 separately requires full
Expression syntax for inline nonconstant arrays. A syntax-only AST therefore
must preserve `Vec<Vec<Expr>>` row structure, including ragged rows, and expose
rectangularity/constantness as a capability or validation result. Rejecting a
ragged row at grammar admission would impose an artificial subset.

## Whitespace and compatibility boundary

Section 5.14 defines whitespace as exactly SPACE, TAB, LF, and CR. It is
ignored around expressions, operators, constants, arrays, parentheses, and
function closing parentheses, and before a function name. It must not separate
a function name from its opening parenthesis, and it is forbidden inside a
terminating lexical rule unless that rule permits it. Whitespace inside double
or single quotes is content. The current tokenizer's broad ASCII-whitespace
skip and acceptance of `SUM (` are compatibility behavior; a strict grammar
path must distinguish that from normative conformance. Formula XML escaping is
an embedding concern from §5.1 and should not alter retained decoded source
text.

## API and finite-resource constraints

The existing public `Formula { text, tokens }` and `Token` enum are used by
legacy callers. Adding fields or variants would break public struct literals
and exhaustive matches, and the flat token shape cannot express precedence,
missing arguments, operator context, labels, names, or array rows. The bounded
compatibility design is a separate `ParsedExpression`/`FormulaAst` result with
inert expression nodes for literals, references, names, labels, errors,
function calls, unary/postfix/binary operators, explicit groups, and arrays;
it should retain exact source text and forced-recalculation metadata. Existing
tokenization and legacy projections can remain available independently.

ODF §3.7's basic-limit floor is 1024 interchange characters, 30 list
parameters, 32,767 ASCII string characters, and seven function nesting levels.
The project already has finite 1 MiB formula-byte and 65,536-token defaults;
these are larger ceilings, not reasons to omit the standard minimums. The AST
path needs explicit checked limits for nodes/parameters, nesting depth, array
rows/columns/elements, and owned identifier/label/error text, with checked
row×column arithmetic. It must admit before fallible vector/string growth,
charge the existing `InputBytes`/`Objects`/`Depth`/`Work` dimensions where a
budget context is supplied, and avoid recursive destruction or an unbounded
call stack. Resource failures remain typed; no name resolution, function
execution, workbook access, or external I/O is part of this feature.

## Implementation review status

The frozen implementation was reviewed against the requirements above. The
strict parser consumes one complete root through EOF, retains prefix/postfix
and multi-character comparison operators with the §5.5 precedence and
associativity, preserves explicit missing function-parameter slots and ragged
array rows, and retains source spans for grouping and lexical content. The
focused receipt in `gates/focused-test.log` covers these paths,
including allocator-injection and hard-depth subprocess checks.

The tree is flat and append-only: node and edge counts are checked before
growth, child ranges are checked when exposed, reference IDs are separately
bounded, array cells are charged before parsing each cell, and parser-owned
source, vectors, temporary child lists, and decoded source components use
fallible reservations. Recursive descent is capped at 256, while long
left-associative chains remain iterative. A failed parse drops the staging
arena and cannot publish a partial expression. Bracketed references and named
expression sources are parsed and validated lexically; they are retained as
inert metadata and never resolved, opened, or executed.

Reviewed implementation digests are:

```text
crates/litchi-ods/src/codec/formula.rs                       e5681591c4c1c196e21141eadee7fd046bbb668b6defc8ab5ced1706854ee072
crates/litchi-ods/src/codec/formula/expression.rs             b75b78b75b0b1dbf660f736eadddee335c18c2121fd0bc78198c7eca1c58ef04
crates/litchi-ods/src/codec/formula/expression/names.rs      79631c602b587391baa9efe695da6e6c480434a12a7ab826e32d4076c9fb53b2
crates/litchi-ods/src/codec/formula/reference.rs             163e5131e0f787059641df30fda0b0cb2b7582fdf414f6263d1ebfbb80a47c40
crates/litchi-ods/src/codec/formula/reference/iri.rs         194bbad30c601938e27565b408e01509e24534364f5e614b1e5ae2b6eb51d306
crates/litchi-ods/tests/ods_formula_expressions.rs          2e8580024a04e6f740d8b7b38c211394268d54f31d40135a141e37645c8ae2ba
```

Section 5.14's first sentence makes whitespace globally ignorable except
inside string and single-quoted content. Its later clauses forbid whitespace
inside terminating lexical rules and explicitly forbid separating a function
name from its opening parenthesis; they do not restrict whitespace at the
boundaries of named-expression or reference subrules. Consequently all of
`'Sheet' . Name`, `$$ Name`, `$ 'Sheet' . Name`, and `'source'# Name` should
parse. The expression parser now admits those boundaries, including spaces
after `.` and after the `$$` or absolute `$` component. `S UM()`, `SUM (1)`,
`$ $Name`, and whitespace inside an unquoted identifier remain invalid because
those split a terminating token or violate the explicit function-call rule.
Section 5.11 also discusses an empty quoted sheet as a lookup starting point,
despite the shared `SingleQuoted` `+` production; the implementation's
acceptance of `''` for that component follows the prose interpretation of this
inconsistency.

The same boundary rule applies to the delegated §5.8 reference grammar. For
example, `[ . $A$1 ]`, `[Sheet . A 1]`, and `[. A 1 : . B 2]` are syntactically
valid with whitespace between nonterminal components, while whitespace inside
`$A`, `$1`, or an unquoted `SheetName` is not. The reference parser now applies
the same exact four-character boundary skip, including source `#`, endpoint
`.`, sheet `.`, and column/row boundaries, while retaining lexical terminals
as contiguous. Bytes outside that four-character set are not silently treated
as formula whitespace; XML embedding character validation remains a separate
layer. The added expression and reference regressions cover the valid and
refusal cases. No source edit was made in this review.

Disposition: no boundedness, arena, precedence, array-shape, UTF-8/string,
reference-inertness, allocation, or Chapter 5 whitespace blocker was found.

# 0005 — M0b: path/selector grammar in `litchi-core::selector`

- Milestone: M0b (design §2, §15)
- Status: implemented on branch `uo/0005-selector-grammar`; awaiting coordinator review
- Scope: format-agnostic **syntax only**. Resolution against documents
  (aliases, case-insensitive matching, header-column binding, `path_not_found`)
  is M1 work.

## What changed

`litchi-core::selector` now holds two layers:

1. The existing typed vocabulary, unchanged: `Position` (zero-based) and
   `Selector<'a, Id>::{Name, Position, Id}`. No caller needed changes
   (`litchi-xlsx`, `litchi-odt` and others were re-checked). `Position` gained
   `from_one_based` / `one_based`, the single named boundary between the
   1-based string grammar and the zero-based typed value.
2. The new string grammar as typed ASTs:
   - `Path<'a>` / `Step<'a>` for element addressing (`/body/p[3]/r[1]`).
   - `Predicate<'a>::{Index(Position), Last, Filter(Expr)}`,
     `Expr::{Condition, And, Or, Not}`,
     `Condition::{Exists(Key), Compare { key, op, value }}`, `Key` (with the
     `@` attribute flag), `CompareOp`, `Value::{Text, Regex}`.
   - `Query<'a>` (optional `Scope::{Path, Sheet}` + comma alternatives),
     `Complex` (compounds joined by `Combinator::{Child, Descendant}`),
     `Compound` (`Element::{Implicit, Any, Named}` + predicates + `Pseudo`s),
     `Pseudo::{Contains, Empty, NoAlt, Has, Not}`.
   - `Axis` (story/part classification of a step), `CellRef`, `CellRange`,
     `column_number` (syntactic A1 helpers).
   - `ParseError` (`Copy`, 24 bytes): `grammar()`, byte `offset()`,
     `kind()` (`ErrorKind` incl. `Unexpected { expected: Expected, found: Found }`),
     `code()` (`ErrorCode::{InvalidPath, InvalidSelector, MaxDepthExceeded}`,
     plus the reserved resolution-time `PathNotFound`), and `suggestion()`.
   - `selector::limits` constants.

Strings are an edge layer: `Path::parse(&str)` / `Path::parse_relative` /
`Query::parse(&str)` borrow names and values (`Cow::Borrowed`; escapes are the
only reason to own), `FromStr` yields owned `'static` values, and `into_owned`
converts. `Display` renders a **canonical** form; `parse(display(x)) == x` for
every parsed value, and display is a fixed point. All AST types are covariant
in their lifetime (no `SmallVec` projection invariance), `Send + Sync`, and
immutable after parse, so a parsed query is reusable across nodes and threads.

Files:

- `crates/litchi-core/src/selector.rs` (module root, `Position` helpers)
- `crates/litchi-core/src/selector/{cursor,path,predicate,query,axis,error,limits,text,store}.rs`
- `crates/litchi-core/src/selector/grammar_tests.rs` (unit tests per feature)
- `crates/litchi-core/tests/selector_corpus.rs` + `tests/fixtures/officecli_{paths,selectors}.txt`
- `crates/litchi-core/tests/selector_roundtrip.rs` (seeded property tests, mutation robustness, hostile inputs)
- `crates/litchi-core/tests/selector_allocations.rs` (counting allocator)
- `crates/litchi-core/examples/selector_parse_bench.rs`
- `crates/litchi-core/fuzz/fuzz_targets/selector_grammar.rs` + seed corpus
  `crates/litchi-core/fuzz/corpus/selector_grammar/` (existing litchi-core fuzz
  crate; README section added)

## Grammar (EBNF)

Keywords (`and`, `or`, `not`, `last()`, pseudo-class names) match ASCII
case-insensitively and display in lower case. `ws` is ASCII whitespace.

```ebnf
(* ---------- paths ---------- *)
path          = "/" | "/" step { "/" step } ;
relative-path = step { "/" step } ;                 (* Path::parse_relative *)
step          = name { predicate } ;
name          = quoted | bare-name ;
bare-name     = name-char { name-char } ;           (* first char is not a quote *)
name-char     = ? any char except "/", "[", "]", ASCII control ? ;

quoted        = '"' { qchar } '"' | "'" { qchar } "'" ;
qchar         = "\" any-char | ? any char except the delimiter and "\" ? ;
                (* \<delimiter> and \\ drop the backslash; any other \x is literal *)

predicate     = "[" ws ( index | "last()" | or-expr ) ws "]" ;
index         = digit { digit } ;                   (* 1-based: 1 ..= usize::MAX *)

or-expr       = and-expr { "or" and-expr } ;        (* keywords need ws or "(" after *)
and-expr      = factor { "and" factor } ;
factor        = "not" ws "(" or-expr ")" | "(" or-expr ")" | condition ;
condition     = key [ ws op ws value ] ;            (* key alone = exists / bare token *)
key           = [ "@" ] ( quoted | key-char { key-char } ) ;
key-char      = ? Unicode alphanumeric ? | "." | "_" ;
op            = ">=" | "<=" | "!=" | "\!=" | "~=" | "=" | ">" | "<" ;
value         = quoted | regex | bare-value ;
regex         = "r" quoted ;                        (* raw: backslashes kept; only with ~= *)
bare-value    = ? chars up to "]" or ")" or ws-before-"and"/"or"; no quotes, "[",
                  ASCII control or "*"; must not start with "=", "~", "!";
                  trailing ws trimmed ? ;

(* ---------- queries ---------- *)
query         = ws [ scope ] selector-list ws ;
scope         = path-scope "/" | sheet "!" ;
path-scope    = [ "/" step { "/" step } ] ;         (* text before the last top-level "/" *)
sheet         = "'" { ? char except "'" ? | "''" } "'"
              | sheet-char { sheet-char } ;         (* "!" must not be followed by "=" *)
sheet-char    = ? any char except "!", "[", "(", ":", ",", ">", "/", quotes ? ;

selector-list = complex { ws "," ws complex } ;
complex       = compound { combinator compound } ;
combinator    = ws ">" ws | ws1 ;                   (* "+" and "~" are rejected *)
compound      = element { predicate | pseudo }
              | ( predicate | pseudo ) { predicate | pseudo } ;
element       = "*" | cell-range | ident ;
ident         = ( ALPHA | DIGIT | "_" ) { ALPHA | DIGIT | "_" | "-" } ;
cell-range    = cell [ ":" cell ] ;                 cell = 1*3ALPHA 1*7DIGIT ;
pseudo        = ":contains(" ws ( quoted | bare-text ) ws ")"
              | ":has(" ws selector-list ws ")"
              | ":not(" ws selector-list ws ")"
              | ":empty" | ":no-alt"
              | ":" shorthand ;                     (* = :contains(shorthand) *)
bare-text     = ? chars up to ")"; no quotes, "[", "]", "(", ASCII control ? ;
shorthand     = ? chars up to ws , > ( ) [ ] : \ quote or end ? ;
```

Notes:

- A query whose text starts with `/` is split at the last `/` that is outside
  `[...]`, `(...)` and quoted spans: the prefix is a `Path` scope, the rest a
  selector list (`/Assumptions/cell[type=Number]`, `/slide[1]/shape:contains("a/b")`).
  `/body` is the root scope plus `body`.
- Stacked brackets `[a][b]` are separate predicates (AND); inside one bracket
  `and` binds tighter than `or`.
- Bare and quoted spellings of the same text are the same AST value; display
  uses the bare form only when it reads back identically, otherwise `"..."`
  with `\"`/`\\` escapes (sheet scopes use Excel `'...'` with `''`).
- An all-digit bare token is always an index; an all-digit key displays quoted.

## Story axes

`Step::axis()` / `Path::story()` / `Path::axes()` classify step names
(case-insensitive, allocation-free sorted table):

| Axis | Accepted names | OfficeCLI forms harvested |
|---|---|---|
| Body | `body` | `/body/p[N]` |
| Header / Footer | `header`, `footer` | `/header[N]/p[M]/r[K]`, `/header[@type=default]` |
| Footnote / Endnote | `footnote(s)`, `endnote(s)` | `/footnote[@footnoteId=N]`, `/footnotes/footnote[N]`, `/endnote[N]` |
| Comment | `comment(s)`, `modernComment` | `/comments/comment[N]`, `/slide[N]/modernComment[K]/reply[R]` |
| Notes | `notes`, `noteSlide` | `/slide[N]/notes`, `/noteSlide[N]` |
| NotesMaster | `notesMaster` | `/notesMaster` |
| Slide / SlideMaster / SlideLayout | `slide`, `slideMaster`, `slideLayout`, `layout` | `/slidemaster[N]/slidelayout[M]/shape[K]`, design's `/slideMaster[1]/layout[2]` |
| Chart | `chart` | `/chart[N]/series[K]`, `/Sheet1/chart[N]/axis[@role=value]`, `/slide[N]/chart[N]` |
| Sheet | `sheet` | `/sheet[N]/A1`, `/sheet[last()]` |
| Parts | `styles`, `numbering`, `theme`, `settings`, `watermark`, `customXml` | `/styles/Heading1`, `/numbering/abstractNum[@id=N]/level[L]` |

Sheets addressed by name (`/Sheet1/...`) are not axes; their names are free
text and only resolution knows they are sheets.

## Supported vs OfficeCLI deltas

OfficeCLI's grammar is not one parser: the Word path navigator, the
`AttributeFilter` expression parser, and per-handler regex selectors each
accept slightly different languages. This grammar is one recursive-descent
parser covering their union, with the deltas below.

| # | Area | OfficeCLI | Litchi | Disposition |
|---|---|---|---|---|
| 1 | Stacked path predicates | Word path parser rejects `p[1][2]`; schemas use `placeholder[@type=body][@idx=K]`, `revision[@id=N][@type=T]` | Accepted everywhere (AND) | superset |
| 2 | Whitespace in `[ 1 ]` | Word path rejects; filter parser trims | Accepted, displays `[1]` | superset |
| 3 | `@attr` value quoting in paths | bare or `"..."` without embedded `"` | bare, `"..."` or `'...'` with backslash escapes | superset |
| 4 | Attribute / key names | path: `[A-Za-z_][A-Za-z0-9_]*`; filter: letters/digits/`._` or quoted | Unicode alphanumeric/`._` or quoted; `@w:paraId` rejected in both | superset |
| 5 | Bare tokens in path brackets (`col[B]`, `formfield[x]`, `placeholder[ctrTitle]`, `namedrange[NAME]`) | per-handler special cases; Word path rejects `p[XYZ]` | Any bare key parses (`Predicate::bare_key`); resolution decides meaning | deferred to M1 resolution |
| 6 | Element aliases (`paragraph`→`p`), case-insensitive types | applied while parsing | names kept verbatim; aliasing/case folding at resolution | deferred to M1 |
| 7 | `[0]`, `[-1]`, overflow | rejected (message) | rejected: `IndexZero` / `IndexNegative` / `IndexOverflow` | same, typed |
| 8 | `last()` | paths and selectors | same (`Predicate::Last`) | same |
| 9 | Quote inside a bare value (`[name=O'Brien]`) | unbalanced-quote error | `QuoteInBareValue` | same, typed |
| 10 | `*` in values | rejected even when quoted | rejected only when bare; `"x*"` is a literal | wontfix (quoted literal is unambiguous) |
| 11 | Regex `r"..."` | only with `~=` | same (`RegexRequiresContains`); stored, not compiled | same |
| 12 | zsh `\!=` | accepted | accepted, displays `!=` | same |
| 13 | `:contains()` argument | Word: up to first `)`; Excel: lazy regex | quoted, or bare up to `)` without quotes/brackets/`(` | same for harvested forms |
| 14 | Shorthand `cell:text` | rest of selector | text up to ws, `,`, `>`, brackets, parens, `:`, `\`, quote; normalizes to `:contains(text)` | narrower; quote with `:contains("...")` |
| 15 | `:has()` | only `:has(formula)` (Excel) | any selector list | superset; resolution restricts |
| 16 | `:not()` | only `:not(:empty)` / `:not(:has(formula))` by substring | any selector list | superset |
| 17 | Combinators | `>` one level (Word), ancestor advisory (PPT); descendant only after `slide`; `+`/`~` rejected | arbitrary `>` / descendant chains; `+`/`~` → `UnsupportedCombinator` | superset syntax; semantics at M1 |
| 18 | Union `,` | top-level, quote/bracket/paren aware | same | same |
| 19 | `Sheet!` scope | top-level `!` not followed by `=`, Excel `'...'` quoting | same; display canonicalizes `My Sheet!` to `'My Sheet'!` | same |
| 20 | Path-scoped query `/Sheet1/cell[...]` | Excel: first segment is the sheet; PPT strips the slide prefix | last top-level `/` splits a `Path` scope from the selector; scope names with unbalanced parentheses must be quoted | generalized |
| 21 | Segment quoting | none (split on `/`) | `/"a/b"/A1` for names with `/ [ ]` or a leading quote | superset |
| 22 | `//x`, `x/`, `/a//b` | rejected | `EmptySegment` | same, typed |
| 23 | `Sheet1!A1` as a mutation target | accepted by set/remove | parses as a `Query` (sheet scope + `A1`); `Path::parse` requires `/` | facade converts at M1/M2 |
| 24 | `bare_selector_rejected` guard | CLI verb layer | not grammar; `Query::scope()` and `Path::is_absolute()` expose what the guard needs | deferred to M2 verbs |
| 25 | `find:` anchors, `selected` watch path | `--after/--before find:...`, `get selected` | not part of the grammar | deferred (anchor mini-syntax / resident server, design §11) |
| 26 | Content-filter vs locator paths | `IsContentFilterPath` | `Predicate::is_structural`, `Path::is_structural` with the same rule | same |
| 27 | `col.`/`column.` header prefix | stripped at match time | kept in the key text (`col.Salary`) | resolution strips (M1) |
| 28 | Whole-column/row ranges `A:C`, `1:5`, `$A$1` | no usage harvested | not recognized as cell ranges (`A1:B` reads `B` as `:contains` shorthand in a query) | deferred |
| 29 | Separator whitespace | `char.IsWhiteSpace` (Unicode) | ASCII whitespace only; Unicode spaces are ordinary characters | wontfix |
| 30 | Errors | `CliException` message + `invalid_selector`; path errors are untyped `ArgumentException` | `ParseError { grammar, offset, kind, code, suggestion }`; `invalid_path` vs `invalid_selector`; limits → `max_depth_exceeded` | improvement |

## Limits (ADR 0006)

All in `litchi_core::selector::limits`; exceeding one returns
`ErrorKind::LimitExceeded(Limit)`. Recursion happens only through
parentheses, `not(...)`, `:has(...)`, and `:not(...)`, all of which share the
nesting counter, so stack depth is bounded regardless of input.

| Constant | Value | Code |
|---|---|---|
| `MAX_INPUT_BYTES` | 16 KiB | `invalid_path` / `invalid_selector` |
| `MAX_STEPS` | 64 | `max_depth_exceeded` |
| `MAX_PREDICATES` (brackets + pseudos per step/compound) | 16 | `invalid_*` |
| `MAX_NESTING` | 32 | `max_depth_exceeded` |
| `MAX_TERMS` (conditions per bracket) | 64 | `invalid_*` |
| `MAX_ALTERNATIVES` | 64 | `invalid_*` |
| `MAX_COMPOUNDS` | 32 | `invalid_*` |

`hostile_inputs_are_bounded` feeds 100k open parentheses, 100k `:not(`
nestings, 10k-step paths, a 1 MiB name, and multi-byte overflows; all fail
typed, and oversize errors report a char-boundary offset.

## Performance evidence

Allocation (counting global allocator, `tests/selector_allocations.rs`):

- **Zero allocations** on the success path for 19 path forms covering word,
  slide, sheet, story-axis, stable-identity, and 6-step table-cell paths
  (`/body/tbl[1]/tr[2]/tc[3]/p[1]/r[1]`), and for 17 simple queries
  (`paragraph[style=Heading1]`, `row[Salary>5000]`, `Sheet1!B[value!=0]`,
  `'My Data (2024)'!row[Amount>5000]`, `cell:contains("TODO")`, `p:empty`, ...).
- **At most one allocation** for `and`/`or` expressions, combinator chains,
  unions, `:has(...)`/`:not(...)`, and path-scoped queries (boxed scope).
- Rejections, `Axis::classify`, and `CellRange::parse` allocate nothing.

How: names/values are `Cow::Borrowed`; a path keeps up to 6 steps in a
covariant inline array (`store::Inline`), and single predicates /
pseudo-classes / alternatives sit in `store::OneOrMany` (no `unsafe`, no new
dependency, no regex). Sizes: `Path` 760 B, `Query` 248 B, `ParseError` 24 B.

Throughput (`examples/selector_parse_bench.rs`, release, 5 consecutive runs on
one AMD EPYC 9R45 host; raw output in
`docs/unified-ops/results/0005/selector_parse_bench.txt`):

| Input | ns/parse (min–max of 5 runs) |
|---|---|
| 950 harvested paths (mean) | 56.6–57.8 |
| 132 harvested selectors (mean) | 114.4–130.0 |
| `/body/p[3]/r[1]` | 56.1–59.2 |
| `/slide[1]/shape[@name=Title]` | 78.6–85.9 |
| `/Sheet1/A1:C5` | 41.3–41.8 |
| `row[Salary>5000 and Region=EMEA]` | 140.9–157.8 |

These are indications for this machine, not budgets; there is no prior
implementation to compare against.

## Tests and gates

- Unit tests per grammar feature (26) in `selector/grammar_tests.rs`: paths,
  root/relative, `last()`, identities, filters, bare keys, spreadsheet forms,
  story axes, path errors with exact offsets/expected tokens, limits, borrowing,
  spill, covariance, position conversion, `Send + Sync` and size bounds; selector types, operators, regex,
  booleans and precedence, pseudo-classes, combinators/unions, scopes,
  selector errors, limits, builders, unchanged vocabulary.
- Corpus: 959-line path fixture (950 paths) and 138-line selector fixture
  (132 selectors) harvested from OfficeCLI examples, README/SKILL docs, skills,
  and `schemas/help/**/*.json` (templates made concrete), plus forms quoted in
  OfficeCLI help text and source comments; every entry parses, round-trips,
  and agrees with owned parsing.
- Round-trip properties (no `proptest` in the workspace, so a seeded
  SplitMix64 generator): 20k random path ASTs and 20k random query ASTs over
  adversarial names/keys/values satisfy `parse(display(x)) == x` and display
  fixed points; 60k mutated corpus strings either fail typed (offset on a char
  boundary) or round-trip.
- Fuzz target `selector_grammar` added to the existing
  `crates/litchi-core/fuzz` crate (same pattern as `detect_format`), with eight
  reviewable seeds. It type-checks; it was **not run** under libFuzzer here
  (no nightly toolchain or `cargo-fuzz` on this host). The in-tree mutation
  test exercises the same invariants on stable.

Commands run (all green):

```sh
cargo fmt --check -p litchi-core                 # workspace-wide check also flags pre-existing
                                                 # diffs in litchi-iwa-archive and soapberry-zip
cargo clippy -p litchi-core --all-features --all-targets -- -D warnings
cargo test -p litchi-core                        # 219 lib + 8 + 4 + 3 + 4 tests, 55 doctests
RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi-core --no-deps
cargo check -p litchi-xlsx -p litchi-odt         # existing Selector/Position callers
cargo check --manifest-path crates/litchi-core/fuzz/Cargo.toml
```

## What is not claimed

- No resolution: nothing here checks that `/body/p[3]` exists, binds
  `row[Salary>5000]` to a table header, applies aliases or case folding, or
  evaluates predicates. `ErrorCode::PathNotFound` is reserved, never produced.
- No predicate evaluation engine and no regex compilation; `Value::Regex` is
  stored text. Numeric/unit comparison semantics (`24pt`, `10cm`) are M1.
- No claim that every OfficeCLI string is accepted: the corpus covers the
  harvested strings; deltas above are the known differences.
- The fuzz target has not been run under libFuzzer.
- Throughput numbers are single-host indications without a baseline; no
  claim about end-to-end `get`/`query` cost.
- `Display` round-trip is guaranteed for parsed values and for builder values
  within the documented invariants (`And`/`Or` with at least two operands,
  absolute scope paths, representable regex patterns); arbitrary hand-built
  element/pseudo names outside the token grammar are not re-validated.

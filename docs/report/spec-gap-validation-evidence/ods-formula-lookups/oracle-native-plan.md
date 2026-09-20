# ODF lookup-function oracle and native-evidence plan

Status: retained evidence-design plan. The contract-bound oracle and native
capture are now retained in this directory; see [oracle-review.md](oracle-review.md),
[oracle-owner-correction.md](oracle-owner-correction.md), and
[native/README.md](native/README.md) for their actual scope, results, and limits.
The proposed work below records the design intent and is not a passing receipt.

## Normative inputs

The independent model in `lookup_oracle.py` will use only the repository-local ODF 1.4 Part 4
material:

| Input | Identity |
| --- | --- |
| package | `3rdparty/specs/OpenDocument-v1.4-os.zip` |
| package SHA-256 | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| formula member | `part4-formula/OpenDocument-v1.4-os-part4-formula.html` |
| member SHA-256 | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |
| sections | 6.14.2, 6.14.3, 6.14.5–6.14.9, 6.14.11, 6.14.12 |
| conversion sections | 6.3.2, 6.3.3, 6.3.4, 6.3.5, 6.3.6, 6.3.14 |

The final evidence bundle must also pin the lookup contract SHA-256. Until that
file exists and the semantic reviewer resolves the questions below, no native
result is promoted to evidence.

## Independent oracle

The planned `lookup_oracle.py` will own its Python value and reference types;
it will not import the Rust evaluator or copy its lookup algorithm. Its output
will be `lookup-goldens.json` with this bounded shape:

```text
schema, contract_sha256, normative_archive_sha256, normative_member_sha256,
functions, workbook, profile, observations[]
```

Each observation will retain `case`, exact `formula`, function, scalar or
matrix mode, explicit formula position, expected typed result, expected cell
reads, and any returned reference descriptor. A descriptor will retain source
identity, ordered records, sheet span, two-dimensional bounds, and list/record
boundaries. `INDEX` area selection, `INDIRECT`, `OFFSET`, and selected `CHOOSE`
references must remain distinguishable from a materialized cell Array. A
ReferenceList will retain order and duplicates rather than being flattened.

The resolver-free portion will cover ADDRESS, CHOOSE scalar values, text and
inline arrays, invalid indexes, and all known shape/type refusals. A bounded
synthetic resolver will supply lookup tables and record reads for reference
and ForceArray cases. It will log reads in order and expose finite sheet order,
extents, and cell contents. The model will compare read counts and selected
coordinates, not merely the final scalar.

The proposed local workbook profile uses four ordered sheets (`Main`, `Data`,
`Hidden`, `Archive`) and keeps `Hidden` in resolver order/count. The lookup
tables will include:

- ascending numeric keys with duplicate runs and a key below the query;
- descending numeric keys with duplicate runs;
- case variants such as `Alpha`/`alpha`;
- mixed Number/Text/Logical values to exercise the specified type ordering and
  Number-versus-Text mismatch rules;
- empty cells and formula Error cells after the searched portion and in the
  returned portion; and
- equal-shape local tables on multiple sheets plus ordered ReferenceList and
  3-D reference spellings.

The first oracle corpus will be organized as follows. `lookup_oracle.py
--self-check` runs the model without writing JSON; `--write` and `--check` bind
the corpus to `contract.md`. The current rows include all nine function names,
scalar and matrix CHOOSE values and descriptors, inline and reference lookup
arrays, exact/ascending/descending duplicate ties, Empty search-versus-result
behavior, formula-error suffix scanning, scalar/list/shape refusals, bounded
LOOKUP extension, ADDRESS A1/R1C1 formatting and quoting, descriptor-only
INDIRECT/OFFSET, relative R1C1 interpretation for INDIRECT, and pinned text
observations for Straße/STRASSE, sigma variants, and dotted-I no-normalization.

| Function | Planned independent cases |
| --- | --- |
| ADDRESS | default `Abs=1`, all four absolute modes, optional sheet text including spaces and apostrophes, A1 and R1C1 output, current-position-relative R1C1, integer/bounds refusals, omitted versus explicit empty optional slots |
| CHOOSE | first/middle/last index, zero/negative/out-of-range and fractional index policy, logical/text index conversion, selected value/reference/error, and proof that unselected branches do no reads or provider work |
| HLOOKUP | exact `FALSE/0`, approximate omitted/TRUE/nonzero, duplicate last-match tie, ascending mixed-type order, row bounds, array versus reference data, 3-D/list shape, searched errors, and returned error/read precedence |
| INDEX | Array and ReferenceList inputs, direct 3-D record, row/column omission or zero returning full row/column/area, `AreaNumber` record selection, duplicate records, bounds, returned reference identity, and selected cell reads |
| INDIRECT | A1 local cell/range, dotted and exclamation sheet separators, quoted sheet names, R1C1 absolute/relative text, A1 flag defaults, current position, invalid text, source/external text, and no read until a consuming function evaluates the returned Reference |
| LOOKUP | two-parameter orientation, three-parameter result vector, approximate largest-less-than-or-equal rule, duplicate last tie, case-insensitive text, mixed type ordering, type mismatch, result length/extension behavior, and unsorted/error cases |
| MATCH | omitted/`1`, `0`, and `-1`, exact first tie, approximate last tie, case-insensitive text, mixed type ordering, Number/Text mismatch, vector shape, fractional/invalid match type policy, and error/read precedence |
| OFFSET | row/column offsets, default and explicit height/width, empty height slot, 3-D/reference-list identity, bounds and overflow, source refusal, and zero reads until a downstream consumer selects cells |
| VLOOKUP | exact `FALSE/0`, approximate omitted/TRUE/nonzero, duplicate last-match tie, ascending mixed-type order, column bounds, array versus reference data, 3-D/list shape, searched errors, and returned error/read precedence |

Approximate lookup rows will state the sorted-input assumption explicitly. The
oracle will not silently turn an unsorted input into a normative result. Native
observations for unsorted inputs will be retained as host behavior only.

## Reference and conversion seams to test

The contract must decide how each `Reference|Array` or `ReferenceList|Array`
parameter treats a direct Reference, a 3-D cuboid, an ordered list, and a
computed reference. The oracle will retain these distinctions through every
consumer. It will test that a direct `OFFSET` or `INDIRECT` descriptor can be
passed to `INDEX`/lookup without an implicit scalar intersection, while a
consumer that needs one cell performs the documented intersection and charges
the resulting read.

The conversion cases will make the selected profile visible:

- Integer parameters: Number, Logical, Text, fractional Number, formula Error,
  Empty, and missing argument, with the function-specific rounding/refusal
  decision recorded in the contract.
- Logical `RangeLookup` and A1 flags: TRUE/FALSE, zero/nonzero Number, Text,
  Empty, and formula Error, including whether omitted and explicit empty slots
  select the same default.
- Text lookup and address names: case-folding for comparisons, exact sheet-name
  matching, quoting/escaping, and the treatment of Number/Logical values where
  a Text parameter is required.
- Lookup ordering: Numbers before Text before Logical for mixed approximate
  data, case-insensitive Text ordering, and the separate Number-versus-Text
  mismatch rule when a candidate is selected.

## Lazy evaluation and resource receipts

The semantic rows will include selected and unselected `CHOOSE`, `IF`,
`IFERROR`, and `IFNA` branches containing references, formula Errors, and
provider/source failures. The expected result records whether the unselected
branch was untouched. For lookups and INDEX, the resolver log will distinguish
searched-cell reads from the returned-cell read and will cover the first/last
duplicate tie without relying on a host's hidden cache.

The eventual native and focused receipts should prove:

- no reads for ADDRESS, reference construction by INDIRECT/OFFSET, rejected
  shapes, invalid indexes, or unselected branches;
- bounded ordered reads for exact and approximate searches;
- formula Error values remain values, while typed provider, cancellation,
  resource, and source failures supersede them;
- checked dimensions and `max_reference_cells`/array limits precede materialized
  storage; and
- source-version and cancellation fences surround the complete evaluation and
  publication.

## Native fixture, after contract freeze

`native/` will retain a locally authored FODS input, fresh-profile recalculated
ODS, extracted `content.xml`, typed `native-results.json`, `provenance.json`,
and a small `reproduce.py`. The reproduction will invoke `/usr/bin/libreoffice`
headlessly with a new temporary profile and `C.UTF-8`, verify input/output/
content/result hashes and formula order, and remove its temporary profile and
output tree on every exit path. No network or downloaded fixture is allowed.

The fixture will keep native formulas bounded and inspectable: one sheet for
the lookup tables, one for 3-D/list and reference-construction cases, and one
hidden sheet for order/count behavior. Unsupported host syntax or host error
tokens will remain verbatim in the result receipt and be listed as divergences
from the normative profile. Native output will never be used to fill an
independent expected value.

## Semantic-review questions before capture

These decisions are intentionally open and should be answered in the lookup
contract or semantic review before `lookup_oracle.py` and the native fixture
are generated:

1. Which of HLOOKUP, INDEX, LOOKUP, MATCH, and VLOOKUP admit an ordered
   ReferenceList, and when does a one-record list remain a list rather than a
   direct Reference? Does `AreaNumber` count retained records or physical
   planes of a 3-D record?
2. For lookup DataSource values that are 3-D References or lists, is the
   searched plane the first plane, a flattened logical sequence, or a typed
   refusal? What shape does a returned 3-D reference preserve?
3. Do ADDRESS, INDIRECT, and OFFSET preserve reference identity and source
   metadata through nesting, and which consumers are allowed to resolve a
   source-qualified descriptor in this limited resolver profile?
4. What exact integer conversion applies to fractional row/column/index,
   AreaNumber, and MatchType parameters? How do Logical, Text, Empty, formula
   Error, missing, and explicit empty slots differ?
5. For ADDRESS, what are the exact default and refusal rules for `Abs`, A1,
   optional Sheet, row/column zero or negative values, quoting, apostrophe
   escaping, and R1C1 relative coordinates at a specified formula position?
6. For INDIRECT, must both `.` and `!` sheet separators be accepted in each
   A1/R1C1 mode, and what happens for external/source text, whole rows/columns,
   3-D text, and malformed or missing names?
7. For CHOOSE, are array-valued indexes and values iterated in matrix mode,
   and does a selected Reference remain a descriptor while an unselected
   Reference branch remains completely lazy?
8. For approximate HLOOKUP/LOOKUP/MATCH/VLOOKUP, confirm default flags,
   duplicate tie direction, case-folding domain, Number/Text mismatch, mixed
   Number/Text/Logical ordering, and whether unsorted input is a refusal,
   formula Error, or explicitly implementation-dependent.
9. For LOOKUP, confirm two- versus three-argument orientation, result-vector
   length mismatch, cell-range extension, Array mismatch, and error precedence.
10. For INDEX, confirm omitted/empty/zero row and column output shapes, whether
    scalar publication may project a returned row/column, and how an
    `AreaNumber` selection behaves for duplicate and 3-D records.
11. For OFFSET, confirm empty height/width slots, negative offsets, overflow and
    sheet-boundary checks, 3-D/list behavior, and whether the returned
    descriptor is cacheable before a cell consumer reads it.
12. What exact read/work/cancellation/resource order is required for approximate
    search, duplicate ties, selected errors, and lazy branches? This controls
    both independent `expected_reads` and the native comparison boundary.

No native capture should begin until these questions have contract answers and
the reviewer has fixed the resolver sheet order, test table contents, and
expected typed-failure policy.

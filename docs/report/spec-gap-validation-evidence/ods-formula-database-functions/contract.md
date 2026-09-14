# ODF database functions: normative contract

This contract scopes the database-function family in OpenFormula 1.4 Part 4. It
is an evidence and implementation contract for a future evaluator; it does not
claim that the current evaluator implements this family.

The source was read from the local distribution:

- archive: `3rdparty/specs/OpenDocument-v1.4-os.zip`
- archive SHA-256: `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4`
- member: `part4-formula/OpenDocument-v1.4-os-part4-formula.html`
- member SHA-256: `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1`
- relevant sections: §§4.11.8–4.11.11 and 6.9, with the common conversion and
  host-behavior rules in §§3.4, 6.1–6.3 and the aggregate-function rules in
  §§6.13, 6.16 and 6.18.

## Function set

The complete §6.9 family contains these twelve functions. `D`, `F` and `C`
denote the `Database`, `Field` and `Criteria` pseudotypes described below.
Unless a row says otherwise, the result is a `Number` and the operation is
performed on field values from records selected by `C`.

| Function | Signature | Required operation |
| --- | --- | --- |
| `DAVERAGE` | `DAVERAGE(D; F; C)` | Apply `AVERAGE` to the selected values. |
| `DCOUNT` | `DCOUNT(D; [F]; C)` | Apply `COUNT`; when `F` is omitted, count every record satisfying `C`. |
| `DCOUNTA` | `DCOUNTA(D; [F]; C)` | Apply `COUNTA`; when `F` is omitted, count every record satisfying `C`. |
| `DGET` | `DGET(D; F; C)` | Extract the selected field value from exactly one matching record; return `Error` when there are zero or more than one matching records. |
| `DMAX` | `DMAX(D; F; C)` | Apply `MAX` to the selected values. |
| `DMIN` | `DMIN(D; F; C)` | Apply `MIN` to the selected values. |
| `DPRODUCT` | `DPRODUCT(D; F; C)` | Multiply the selected values. |
| `DSTDEV` | `DSTDEV(D; F; C)` | Apply sample `STDEV` to the selected values. |
| `DSTDEVP` | `DSTDEVP(D; F; C)` | Apply population `STDEVP` to the selected values. |
| `DSUM` | `DSUM(D; F; C)` | Apply `SUM` to the selected values. |
| `DVAR` | `DVAR(D; F; C)` | Apply sample `VAR` to the selected values. |
| `DVARP` | `DVARP(D; F; C)` | Apply population `VARP` to the selected values. |

The optional brackets around `F` in `DCOUNT` and `DCOUNTA` are part of the
signature. They are omitted-field syntax, not an empty-string field name.
Section 6.9 states no additional function-specific constraints, so the common
parameter-conversion, error, sequence and aggregate rules remain applicable.

## Database, field and criteria values

### Database (§4.11.9)

A database is a rectangular, organized set of data with one or more fields and
zero or more records. Every record has a value for every field, although a
field's value may be empty. An evaluator that implements any database function
shall support a range as a database. The first row of such a range is the set
of field names; remaining rows are data records.

A single cell containing text is also a valid database: it has one field and no
data records. The specification advises that field names be unique without
regard to case for interoperability. It does not prescribe a particular
behavior for duplicate, blank, non-text, or case-colliding headers. That
behavior must be an explicit implementation profile, rather than silently
borrowing spreadsheet behavior from another application.

The formula `Database` pseudotype is a logical tabular input. It is distinct
from any package-level database-range declaration or named-range metadata. A
resolver must define how a formula reference becomes the rectangular database,
including sheet identity and finite extent, and must not infer a database from
unrelated workbook metadata.

### Field (§4.11.10)

`Field` is either `Text` or `Number`:

- Text selects the database field with the same name. Evaluators should match
  the name case-insensitively.
- Number selects the field numbered from left to right, starting at 1. The
  number is required to be a positive integer.
- A function accepting a field shall return `Error` if the selected field does
  not exist.

The specification does not define conversion of fractional, negative, zero,
or nonnumeric text field selectors, nor does it define duplicate-name or blank
header resolution. These are explicit profile decisions. The repository profile
uses a positive integer selector as an ordinal even when the selected header is
`Text`, `Empty`, `Logical`, `Number` or `Error`, and even when header names are
duplicated. A Text selector searches only Text headers case-insensitively and
requires exactly one match; a missing or ambiguous match is a typed `Value`
error. A unique empty Text header can therefore be selected by empty Text; an
`Empty` cell is not silently converted into an empty Text header. Fractional,
negative, zero and nonnumeric selectors are rejected. These conversions and
their error precedence must remain documented and budgeted.

### Criteria (§4.11.8 and §4.11.11)

A criterion is one cell `Reference`, `Number`, or `Text`, used to compare with
cell contents. A reference to an empty cell is converted to numeric `0`; this
conversion does not make every empty-cell comparison match. In the repository
profile, an Empty criterion cell read from a criteria range follows that
reference conversion and is a numeric-zero criterion. An explicitly empty Text
criterion (for example `=""`) retains the empty-value matching rules below.

Criteria is a rectangular set with at least one column and two rows. Its first
row names the fields to which the expressions in later rows apply. To select a
record, all expressions in one criteria row shall match. Thus the conjunction
within a row is normative. Section 4.11.11 does not state the aggregation rule
between multiple expression rows. The usual database-function interpretation
is “any matching criteria row” (OR between rows), but that cross-row OR should
be recorded as a chosen profile rule until the specification supplies a more
explicit sentence. An implementation must not accidentally treat rows as an
unbounded conjunction.

For a criterion value with an operator:

- A Number or Logical criterion matches equal cell content.
- A value may begin with `<`, `<=`, `>`, `>=`, or use infix `=` or `<>`.
- `=` with an empty value matches empty cells.
- `<>` with an empty value matches non-empty cells.
- `<>` with a non-empty value matches every cell content except that value,
  including empty cells.
- The text criterion `=0` explicitly does not match an empty cell, even though
  a reference to an empty cell converts to numeric zero.

For `=` and `<>` with a non-empty value that is not interpretable as a Number,
the host property `HOST-SEARCH-CRITERIA-MUST-APPLY-TO-WHOLE-CELL` determines
whether the entire cell or a matching subpart is used. The same whole-cell
versus subpart choice applies to other Text criteria. This is a host-controlled
matching mode, not permission to treat an arbitrary expression as a formula.

The exact text comparison is also affected by `HOST-CASE-SENSITIVE`. Errors in
criteria or in a candidate cell need an explicit precedence policy; they must
not be converted to an ordinary empty value merely to make a row match.

## Host-controlled matching and conversion

Section 6.9.1 makes database results dependent on the host properties in §3.4.
The implementation must expose or document these choices:

| Host property | Database consequence |
| --- | --- |
| `HOST-CASE-SENSITIVE` | Controls case sensitivity of text equality, inequality and ordered comparisons used by criteria. |
| `HOST-SEARCH-CRITERIA-MUST-APPLY-TO-WHOLE-CELL` | Selects whole-cell versus subpart matching for the applicable non-empty text criteria. |
| `HOST-USE-REGULAR-EXPRESSIONS` | Enables regular-expression matching in character-string comparisons/searches. The regex dialect, invalid-pattern result and interaction with wildcards are not specified here and require a profile. |
| `HOST-USE-WILDCARDS` | Enables `?` and `*`; `~` escapes a wildcard. The precedence when regex mode is also enabled requires a profile. |
| `HOST-LOCALE` | Affects locale-sensitive text-to-number/date conversion and comparison. |
| `HOST-NULL-YEAR`, `HOST-NULL-DATE` | Affect date parsing and serial-date interpretation where a text criterion is converted as a date. |
| `HOST-PRECISION-AS-SHOWN` | May affect numeric values obtained from referenced cells before comparison. |

The host-property names are normative; their defaults and some interactions are
not. In particular, do not claim a single fixed regex dialect, wildcard mode,
case policy or date format as universal ODF behavior.

The common conversion rules apply when the pseudotype is not already satisfied
(§§6.1–6.3):

- A single-cell reference is dereferenced; a multi-cell reference is subject
  to the specified implicit-intersection rule before scalar conversion.
- A reference to an empty cell converts to numeric zero for `Number`, and to an
  empty string for `Text`.
- `Number` and `Logical` can enter number sequences; text and logical treatment
  in references follows the relevant `NumberSequence`/`NumberSequenceList`
  rules. In particular, reference sequence extraction omits Empty and Text
  cells in the standard sequence conversion.
- Text-to-number conversion is implementation-defined and may be zero, an
  `Error`, or numeric parsing; locale can affect it. A database implementation
  must state which policy it uses.
- If any supplied value is an `Error`, the common rule is to return that error;
  when more than one error is encountered, the leftmost should be returned
  (§6.1). Pseudotype conversion failure is implementation-defined under §6.2,
  but it must remain a typed formula error and retain the caller's resource and
  cancellation boundaries.

Dates are numeric serial subtypes. A numeric date criterion therefore compares
by its already-valued serial Number; this profile performs no additional epoch
conversion. A textual date criterion may pass through locale/date conversion
(`DATEVALUE`/`VALUE` rules and the null-date host properties), but Part 4 does
not make one universal textual-date conversion mandatory for a database
criterion. This profile's finite numeric criterion grammar does not consult an
ambient locale; textual date parsing remains a separately documented host
choice if it is added later.

## Aggregate behavior inherited by the D-functions

The D-functions delegate their selected field values to the corresponding
aggregate, so the evaluator must preserve the aggregate's input sequence and
empty/error contract rather than applying a generic “all cells are numbers”
rule.

| Aggregate | Relevant Part 4 behavior |
| --- | --- |
| `AVERAGE` | Number sequence requires at least one Number; no numbers is `Error` (§6.18.3). |
| `COUNT` | Counts numbers and ignores other values in the number sequence; errors do not propagate; omitted-argument zero-parameter behavior is implementation-defined (§6.13.6). |
| `COUNTA` | Counts every nonblank value, including errors and empty strings; an empty cell is blank (§6.13.7). |
| `MAX` | Ignores nonnumbers; this profile returns 0 when no Number is selected, matching its `MIN` choice. The no-number result is not fully specified in §6.18.45. |
| `MIN` | Ignores nonnumbers and returns 0 when no numbers are present (§6.18.48). |
| `PRODUCT` | Multiplies numbers from `NumberSequenceList`; text in ranges is not included (§6.16.47). This profile returns the multiplicative identity 1 for an empty selection. |
| `SUM` | Adds numbers from `NumberSequenceList`; the sequence is constrained to be nonempty (§6.16.61), but that section explicitly permits evaluators to evaluate expressions that do not meet the constraint. This profile returns 0 for an empty selection. |
| `STDEV` | Requires at least two numbers and uses the sample formula (§6.18.72); fewer is `Error`. |
| `STDEVP` | Requires at least one Number and uses the population formula. This profile follows the typed `NumberSequence` reference rules and omits referenced Text, Logical and Empty cells; the contrary Text/Logical wording in the §6.18.74 summary is recorded as a specification inconsistency (§6.18.74 and §6.3.7). |
| `VAR` | Requires at least two numbers and uses the sample variance formula (§6.18.82); fewer is `Error`. |
| `VARP` | Requires at least one Number and returns 0 for one number (§6.18.84). |

For the selected empty profile, `DSUM` returns 0, `DMAX` returns 0 and
`DPRODUCT` returns 1. `DAVERAGE` and `DSTDEV`/`DVAR` retain their minimum-count
constraints; `DSTDEVP`/`DVARP` require at least one Number. These are explicit
repository choices wherever the individual aggregate section is silent or
uses a constraint rather than a result rule, and must be tested as such.

`DGET` is special: its function result is declared `Number`, but its prose says
to extract the value from the selected field. The implementation must define
numeric conversion of the extracted value and the error for a nonnumeric value;
it must not silently return an untyped Text/Logical value under a Number API.

For `DCOUNT` and `DCOUNTA`, omitted `F` changes the counted unit from selected
field values to records. It must not be represented as an empty field selector,
and it must not cause the implementation to scan or materialize an unrelated
field for every record.

### Selected error and short-circuit profile

An `Error` supplied as a literal/function argument for `D`, `F` or `C` is
propagated before database or criteria source preparation, following the common
leftmost-error rule in §6.1. During a scan, only criterion columns needed to
decide the current row and the selected field of a matching row are evaluated.
Errors in uninspected fields or records therefore do not fail the query. Within
one criteria row, expressions are checked left-to-right and a false expression
stops that row. Criteria rows are checked in order for each database record; the
first matching row stops further criteria-row checks for that record, while the
database scan continues for the aggregate. A needed criterion-cell Error is
returned as a typed formula error under this profile. This short-circuit
behavior is an implementation choice where Part 4 does not specify error
ordering, and must not be replaced by a global pre-scan for errors.

`COUNT` ignores Error values in its selected field, and `COUNTA` counts them as
nonblank, as their aggregate sections explicitly override the common error
propagation rule. For the other numeric aggregates and for `DGET`, an Error in
an evaluated selected-field value propagates before the aggregate result is
published. Empty selections still use the function-specific results above.

## Required resolver and resource boundaries

A conforming implementation needs an immutable database view with:

1. a stable semantic sheet/source identity for each referenced range;
2. a finite rectangular extent for the database and criteria ranges;
3. borrowed cell values where the provider can guarantee lifetime, or an owned
   conversion whose text, row, field and criteria storage is charged;
4. explicit handling of Empty, Text, Number, Logical, Date/Time subtypes and
   Error; and
5. a caller-retained execution context for cancellation, work and memory
   admission across range discovery, criteria matching, selected-record scans
   and aggregation.

Formula-bearing cells, cached values, unknown cell types and external names
need an explicit provider contract. A database evaluator must not silently use
an inert cached formula result as if it were freshly evaluated. If formulas are
unsupported in the provider, return a typed unsupported/error result before
partially updating an aggregate.

All dimensions and products need checked arithmetic. The implementation should
charge work for each physically inspected record/cell, criteria expression and
aggregate input, and charge retained indexes or owned text before allocation.
It must enforce `max_reference_cells` (or the equivalent aggregate read limit)
across the entire database/criteria operation, rather than resetting a counter
per field or per criteria row. Cancellation must be checked before provider
calls, before large reservations, during scans, and before publication of the
result. A refusal is atomic and preserves the structured resource kind,
observed amount, limit and scope.

The implementation should avoid expanding a rectangular range into one object
per cell when the provider can stream rows or use bounded run indexes. Any
index must retain the immutable snapshot identity and release its reservation
after its backing storage is dropped. Criteria-row matching should short-circuit
after a false expression while preserving the specified leftmost-error policy.

## Profile choices and proposed repository profile

These points are not safely inferred from the twelve function summaries and
must be recorded in the implementation ADR/tests. The implementation profile
below resolves some of them; the rest remain open until the evaluator ADR
chooses them:

- The treatment of a header cell that is Empty rather than an empty Text value
  when a Text field selector is supplied.
- The §6.18.74 STDEVP summary's treatment of Text/Logical values versus the
  reference `NumberSequence` omission rules.
- The exact diagnostic ordering when several needed criteria cells fail at
  once; this profile evaluates criteria rows and columns in source order and
  propagates the first needed typed error.
- Textual date-criterion parsing, locale, null-year/null-date handling and any
  timezone policy; already-valued numeric date serials require no extra epoch
  conversion in this profile.

### Proposed deterministic repository profile

The following choices are valid implementation-profile decisions, but should
not be described as requirements imposed by ODF where the cited text is only a
recommendation or leaves behavior implementation-defined:

- Criteria rows are ORed: a record is selected when every expression in at
  least one post-header row matches. The AND rule within each row is normative
  (§4.11.11); the cross-row OR is a repository choice because §4.11.11 does
  not state it explicitly. An empty criteria cell remains the §4.11.8
  empty-reference-to-Number-0 criterion; it is never treated as an unconditional
  wildcard row.
- Field text lookup is case-insensitive, as §4.11.10 says evaluators
  *should* match field names case-insensitively. A Text selector resolves only
  when exactly one Text header matches under that comparison; a missing or
  ambiguous match is the repository's typed `Value` error. A unique empty
  Text header is therefore selectable by the empty Text selector. An Empty
  header cell is not silently treated as an empty Text header. A positive
  integer selector always selects the corresponding field from left to right,
  starting at 1, regardless of whether that field's header is Text, Empty,
  Logical, Number, Error or duplicated with another header. Non-positive, fractional,
  and out-of-range numeric selectors are typed `Value` errors. The
  specification mandates an `Error` for a nonexistent selected field, but does
  not mandate this particular error subtype; §6.2 leaves pseudotype conversion
  failures implementation-defined.
- Criteria headers are `Field` selectors, so a numeric criteria header is
  valid even when the database header values are non-text. Repeated criteria
  columns selecting the same field are valid and combine with the normative
  AND within that criteria row; they are needed for bounds such as `>=1` and
  `<=10`. An Empty (blank, non-Text) criteria-header cell is a missing Field
  selector and is rejected with `Value`; it is not an ignored or wildcard
  column. A unique empty Text header remains selectable by an empty Text Field
  selector. Database header values are not globally rejected merely for being
  non-text, duplicated or case-colliding, including an Error-valued header when
  a numeric selector addresses it.
- Criteria text uses the existing `Sensitive` text-comparison policy,
  whole-cell matching (`HOST-SEARCH-CRITERIA-MUST-APPLY-TO-WHOLE-CELL = true`),
  regular expressions disabled, and wildcards disabled. These are fixed host
  settings for deterministic evaluation, not universal ODF defaults: §3.4
  expressly makes case, whole-cell, regular-expression and wildcard behavior
  host-defined. Regex dialect and regex/wildcard precedence therefore do not
  apply in this profile, though a future configurable host may define them.
- Numeric recognition of scalar Text criteria uses the existing
  locale-independent finite-number grammar. If recognition fails, the value
  follows the text-criterion path under whole-cell matching. This is permitted
  by §6.3.5, which makes Text-to-Number conversion implementation-defined and
  notes locale dependence; it must be documented as a profile rather than
  presented as the only ODF behavior. Number cells, including Date/Time
  numeric subtypes, compare by their already-valued serial Number. No ambient
  locale is consulted for this profile's numeric criterion parsing.
- `DGET` follows its §6.9.5 return declaration and converts the single selected
  field to `Number` using the profile's existing conversion rules: Empty becomes
  numeric 0 through the reference-to-Number rule, Logical becomes 0 or 1,
  numeric Date/Time values remain serial Numbers, and unconvertible Text or
  Complex values produce a typed conversion error. The section's prose says
  “Extracts the value,” but its explicit `Returns: Number` line is the stronger
  result contract. Preserving Text, Logical, Complex or Empty as a generic
  public `Value` is possible as a separate raw-value API, but making that the
  result of `DGET` would be a deliberate variance from §6.9.5 and must be
  labelled as such, not silently called exact conformance.

The profile's `Value` error choices above are repository policy. Part 4 usually
says only `Error` for missing fields or conversion failures, and §6.2 makes
pseudotype mismatch behavior implementation-defined; tests must assert the
chosen subtype and its precedence explicitly.

Tests should cover at least: one-cell and multi-record databases; omitted and
explicit fields; field names with case differences; missing/duplicate/blank
headers; empty records; `=0`, `=\"\"`, `<>\"\"`, `<>nonempty` and a reference to
an empty criterion cell; AND within a row and multiple criteria rows; Number,
Logical, Text, Date and Error cells; every host matching mode; regex/wildcard
edge cases; zero/one/many DGET matches; empty aggregates; sample versus
population statistics; cancellation and each resource limit during scanning;
and an immutable provider whose cached formula values are deliberately inert.

This contract is intentionally a complete dependency and ambiguity inventory.
It leaves no database function outside the twelve §6.9 entries, while keeping
host-selected behavior and currently unimplemented provider/evaluation work
visible for subsequent implementation batches.

# Independent review of the ODF conditional aggregate batch

## Review status

The frozen implementation is accepted for the bounded repository profile.
Semantic and resource reviews found no remaining blocker. All seven isolated
gates passed, and root independently verified the exact oracle, native receipt,
and all 1,260 performance samples. The receipts and limitations below define
this acceptance; they do not establish whole-workbook recalculation support.

At the latest independent snapshot, the conditional suite passed 22/22, the
exact Fraction oracle passed 6/6 in both scalar and matrix demand modes, the
resource/ownership suite passed 11/11, the shared database suite passed 20/20,
the shared criterion text profile passed 1/1, and the selected native receipt
passed 2/2. The native receipt contains 64 selected rows across six functions
and 16 explicit exclusions. The R38 bare-numeric Text observation is now
selected under the corrected matcher; the stale R39 cache remains excluded.

## Source basis

The normative source is the local ODF 1.4 Part 4 member recorded in
`contract.md`:

| Input | SHA-256 |
| --- | --- |
| `3rdparty/specs/OpenDocument-v1.4-os.zip` | `9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4` |
| `part4-formula/OpenDocument-v1.4-os-part4-formula.html` | `ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1` |

The key normative sentences are §4.11.8's distinction between a reference to
an Empty cell (numeric zero) and an explicitly empty Text criterion, the
Reference/ReferenceList distinction in §§4.8–4.9, and the range and generated
destination constraints in §§6.13.9–10, 6.16.62–63, and 6.18.5–6.

## Decisions that must remain visible in code and tests

The following are deliberate profile choices or direct requirements. They
must not be inferred from the existing database matcher solely because that
module has similarly named criterion code.

* `SUMIF` and `COUNTIF` accept ordered `ReferenceList` criterion ranges. The
  list is traversed by logical reference occurrence, then 3-D plane order,
  then row-major cell order. Duplicate and overlapping occurrences count
  independently. A `SUMIF` call with an optional destination rejects a list
  containing more than one logical reference with formula `#VALUE!`, as the
  specification explicitly requires.
* `SUMIFS`, `COUNTIFS`, and `AVERAGEIFS` use one logical `Reference` per range,
  and every range must have equal plane count, rows, and columns. A
  ReferenceList in one of these slots is a formula pseudotype error and is
  never flattened or implicitly intersected. A single logical 3-D reference
  is retained as one reference and is paired plane-for-plane.
* A criterion is scalar. A `Number`, `Logical`, or `Text`, or one single-cell
  reference to one of those values, is valid. A multicell reference, a 3-D
  reference with more than one cell, an inline Array, and a ReferenceList are
  formula `#VALUE!`; they are not silently reduced by the caller's implicit
  intersection. A reference to an Empty cell is numeric zero. A syntactically
  empty required parameter remains `#VALUE!`.
* Optional `SUMIF`/`AVERAGEIF` destination ranges are top-left anchors. A
  one-cell anchor such as `C1` expands to the criterion range's rows and
  columns. The supplied destination extent is ignored. `SUMIF` follows its
  explicit §6.16.62 rule and silently clips generated rows and columns beyond
  sheet bounds. §6.18.5 gives `AVERAGEIF` no such permission, so this profile
  returns `#REF!` if any generated average destination is outside the sheet.
  A destination cell is read only after its corresponding criteria match.
  For a 3-D criterion range, a one-plane destination is still a valid anchor:
  generated destination planes begin at the anchor's plane and follow the
  criterion range's sheet order; actual destination extent is ignored. A
  generated plane beyond the provider's sheet order is `#REF!` for both
  functions; SUMIF's silent clipping applies only to generated rows/columns.
* `AVERAGEIF` and `AVERAGEIFS` return the existing average error
  `#DIV/0!` (`ScalarError::DivisionByZero`) when there are no selected Number
  values, including no matches. `SUMIF`/`SUMIFS` return exact numeric zero
  for an empty selected set. Text, Empty, Logical, and Complex destination
  cells are omitted from the sum/average; selected formula Errors propagate.
* Criterion arguments are eagerly evaluated in source order. A direct
  criterion Error, or an Error in its one-cell reference, wins before any
  range scan. Geometry errors are generated before the scan. Once scanning
  begins, positions are visited by list occurrence, plane, row, and column;
  criterion pairs are read left-to-right and the selected destination is read
  last. Thus an earlier selected destination Error wins over a later
  criterion-range Error, while a false earlier `IFS` criterion prevents a
  later criterion cell at that position from being read. Criterion Errors are
  never coerced to Empty or a nonmatch. Optional destination Errors are
  observed only for selected positions, including the explicitly lazy
  `AVERAGEIFS` value read. Typed provider, cancellation, source-version,
  allocation, and resource failures remain evaluator failures.
* The fixed matching profile is case-sensitive, whole-cell, literal text;
  regular expressions and wildcards are disabled. Only operator-prefixed
  numeric comparator text is parsed with the existing finite parser. A bare
  numeric-looking Text criterion remains Text and matches a Text candidate,
  not a Number candidate; §4.11.8's “Other Text value” branch makes this a
  normative criterion distinction rather than host matching latitude. Explicit
  `"=0"` does not match an Empty cell, whereas `"="` matches Empty and `"<>"`
  matches non-Empty. Whitespace remains significant. This profile is separate
  from any future host-configured regex/wildcard/substring behavior.
* The resolver-free scalar evaluator has no reference capability. A scalar
  call with only constant range arguments produces formula `#VALUE!` after
  eager arguments are consumed; a reference argument reaches the established
  typed `Unsupported(Reference)` boundary. Actual conditional execution
  belongs to the value evaluator.

## Current shared matcher review

The new private `value/criteria.rs` extraction is a sound ownership seam for
sharing criterion parsing with the database functions, provided the
conditional caller supplies the correct input kind. It retains borrowed text,
charges criterion and candidate text bytes, and preserves the accepted
case-sensitive literal comparison profile. The following review checks are
closed by the focused database, conditional, and criterion-profile suites:

1. `CriterionValue::from_runtime_value` currently maps `RuntimeValue::Empty`
   and `RuntimeValue::Missing` to `CriterionValue::Empty`. The conditional
   caller must convert a syntactically missing required criterion to formula
   `#VALUE!` before this helper is called. Only a physically Empty cell
   reached through a single-cell reference may become numeric-zero criterion
   state.
2. A `RuntimeValue::ScalarCell` is projected through the resolver before
   conversion. The latest conditional caller handles this case directly;
   retain the read-count regression so a future refactor cannot turn valid
   single-cell criterion references into a generated Value error.
3. An array or reference-list range must be rejected from the conditional
   range slot without dereferencing its cells. Constant arrays produce a
   formula Value; a resolver capability refusal stays typed. This distinction
   is observable through read counters and `IFERROR`.
4. The matcher currently propagates every candidate `CriterionValue::Error`.
   Conditional range scans must preserve that behavior, but destination values
   must be read only after all criterion matchers return true. Reusing the
   matcher must not introduce a destination pre-scan. Retaining a formula
   Error while scanning is consistent with the database evaluator; any later
   typed provider failure must still propagate and supersede that pending
   formula result because the observation is incomplete.
5. The helper's Empty criterion branch is correct only for an explicit empty
   Text criterion. A physically Empty criterion reference must be converted to
   numeric zero before compilation. Tests need both `COUNTIF(R;"")` and
   `COUNTIF(R;[.empty])` to prevent this semantic collapse.
6. The shared matcher must not parse a bare numeric-looking Text criterion as a
   Number. The latest matcher gates numeric parsing on an explicit operator.
   The conditional and database callers still need a regression with a Text
   candidate `"3"` and a Number candidate `3`: criterion Text `"3"` selects
   only the Text candidate, while explicit `"=3"` uses numeric comparator
   semantics. The former parse-all-finite-text behavior was a pre-existing
   matcher defect, not a host-profile option.

The database and conditional suites both retain their empty-criterion cases,
and the shared matcher correction deliberately changes bare numeric-looking
Text from the former parse-all behavior to the §4.11.8 “Other Text value” rule.
No public criterion type was added.

## Current implementation disposition

The earlier construction blockers are closed in the latest snapshot: borrowed
criterion text has source-tied lifetimes; empty averages return
`DivisionByZero`; a one-record ReferenceList is admitted for `SUMIF` with `S`;
no-alternate scans reuse the criterion cell; generated anchor planes advance in
sheet order; `AVERAGEIF` geometry is checked before reads; match counts use
checked arithmetic; criterion arguments enter scalar context; and missing
criteria remain `#VALUE!`. The projection fixture records its parser shape and
the limits helper no longer returns a value borrowing a local expression.

All semantic review items are closed in the current snapshot:

* The shared matcher correction now has a passing focused regression in
  `ods_formula_criterion_text_profile.rs`: a Text candidate `"3"` and a Number
  candidate `3` distinguish bare Text `"3"` from explicit `"=3"` in both DSUM
  and SUMIF. This test passes in the final gate.
* The focused error-order test now covers an observed Formula Error followed by
  a typed provider failure; the typed failure supersedes the retained formula
  Error because the observation is incomplete. The final-overflow fixture also
  confirms that a selected cell Formula Error wins over a later generated
  `#NUM!` conversion error, as required by `contract.md`. Both tests pass in
  the final gate.
* The isolated full test, all-target Clippy, rustdoc, and format gates now pass.
  Boundary and performance receipts also pass independent verification.

## Performed implementation review

The frozen conditional module was independently inspected for:

* arity and pair-parity validation, including direct Error precedence before
  generated arity/shape errors;
* range shape admission before cell reads, cumulative reference-cell limits,
  checked coordinate arithmetic, and no hidden materialization of all ranges;
* correct 3-D plane pairing and ReferenceList occurrence order;
* top-left anchor construction independent of the optional destination's
  actual dimensions, with explicit out-of-bounds clipping and no accidental
  resolver read beyond the sheet extent;
* row-major criterion and destination read order, lazy AND short-circuiting,
  selected-value lazy reads, and leftmost formula-error retention;
* exact SUM/AVERAGE accumulator reuse, one final binary64 conversion, and
  `DivisionByZero` for no average values;
* projection-cache identity and lazy-branch behavior; and
* a resolver-free scalar refusal test that distinguishes typed Reference
  capability from formula Value for constant range arguments.

The full isolated integration, 288-observation Fraction oracle, and native
receipt pass. The independent resource review closes all four findings, and
final boundary and performance receipts are verified.

## Performance review boundary

The performance harness must prepare and preflight all scalar, matrix,
reference-list, 3-D, anchor, shape-mismatch, empty-average, and lazy-value
cases before any final samples. It must record cumulative resolver reads,
work, retained storage, allocator balance, and external RSS in fresh child
processes. The final report may claim only the bounded resolver-backed
evaluator profile; it may not claim whole-workbook recalculation, native
producer acceptance, regex/wildcard host coverage, or a language-level RSS
bound. Any material RSS shift requires an explicit review note rather than a
noise assumption.

## Final isolated gate receipt

Root verified all seven retained gates against the frozen source closure: 1,313
ODS tests passed with zero failures or ignored tests, strict all-target Clippy,
strict rustdoc, both formatting checks, crate boundaries, and diff checks.
The independent resource review closes all four resource findings. Performance
conclusions are recorded separately in the final performance report.

## Final performance disposition

Root independently verified all 270 baseline and 990 candidate raw samples,
source and profile identity, allocation balance, reference-read bounds, and
reported medians. Matched allocation counts, requested bytes, and peak live
bytes are unchanged. Unrounded time shifts are −3.17% to +4.30% and RSS shifts
are −2.56% to +4.51%; no matched metric exceeds the 5% review threshold.
Repeated sheet metadata lookup remains a bounded optimization opportunity.
No material performance blocker was found for these measured workloads.

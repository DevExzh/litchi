# Database evaluator source review

This is a source-only review of the current database dispatch and query path;
it is not a build or runtime acceptance report. The reviewed source snapshot
has these SHA-256 digests:

- `crates/litchi-ods/src/codec/formula/evaluation/value.rs`: `35d441d137123348f7b57a7fbb222e960c7e4bc5137eeeec25077c45bf581187`
- `crates/litchi-ods/src/codec/formula/evaluation/value/database.rs`: `99df2353ba7086dc8bedd08a12e979f415ef281f484133b78257a0651421a281`
- [the normative contract](contract.md): `e46a6e31fd13d848bd28a2f838a63aae818932f4eeacf2878d94d77485068f13`

## Disposition

The previously identified source blockers are resolved in this snapshot:

- `DCOUNT` and `DCOUNTA` admit both the two-argument omitted-field form and
  the explicit three-argument form. Database and criteria arguments are
  scheduled in matrix context, while the field remains scalar.
- Error-valued database headers are retained, so positive numeric selectors
  remain ordinal. Text selectors inspect only Text headers.
- Empty criteria cells compile as numeric-zero criteria. An Empty criteria
  header is rejected as a missing Field selector; a unique empty Text header
  remains selectable.
- Criteria rows are ORed and clauses within a row are ANDed. Database cells
  needed by clauses are fetched lazily in clause order, with reuse across
  criteria rows. The selected field is fetched only after a row matches.
- Provider Text reads are checked against `max_text_bytes`, and an internal
  `RuntimeElement::Missing` is preserved as typed `NotAvailable` rather than
  being converted to Empty.

`DGET` records only the first selected value while scanning. Its result checks
  the matching-record cardinality before converting that value: zero or
  multiple matches return `Value`, including when a selected value in a
  multi-match result is an Error; exactly one match then applies the declared
  Number conversion and preserves a selected Error.

I found no additional concrete blocker in the reviewed dispatch, query,
provider-read, or aggregate boundary. The pending public out-of-shape
regression and the post-fix full gates are runtime evidence owned by the test
and root agents; this review does not claim they passed.

## Root mutation check

The public computed-matrix regression passes, but also passes after replacing
only `RuntimeElement::Missing => DbCell::Error(NotAvailable)` with the old Empty
conversion in the isolated workspace. The tested public matrix paths already
materialize an Error before this private boundary. Consequently, these tests
prove end-to-end error preservation for the exercised matrix expressions; they
do not prove that the private Missing arm is publicly reachable or distinguish
its conversion. The corrected arm retains the VM's documented invariant as a
defensive boundary fix. `missing-mutation-09` records the exact replacement,
original/mutated source hashes, command, log, successful status, and restoration.
Canonical production source was never mutated for this check.

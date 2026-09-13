# ODS DDE metadata transactions

This batch addresses the ODF spreadsheet audit's missing public mutation path
for inert DDE source declarations and formula-link cached tables. The baseline
is `5347801fc`; the existing `dde::Snapshot` supplied inspection only.

The implementation belongs to the ODS format owner. Source declarations retain
application, topic, item, optional name, conversion mode, and automatic-update
presence as data. No operation starts a DDE conversation, resolves a source,
refreshes cached values, or evaluates formulas.

## Specification and preservation contract

The normative reference is ODF 1.4 Part 3 §§9.8 and 14.7 and the bundled
`schemas/OpenDocument-v1.4-schema.rng` in
`3rdparty/specs/OpenDocument-v1.4-os.zip`.

- A worksheet source is a direct `office:dde-source` child of `table:table`,
  after optional title, description, and table-source owners and before optional
  scenario, forms, shapes, column, and row content.
- Formula links belong to the direct spreadsheet `table:dde-links` epilogue.
  A present container contains at least one `table:dde-link`; each link contains
  one source followed by one cached `table:table`.
- Section 14.7.1 requires named connections although the RNG makes `office:name`
  optional. New or replaced declarations therefore require an explicit name;
  unnamed producer declarations remain readable and exactly preservable.
  Editing these declarations does not rewrite formulas or text-field connection
  declarations, whose references belong to separate owners.
- The RNG reuses the general table grammar for a cache, but §14.7.4 additionally
  specifies that only cell-attribute data is used, cells remain empty, and the
  table contains no style information. Authored text therefore belongs in
  `office:string-value`, not a `text:p` child. Schema success alone cannot prove
  this cache-specific semantic requirement. Existing richer producer fragments
  remain opaque for exact preservation.
- Changing a source declaration must retain the associated cached table bytes.
  Replacing a plain scalar cache retains its existing table name. Styles,
  covered/merged cells, grouping, presentation text, and extension structures
  outside the scalar model make a cache opaque to typed replacement.
  Unsupported attributes, namespace scopes, comments, processing instructions,
  and extension markup must survive a supported edit or cause a typed refusal.
- Exact no-ops and inverse patches retain exact source XML. Changed package
  publication refuses signed sources; source-backed publication also checks the
  captured live source before writing.

`validate_schema.py` validates emitted DDE owners against the bundled RNG and
reports whole-content validity separately. It never dereferences links or starts
a native application. Run it with `--spec <OpenDocument-v1.4-os.zip> --content
<generated-content.xml>`; owner validation success does not certify every other
part of the package.

## Validation scope

The final all-feature/all-target ODS gate passes 770 tests in 45 targets,
including 33 transaction and 12 facade regressions. Strict clippy, rustdoc,
doctests, crate boundaries, and scoped formatting pass. There are no doctests
in the current ODS crate. Commands and statuses are in
[`gates/gate-results.json`](gates/gate-results.json); the two source manifests
record unchanged source across the final gate sequence.

`ods_dde_transactions.rs` covers source/link CRUD, exact inverses, ambiguous
selectors, source-only edits, opaque cache text and references, control-character
escaping, empty caches, flat roots, cancellation, output limits, and retained
patch memory. `ods_dde_facades.rs` covers ordinary/mutable failure atomicity,
source-backed publication and raw unrelated-member retention, stale sources,
signed-source refusal, destination limits, diagnostic scope, and BOM coordinates.
Memory regressions cover admission of caller-owned cache buffers, failed
admission without retained draft changes, and reservation release when staged
payloads are replaced or removed. Admission counts buffer capacity. A link
specification computes that bound when constructed; staging can then reserve
the payload without rescanning its cells. Edits have exclusive ownership;
snapshots and patches share their retained source buffers.
The parser's 12 regressions include namespace storage admission and active-depth
budgets. Final review dispositions are in [spec-review.md](spec-review.md) and
[resource-review.md](resource-review.md).

The replayable [authoring runner](authoring-runner/src/main.rs) emits
[`fixtures/api-authored-content.xml`](fixtures/api-authored-content.xml) through
the public API. It passes the whole-content ODF 1.4 RNG and cache-specific prose
checks. Independent lxml decoding confirms exact tab/LF/CR values and the empty
cache cell. These are synthetic authoring fixtures. The wider suite includes
existing producer fixtures, but this batch does not establish native DDE
save/reopen interoperability or DDE execution.

Performance measurements distinguish parsing, staging, commit, and end-to-end
work. Work and Memory are execution-budget counters; process maximum RSS is a
separate operating-system measurement. The read-only baseline is comparable only
where the same operation and input exist on both revisions. The
[final profile](candidate/report.md) records a large-cache parse p50 of 29.9 ms
versus 8.2 ms at baseline (3.65× slower on the shared host), with corrected
staging admission of 2,103,553 bytes for the prebuilt cache. No transaction
speedup is claimed against the read-only baseline.

The isolated baseline checkout with the scoped ODS overlay also passes all
45 DDE transaction/facade tests. [Root verification](root-verification.json)
records source and artifact hashes, 24 verified benchmark lanes, exact patch
replay, and cleanup. Root gates ran in the shared workspace; isolated checks
exclude its unrelated OPC/XLSX edits.

The profiling generator emits one self-closing `table:table-column` in every
worksheet and cache table, and keeps cached cells attribute-only. These
synthetic streams exercise the DDE owner shape and are content.xml fixtures,
not complete ZIP packages. The final generator uses
`office:version="1.4"`, supplies one empty cell in each worksheet row, and the
three generated shapes pass complete-content RNG validation in
[`fixtures/generated-shapes-schema-validation.json`](fixtures/generated-shapes-schema-validation.json).

The wider specification audit remains open, including DDE sessions and refresh
as intentionally unimplemented execution capabilities.

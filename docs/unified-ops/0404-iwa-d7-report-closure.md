# 0404 — D7 report and CRUD closure

## Scope

This isolated harness-only repair starts at coordinator source
`c07e6208ccd39387665ac2821b2aeb895f00bc22`, whose runtime is unchanged by
this slice. It closes the row-25 report failures found by the c07 full gate:

* the tracked Numbers two-sheet chart-font fixture was missing from the
  Numbers CRUD corpus, so the corpus fidelity and query digest goldens were
  stale;
* Pages and Numbers chart-grid fixtures were absent from `crud_matrix`,
  leaving the ten claimed chart `get`/`find`/`query`/`view`/`set` cells
  unreached;
* the Pages table recipe used the old `tbl` spelling and a fixture that does
  not provide the current table-copy surface;
* the report's generic section-removal scan stopped before the seeded,
  non-empty `/section[2]` sibling and therefore only exercised Pages'
  intentionally non-removable first section.

The owned paths are limited to
`crates/litchi-ops/tests/crud_matrix.rs`, its generated report outputs under
`docs/unified-ops/results/0289/`, the query and registry golden files under
`crates/litchi-ops/tests/golden/`, and this record. No production runtime,
leaf/shared engine, native fixture, capability schema, ledger, or `PROGRESS.md`
was changed.

## Harness changes

`crud_matrix` now includes the existing tracked native chart-grid fixtures:

* `test-data/iwork/pages/chart-grid-authored-native.pages`;
* `test-data/iwork/numbers/chart-grid-authored-native.numbers`.

Pages also retains a table-copy fixture at
`test-data/iwork/pages/body-table-visible.pages`. Its recipe uses the current
canonical `table` element name and `{rows, cols, data}` properties. This is a
test-harness correction, not a new public route.

The Pages seed creates an introductory paragraph, adds a named second
section, and sets its text to `Tail`. The bounded remove candidate window is
24 instances, enough to reach that sibling while retaining a finite scan.
The first section is still attempted and still receives its typed refusal;
the successful row removes only the seeded second section. Chart CRUD writes
use a known changed point size (`17pt`, or `19pt` when already `17pt`) so the
report proves mutation, reopen, edit and batch paths rather than a chart
no-op.

## Generated closure

The affected generated files were regenerated from the c07 source:

* `fidelity.{md,tsv}` now includes the existing two-sheet Numbers fixture;
* `tests/golden/query_corpus/digests.json` changes only the Numbers digest
  family attributable to that fixture (40 opened files, 43 sheets and 42
  tables in the current query corpus);
* Pages, Keynote and Numbers registry snapshots contain the chart
  `font.size` element/property declarations already present in the migrated
  runtime;
* `crud-checklist.{md,tsv}` and `crud-adr0008.tsv` contain successful Pages
  and Numbers chart read/query/view/set/edit/batch rows, zero unreached
  claimed cells, and successful pre-existing Pages section/table coverage.

No generated report adds a new failure, changes a native claim, or relaxes a
checker. The report outputs remain reduced text/JSON/TSV artifacts; no raw
captures, binaries, or large evidence were added.

## Verification

All commands used the coordinator lock copied from the frozen source,
`--locked --offline`, the capped cargo shim, `CARGO_BUILD_JOBS=4`,
`CARGO_INCREMENTAL=0`, and an external target directory:

```text
LITCHI_BLESS=1 cargo test --locked --offline -p litchi-ops --all-features \
  --test crud_matrix -- crud_checklist_per_format                 # PASS
LITCHI_BLESS=1 cargo test --locked --offline -p litchi-ops --all-features \
  --test corpus_fidelity -- open_empty_edit_save_is_byte_identical # PASS
LITCHI_BLESS_GOLDEN=1 cargo test --locked --offline -p litchi-ops \
  --all-features --test query_corpus_differential \
  -- corpus_queries_return_what_they_returned_before                 # PASS
LITCHI_BLESS=1 cargo test --locked --offline -p litchi-ops --all-features \
  --test registry -- registry_goldens_follow_enabled_formats        # PASS
```

The focused CRUD result reports zero unreached claimed cells. The corpus
scan passes 1,022 claimed-format inputs; the generated fidelity table has
1,000 opened rows after its existing refusal/extension classification. The
query differential passes all 895 digest keys after the Numbers-only golden
refresh. Registry goldens pass for Pages, Keynote and Numbers.

The original failing c07 receipt remains preserved externally at
`.coord/resume-20261028/full-c07e6208cc/25-tests-ops.out`; it is historical
failure evidence, not a green result. The current integrated candidate still
requires the coordinator's complete gate on one clean commit. Native
verification and performance qualification are outside this report-closure
slice and are not claimed here.
